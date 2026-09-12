#!/usr/bin/env python3
"""Trojan AI for Office 本地服务：静态文件 + 模型列表代理（绕过浏览器跨域限制）

安全模型：服务只绑 127.0.0.1，但这挡不住「别的网页借用户浏览器来访问」。
所以所有 /api/* 端点要求启动时生成的随机 token（通过 taskpane.html 注入给插件页面），
并且不下发任何 CORS 头——跨域页面既读不到 taskpane.html 里的 token，也读不了 /api/* 的响应。
"""
import json
import os
import posixpath
import secrets
import threading
import time
import urllib.request
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs, unquote, quote

API_TOKEN = secrets.token_hex(16)

# 【批次 B】共享设置的「读现状-贫瘠闸-写文件」整段串行化：ThreadingHTTPServer 允许并发请求，
# 两个 Office 应用同时保存时交错读写会把对方的结果静默吞掉（review 第 5 节）。锁只护这一段，
# 不影响静态文件与其他 API 的并发。
SETTINGS_LOCK = threading.Lock()

# ---- 【第 11 节】真机测试桥：内存命令队列 + 事件转发。----
# 安全模型：/api/test/* 与其他 /api/* 同级 token 鉴权 + 运行期短期身份（runId/instanceId）；
# 只接受有限命令白名单，不接受任意代码/路径；纯内存不落盘；单宿主同时只允许一个测试运行；
# 【R1】命令定向投递：每条命令带 targetInstanceId；写命令必须先 bind（bind 需该实例先有成功
# attach_result）；回执按实例归属结算，别的实例伪发回执不能结算；runner 租约独立于桥心跳。
# 生产正常使用时无运行，插件侧只做低频可用性探测。
TEST_LOCK = threading.Lock()
TEST_RUNS = {}  # runId -> run dict，结构见 /api/test/run
TEST_COMMANDS_ALLOWED = ('attach_fixture', 'send_prompt', 'decide_card', 'snapshot', 'stop', 'finish', 'end_model_script', 'capability_probe', 'host_metadata')
TEST_WRITE_COMMANDS = ('send_prompt', 'decide_card')  # 【R1】必须先绑定才准入队
TEST_CONTROL_COMMANDS = ('stop', 'finish')            # 优先投递，不被普通命令串行阻塞
# 【R5】命令状态机：queued → delivered → running → completed/failed；只有匹配类型的合法终态事件能结算
TEST_SETTLE_EVENT = {
    'attach_fixture': 'attach_result', 'send_prompt': 'prompt_accepted',
    'decide_card': 'card_settled', 'snapshot': 'snapshot_result',
    'stop': 'stopped', 'finish': 'finished', 'end_model_script': 'model_script_ended',
    'capability_probe': 'capability_result',
    'host_metadata': 'host_metadata_result'   # 【R5】只读元数据命令的合法终态
}

def _env_ms(name, default):
    try:
        return max(1, int(os.environ.get(name) or default))
    except ValueError:
        return default

TEST_TTL_SECONDS = 30 * 60
TEST_MAX_COMMANDS = 50
TEST_MAX_EVENTS = _env_ms('OAI_TEST_MAX_EVENTS', 2000)
TEST_REDELIVER_MS = _env_ms('OAI_TEST_REDELIVER_MS', 60000)   # 投递后多久无回执允许补投
TEST_RUNNER_LEASE_MS = _env_ms('OAI_TEST_RUNNER_LEASE_MS', 120000)  # runner 租约：断线失效触发停止

def _append_test_event(run, etype, payload, instance_id, command_id=''):
    """事件只增序号（裁剪不回退）；游标落后于保留区时由 events 端点报缺口。"""
    run['eventSeq'] += 1
    run['events'].append({'seq': run['eventSeq'], 'ts': time.time(),
                          'type': str(etype)[:60], 'instanceId': str(instance_id)[:64],
                          'commandId': str(command_id)[:64], 'payload': payload or {}})
    if len(run['events']) > TEST_MAX_EVENTS:
        run['events'] = run['events'][-TEST_MAX_EVENTS // 2:]

def _gc_test_runs():
    """【R3/S4】runner 租约失效 → 进入停止流程：
    已绑定的运行先转 stop_requested 并向绑定实例定向投递停止命令（控制通道保持，
    桥还能收到并走原停止流程）；二次宽限仍无回执才强制作废（保留事件证据）。
    未绑定的准备运行直接取消——没有测试任务可停，不广播任何 UI 停止动作（S3）。"""
    now = time.time()
    for rid in list(TEST_RUNS.keys()):
        run = TEST_RUNS[rid]
        lease = TEST_RUNNER_LEASE_MS / 1000.0
        if not run['stopped'] and now - run['runnerLastSeen'] > lease:
            bound = run['bound']
            if bound and not run['stopRequested']:
                run['stopRequested'] = True
                _append_test_event(run, 'runner_lost', {'reason': 'runner lease expired, stop requested'}, '')
                if not any(c['command'] == 'stop' and c['state'] in ('queued', 'delivered', 'running') for c in run['commands']):
                    run['commands'].insert(0, {'commandId': secrets.token_hex(6), 'command': 'stop',
                                               'params': {}, 'target': bound['instanceId'],
                                               'state': 'queued', 'idempotencyKey': '',
                                               'deliveredTo': None, 'deliveredAt': 0.0})
            elif bound and run['stopRequested'] and now - run['runnerLastSeen'] > 2 * lease:
                run['stopped'] = True  # 二次宽限仍无确认：强制作废，事件已留痕
            elif not bound:
                run['stopped'] = True
                _append_test_event(run, 'runner_lost', {'reason': 'runner lease expired (prep run, cancelled)'}, '')
        if run['stopped'] and now - run['runnerLastSeen'] > 3600:
            TEST_RUNS.pop(rid, None)

def _target_host_allowed(url):
    """【SSRF 防护】服务端发起的外网请求仅允许 http/https，且解析后的 IP 必须是公网地址：
    拒绝环回/私有/保留/链路本地/组播/未指定地址（含域名解析到私网的情况）。"""
    import socket, ipaddress
    parsed = urlparse(url)
    if parsed.scheme not in ('http', 'https') or not parsed.hostname:
        return False, 'url must be an absolute http(s) link'
    try:
        infos = socket.getaddrinfo(parsed.hostname, parsed.port or (443 if parsed.scheme == 'https' else 80), proto=socket.IPPROTO_TCP)
    except Exception as e:
        return False, 'cannot resolve host: %s' % str(e)[:120]
    for info in infos:
        ip = ipaddress.ip_address(info[4][0])
        if ip.is_loopback or ip.is_private or ip.is_reserved or ip.is_link_local or ip.is_multicast or ip.is_unspecified:
            return False, 'target host resolves to non-public address %s' % ip
    return True, ''

class Handler(SimpleHTTPRequestHandler):
    def end_headers(self):
        # 禁缓存：Office 侧边栏的 WebView 会按 URL 缓存静态文件，没有这几行，
        # 改了代码但版本号没变就会一直加载旧文件（改代码不生效的「灵异事件」根源）
        self.send_header('Cache-Control', 'no-store, must-revalidate')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

    def check_api_token(self, q):
        token = (q.get('token') or [''])[0] or self.headers.get('X-Local-Token') or ''
        return secrets.compare_digest(token, API_TOKEN)

    # 【任务 A】静态文件白名单：/api/* 之外的一切路径默认拒绝。此前仅 /api/* 校验 token，
    # 其余请求全部交给 SimpleHTTPRequestHandler——仓库目录下的 data/（含共享设置+key）、
    # .git/、源码都能被无 token 直接下载。白名单只留插件真实需要的入口与目录。
    STATIC_ALLOWED_TOP = ('/', '/taskpane.html', '/commands.html', '/support.html')
    STATIC_ALLOWED_ROOTS = ('assets', 'skills')

    def is_static_path_allowed(self, url_path):
        decoded = unquote(url_path)
        if '\x00' in decoded:
            return False
        norm = posixpath.normpath(decoded)
        if norm in self.STATIC_ALLOWED_TOP:
            return True
        parts = [p for p in norm.split('/') if p]
        if not parts or parts[0] not in self.STATIC_ALLOWED_ROOTS:
            return False
        # URL 已规范化仍需核对真实落点：拒绝越界与指向白名单之外的符号链接（realpath）
        root = os.path.dirname(os.path.abspath(__file__))
        base = os.path.realpath(os.path.join(root, parts[0]))
        full = os.path.realpath(os.path.join(root, *parts))
        if not (full == base or full.startswith(base + os.sep)):
            return False
        return os.path.isfile(full)  # 目录一律拒绝，顺带关闭目录列表

    def send_head(self):
        # GET（静态回退分支）与 HEAD 的共同分发点：SimpleHTTPRequestHandler.do_HEAD 只走
        # send_head，白名单放这里才能同时堵住两条入口；动态 /api/* 在 do_GET 更早处已返回。
        if not self.is_static_path_allowed(urlparse(self.path).path):
            data = json.dumps({'error': 'forbidden: static path not allowed'}).encode()
            self.send_response(403)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(data)))
            self.end_headers()
            if self.command != 'HEAD':
                self.wfile.write(data)
            return None
        return super().send_head()

    def serve_taskpane(self):
        # taskpane.html 里的 __LOCAL_API_TOKEN__ 占位符在这里换成启动时生成的真 token
        try:
            with open('taskpane.html', 'r', encoding='utf-8') as fp:
                html = fp.read().replace("'__LOCAL_API_TOKEN__'", repr(API_TOKEN))
        except OSError as e:
            self.send_json(500, {'error': str(e)})
            return
        data = html.encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == '/taskpane.html' or parsed.path == '/':
            self.serve_taskpane()
            return
        if parsed.path.startswith('/api/'):
            if not self.check_api_token(parse_qs(parsed.query)):
                self.send_json(403, {'error': 'forbidden: missing or wrong token'})
                return
        if parsed.path == '/api/models':
            self.proxy_models(parse_qs(parsed.query))
            return
        if parsed.path == '/api/fetch-image':
            self.fetch_image(parse_qs(parsed.query))
            return
        if parsed.path == '/api/extract-images':
            self.extract_images(parse_qs(parsed.query))
            return
        if parsed.path == '/api/settings':
            self.read_settings()
            return
        if parsed.path == '/api/skills':
            self.list_skills()
            return
        if parsed.path == '/api/file':
            self.read_local_file(parse_qs(parsed.query))
            return
        if parsed.path.startswith('/api/test/'):
            self.send_json(404, {'error': 'test endpoints disabled in trial'})
            return
        super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path.startswith('/api/'):
            if not self.check_api_token(parse_qs(parsed.query)):
                self.send_json(403, {'error': 'forbidden: missing or wrong token'})
                return
        if parsed.path == '/api/settings':
            self.write_settings()
            return
        if parsed.path == '/api/chat':
            self.proxy_chat()
            return
        if parsed.path.startswith('/api/test/'):
            self.send_json(404, {'error': 'test endpoints disabled in trial'})
            return
        self.send_json(405, {'error': 'method not allowed'})

    # ---- 【第 11 节】测试桥 API：runner 与插件内 test-bridge.js 的消息面 ----
    # GET  /api/test/register?host&instanceId&documentId&sessionId  插件低频探测+心跳，上报身份
    # GET  /api/test/commands?runId&instanceId      桥拉取下一条命令（定向+串行+超时补投）
    # GET  /api/test/events?runId&since=N           runner 增量拉事件（游标落后于裁剪区报缺口）
    # GET  /api/test/status?runId                   runner/测试查运行摘要（绑定、命令状态、身份核对）
    # POST /api/test/run     {host}                 建运行（单宿主单运行，冲突 409）
    # POST /api/test/bind    {runId,instanceId,documentId,sessionId}  绑定测试目标（需先 attach 成功）
    # POST /api/test/enqueue {runId,command,params,targetInstanceId?,idempotencyKey?}
    # POST /api/test/event   {runId,instanceId,type,payload,commandId?}
    # POST /api/test/stop    {runId}                请求停止（绑定与事件通道保留，等桥确认）
    def handle_test_api(self, parsed):
        q = parse_qs(parsed.query)
        path = parsed.path
        if path == '/api/test/register':
            host = (q.get('host') or [''])[0]
            instance_id = (q.get('instanceId') or [''])[0]
            document_id = (q.get('documentId') or [''])[0][:128]
            session_id = (q.get('sessionId') or [''])[0][:128]
            if host not in ('word', 'excel', 'ppt') or not instance_id:
                self.send_json(400, {'error': 'host/instanceId required'})
                return
            with TEST_LOCK:
                _gc_test_runs()
                active = None
                bound_to_other = False
                for rid, run in TEST_RUNS.items():
                    if run['host'] == host and not run['stopped']:
                        active = rid
                        run['instances'][instance_id] = {'lastSeen': time.time(),
                                                         'documentId': document_id, 'sessionId': session_id}
                        if run['bound'] and run['bound']['instanceId'] != instance_id:
                            bound_to_other = True
                        # 绑定实例的身份变化（换文档）→【S3】实际封锁：记 binding_lost 并终止运行，
                        # 后续命令不再投递（防测试动作落到漂移后的新业务文档上）。
                        # 【W3W5-隔离】仅 sessionId 变化不再封锁：测试运行会受控切换到临时测试会话
                        # （桥 install 注入时迁移绑定会话身份）；非受控的会话漂移仍由桥侧写前复核
                        # （文档标记+会话一致）拦截。文档身份才是防误写的硬边界。
                        if run['bound'] and run['bound']['instanceId'] == instance_id:
                            drifted = bool(run['bound']['documentId'] and document_id and run['bound']['documentId'] != document_id)
                            if drifted and not run['stopped']:
                                run['stopped'] = True
                                _append_test_event(run, 'binding_lost', {
                                    'instanceId': instance_id,
                                    'boundDocumentId': run['bound']['documentId'], 'newDocumentId': document_id,
                                    'boundSessionId': run['bound']['sessionId'], 'newSessionId': session_id
                                }, instance_id)
                            run['bound']['currentDocumentId'] = document_id
                            run['bound']['currentSessionId'] = session_id
                self.send_json(200, {'active': bool(active), 'runId': active, 'boundToOther': bound_to_other})
            return
        if path == '/api/test/commands':
            run_id = (q.get('runId') or [''])[0]
            instance_id = (q.get('instanceId') or [''])[0]
            with TEST_LOCK:
                run = TEST_RUNS.get(run_id)
                if not run or run['stopped']:
                    self.send_json(404, {'error': 'no such test run'})
                    return
                inst = run['instances'].setdefault(instance_id, {'lastSeen': time.time(), 'documentId': '', 'sessionId': ''})
                inst['lastSeen'] = time.time()
                my = [c for c in run['commands'] if c.get('target') in (None, instance_id)]
                # 控制命令（stop/finish）优先：不被普通命令串行阻塞（R3）
                for c in my:
                    if c['state'] == 'queued' and c['command'] in TEST_CONTROL_COMMANDS:
                        c.update(state='delivered', deliveredTo=instance_id, deliveredAt=time.time())
                        self.send_json(200, {'command': c})
                        return
                # 串行：在途（delivered/running）未回执时不出下一条；超时补投不再被串行挡住（R5）
                outstanding = [c for c in my if c['state'] in ('delivered', 'running')]
                if outstanding:
                    c = outstanding[0]
                    if time.time() - c['deliveredAt'] > TEST_REDELIVER_MS / 1000.0:
                        c['deliveredAt'] = time.time()
                        self.send_json(200, {'command': c})
                        return
                    self.send_json(200, {'command': None})
                    return
                for c in my:
                    if c['state'] == 'queued':
                        c.update(state='delivered', deliveredTo=instance_id, deliveredAt=time.time())
                        self.send_json(200, {'command': c})
                        return
            self.send_json(200, {'command': None})  # 无命令（204 不得带响应体，统一 200）
            return
        if path == '/api/test/events':
            run_id = (q.get('runId') or [''])[0]
            try:
                since = int((q.get('since') or ['0'])[0])
            except ValueError:
                since = 0
            with TEST_LOCK:
                run = TEST_RUNS.get(run_id)
                if not run:
                    self.send_json(404, {'error': 'no such test run'})
                    return
                run['runnerLastSeen'] = time.time()
                if since > 0 and run['events'] and since < run['events'][0]['seq'] - 1:
                    # 游标落在被裁剪区：明确报缺口，不让 runner 永远等（R5）
                    self.send_json(200, {'gap': True, 'firstSeq': run['events'][0]['seq'],
                                         'events': [], 'next': run['events'][0]['seq'],
                                         'stopped': run['stopped']})
                    return
                events = [e for e in run['events'] if e['seq'] > since][:200]
                nxt = events[-1]['seq'] if events else max(since, run['eventSeq'])
                self.send_json(200, {'gap': False, 'events': events, 'next': nxt,
                                     'stopped': run['stopped']})
            return
        if path == '/api/test/status':
            run_id = (q.get('runId') or [''])[0]
            with TEST_LOCK:
                run = TEST_RUNS.get(run_id)
                if not run:
                    self.send_json(404, {'error': 'no such test run'})
                    return
                run['runnerLastSeen'] = time.time()
                b = run['bound']
                identity_match = bool(b) and (
                    b.get('currentDocumentId', b['documentId']) == b['documentId']
                    and b.get('currentSessionId', b['sessionId']) == b['sessionId'])
                self.send_json(200, {
                    'host': run['host'], 'bound': b, 'stopped': run['stopped'],
                    'stopRequested': run['stopRequested'], 'identityMatch': identity_match,
                    'eventSeq': run['eventSeq'],
                    'commands': [{'commandId': c['commandId'], 'command': c['command'],
                                  'state': c['state'], 'target': c.get('target')} for c in run['commands']],
                    'instances': {k: {'documentId': v['documentId'], 'sessionId': v['sessionId']}
                                  for k, v in run['instances'].items()}})
            return
        # ---- POST 分支 ----
        try:
            length = int(self.headers.get('Content-Length') or 0)
            body = json.loads(self.rfile.read(length) or b'{}')
        except Exception as e:
            self.send_json(400, {'error': f'bad json: {e}'})
            return
        if path == '/api/test/run':
            host = body.get('host')
            if host not in ('word', 'excel', 'ppt'):
                self.send_json(400, {'error': 'host must be word/excel/ppt'})
                return
            with TEST_LOCK:
                _gc_test_runs()
                for rid, run in TEST_RUNS.items():
                    if run['host'] == host and not run['stopped']:
                        self.send_json(409, {'error': f'{host} already has an active test run', 'runId': rid})
                        return
                rid = secrets.token_hex(8)
                TEST_RUNS[rid] = {'host': host, 'commands': [], 'events': [], 'eventSeq': 0,
                                  'instances': {}, 'bound': None, 'created': time.time(),
                                  'stopped': False, 'stopRequested': False,
                                  'runnerLastSeen': time.time()}
            self.send_json(200, {'runId': rid})
            return
        if path == '/api/test/bind':
            run_id = body.get('runId')
            instance_id = str(body.get('instanceId') or '')[:64]
            document_id = str(body.get('documentId') or '')[:128]
            session_id = str(body.get('sessionId') or '')[:128]
            with TEST_LOCK:
                run = TEST_RUNS.get(run_id)
                if not run or run['stopped']:
                    self.send_json(404, {'error': 'no such test run'})
                    return
                if run['bound']:
                    if run['bound']['instanceId'] == instance_id:
                        self.send_json(200, {'ok': True, 'already': True})
                    else:
                        self.send_json(409, {'error': 'run already bound to another instance'})
                    return
                # 必须有该实例自己的成功 attach_result 且 documentId 一致——attach 失败不留写权限（R1）
                ok_attach = any(e['instanceId'] == instance_id and e['type'] == 'attach_result'
                                and e['payload'].get('ok') and e['payload'].get('documentId') == document_id
                                for e in run['events'])
                if not ok_attach:
                    self.send_json(409, {'error': 'bind requires a successful attach_result from this instance with matching documentId'})
                    return
                run['bound'] = {'instanceId': instance_id, 'documentId': document_id,
                                'sessionId': session_id, 'boundAt': time.time()}
                _append_test_event(run, 'bound', {'instanceId': instance_id,
                                                  'documentId': document_id, 'sessionId': session_id}, instance_id)
            self.send_json(200, {'ok': True})
            return
        if path == '/api/test/enqueue':
            run_id = body.get('runId')
            command = body.get('command')
            if command not in TEST_COMMANDS_ALLOWED:
                self.send_json(400, {'error': 'command not allowed: ' + str(command)})
                return
            idem = str(body.get('idempotencyKey') or '')[:80]
            with TEST_LOCK:
                run = TEST_RUNS.get(run_id)
                if not run or run['stopped']:
                    self.send_json(404, {'error': 'no such test run'})
                    return
                run['runnerLastSeen'] = time.time()
                if idem:
                    for c in run['commands']:
                        if c.get('idempotencyKey') == idem:
                            self.send_json(200, c)  # 同一逻辑命令的 HTTP 重试不再生成第二条（R5）
                            return
                # 【R1】目标解析：显式 targetInstanceId 优先；写命令必须已绑定且指向绑定实例
                target = body.get('targetInstanceId')
                if command in TEST_WRITE_COMMANDS:
                    if not run['bound']:
                        self.send_json(409, {'error': 'write commands require bind first'})
                        return
                    if target and target != run['bound']['instanceId']:
                        self.send_json(409, {'error': 'write commands may only target the bound instance'})
                        return
                    target = run['bound']['instanceId']
                else:
                    if target:
                        if target not in run['instances']:
                            self.send_json(400, {'error': 'unknown targetInstanceId'})
                            return
                    elif run['bound']:
                        target = run['bound']['instanceId']
                    elif command not in TEST_CONTROL_COMMANDS:
                        self.send_json(400, {'error': 'pre-bind commands need explicit targetInstanceId'})
                        return
                if len(run['commands']) >= TEST_MAX_COMMANDS:
                    self.send_json(429, {'error': 'command queue full'})
                    return
                cmd = {'commandId': secrets.token_hex(6), 'command': command,
                       'params': body.get('params') or {}, 'target': target,
                       'state': 'queued', 'idempotencyKey': idem,
                       'deliveredTo': None, 'deliveredAt': 0.0}
                if command in TEST_CONTROL_COMMANDS:
                    run['commands'].insert(0, cmd)  # 控制命令插队
                else:
                    run['commands'].append(cmd)
            self.send_json(200, cmd)
            return
        if path == '/api/test/event':
            run_id = body.get('runId')
            instance_id = str(body.get('instanceId') or '')[:64]
            etype = str(body.get('type') or 'unknown')[:60]
            command_id = str(body.get('commandId') or '')[:64]
            with TEST_LOCK:
                run = TEST_RUNS.get(run_id)
                if not run:
                    self.send_json(404, {'error': 'no such test run'})
                    return
                _append_test_event(run, etype, body.get('payload') or {}, instance_id, command_id)
                if command_id:
                    for cmd in run['commands']:
                        if cmd['commandId'] != command_id:
                            continue
                        # 【R1】回执归属：只有命令实际投递的实例能推进它的状态
                        #（未投递的命令没有任何实例能凭 commandId 结算它）
                        if cmd.get('deliveredTo') != instance_id:
                            break
                        if etype == 'command_started':
                            if cmd['state'] == 'delivered':
                                cmd['state'] = 'running'  # 开始≠完成（R5）
                        elif etype == 'error':
                            if cmd['state'] != 'completed':
                                cmd['state'] = 'failed'
                        elif etype == TEST_SETTLE_EVENT.get(cmd['command']):
                            ok = (body.get('payload') or {}).get('ok')
                            cmd['state'] = 'failed' if ok is False else 'completed'
                        break
                if etype == 'stopped':
                    # 停止确认只认绑定实例（未绑定时任何实例都可能是测试目标）
                    if not run['bound'] or run['bound']['instanceId'] == instance_id:
                        run['stopped'] = True
                if etype == 'stop_unconfirmed':
                    # 停止未在时限内确认：本轮测试终止并如实留痕，不显示「已安全停止」（R3）
                    if not run['bound'] or run['bound']['instanceId'] == instance_id:
                        run['stopped'] = True
                if etype == 'bridge_detached':
                    run['stopped'] = True
            self.send_json(200, {'ok': True})
            return
        if path == '/api/test/stop':
            run_id = body.get('runId')
            resp = {'ok': True, 'found': bool(TEST_RUNS.get(run_id))}
            with TEST_LOCK:
                run = TEST_RUNS.get(run_id)
                if run and not run['stopped']:
                    run['runnerLastSeen'] = time.time()
                    bound = run['bound']
                    if bound:
                        # 【S3/S4】已绑定且本轮拥有任务：定向停止——只投给绑定实例，
                        # 走原 UI 停止流程；身份漂移时命令通道已因 binding_lost 关闭。
                        run['stopRequested'] = True
                        stop_cmd = next((c for c in run['commands'] if c['command'] == 'stop' and c['state'] in ('queued', 'delivered', 'running')), None)
                        if not stop_cmd:
                            stop_cmd = {'commandId': secrets.token_hex(6), 'command': 'stop',
                                        'params': {}, 'target': bound['instanceId'],
                                        'state': 'queued', 'idempotencyKey': '',
                                        'deliveredTo': None, 'deliveredAt': 0.0}
                            run['commands'].insert(0, stop_cmd)  # 优先投递，不被普通队列阻塞（R3）
                            _append_test_event(run, 'stop_requested', {}, 'runner')
                        # 【U1】回传停止命令凭据：runner 据此匹配「有来源的停止终态回执」，
                        # 不再拿 status.stopped 冒充确认（漂移/stop_unconfirmed 下它同样为真）
                        resp['stopCommandId'] = stop_cmd['commandId']
                        resp['targetInstanceId'] = bound['instanceId']
                    else:
                        # 【S3】未绑定的准备运行：直接取消，不广播任何 UI 停止动作——
                        # 没有测试任务可停，业务窗口的停止按钮不能被探针触碰。
                        run['stopped'] = True
                        _append_test_event(run, 'cancelled', {'reason': 'prep run cancelled before bind'}, 'runner')
                        resp['cancelled'] = True
            self.send_json(200, resp)
            return
        self.send_json(404, {'error': 'unknown test endpoint'})

    SETTINGS_FILE = os.path.join(os.environ['OFFICE_AI_TRIAL_DATA'], 'shared-settings.json')
    # 【任务 C】上一份有效设置的备份（受任务 A 静态白名单保护）：仅成功保存前更新，只保留一份
    SETTINGS_BACKUP_FILE = os.path.join(os.environ['OFFICE_AI_TRIAL_DATA'], 'shared-settings.backup.json')
    SKILLS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'skills')
    SKILL_FILE_EXTENSIONS = {'.md', '.txt', '.yaml', '.yml', '.json', '.csv', '.js', '.ts', '.py', '.zip'}

    # 本地文件读取（供插件的 read_local_file 工具）：允许个人文件夹内的常用位置
    # （桌面/文档/下载/工作目录等），但排除隐藏路径和 ~/Library（微信收到的文件除外），防任意敏感文件读取
    HOME_DIR = os.path.realpath(os.path.expanduser('~'))
    WECHAT_DIR = os.path.realpath(os.path.expanduser('~/Library/Containers/com.tencent.xinWeChat'))
    ALLOWED_FILE_EXTENSIONS = {'.txt', '.md', '.csv', '.json', '.pdf', '.docx', '.pptx'}
    MAX_FILE_BYTES = 30 * 1024 * 1024

    def check_path_allowed(self, full):
        if not (full == self.HOME_DIR or full.startswith(self.HOME_DIR + os.sep)):
            return '只允许读取你个人文件夹内的文件'
        # 拒绝路径里任何点开头的隐藏目录/文件（.env、.ssh 之类）
        if any(part.startswith('.') for part in full.split(os.sep) if part):
            return '不允许读取隐藏文件或隐藏目录'
        library = self.HOME_DIR + os.sep + 'Library'
        if (full == library or full.startswith(library + os.sep)) and not full.startswith(self.WECHAT_DIR + os.sep):
            return '系统目录（~/Library）不可读取；微信里收到的文件可以直接读'
        return ''

    def read_local_file(self, q):
        raw = unquote((q.get('path') or [''])[0]).strip()
        if not raw:
            self.send_json(400, {'error': 'missing path'})
            return
        full = os.path.realpath(os.path.expanduser(raw))
        denied = self.check_path_allowed(full)
        if denied:
            self.send_json(403, {'error': denied})
            return
        ext = os.path.splitext(full)[1].lower()
        if ext not in self.ALLOWED_FILE_EXTENSIONS:
            self.send_json(415, {'error': f'暂不支持 {ext or "该"} 格式（支持 PDF/DOCX/PPTX/TXT/MD/CSV/JSON）'})
            return
        try:
            if os.path.getsize(full) > self.MAX_FILE_BYTES:
                self.send_json(413, {'error': '文件超过 30MB，太大了'})
                return
            with open(full, 'rb') as fp:
                data = fp.read()
        except FileNotFoundError:
            self.send_json(404, {'error': '文件不存在：' + full})
            return
        except PermissionError:
            self.send_json(500, {'error': '插件没有该文件夹的访问权限。请改用附件上传，或检查系统隐私设置。'})
            return
        except OSError as e:
            self.send_json(500, {'error': str(e)})
            return
        self.send_response(200)
        self.send_header('Content-Type', 'application/octet-stream')
        self.send_header('X-File-Name', quote(os.path.basename(full)))
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def list_skills(self):
        files = []
        try:
            if os.path.isdir(self.SKILLS_DIR):
                for root, dirs, names in os.walk(self.SKILLS_DIR):
                    dirs[:] = sorted(d for d in dirs if not d.startswith('.') and d != '__MACOSX')
                    for name in sorted(names):
                        if name.startswith('.') or name == 'index.json':
                            continue
                        if os.path.splitext(name)[1].lower() not in self.SKILL_FILE_EXTENSIONS:
                            continue
                        full_path = os.path.join(root, name)
                        relative = os.path.relpath(full_path, self.SKILLS_DIR).replace(os.sep, '/')
                        files.append(relative)
            self.send_json(200, {'files': sorted(files)})
        except Exception as e:
            self.send_json(500, {'error': str(e)})

    @staticmethod
    def _settings_etag(data):
        """【任务 C】设置内容版本标记：磁盘原始字节的 SHA-256，带合法引号，对客户端视为不透明值。"""
        import hashlib
        return '"' + hashlib.sha256(data).hexdigest() + '"'

    def read_settings(self):
        try:
            with open(self.SETTINGS_FILE, 'rb') as fp:
                data = fp.read()
        except FileNotFoundError:
            data = b'{}'  # 空层也定义稳定初始版本：GET 返回它，首次保存走同一套条件检查
        except Exception as e:
            self.send_json(500, {'error': str(e)})
            return
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('ETag', self._settings_etag(data))  # 正文与 ETag 出自同一次读取的同一快照
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    # 【2026-08-31 事故防线】与客户端 state.js 同一道闸的双保险：
    # 拒绝「无任何带 key 服务商的贫瘠配置」覆盖「有带 key 服务商的丰富配置」。
    # 当天 Excel 残留默认态就这样抹掉了全部三应用配置。
    @staticmethod
    def _providers_with_keys(settings):
        try:
            providers = settings.get('serviceProviders') or []
            return sum(1 for p in providers
                       if isinstance(p, dict) and isinstance(p.get('apiKey'), str) and p['apiKey'].strip())
        except Exception:
            return 0

    def _atomic_write_bytes(self, target, data):
        """原子写入：唯一临时文件（pid+线程 id 防并发互踩）+ os.replace。"""
        os.makedirs(os.path.dirname(target), exist_ok=True)
        tmp = '%s.%d.%d.tmp' % (target, os.getpid(), threading.get_ident())
        with open(tmp, 'wb') as fp:
            fp.write(data)
        os.replace(tmp, target)

    def write_settings(self):
        try:
            length = int(self.headers.get('Content-Length') or 0)
            body = self.rfile.read(length)
            incoming = json.loads(body)  # 校验合法 JSON，防写坏
        except Exception as e:
            self.send_json(500, {'error': str(e)})
            return
        # 【任务 C】条件写入合同：缺 If-Match 一律 428，不保留无条件写入后门。
        # 旧版页面没有这个头，被拒是预期行为——防止旧配置静默覆盖新配置；提示重开插件加载新版。
        if_match = (self.headers.get('If-Match') or '').strip()
        if not if_match:
            self.send_json(428, {'error': 'missing If-Match; reload the add-in to load the new version',
                                 'code': 'settings_precondition_required'})
            return
        with SETTINGS_LOCK:
            try:
                with open(self.SETTINGS_FILE, 'rb') as fp:
                    current_bytes = fp.read()
            except FileNotFoundError:
                current_bytes = b'{}'
            except Exception as e:
                self.send_json(500, {'error': str(e)})
                return
            # 版本核对：客户端基线与磁盘现状不符即冲突，共享文件原样不动
            if if_match != self._settings_etag(current_bytes):
                self.send_json(409, {'error': 'settings were updated by another Office app; this save was not applied',
                                     'code': 'settings_conflict'})
                return
            try:
                current = json.loads(current_bytes)
            except Exception:
                current = None  # 损坏文件按「现状未知」放行本次已过版本核对的写入，顺便自愈
            if (current is not None
                    and self._providers_with_keys(incoming) == 0
                    and self._providers_with_keys(current) > 0):
                self.send_json(409, {'error': 'refused: overwrite would drop all configured providers',
                                     'hint': '本地配置为空（无任何带 key 的服务商），疑似残留默认态；已拒绝覆盖共享配置',
                                     'code': 'settings_barren'})
                return
            # 备份上一份有效版本：仅成功路径更新，拒绝/坏 JSON/网络失败都不动它
            if current is not None:
                try:
                    self._atomic_write_bytes(self.SETTINGS_BACKUP_FILE, current_bytes)
                except Exception:
                    pass  # 备份失败不阻塞保存本身
            try:
                self._atomic_write_bytes(self.SETTINGS_FILE, body)
            except Exception as e:
                self.send_json(500, {'error': str(e)})
                return
        self.send_json(200, {'ok': True, 'etag': self._settings_etag(body)})

    def extract_images(self, q):
        """【图片链路】网页 -> 图片直链列表。web_search 找到文章页后，模型用它拿到页内图片 URL，
        再交给 insert_image(url)。白名单 http/https、限 text/html、限 2MB、最多返回 20 条。"""
        from urllib.parse import urljoin
        import re as _re
        url = unquote((q.get('url') or [''])[0]).strip()
        if not url:
            self.send_json(400, {'error': 'missing url'})
            return
        ok, reason = _target_host_allowed(url)
        if not ok:
            self.send_json(400, {'error': reason})
            return
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Macintosh) TrojanAI-Office/1.0'})
        try:
            with urllib.request.urlopen(req, timeout=20) as r:
                ctype = (r.headers.get('Content-Type') or '').split(';')[0].strip().lower()
                if ctype and ctype != 'text/html':
                    self.send_json(415, {'error': f'not an html page (content-type {ctype})'})
                    return
                html = r.read(2 * 1024 * 1024 + 1)
                if len(html) > 2 * 1024 * 1024:
                    self.send_json(413, {'error': 'page over 2MB'})
                    return
                charset = 'utf-8'
                m = _re.search(rb'charset=["\']?([\w-]+)', r.headers.get('Content-Type', '').encode() + html[:2048])
                if m:
                    try: charset = m.group(1).decode()
                    except Exception: pass
                text = html.decode(charset, 'replace')
            # 提取 <img src>，补全相对地址，过滤 data:/js 伪协议，绝对化后去重
            seen, images = set(), []
            for m in _re.finditer(r'<img[^>]+src=["\']([^"\']+)["\']', text, _re.I):
                src = m.group(1).strip()
                if src.startswith(('data:', 'javascript:')) or src.startswith('#'):
                    continue
                absolute = urljoin(url, src)
                if absolute in seen:
                    continue
                seen.add(absolute)
                images.append(absolute)
                if len(images) >= 20:
                    break
            self.send_json(200, {'count': len(images), 'images': images})
        except Exception as e:
            code = getattr(e, 'code', 502)
            self.send_json(code if isinstance(code, int) and code >= 400 else 502, {'error': str(e)[:300]})

    def fetch_image(self, q):
        """【图片链路】下载外网图片转 base64（WKWebView 直连外网图片会被 CORS 拦，只能走本机）。
        白名单 http/https、content-type 必须 image/*、上限 5MB。带 token 校验（与其他 /api/* 一致）。"""
        url = unquote((q.get('url') or [''])[0]).strip()
        if not url:
            self.send_json(400, {'error': 'missing url'})
            return
        ok, reason = _target_host_allowed(url)
        if not ok:
            self.send_json(400, {'error': reason})
            return
        req = urllib.request.Request(url, headers={'User-Agent': 'TrojanAI-Office/1.0'})
        try:
            with urllib.request.urlopen(req, timeout=20) as r:
                ctype = (r.headers.get('Content-Type') or '').split(';')[0].strip().lower()
                if not ctype.startswith('image/'):
                    self.send_json(415, {'error': f'not an image (content-type {ctype or "unknown"})'})
                    return
                data = r.read(5 * 1024 * 1024 + 1)
                if len(data) > 5 * 1024 * 1024:
                    self.send_json(413, {'error': 'image over 5MB'})
                    return
            import base64 as b64mod
            self.send_json(200, {'base64': b64mod.b64encode(data).decode(), 'contentType': ctype, 'size': len(data)})
        except Exception as e:
            code = getattr(e, 'code', 502)
            self.send_json(code if isinstance(code, int) and code >= 400 else 502, {'error': str(e)[:300]})

    def proxy_models(self, q):
        base = unquote((q.get('base') or [''])[0]).rstrip('/')
        key = unquote((q.get('key') or [''])[0])
        protocol = (q.get('protocol') or ['openai'])[0]
        if not base:
            self.send_json(400, {'error': 'missing base'})
            return
        if protocol == 'anthropic':
            # Anthropic Messages 协议：模型列表在 /v1/models，鉴权用 x-api-key
            url = (base if base.endswith('/v1') else base + '/v1') + '/models'
            req = urllib.request.Request(url, headers={
                'x-api-key': key,
                'anthropic-version': '2023-06-01',
                'Authorization': f'Bearer {key}',
                'User-Agent': 'TrojanAI-Office/1.0'
            })
        else:
            url = base + '/models'
            req = urllib.request.Request(url, headers={
                'Authorization': f'Bearer {key}',
                'User-Agent': 'TrojanAI-Office/1.0'
            })
        try:
            with urllib.request.urlopen(req, timeout=20) as r:
                data = r.read()
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(data)
        except Exception as e:
            code = getattr(e, 'code', 502)
            body = ''
            try:
                body = e.read().decode()[:300]
            except Exception:
                body = str(e)[:300]
            self.send_json(code if isinstance(code, int) and code >= 400 else 502, {'error': body or str(e)})

    def proxy_chat(self):
        # 【v142】对话流式代理：浏览器 taskpane 受 CORS 限制无法直连部分上游（如
        # Anthropic 官方端点、未开 CORS 的中转），由本机 server 转发。SSE 逐块透传。
        # 安全模型：仅本机监听 + /api/* token 鉴权（与 settings 同级）。
        try:
            length = int(self.headers.get('Content-Length') or 0)
            payload = json.loads(self.rfile.read(length))
        except Exception as e:
            self.send_json(400, {'error': f'bad request body: {e}'})
            return
        url = str(payload.get('url') or '')
        if not (url.startswith('http://') or url.startswith('https://')):
            self.send_json(400, {'error': 'url must be http(s)'})
            return
        headers = dict(payload.get('headers') or {})
        body = payload.get('body')
        data = json.dumps(body).encode() if not isinstance(body, (bytes, str)) else (
            body.encode() if isinstance(body, str) else json.dumps(body).encode())
        if 'Content-Type' not in {k.title() for k in headers}:
            headers['Content-Type'] = 'application/json'
        req = urllib.request.Request(url, data=data, headers=headers, method='POST')
        try:
            upstream = urllib.request.urlopen(req, timeout=300)
        except Exception as e:
            code = getattr(e, 'code', 502)
            detail = ''
            try:
                detail = e.read().decode()[:500]
            except Exception:
                detail = str(e)[:500]
            self.send_json(code if isinstance(code, int) and code >= 400 else 502, {'error': detail or str(e)})
            return
        try:
            self.send_response(upstream.status)
            ctype = upstream.headers.get('Content-Type') or 'text/event-stream'
            self.send_header('Content-Type', ctype)
            # SSE 流式：不设 Content-Length，逐块写，靠连接关闭表示结束
            self.end_headers()
            # 【任务 B】read1(n) 是「最多读 n 字节，有多少读多少」，不等攒满：
            # 旧 read(1024) 在分块上游上会等够 1024 字节或连接结束才返回，
            # 已到达的小 SSE 事件被扣住，模型首字显示延迟。逐块即读即写即 flush。
            read1 = getattr(upstream, 'read1', None)
            while True:
                chunk = read1(8192) if read1 else upstream.read(1)
                if not chunk:
                    break
                self.wfile.write(chunk)
                self.wfile.flush()
        except Exception as e:
            # 头已发出后无法再报错，只能断流（客户端读流会得到 done/error）
            self.log_message('proxy_chat stream broken: %s', e)
        finally:
            try:
                upstream.close()
            except Exception:
                pass

    def send_json(self, code, obj):
        data = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, *a):
        pass

if __name__ == '__main__':
    import sys
    directory = sys.argv[1] if len(sys.argv) > 1 else '.'
    port = int(sys.argv[2]) if len(sys.argv) > 2 else 8080  # 测试用随机端口注入；生产不传保持 8080
    import os
    os.chdir(directory)
    print('LISTENING', port, flush=True)  # 测试据此判断就绪
    ThreadingHTTPServer(('127.0.0.1', port), Handler).serve_forever()

# 从源码构建（开发者）

普通使用者直接安装 Mac `.pkg` 即可。安装包已包含 Python 运行环境，无需执行以下步骤。

以下仅供需要自行构建安装包的开发者使用。构建环境为 Mac，需要 Python 3 和 macOS 命令行工具。

1. 从 [Python 官方下载页](https://www.python.org/downloads/release/python-31315/) 下载 macOS installer，保存为 `build-cache/python-mac.pkg`。脚本会核对文件校验值。
2. 运行 `pkgutil --expand-full build-cache/python-mac.pkg build-cache/python-mac-expanded`，解开运行环境。
3. 使用空的 `staging` 目录，运行 `python3 tools/package.py`。Mac 安装包输出到 `releases`。
4. 运行 `python3 -m unittest discover -s tests`，并在实际 Office 中验证安装、使用和卸载。

## PPT / Excel 恢复逻辑回归

安装 Node.js 22 或更新版本后，运行 `node --test tests/*.test.js`。这些测试使用仿真的 Office 对象和脚本化模型回复，覆盖锚点失配、重新读取后恢复及重复失败停止；不需要真实模型密钥。

离线通过不替代实际 Office 验收。合并本轮恢复修复前，还需在试用版副本上核对 PowerPoint/Excel 的恢复、停止、重连，并复查 Word 共享循环行为。安装包需另行构建和验证。

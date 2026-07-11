# Electron桌面背单词工具

## 启动方式

1. 使用 VSCode 打开项目根目录。
2. 在终端执行 `npm install` 安装依赖。
3. 在终端执行 `npm start` 启动 Electron 客户端。

## uv 安装入口

本项目当前运行链路仍是 Node.js / Electron；`uv` 作为本机 Python 工具管理器引入到工程辅助脚本中，便于后续增加 Python 脚本或本地自动化时统一安装和校验。

仅检查当前机器是否已安装 `uv`：

```powershell
powershell -ExecutionPolicy Bypass -File .\setup-uv.ps1
```

按官方推荐的 Windows 包管理方式安装 `uv`：

```powershell
powershell -ExecutionPolicy Bypass -File .\setup-uv.ps1 -Install
```

如果本机没有 `winget`，可改用官方独立安装脚本：

```powershell
powershell -ExecutionPolicy Bypass -File .\setup-uv.ps1 -Install -Method standalone
```

如果当前网络对 GitHub 下载不稳定，也可以使用官方 PyPI 包安装：

```powershell
powershell -ExecutionPolicy Bypass -File .\setup-uv.ps1 -Install -Method pip
```

安装后可用 `uv --version` 校验；若刚安装后当前终端找不到 `uv`，重新打开终端即可。

## 无npm本地启动方式

当 npm 依赖源无法访问时，可使用本地浏览器版：

```powershell
powershell -ExecutionPolicy Bypass -File .\start-local.ps1
```

该入口启动 `http://127.0.0.1:3765/`，首次启动不提供默认词库，需要在页面点击 `选择xlsx文件` 选择本地 Excel。
选择后的文件会缓存到 `.word-memory-cache`，下次启动继续使用上一次选择的词库。
由于 `.xlsx` 文件不会上传 GitHub，从 GitHub 下载代码后第一次启动也会进入选择文件状态。

如确实需要从命令行指定某个 Excel，可传入 `-ExcelPath`：

```powershell
powershell -ExecutionPolicy Bypass -File .\start-local.ps1 -ExcelPath .\Vocabulary.xlsx
```

## Excel文件要求

仅支持 `.xlsx` 文件。Excel 标准结构详见 `REQUIREMENTS.md`。

两工作表格式不强制工作表名称完全一致。程序会优先找两张有数据的表，列数较少的按总表逻辑解析，列数较多的按汇总表逻辑解析。
字段名也会模糊匹配；若字段名无法识别，则按下面的位置兜底：

- 总表逻辑：第 1 列作为单元，第 3 列作为英文单词，第 4 列作为组合释义。
- 汇总表逻辑：第 1 列作为单元，第 3 列作为英文单词，第 4/5/6 列作为音标/词性/中文意思。

单词按英文文本去重合并。两张表同时存在同一英文单词时，展示信息优先使用汇总表逻辑里的音标、词性、中文意思。`Page` 和 `总表里有` 这类辅助列不参与单词字段解析。

单工作表格式：

- 表头包含：`Day`、`序号`、`单词`、`中文意思`
- `Day` 作为单元，`单词` 作为英文，`中文意思` 作为组合释义；若组合释义里包含音标或 `n.`、`v.`、`adj.` 等词性标记，会自动拆到音标/词性/中文意思。

## 本地数据

本地缓存保存在项目目录下的 `.word-memory-cache` 文件夹，包含上次选择的 Excel 路径、权重、知道/不知道记录、首次抽取状态和上次随机数量。程序重启后这些数据会保留。

## 当前环境说明

当前本机已可通过 Node.js 运行无 npm 本地版。Electron 版本仍需要 npm 依赖可下载后再启动。

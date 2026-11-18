# PhotoExchangeFormula

Windows 桌面工具：从图片识别数学公式，输出 LaTeX，并可复制图片或导出可编辑的 Word 公式。

## 运行

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe gui_version.py
```

也可以在安装依赖后双击 `run_gui.bat`。识别引擎使用 [pix2tex](https://github.com/lukas-blecher/LaTeX-OCR)；首次识别需要下载模型，缓存保存在程序目录的 `cache/` 中。

选择图片后可框选公式区域、修改识别出的 LaTeX，再复制 LaTeX、复制原图或保存 Word 文档。导出原生 Word 公式需要本机安装 Microsoft Office；转换失败时请使用 LaTeX 手动粘贴。

## 许可证

MIT。详见 [LICENSE](LICENSE)。

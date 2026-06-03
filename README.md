# extract-to-md

中文 | [English](README.en.md)

`extract-to-md` 是一个本地优先的文档转 Markdown 命令行工具，适合把 PDF、Office、图片、表格等文件整理成方便 LLM / Agent 阅读的 Markdown。

它可以在可用时调用 [Microsoft MarkItDown](https://github.com/microsoft/markitdown)，同时增加了更适合本地文档处理的兜底逻辑：PDF 文本层抽取、扫描件 OCR、图片 OCR、PPT 渲染后 OCR、Excel 行数限制等。

## 支持格式

| 类型 | 支持情况 | 需要的外部工具 |
| --- | --- | --- |
| `.md` / `.txt` / `.csv` / `.json` / `.xml` / `.html` / `.yaml` | 直接读取文本 | 无 |
| `.docx` | 提取正文和表格；可优先使用 MarkItDown | 可选 `markitdown` |
| `.xlsx` | 提取工作表为 Markdown 表格 | 无 |
| `.pptx` | 提取真实文本和演讲者备注 | 无 |
| 图片型 `.pptx` | 转 PDF 后 OCR | `soffice`、`pdftoppm`、`tesseract` |
| `.pdf` 有文本层 | 使用 `pdftotext -layout` 保留布局 | Poppler 的 `pdftotext` |
| 扫描版 `.pdf` | 渲染每页图片后 OCR | Poppler 的 `pdftoppm`、`tesseract` |
| `.png` / `.jpg` / `.jpeg` / `.webp` / `.bmp` / `.tif` / `.tiff` | OCR 图片文字 | `tesseract` |
| 其他 MarkItDown 支持的格式 | 作为最后兜底尝试 | 可选 `markitdown` |

默认 OCR 语言是 `chi_sim+eng`，也就是简体中文 + 英文。

## 快速安装

推荐使用 `uv tool install` 从 GitHub 安装：

```powershell
uv tool install "extract-to-md[markitdown] @ git+https://github.com/byezero/extract-to-md.git"
```

安装后确认命令可用：

```powershell
extract-to-md --help
```

如果你已经 clone 了这个仓库，也可以在仓库目录中安装：

```powershell
uv tool install ".[markitdown]"
```

不需要 MarkItDown 兜底时，可以安装最小版本：

```powershell
uv tool install git+https://github.com/byezero/extract-to-md.git
```

## Windows 安装外部工具

在 Windows 上推荐用 `winget` 安装依赖：

```powershell
winget install -e --id Python.Python.3.12
winget install -e --id astral-sh.uv
winget install -e --id oschwartz10612.Poppler
winget install -e --id UB-Mannheim.TesseractOCR
winget install -e --id TheDocumentFoundation.LibreOffice
```

安装后新开一个 PowerShell，确认这些命令都能找到：

```powershell
python --version
uv --version
pdftotext -v
pdftoppm -v
tesseract --version
soffice --version
```

如果 `tesseract` 报 `chi_sim` 语言不存在，需要安装简体中文语言数据。常见做法是把 `chi_sim.traineddata` 放到 Tesseract 的 `tessdata` 目录，然后再运行：

```powershell
tesseract --list-langs
```

列表里应该能看到：

```text
chi_sim
eng
```

## macOS 安装外部工具

```bash
brew install python uv poppler tesseract libreoffice
brew install tesseract-lang
```

然后确认：

```bash
python3 --version
uv --version
pdftotext -v
pdftoppm -v
tesseract --version
soffice --version
```

## Linux 安装外部工具

Ubuntu / Debian 示例：

```bash
sudo apt update
sudo apt install -y python3 python3-pip poppler-utils tesseract-ocr tesseract-ocr-chi-sim libreoffice
curl -LsSf https://astral.sh/uv/install.sh | sh
```

然后确认：

```bash
python3 --version
uv --version
pdftotext -v
pdftoppm -v
tesseract --version
soffice --version
```

## 使用方法

转换文件并输出到终端：

```powershell
extract-to-md input.pdf
```

转换文件并写入 Markdown：

```powershell
extract-to-md input.pdf -o output.md
```

强制 OCR 扫描版 PDF：

```powershell
extract-to-md scanned.pdf --force-ocr -o scanned.md
```

只识别英文图片：

```powershell
extract-to-md image.png --lang eng -o image.md
```

处理版式比较散的 PPT：

```powershell
extract-to-md deck.pptx --psm 11 -o deck.md
```

限制 Excel 每个工作表最多读取 200 行：

```powershell
extract-to-md workbook.xlsx --max-rows-per-sheet 200 -o workbook.md
```

## 常见问题

### 和 MarkItDown 是什么关系？

MarkItDown 是微软的通用文档转 Markdown 工具。`extract-to-md` 会在适合的时候调用它，但不会完全依赖它。

这个工具额外做了几件事：

- PDF 优先用 `pdftotext -layout` 抽文本，文本太少再 OCR
- 扫描版 PDF 可以自动或强制 OCR
- 图片直接用 Tesseract OCR
- 图片型 PPTX 会先用 LibreOffice 转成 PDF，再逐页 OCR
- Excel 会限制读取行数，避免输出过大
- 所有输出统一整理成 Markdown

### 为什么 PDF 必须安装 Poppler？

当前 PDF 路径直接使用 `pdftotext` / `pdftoppm`，这样可控、离线、布局保留比较稳定。没有 Poppler 时，PDF 转换会失败。

### 为什么扫描件识别不准？

OCR 准确率取决于图片清晰度、语言包、DPI 和页面布局。可以尝试：

```powershell
extract-to-md scanned.pdf --force-ocr --dpi 400 --psm 6 -o scanned.md
extract-to-md scanned.pdf --force-ocr --dpi 300 --psm 11 -o scanned.md
```

### 输出 Markdown 开头的注释是什么？

输出会带一行来源注释：

```html
<!-- extracted-from: input.pdf -->
```

这是为了让 Agent 或后续处理流程知道 Markdown 来自哪个文件。

## 开发

克隆仓库：

```powershell
git clone https://github.com/byezero/extract-to-md.git
cd extract-to-md
```

本地运行：

```powershell
python src\extract_to_md\cli.py --help
```

构建包：

```powershell
uv build
```

## License

MIT

# DocReplacer (文档批量替换工具)

[English Version](#english-version) | [中文版](#中文版)

---

## <a name="中文版"></a>中文版

**DocReplacer** 是一个简单易用的桌面应用程序，旨在帮助用户依据自定义的规则，一键批量替换 Word 文档（.doc, .docx）中的文本词汇。

### 🌟 核心功能
* **批量文档处理**：支持一次性导入多个 `.doc` 或 `.docx` 文件，快速完成指定内容的统一替换。
* **自定义规则导入**：支持导入 `.xlsx` 格式的替换规则表格（需包含“原词”和“替换词”两列）。
* **内置规则编辑器**：随时在软件界面内直接添加、修改、或者删除替换规则，并且可以全选/反选特定的生效规则。
* **规则导出备份**：支持将界面上配置好的当前规则重新导出为 Excel 表格，方便未来重复使用。
* **安全替换引擎**：在执行替换时，会自动备份原文档，并在需要时提供“一键撤销（还原）”功能，保证文档安全。
* **跨平台支持**：支持打包为 Windows (`.exe`)、macOS (`.app`) 以及 Ubuntu Linux 系统的可执行程序。

### 🚀 如何使用
1. **下载软件**：在 GitHub 的 [Actions](https://github.com/Shengda21/DocReplacer/actions) 页面下载适合您系统的最新打包版本。
2. **导入或添加规则**：您可以点击“导入规则”加载准备好的 Excel 表格，或者点击“新建规则”手动输入需要替换的词语。
3. **选择要处理的文档**：点击“选择文件”导入您需要进行文字替换的 Word 文件夹或单个文件。
4. **一键替换**：点击“开始替换”，程序会自动在新文件中生成结果。如果觉得替换效果不满意，您可以点击“还原”。

---

## <a name="english-version"></a>English Version

**DocReplacer** is an easy-to-use desktop application designed to help users batch replace text words in Word documents (.doc, .docx) based on customized rules with a single click.

### 🌟 Core Features
* **Batch Document Processing**: Supports importing multiple `.doc` or `.docx` files at once to quickly complete uniform content replacement.
* **Custom Rule Import**: Supports importing replacement rule tables in `.xlsx` format (must contain "Original Word" and "Replacement Word" columns).
* **Built-in Rule Editor**: Add, modify, or delete replacement rules directly within the software interface at any time, with options to select/deselect specific active rules.
* **Rule Export & Backup**: Supports exporting currently configured rules from the interface back to an Excel table for convenient future reuse.
* **Safe Replacement Engine**: Automatically backs up original documents when performing replacements, offering a "One-Click Undo (Restore)" function to ensure document safety.
* **Cross-Platform Support**: Supports packaging into executable programs for Windows (`.exe`), macOS (`.app`), and Ubuntu Linux systems.

### 🚀 How to Use
1. **Download Software**: Download the latest packaged version for your system from the GitHub [Actions](https://github.com/Shengda21/DocReplacer/actions) page.
2. **Import or Add Rules**: Click "Import Rules" to load a prepared Excel table, or click "New Rule" to manually enter words to be replaced.
3. **Select Documents to Process**: Click "Select Files" to import the Word folder or individual files you need to modify.
4. **One-Click Replace**: Click "Start Replace", and the program will automatically generate results in new files. If unsatisfied, you can click "Restore".

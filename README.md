# Excel 效率工具台

基于 Python、Tkinter 和 pywin32 构建的轻量级 Excel 操作工具。

## 功能特性
- **安全低风险操作**：工具专注处理 Excel 本地文件和正在运行的 Excel 实例，不包含删除源文件、重命名源工作表及网络相关的高风险操作。
- **获取工作表列表**：一键获取当前活动工作簿中的所有工作表名称，并填入活动工作表的 A 列（从 A1 开始，A1为表头“工作表名称”）。
- **一簿按工作表拆分为多个文件**：选择一个源工作簿和输出目录后，按工作表分别导出为独立 `.xlsx` 文件。该功能需要 Windows + Microsoft Excel，源工作簿只读打开，不保存、不覆盖、不删除源文件；输出目录中已有同名文件时会自动追加 `_2`、`_3` 等序号。

## 技术栈
- 语言：Python 3
- GUI：Tkinter
- 核心库：pywin32 (win32com)

## 环境准备与依赖安装
本程序依赖 Windows 操作系统且需提前安装 Microsoft Excel。

```bash
pip install -r requirements.txt
```

## 运行方式
```bash
python -m src.excel_efficiency_toolkit.app
```

## 运行测试
本项目包含了不依赖真实 Excel 的单元测试（采用 Mock 机制）。
```bash
pytest tests/
```

## 手工验收步骤：获取工作表列表 (Windows + Excel)
1. 启动 Windows 下的 Microsoft Excel。
2. 新建或打开任意一个包含多个工作表的工作簿。
3. 确保目标工作簿是当前活跃的窗口。
4. 运行本工具台：`python -m src.excel_efficiency_toolkit.app`。
5. 在界面上点击【列出当前工作簿所有工作表】按钮。
6. 查看 Excel 的当前活动工作表，A 列应该已被填入“工作表名称”以及对应的工作表列表。
7. **注意事项**：该操作不执行保存动作，也不会关闭您的 Excel 文件。

## 手工验收步骤：一簿按工作表拆分为多个文件 (Windows + Excel)
1. 准备一个包含多个工作表的 Excel 文件，例如 `源文件.xlsx`。
2. 准备一个空输出目录；也可以预先放入一个同名 `.xlsx` 文件，用于验证自动追加序号。
3. 运行本工具台：`python -m src.excel_efficiency_toolkit.app`。
4. 点击【一簿按工作表拆分为多个文件】按钮。
5. 在文件选择框中选择源 Excel 文件。
6. 在文件夹选择框中选择输出目录。
7. 查看运行日志，应显示源文件、输出目录以及成功导出的文件数量。
8. 查看输出目录，每个工作表应各自生成一个 `.xlsx` 文件；如果已有同名文件，新文件名应自动追加 `_2`、`_3` 等序号。
9. 确认源工作簿未被保存、覆盖或删除。

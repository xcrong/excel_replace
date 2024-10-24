# batch-modify-excel

## 项目简介

`batch-modify-excel` 是一个用于批量清理 Excel 文件的工具。用户可以选择单个 Excel 文件或包含多个 Excel 文件的文件夹，输入要清理的列和行，以及要清理的字符。该工具会根据用户的设置处理 Excel 文件，并提供处理进度反馈。

## 功能特性

- 支持选择单个 Excel 文件或文件夹。
- 用户可以指定要清理的列和行。
- 支持清理指定字符。
- 处理进度实时反馈。
- 支持覆盖原文件或保存为新文件。

## 安装

1. 确保您已安装 Python 3.8 或更高版本。
2. 克隆此仓库：

   ```bash
   git clone https://github.com/yourusername/batch-modify-excel.git
   cd batch-modify-excel
   ```

3. 安装依赖项：

   ```bash
   pip install -r requirements.txt
   ```

## 使用方法

1. 运行程序：

   ```bash
   python -m batch_modify_excel
   ```

2. 在界面中选择要清理的 Excel 文件或文件夹。
3. 输入要清理的列（如 A B C）和行（如 1 2 3）。
4. 输入要清理的字符（如 "（ 免费）"）。
5. 点击“执行清理”按钮开始处理。

## 配置

程序的配置文件为 `config.json`，您可以在其中设置默认值，例如：

```json
{
    "columns": "A",
    "rows": "",
    "clean_char": "（ 免费）",
    "select_folder": true,
    "overwrite": true,
    "file_path": "",
    "last_file_path": "D:/space/rust/fastoffice/data/总"
}
```

## 贡献

欢迎任何形式的贡献！请提交问题或拉取请求。

## 许可证

此项目采用 MIT 许可证，详细信息请查看 LICENSE 文件。

## 联系信息

如有任何问题，请联系作者：

- 邮箱：zlz_gty@foxmail.com

---

请根据您的具体需求和项目情况进行调整。如果您有其他特定内容需要添加，请告诉我！
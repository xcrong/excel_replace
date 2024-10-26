# Excel 批量替换 

## 项目简介

`Excel 批量替换` 是一个用于批量处理 Excel 文件的工具。用户可以选择单个 Excel 文件或包含多个 Excel 文件的文件夹，输入要处理的列和行，以及要处理的字符。该工具会根据用户的设置处理 Excel 文件，并提供处理进度反馈。

## 功能特性

- 支持选择单个 Excel 文件或文件夹。
- 用户可以指定要处理的列和行。
- 支持处理指定字符。
- 处理进度实时反馈。
- 支持覆盖原文件或保存为新文件。

## 安装

1. 确保您已安装 Python 3.8 或更高版本。
2. 克隆此仓库：

   ```bash
   git clone https://github.com/xcrong/excel_replace.git
   cd excel_replace
   ```

3. 安装依赖项：

   ```bash
   rye sync # rye 工具官网为： https://rye.astral.sh/
   ```

## 使用方法

1. 运行程序：

   ```bash
   rye run test
   ```

2. 在界面中选择要处理的 Excel 文件或文件夹。
3. 输入要处理的列（如 A B C）和行（如 1 2 3）。
4. 输入要处理的字符（如 "（ 免费）"）。
5. 点击“执行处理”按钮开始处理。


## 贡献

欢迎任何形式的贡献！请提交问题或拉取请求。

## 许可证

此项目采用 MIT 2.0  许可证，详细信息请查看 LICENSE 文件。

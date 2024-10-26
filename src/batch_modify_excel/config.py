class Config:
    # 配置文件路径
    CONFIG_FILE = "config.json"
    # 默认处理字符
    DEFAULT_target_char = "（ 免费）"
    DEFAULT_replace_char = ""
    # 分隔符列表
    SEPARATORS = [",", "，", " ", "、", ", "]
    # 项目信息
    PROJECT_INFO = """Excel 批量处理工具
版本：0.1.0
这是一个用于批量处理Excel文件的工具。它可以帮助您快速地处理Excel文件中的指定字符，提高工作效率。"""
    # 作者信息
    AUTHOR_INFO = """如果您有任何问题或建议，请随时与我联系。
    xcrong: zlz_gty@foxmail.com
    """
    # 使用说明
    INSTRUCTIONS = """1. 选择要处理的Excel文件或文件夹
2. 输入要处理的列和行（可用英文逗号、中文逗号、顿号或空格分隔）
3. 输入要处理的字符
4. 点击'执行处理'按钮开始处理
5. 根据需要，选择是否覆盖原文件
6. 等待处理完成，查看处理结果"""
    # 窗口标题
    WINDOW_TITLE = "Excel 批量处理工具"
    # 关于菜单标签
    ABOUT_MENU_LABEL = "关于"
    # 项目信息标签
    PROJECT_INFO_LABEL = "项目信息"
    # 作者信息标签
    AUTHOR_INFO_LABEL = "作者信息"
    # 使用说明标签
    INSTRUCTIONS_LABEL = "使用说明"
    # 选择Excel文件按钮文本
    SELECT_EXCEL_BUTTON_TEXT = "选择Excel文件"
    # 选择文件夹按钮文本
    SELECT_FOLDER_BUTTON_TEXT = "选择含Excel的文件夹"
    # 覆盖原文件复选框文本
    OVERWRITE_CHECKBOX_TEXT = "覆盖原文件"
    # 执行处理按钮文本
    EXECUTE_BUTTON_TEXT = "执行处理"
    # 已选择文件夹消息
    FOLDER_SELECTED_MESSAGE = "已选择文件夹: {}"
    # 已选择文件消息
    FILE_SELECTED_MESSAGE = "已选择文件: {}"
    # 处理完成消息
    PROCESS_COMPLETE_MESSAGE = "文件夹内所有Excel文件处理完毕！"
    # 单个文件处理完成消息
    SINGLE_FILE_COMPLETE_MESSAGE = "文件处理完毕！"
    # 警告消息
    WARNING_MESSAGE = "警告"
    # 选择文件警告
    SELECT_FILE_WARNING = "请先选择一个文件或文件夹"
    # 输入错误消息
    INPUT_ERROR_MESSAGE = "输入错误"
    # 行号错误消息
    ROW_NUMBER_ERROR = "行号必须是数字。"
    # 配置保存错误消息
    CONFIG_SAVE_ERROR = "保存配置时出错：{}"
    # 文件保存消息
    FILE_SAVED_MESSAGE = "文件已保存至 {}"
    # 进度标签文本
    PROGRESS_LABEL_TEXT = "处理进度："

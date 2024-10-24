# ui.py
import tkinter as tk  # 导入tkinter库
from tkinter import (
    filedialog,
    messagebox,
)  # 导入文件对话框和消息框
import ttkbootstrap as ttk  # 导入ttkbootstrap库
from ttkbootstrap.constants import *  # type: ignore # 导入ttkbootstrap库的常量  # noqa: F403

from batch_modify_excel.config import Config  # 导入配置文件


class ExcelCleanerUI:  # 定义ExcelCleanerUI类
    def __init__(self, root, logic):  # 初始化方法
        self.root = root  # 设置根窗口
        self.logic = logic  # 设置逻辑处理
        self.root.title(
            Config.WINDOW_TITLE
        )  # 设置窗口标题为配置文件中的窗口标题

        # UI components  # 定义UI组件
        self.columns_entry = None  # 列输入框
        self.rows_entry = None  # 行输入框
        self.target_char_entry = None  # 处理字符输入框
        self.file_path_entry = None  # 文件路径输入框
        self.progress_label = None  # 进度标签
        self.execute_button = None  # 执行按钮
        self.select_folder_var = (
            tk.BooleanVar()
        )  # 选择文件夹变量
        self.overwrite_var = tk.BooleanVar()  # 盖变量
        self.select_button = (
            None  # 文件选择按钮  # 添加这行来存储按钮引用
        )

        self.setup_menu()  # 设置菜单
        self.setup_ui()  # 设置UI

    def setup_menu(self):  # 设置菜单方法
        self.menubar = ttk.Menu(self.root)  # 创建菜单栏
        self.root.config(
            menu=self.menubar
        )  # 设置根窗口的菜单栏

        self.about_menu = ttk.Menu(
            self.menubar, tearoff=0
        )  # 创建关于菜单
        self.menubar.add_cascade(
            label=Config.ABOUT_MENU_LABEL,
            menu=self.about_menu,
        )  # 添加关于菜单到菜单栏
        self.about_menu.add_command(
            label=Config.PROJECT_INFO_LABEL,
            command=self.show_project_info,
        )  # 添加项目信息到关于菜单
        self.about_menu.add_command(
            label=Config.AUTHOR_INFO_LABEL,
            command=self.show_author_info,
        )  # 添加作者信息到关于菜单
        self.about_menu.add_command(
            label=Config.INSTRUCTIONS_LABEL,
            command=self.show_instructions,
        )  # 添加使用说明到关于菜单

    def setup_ui(self):  # 设置UI方法
        # Configure grid  # 配置网格
        for i in range(9):  # 遍历9次
            self.root.rowconfigure(
                i, weight=1
            )  # 设置行的权重为1
        for i in range(3):  # 遍历3次
            self.root.columnconfigure(
                i, weight=1
            )  # 设置列的权重为1

        self.create_input_fields()  # 创建输入字段
        self.create_buttons_and_checkboxes()  # 创建按钮和复选框
        self.create_progress_label()  # 创建进度标签

        # Set initial values from config  # 从配置文件中设置初始值
        self.update_ui_from_config(
            self.logic.config
        )  # 更新UI从配置文件中的值

    def create_input_fields(self):  # 创建输入字段方法
        fields = [  # 定义字段列表
            ("要处理的列 (如A B C):", "columns"),  # 列字段
            ("要处理的行 (如1 2 3):", "rows"),  # 行字段
            ("目标字符", "target_char"),
            ("替换为:", "replace_char"),  # 替换字符字段
        ]

        for i, (label_text, config_key) in enumerate(
            fields
        ):  # 遍历字段
            label = ttk.Label(
                self.root, text=label_text
            )  # 创建标签
            label.grid(
                row=i,
                column=0,
                padx=10,
                pady=10,
                sticky="ew",
            )  # 设置标签的位置

            entry = ttk.Entry(self.root)  # 创建输入框
            entry.grid(
                row=i,
                column=1,
                columnspan=2,
                padx=10,
                pady=10,
                sticky="ew",
            )  # 设置输入框的位置
            setattr(
                self, f"{config_key}_entry", entry
            )  # 设置输入框的属性

    def create_buttons_and_checkboxes(
        self,
    ):  # 创建按钮和复选框方法
        # File selection  # 文件选择
        self.select_button = ttk.Button(
            self.root,
            text=Config.SELECT_EXCEL_BUTTON_TEXT,  # 创建文件选择按钮
            command=self.select_file,
        )  # 设置文件选择按钮的命令
        self.select_button.grid(
            row=4, column=0, padx=10, pady=10, sticky="ew"
        )  # 设置文件选择按钮的位置

        self.file_path_entry = ttk.Entry(
            self.root
        )  # 创建文件路径输入框
        self.file_path_entry.grid(
            row=4, column=1, padx=10, pady=10, sticky="ew"
        )  # 设置文件路径输入框的位置

        # Checkboxes  # 复选框
        select_folder_cb = ttk.Checkbutton(
            self.root,
            text=Config.SELECT_FOLDER_BUTTON_TEXT,  # 创建选择文件夹复选框
            variable=self.select_folder_var,  # 设置选择文件夹复选框的变量
            command=self.update_select_button_text,
        )  # 设置选择文件夹复选框的命令
        select_folder_cb.grid(
            row=4, column=2, padx=10, pady=10, sticky="w"
        )  # 设置选择文件夹复选框的位置

        overwrite_cb = ttk.Checkbutton(
            self.root,
            text=Config.OVERWRITE_CHECKBOX_TEXT,  # 创建覆盖复选框
            variable=self.overwrite_var,
        )  # 设置覆盖复选框的变量
        overwrite_cb.grid(
            row=5,
            column=0,
            columnspan=3,
            padx=10,
            pady=10,
            sticky="w",
        )  # 设置覆盖复选框的位置

        # Execute button  # 执行按钮
        self.execute_button = ttk.Button(
            self.root,
            text=Config.EXECUTE_BUTTON_TEXT,  # 创建执行按钮
            command=self.execute_cleaning,
        )  # 设置执行按钮的命令
        self.execute_button.grid(
            row=6,
            column=0,
            columnspan=3,
            padx=10,
            pady=10,
            sticky="ew",
        )  # 设置执行按钮的位置

    def create_progress_label(self):  # 创建进度标签方法
        self.progress_label = ttk.Label(
            self.root, text=Config.PROGRESS_LABEL_TEXT
        )  # 创建进度标签
        self.progress_label.grid(
            row=7,
            column=0,
            columnspan=3,
            padx=10,
            pady=10,
            sticky="ew",
        )  # 设置进度标签的位置

    def update_ui_from_config(
        self, config
    ):  # 从配置文件中更新UI方法
        """Update UI components with values from config"""  # 更新UI组件的方法
        self.columns_entry.insert(  # type: ignore
            0, config.get("columns", "")
        )  # type: ignore # 更新列输入框的值
        self.rows_entry.insert(0, config.get("rows", ""))  # type: ignore # 更新行输入框的值
        self.target_char_entry.insert(  # type: ignore
            0,
            config.get(
                "target_char", Config.DEFAULT_target_char
            ),
        )  # type: ignore # 更新处理字符输入框的值
        self.replace_char_entry.insert(  # type: ignore
            0,
            config.get(
                "replace_char", Config.DEFAULT_replace_char
            ),
        )  # type: ignore # 更新替换字符输入框的值
        self.file_path_entry.insert(  # type: ignore
            0, config.get("last_file_path", "")
        )  # type: ignore # 更新文件路径输入框的值
        self.select_folder_var.set(
            config.get("select_folder", False)
        )  # 更新选择文件夹复选框的值
        self.overwrite_var.set(
            config.get("overwrite", False)
        )  # 更新覆盖复选框的值

    def get_ui_values(self):  # 获取UI值的方法
        """Get current values from UI components"""  # 获取当前UI组件的值
        return {
            "columns": self.columns_entry.get(),  # 取列输入框的值 # type: ignore
            "rows": self.rows_entry.get(),  # 获取行输入框的值 # type: ignore
            "target_char": self.target_char_entry.get(),  # 获取处理字符输入框的值 # type: ignore
            "replace_char": self.replace_char_entry.get(),  # 获取替换字符输入框的值 # type: ignore
            "select_folder": self.select_folder_var.get(),  # 获取选择文件夹复选框的值
            "overwrite": self.overwrite_var.get(),  # 获取覆盖复选框的值
            "last_file_path": self.file_path_entry.get(),  # 获取文件路径输入框的值 # type: ignore
        }

    def update_progress(
        self, current, total, filename
    ):  # 更新进度的方法
        """Update progress label"""  # 更新进度标签
        progress_text = f"{Config.PROGRESS_LABEL_TEXT} {current}/{total} - 正在处理: {filename}"  # 进度文本
        self.progress_label.config(text=progress_text)  # type: ignore # 设置进度标签的文本
        self.root.update_idletasks()  # 更新根窗口的任务

    def update_select_button_text(
        self,
    ):  # 更新文件选择按钮的文本的方法
        """更新文件选择按钮的文本"""  # 更新文件选择按钮的文本
        if (
            self.select_folder_var.get()
        ):  # 如果选择文件夹复选框被选中
            self.select_button.config(  # type: ignore
                text=Config.SELECT_FOLDER_BUTTON_TEXT
            )  # type: ignore # 设置文件选择按钮的文本为选择文件夹按钮文本
        else:  # 如果选择文件夹复选框未被选中
            self.select_button.config(  # type: ignore
                text=Config.SELECT_EXCEL_BUTTON_TEXT
            )  # type: ignore # 设置文件选择按钮的文本为选择Excel文件按钮文本

    def select_file(self):  # 选择文件的方法
        """处理文件/文件夹选择"""  # 处理文件/文件夹选择
        if (
            self.select_folder_var.get()
        ):  # 如果选择文件夹复选框被选中
            path = filedialog.askdirectory()  # 选择文件夹
            message = (
                Config.FOLDER_SELECTED_MESSAGE
            )  # 文件夹选择消息
            title = (
                Config.SELECT_FOLDER_BUTTON_TEXT
            )  # 文件夹选择标题
        else:  # 如果选择文件夹复选框未被选中
            path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx")]
            )  # 选择Excel文件
            message = (
                Config.FILE_SELECTED_MESSAGE
            )  # Excel文件选择消息
            title = (
                Config.SELECT_EXCEL_BUTTON_TEXT
            )  # Excel文件选择标题

        if path:  # 如果路径存在
            self.file_path_entry.delete(0, tk.END)  # type: ignore # 删除文件路径输入框的值
            self.file_path_entry.insert(0, path)  # type: ignore # 插入文件路径
            messagebox.showinfo(
                title, message.format(path)
            )  # 显示文件选择消息

    def execute_cleaning(self):  # 执行处理的方法
        """Execute the cleaning process"""  # 执行处理过程
        values = self.get_ui_values()  # 获取UI值
        if not values[
            "last_file_path"
        ]:  # 如果文件路径不存在
            messagebox.showwarning(
                Config.WARNING_MESSAGE,
                Config.SELECT_FILE_WARNING,
            )  # 显示选择文件警告
            return

        if values["overwrite"]:  # 如果覆盖复选框被选中
            response = messagebox.askyesno(
                "备份提醒",
                "请确认您已经提前备份了文件。是否继续执行？",
            )  # 提醒用户已经提前备份了文件并询问是否继续
            if not response:  # 如果用户选择不继续
                return  # 停止执行

        def on_complete(
            success_count, total_count
        ):  # 处理完成的方法
            self.execute_button.config(  # type: ignore
                text=Config.EXECUTE_BUTTON_TEXT,
                state="normal",
            )  # type: ignore # 设置执行按钮的文本和状态
            if total_count > 0:  # 如果总数大于0
                if (
                    success_count == total_count
                ):  # 如果成功数等于总数
                    messagebox.showinfo(
                        "完成",
                        Config.PROCESS_COMPLETE_MESSAGE,
                    )  # 显示处理完成消息
                else:  # 如果成功数不等于总数
                    messagebox.showwarning(
                        "处理完成",
                        f"共处理 {total_count} 个文件，成功 {success_count} 个，失败 {total_count - success_count} 个",
                    )  # 显示处理完成消息

        # Disable execute button and change text while processing  # 在处理过程中禁用执行按钮并更改文本
        self.execute_button.config(  # type: ignore
            text="处理中...", state="disabled"
        )  # type: ignore # 设置执行按钮的文本和状态

        # Start processing  # 开始处理
        if not self.logic.execute_cleaning(
            values, self.update_progress, on_complete
        ):  # 如果逻辑处理未开始
            messagebox.showwarning(
                "警告", "正在处理中，请等待当前任务完成"
            )  # 显示警告
            return

    def show_project_info(self):  # 显示项目信息的方法
        messagebox.showinfo(
            Config.PROJECT_INFO_LABEL, Config.PROJECT_INFO
        )  # 显示项目信息

    def show_author_info(self):  # 显示作者信息的方法
        messagebox.showinfo(
            Config.AUTHOR_INFO_LABEL, Config.AUTHOR_INFO
        )  # 显示作者信息

    def show_instructions(self):  # 显示使用说明的方法
        messagebox.showinfo(
            Config.INSTRUCTIONS_LABEL, Config.INSTRUCTIONS
        )  # 显示使用说明

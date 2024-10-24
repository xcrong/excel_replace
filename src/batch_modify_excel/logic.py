# logic.py
import os  # 导入os模块，用于文件路径操作
import json  # 导入json模块，用于处理配置文件
import threading  # 导入threading模块，用于多线程处理
from tkinter import messagebox  # 从tkinter导入messagebox模块，用于显示消息框
from openpyxl import load_workbook  # 从openpyxl导入load_workbook函数，用于加载Excel文件
from batch_modify_excel.config import Config  # 从配置文件中导入Config类

class ExcelCleanerLogic:
    def __init__(self):
        self.config_file = Config.CONFIG_FILE  # 设置配置文件路径
        self.config = self.load_config()  # 加载配置文件
        self.processing = False  # 初始化处理标志为False
        self.current_thread = None  # 初始化当前线程为None

    def load_config(self):
        """从文件中加载配置"""
        if os.path.exists(self.config_file):  # 检查配置文件是否存在
            with open(self.config_file, 'r', encoding='utf-8') as f:  # 打开配置文件
                return json.load(f)  # 返回配置文件的内容
        return {}  # 如果配置文件不存在，返回空字典

    def save_config(self, values):
        """将配置保存到文件中"""
        try:
            self.config.update(values)  # 更新配置字典
            with open(self.config_file, 'w', encoding='utf-8') as f:  # 打开配置文件
                json.dump(self.config, f, ensure_ascii=False, indent=4)  # 将配置保存到文件中
        except Exception as e:
            print(Config.CONFIG_SAVE_ERROR.format(str(e)))  # 如果保存配置时出错，打印错误信息

    def parse_input(self, input_str, is_row=False):
        """将输入字符串解析为列表"""
        if not input_str:  # 如果输入字符串为空
            return None  # 返回None
        
        for sep in Config.SEPARATORS:  # 遍历分隔符列表
            if sep in input_str:  # 如果输入字符串中包含分隔符
                items = [item.strip() for item in input_str.split(sep)]  # 将输入字符串按分隔符分割，并去除空格
                if is_row:  # 如果是解析行号
                    try:
                        return [int(item) for item in items]  # 尝试将每个项转换为整数
                    except ValueError:
                        messagebox.showerror(Config.INPUT_ERROR_MESSAGE, Config.ROW_NUMBER_ERROR)  # 如果转换失败，显示错误信息
                        return None
                return items  # 返回解析后的列表
        return [input_str.strip()]  # 如果输入字符串中不包含分隔符，返回包含输入字符串的列表

    def clean_excel(self, file_path, columns, rows, clean_char, overwrite):
        """根据指定的参数清理Excel文件"""
        try:
            wb = load_workbook(file_path)  # 加载Excel文件
            
            for sheet in wb.worksheets:  # 遍历工作表
                if columns:  # 如果指定了列
                    for col in columns:  # 遍历列
                        for cell in sheet[col]:  # 遍历列中的每个单元格
                            if cell.value and clean_char in str(cell.value):  # 如果单元格值包含要清理的字符
                                cell.value = str(cell.value).replace(clean_char, "")  # 清理单元格值
                                 
                if rows:  # 如果指定了行
                    for row_num in rows:  # 遍历行号
                        for cell in sheet[row_num]:  # 遍历行中的每个单元格
                            if cell.value and clean_char in str(cell.value):  # 如果单元格值包含要清理的字符
                                cell.value = str(cell.value).replace(clean_char, "")  # 清理单元格值
                                 
                if not columns and not rows:  # 如果既没有指定列也没有指定行
                    for row in sheet.iter_rows():  # 遍历工作表中的每一行
                        for cell in row:  # 遍历行中的每个单元格
                            if cell.value and clean_char in str(cell.value):  # 如果单元格值包含要清理的字符
                                cell.value = str(cell.value).replace(clean_char, "")  # 清理单元格值

            save_path = file_path if overwrite else os.path.join(
                os.path.dirname(file_path),
                "new_" + os.path.basename(file_path)
            )  # 根据是否覆盖原文件决定保存路径
            
            wb.save(save_path)  # 保存Excel文件
            print(Config.FILE_SAVED_MESSAGE.format(save_path))  # 打印文件保存的消息
            return True  # 返回成功标志
        except Exception as e:
            print(f"处理文件 {file_path} 时出错: {str(e)}")  # 如果处理文件时出错，打印错误信息
            return False  # 返回失败标志

    def process_files(self, values, progress_callback, complete_callback):
        """在后台线程中处理文件"""
        file_path = values['last_file_path']  # 获取文件路径
        columns = self.parse_input(values['columns'])  # 解析列输入
        rows = self.parse_input(values['rows'], is_row=True)  # 解析行输入
        clean_char = values['clean_char']  # 获取要清理的字符
        
        try:
            if values['select_folder']:  # 如果选择了文件夹
                total_files = sum(1 for _, _, files in os.walk(file_path) 
                                for file in files if file.endswith('.xlsx'))  # 计算文件夹中所有Excel文件的数量
                processed_files = 0  # 初始化已处理文件的数量
                success_count = 0  # 初始化成功处理的文件数量
                
                for root, dirs, files in os.walk(file_path):  # 遍历文件夹
                    if not self.processing:  # 检查是否应继续处理
                        break
                        
                    for file in files:  # 遍历文件
                        if not self.processing:  # 检查是否应继续处理
                            break
                            
                        if file.endswith('.xlsx'):  # 如果文件是Excel文件
                            full_path = os.path.join(root, file)  # 获取文件的完整路径
                            if self.clean_excel(full_path, columns, rows, 
                                           clean_char, values['overwrite']):  # 调用clean_excel方法处理文件
                                success_count += 1  # 如果处理成功，增加成功处理的文件数量
                            processed_files += 1  # 增加已处理文件的数量
                            progress_callback(processed_files, total_files, file)  # 调用进度回调函数
                complete_callback(success_count, processed_files)  # 调用完成回调函数
            else:  # 如果没有选择文件夹
                filename = os.path.basename(file_path)  # 获取文件名
                success = self.clean_excel(file_path, columns, rows, 
                                      clean_char, values['overwrite'])  # 调用clean_excel方法处理文件
                progress_callback(1, 1, filename)  # 调用进度回调函数
                complete_callback(1 if success else 0, 1)  # 调用完成回调函数
                
        except Exception as e:
            print(f"处理文件时出错: {str(e)}")  # 如果处理文件时出错，打印错误信息
            complete_callback(0, 1)  # 调用完成回调函数，表示处理失败
        finally:
            self.processing = False  # 设置处理标志为False

    def execute_cleaning(self, values, progress_callback, complete_callback):
        """开始清理过程在一个单独的线程中"""
        if self.processing:  # 检查是否已经在处理
            return False  # 如果已经在处理，返回False
            
        self.processing = True  # 设置处理标志为True
        self.save_config(values)  # 保存配置
        
        # 在新线程中开始处理
        self.current_thread = threading.Thread(
            target=self.process_files,
            args=(values, progress_callback, complete_callback)
        )
        self.current_thread.daemon = True  # 设置线程为守护线程，确保主程序退出时线程也会退出
        self.current_thread.start()  # 启动线程
        return True  # 返回成功标志

    def stop_processing(self):
        """停止当前的处理过程"""
        self.processing = False  # 设置处理标志为False
        if self.current_thread and self.current_thread.is_alive():  # 检查当前线程是否仍然活着
            self.current_thread.join(timeout=1.0)  # 等待线程完成或超时

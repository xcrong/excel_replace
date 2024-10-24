# main.py
import ttkbootstrap as ttk  # 导入ttkbootstrap库，用于创建图形用户界面
from batch_modify_excel.ui import ExcelCleanerUI  # 从ui模块中导入ExcelCleanerUI类，用于创建用户界面
from batch_modify_excel.logic import ExcelCleanerLogic  # 从logic模块中导入ExcelCleanerLogic类，用于处理Excel文件的清理逻辑

def main():
    root = ttk.Window(themename="litera")  # 创建一个ttkbootstrap窗口，主题名为"litera"
    root.title("Excel 批量清理工具")  # 设置窗口标题为"Excel 批量清理工具"
    
    # 设置窗口初始大小为1200x600，并且不允许用户调整窗口大小
    root.geometry("1200x600")  
    root.resizable(False, False)
    
    logic = ExcelCleanerLogic()  # 创建一个ExcelCleanerLogic实例，用于处理Excel文件的清理逻辑
    app = ExcelCleanerUI(root, logic)  # 创建一个ExcelCleanerUI实例，用于创建用户界面
    
    # 定义一个函数，用于处理窗口关闭事件
    def on_closing():
        logic.save_config(app.get_ui_values())  # 在窗口关闭前，保存用户界面中的值到配置文件中
        root.destroy()  # 销毁窗口
    
    # 将on_closing函数绑定到窗口关闭事件上
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()  # 启动事件循环，开始处理用户交互

if __name__ == "__main__":
    main()  # 如果脚本直接运行（不是被导入），则调用main函数

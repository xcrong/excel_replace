import ttkbootstrap as ttk
from batch_modify_excel.ui import ExcelCleanerUI
from batch_modify_excel.logic import ExcelReplicerLogic
import random

light_themes = [
    "cosmo",
    "flatly",
    "litera",
    "minty",
    "lumen",
    "sandstone",
    "yeti",
    "pulse",
    "united",
    "morph",
    "journal",
    "simplex",
    "cerculean",
]


def main():
    root = ttk.Window(themename=random.choice(light_themes))
    root.title("Excel 批量处理工具")
    root.geometry("1200x600")
    root.resizable(False, False)

    logic = ExcelReplicerLogic()
    app = ExcelCleanerUI(root, logic)

    def on_closing():
        logic.save_config(app.get_ui_values())
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()

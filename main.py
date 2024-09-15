import tkinter as tk
import gui
from database import create_table

def main():
    # 創建窗口
    window = tk.Tk()
    window.geometry("1270x740")
    window.title("五星表單工具")
    # 創建資料庫
    create_table()
    # GUI介面
    app = gui.Application(window)

    # 執行窗口
    window.mainloop()

if __name__ == "__main__":
    main()
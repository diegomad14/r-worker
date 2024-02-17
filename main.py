import tkinter as tk
from app_logic import AppLogic

if __name__ == "__main__":
    root = tk.Tk()
    app = AppLogic(root)

    root.mainloop()
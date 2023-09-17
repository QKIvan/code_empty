from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from PIL import Image, ImageTk
import ctypes
import openpyxl

# Get screen resolution
user32 = ctypes.windll.user32
screen_width = user32.GetSystemMetrics(0)
screen_height = user32.GetSystemMetrics(1)

# Intial name_list , checked_name_list
name_list = ['小明', '小红', '小张', '小李']
checked_name_list = []

# Intial Tkinter
root = Tk()
root.title("Achecklist")
root.geometry("{}x{}".format(screen_width, screen_height))

# Define handle_check function
def handle_check(name):
    if name in checked_name_list:
        checked_name_list.remove(name)
        # Update icon when remove name
        var_dict[name].set(False)
    else:
        checked_name_list.append(name)
        # Update icon when add name
        var_dict[name].set(True)

# Load icon
checked_image = Image.open("checked.png")
checked_image = checked_image.resize((30, 30))
checked_icon = ImageTk.PhotoImage(checked_image)
var_dict = {} 

# Create Label , CheckBox
for i, name in enumerate(name_list):
    var_dict[name] = BooleanVar()
    cb = Checkbutton(root, text=name, variable=var_dict[name], command=lambda name=name: handle_check(name))
    cb.grid(row=i, column=0)

    # Set default mode
    var_dict[name].set(0)

# Create confirmation button
btn = Button(root, text="确认")
btn.grid(row=len(name_list), column=0)
# Print selected names
def print_checked_names():
    print("共勾选了 {} 个名字：".format(len(checked_name_list)))
    for name in checked_name_list:
        print(name)
    export_to_excel()

btn.config(command=print_checked_names)

# Export the names to Excel
def export_to_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Selected Names"
    for i, name in enumerate(checked_name_list):
        ws.cell(row=i+1, column=1, value=name)
    wb.save("selected_names.xlsx")

# Menu bar
menu_bar = Menu(root)
file_menu = Menu(menu_bar, tearoff=0)
help_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="文件", menu=file_menu)
resolutions = ["800x600", "1024x768", "1280x720", "1366x768", "1920x1080", "2560x1440", "3840x2160"]
selected_resolution = StringVar()
selected_resolution.set("{}x{}".format(screen_width, screen_height))
resolution_menu = Menu(menu_bar, tearoff=0)
for res in resolutions:
    resolution_menu.add_command(label=res, command=lambda res=res: root.geometry(res))
menu_bar.add_cascade(label="分辨率", menu=resolution_menu)
help_menu.add_command(label="关于", command=lambda: messagebox.showinfo("关于", "Github: https://github.com/QKIvan/Achecklist\nContributors: QKIvan、Multicolo"))
menu_bar.add_cascade(label="帮助", menu=help_menu)

root.config(menu=menu_bar)

root.mainloop()

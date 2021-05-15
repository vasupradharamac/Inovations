from tkinter import *
from tkinter import filedialog
from tkinter import font
import ctypes
from tkinter import colorchooser
import os, sys
import win32print
import win32api
import tkinter.messagebox

root = Tk()
root.title('Code-Docx')
root.iconbitmap(r"code_docs.ico")
root.geometry("1200x660")
ctypes.windll.shcore.SetProcessDpiAwareness(1)

# random string
myappid = 'Vasu"s Codex'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

# set variable for open filename
global open_status_name
open_status_name = False

global selected
selected = False


# Exiting editor
def exit_editor(event=None):
    if tkinter.messagebox.askokcancel("Exit", "Are you sure you want to Quit?"):
        root.destroy()


# New File function
def new_file(e):
    my_text.delete("1.0", END)
    # updating statusbar
    root.title("New File - Textpad")
    status_bar.config(text="New File        ")
    global open_status_name
    open_status_name = False


# Open File function
def open_file(e):
    # delete the previously existing texts
    my_text.delete("1.0", END)

    # Grab FileName
    text_file = filedialog.askopenfilename(initialdir="C:/gui/", title="Open File", filetypes=(
        ("Text Files", "*.txt"), ("HTML Files", "*.html"), ("Python Files", "*.py"), ("All Files", "*.*")))
    # checking if the file exists
    if text_file:
        # making file name global
        global open_status_name
        open_status_name = text_file

    name = text_file
    status_bar.config(text=f'{name}        ')
    name = name.replace("C:/gui/", "")
    root.title(f'{name} - Textpad')

    # Opening the file
    text_file = open(text_file, 'r')
    stuff = text_file.read()

    # Adding file to textbox
    my_text.insert(END, stuff)
    # closing the opened file
    text_file.close()


# Save As function
def save_as_file():
    text_file = filedialog.asksaveasfilename(defaultextension='.*', initialdir="C:/gui/", title="Save File", filetypes=(
        ("Text Files", "*.txt"), ("HTML Files", "*.html"), ("Python Files", "*.py"), ("All Files", "*.*")))
    if text_file:
        name = text_file
        status_bar.config(text=f'Saved: {name}        ')
        name = name.replace("C:/gui/", "")
        root.title(f'{name} - Textpad')

        # save the file
        text_file = open(text_file, 'w')
        text_file.write(my_text.get(1.0, END))
        # close the file
        text_file.close()


def save_file(e):
    global open_status_name
    if open_status_name:
        # save the file
        text_file = open(open_status_name, 'w')
        text_file.write(my_text.get(1.0, END))
        text_file.close()

        # status update or popup code

        status_bar.config(text=f'Saved: {open_status_name}')
    else:
        save_as_file()


def cut_text(e):
    global selected

    if e:
        selected = root.clipboard_get()
    else:
        if my_text.selection_get():
            selected = my_text.selection_get()
            # deleting the selected text
            my_text.delete("sel.first", "sel.last")

            # delete the already existing or previous text on clipboard and append it with the selected text.
            root.clipboard_clear()
            root.clipboard_append(selected)


def copy_text(e):
    global selected

    if e:
        selected = root.clipboard_get()

    if my_text.selection_get():
        # grabbing selected text from text box
        selected = my_text.selection_get()

        # delete the already existing or previous text on clipboard and append it with the selected text.
        root.clipboard_clear()
        root.clipboard_append(selected)


def paste_text(e):
    global selected

    if e:
        selected = root.clipboard_get()
    else:
        if selected:
            position = my_text.index(INSERT)
            my_text.insert(position, selected)


def bold_text(e):
    # Creating font
    bold_font = font.Font(my_text, my_text.cget("font"))
    bold_font.configure(weight="bold")

    # Configuring tag
    my_text.tag_configure("bold", font=bold_font)

    # Defining current tags
    current_tags = my_text.tag_names("sel.first")

    # Check if tag has been set
    if "bold" in current_tags:
        my_text.tag_remove("bold", "sel.first", "sel.last")
    else:
        my_text.tag_add("bold", "sel.first", "sel.last")


def italics_text(e):
    # Creating font
    italics_font = font.Font(my_text, my_text.cget("font"))
    italics_font.configure(slant="italic")

    # Configuring tag
    my_text.tag_configure("italic", font=italics_font)

    # Defining current tags
    current_tags = my_text.tag_names("sel.first")

    # Check if tag has been set
    if "italic" in current_tags:
        my_text.tag_remove("italic", "sel.first", "sel.last")
    else:
        my_text.tag_add("italic", "sel.first", "sel.last")


# Change selected text color
def text_color():
    # Pick a color
    my_color = colorchooser.askcolor()[1]
    if my_color:
        # Creating font
        color_font = font.Font(my_text, my_text.cget("font"))

        # Configuring tag
        my_text.tag_configure("colored", font=color_font, foreground=my_color)

        # Defining current tags
        current_tags = my_text.tag_names("sel.first")

        # Check if tag has been set
        if "colored" in current_tags:
            my_text.tag_remove("colored", "sel.first", "sel.last")
        else:
            my_text.tag_add("colored", "sel.first", "sel.last")


# Changing background color
def bg_color():
    my_color = colorchooser.askcolor()[1]
    if my_color:
        my_text.config(bg=my_color)


# Changing the entire text color
def all_text_color():
    my_color = colorchooser.askcolor()[1]
    if my_color:
        my_text.config(fg=my_color)


def print_file(e):
    # printer_name = win32print.GetDefaultPrinter()
    # status_bar.config(text=printer_name)

    # Grab FileName
    print_to_file = filedialog.askopenfilename(initialdir="C:/gui/", title="Open File", filetypes=(
        ("Text Files", "*.txt"), ("HTML Files", "*.html"), ("Python Files", "*.py"), ("All Files", "*.*")))
    if print_to_file:
        win32api.ShellExecute(0, "print", print_to_file, None, ".", 0)


# Select all text
def select_all(e):
    # add "sel" tag for selecting all text
    my_text.tag_add('sel', '1.0', 'end')


# clear all text
def clear_all():
    my_text.delete(1.0, END)


# Finding and replacing texts
def find_text(event=None):
    search_toplevel = Toplevel(root)
    search_toplevel.title('Find Text')
    search_toplevel.transient(root)
    search_toplevel.resizable(False, False)
    Label(search_toplevel, text="Find All:").grid(row=0, column=0, sticky='e')
    search_entry_widget = Entry(search_toplevel, width=25)
    search_entry_widget.grid(row=0, column=1, padx=2, pady=2, sticky='we')
    search_entry_widget.focus_set()
    ignore_case_value = IntVar()
    Checkbutton(search_toplevel, text='Ignore Case', variable=ignore_case_value).grid(row=1, column=1, sticky='e',
                                                                                      padx=2, pady=2)
    Button(search_toplevel, text="Find All", underline=0,
           command=lambda: search_output(
               search_entry_widget.get(), ignore_case_value.get(),
               my_text, search_toplevel, search_entry_widget)
           ).grid(row=0, column=2, sticky='e' + 'w', padx=2, pady=2)

    def close_search_window():
        my_text.tag_remove('match', '1.0', END)
        search_toplevel.destroy()

    search_toplevel.protocol('WM_DELETE_WINDOW', close_search_window)
    return "break"


def search_output(needle, if_ignore_case, my_text, search_toplevel, search_box):
    my_text.tag_remove('match', '1.0', END)
    matches_found = 0
    if needle:
        start_pos = '1.0'
        while True:
            start_pos = my_text.search(needle, start_pos, nocase=if_ignore_case, stopindex=END)
            if not start_pos:
                break
            end_pos = '{} + {}c'.format(start_pos, len(needle))
            my_text.tag_add('match', start_pos, end_pos)
            matches_found += 1
            start_pos = end_pos
        my_text.tag_config('match', background="#ADD8E6", foreground='blue')
    search_box.focus_set()
    search_toplevel.title('{} matches found'.format(matches_found))


# Turn night mode on
def night_on():
    main_color = "#000000"
    second_color = "#373737"
    text_color = "green"
    text_color2 = "#ffffff"

    root.config(bg=main_color)
    status_bar.config(bg=second_color, fg="red")
    my_text.config(bg=main_color, fg=text_color2)

    # Changing menu colors
    file_menu.config(bg=second_color, fg=text_color2)
    edit_menu.config(bg=second_color, fg=text_color2)
    style_menu.config(bg=second_color, fg=text_color2)
    options_menu.config(bg=second_color, fg=text_color2)


# Turn night mode off
def night_off():
    main_color = "SystemButtonFace"
    second_color = "SystemButtonFace"
    text_color = "black"
    text_color2 = "black"

    root.config(bg=main_color)
    status_bar.config(bg=second_color, fg="black")
    my_text.config(bg="white", fg=text_color2)

    # Changing menu colors
    file_menu.config(bg=second_color, fg=text_color2)
    edit_menu.config(bg=second_color, fg=text_color2)
    style_menu.config(bg=second_color, fg=text_color2)
    options_menu.config(bg=second_color, fg=text_color2)


def ash_theme():
    root.config(bg="#D1D4D1")
    status_bar.config(bg="#D1D4D1", fg="black")
    my_text.config(bg="#696969", fg="#ffffff")

    # Changing menu colors
    file_menu.config(bg="#808080", fg="#ffffff")
    edit_menu.config(bg="#808080", fg="#ffffff")
    style_menu.config(bg="#808080", fg="#ffffff")
    options_menu.config(bg="#808080", fg="#ffffff")


def aquaminerale_theme():
    root.config(bg="#ffffff")
    status_bar.config(bg="#ffffff", fg="black")
    my_text.config(bg="#87CEEB", fg="#000000")

    # Changing menu colors
    file_menu.config(bg="#F0F8FF", fg="#000000")
    edit_menu.config(bg="#F0F8FF", fg="#000000")
    style_menu.config(bg="#F0F8FF", fg="#000000")
    options_menu.config(bg="#F0F8FF", fg="#000000")


def caramel_theme():
    root.config(bg="#ffffff")
    status_bar.config(bg="#ffffff", fg="black")
    my_text.config(bg="#C68E17", fg="#000000")

    # Changing menu colors
    file_menu.config(bg="#F0E68C", fg="#000000")
    edit_menu.config(bg="#F0E68C", fg="#000000")
    style_menu.config(bg="#F0E68C", fg="#000000")
    options_menu.config(bg="#F0E68C", fg="#000000")


def olive_theme():
    root.config(bg="#D1D4D1")
    status_bar.config(bg="#D1D4D1", fg="black")
    my_text.config(bg="#808000", fg="#ffffff")

    # Changing menu colors
    file_menu.config(bg="#BDB76B", fg="#000000")
    edit_menu.config(bg="#BDB76B", fg="#000000")
    style_menu.config(bg="#BDB76B", fg="#000000")
    options_menu.config(bg="#BDB76B", fg="#000000")

def dynamic_resize(widget):
    # getting current window height and width
    curh,curw=root.winfo_height(),root.winfo_width()
    # using them to update text_box
    my_text.configure(height=curh,width=curw)



# Main Frame
my_frame = Frame(root)
my_frame.pack(pady=5)

# Vertical Scrollbar for text box
text_scroll = Scrollbar(my_frame)
text_scroll.pack(side=RIGHT, fill=Y)

# Horizontal scrollbar
hor_scroll = Scrollbar(my_frame, orient="horizontal")
hor_scroll.pack(side=BOTTOM, fill=X)

# text box
my_text = Text(my_frame, width=97, height=25, font=("Lato", 16), selectbackground="#ADD8E6", selectforeground="black",
               undo=True, yscrollcommand=text_scroll.set, wrap="none", xscrollcommand=hor_scroll.set)
my_text.pack()

# configuring scrollbar
text_scroll.config(command=my_text.yview)
hor_scroll.config(command=my_text.xview)

# creating_Menu
my_menu = Menu(root)
root.config(menu=my_menu)

# File_Menu
file_menu = Menu(my_menu, tearoff=False)
my_menu.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="New", command=lambda: new_file(False), accelerator="(Ctrl+N)")
file_menu.add_command(label="Open", command=lambda: open_file(False), accelerator="(Ctrl+O)")
file_menu.add_command(label="Save", command=lambda: save_file(False), accelerator="(Ctrl+S)")
file_menu.add_command(label="Save As", command=save_as_file)
file_menu.add_separator()
file_menu.add_command(label="Print File", command=lambda: print_file(False), accelerator="(Ctrl+P)")
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.quit)

# Edit_Menu
edit_menu = Menu(my_menu, tearoff=False)
my_menu.add_cascade(label="Edit", menu=edit_menu)
edit_menu.add_command(label="Cut", command=lambda: cut_text(False), accelerator="(Ctrl+X)")
edit_menu.add_command(label="Copy", command=lambda: copy_text(False), accelerator="(Ctrl+C)")
edit_menu.add_command(label="Paste", command=lambda: paste_text(False), accelerator="(Ctrl+V)")
edit_menu.add_separator()
edit_menu.add_command(label="Undo", command=my_text.edit_undo, accelerator="(Ctrl+Z)")
edit_menu.add_command(label="Redo", command=my_text.edit_redo, accelerator="(Ctrl+Y)")
edit_menu.add_command(label="Find Text", accelerator="(Ctrl+F)", command=lambda: find_text(False))
edit_menu.add_separator()
edit_menu.add_command(label="Select All", command=lambda: select_all(False), accelerator="(Ctrl+A)")
edit_menu.add_command(label="Clear", command=clear_all)

# style_Menu
style_menu = Menu(my_menu, tearoff=False)
my_menu.add_cascade(label="Style", menu=style_menu)
style_menu.add_command(label="Bold", command=lambda: bold_text(False), accelerator="(Ctrl+B)")
style_menu.add_command(label="Italics", command=lambda: italics_text(False), accelerator="(Ctrl+I)")
style_menu.add_separator()
style_menu.add_command(label="Foreground Color", command=all_text_color)
style_menu.add_separator()
style_menu.add_command(label="Selected Text Color", command=text_color)
style_menu.add_command(label="Background Color", command=bg_color)

# Options_Menu
options_menu = Menu(my_menu, tearoff=False)
my_menu.add_cascade(label="Themes", menu=options_menu)
options_menu.add_command(label="Default", command=night_off)
options_menu.add_separator()
options_menu.add_command(label="Dark Mode", command=night_on)
options_menu.add_separator()
options_menu.add_command(label="Ash", command=ash_theme)
options_menu.add_command(label="Aqua Minerale", command=aquaminerale_theme)
options_menu.add_command(label="Caramel", command=caramel_theme)
options_menu.add_command(label="Olive", command=olive_theme)

# Adding status bar to the bottom
status_bar = Label(root, text='Ready        ')
status_bar.pack(fill=X,  ipady=15)

# Binding edit keys with keyboard shortcut
root.bind('<Control-Key-x>', cut_text)
root.bind('<Control-Key-c>', copy_text)
root.bind('<Control-Key-v>', paste_text)
root.bind('<Control-Key-b>', bold_text)
root.bind('<Control-Key-i>', italics_text)
root.bind('<Control-Key-a>', select_all)
root.bind('<Control-Key-p>', print_file)
root.bind('<Control-Key-f>', find_text)
# root.bind('<Control-key-s>', save_file)
root.bind('<Control-Key-o>', open_file)
# root.bind('<Control-Key-shift_R-s>', save_as_file)
root.bind('<Control-Key-n>', new_file)

root.protocol('WM_DELETE_WINDOW', exit_editor)
# root.resizable(height = None, width = None)

# binds resize call back to dynamic resize method
root.bind("<Configure>",dynamic_resize)
root.resizable(True,True)
root.mainloop()

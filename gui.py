"""
    CopyRight Dr. Ahmad Hamdi Emara 2020
    Adam Medical Company
"""
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# IMPORTS REGION.
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
from dataclasses import dataclass
import datetime

from eval import *
from contacts import *

from tkinter import Tk, filedialog, messagebox, Label, Button, Canvas

# =-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# GLOBAL VARIABLES REGION.
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
now = datetime.datetime.now()
cpy = f"© Dr. Ahmad Hamdi Emara {now.year}"

# =-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# THEME REGION.
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
@dataclass
class Color:
    component_bg: str = '#F7DC6F'
    fg: str = '#B7950B'
    windows_bg: str = '#FCF3CF'

colors = Color()

# =-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# START OF UI SCRIPT.
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
root = Tk()
class Tools():
    def __init__(self):
        root.resizable(False, False)

        main_canvas =Canvas(root, width = 300, height = 300, bg = colors.windows_bg, relief = 'raised')
        main_canvas.pack()
        
        root.title('Adam Helper Tools')
        self.center(root)

        lbl = Label(root, text='Choose Your Tool', bg = colors.windows_bg, fg = colors.fg)
        lbl.config(font=('helvetica', 20))
        main_canvas.create_window(150, 60, window=lbl)

        browseButton_ContactsExcel = Button(root, text="      Import Contacts File     ", command=self.getContactsExcelFile, bg=colors.component_bg, fg=colors.fg, font=('helvetica', 14, 'bold'))
        main_canvas.create_window(150, 150, window=browseButton_ContactsExcel)

        browseButton_OrderExcel = Button(root, text="         Import Order File        ", command=self.getOrderExcelFile, bg=colors.component_bg, fg=colors.fg, font=('helvetica', 14, 'bold'))
        main_canvas.create_window(150, 220, window=browseButton_OrderExcel)

        
        lbl = Label(root, text=cpy, bg = colors.windows_bg, fg=colors.fg)
        lbl.config(font=('helvetica', 12))
        main_canvas.create_window(150, 280, window=lbl)

        root.mainloop()

    def getContactsExcelFile(self):
        global read_file
    
        import_file_path = filedialog.askopenfilename()
        read_file = pd.read_excel(import_file_path, dtype={'MOBILE NO.':str})

        if convertToCSV(read_file, import_file_path):
            self.done("Contacts converted successfully !!")
        else:
            self.error("")

    def getOrderExcelFile(self):
            global read_file
            
            import_file_paths = filedialog.askopenfilenames(parent=root, title='Choose the order excel file')


            for import_file_path in import_file_paths:
                read_file = pd.read_excel(import_file_path, header=17, skipfooter=1)
                
                temp = pd.read_excel(import_file_path)

                pharmacy_name = ''
                try:
                    pharmacy_name = str(temp.iat[8, 1]).split('FROM : ')[1]
                except Exception as e:
                    self.error(e)
                    print(e)

                if evaluateOrder(import_file_path, read_file, pharmacy_name):
                    self.done("Order Evaluation Done !!")
                else:
                    err = f'"The excel file for {pharmacy_name} is corrupt, please contact ABC System Support!"'
                    self.error(err)

    def done(self, message):
        messagebox.showinfo("ADAM CO.", message)

    def error(self, err):
        messagebox.showerror("ADAM CO.", err)

    def info(self, inf):
        messagebox.showinfo('ADAM CO.', inf)

    def center(self, win):
        """
        centers a tkinter window
        :param win: the root or Toplevel window to center
        """
        win.update_idletasks()
        width = win.winfo_width()
        frm_width = win.winfo_rootx() - win.winfo_x()
        win_width = width + 2 * frm_width
        height = win.winfo_height()
        titlebar_height = win.winfo_rooty() - win.winfo_y()
        win_height = height + titlebar_height + frm_width
        x = win.winfo_screenwidth() // 2 - win_width // 2
        y = win.winfo_screenheight() // 2 - win_height // 2
        win.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        win.deiconify()

def main():
    Tools()

if __name__ == "__main__":
    main()
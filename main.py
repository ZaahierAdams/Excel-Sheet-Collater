'''

Developed by: Zaahier Adams
https://github.com/ZaahierAdams

'''

from tkinter import *
from tkinter import messagebox
from PIL import ImageTk, Image
from Collater import Excel_Collater
from os import path
import Texts


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = path.abspath(".")
    return path.join(base_path, relative_path)

def Query(): 
    index_name = entry_1.get()
    Excel_Collater(index_name)
    if Excel_Collater.success is True:
        messagebox.showinfo('Success','Excel files Successfully Collated\n Collated file saved in directory')
    else:
        messagebox.showinfo('ERROR','An error occured')
    
def About():
    messagebox.showinfo('About',Texts.Description())
def Version():
    messagebox.showinfo('About',Texts.Version())

#def main():
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~      GUI     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Master Window ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   
root = Tk()
root.title('Collater')
root.iconbitmap(resource_path("icon.ico"))
root.configure(background="white")
root.resizable(False, False)

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Dropdown Menu ~~~~~~~~~~~~~~~~~~~~~~~~~~~~

menu = Menu(root)
root.config(menu=menu)
dd_Help = Menu(menu)
menu.add_cascade(label="Help",menu=dd_Help)
dd_Help.add_command(label="About",command=About)
dd_Help.add_command(label="Version",command=Version)


# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~  Frame() ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
           
Frame_1 = Frame(root, background='white' )
Frame_3 = Frame(root, background='#9CD7D5' )
Frame_4 = Frame(root, background='#9CD7D5' )


# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~  Label() ~~~~~~~~~~~~~~~~~~~~~~~~~~~~

image_location1 = resource_path("Cover.png")
photo = ImageTk.PhotoImage(Image.open(image_location1))
Label_1 = Label(Frame_1, width=333, height=333, bg='white', image=photo)
Label_1.pack()


Label_3 = Label(Frame_3, width=50, height=1, bg='#9CD7D5', fg ='black', text = 'Index name:', font=("Helvetica", 12), anchor=N)
Label_3.grid(row=0, column =0,  sticky=E)



# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~  Button() ~~~~~~~~~~~~~~~~~~~~~~~~~~~~

button_2 = Button(Frame_4, text = "Collate Excel Sheets", bg ='#9CD7D5', fg='Black',command = Query)
button_2.grid(row=2, column =0, sticky=W, padx =170, pady=0)

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~  Text Box ~~~~~~~~~~~~~~~~~~~~~~~~~~~~

entry_1 = Entry(Frame_4)
entry_1.grid(row=0, column =0,padx =50)


# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~  pack() ~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Frame_1.pack(fill=X)   
Frame_3.pack(fill=X)
Frame_4.pack(fill=X)    


                 
root.mainloop()
    
#if __name__ == '__main__':
#    main()
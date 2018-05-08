import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
import sys
import os
import comtypes.client
from ctypes import windll
windll.shcore.SetProcessDpiAwareness(1)

wdFormatPDF = 17

root =tk.Tk()
root.title("Word To Pdf")

my_label=''
input_file_name=''
def response(): 
    global my_label,input_file_name
    input_file_name = tkinter.filedialog.askopenfilename(defaultextension=".txt", filetypes=[("All Files", "*.*"), ("Text Documents", "*.txt")])
    print(input_file_name)
    try:
        my_label.pack_forget()
    except:pass
    my_label = tk.Label(root, text=input_file_name)
    my_label.pack()
    
def Convert(event=None):
    out_file=os.path.abspath('/'.join([x for x in input_file_name.split('/')][:-1])+'\\'+ my_text.get("1.0","end-1c"))
    in_file=os.path.abspath(input_file_name)
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
    label2 = tk.Label(root, text='\npdf created!')
    label2.pack()
    
my_button = tk.Button(root, text="Choose File..",command=response) 
l1 = tk.Label(root, text='')
l2=tk.Label(root, text='') 
l3=tk.Label(root, text='') 
my_text = tk.Text(root, height=1, width=40)

convert_button=tk.Button(root, text="convert",command=Convert,height=1,width=7)

my_button.pack() 
l1.pack()
my_text.pack()
l2.pack()
convert_button.pack()
l3.pack()
root.mainloop()
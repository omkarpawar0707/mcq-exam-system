import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import openpyxl
from openpyxl import load_workbook

win = tk.Tk()
win.title("Student MCQ Examination Demo")
win.geometry("400x800")

path = r"C:\Users\hp\OneDrive\Documents\Python Mcq demo.xlsx"
A = openpyxl.load_workbook(path)
B = A["Sheet1"]

def btn_show():
    name = stnlt.get()
    roll_no = rnlt.get()
    
    que1 = value_rb1.get()
    if que1 == 1:
        question_1 = "show"
    elif que1 == 2:
        question_1 = "print"
    elif que1 == 3:
        question_1 = "display"
    elif que1 == 4:
        question_1 = "output"
    else:
        ""
        
    que2 = value_rb2.get()
    if que2 == 1:
        question_2 = "//"
    elif que2 == 2:
        question_2 = "<!-- -->"
    elif que2 == 3:
        question_2 = "#"
    elif que2 == 4:
        question_2 = "/* */"
    else:
        ""
        
    que3 = value_rb3.get()
    if que3 == 1:
        question_3 = ".pt"
    elif que3 == 2:
        question_3 = ".pyt"
    elif que3 == 3:
        question_3 = ".py"
    elif que3 == 4:
        question_3 = ".python"
    else:
        ""
        
    que4 = value_rb4.get()
    if que4 == 1:
        question_4 = "input()"
    elif que4 == 2:
        question_4 = "scan()"
    elif que4 == 3:
        question_4 = "read()"
    elif que4 == 4:
        question_4 = "get()"
    else:	
        ""
    
    que5 = value_rb5.get()
    if que5 == 1:
        question_5 = "float"
    elif que5 == 2:
        question_5 = "string"
    elif que5 == 3:
        question_5 = "int"
    elif que5 == 4:
        question_5 = "bool"
    else:
        ""
    B.append([name,roll_no,que1,que2,que3,que4,que5])
    A.save(path)
    
    if name and roll_no and que1 and que2 and que3 and que4 and que5:
        messagebox.showinfo("Status","You Data is Successfully Captuchered...!")
    else:
        messagebox.showwarning("Try Again","Please do all the MCQs")        
        

h1 = tk.Label(win,text=("------- PYTHON MCQ EXAMINATION -------"),font=("Times New Roman",12),fg="red")
h1.pack()

stnl = tk.Label(win,text=("Enter Your Full Name:"),font=("Calibri",12,"underline","bold"))
stnl.pack()
stnlt = tk.Entry(win,width="40",justify="center",bg="#c8ccc8")
stnlt.pack()

rnl = tk.Label(win,text=("Enter Your Roll Number:"),font=("Calibri",12,"underline","bold"))
rnl.pack()
rnlt = tk.Entry(win,width="20",justify="center",bg="#c8ccc8")
rnlt.pack()

#Question 1 from here
q1 = tk.Label(win,text="1. Which keyword is used to display output in Python?",font=("Arial",8,"bold"))
q1.pack(anchor=tk.W,padx=10)

value_rb1 = tk.IntVar()
ans1 = ttk.Radiobutton(win, text="a) show", variable=value_rb1, value=1)
ans1.pack(anchor=tk.W, padx=20)

ans2 = ttk.Radiobutton(win, text="b) print", variable=value_rb1, value=2)
ans2.pack(anchor=tk.W, padx=20)

ans3 = ttk.Radiobutton(win, text="c) display", variable=value_rb1, value=3)
ans3.pack(anchor=tk.W, padx=20)

ans4 = ttk.Radiobutton(win, text="d) output", variable=value_rb1, value=4)
ans4.pack(anchor=tk.W, padx=20)

#Question 2 from here
q2 = tk.Label(win,text="2. Which symbol is used for comments in Python?",font=("Arial",8,"bold"))
q2.pack(anchor=tk.W,padx=10)

value_rb2 = tk.IntVar()
ans1 = ttk.Radiobutton(win, text="a) //", variable=value_rb2, value=1)
ans1.pack(anchor=tk.W, padx=20)

ans2 = ttk.Radiobutton(win, text="b) <!-- -->", variable=value_rb2, value=2)
ans2.pack(anchor=tk.W, padx=20)

ans3 = ttk.Radiobutton(win, text="c) #", variable=value_rb2, value=3)
ans3.pack(anchor=tk.W, padx=20)

ans4 = ttk.Radiobutton(win, text="d) /* */", variable=value_rb2, value=4)
ans4.pack(anchor=tk.W, padx=20)

#Question 3 from here
q3 = tk.Label(win,text="3. What is the correct file extension for Python files?",font=("Arial",8,"bold"))
q3.pack(anchor=tk.W,padx=10)

value_rb3 = tk.IntVar()
ans1 = ttk.Radiobutton(win, text="a) .pt", variable=value_rb3, value=1)
ans1.pack(anchor=tk.W, padx=20)

ans2 = ttk.Radiobutton(win, text="b) .pyt", variable=value_rb3, value=2)
ans2.pack(anchor=tk.W, padx=20)

ans3 = ttk.Radiobutton(win, text="c) .py", variable=value_rb3, value=3)
ans3.pack(anchor=tk.W, padx=20)

ans4 = ttk.Radiobutton(win, text="d) .python", variable=value_rb3, value=4)
ans4.pack(anchor=tk.W, padx=20)

q4 = tk.Label(win,text="4. Which of the following is used to take input from the user?",font=("Arial",8,"bold"))
q4.pack(anchor=tk.W,padx=10)

value_rb4 = tk.IntVar()
ans1 = ttk.Radiobutton(win, text="a) input()", variable=value_rb4, value=1)
ans1.pack(anchor=tk.W, padx=20)

ans2 = ttk.Radiobutton(win, text="b) scan()", variable=value_rb4, value=2)
ans2.pack(anchor=tk.W, padx=20)

ans3 = ttk.Radiobutton(win, text="c) int", variable=value_rb4, value=3)
ans3.pack(anchor=tk.W, padx=20)

ans4 = ttk.Radiobutton(win, text="d) get()", variable=value_rb4, value=4)
ans4.pack(anchor=tk.W, padx=20)


q5 = tk.Label(win,text="5. Which data type is used to store whole numbers?",font=("Arial",8,"bold"))
q5.pack(anchor=tk.W,padx=10)

value_rb5 = tk.IntVar()
ans1 = ttk.Radiobutton(win, text="a) float", variable=value_rb5, value=1)
ans1.pack(anchor=tk.W, padx=20)

ans2 = ttk.Radiobutton(win, text="b) string", variable=value_rb5, value=2)
ans2.pack(anchor=tk.W, padx=20)

ans3 = ttk.Radiobutton(win, text="c) int", variable=value_rb5, value=3)
ans3.pack(anchor=tk.W, padx=20)

ans4 = ttk.Radiobutton(win, text="d) bool", variable=value_rb5, value=4)
ans4.pack(anchor=tk.W, padx=20)

sbtn = tk.Button(win,text = "Submit",command=btn_show)
sbtn.pack()

win.mainloop()
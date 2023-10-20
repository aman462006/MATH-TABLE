from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import random
import datetime as dt
import os
import openpyxl




window = Tk()
window.title("MathTable")


frame = Frame(window)
frame.grid(row=0,column=0)


#USER INFORMATION
user_info_frame=LabelFrame(frame,text="User Information",font=("Aerial",15))
user_info_frame.grid(row=0 , column=0,padx=20,pady=20,sticky='news')
name_label= Label(user_info_frame,text="Name",font=("Aerial",12))
name_label.grid(row=0,column=0)
name_entry= Entry(user_info_frame,text="Name",width=105)
name_entry.grid(row=0,column=1,columnspan=5)


Standard=Label(user_info_frame,text="Standard",font=("Aerial",11))
Standard.grid(row=1,column=0)
Standard_entry= Entry(user_info_frame,text="Standard")
Standard_entry.grid(row=1,column=1)

Division=Label(user_info_frame,text="Division",font=("Aerial",12))
Division.grid(row=1,column=2)
Division_entry= Entry(user_info_frame,text="Division")
Division_entry.grid(row=1,column=3)

Date=Label(user_info_frame,text="Date :" ,font=("Aerial",12))
Date.grid(row=1,column=4)
date = dt.datetime.now()
date_label = Label(user_info_frame, text=f"{date:%A, %B %d, %Y}",font=("Aerial",12))
date_label.grid(row=1,column=5)



for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10,pady=5)

ans = 0
nqp = -1
nqc = 0
nqi = 0

def num():
    func = function_combobox.get()[-2]
    lvl = level_combobox.get()
    if func==" ":
        func=random.choice(["+","-","/","*"])
    if lvl==" ":
        lvl=random.choice(["EASY","MEDIUM","DIFFICULT"])

    if func=="+" or func == "-":
        if lvl=="EASY":
           return (10,0,func)
        elif lvl=="MEDIUM":
            return (100,10,func)
        elif lvl=="DIFFICULT":
            return (1000,100,func)
    elif func=="*" or func=="/":
        if lvl=="EASY":
            return(10,1,func)
        elif lvl=="MEDIUM":
            return(20,10,func)
        elif lvl=="DIFFICULT":
           return(100,20,func)


# QUESTION_FRAME
def question_generator():
    global ans

    note_label = Label(question_frame,text="If answer is in decimal round it upto two digits.",font=('Times', 15))
    note_label.grid(row=2,column=0,columnspan=2)

    upper_range,lower_range,func = num()
    n1 = random.randint(lower_range,upper_range)
    if func =="/" and level_combobox.get()=="EASY":
        n2 = n1*random.randint(lower_range,upper_range)
        (n1,n2)=(n2,n1)
    else:
        n2 = random.randint(lower_range,upper_range)
        
    question_label = Label(question_frame,text=str(n1)+func+str(n2),font=('Times', 24),width=20)
    question_label.grid(row=1,column=0,padx=50,pady=20)
    ans = round(eval(str(n1)+func+str(n2)),2)
    # print(ans)
    
def enter(jk):
    global ans,nqp,nqc,nqi
    if name_entry.get()=="" or Standard_entry.get()=="" or Division_entry.get()=="" or function_combobox.get()=="" or level_combobox.get()=="":
        messagebox.showwarning(title="Error",message="All information is compulsory.")
    elif answer_entry.get()=="" and nqp!=-1:
        messagebox.showwarning(title="Error",message="Please enter your answer.")
        
    else:
        if nqp == -1:
            question_generator()
            nqp+=1
            answer_entry.delete(0,END)
        elif float(answer_entry.get()) == ans:
            nqp = nqp + 1
            nqc = nqc + 1
            question_generator()
            answer_entry.delete(0,END)
        elif float(answer_entry.get()) != ans:
            nqp = nqp + 1
            nqi= nqi + 1
            answer_entry.delete(0,END)
        # print(ans,nqp,nqc,nqi)
        Label(result_frame,text="No. of Question practice : "+ str(nqp),font=("Aerial",12)).grid(row=0,column=0)

        Label(result_frame,text="No. of Question correct : "+ str(nqc),font=("Aerial",12)).grid(row=1,column=0)

        Label(result_frame,text="No. of Question incorrect :"+ str(nqi),font=("Aerial",12)).grid(row=2,column=0)
        if nqp!=0:
            Label(result_frame,text="Accuracy :"+ str(round((nqc/nqp*100),2)),font=("Aerial",12)).grid(row=3,column=0)
        else:
            Label(result_frame,text="Accuracy : ",font=("Aerial",12)).grid(row=3,column=0)
            
    for widget in result_frame.winfo_children():
        widget.grid_configure(padx=300,pady=10)

question_frame=LabelFrame(frame,text="QUESTION",font=('Aerial', 15))
question_frame.grid(row=1,column=0 ,padx=20,pady=20,sticky="news")



function_combobox=ttk.Combobox(question_frame,values=["Addition(+)","Subtraction(-)","Division(/)","Multiplication(*)","RANDOM( )"],font=('Aerial', 15))
function_combobox.grid(row=0,column=0,padx=70,pady=10)

level_combobox=ttk.Combobox(question_frame , values=["EASY","MEDIUM","DIFFICULT","RANDOM"],font=('Aerial', 15))
level_combobox.grid(row=0,column=1,padx=70,pady=10)


answer_entry=Entry(question_frame,text="Answer",font=('Aerial', 15),width=22)
answer_entry.grid(row=1,column=1,)



# RESULT FRAME
result_frame=LabelFrame(frame,text="RESULT",font=('Aerial', 15))
result_frame.grid(row=2,column=0,padx=20,pady=20,sticky='news')


def quit():
    ans = messagebox.askokcancel("askokcancel", "SAVE") 
    if ans:
        filepath = "S:\MATH TABLE\data.xlsx"
        if not os.path.exists(filepath):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            heading = ["Name","Standard","Divison","Date and Time","No. of question practice","No. of question correct","No. of question incorrect","Accuracy"]
            sheet.append(heading)
            workbook.save(filepath)
            
        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook.active
        sheet.append([name_entry.get(),Standard_entry.get(),Division_entry.get(),date,nqp,nqc,nqi,str(round((nqc/nqp*100),2))])
        workbook.save(filepath)
    window.destroy()




Exit_frame=LabelFrame(frame)
Exit_frame.grid(row=3 , column=0,sticky="news",padx=20,pady=20)
Label(Exit_frame,text="").grid(row=0,column=0,padx=85)
Exit_Button=Button(Exit_frame,text="EXIT",padx=50,width=85,command=lambda: quit())
Exit_Button.grid(row=0,column=0,padx=35,pady=10)





window.bind('<Return>', enter)
  


window.mainloop()
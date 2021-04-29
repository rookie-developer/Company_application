import datetime, xlrd
import smtplib
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from datetime import date
import PIL as p
import PIL.ImageTk as ptk


root = Tk()

root.title("company application")
root.geometry("1080x480")


############################################## to import logo for username ##############################
user_pic_path = "username_logo.png"
user_pic_image = p.Image.open(user_pic_path)
user_pic = ptk.PhotoImage(user_pic_image)

############################################## to import logo for password ##############################
pass_pic_path = "password_logo.png"
pass_pic_image = p.Image.open(pass_pic_path)
pass_pic = ptk.PhotoImage(pass_pic_image)

######################## this function is for importing the path of the file ##############################
def get_path():

    global import_file_path

    import_file_path = filedialog.askopenfilename()

    get_excel(import_file_path)
#####################################################################################################
#####################################################################################################



################################## this fonction is to get the excel file ########################
def get_excel(path):

    today = date.today()
    day =  today.strftime("%d/%m/%Y")
    command = str(day)[:5]

    inputworkbook = xlrd.open_workbook(path)
    inputworksheet = inputworkbook.sheet_by_index(0)

    main_frame = Frame(root,bg='white')
    main_frame.place(relwidth="0.72",relheight="0.6",relx="0.25",rely="0.01")

    my_canvas = Canvas(main_frame)
    my_canvas.pack(side=LEFT,fill=BOTH,expand=1)

    my_scrollbar = ttk.Scrollbar(main_frame,orient=VERTICAL,command=my_canvas.yview)
    my_scrollbar.pack(side=RIGHT,fill=Y)

    my_scrollbar1 = ttk.Scrollbar(main_frame,orient=HORIZONTAL,command=my_canvas.xview)
    my_scrollbar1.pack(side=BOTTOM,fill=X)

    my_canvas.configure(yscrollcommand=my_scrollbar.set)
    my_canvas.configure(xscrollcommand=my_scrollbar1.set)
    my_canvas.bind('<Configure>', lambda e : my_canvas.configure(scrollregion = my_canvas.bbox("all")))

    second_frame = Frame(my_canvas)

    my_canvas.create_window((0,0),window = second_frame , anchor="nw")

    for i in range(int(inputworksheet.nrows)):
            x=0.1+i/10
            y=0.1
            for j in range(int(inputworksheet.ncols)):
                if i == 0:
                    bg= '#fef26a'
                elif j == 0 :
                    bg='#01ccfc'
                else :
                    bg='white'
                a1 = inputworksheet.cell_value(rowx=i, colx=j)
                try :
                    string = datetime.datetime(*xlrd.xldate_as_tuple(a1, inputworkbook.datemode))
                    dat = str(string)
                    string_to_print = dat[0:10]
                    Label(second_frame,height=1,width=15,bg=bg,text=string_to_print).grid(row=i,column=j,pady=3,padx=3)
                except:
                    string = str(inputworksheet.cell_value(i,j))
                    Label(second_frame,height=1,width=15,bg=bg,text=string).grid(row=i,column=j,pady=3,padx=3)
#####################################################################################################
#####################################################################################################




################################## this fonction is to get the notifications  ########################
def get_notification():

    try :
        today = date.today()
        day =  today.strftime("%d/%m/%Y")
        command = str(day)[:5]

        inputworkbook = xlrd.open_workbook(import_file_path)
        inputworksheet = inputworkbook.sheet_by_index(0)

        for i in range(int(inputworksheet.ncols)):
            x=0.1+i/10
            y=0.1
            for j in range(int(inputworksheet.nrows)):
                a1 = inputworksheet.cell_value(rowx=j, colx=i)
                try :
                    string = datetime.datetime(*xlrd.xldate_as_tuple(a1, inputworkbook.datemode))
                    dat = str(string)
                    string_to_print = dat[0:10]
                    date1 = dat[5:7]
                    date2 = dat[8:10]
                    date_excel = date2+"/"+date1
                    if(command == date_excel):
                        response = messagebox.askquestion(title="birthday", message=("today is "+str(inputworksheet.cell_value(j,1))+" birthday !\n would you like to send an email"))
                        if(response == 'yes'):
                            message = mail_text.get(1.0 , END)
                            for v in range(int(inputworksheet.ncols)):
                                mail = str(inputworksheet.cell_value(j,v))
                                if '@' in mail :
                                    mail = mail.replace("*",".") 
                            send_mail(entry_username.get(),entry_password.get(),mail,message)
                        else:
                            continue
                except:
                    string = str(inputworksheet.cell_value(j,i))
    except:
        messagebox.showerror(title="error", message="load the file first")
#####################################################################################################
#####################################################################################################



################################## this fonction is to send mail (auto)  ########################
def auto_send():
    count = 0
    try :
        today = date.today()
        day =  today.strftime("%d/%m/%Y")
        command = str(day)[:5]

        inputworkbook = xlrd.open_workbook(import_file_path)
        inputworksheet = inputworkbook.sheet_by_index(0)

        for i in range(int(inputworksheet.ncols)):
            x=0.1+i/10
            y=0.1
            for j in range(int(inputworksheet.nrows)):
                a1 = inputworksheet.cell_value(rowx=j, colx=i)
                try :
                    string = datetime.datetime(*xlrd.xldate_as_tuple(a1, inputworkbook.datemode))
                    dat = str(string)
                    string_to_print = dat[0:10]
                    date1 = dat[5:7]
                    date2 = dat[8:10]
                    date_excel = date2+"/"+date1
                    if(command == date_excel):
                        count = count + 1
                except:
                    pass
        response = messagebox.askquestion(title="birthday", message=("there is "+str(count)+' individuals birthday\n would you like to send them an automated mail ?'))
        if(response == 'yes'):
            for i in range(int(inputworksheet.ncols)):
                x=0.1+i/10
                y=0.1
                for j in range(int(inputworksheet.nrows)):
                    a1 = inputworksheet.cell_value(rowx=j, colx=i)
                    try :
                        string = datetime.datetime(*xlrd.xldate_as_tuple(a1, inputworkbook.datemode))
                        dat = str(string)
                        string_to_print = dat[0:10]
                        date1 = dat[5:7]
                        date2 = dat[8:10]
                        date_excel = date2+"/"+date1
                        if(command == date_excel):
                            count = count - 1
                            message = mail_text.get(1.0 , END)
                            for v in range(int(inputworksheet.ncols)):
                                mail = str(inputworksheet.cell_value(j,v))
                                if '@' in mail :
                                    mail = mail.replace("*",".") 
                            auto_send_mail(entry_username.get(),entry_password.get(),mail,message)
                            if(count == 0):
                                messagebox.showinfo(title="success", message="all mails have been sent successfully")
                        else:
                            continue
                    except:
                        pass                
    except:
        messagebox.showerror(title="error", message="load the file first")
#####################################################################################################
#####################################################################################################




#################################### this function is to send the email (individuely) ###############
def send_mail(username,password,email,msg):
    try:
        try:
            server = smtplib.SMTP('smtp.gmail.com:587')  ## gmail code is 587 (google the others)
        except:
            server = smtplib.SMTP('smtp-mail.outlook.com:587')  ## gmail code is 587 (google the others)
        server.ehlo()
        server.starttls()
        server.login(username, password)   ##get in e_mail
        message = msg  ##the message
        server.sendmail(username,email,message)   ##sending the email
        server.quit()
        messagebox.showinfo(title="success", message="email sent successfully")
    except:
        messagebox.showerror(title="error", message="email unsent")
        messagebox.showwarning(title="check message", message="please check your username or password")
#####################################################################################################
#####################################################################################################


#################################### this function is to send the email (automaticly) ###############
def auto_send_mail(username,password,email,msg):
    try:
        try:
            server = smtplib.SMTP('smtp.gmail.com:587')  ## gmail code is 587 (google the others)
        except:
            server = smtplib.SMTP('smtp-mail.outlook.com:587')  ## outlook code is 587 (google the others)
        server.ehlo()
        server.starttls()
        server.login(username, password)   ##get in e_mail
        message = msg  ##the message
        server.sendmail(username,email,message)   ##sending the email
        server.quit()
    except:
        pass
#####################################################################################################
#####################################################################################################




######################################this is design section ########################################
side_frame = Frame(root,bg='#dedddd')
side_frame.place(relwidth="0.2",relheight="0.98",relx="0.01",rely="0.01")

e_mail_frame = Frame(root,bg='#dedddd')
e_mail_frame.place(relwidth="0.77",relheight="0.37",relx="0.22",rely="0.62")

button = Button(side_frame,bg='gray',fg='white',text='import file',activebackground='#fef26a',activeforeground='#01ccfc',font=("Calibri 16"),command=get_path)
button.place(relwidth="0.7",relheight="0.08",relx="0.2",rely="0.65")

button_notification = Button(side_frame,bg='gray',fg='white',text='individuel mail',activebackground='#fef26a',activeforeground='#01ccfc',font=("Calibri 16"),command=get_notification)
button_notification.place(relwidth="0.7",relheight="0.08",relx="0.2",rely="0.55")

button_notification = Button(side_frame,bg='gray',fg='white',text='auto mail sender',activebackground='#fef26a',activeforeground='#01ccfc',font=("Calibri 16"),command=auto_send)
button_notification.place(relwidth="0.7",relheight="0.08",relx="0.2",rely="0.45")

Label_username = Label(side_frame,image=user_pic)
Label_username.place(relx="0.1",rely="0.2")

Label_username = Label(side_frame,image=pass_pic)
Label_username.place(relx="0.1",rely="0.3")

entry_username = Entry(side_frame , bg='gray',fg='white',font=("Calibri 15"),borderwidth=3)
entry_username.place(relwidth="0.7",relheight="0.07",relx="0.25",rely="0.2")

entry_password = Entry(side_frame , bg='gray',fg='white',font=("Calibri 15"),borderwidth=3,show='*')
entry_password.place(relwidth="0.7",relheight="0.07",relx="0.25",rely="0.3")

Label_info = Label(side_frame,text='please enter both \n the e-mail and the password \n in order for the mail to be sent',bg='#dedddd',fg='black')
Label_info.place(relx="0.17",rely="0.8")

mail_text = Text(e_mail_frame)
mail_text.place(relwidth="0.895",relheight="0.98",relx="0.1",rely="0.01")

mail_label = Label(e_mail_frame,text='enter mail\nhere :',bg='#dedddd')
mail_label.place(relx="0.001",rely="0.01")

#####################################################################################################
#####################################################################################################

root.mainloop()
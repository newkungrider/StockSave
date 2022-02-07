import tkinter as tk
from datetime import datetime
import threading
import gspread 
import datetime
from oauth2client.service_account import ServiceAccountCredentials 
from time import sleep
from tkinter import *
from tkinter import ttk
import yfinance as yf
from tkinter import messagebox
T=""
x = "กำลังทำงาน"
p = "กำลังบันทึก"
y = "หยุดทำงาน"
s = "บันทึกสำเร็จ"
chacktime = T
wait_time = 1
lbsh=chacktime
filename = 'ONE-UGG-RA'

#เชื่อมต่อกับชีท
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(r"C:/creds.json", scope)
client = gspread.authorize(creds) 
worksheet1 = client.open(filename).sheet1
worksheet2 = client.open(filename).get_worksheet(1)
# worksheet = client.open(filename).sheet1

def startProgram():
    doit = True
    btRun["state"] = "disabled"
    chacktime = _select_date_time()  
    while btRun["state"] == "disabled":
        lb2["text"] = x
        # lbsh["text"] = chacktime
        timenow = (str(datetime.datetime.now().strftime("%X")))
        if timenow == chacktime and doit == True :
            updata = []
            name = worksheet2.row_values(1)
            Fupdata = [str(datetime.datetime.now())[0:10]]
            for i in range (len(name)) :
                comm = yf.download(tickers = name[i],period='1d')
                z = 2
                while comm.empty :
                    comm = yf.download(tickers = name[i],period=str(str(z)+'d'))
                    z = z+1
                namekey = list(comm.keys())
                updata.append(str(name[i]))
                for i in range (len(namekey)) :
                     updata.append(float(comm.get(namekey[i]).values))
                    # Fupdata.append(float(comm.get(namekey[i]).values))
            # df_data=pd.DataFrame(TSLA)
            # df_data.describe()
            # df_data['Close']
            # xyz =  { 'Open' : str(TSLA['Open'].values) , 'High' : str(TSLA['High'].values), 'Low' : str(TSLA['Low'].values), 'Close' : str(TSLA['Close'].values), 'Adj Close' : str(TSLA['Adj Close'].values), 'Volume' : str(TSLA['Volume'].values) }
            # str(xyz['Open'])
            # de = xyz['Open'].replace("[", '').replace("]", '')
            # updata = [str(datetime.datetime.now())[0:10],float(xyz['Open'].replace("[", '').replace("]", '')), float(xyz['High'].replace("[", '').replace("]", '')), float(xyz['Low'].replace("[", '').replace("]", '')), float(xyz['Close'].replace("[", '').replace("]", '')), float(xyz['Adj Close'].replace("[", '').replace("]", '')), float(xyz['Volume'].replace("[", '').replace("]", ''))]
            lb2["text"] = p
            target_row = 3
            worksheet1.insert_row(Fupdata,target_row)
            if worksheet1.col_count <= ((len(name)*7)+1) :
                worksheet1.add_cols(((len(name)*7)+1)-worksheet1.col_count)
            for i in range (len(updata)) :
                worksheet1.update_cell(target_row,(i+2),updata[i])
                sleep(1)
            doit = False
            lb2["text"] = s
            sleep(10)

        elif timenow == chacktime and doit == False :
            sleep(wait_time)

        elif timenow != chacktime and doit == False :
            doit = True
            sleep(wait_time)

        elif timenow != chacktime and doit == True :
            print('Timenow :' + timenow)
            print('Chacktime :' + chacktime)
            sleep(wait_time)

        else :
            print('Timenow :' + timenow)
            print('Chacktime :' + chacktime)
            sleep(wait_time)

def stopProgram():
    lb2["text"] = y
    btRun["state"] = "normal"

def show():
    if hour_combobox.get() == '' or hour_combobox.get().replace(' ','') == '' and minutes_combobox.get() == '' or minutes_combobox.get().replace(' ','') == ''and seconds_combobox.get() == '' or seconds_combobox.get().replace(' ','') == ''  :
        msg="กรุณาตั้งเวลา"
        messagebox.showinfo("กรุณาตั้งเวลา",msg)
    else:
        _select_date_time(),runThread()
 
def openNewWindow():
    newWindow = Toplevel(master=window)
    newWindow.title("เพิ่มจำนวชื่อหุ้น")
    newWindow.resizable(0,0) 
    def insert():

        if entry.get() == '' or entry.get().replace(' ','') == '' :
            msg="กรุณากรอกใหม่"
            messagebox.showinfo("กรุณากรอกใหม่",msg)
        else:
            
            comm = yf.download(tickers = entry.get(),period='10d')
            if comm.empty :
                msg="บันทึก"+entry.get()+"ไม่สำเร็จกรุณากรอกใหม่"
                messagebox.showinfo("กรุณากรอกใหม่",msg)
            else :
                name = worksheet2.row_values(1)
                add = entry.get()
                if worksheet2.col_count <= (len(name)+1) :
                    worksheet2.add_cols((len(name)+1)-worksheet2.col_count)
                worksheet2.update_cell(1,(len(name)+1),add)
                msg="บันทึก"+entry.get()+"สำเร็จ"
                messagebox.showinfo("บันทึกสำเร็จ",msg,)
                newWindow.destroy()
                
             
    

    LS=Label(newWindow,text ="เพิ่มจำนวนชื่อหุ้น")
    LS.grid(row=0, column=0, padx=(20,5), pady=(20,5))
    LS.config(font=("Tahoma", 28))
    ES=Entry(newWindow,textvariable=entry)
    ES.grid(row=1, column=0, padx=(5,5), pady=(20,5))
    ES.config(font=("Tahoma", 25))
    BS=Button(newWindow,text="บันทึก", command = insert)
    BS.grid(row=1, column=1, padx=(20,5), pady=(20,5))
    BS.config(font=("Tahoma", 25))

def runThread():
    
    x = threading.Thread(target=startProgram)
    x.start()
# def test0():
#     a = _select_date_time()
#     print('a=%d'%a)

def _select_date_time():
    hour = hour_combobox.get()
    minutes = minutes_combobox.get()
    seconds = seconds_combobox.get()
    time25 = ""
    date_time = {}

    if int(hour) < 10:
        hour = '0' + hour
    
    if int(minutes) < 10:
        minutes = '0' + minutes
    
    if int(seconds) < 10:
        seconds = '0' + seconds

    date_time['hour'] = hour
    date_time['minutes'] = minutes
    date_time['seconds'] = seconds

    date = date_time
    if date:
        hour = date['hour']
        minutes = date['minutes']
        seconds = date['seconds']

        time2 = hour + ':' + minutes + ':' + seconds
        time25 = time2
    return(time25)

# def my_time():
#     time_string = strftime('%H:%M:%S') # time format 
#     l1.config(text=time_string)
#     l1.after(1000,my_time) # time delay of 1000 milliseconds 
	
#GUI โปรแกรม
window = tk.Tk()
entry = StringVar()  
window.title("โปรแกรมบันทึกข้อมูลหุ้นอัตโนมัติ")
window.resizable(0,0)  #don't allow resize in x or y direction
lb1 = tk.Label(master=window, text="โปรแกรมบันทึกข้อมูลหุ้นอัตโนมัติ")
lb1.grid(row=0, padx=(15,15), pady=(5,5))  #padx=(left,right), pady=(top,button)
lb1.config(font=("Tahoma", 28))

lb1 = tk.Label(master=window, text="สถานะระบบ: ")
lb1.grid(row=1, column=0, padx=(20,250), pady=(5,20))  #padx=(left,right), pady=(top,button)
lb1.config(font=("Tahoma", 28))

lb2 = tk.Label(master=window, text="")
lb2.grid(row=1, column=0, padx=(250,20), pady=(5,20))
lb2.config(font=("Tahoma", 28))


lbtn = tk.Label(master=window, text="ตั้งเวลา")
lbtn.grid(row=2, column=0, padx=(20,275), pady=(0,0)) 
lbtn.config(font=("Tahoma",28))

lbtnow = tk.Label(master=window, text="เวลาตอนนี้")
lbtnow.grid(row=2, column=0, padx=(275,20), pady=(0,0)) 
lbtnow.config(font=("Tahoma",28))

lbtns = tk.Label(master=window, text="เวลาบันทึก")
lbtns.grid(row=3, column=0, padx=(275,20), pady=(0,0)) 
lbtns.config(font=("Tahoma",28))


lbh=Label(master=window, text='Hours')
lbh.grid(row=3, column=0, padx=(20,400), pady=(40,5))
hour_combobox = ttk.Combobox(window.master, width=3, values=[*range(0,24)])
hour_combobox.grid(row=3, column=0, padx=(20,400), pady=(0,0))


lbchm=Label(master=window, text=':')
lbchm.grid(row=3, column=0, padx=(20,337), pady=(0,0))
lbchm.config(font=("Tahoma",10))

lbm=Label(master=window, text='Minutes')
lbm.grid(row=3, column=0, padx=(20,275), pady=(40,5))
minutes_combobox = ttk.Combobox(window.master, width=3, values=[*range(0,60)])
minutes_combobox.grid(row=3, column=0, padx=(20,275), pady=(0,0))


lbcms=Label(master=window, text=':')
lbcms.grid(row=3, column=0, padx=(20,212), pady=(0,0))
lbcms.config(font=("Tahoma",10))

lbs=Label(window.master, text='Seconds')
lbs.grid(row=3, column=0, padx=(20,150), pady=(40,5))
seconds_combobox = ttk.Combobox(window.master, width=3,values=[*range(0,60)])
seconds_combobox.grid(row=3, column=0, padx=(20,150), pady=(0,0))


# lbsh = tk.Label(master=window, text="")
# lbsh.grid(row=2, column=1)
# lbsh.config(font=(20))

btRun = tk.Button(master=window, text="start", command=show)
btRun.grid(row=4, column=0, padx=(20,250), pady=(5,20))
btRun.config(font=("Tahoma", 25), width=10)

btStop = tk.Button(master=window, text="stop", command=stopProgram)
btStop.grid(row=4, column=0, padx=(250,20), pady=(5,20))
btStop.config(font=("Tahoma", 25), width=10)

btn = Button(master=window,text ="เพิ่มชื่อหุ้นที่ต้องการบันทึก",command = openNewWindow)
btn.grid(row=5, column=0, padx=(20,5), pady=(5,20))
btn.config(font=("Tahoma",15))

window.mainloop()






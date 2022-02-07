from tkinter import *
from datetime import datetime

window = Tk()
window.title('Clock')
window.geometry('200x60')

lb_clock = Label(font='times 16')
lb_clock.pack(anchor=CENTER, expand=YES)

def tick():
    global curtime
    curtime = datetime.now().time()
    ftime = curtime.strftime('%H:%M:%S')
    lb_clock.config(text=ftime)
    lb_clock.after(1000, tick)       #ให้เรียกฟังก์ชันตัวมันเองทุก 1 วินาที

tick()          #เรียกฟังก์ชันขึ้นมาทำงานครั้งแรก
mainloop()
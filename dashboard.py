# modbus communication
from itertools import count
from multiprocessing.connection import Client
import sqlite3
import pymodbus
import serial
from pymodbus.pdu import ModbusRequest
from pymodbus.client.sync import ModbusSerialClient
from pymodbus.client.sync import ModbusSerialClient as ModbusClient
# from evdev import InputDevice, ecodes, list_devices
from select import select
from pymodbus.client.sync import ModbusTcpClient
from datetime import datetime
from time import sleep
#reports 
from tkcalendar import DateEntry
import pandas as pd
from glob import glob; from os.path import expanduser
import pdfkit
from datetime import date
#graph
import numpy as np
import matplotlib
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import matplotlib.font_manager as font_manager
#tkinter
import string
import subprocess
import threading
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tktooltip import ToolTip
from tkinter import messagebox
root=tk.Tk()
root.title('Sharda Motors')
root.geometry("1600x900")
root.configure(bg="#2E4053")
root.resizable(False, False)
root.iconbitmap('images/wayzon.ico')

# Navbar frame
navbar = tk.Frame(root, width=1600, height=35, bg="#34495E")
navbar.place(x=0, y=0)
heading1 = tk.Label(navbar, text='Sharda Motors', bg="#34495E", fg='#D35400', font=("Arial Bold", 18))
heading1.place(x=5, y=1)
heading2 = tk.Label(navbar, text='Leak Marking machine', bg="#34495E", fg='#E5E8E8', font=("Arial Bold", 15))
heading2.place(x=300, y=5)

# =======================   settings dashboard  =====================



running=True
def start_program(stop_event):
    device_id='ngna_sm2'
    third=0
    while not stop_event.is_set():
        try:
            IP_Address = '127.0.0.1'
            client = ModbusTcpClient(IP_Address)
            connect=client.connect()
            print(connect)
            if connect==True:
                Software_Ready_config()
                sleep(1)
            rr=client.read_holding_registers(address=0,count=40, unit=1)
            first1=rr.registers[0]
            print(first1)
            #first1=1
            print("Getting OK Signal From PLC for scanning.")
            # sleep(1)
            #After getting 1 in first1 wait for scan data.
            if first1==1:
                scan_loop=0
                barcode = ""
                #In while loop till scanning qr code
                while scan_loop == 0:
                    print("In qrcode checking state")
                    barcode='XXSM201563452:235453:2462XXX'
                    if (len (barcode)) > 26:
                        global barcode1
                        barcode1=barcode.replace('X','')
                        sqliteConnection = sqlite3.connect('sharda_motors.db')
                        cursor = sqliteConnection.cursor()
                        cursor.execute("SELECT * FROM scan_data WHERE first_scan = ? and ok_signal = '(OK1)'", [barcode1])
                        records = cursor.fetchall()
                        record_length=len(records)
                        if(record_length == 0 ):
                            split_barcode=list(barcode1.split(":"))
                            varient=split_barcode[0]
                            #send varient to plc for marking selection 
                            if(varient ==' '):
                                client.write_registers(21,1,unit=1)
                            else:
                                client.write_registers(22,1,unit=1)
                            print(barcode1,"Scanned QR code from first scanner")
                            scan_data()
                            scan_done_config()                                            # barcode data send to scan data frame
                            sleep(1)
                            date_time1 = datetime.now()
                            date=date_time1.strftime("%Y-%m-%d %H:%M:%S")
                            try:
                                sqliteConnection = sqlite3.connect('sharda_motors.db')
                                cursor = sqliteConnection.cursor()
                                sqlite_insert_query = """INSERT INTO scan_data(first_scan,first_scan_datetime,device_id,varient) 
                                            VALUES(?,?,?,?)"""
                                count = cursor.execute(sqlite_insert_query,(barcode1,date,device_id,varient))
                                last_id=cursor.lastrowid
                                print(last_id,"Get Last id from sqlite and use it for all update quires bellow")
                                sqliteConnection.commit()
                                cursor.close()
                            except sqlite3.Error as error:
                                print('failed to insert data into sqlite table', error)
                            finally:
                                if sqliteConnection:
                                    sqliteConnection.close()
                            #send scan done bit to plc from register number 20
                            client.write_registers(20,1,unit=1)
                            #sleep(2)
                            #client.write_registers(20,0)
                            
                            scan_loop=1
                            leak_test_loop=0
                            #In while loop till getting data from leak test machine
                            while leak_test_loop == 0:
                                print("Getting OK Signal From PLC for receving leak test data.")
                                try:
                                    # read_register=client.read_holding_registers(address=0,count=40, unit=1)
                                    # second=read_register.registers[1]
                                    # third=read_register.registers[2]
                                    second=1
                                    #After getting 1 in second variable wait for leak test data.
                                    if second == 1:
                                        cycle_start_config()
                                        sleep(1)
                                        client.write_registers(27,0,unit=1)
                                        str_enc= b'<01>:\r\n<01>:10/09/2022 17:40:31\r\n<01>:0.321 bar:(OK):0.019 l/mn\r\n'
                                        decode_string=str_enc.decode('utf8', 'ignore')
                                        global leak_rate,ok,air_presure
                                        if(decode_string!=''):
                                            str_enc=decode_string.replace(':','-')
                                            leak_value = list(str_enc.split("-"))
                                            air_presure=leak_value[5]
                                            ok=leak_value[6]
                                            leak_rate=leak_value[7]
                                            print(leak_value)
                                            print(ok)
                                            print(air_presure)
                                            print(leak_rate)
                                            leak_test()
                                            leak_result_config()                          # leak data send to leak data frame
                                            sleep(1)
                                            sticker_data()                                # sticker data to sticker data frame
                                            data_send_config()
                                            try:
                                                date_time2 = datetime.now()
                                                date2=date_time2.strftime("%Y-%m-%d %H:%M:%S")
                                                sqliteConnection = sqlite3.connect('sharda_motors.db')
                                                cursor = sqliteConnection.cursor()
                                                sql_update_query = """Update scan_data set air_pressure=?,leak_rate=? ,ok_signal = ? , test_datetime= ?  where id = ?"""
                                                data =(air_presure,leak_rate,ok,date2,last_id)
                                                # data = (master_id, slave_id)
                                                cursor.execute(sql_update_query, data)
                                                sqliteConnection.commit() 
                                                cursor.close()
                                                
                                            except sqlite3.Error as error:
                                                print('failed to insert data into sqlite table', error)
                                            finally:
                                                if sqliteConnection:
                                                    sqliteConnection.close()
                                            sleep(4)
                                            destroy_all_data()
                                            destroy_all_config()
                                            if ok=='(OK)':
                                                print("OK")
                                                displaying_data()
                                                fetch_todays_count()
                                                graph()
                                            else:
                                                sleep(2)
                                            print(third)
                                            third=1
                                            leak_test_loop = 1
                                            while third == 0:
                                                try:
                                                    read_register=client.read_holding_registers(address=0,count=40, unit=1)
                                                    third=read_register.registers[2]
                                                    print(third)
                                                    
                                                    print("In Cycle complete state")
                                                    client.write_registers(20,1,unit=1)
                                                    #sleep(1)
                                                except:
                                                    print("In cycle complete state.Connection not made with PLC")
                                                    #sleep(1)
                                        else:
                                            print('Leak velue missed, request for retest')
                                            client.write_registers(27,1,unit=1)
                                            sleep(2)
                                            read_register=client.read_holding_registers(address=0,count=40, unit=1)
                                            print(read_register.registers[27])
                                            client.write_registers(27,0,unit=1)
                                            
                                            sleep(2)
                                            client.write_registers(27,0,unit=1)
                                except:
                                    print("In Leak Test Loop.Connection not made with PLC")
                                    if stop_event.is_set():
                                            return
                        else:
                            client.write_registers(25,1,unit=1)    
        except:        
            print("Getting OK Signal From PLC for scanning. Connection not made withPLC ")
            if stop_event.is_set():
                    return
# modbus program start /stop          
# def start_program():
#     global running
#     running = True
#     start_program1()

def stop_program():
    global running
    running = False
    print("stop program")
    stop_event.set()
    root.destroy()
     
def destroy_all_data():
    leak_rate_l.destroy()
    ok_l.destroy()
    air_presure_l.destroy()
    barcode_l.destroy()
    sticker_l.destroy()
    drawing_rev_no_l.destroy()
    vendor_code_l.destroy()


def leak_test():
    global leak_rate_l,ok_l,air_presure_l
    leak_rate_l=Label(leak_test_frame,text=f'{leak_rate}',font=("Arial Bold", 19),fg='#5B2C6F')
    leak_rate_l.place(x=80,y=50)
    if ok=='(OK)':
        ok_l=Label(leak_test_frame,text=f'{ok}',font=("Arial Bold", 19),fg='#5B2C6F')
    else:
        ok_l=Label(leak_test_frame,text=f'{ok}',font=("Arial Bold", 19),fg='tomato')
    ok_l.place(x=105,y=90)
    air_presure_l=Label(leak_test_frame,text=f'{air_presure}',font=("Arial Bold", 18),fg='#5B2C6F')
    air_presure_l.place(x=80,y=130)

leak_test_frame=Frame(root,width=280,height=230)
leak_test_frame.place(x=20,y=42) 

heading=Label(leak_test_frame,text='Leak Test Value',bg="#1FBEF1",font=("Arial Bold", 13))
heading.place(x=5,y=5)
# ================  scan_data_frame frame  ======================
def scan_data():
    global barcode_l
    # barcode1='SM2010022602:090323:0222'
    barcode_l=Label(scan_data_frame1,text=f'{barcode1}',font=("Arial Bold", 16),fg='#5B2C6F')
    barcode_l.place(x=7,y=80)

scan_data_frame1=Frame(root,width=300,height=230)
scan_data_frame1.place(x=315,y=42)

heading=Label(scan_data_frame1,text='Scan Data',bg="#1FBEF1",font=("Arial Bold", 13))
heading.place(x=5,y=5)

# ================  sticker data frame  ======================
def sticker_data():
    # part_no=f'{varient}'
    # drawing_rev_no=f'{product_code}'
    # vendor_code=f'{month_year}'
    part_no='00571514110251'
    drawing_rev_no='00'
    vendor_code='S75203'
    global sticker_l,drawing_rev_no_l,vendor_code_l
    sticker_l=Label(sticker_data_frame,text=f'Part no:-{part_no}',font=("Arial Bold", 17),fg='#5B2C6F')
    sticker_l.place(x=20,y=50)
    drawing_rev_no_l=Label(sticker_data_frame,text=f'Drawing rev no:-{drawing_rev_no}',font=("Arial Bold", 17),fg='#5B2C6F')
    drawing_rev_no_l.place(x=20,y=90)
    vendor_code_l=Label(sticker_data_frame,text=f'Vendor Code:-{vendor_code}',font=("Arial Bold", 17),fg='#5B2C6F')
    vendor_code_l.place(x=20,y=130)
sticker_data_frame=Frame(root,width=330,height=230)
sticker_data_frame.place(x=630,y=42)

heading=Label(sticker_data_frame,text='Sticker Data',bg="#1FBEF1",font=("Arial Bold", 13))
heading.place(x=5,y=5)


#  ======================   graph frame =================
graph_frame=Frame(root,width=350,height=250,bg='white')
graph_frame.place(x=985,y=41)

def graph():
    from_date=datetime.today().strftime('%Y-%m-%d')
    conn=sqlite3.connect('sharda_motors.db')
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c=conn.execute(f'''select strftime('%H',first_scan_datetime) as time,count(ok_signal) as count from scan_data where ok_signal='(OK)'and first_scan_datetime BETWEEN '{from_date} 00:00:00' AND '{from_date} 24:00:00' group by strftime ('%H',first_scan_datetime);''')
    row = c.fetchall()
    # print(row)
    array=[]
    for r in row:
        # print(dict(r))
        array.append(dict(r))
        
    # print(array)
    array2=[]
    for d in array:
        array2.append(d['count'])
    # print(array2)
    array1=[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23]
    array3 = []
        #diffrence between array1 & array2
    for element in array1:
        if element not in array2:
            array3.append(element)
    # print(array3)
    var=[]
    for i in array3:
        var={'time':f'{i}','count':0}
        # print(var)
        array.append(var)
    # print(array)
    sorted_list=sorted(array,key=lambda x:int(x['time']))
    # print(sorted_list)
    
    keys=[]
    values=[]
    for d in sorted_list:
        keys.append(d['time'])
        values.append(d['count'])
    # print(keys)
    # print(values)

    converte_keys=[]
    for item in keys:
        if item.isdigit():
            converte_keys.append(int(item))
        else:
            converte_keys.append(item)
    # print(converte_keys)
    
    datalst = values
    ff = Figure(figsize=(9,5.5), dpi=42)
    xx = ff.add_subplot(111)
    # ind = np.arange(len(datalst))
    rects1 = xx.bar(converte_keys, datalst, 0.6)
    canvas = FigureCanvasTkAgg(ff, master=graph_frame)
    plt.rcParams.update({'font.size': 25})
    xx.set_title('Graph')
    xx.set_ylabel('count',fontsize=17)
    xx.set_xlabel('timing',fontsize=17)
    # xx.bar_label(rects1, fmt='{:,.0f}',fontsize=17)
    # xx.bar_label(xx.containers[0], label_type='edge', color='blue', rotation=90, fontsize=17, padding=3)
    # xx.bar(datalst, color=['#C94845', '#4958B5', '#49D845', '#777777'])
    font = font_manager.FontProperties(size=15)

    # Set the font size of the x-axis and y-axis labels
    for label in xx.xaxis.get_ticklabels():
        label.set_fontproperties(font)
    for label in xx.yaxis.get_ticklabels():
        label.set_fontproperties(font)
    canvas.draw()
    canvas.get_tk_widget().place(x=0,y=0)
graph()


# ================  leak_test_process frame  ======================
leak_test_process_frame=Frame(root,width=350,height=363.5,bg='#7F8C8D')
leak_test_process_frame.place(x=985,y=300)

heading=Label(leak_test_process_frame,text='Process',bg="#1FBEF1",font=("Arial Bold", 13))
heading.place(x=5,y=5)


def leak_result_config():
    if ok=='(OK)':
        Leak_Result_label.config(text = f"Leak Result:- {ok}",fg='green')
    else:
        Leak_Result_label.config(text = f"Leak Result:- {ok}",fg='tomato')

def cycle_start_config():
    # current_color = cycle_Start_label.cget("fg")
    # next_color = "green" if current_color == "white" else "white"
    # cycle_Start_label.config(fg=next_color)
    # cycle_Start_label.after(500,cycle_start_config)
    cycle_Start_label.config(fg='green')

def data_send_config():
    data_send_label.config(fg='green')
    marking_label.config(fg='green')
def scan_done_config():
    # current_color = scan_label.cget("fg")
    # next_color = "green" if current_color == "white" else "white"
    # scan_label.config(fg=next_color)
    # scan_label.after(500,scan_done_config)
    scan_label.config(fg='green')
def destroy_all_config():
    data_send_label.config(fg='white')
    marking_label.config(fg='white')
    cycle_Start_label.config(fg='white')
    Leak_Result_label.config(text = "Leak Result",fg='white')
    scan_label.config(fg='white')

def Software_stop_config():
    software_ready_label.config(fg='white',text = "Software not ready")
    # current_color = software_ready_label.cget("fg")
    # next_color = "red" if current_color == "white" else "white"
    # software_ready_label.config(fg=next_color)
    # software_ready_label.after(500,Software_stop_config)
def Software_Ready_config():
    software_ready_label.config(fg='green')
    # current_color = software_ready_label.cget("fg")
    # next_color = "green" if current_color == "white" else "white"
    # software_ready_label.config(fg=next_color)
    # software_ready_label.after(1000,Software_Ready_config)

software_ready_label= Label(leak_test_process_frame, text = "Software Ready",font=("Arial Bold", 14),fg='white',bg='#7F8C8D')
software_ready_label.place(x=60,y=35)
scan_label = Label(leak_test_process_frame, text = "Scan Done",font=("Arial Bold", 14),fg='white',bg='#7F8C8D')
scan_label.place(x=60,y=75)
cycle_Start_label = Label(leak_test_process_frame, text = "Cycle Start",font=("Arial Bold", 14),fg='white',bg='#7F8C8D')
cycle_Start_label.place(x=60,y=115)
Leak_Result_label = Label(leak_test_process_frame, text = "Leak Result",font=("Arial Bold", 14),fg='white',bg='#7F8C8D')
Leak_Result_label.place(x=60,y=155)

data_send_label = Label(leak_test_process_frame, text = "Data send to Marking",font=("Arial Bold", 14),fg='white',bg='#7F8C8D')
data_send_label.place(x=60,y=195)
marking_label = Label(leak_test_process_frame, text = "Marking in Progress",font=("Arial Bold", 14),fg='white',bg='#7F8C8D')
marking_label.place(x=60,y=235)


#             ======        reprint button    ============
def reprint():
    print("reprint")
    conn=sqlite3.connect('sharda_motors.db')
    c = conn.cursor()           
    c.execute("select max(id),first_scan from scan_data;")
    row = c.fetchone()
    last_print=row[1]
    print(last_print)
    conn.close()
    # ser = serial.Serial ("/dev/ttyUSB0", 9600)
    # writedata=(f'SIZE 50.8 mm, 25.4 mm\nGAP 3 mm, 0 mm\nDIRECTION 0,0\nREFERENCE 0,0\nOFFSET 0 mm\nSET PEEL OFF\nSET CUTTER OFF\nSET PARTIAL_CUTTER OFF\nSET TEAR ON\nCLS\nQRCODE 258,153,L,5,A,180,M2,S7,"{last_print}"\nPRINT 1,1\n')
    # ser.write(writedata.encode())
    # ser.flush()
    # ser.close()

start_button=Button(leak_test_process_frame,width=8,pady=7,text='Reprint',fg='white',bg='#0C84DC',command=reprint).place(x=200,y=310)
# stop_button=Button(leak_test_process_frame,width=8,pady=7,text='stop',fg='white',bg='red', command=threading.Thread(target=stop_program).start).place(x=200,y=310)

# =================   leak_test_data_show_frame ================
leak_test_show_frame=Frame(root,width=860,height=390)
leak_test_show_frame.place(x=20,y=280)
def fetch():
    from_date=datetime.today().strftime('%Y-%m-%d')
    conn=sqlite3.connect('sharda_motors.db')
    aa=conn.execute(f"SELECT first_scan,first_scan_datetime,air_pressure,leak_rate,ok_signal FROM scan_data WHERE ok_signal='(OK)' and first_scan_datetime BETWEEN '{from_date} 00:00:00' AND '{from_date} 24:00:00'")
    rows1=aa.fetchall()
    return rows1


def displaying_data():
    tv.delete(*tv.get_children())
    global i
    i=1
    for row in fetch():
        tv.insert("",END,values=(i,row[0],row[1],row[2],row[3],row[4]))
        i=i+1


style=ttk.Style()
style.configure("mystyle.Treeview",background='#E8EAEA',rowheight=35,font=('Calibri', 14))
style.configure("mystyle.Treeview.Heading",font=('Times New Roman', 16, "bold"))
style.layout("mystyle.Treeview",[('mystyle.Treeview.treearea',{'sticky':'nswe','border':'5'})])

tree_scroll=Scrollbar(leak_test_show_frame)
tree_scroll.pack(side=RIGHT,fill=Y)

column_names=(0,2,3,4,5,6)
tv=ttk.Treeview(leak_test_show_frame,columns=column_names,show='headings',style="mystyle.Treeview",yscrollcommand=tree_scroll.set)
tv.heading("0",text="No",anchor= CENTER)
tv.column("0",width=45,anchor= CENTER)
# tv.heading("1",text="Sr No",anchor= CENTER)
# tv.column("1",width=60,anchor= CENTER)
tv.heading("2",text="Scan Data",anchor= CENTER)
tv.column("2",width=260,anchor= CENTER)
tv.heading("3",text="Date Time",anchor= CENTER)
tv.column("3",width=220,anchor= CENTER)
tv.heading("4",text="Air Pressure",anchor= CENTER)
tv.column("4",width=150,anchor= CENTER)
tv.heading("5",text="Leak Rate",anchor= CENTER)
tv.column("5",width=125,anchor= CENTER)
tv.heading("6",text="Ok signal",anchor= CENTER)
tv.column("6",width=123,anchor= CENTER)
tree_scroll.config(command=tv.yview)
tv.pack()
displaying_data()


from_date=datetime.today().strftime('%Y-%m-%d')
def fetch_todays_count():
    conn=sqlite3.connect('sharda_motors.db')
    aa=conn.execute(f"SELECT count(*) FROM scan_data WHERE ok_signal='(OK)' and first_scan_datetime BETWEEN '{from_date} 00:00:00' AND '{from_date} 24:00:00'")
    todyas_count=str(aa.fetchall())
    new_string = todyas_count.translate(str.maketrans('','', string.punctuation))
    # print(new_string)
    footer_text.config(fg='black',text=f'todays count:-{new_string}')
    return new_string
    
footer=Frame(root,width=1490,height=33,bg='#2980B9')
footer.place(x=0,y=670)

footer_text=Label(footer, text = f"todays count:-",font=("Arial Bold", 15),fg='black',bg='skyblue')
footer_text.place(x=600,y=0)

fetch_todays_count()
def reports():
    
    global new_window
    new_window = tk.Toplevel(root)
    new_window.geometry("740x650")  # Size of the window 
    new_window.title("Report")  # Adding a title
    sel=tk.StringVar()
    cal=DateEntry(new_window,selectmode='dayt',textvariable=sel,date_pattern='yyyy-mm-dd')
    cal.place(x=170,y=40)
    sel1=tk.StringVar()
    cal1=DateEntry(new_window,selectmode='dayt',textvariable=sel1,date_pattern='yyyy-mm-dd')
    cal1.place(x=350,y=40)
    def my_upd(): # triggered when value of string varaible changes
        if(len(sel.get())>4 or len(sel1.get())>4):
            from_date=sel.get()
            to_date=sel1.get()
            dt=cal.get_date()
            dt1=cal1.get_date()
            # print("from_date date",from_date) # get selected date object from calendar
            # print("to date",to_date)
            now = datetime.now() 
            dt3=dt1.strftime("%d-%B-%Y") #format for MySQL date column 
            dt2=dt.strftime("%d-%B-%Y") #format to display at label 
            l1.config(text=dt2+" To "+dt3) # display date at Label
            conn=sqlite3.connect('sharda_motors.db')
            global query
            query=f"SELECT * FROM scan_data WHERE ok_signal='(OK)' and first_scan_datetime BETWEEN '{from_date} 00:00:00' AND '{to_date} 24:00:00'"
            r_set=conn.execute(query) # execute query with data
            for item in trv.get_children(): # delete all previous listings
                trv.delete(item)
            # to store total sale of the selected date
            i=1
            for dt in r_set:
                # print(dt)
                trv.insert("", 'end',iid=dt[0], text=dt[0],
                values =(i,dt[1],dt[2],dt[3],dt[4],dt[5]))
                i=i+1
                # show total value
    l1=tk.Label(new_window,font=('Times',20,'bold'),fg='blue')
    l1.place(x=150,y=80)
    l3=tk.Label(new_window,font=('Times',20,'bold'),text="To",fg='black').place(x=285,y=33)
    tk.Button(new_window, text = "Get data", command = my_upd,fg='green').place(x=490,y=37)

    # tree_scroll1=Scrollbar(new_window)
    # tree_scroll1.pack(side=RIGHT,fill=Y)
    trv = ttk.Treeview(new_window, selectmode ='browse')

    trv.place(x=70,y=120)
    # number of columns
    trv["columns"] = ("0","1", "2", "3","4","5")
    trv['height']  =20
    # Defining heading
    trv['show'] = 'headings'
    
    # width of columns and alignment 
    trv.column("0", width = 30, anchor ='c')
    trv.column("1", width = 160, anchor ='c')
    trv.column("2", width = 150, anchor ='c')
    trv.column("3", width = 90, anchor ='c')
    trv.column("4", width = 80, anchor ='c')
    trv.column("5", width = 90, anchor ='c')
    
    # Headings  
    # respective columns
    trv.heading("0", text ="id")
    trv.heading("1", text ="scan data")
    trv.heading("2", text ="date time")
    trv.heading("3", text ="air pressure")
    trv.heading("4", text ="leak rate")  
    trv.heading("5", text ="ok")

    sel.trace('w',my_upd)
    # tree_scroll1.config(command=trv.yview)

    today = date.today()
    today_date = today.strftime("%d_%m")
    def download_pdf():
        global query
        conn = sqlite3.connect('sharda_motors.db')
        cursor = conn.cursor()
        clients = pd.read_sql(query ,conn)
        clients.to_excel(f"example{today_date}.xlsx")
        # clients.to_csv(f'csvdata{today_date}.csv', index=False)

        # df1 = pd.read_csv(f'csvdata{today_date}.csv')
        # print("The dataframe is:")
        # html_string = df1.to_html()
        # pdfkit.from_string(html_string, f"output_file{today_date}.pdf")

    l2=tk.Label(new_window,font=('Times',22,'bold'),fg='red')
    l2.grid(row=1,column=2,sticky='ne',pady=20)
    tk.Button(new_window, text = "download excel", command = download_pdf,fg='white',bg='green').place(x=300,y=560)

    
    # new_window = subprocess.run(["python", "show_data20_3.py"], capture_output=True, text=True)
    # print(new_window.stdout)
reports=Button(footer,width=6,pady=2,text='Reports',fg='black',font=("Arial Bold", 12),bg='skyblue',command=reports).place(x=350,y=0)
# start_func=threading.Thread(target=start_program).start()  # strat modbus program

root.protocol("WM_DELETE_WINDOW", stop_program)

stop_event = threading.Event()
func1_thread = threading.Thread(target=start_program, args=(stop_event,))
func1_thread.start()

root.mainloop()
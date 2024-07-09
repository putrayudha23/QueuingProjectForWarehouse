import customtkinter
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import ttk
# from tkinter import ttk
from tkinter import messagebox
import time
# from datetime import datetime
import threading

from gtts import gTTS
from playsound import playsound
import os

import serial
import serial.tools.list_ports

# database
import pyodbc
import sqlite3 as sq

class AppKios(customtkinter.CTk):
    #UI
    def __init__(self):
        super().__init__()

        self.title("AKR AppKios Queuing")
        # self.geometry("1024x768")
        self.geometry("950x768")
        customtkinter.set_appearance_mode("Light")
        
        # set grid layout
        self.grid_rowconfigure((0,2), weight=1)
        self.grid_columnconfigure(0, weight=1)

        # image
        self.home_image = Image.open("./image_kios/bg_kios.jpg")
        self.buttonInput_image = tk.PhotoImage(file = "./image_kios/button_input.png")

        # untuk input touchscreen
        global exp, equation
        exp = ""

        # connect local database (sqlite3)
        self.database()

        # frame bg
        self.bg_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="#0e0440")
        self.bg_frame.grid(row=0, column=0, sticky="nsew")
        self.bg_label = customtkinter.CTkLabel(self.bg_frame,text="")
        self.bg_label.pack(expand=tk.YES, fill=tk.BOTH)
        self.bg_frame.bind("<Configure>", self.resizer_bg)
        # frame entry
        self.entry_plat_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="#0e0440")
        self.entry_plat_frame.grid(row=1, column=0, sticky="nsew")
        equation = tk.StringVar()
        self.entry_plat = customtkinter.CTkEntry(self.entry_plat_frame, placeholder_text="", fg_color="#0e0440", text_color="white", height=130, font=customtkinter.CTkFont("bold",size=50), corner_radius=0, placeholder_text_color="#0e0440", justify=tk.CENTER, textvariable=equation)
        self.entry_plat.pack(expand=tk.YES, fill=tk.BOTH)
        # frame virtual keyboard
        self.vKeyboard_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="#2a097d")
        self.vKeyboard_frame.grid(row=2, column=0, sticky="nsew")
        self.vKeyboard_frame.grid_rowconfigure(0, weight=1)
        self.vKeyboard_frame.grid_columnconfigure(0, weight=1)
        self.keyBtn_frame = customtkinter.CTkFrame(self.vKeyboard_frame, corner_radius=0, fg_color="transparent")
        self.keyBtn_frame.grid(row=0, column=0, pady=10)
        num0 = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="0", hover_color="gray", command=lambda: self.press('0'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        num0.grid(row=0, column=0, padx=(10,5), pady=5)
        num1 = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="1", hover_color="gray", command=lambda: self.press('1'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        num1.grid(row=0, column=1, padx=5, pady=5)
        num2 = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="2", hover_color="gray", command=lambda: self.press('2'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        num2.grid(row=0, column=2, padx=5, pady=5)
        num3 = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="3", hover_color="gray", command=lambda: self.press('3'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        num3.grid(row=0, column=3, padx=5, pady=5)
        num4 = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="4", hover_color="gray", command=lambda: self.press('4'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        num4.grid(row=0, column=4, padx=5, pady=5)
        num5 = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="5", hover_color="gray", command=lambda: self.press('5'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        num5.grid(row=0, column=5, padx=5, pady=5)
        num6 = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="6", hover_color="gray", command=lambda: self.press('6'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        num6.grid(row=0, column=6, padx=5, pady=5)
        num7 = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="7", hover_color="gray", command=lambda: self.press('7'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        num7.grid(row=0, column=7, padx=5, pady=5)
        num8 = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="8", hover_color="gray", command=lambda: self.press('8'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        num8.grid(row=0, column=8, padx=5, pady=5)
        num9 = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="9", hover_color="gray", command=lambda: self.press('9'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        num9.grid(row=0, column=9, padx=5, pady=5)

        # Define the alphabet
        alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

        # Loop through each letter and create the buttons
        for i, letter in enumerate(alphabet):
            padx_value = (10, 5) if letter in ['A', 'K', 'U'] else 5  # Set padx value based on letter
            button = customtkinter.CTkButton(
                self.keyBtn_frame,
                fg_color="white",
                text=letter,
                hover_color="gray",
                command=lambda l=letter: (self.press(l), self.entry_plat.icursor(tk.END)),  # Passing letter as a default argument to lambda
                height=60,
                width=60,
                text_color="black",
                font=customtkinter.CTkFont("bold", size=30)
            )
            button.grid(row=i // 10 + 1, column=i % 10, padx=padx_value, pady=5)  # Use integer division and modulus to place buttons

        # A = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="A", hover_color="gray", command=lambda: self.press('A'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # A.grid(row=1, column=0, padx=(10,5), pady=5)
        # B = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="B", hover_color="gray", command=lambda: self.press('B'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # B.grid(row=1, column=1, padx=5, pady=5)
        # C = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="C", hover_color="gray", command=lambda: self.press('C'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # C.grid(row=1, column=2, padx=5, pady=5)
        # D = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="D", hover_color="gray", command=lambda: self.press('D'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # D.grid(row=1, column=3, padx=5, pady=5)
        # E = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="E", hover_color="gray", command=lambda: self.press('E'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # E.grid(row=1, column=4, padx=5, pady=5)
        # F = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="F", hover_color="gray", command=lambda: self.press('F'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # F.grid(row=1, column=5, padx=5, pady=5)
        # G = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="G", hover_color="gray", command=lambda: self.press('G'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # G.grid(row=1, column=6, padx=5, pady=5)
        # H = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="H", hover_color="gray", command=lambda: self.press('H'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # H.grid(row=1, column=7, padx=5, pady=5)
        # I = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="I", hover_color="gray", command=lambda: self.press('I'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # I.grid(row=1, column=8, padx=5, pady=5)
        # J = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="J", hover_color="gray", command=lambda: self.press('J'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # J.grid(row=1, column=9, padx=5, pady=5)
        # K = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="K", hover_color="gray", command=lambda: self.press('K'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # K.grid(row=2, column=0, padx=(10,5), pady=5)
        # L = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="L", hover_color="gray", command=lambda: self.press('L'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # L.grid(row=2, column=1, padx=5, pady=5)
        # M = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="M", hover_color="gray", command=lambda: self.press('M'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # M.grid(row=2, column=2, padx=5, pady=5)
        # N = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="N", hover_color="gray", command=lambda: self.press('N'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # N.grid(row=2, column=3, padx=5, pady=5)
        # O = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="O", hover_color="gray", command=lambda: self.press('O'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # O.grid(row=2, column=4, padx=5, pady=5)
        # P = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="P", hover_color="gray", command=lambda: self.press('P'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # P.grid(row=2, column=5, padx=5, pady=5)
        # Q = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="Q", hover_color="gray", command=lambda: self.press('Q'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # Q.grid(row=2, column=6, padx=5, pady=5)
        # R = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="R", hover_color="gray", command=lambda: self.press('R'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # R.grid(row=2, column=7, padx=5, pady=5)
        # S = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="S", hover_color="gray", command=lambda: self.press('S'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # S.grid(row=2, column=8, padx=5, pady=5)
        # T = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="T", hover_color="gray", command=lambda: self.press('T'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # T.grid(row=2, column=9, padx=5, pady=5)
        # U = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="U", hover_color="gray", command=lambda: self.press('U'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # U.grid(row=3, column=0, padx=(10,5), pady=5)
        # V = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="V", hover_color="gray", command=lambda: self.press('V'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # V.grid(row=3, column=1, padx=5, pady=5)
        # W = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="W", hover_color="gray", command=lambda: self.press('W'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # W.grid(row=3, column=2, padx=5, pady=5)
        # X = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="X", hover_color="gray", command=lambda: self.press('X'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # X.grid(row=3, column=3, padx=5, pady=5)
        # Y = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="Y", hover_color="gray", command=lambda: self.press('Y'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # Y.grid(row=3, column=4, padx=5, pady=5)
        # Z = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="Z", hover_color="gray", command=lambda: self.press('Z'), height=60, width=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        # Z.grid(row=3, column=5, padx=5, pady=5)
            
        Hapus = customtkinter.CTkButton(self.keyBtn_frame, fg_color="white", text="Hapus", hover_color="gray", command=self.Backspace, height=60, text_color="black", font=customtkinter.CTkFont("bold",size=30))
        Hapus.grid(row=3, column=6, columnspan=4, padx=5, pady=5, sticky="nsew")

        self.inputBtn_frame = customtkinter.CTkFrame(self.vKeyboard_frame, corner_radius=0, fg_color="transparent")
        self.inputBtn_frame.grid(row=0, column=1, sticky="nsew")
        self.input_button = tk.Button(self.inputBtn_frame, text = 'Click Me !', image = self.buttonInput_image, bg="#2a097d", border=0, activebackground="#2a097d", command=self.input_noPlat)
        self.input_button.grid(row=0, column=0, pady=10, padx=(20,25))
        self.checkBox_frame = customtkinter.CTkFrame(self.inputBtn_frame, corner_radius=0, fg_color="transparent")
        self.checkBox_frame.grid(row=1, column=0)
        self.checkBox_loading = customtkinter.CTkCheckBox(self.checkBox_frame, text="Loading", width=20, onvalue="Loading", offvalue="off Loading", text_color="white", border_color="white",command=self.checkbox_event)
        self.checkBox_loading.pack(side="left", padx=(25,20), pady=10)
        self.checkBox_unloading = customtkinter.CTkCheckBox(self.checkBox_frame, text="Unloading", width=20, onvalue="Unloading", offvalue="off Unloading", text_color="white", border_color="white",command=self.checkbox_event)
        self.checkBox_unloading.pack(side="left", pady=10, padx=(0,40))
        # frame button
        self.button_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="#30186e")
        self.button_frame.grid(row=3, column=0, sticky="nsew")
        self.openOperator_button = customtkinter.CTkButton(self.button_frame, fg_color="white", text="Open TV Display", hover_color="gray", command=self.openOperator_window, height=40, text_color="black")
        self.openOperator_button.pack(side="right", pady=10, padx=50)

        self.set_noantrian()

        # when close stop thread
        self.wm_protocol("WM_DELETE_WINDOW", self.stop_main)

    def resizer_bg(self,e):
        global resize_bg
        bg_image = Image.open("./image_kios/bg_kios.jpg")
        resize_bg = bg_image.resize((e.width, e.height), Image.Resampling.LANCZOS)
        resize_bg = ImageTk.PhotoImage(resize_bg)
        self.bg_label.configure(image=resize_bg)
        self.entry_plat.focus()

    def press(self,num):
        global exp
        exp = exp + str(num)
        equation.set(exp)

    def Backspace(self):
        global exp
        exp = exp[:-1]
        equation.set(exp)

    def database(self):
        global driver_db, server_db, dbname, username, password

        file = sq.connect("./dataBase/dataBase.db")
        c = file.cursor()
        c.execute("SELECT * FROM tb_database")
        rec = c.fetchall()
        for i in rec:
            driver_db = "ODBC Driver 17 for SQL Server"
            server_db = i[1]
            dbname = i[2]
            username = i[3]
            password = i[4]
        file.commit()
        file.close()

        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # buat tabel vetting truck
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_vetting')
            CREATE TABLE tb_vetting(
                company nvarchar(255),
                no_polisi nvarchar(255),
                reg_date nvarchar(255),
                expired_date nvarchar(255),
                remark nvarchar(255),
                status nvarchar(255),
                rfid1 nvarchar(255),
                rfid2 nvarchar(255),
                category nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel unloading truck
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_unloading')
            CREATE TABLE tb_unloading(
                company nvarchar(255),
                no_polisi nvarchar(255),
                packaging nvarchar(255),
                product nvarchar(255),
                date nvarchar(255),
                time nvarchar(255),
                remark nvarchar(255),
                status nvarchar(255),
                quota nvarchar(255),
                berat nvarchar(255),
                wh_id nvarchar(255),
                no_so nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel loading list
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_loadingList')
            CREATE TABLE tb_loadingList(
                company nvarchar(255),
                no_polisi nvarchar(255),
                packaging nvarchar(255),
                product nvarchar(255),
                date nvarchar(255),
                time nvarchar(255),
                remark nvarchar(255),
                status nvarchar(255),
                quota nvarchar(255),
                no_so nvarchar(255),
                satuan nvarchar(255),
                berat nvarchar(255),
                wh_id nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel displan data
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_displanData')
            CREATE TABLE tb_displanData(
                company nvarchar(255),
                truck nvarchar(255),
                packaging nvarchar(255),
                product nvarchar(255),
                date nvarchar(255),
                time nvarchar(255),
                pkg_remark nvarchar(255),
                remark nvarchar(255),
                no_so nvarchar(255),
                tr_part nvarchar(255),
                satuan nvarchar(255),
                berat nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel packaging registration
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_packagingReg')
            CREATE TABLE tb_packagingReg(
                packaging nvarchar(255),
                prefix nvarchar(255),
                remark nvarchar(255),
                status nvarchar(255),
                sla nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel warehouse sla
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_warehouseSla')
            CREATE TABLE tb_warehouseSla(
                packaging nvarchar(255),
                product_grp nvarchar(255),
                weight_min nvarchar(255),
                weight_max nvarchar(255),
                wh_sla nvarchar(255),
                vehicle nvarchar(255),
                waiting_time nvarchar(255),
                w_bridge1 nvarchar(255),
                w_bridge2 nvarchar(255),
                covering_time nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel product registration
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_productReg')
            CREATE TABLE tb_productReg(
                packaging nvarchar(255),
                product_grp nvarchar(255),
                remark nvarchar(255),
                status nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel tapping point registration
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_tappingPointReg')
            CREATE TABLE tb_tappingPointReg(
                hw_id nvarchar(255),
                tap_id nvarchar(255),
                remark nvarchar(255),
                packaging nvarchar(255),
                product nvarchar(255),
                activity nvarchar(255),
                block nvarchar(255),
                max nvarchar(255),
                functions nvarchar(255),
                status nvarchar(255),
                alias nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel warehouse flow registration
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_warehouseFlow')
            CREATE TABLE tb_warehouseFlow(
                packaging nvarchar(255),
                activity nvarchar(255),
                registration nvarchar(255),
                weightbridge1 nvarchar(255),
                warehouse nvarchar(255),
                weightbridge2 nvarchar(255),
                covering nvarchar(255),
                unregistration nvarchar(255),
                status nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel waiting list
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_waitinglist')
            CREATE TABLE tb_waitinglist(
                ticket nvarchar(255),
                no_polisi nvarchar(10) UNIQUE,
                activity nvarchar(255),
                arrival nvarchar(255),
                packaging nvarchar(255),
                product nvarchar(255),
                skip nvarchar(255),
                next_call nvarchar(255),
                next_id nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel alern
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_alert')
            CREATE TABLE tb_alert(
                date nvarchar(255),
                time nvarchar(255),
                status nvarchar(255),
                category nvarchar(255),
                message nvarchar(255),
                notes nvarchar(255),
                no_polisi nvarchar(255),
                arrival nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel history
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_history')
            CREATE TABLE tb_history(
                date nvarchar(255),
                time nvarchar(255),
                status nvarchar(255),
                category nvarchar(255),
                message nvarchar(255),
                operator nvarchar(255),
                warehouse nvarchar(255),
                notes nvarchar(255),
                no_polisi nvarchar(255),
                arrival nvarchar(255),
                solving_time nvarchar(255),
                solving_date nvarchar(255),
                solving_total_time nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")
        
        # buat tabel queuing status
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_queuing_status')
            CREATE TABLE tb_queuing_status(
                rfid nvarchar(255),
                ticket nvarchar(255),
                no_polisi nvarchar(10) UNIQUE,
                activity nvarchar(255),
                packaging nvarchar(255),
                product nvarchar(255),
                arrival nvarchar(255),
                calling nvarchar(255),
                pairing nvarchar(255),
                wb1_in nvarchar(255),
                wb1_out nvarchar(255),
                wh_in nvarchar(255),
                wh_out nvarchar(255),
                wb2_in nvarchar(255),
                wb2_out nvarchar(255),
                cc_in nvarchar(255),
                cc_out nvarchar(255),
                unpairing nvarchar(255),
                total_minute nvarchar(255),
                wh_id nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel no antrian
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_noantrian')
            CREATE TABLE tb_noantrian(
                prefix nvarchar(255),
                no_antrian nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel callig time
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_calling_time')
            CREATE TABLE tb_calling_time(
                no_polisi nvarchar(10) UNIQUE,
                calling_time nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")
        
        # buat tabel tb_ticket_skip
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_ticket_skip')
            CREATE TABLE tb_ticket_skip(
                ticket nvarchar(10) UNIQUE,
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")
        
        # buat tabel tb_ticket_history
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_ticket_history')
            CREATE TABLE tb_ticket_history(
                ticket nvarchar(255),
                no_polisi nvarchar(255),
                activity nvarchar(255),
                date nvarchar(255),
                arrival nvarchar(255),
                packaging nvarchar(255),
                product nvarchar(255),
                company nvarchar(255),
                status nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel tb_temporaryRFID
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_temporaryRFID')
            CREATE TABLE tb_temporaryRFID(
                rfid nvarchar(255),
                no_polisi nvarchar(255),
                location nvarchar(255),
                time_detected nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel start container
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_temporaryRFID_start')
            CREATE TABLE tb_temporaryRFID_start(
                rfid nvarchar(255),
                no_polisi nvarchar(10) UNIQUE,
                location nvarchar(255),
                time_detected nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel last container
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_temporaryRFID_last')
            CREATE TABLE tb_temporaryRFID_last(
                rfid nvarchar(255),
                no_polisi nvarchar(10) UNIQUE,
                location nvarchar(255),
                time_detected nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel queuing flow
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_queuing_flow')
            CREATE TABLE tb_queuing_flow(
                no_polisi nvarchar(10) UNIQUE,
                getin nvarchar(255),
                weight_bridge1 nvarchar(255),
                warehouse nvarchar(255),
                cargo_covering nvarchar(255),
                weight_bridge2 nvarchar(255),
                getout nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel calling
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_calling')
            CREATE TABLE tb_calling(
                no_polisi nvarchar(10) UNIQUE,
                call_wb1 nvarchar(255),
                call_wh nvarchar(255),
                call_wb2 nvarchar(255),
                call_cc nvarchar(255),
                call_getout nvarchar(255),
                call_end nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel tb_overSLA
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_overSLA')
            CREATE TABLE tb_overSLA(
                no_polisi nvarchar(10) UNIQUE,
                over_wb1 nvarchar(255),
                over_wh nvarchar(255),
                over_wb2 nvarchar(255),
                over_cc nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tabel tb_queuing_history
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_queuing_history')
            CREATE TABLE tb_queuing_history(
                date nvarchar(255),
                rfid nvarchar(255),
                ticket nvarchar(10),
                no_polisi nvarchar(255),
                activity nvarchar(255),
                packaging nvarchar(255),
                product nvarchar(255),
                arrival nvarchar(255),
                calling nvarchar(255),
                pairing nvarchar(255),
                wb1_in nvarchar(255),
                wb1_out nvarchar(255),
                wh_in nvarchar(255),
                wh_out nvarchar(255),
                wb2_in nvarchar(255),
                wb2_out nvarchar(255),
                cc_in nvarchar(255),
                cc_out nvarchar(255),
                unpairing nvarchar(255),
                total_minute nvarchar(255),
                wh_id nvarchar(255),
                warehouse nvarchar(255),
                berat nvarchar(255),
                waiting_minute nvarchar(255),
                wb1_minute nvarchar(255),
                wh_minute nvarchar(255),
                wb2_minute nvarchar(255),
                cc_minute nvarchar(255),
                company nvarchar(255),
                category nvarchar(255),
                no_so nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tb_displan_data_container
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_displan_data_container')
            CREATE TABLE tb_displan_data_container(
                company nvarchar(10),
                truck nvarchar(255),
                packaging nvarchar(255),
                product nvarchar(255),
                date nvarchar(255),
                time nvarchar(255),
                pkg_remark nvarchar(255),
                remark nvarchar(255),
                no_so nvarchar(255),
                tr_part nvarchar(255),
                satuan nvarchar(255),
                berat nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")
        
        # buat tb_loadingScheduleSetting
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_loadingScheduleSetting')
            CREATE TABLE tb_loadingScheduleSetting(
                allow_h1 BIT,
                allow_h2 BIT,
                allow_h3 BIT,
                allow_h_minus_1 BIT,
                allow_h_minus_2 BIT,
                allow_h_minus_3 BIT
            )""")
        
        # buat tb_userRegistration
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_userRegistration')
            CREATE TABLE tb_userRegistration(
                username nvarchar(255),
                password nvarchar(255),
                email nvarchar(255),
                remark nvarchar(255),
                alertNote BIT,
                reportViewer BIT,
                packaging BIT,
                productReg BIT,
                tappingPoint BIT,
                warehouseFlow BIT,
                rfidManualCall BIT,
                vetting BIT,
                unloading BIT,
                distribution BIT,
                databaseSetting BIT,
                displanApi BIT,
                otherSetting BIT,
                userProfile BIT,
                userRegistration BIT,
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tb_call_rules
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_call_rules')
            CREATE TABLE tb_call_rules(
                delay_repeat nvarchar(255),
                skip_call_time nvarchar(255),
                delete_call bit
            )""")

        # buat tb_log_report
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_log_report')
            CREATE TABLE tb_log_report(
                warehouse nvarchar(255),
                date_log nvarchar(255),
                time_log nvarchar(255),
                transaction_log nvarchar(255),
                operator nvarchar(255),
                status nvarchar(255),
                remark nvarchar(255)
            )""")
        
        # buat tb_playSound
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_playSound')
            CREATE TABLE tb_playSound(
                location nvarchar(255),
                activity nvarchar(255),
                ticket nvarchar(255),
                covering_status nvarchar(255),
                oid INT IDENTITY(1,1) PRIMARY KEY
            )""")

        # buat tb_playSound
        c.execute("""IF NOT EXISTS (SELECT * FROM sys.tables WHERE name='tb_inPlaySound')
            CREATE TABLE tb_inPlaySound(
                play bit
            )""")

        file.commit()
        file.close()

        # Check if the table is empty and insert a value if it is
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT COUNT(*) FROM tb_inPlaySound")
        if c.fetchone()[0] == 0:
            c.execute("INSERT INTO tb_inPlaySound (play) VALUES (0)")  # Inserting 0 or False depending on your preference
        file.commit()
        file.close()

    # ui calling ada disini
    def openOperator_window(self):
        global operator_window
        operator_window = customtkinter.CTkToplevel()
        operator_window.title("AKR AppKios Queuing (Operator Window)")
        operator_window.geometry("1400x768")
        # set grid layout
        operator_window.grid_rowconfigure(1, weight=1)
        operator_window.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use("winnative")
        style.configure("Treeview",
            background = "white",
            foreground = "black",
            rowheight = 30,
            fieldbackground = "white"
        )

        #>>>> frame top
        self.top_frame = customtkinter.CTkFrame(operator_window, corner_radius=20, fg_color="#19154a")
        self.top_frame.grid(row=0, column=0, sticky="nsew", pady=10, padx=10)
        self.top_frame.grid_rowconfigure(0, weight=1)
        self.top_frame.grid_columnconfigure(0, weight=1)
        # display antrian
        self.no_antrian_display_frame = customtkinter.CTkFrame(self.top_frame, corner_radius=20, fg_color="#19154a")
        self.no_antrian_display_frame.grid(row=0, column=0, padx=10, pady=10)
        self.title_no_antrian = customtkinter.CTkLabel(self.no_antrian_display_frame, text="No Antrian", font=customtkinter.CTkFont(size=60,weight="bold"),anchor="center", text_color="yellow")
        self.title_no_antrian.pack()
        self.no_antrian = customtkinter.CTkLabel(self.no_antrian_display_frame, text="XXXX", font=customtkinter.CTkFont(size=70,weight="bold"),anchor="center", text_color="white")
        self.no_antrian.pack(pady=10)
        self.title_no_kendaraan = customtkinter.CTkLabel(self.no_antrian_display_frame, text="No Kendaraan", font=customtkinter.CTkFont(size=40,weight="bold"),anchor="center", text_color="yellow")
        self.title_no_kendaraan.pack()
        self.no_kendaraan = customtkinter.CTkLabel(self.no_antrian_display_frame, text="YYYY", font=customtkinter.CTkFont(size=60,weight="bold"),anchor="center", text_color="white")
        self.no_kendaraan.pack(pady=10)
        self.title_packaging = customtkinter.CTkLabel(self.no_antrian_display_frame, text="Jenis Packaging / Aktifitas", font=customtkinter.CTkFont(size=40,weight="bold"),anchor="center", text_color="yellow")
        self.title_packaging.pack()
        self.packaging = customtkinter.CTkLabel(self.no_antrian_display_frame, text="packaging / aktivitas", font=customtkinter.CTkFont(size=50,weight="bold"),anchor="center", text_color="white")
        self.packaging.pack(pady=10)
        # tabel penugasan warehouse
        self.warehouseTable_display_frame = customtkinter.CTkFrame(self.top_frame, corner_radius=10, fg_color="#19154a")
        self.warehouseTable_display_frame.grid(row=0, column=1, sticky="nsew", padx=(0,10), pady=10)
        self.warehouseTable_display_frame.grid_rowconfigure((1,3), weight=1)
        self.warehouseTable_display_frame.grid_columnconfigure(0, weight=1)

        self.warehouseTable_titelLoading_frame = customtkinter.CTkFrame(self.warehouseTable_display_frame, corner_radius=10, fg_color="transparent")
        self.warehouseTable_titelLoading_frame.grid(row=0, column=0, sticky="nsew")
        self.warehouseTable_titelLoading = customtkinter.CTkLabel(self.warehouseTable_titelLoading_frame, text="Loading", font=customtkinter.CTkFont(size=17),anchor="w", text_color="white")
        self.warehouseTable_titelLoading.grid(row=0, column=0, sticky="nsew", padx=10)
        self.warehouseTable_tableLoading_frame = customtkinter.CTkFrame(self.warehouseTable_display_frame, corner_radius=10, fg_color="transparent")
        self.warehouseTable_tableLoading_frame.grid(row=1, column=0, sticky="nsew")
        
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.warehouse_assignLoading_table = ttk.Treeview(self.warehouseTable_tableLoading_frame, height=8)
        self.warehouse_assignLoading_table.pack(side="left", expand=tk.NO, fill=tk.BOTH, padx=(10,0))
        self.warehouse_assignLoading_table["columns"] = ("Plat Nomor", "Warehouse")
        self.warehouse_assignLoading_table.column("#0", width=0,stretch=tk.NO)
        self.warehouse_assignLoading_table.column("Plat Nomor", anchor=tk.CENTER, width=300,minwidth=0,stretch=tk.YES)
        self.warehouse_assignLoading_table.column("Warehouse", anchor=tk.CENTER, width=300, minwidth=0,stretch=tk.YES)
        self.warehouse_assignLoading_table.heading("#0", text="", anchor=tk.W)
        self.warehouse_assignLoading_table.heading("Plat Nomor", text="Plat Nomor", anchor=tk.CENTER)
        self.warehouse_assignLoading_table.heading("Warehouse", text="Warehouse", anchor=tk.CENTER)
        self.warehouse_assignLoading_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.warehouse_assignLoading_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.warehouse_assignLoading_table_scrollbar = customtkinter.CTkScrollbar(self.warehouseTable_tableLoading_frame, hover=True, button_hover_color="dark blue", command=self.warehouse_assignLoading_table.yview)
        self.warehouse_assignLoading_table_scrollbar.pack(side="left",fill=tk.Y, padx=(0,5))
        self.warehouse_assignLoading_table.configure(yscrollcommand=self.warehouse_assignLoading_table_scrollbar.set)

        self.warehouseTable_titelUnloading_frame = customtkinter.CTkFrame(self.warehouseTable_display_frame, corner_radius=10, fg_color="transparent")
        self.warehouseTable_titelUnloading_frame.grid(row=2, column=0, sticky="nsew")
        self.warehouseTable_titelUnloading = customtkinter.CTkLabel(self.warehouseTable_titelUnloading_frame, text="Unloading", font=customtkinter.CTkFont(size=17),anchor="w", text_color="white")
        self.warehouseTable_titelUnloading.grid(row=0, column=0, sticky="nsew", padx=10)
        self.warehouseTable_tableUnloading_frame = customtkinter.CTkFrame(self.warehouseTable_display_frame, corner_radius=10, fg_color="transparent")
        self.warehouseTable_tableUnloading_frame.grid(row=3, column=0, sticky="nsew")
        
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.warehouse_assignUnloading_table = ttk.Treeview(self.warehouseTable_tableUnloading_frame, height=8)
        self.warehouse_assignUnloading_table.pack(side="left", expand=tk.NO, fill=tk.BOTH, pady=(0,10), padx=(10,0))
        self.warehouse_assignUnloading_table["columns"] = ("Plat Nomor", "Warehouse")
        self.warehouse_assignUnloading_table.column("#0", width=0,stretch=tk.NO)
        self.warehouse_assignUnloading_table.column("Plat Nomor", anchor=tk.CENTER, width=300,minwidth=0,stretch=tk.YES)
        self.warehouse_assignUnloading_table.column("Warehouse", anchor=tk.CENTER, width=300, minwidth=0,stretch=tk.YES)
        self.warehouse_assignUnloading_table.heading("#0", text="", anchor=tk.W)
        self.warehouse_assignUnloading_table.heading("Plat Nomor", text="Plat Nomor", anchor=tk.CENTER)
        self.warehouse_assignUnloading_table.heading("Warehouse", text="Warehouse", anchor=tk.CENTER)
        self.warehouse_assignUnloading_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.warehouse_assignUnloading_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.warehouse_assignUnloading_table_scrollbar = customtkinter.CTkScrollbar(self.warehouseTable_tableUnloading_frame, hover=True, button_hover_color="dark blue", command=self.warehouse_assignUnloading_table.yview)
        self.warehouse_assignUnloading_table_scrollbar.pack(side="left",fill=tk.Y, padx=(0,5), pady=(0,10))
        self.warehouse_assignUnloading_table.configure(yscrollcommand=self.warehouse_assignUnloading_table_scrollbar.set)

        #>>>> frame down
        self.down_frame = customtkinter.CTkFrame(operator_window, corner_radius=0, fg_color="transparent")
        self.down_frame.grid(row=1, column=0, sticky="nsew")
        self.down_frame.grid_rowconfigure(0, weight=1)
        self.down_frame.grid_columnconfigure((0,1,2,3,4), weight=1)
        
        self.loading_zak_frame = customtkinter.CTkFrame(self.down_frame, corner_radius=10, fg_color="light blue")
        self.loading_zak_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=(0,10))
        self.loading_zak_frame.grid_rowconfigure(1, weight=1)
        self.loading_zak_frame.grid_columnconfigure(0, weight=1)
        self.title_loading_zak = customtkinter.CTkLabel(self.loading_zak_frame, text="Loading Zak", font=customtkinter.CTkFont(size=17),anchor="w")
        self.title_loading_zak.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10,0))
        self.tableFrame_loading_zak = customtkinter.CTkFrame(self.loading_zak_frame, corner_radius=10, fg_color="white")
        self.tableFrame_loading_zak.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.loading_zak_tableMenunggu = ttk.Treeview(self.tableFrame_loading_zak)
        self.loading_zak_tableMenunggu.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10, padx=(10,0))
        self.loading_zak_tableMenunggu["columns"] = ("Menunggu")
        self.loading_zak_tableMenunggu.column("#0", width=0,stretch=tk.NO)
        self.loading_zak_tableMenunggu.column("Menunggu", anchor=tk.CENTER, width=100,minwidth=0,stretch=tk.YES)
        self.loading_zak_tableMenunggu.heading("#0", text="", anchor=tk.W)
        self.loading_zak_tableMenunggu.heading("Menunggu", text="Menunggu", anchor=tk.CENTER)
        self.loading_zak_tableMenunggu.tag_configure("oddrow",background="white",font=(None, 13))
        self.loading_zak_tableMenunggu.tag_configure("evenrow",background="lightblue",font=(None, 13))
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.loading_zak_tableDipanggil = ttk.Treeview(self.tableFrame_loading_zak)
        self.loading_zak_tableDipanggil.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10)
        self.loading_zak_tableDipanggil["columns"] = ("Dipanggil")
        self.loading_zak_tableDipanggil.column("#0", width=0,stretch=tk.NO)
        self.loading_zak_tableDipanggil.column("Dipanggil", anchor=tk.CENTER, width=100, minwidth=0,stretch=tk.YES)
        self.loading_zak_tableDipanggil.heading("#0", text="", anchor=tk.W)
        self.loading_zak_tableDipanggil.heading("Dipanggil", text="Dalam Proses", anchor=tk.CENTER)
        self.loading_zak_tableDipanggil.tag_configure("oddrow",background="white",font=(None, 13))
        self.loading_zak_tableDipanggil.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.loading_zak_table_scrollbar = customtkinter.CTkScrollbar(self.tableFrame_loading_zak, hover=True, button_hover_color="dark blue", command=self.loading_zak_tableDipanggil.yview)
        self.loading_zak_table_scrollbar.pack(side="left",fill=tk.Y, pady=10, padx=(0,5))
        self.loading_zak_tableMenunggu.configure(yscrollcommand=self.loading_zak_table_scrollbar.set)
        self.loading_zak_tableDipanggil.configure(yscrollcommand=self.loading_zak_table_scrollbar.set)

        self.loading_curah_frame = customtkinter.CTkFrame(self.down_frame, corner_radius=10, fg_color="light blue")
        self.loading_curah_frame.grid(row=0, column=1, sticky="nsew", padx=(0,10), pady=(0,10))
        self.loading_curah_frame.grid_rowconfigure(1, weight=1)
        self.loading_curah_frame.grid_columnconfigure(0, weight=1)
        self.title_loading_curah = customtkinter.CTkLabel(self.loading_curah_frame, text="Loading Curah", font=customtkinter.CTkFont(size=17),anchor="w")
        self.title_loading_curah.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10,0))
        self.tableFrame_loading_curah = customtkinter.CTkFrame(self.loading_curah_frame, corner_radius=10, fg_color="white")
        self.tableFrame_loading_curah.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.loading_curah_tableMenunggu = ttk.Treeview(self.tableFrame_loading_curah)
        self.loading_curah_tableMenunggu.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10, padx=(10,0))
        self.loading_curah_tableMenunggu["columns"] = ("Menunggu")
        self.loading_curah_tableMenunggu.column("#0", width=0,stretch=tk.NO)
        self.loading_curah_tableMenunggu.column("Menunggu", anchor=tk.CENTER, width=100,minwidth=0,stretch=tk.YES)
        self.loading_curah_tableMenunggu.heading("#0", text="", anchor=tk.W)
        self.loading_curah_tableMenunggu.heading("Menunggu", text="Menunggu", anchor=tk.CENTER)
        self.loading_curah_tableMenunggu.tag_configure("oddrow",background="white",font=(None, 13))
        self.loading_curah_tableMenunggu.tag_configure("evenrow",background="lightblue",font=(None, 13))
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.loading_curah_tableDipanggil = ttk.Treeview(self.tableFrame_loading_curah)
        self.loading_curah_tableDipanggil.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10)
        self.loading_curah_tableDipanggil["columns"] = ("Dipanggil")
        self.loading_curah_tableDipanggil.column("#0", width=0,stretch=tk.NO)
        self.loading_curah_tableDipanggil.column("Dipanggil", anchor=tk.CENTER, width=100, minwidth=0,stretch=tk.YES)
        self.loading_curah_tableDipanggil.heading("#0", text="", anchor=tk.W)
        self.loading_curah_tableDipanggil.heading("Dipanggil", text="Dalam Proses", anchor=tk.CENTER)
        self.loading_curah_tableDipanggil.tag_configure("oddrow",background="white",font=(None, 13))
        self.loading_curah_tableDipanggil.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.loading_curah_table_scrollbar = customtkinter.CTkScrollbar(self.tableFrame_loading_curah, hover=True, button_hover_color="dark blue", command=self.loading_curah_tableDipanggil.yview)
        self.loading_curah_table_scrollbar.pack(side="left",fill=tk.Y, pady=10, padx=(0,5))
        self.loading_curah_tableMenunggu.configure(yscrollcommand=self.loading_curah_table_scrollbar.set)
        self.loading_curah_tableDipanggil.configure(yscrollcommand=self.loading_curah_table_scrollbar.set)

        self.loading_jumbo_frame = customtkinter.CTkFrame(self.down_frame, corner_radius=10, fg_color="light blue")
        self.loading_jumbo_frame.grid(row=0, column=2, sticky="nsew", padx=(0,10), pady=(0,10))
        self.loading_jumbo_frame.grid_rowconfigure(1, weight=1)
        self.loading_jumbo_frame.grid_columnconfigure(0, weight=1)
        self.title_loading_jumbo = customtkinter.CTkLabel(self.loading_jumbo_frame, text="Loading Jumbo", font=customtkinter.CTkFont(size=17),anchor="w")
        self.title_loading_jumbo.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10,0))
        self.tableFrame_loading_jumbo = customtkinter.CTkFrame(self.loading_jumbo_frame, corner_radius=10, fg_color="white")
        self.tableFrame_loading_jumbo.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.loading_jumbo_tableMenunggu = ttk.Treeview(self.tableFrame_loading_jumbo)
        self.loading_jumbo_tableMenunggu.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10, padx=(10,0))
        self.loading_jumbo_tableMenunggu["columns"] = ("Menunggu")
        self.loading_jumbo_tableMenunggu.column("#0", width=0,stretch=tk.NO)
        self.loading_jumbo_tableMenunggu.column("Menunggu", anchor=tk.CENTER, width=100,minwidth=0,stretch=tk.YES)
        self.loading_jumbo_tableMenunggu.heading("#0", text="", anchor=tk.W)
        self.loading_jumbo_tableMenunggu.heading("Menunggu", text="Menunggu", anchor=tk.CENTER)
        self.loading_jumbo_tableMenunggu.tag_configure("oddrow",background="white",font=(None, 13))
        self.loading_jumbo_tableMenunggu.tag_configure("evenrow",background="lightblue",font=(None, 13))
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.loading_jumbo_tableDipanggil = ttk.Treeview(self.tableFrame_loading_jumbo)
        self.loading_jumbo_tableDipanggil.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10)
        self.loading_jumbo_tableDipanggil["columns"] = ("Dipanggil")
        self.loading_jumbo_tableDipanggil.column("#0", width=0,stretch=tk.NO)
        self.loading_jumbo_tableDipanggil.column("Dipanggil", anchor=tk.CENTER, width=100, minwidth=0,stretch=tk.YES)
        self.loading_jumbo_tableDipanggil.heading("#0", text="", anchor=tk.W)
        self.loading_jumbo_tableDipanggil.heading("Dipanggil", text="Dalam Proses", anchor=tk.CENTER)
        self.loading_jumbo_tableDipanggil.tag_configure("oddrow",background="white",font=(None, 13))
        self.loading_jumbo_tableDipanggil.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.loading_jumbo_table_scrollbar = customtkinter.CTkScrollbar(self.tableFrame_loading_jumbo, hover=True, button_hover_color="dark blue", command=self.loading_jumbo_tableDipanggil.yview)
        self.loading_jumbo_table_scrollbar.pack(side="left",fill=tk.Y, pady=10, padx=(0,5))
        self.loading_jumbo_tableMenunggu.configure(yscrollcommand=self.loading_jumbo_table_scrollbar.set)
        self.loading_jumbo_tableDipanggil.configure(yscrollcommand=self.loading_jumbo_table_scrollbar.set)

        self.loading_pallet_frame = customtkinter.CTkFrame(self.down_frame, corner_radius=10, fg_color="light blue")
        self.loading_pallet_frame.grid(row=0, column=3, sticky="nsew", padx=(0,10), pady=(0,10))
        self.loading_pallet_frame.grid_rowconfigure(1, weight=1)
        self.loading_pallet_frame.grid_columnconfigure(0, weight=1)
        self.title_loading_pallet = customtkinter.CTkLabel(self.loading_pallet_frame, text="Loading Pallet", font=customtkinter.CTkFont(size=17),anchor="w")
        self.title_loading_pallet.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10,0))
        self.tableFrame_loading_pallet = customtkinter.CTkFrame(self.loading_pallet_frame, corner_radius=10, fg_color="white")
        self.tableFrame_loading_pallet.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.loading_pallet_tableMenunggu = ttk.Treeview(self.tableFrame_loading_pallet)
        self.loading_pallet_tableMenunggu.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10, padx=(10,0))
        self.loading_pallet_tableMenunggu["columns"] = ("Menunggu")
        self.loading_pallet_tableMenunggu.column("#0", width=0,stretch=tk.NO)
        self.loading_pallet_tableMenunggu.column("Menunggu", anchor=tk.CENTER, width=100,minwidth=0,stretch=tk.YES)
        self.loading_pallet_tableMenunggu.heading("#0", text="", anchor=tk.W)
        self.loading_pallet_tableMenunggu.heading("Menunggu", text="Menunggu", anchor=tk.CENTER)
        self.loading_pallet_tableMenunggu.tag_configure("oddrow",background="white",font=(None, 13))
        self.loading_pallet_tableMenunggu.tag_configure("evenrow",background="lightblue",font=(None, 13))
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.loading_pallet_tableDipanggil = ttk.Treeview(self.tableFrame_loading_pallet)
        self.loading_pallet_tableDipanggil.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10)
        self.loading_pallet_tableDipanggil["columns"] = ("Dipanggil")
        self.loading_pallet_tableDipanggil.column("#0", width=0,stretch=tk.NO)
        self.loading_pallet_tableDipanggil.column("Dipanggil", anchor=tk.CENTER, width=100, minwidth=0,stretch=tk.YES)
        self.loading_pallet_tableDipanggil.heading("#0", text="", anchor=tk.W)
        self.loading_pallet_tableDipanggil.heading("Dipanggil", text="Dalam Proses", anchor=tk.CENTER)
        self.loading_pallet_tableDipanggil.tag_configure("oddrow",background="white",font=(None, 13))
        self.loading_pallet_tableDipanggil.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.loading_pallet_table_scrollbar = customtkinter.CTkScrollbar(self.tableFrame_loading_pallet, hover=True, button_hover_color="dark blue", command=self.loading_pallet_tableDipanggil.yview)
        self.loading_pallet_table_scrollbar.pack(side="left",fill=tk.Y, pady=10, padx=(0,5))
        self.loading_pallet_tableMenunggu.configure(yscrollcommand=self.loading_pallet_table_scrollbar.set)
        self.loading_pallet_tableDipanggil.configure(yscrollcommand=self.loading_pallet_table_scrollbar.set)

        self.unloading_frame = customtkinter.CTkFrame(self.down_frame, corner_radius=10, fg_color="light green")
        self.unloading_frame.grid(row=0, column=4, sticky="nsew", padx=(0,10), pady=(0,10))
        self.unloading_frame.grid_rowconfigure(1, weight=1)
        self.unloading_frame.grid_columnconfigure(0, weight=1)
        self.title_unloading = customtkinter.CTkLabel(self.unloading_frame, text="Unloading", font=customtkinter.CTkFont(size=17),anchor="w")
        self.title_unloading.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10,0))
        self.tableFrame_unloading = customtkinter.CTkFrame(self.unloading_frame, corner_radius=10, fg_color="white")
        self.tableFrame_unloading.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.unloading_tableMenunggu = ttk.Treeview(self.tableFrame_unloading)
        self.unloading_tableMenunggu.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10, padx=(10,0))
        self.unloading_tableMenunggu["columns"] = ("Menunggu")
        self.unloading_tableMenunggu.column("#0", width=0,stretch=tk.NO)
        self.unloading_tableMenunggu.column("Menunggu", anchor=tk.CENTER, width=100,minwidth=0,stretch=tk.YES)
        self.unloading_tableMenunggu.heading("#0", text="", anchor=tk.W)
        self.unloading_tableMenunggu.heading("Menunggu", text="Menunggu", anchor=tk.CENTER)
        self.unloading_tableMenunggu.tag_configure("oddrow",background="white",font=(None, 13))
        self.unloading_tableMenunggu.tag_configure("evenrow",background="lightblue",font=(None, 13))
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.unloading_tableDipanggil = ttk.Treeview(self.tableFrame_unloading)
        self.unloading_tableDipanggil.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10)
        self.unloading_tableDipanggil["columns"] = ("Dipanggil")
        self.unloading_tableDipanggil.column("#0", width=0,stretch=tk.NO)
        self.unloading_tableDipanggil.column("Dipanggil", anchor=tk.CENTER, width=100, minwidth=0,stretch=tk.YES)
        self.unloading_tableDipanggil.heading("#0", text="", anchor=tk.W)
        self.unloading_tableDipanggil.heading("Dipanggil", text="Dalam Proses", anchor=tk.CENTER)
        self.unloading_tableDipanggil.tag_configure("oddrow",background="white",font=(None, 13))
        self.unloading_tableDipanggil.tag_configure("evenrow",background="lightblue",font=(None, 13))

        self.unloading_table_scrollbar = customtkinter.CTkScrollbar(self.tableFrame_unloading, hover=True, button_hover_color="dark blue", command=self.unloading_tableDipanggil.yview)
        self.unloading_table_scrollbar.pack(side="left",fill=tk.Y, pady=10, padx=(0,5))
        self.unloading_tableMenunggu.configure(yscrollcommand=self.unloading_table_scrollbar.set)
        self.unloading_tableDipanggil.configure(yscrollcommand=self.unloading_table_scrollbar.set)

        #> masukan item ke tabel penugasa warehouse (tb loading/unloading)
        # loading
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        for record in self.warehouse_assignLoading_table.get_children():
            self.warehouse_assignLoading_table.delete(record)
        count = 0
        c.execute("SELECT no_polisi, wh_id FROM tb_loadingList")
        rec = c.fetchall()
        for i in rec:
            if i[1] != "Auto":
                if count % 2 == 0:
                    self.warehouse_assignLoading_table.insert(parent="",index="end", iid=count, text="", values=(i[0], "Gudang {}".format(int(i[1]))), tags=("evenrow",))
                else:
                    self.warehouse_assignLoading_table.insert(parent="",index="end", iid=count, text="", values=(i[0], "Gudang {}".format(int(i[1]))), tags=("oddrow",))
                count+=1
            else:
                if count % 2 == 0:
                    self.warehouse_assignLoading_table.insert(parent="",index="end", iid=count, text="", values=(i[0], "-"), tags=("evenrow",))
                else:
                    self.warehouse_assignLoading_table.insert(parent="",index="end", iid=count, text="", values=(i[0], "-"), tags=("oddrow",))
                count+=1
        file.commit()
        file.close()
        # unloading
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        for record in self.warehouse_assignUnloading_table.get_children():
            self.warehouse_assignUnloading_table.delete(record)
        count = 0
        c.execute("SELECT no_polisi, wh_id FROM tb_unloading")
        rec = c.fetchall()
        for i in rec:
            if i[1] != "Auto":
                if count % 2 == 0:
                    self.warehouse_assignUnloading_table.insert(parent="",index="end", iid=count, text="", values=(i[0], "Gudang {}".format(int(i[1]))), tags=("evenrow",))
                else:
                    self.warehouse_assignUnloading_table.insert(parent="",index="end", iid=count, text="", values=(i[0], "Gudang {}".format(int(i[1]))), tags=("oddrow",))
                count+=1
            else:
                if count % 2 == 0:
                    self.warehouse_assignUnloading_table.insert(parent="",index="end", iid=count, text="", values=(i[0], "-"), tags=("evenrow",))
                else:
                    self.warehouse_assignUnloading_table.insert(parent="",index="end", iid=count, text="", values=(i[0], "-"), tags=("oddrow",))
                count+=1
        file.commit()
        file.close()

        # masukan item waiting list ke tabel2 UI
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT ticket, activity, packaging FROM tb_waitinglist")
        rec = c.fetchall()
        for i in rec:
            if i[1] != "UNLOADING":
                if i[2] == "ZAK":
                    for record in self.loading_zak_tableMenunggu.get_children():
                        self.loading_zak_tableMenunggu.delete(record)
                    self.loading_zak_tableMenunggu.insert(parent="",index="end", iid=count, text="", values=(i[0]), tags=("oddrow",))
                elif i[2] == "CURAH":
                    for record in self.loading_curah_tableMenunggu.get_children():
                        self.loading_curah_tableMenunggu.delete(record)
                    self.loading_curah_tableMenunggu.insert(parent="",index="end", iid=count, text="", values=(i[0]), tags=("oddrow",))
                elif i[2] == "JUMBO":
                    for record in self.loading_jumbo_tableMenunggu.get_children():
                        self.loading_jumbo_tableMenunggu.delete(record)
                    self.loading_jumbo_tableMenunggu.insert(parent="",index="end", iid=count, text="", values=(i[0]), tags=("oddrow",))
                elif i[2] == "PALLET":
                    for record in self.loading_pallet_tableMenunggu.get_children():
                        self.loading_pallet_tableMenunggu.delete(record)
                    self.loading_pallet_tableMenunggu.insert(parent="",index="end", iid=count, text="", values=(i[0]), tags=("oddrow",))
            else:
                for record in self.unloading_tableMenunggu.get_children():
                    self.unloading_tableMenunggu.delete(record)
                self.unloading_tableMenunggu.insert(parent="",index="end", iid=count, text="", values=(i[0]), tags=("oddrow",))
        file.commit()
        file.close()

        # thread for logic calling
        global stop_event, thread
        stop_event = threading.Event()
        thread = threading.Thread(target=self.calling_ticket, args=(stop_event,))
        thread.start()
        operator_window.wm_protocol("WM_DELETE_WINDOW", self.stop)

    #> logic calling no antrian
    def calling_ticket(self, stop_event):
        value_getin = 0
        while not stop_event.is_set():

            # # cek tiket yang sudah tidak ada di queuing status. delete dari tabelDipanggil
            # file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            # c = file.cursor()
            # c.execute("SELECT ticket, no_polisi, packaging, activity FROM tb_queuing_status")
            # rec = c.fetchall()
            # print(rec)
            # if rec:  # Check if there are records
            #     tickets_to_delete = set()  # Create a set to store tickets to delete
            #     for ticket, no_polisi, packaging, activity in rec:
            #         # Check if activity is not "UNLOADING"
            #         if activity != "UNLOADING":
            #             # Check packaging type and corresponding table
            #             table_to_check = None
            #             if packaging == "ZAK":
            #                 table_to_check = self.loading_zak_tableDipanggil
            #             elif packaging == "CURAH":
            #                 table_to_check = self.loading_curah_tableDipanggil
            #             elif packaging == "JUMBO":
            #                 table_to_check = self.loading_jumbo_tableDipanggil
            #             elif packaging == "PALLET":
            #                 table_to_check = self.loading_pallet_tableDipanggil

            #             # If table_to_check is not None, iterate over its children
            #             if table_to_check:
            #                 for record in table_to_check.get_children():
            #                     value = table_to_check.item(record, "values")
            #                     print(value[0])
            #                     print("----")
            #                     print(ticket)
            #                     if value[0] == ticket:
            #                         print("hehe")
            #                         break  # Ticket found, no need to delete
            #                 else:
            #                     print("lala")
            #                     tickets_to_delete.add(ticket)  # Ticket not found, add to set
            #         else:  # If activity is "UNLOADING"
            #             # Iterate over unloading table's children
            #             for record in self.unloading_tableDipanggil.get_children():
            #                 value = self.unloading_tableDipanggil.item(record, "values")
            #                 if value[0] == ticket:
            #                     break  # Ticket found, no need to delete
            #             else:
            #                 tickets_to_delete.add(ticket)  # Ticket not found, add to set
            #     # Delete tickets that are not in records
            #     print(tickets_to_delete)
            #     for ticket in tickets_to_delete:
            #         for record in self.unloading_tableDipanggil.get_children():
            #             value = self.unloading_tableDipanggil.item(record, "values")
            #             if value[0] == ticket:
            #                 self.unloading_tableDipanggil.delete(record)
            #                 break
            #         else:  # Ticket not found in unloading table, check other tables
            #             for table in (self.loading_zak_tableDipanggil, self.loading_curah_tableDipanggil,
            #                         self.loading_jumbo_tableDipanggil, self.loading_pallet_tableDipanggil):
            #                 for record in table.get_children():
            #                     value = table.item(record, "values")
            #                     if value[0] == ticket:
            #                         table.delete(record)
            #                         break

            #> Pairing/Registration Check
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT no_polisi FROM tb_temporaryRFID_start where location = ?",("Pairing/Registration",))
            rec = c.fetchone()
            try:
                value_getin = len(rec)
            except:
                value_getin = 0

            file.commit()
            file.close()

            # calling setting
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT delay_repeat, skip_call_time, delete_call FROM tb_call_rules")
            rec = c.fetchone()
            delay_repeat = rec[0]
            skip_call_time = rec[1]
            delete_call = rec[2]
            file.commit()
            file.close()

            # get max getin di tb tapping point
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT max FROM tb_tappingPointReg where functions = ?",("Pairing/Registration",))
            rec = c.fetchone()
            max_value_getin = rec[0]
            file.commit()
            file.close()

            # kurangi MAX getin dengan jumlah lokasi Pairing/Registration di tb_temporaryRFID_start
            max_value_getin = int(max_value_getin) - value_getin
            
            if max_value_getin > 0:
                # get ticket dan no plat berdasarkan oid terkecil di tb waitinglist
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                c.execute("SELECT ticket, no_polisi, packaging, activity FROM tb_waitinglist where oid = (SELECT MIN(oid) FROM tb_waitinglist)")
                rec = c.fetchone()
                try:
                    ticket = rec[0]
                    no_polisi = rec[1]
                    packaging = rec[2]
                    activity = rec[3]
                except:
                    ticket = "XXX"
                    no_polisi = "YYYY"
                    packaging = "Packaging"
                    activity = "Activity"
                file.commit()
                file.close()

                if rec != None:
                    # get warehouse yang ditugaskan (berdasarkan no plat dan aktivity -> tb loading/unloading)
                    if activity == "LOADING":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()
                        c.execute("SELECT wh_id FROM tb_loadingList where no_polisi = ?",(no_polisi,))
                        rec = c.fetchone()
                        wh_id = rec[0]
                        file.commit()
                        file.close()
                    else:
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()
                        c.execute("SELECT wh_id FROM tb_unloading where no_polisi = ?",(no_polisi,))
                        rec = c.fetchone()
                        wh_id = rec[0]
                        file.commit()
                        file.close()

                    # check jumlah warehouse yang ditugaskan di tb_queuing status
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT count(*) FROM tb_queuing_status where wh_id = ?",(wh_id,))
                    rec = c.fetchone()
                    count_warehouse = rec[0]
                    file.commit()
                    file.close()
                    # get max warehouse yang ditugaskan di tb tapping point
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT max FROM tb_tappingPointReg where remark = ?",("Gudang "+str(int(wh_id))))
                    rec = c.fetchone()
                    max_warehouse = rec[0]
                    file.commit()
                    file.close()
                    # kurangi max dengan jumlah  jumlah warehouse yang ditugaskan di tb_queuing status
                    max_value_warehouse = int(max_warehouse) - int(count_warehouse)

                    if max_value_warehouse > 0:
                        # tampilkan no ticket, plat dan activity di UI
                        self.no_antrian.configure(text=ticket)
                        self.no_kendaraan.configure(text=no_polisi)
                        self.packaging.configure(text="{} / {}".format(packaging,activity))
                        # masukan item waiting list ke tabel2 UI (hapus di menunggu -> panggil)
                        if activity != "UNLOADING":
                            if packaging == "ZAK":
                                for record in self.loading_zak_tableMenunggu.get_children():
                                    value = self.loading_zak_tableMenunggu.item(record, "values")
                                    if value[0] == ticket:
                                        self.loading_zak_tableMenunggu.delete(record)
                                self.loading_zak_tableDipanggil.insert(parent="",index="end", text="", values=(ticket), tags=("oddrow",))
                            elif packaging == "CURAH":
                                for record in self.loading_curah_tableMenunggu.get_children():
                                    value = self.loading_curah_tableMenunggu.item(record, "values")
                                    if value[0] == ticket:
                                        self.loading_curah_tableMenunggu.delete(record)
                                self.loading_curah_tableDipanggil.insert(parent="",index="end", text="", values=(ticket), tags=("oddrow",))
                            elif packaging == "JUMBO":
                                for record in self.loading_jumbo_tableMenunggu.get_children():
                                    value = self.loading_jumbo_tableMenunggu.item(record, "values")
                                    if value[0] == ticket:
                                        self.loading_jumbo_tableMenunggu.delete(record)
                                self.loading_jumbo_tableDipanggil.insert(parent="",index="end", text="", values=(ticket), tags=("oddrow",))
                            elif packaging == "PALLET":
                                for record in self.loading_pallet_tableMenunggu.get_children():
                                    value = self.loading_pallet_tableMenunggu.item(record, "values")
                                    if value[0] == ticket:
                                        self.loading_pallet_tableMenunggu.delete(record)
                                self.loading_pallet_tableDipanggil.insert(parent="",index="end", text="", values=(ticket), tags=("oddrow",))
                        else:
                            for record in self.unloading_tableMenunggu.get_children():
                                value = self.unloading_tableMenunggu.item(record, "values")
                                if value[0] == ticket:
                                    self.unloading_tableMenunggu.delete(record)
                            self.unloading_tableDipanggil.insert(parent="",index="end", text="", values=(ticket), tags=("oddrow",))
                        # calling no ticket | calling setiap 30 detik | 3x panggil ga getin no antrian skip antrian berikutnya | next +2x panggil ga getin delete from waiting list
                        times_calling = int(skip_call_time) + 1
                        duration_next_calling = int(delay_repeat)
                        skip = False
                        for i in range(times_calling):
                            # if plat where oid min tidak sama maka break (untuk penanganan call di main app)
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()
                            c.execute("SELECT no_polisi FROM tb_waitinglist where oid = (SELECT MIN(oid) FROM tb_waitinglist)")
                            rec = c.fetchone()
                            if rec != None:
                                no_polisi_new = rec[0]
                            else:
                                no_polisi_new = ""
                            file.commit()
                            file.close()
                            if no_polisi != no_polisi_new:
                                break
                            # check tabel waitinglist
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()
                            c.execute("SELECT * FROM tb_waitinglist where no_polisi = ?",(no_polisi,))
                            rec = c.fetchone()
                            file.commit()
                            file.close()

                            if rec != None:
                                if i != int(skip_call_time):
                                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                    c = file.cursor()
                                    c.execute("SELECT play FROM tb_inPlaySound")
                                    play_value = c.fetchone()[0]
                                    file.commit()
                                    file.close()

                                    # get time calling
                                    try:
                                        calling_time = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("INSERT INTO tb_calling_time (no_polisi,calling_time) VALUES (?,?)",
                                            no_polisi,
                                            calling_time
                                        )
                                        file.commit()
                                        file.close()
                                    except:
                                        pass

                                    # play sound
                                    if play_value == 0:
                                        sound_directory = "./sound/taping/kios/{}/{}.mp3".format(ticket[0],ticket)
                                        playsound(sound_directory)

                                    # wait until can play
                                    while play_value == 1:
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT play FROM tb_inPlaySound")
                                        play_value = c.fetchone()[0]
                                        file.commit()
                                        file.close()
                                        if play_value == 0:
                                            sound_directory = "./sound/taping/kios/{}/{}.mp3".format(ticket[0],ticket)
                                            playsound(sound_directory)

                                    time.sleep(duration_next_calling)
                            else:
                                break

                            if i == int(skip_call_time):
                                skip = True
                                input_tb_ticket_skip = True

                                # input ticket to tb_ticket_skip
                                if input_tb_ticket_skip == True:
                                    try:
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("INSERT INTO tb_ticket_skip (ticket) VALUES (?)",
                                            ticket
                                        )
                                        file.commit()
                                        file.close()
                                    except:
                                        pass
                                input_tb_ticket_skip = True

                            if i == int(skip_call_time)-1:
                                # get all no ticket di tb_ticket_skip | untuk hapus kalo ud pernah skip
                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                c = file.cursor()
                                c.execute("SELECT * FROM tb_ticket_skip")
                                rec = c.fetchall()
                                file.commit()
                                file.close()
                                if rec != None:
                                    for i in rec:
                                        if i[0] == ticket:
                                            if delete_call == True:
                                                # get no plat from tb_waitinglist
                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                c = file.cursor()
                                                c.execute("SELECT no_polisi FROM tb_waitinglist where ticket = ?",(ticket,))
                                                rec = c.fetchone()
                                                no_polisi_delete = rec[0]
                                                file.commit()
                                                file.close()
                                                # delete di tb_calling
                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                c = file.cursor()
                                                c.execute("DELETE from tb_calling where no_polisi = ?",(no_polisi_delete,))
                                                file.commit()
                                                file.close()
                                                # delete di tb_calling_time
                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                c = file.cursor()
                                                c.execute("DELETE from tb_calling_time where no_polisi = ?",(no_polisi_delete,))
                                                file.commit()
                                                file.close()
                                                # delete di tb_queuing_flow
                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                c = file.cursor()
                                                c.execute("DELETE from tb_queuing_flow where no_polisi = ?",(no_polisi_delete,))
                                                file.commit()
                                                file.close()
                                                # delete di tb_overSLA
                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                c = file.cursor()
                                                c.execute("DELETE from tb_overSLA where no_polisi = ?",(no_polisi_delete,))
                                                file.commit()
                                                file.close()
                                                # delete di waitinglist
                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                c = file.cursor()
                                                c.execute("DELETE from tb_waitinglist where ticket = ?",(ticket,))
                                                file.commit()
                                                file.close()
                                                # delete di tb tiket skip
                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                c = file.cursor()
                                                c.execute("DELETE from tb_ticket_skip where ticket = ?",(ticket,))
                                                file.commit()
                                                file.close()
                                                input_tb_ticket_skip = False
                                                # delete from tb ui dipanggil
                                                if activity != "UNLOADING":
                                                    if packaging == "ZAK":
                                                        for record in self.loading_zak_tableDipanggil.get_children():
                                                            value = self.loading_zak_tableDipanggil.item(record, "values")
                                                            if value[0] == ticket:
                                                                self.loading_zak_tableDipanggil.delete(record)
                                                    elif packaging == "CURAH":
                                                        for record in self.loading_curah_tableDipanggil.get_children():
                                                            value = self.loading_curah_tableDipanggil.item(record, "values")
                                                            if value[0] == ticket:
                                                                self.loading_curah_tableDipanggil.delete(record)
                                                    elif packaging == "JUMBO":
                                                        for record in self.loading_jumbo_tableDipanggil.get_children():
                                                            value = self.loading_jumbo_tableDipanggil.item(record, "values")
                                                            if value[0] == ticket:
                                                                self.loading_jumbo_tableDipanggil.delete(record)
                                                    elif packaging == "PALLET":
                                                        for record in self.loading_pallet_tableDipanggil.get_children():
                                                            value = self.loading_pallet_tableDipanggil.item(record, "values")
                                                            if value[0] == ticket:
                                                                self.loading_pallet_tableDipanggil.delete(record)
                                                else:
                                                    for record in self.unloading_tableDipanggil.get_children():
                                                        value = self.unloading_tableDipanggil.item(record, "values")
                                                        if value[0] == ticket:
                                                            self.unloading_tableDipanggil.delete(record)

                                                # reset no ticket, plat dan activity di UI
                                                self.no_antrian.configure(text="XXX")
                                                self.no_kendaraan.configure(text="YYYY")
                                                self.packaging.configure(text="Packaging / Activity")

                        # skip, next row waiting list
                        if skip == True:
                            skip = False
                            # kalo skip no antrian rowid nya di jadikan yang terakhir
                            try:
                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                c = file.cursor()
                                # get all data where MIN(oid)
                                c.execute("SELECT ticket, no_polisi, activity, arrival, packaging, product, skip, next_call, next_id FROM tb_waitinglist where oid = (SELECT MIN(oid) FROM tb_waitinglist)")
                                rec_first_data_waitinglist = c.fetchone()
                                # delete all data where MIN(oid)
                                c.execute("DELETE FROM tb_waitinglist where oid = (SELECT MIN(oid) FROM tb_waitinglist)")
                                file.commit()
                                file.close()
                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                c = file.cursor()
                                # insert all data to table
                                c.execute("INSERT INTO tb_waitinglist (ticket,no_polisi,activity,arrival,packaging,product,skip,next_call,next_id) VALUES (?,?,?,?,?,?,?,?,?)",
                                    rec_first_data_waitinglist[0],
                                    rec_first_data_waitinglist[1],
                                    rec_first_data_waitinglist[2],
                                    rec_first_data_waitinglist[3],
                                    rec_first_data_waitinglist[4],
                                    rec_first_data_waitinglist[5],
                                    rec_first_data_waitinglist[6],
                                    rec_first_data_waitinglist[7],
                                    rec_first_data_waitinglist[8]
                                )
                                file.commit()
                                file.close()
                            except:
                                pass
                            # kembalikan no ticket ke menunggu
                            if activity != "UNLOADING":
                                if packaging == "ZAK":
                                    for record in self.loading_zak_tableDipanggil.get_children():
                                        value = self.loading_zak_tableDipanggil.item(record, "values")
                                        if value[0] == ticket:
                                            self.loading_zak_tableMenunggu.insert(parent="",index="end", text="", values=(ticket), tags=("oddrow",))
                                            self.loading_zak_tableDipanggil.delete(record)
                                elif packaging == "CURAH":
                                    for record in self.loading_curah_tableDipanggil.get_children():
                                        value = self.loading_curah_tableDipanggil.item(record, "values")
                                        if value[0] == ticket:
                                            self.loading_curah_tableMenunggu.insert(parent="",index="end", text="", values=(ticket), tags=("oddrow",))
                                            self.loading_curah_tableDipanggil.delete(record)
                                elif packaging == "JUMBO":
                                    for record in self.loading_jumbo_tableDipanggil.get_children():
                                        value = self.loading_jumbo_tableDipanggil.item(record, "values")
                                        if value[0] == ticket:
                                            self.loading_jumbo_tableMenunggu.insert(parent="",index="end", text="", values=(ticket), tags=("oddrow",))
                                            self.loading_jumbo_tableDipanggil.delete(record)
                                elif packaging == "PALLET":
                                    for record in self.loading_pallet_tableDipanggil.get_children():
                                        value = self.loading_pallet_tableDipanggil.item(record, "values")
                                        if value[0] == ticket:
                                            self.loading_pallet_tableMenunggu.insert(parent="",index="end", text="", values=(ticket), tags=("oddrow",))
                                            self.loading_pallet_tableDipanggil.delete(record)
                            else:
                                for record in self.unloading_tableDipanggil.get_children():
                                    value = self.unloading_tableDipanggil.item(record, "values")
                                    if value[0] == ticket:
                                        self.unloading_tableMenunggu.insert(parent="",index="end", text="", values=(ticket), tags=("oddrow",))
                                        self.unloading_tableDipanggil.delete(record)

                else:
                    # tampilkan no ticket, plat dan activity di UI
                    self.no_antrian.configure(text=ticket)
                    self.no_kendaraan.configure(text=no_polisi)
                    self.packaging.configure(text="{} / {}".format(packaging,activity))

    # untuk stop thread
    def stop(self):
        stop_event.set()
        thread.join()
        operator_window.destroy()

    # untuk stop thread
    def stop_main(self):
        try:
            thread.join()
        except:
            pass
        try:
            operator_window.destroy()
        except:
            pass
        self.destroy()

    def checkbox_event(self):
        if self.checkBox_loading.get() == "Loading":
            self.checkBox_unloading.configure(state="disabled")
        if self.checkBox_loading.get() == "off Loading":
            self.checkBox_unloading.configure(state="normal")
        if self.checkBox_unloading.get() == "Unloading":
            self.checkBox_loading.configure(state="disabled")
        if self.checkBox_unloading.get() == "off Unloading":
            self.checkBox_loading.configure(state="normal")

    def set_noantrian(self):
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("select count(*) from tb_noantrian")
        rec = c.fetchone()
        count = rec[0]
        file.commit()
        file.close()
        if int(count) == 0:
            prefix = ["A","B","C","D","Z"]
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("select count(*) from tb_noantrian")
            for i in prefix:
                c.execute("INSERT INTO tb_noantrian (prefix,no_antrian) VALUES (?,?)",
                    i,
                    "0"
                )
            file.commit()
            file.close()

    def input_noPlat(self):
        if self.checkBox_loading.get() == "off Loading" and self.checkBox_unloading.get() == "off Unloading":
            messagebox.showerror("showerror", "Pilih Loading atau Unloading")
        else:
            # get plate number
            plat_number = self.entry_plat.get()
            # check plat num di vetting (if ada lanjut chek di loading/unloading else tb_history-popup)
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT no_polisi FROM tb_vetting where no_polisi = ?",(plat_number,))
            rec = c.fetchone()
            file.commit()
            file.close()
            if rec == None:
                messagebox.showerror("showerror", "Plat Nomer Tidak Terdaftar")
                # kirim laporan nomorplat not registered to tb history
                date = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
                waktu = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
                status = "WAITING"
                category = "REGISTRATION"
                message = "Truck {} tidak terdaftar di vetting".format(plat_number)
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                c.execute("INSERT INTO tb_alert (date,time,status,category,message,notes,no_polisi,arrival) VALUES (?,?,?,?,?,?,?,?)",
                    date,
                    waktu,
                    status,
                    category,
                    message,
                    "",
                    plat_number,
                    waktu
                )
                file.commit()
                file.close()
            elif plat_number == rec[0]:

                # check plat num dan tanggal di loading list
                if self.checkBox_loading.get() == "Loading":
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT remark, product, date, packaging, wh_id FROM tb_loadingList where no_polisi = ? AND status = ?",(plat_number,"Activate"))
                    rec_loading = c.fetchall()
                    file.commit()
                    file.close()
                    if rec_loading != None:
                        if len(rec_loading) == 1:
                            for i in rec_loading:
                                activity = i[0]
                                product = i[1]
                                tanggal = i[2]
                                pckging = i[3]
                                wrhouse = "Gudang "+i[4]
                                tanggal_now = time.strftime("%Y")+"-"+time.strftime("%m")+"-"+time.strftime("%d")
                                # generate no ticket (get packaging, get prefix packaging)
                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                c = file.cursor()
                                c.execute("SELECT packaging, wh_id, quota FROM tb_loadingList where no_polisi = ?",(plat_number,))
                                rec = c.fetchone()
                                packaging = rec[0]
                                wh_id = rec[1]
                                quota = rec[2]
                                c.execute("SELECT prefix FROM tb_packagingReg where packaging = ?",(packaging,))
                                prefix = c.fetchone()
                                prefix = prefix[0]
                                c.execute("SELECT no_antrian FROM tb_noantrian where prefix = ?",(prefix,))
                                no_antrian = c.fetchone()
                                no_antrian = no_antrian[0]
                                file.commit()
                                file.close()
                                # get no ticket n time (no ticket based on prefix packaging) -> cek dlu ada ga no platnya di tb waitinglist. kalo ga ada lanjut generate ticket
                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                c = file.cursor()
                                c.execute("SELECT no_polisi FROM tb_waitinglist where no_polisi = ?",(plat_number,))
                                rec = c.fetchone()
                                file.commit()
                                file.close()
                                if rec == None:
                                    if int(quota) > 0:
                                        # no antrian
                                        antrian = str(int(no_antrian) + 1)
                                        if antrian != "1000":
                                            # update no_antrian = antrian where prefix di tb_noantrian
                                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                            c = file.cursor()
                                            c.execute("UPDATE tb_noantrian SET no_antrian=? WHERE prefix = ?",
                                                antrian,
                                                prefix
                                            )
                                            file.commit()
                                            file.close()
                                        else:
                                            # update no_antrian = antrian where prefix di tb_noantrian
                                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                            c = file.cursor()
                                            antrian = "1"
                                            c.execute("UPDATE tb_noantrian SET no_antrian=? WHERE prefix = ?",
                                                antrian,
                                                prefix
                                            )
                                            file.commit()
                                            file.close()
                                        # input to waiting list
                                        ticket = prefix+antrian
                                        arrival = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("INSERT INTO tb_waitinglist (ticket,no_polisi,activity,arrival,packaging,product,skip,next_call,next_id) VALUES (?,?,?,?,?,?,?,?,?)",
                                            ticket,
                                            plat_number,
                                            activity,
                                            arrival,
                                            packaging,
                                            product,
                                            "",
                                            "",
                                            ""
                                        )
                                        file.commit()
                                        file.close()
                                        # input to ticket history
                                        ticket_date = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT company FROM tb_vetting where no_polisi = ?",(plat_number,))
                                        rec = c.fetchone()
                                        company = rec[0]
                                        file.commit()
                                        file.close()
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("INSERT INTO tb_ticket_history (ticket,no_polisi,activity,date,arrival,packaging,product,company,status) VALUES (?,?,?,?,?,?,?,?,?)",
                                            ticket,
                                            plat_number,
                                            activity,
                                            ticket_date,
                                            arrival,
                                            packaging,
                                            product,
                                            company,
                                            "Print"
                                        )
                                        file.commit()
                                        file.close()
                                        # kurangi kuota di tb loading
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT quota FROM tb_loadingList where no_polisi = ?",(plat_number,))
                                        rec = c.fetchone()
                                        kuota = rec[0]
                                        file.commit()
                                        file.close()
                                        kuota = int(kuota)-1
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("UPDATE tb_loadingList SET quota=? WHERE no_polisi = ? AND status = ?",
                                            kuota,
                                            plat_number,
                                            "Activate"
                                        )
                                        file.commit()
                                        file.close()
                                        #> tentukan warehouse jika wh id Auto, input to loading list
                                        # check wh_id
                                        if wh_id == "Auto":
                                            #> Hitung jumlah tiap wh_id
                                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                            c = file.cursor()
                                            c.execute("SELECT wh_id, COUNT(*) as wh_id_count FROM tb_queuing_status GROUP BY wh_id")
                                            rec_gudang = c.fetchall()
                                            file.commit()
                                            file.close()
                                            # get gudang2 yang statusnya activate dan ada packagingnya
                                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                            c = file.cursor()
                                            c.execute("SELECT remark, packaging, max, status FROM tb_tappingPointReg where functions = ?",("Warehouse",))
                                            rec = c.fetchall()
                                            file.commit()
                                            file.close()
                                            for i in reversed(rec):
                                                gudang = i[0]
                                                kapasitas_gudang = int(i[2])
                                                if len(rec_gudang) != 0:
                                                    for row in rec_gudang:
                                                        wh = row[0]
                                                        value_wh = row[1]
                                                        if wh.startswith('0'):
                                                            wh = "Gudang " + wh[1:]
                                                        else:
                                                            wh = "Gudang " + wh
                                                        if wh == gudang:
                                                            available_wh = kapasitas_gudang - value_wh
                                                        else:
                                                            available_wh = kapasitas_gudang
                                                        # cek status warehouse
                                                        if i[3] == "Activate":
                                                            packaging_parsing = str(i[1])
                                                            packaging_parsing = packaging_parsing.split("/")
                                                            for j in packaging_parsing:
                                                                # cek packaging di warehouse itu
                                                                if j == packaging:
                                                                    if available_wh > 0:
                                                                        new_wh_id = "0"+gudang[7]
                                                                        # update wh_id dengan new_wh_id di tabel loadinglist
                                                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                                        c = file.cursor()
                                                                        c.execute("UPDATE tb_loadingList SET wh_id=? WHERE no_polisi = ? AND date = ?",
                                                                            new_wh_id,
                                                                            plat_number,
                                                                            tanggal
                                                                        )
                                                                        file.commit()
                                                                        file.close()

                                                                        break
                                                else:
                                                    for i in reversed(rec):
                                                        gudang = i[0]
                                                        kapasitas_gudang = int(i[2])
                                                        available_wh = kapasitas_gudang
                                                        if i[3] == "Activate":
                                                            packaging_parsing = str(i[1])
                                                            packaging_parsing = packaging_parsing.split("/")
                                                            for j in packaging_parsing:
                                                                # cek packaging di warehouse itu
                                                                if j == packaging:
                                                                    if available_wh > 0:
                                                                        new_wh_id = "0"+gudang[7]
                                                                        # update wh_id dengan new_wh_id di tabel loadinglist
                                                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                                        c = file.cursor()
                                                                        c.execute("UPDATE tb_loadingList SET wh_id=? WHERE no_polisi = ? AND date = ?",
                                                                            new_wh_id,
                                                                            plat_number,
                                                                            tanggal
                                                                        )
                                                                        file.commit()
                                                                        file.close()

                                                                        break

                                        #> print ticket
                                        available_ports = [port.device for port in serial.tools.list_ports.comports()]
                                        for port in available_ports:
                                            try:
                                                ser = serial.Serial(port, 9600, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE, bytesize=serial.EIGHTBITS, timeout=1)
                                                header = "Queuing Management System\n"
                                                ser.write(bytes("\x1B\x21\x04", 'utf-8'))  # Set font size to double-width and double-height
                                                ser.write(bytes(header, 'utf-8'))
                                                ser.write('PT AKR Corporindo Tbk\n'.encode('utf-8'))
                                                ser.write('------------------------------\n'.encode('utf-8'))
                                                ser.write(f'No Antrian\t\t{ticket}\n'.encode('utf-8'))
                                                ser.write('------------------------------\n'.encode('utf-8'))
                                                ser.write(f'Penugasan\t\t{wrhouse}\n'.encode('utf-8'))
                                                ser.write(f'Aktivitas\t\t{activity}\n'.encode('utf-8'))
                                                ser.write(f'Packaging\t\t{pckging}\n'.encode('utf-8'))
                                                ser.write('------------------------------\n'.encode('utf-8'))
                                                ser.write('\n'.encode('utf-8'))
                                                ser.write('\n'.encode('utf-8'))
                                                ser.write('\n'.encode('utf-8'))
                                                ser.write('\n'.encode('utf-8'))
                                                ser.write('\n'.encode('utf-8'))
                                                ser.write('\n'.encode('utf-8'))
                                                cut_sequence = b'\x1D\x56\x42\x00'
                                                ser.write(cut_sequence)
                                                ser.close()
                                            except:
                                                pass
                                    else:
                                        messagebox.showerror("showerror", "Kuota untuk kendaraan ini sudah habis")
                                else:
                                    messagebox.showerror("showerror", "Plat nomer sudah ada di waiting list")
                        else:
                            messagebox.showerror("showerror", "Terdapat lebih dari 1 data dengan tugas yang sama\nSilahkan Non Activate salah satu di tabel Loading")
                            # kirim laporan nomorplat not registered to tb alert
                            date = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
                            waktu = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
                            status = "WAITING"
                            category = "REGISTRATION"
                            message = "Truck {} memiliki lebih dari 2 tugas di tabel loading, silahkan Not Activate salah satu".format(plat_number)
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()
                            c.execute("INSERT INTO tb_alert (date,time,status,category,message,notes,no_polisi,arrival) VALUES (?,?,?,?,?,?,?,?)",
                                date,
                                waktu,
                                status,
                                category,
                                message,
                                "",
                                plat_number,
                                waktu
                            )
                            file.commit()
                            file.close()
                    else:
                        messagebox.showerror("showerror", "Plat Nomer Tidak Terdaftar di Loading List\natau Plat Nomer tidak diaktivasi")
                        # kirim laporan nomorplat not registered to tb alert
                        date = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
                        waktu = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
                        status = "WAITING"
                        category = "REGISTRATION"
                        message = "Truck {} tidak terdaftar di loading / tidak diaktivasi".format(plat_number)
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()
                        c.execute("INSERT INTO tb_alert (date,time,status,category,message,notes,no_polisi,arrival) VALUES (?,?,?,?,?,?,?,?)",
                            date,
                            waktu,
                            status,
                            category,
                            message,
                            "",
                            plat_number,
                            waktu
                        )
                        file.commit()
                        file.close()
                    
                            
                # check plat num di unloading list
                if self.checkBox_unloading.get() == "Unloading":
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT remark, product, date, packaging, wh_id FROM tb_unloading where no_polisi = ? AND status = ?",(plat_number,"Activate"))
                    rec_unloading = c.fetchall()
                    file.commit()
                    file.close()
                    if rec_unloading != None:
                        if len(rec_unloading) == 1:
                            for i in rec_unloading:
                                activity = i[0]
                                product = i[1]
                                tanggal = i[2]
                                pckging = i[3]
                                wrhouse = "Gudang "+i[4]
                                tanggal_now = time.strftime("%Y")+"-"+time.strftime("%m")+"-"+time.strftime("%d")
                                # generate no ticket (get packaging, get prefix packaging)
                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                c = file.cursor()
                                c.execute("SELECT packaging, wh_id, quota FROM tb_unloading where no_polisi = ?",(plat_number,))
                                rec = c.fetchone()
                                packaging = rec[0]
                                wh_id = rec[1]
                                quota = rec[2]
                                prefix = "Z"
                                c.execute("SELECT no_antrian FROM tb_noantrian where prefix = ?",(prefix,))
                                no_antrian = c.fetchone()
                                no_antrian = no_antrian[0]
                                file.commit()
                                file.close()
                                # get no ticket n time (no ticket based on prefix packaging) -> cek dlu ada ga no platnya di tb waitinglist. kalo ada lanjut generate ticket
                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                c = file.cursor()
                                c.execute("SELECT no_polisi FROM tb_waitinglist where no_polisi = ?",(plat_number,))
                                rec = c.fetchone()
                                file.commit()
                                file.close()
                                if rec == None:
                                    if int(quota) > 0:
                                        # no antrian
                                        antrian = str(int(no_antrian) + 1)
                                        if antrian != "101":
                                            # update no_antrian = antrian where prefix di tb_noantrian
                                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                            c = file.cursor()
                                            c.execute("UPDATE tb_noantrian SET no_antrian=? WHERE prefix = ?",
                                                antrian,
                                                prefix
                                            )
                                            file.commit()
                                            file.close()
                                        else:
                                            # update no_antrian = antrian where prefix di tb_noantrian
                                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                            c = file.cursor()
                                            antrian = "1"
                                            c.execute("UPDATE tb_noantrian SET no_antrian=? WHERE prefix = ?",
                                                antrian,
                                                prefix
                                            )
                                            file.commit()
                                            file.close()
                                        # input to waiting list
                                        ticket = prefix+antrian
                                        arrival = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("INSERT INTO tb_waitinglist (ticket,no_polisi,activity,arrival,packaging,product,skip,next_call,next_id) VALUES (?,?,?,?,?,?,?,?,?)",
                                            ticket,
                                            plat_number,
                                            activity,
                                            arrival,
                                            packaging,
                                            product,
                                            "",
                                            "",
                                            ""
                                        )
                                        file.commit()
                                        file.close()
                                        # input to ticket history
                                        ticket_date = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT company FROM tb_vetting where no_polisi = ?",(plat_number,))
                                        rec = c.fetchone()
                                        company = rec[0]
                                        file.commit()
                                        file.close()
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("INSERT INTO tb_ticket_history (ticket,no_polisi,activity,date,arrival,packaging,product,company,status) VALUES (?,?,?,?,?,?,?,?,?)",
                                            ticket,
                                            plat_number,
                                            activity,
                                            ticket_date,
                                            arrival,
                                            packaging,
                                            product,
                                            company,
                                            "Print"
                                        )
                                        file.commit()
                                        file.close()
                                        # kurangi kuota di tb unloading
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT quota FROM tb_unloading where no_polisi = ?",(plat_number,))
                                        rec = c.fetchone()
                                        kuota = rec[0]
                                        file.commit()
                                        file.close()
                                        kuota = int(kuota)-1
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("UPDATE tb_unloading SET quota=? WHERE no_polisi = ? AND status = ?",
                                            kuota,
                                            plat_number,
                                            "Activate"
                                        )
                                        file.commit()
                                        file.close()
                                        #> tentukan warehouse jika wh id Auto, input to unloading list
                                        # check wh_id
                                        if wh_id == "Auto":
                                            # get gudang2 yang statusnya activate dan ada packagingnya
                                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                            c = file.cursor()
                                            c.execute("SELECT remark, packaging, max, status FROM tb_tappingPointReg where functions = ?",("Warehouse",))
                                            rec = c.fetchall()
                                            file.commit()
                                            file.close()
                                            for i in reversed(rec):
                                                if i[3] == "Activate":
                                                    packaging_parsing = str(i[1])
                                                    packaging_parsing = packaging_parsing.split("/")
                                                    for j in packaging_parsing:
                                                        if j == packaging:
                                                            if int(i[2]) > 0:
                                                                gudang = i[0]
                                                                new_wh_id = "0"+gudang[7]
                                                                # update wh_id dengan new_wh_id di tabel loadinglist
                                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                                c = file.cursor()
                                                                c.execute("UPDATE tb_unloading SET wh_id=? WHERE no_polisi = ? AND date = ?",
                                                                    new_wh_id,
                                                                    plat_number,
                                                                    tanggal
                                                                )
                                                                file.commit()
                                                                file.close()
                                                                break
                                        # print ticket
                                        available_ports = [port.device for port in serial.tools.list_ports.comports()]
                                        for port in available_ports:
                                            try:
                                                ser = serial.Serial(port, 9600, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE, bytesize=serial.EIGHTBITS, timeout=1)
                                                header = "Queuing Management System\n"
                                                ser.write(bytes("\x1B\x21\x04", 'utf-8'))  # Set font size to double-width and double-height
                                                ser.write(bytes(header, 'utf-8'))
                                                ser.write('PT AKR Corporindo Tbk\n'.encode('utf-8'))
                                                ser.write('------------------------------\n'.encode('utf-8'))
                                                ser.write(f'No Antrian\t\t{ticket}\n'.encode('utf-8'))
                                                ser.write('------------------------------\n'.encode('utf-8'))
                                                ser.write(f'Penugasan\t\t{wrhouse}\n'.encode('utf-8'))
                                                ser.write(f'Aktivitas\t\t{activity}\n'.encode('utf-8'))
                                                ser.write(f'Packaging\t\t{pckging}\n'.encode('utf-8'))
                                                ser.write('------------------------------\n'.encode('utf-8'))
                                                ser.write('\n'.encode('utf-8'))
                                                ser.write('\n'.encode('utf-8'))
                                                ser.write('\n'.encode('utf-8'))
                                                ser.write('\n'.encode('utf-8'))
                                                ser.write('\n'.encode('utf-8'))
                                                ser.write('\n'.encode('utf-8'))
                                                cut_sequence = b'\x1D\x56\x42\x00'
                                                ser.write(cut_sequence)
                                                ser.close()
                                            except:
                                                pass
                                    else:
                                        messagebox.showerror("showerror", "Kuota untuk kendaraan ini sudah habis")
                                else:
                                    messagebox.showerror("showerror", "Plat nomer sudah ada di waiting list")
                        else:
                            messagebox.showerror("showerror", "Terdapat lebih dari 1 data dengan tugas yang sama\nSilahkan Non Activate salah satu di tabel Unloading")
                            # kirim laporan nomorplat not registered to tb alert
                            date = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
                            waktu = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
                            status = "WAITING"
                            category = "REGISTRATION"
                            message = "Truck {} memiliki lebih dari 2 tugas di tabel unloading, silahkan Not Activate salah satu".format(plat_number)
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()
                            c.execute("INSERT INTO tb_alert (date,time,status,category,message,notes,no_polisi,arrival) VALUES (?,?,?,?,?,?,?,?)",
                                date,
                                waktu,
                                status,
                                category,
                                message,
                                "",
                                plat_number,
                                waktu
                            )
                            file.commit()
                            file.close()
                    else:
                        messagebox.showerror("showerror", "Plat Nomer Tidak Terdaftar di Unloading List\natau Plat Nomer tidak diaktivasi")
                        # kirim laporan nomorplat not registered to tb history
                        date = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
                        waktu = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
                        status = "WAITING"
                        category = "REGISTRATION"
                        message = "Truck {} tidak terdaftar di unloading / tidak diaktivasi".format(plat_number)
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()
                        c.execute("INSERT INTO tb_alert (date,time,status,category,message,notes,no_polisi,arrival) VALUES (?,?,?,?,?,?,?,?)",
                            date,
                            waktu,
                            status,
                            category,
                            message,
                            "",
                            plat_number,
                            waktu
                        )
                        file.commit()
                        file.close()
            # clean entry                
            self.entry_plat.delete(0,tk.END)


if __name__ == "__main__":
    app = AppKios()
    app.mainloop()


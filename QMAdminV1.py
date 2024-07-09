import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from datetime import date
from tkinter import messagebox

import customtkinter
from PIL import Image
from collections import OrderedDict

import threading
import serial
import time
from datetime import datetime
from playsound import playsound

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import math
from tkinter import filedialog

# database
import pyodbc
import sqlite3 as sq

# api
import requests
import json

from babel.numbers import *

import time
from datetime import datetime, timedelta

class App(customtkinter.CTk):

    # UI
    def __init__(self):
        super().__init__()

        self.title("AKR QMAdmin")
        self.geometry("1400x750")
        customtkinter.set_appearance_mode("Light")

        # global var
        global check_all, check_zak, check_curah, check_jumbo, check_pallet, productTapPoint, count_alert, tupple_waitinglist, tupple_queuing_status, jumlah_inbox, data_product, data_avg_time, data_total_arrival_time, data_arrival_time, data_date, data_total_queuing, size_diagram, size_arrow, size_container, tupple_dashboard, dict_alert_dashboard, tupple_displan_data, tupple_loadinglist, tupple_soundPlay, loadingDate, date_range

        # username
        global username_log
        username_log = ""

        # set warehouse name
        global wh_name
        wh_name = ""

        # data displan api
        global json_string, start_schedule, save_api_entry
        json_string = ""
        save_api_entry = False
        start_schedule = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")

        # untuk checkbox tapping point
        check_all = False
        check_zak = False
        check_curah = False
        check_jumbo = False
        check_pallet = False
        productTapPoint = ["-"]

        # untuk update tabel thread (cek apakah ada data baru)
        count_alert = 0
        tupple_dashboard = []
        tupple_queuing_status = []
        tupple_waitinglist = []
        tupple_loadinglist = []
        tupple_displan_data = []
        tupple_soundPlay = []
        dict_alert_dashboard = {}

        # untuk ambil data displan
        loadingDateCoba = "2023-03-23" #(nanti dihapus)
        loadingDate = datetime.strptime(loadingDateCoba, "%Y-%m-%d").date() #(nanti dihapus)
        # loadingDate = datetime.now().date()
        date_range = [loadingDate]

        # set grid layout 1x2
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # load images with light and dark mode image
        self.logo_image = customtkinter.CTkImage(Image.open("./image/akr.png"), size=(26, 26))
        self.dashboard_image = customtkinter.CTkImage(Image.open("./image/dashboard.png"), size=(20, 20))
        self.report_image = customtkinter.CTkImage(Image.open("./image/report.png"), size=(20, 20))
        self.qsettings_image = customtkinter.CTkImage(light_image=Image.open("./image/qsettings.png"),size=(20, 20))
        self.truck_image = customtkinter.CTkImage(light_image=Image.open("./image/truck.png"),size=(20, 20))
        self.settings_image = customtkinter.CTkImage(light_image=Image.open("./image/settings.png"),size=(20, 20))
        self.user_image = customtkinter.CTkImage(light_image=Image.open("./image/user.png"),size=(20, 20))
        self.log_image = customtkinter.CTkImage(light_image=Image.open("./image/log.png"),size=(30, 30))
        self.search_image = customtkinter.CTkImage(light_image=Image.open("./image/search.png"),size=(20, 20))
        # image dashboard
        size_diagram = 200
        size_arrow = 70
        size_container = 120
        self.akr_dashboard_image = customtkinter.CTkImage(light_image=Image.open("./image/akr_dashboard.png"),size=(250, 66))
        self.gatein_image = customtkinter.CTkImage(light_image=Image.open("./image/in.png"),size=(size_diagram, size_diagram))
        self.wb1_image = customtkinter.CTkImage(light_image=Image.open("./image/bridge.png"),size=(size_diagram, size_diagram))
        self.wh_image = customtkinter.CTkImage(light_image=Image.open("./image/warehouse.png"),size=(size_diagram, size_diagram))
        self.wb2_image = customtkinter.CTkImage(light_image=Image.open("./image/bridge.png"),size=(size_diagram, size_diagram))
        self.cc_image = customtkinter.CTkImage(light_image=Image.open("./image/cover.png"),size=(size_diagram, size_diagram))
        self.gateout_image = customtkinter.CTkImage(light_image=Image.open("./image/out.png"),size=(size_diagram, size_diagram))
        self.arrow_image = customtkinter.CTkImage(light_image=Image.open("./image/arrow.png"),size=(size_arrow, size_arrow))
        self.container_image = customtkinter.CTkImage(light_image=Image.open("./image/container.png"),size=(size_container, size_container))
        self.sla_alert_image = customtkinter.CTkImage(light_image=Image.open("./image/sla_alert.png"),size=(50, 50))
        self.registration_alert_image = customtkinter.CTkImage(light_image=Image.open("./image/registration_alert.png"),size=(50, 50))
        # image icon
        self.api_image = customtkinter.CTkImage(light_image=Image.open("./image/api.png"),size=(70, 70))
        self.calling_image = customtkinter.CTkImage(light_image=Image.open("./image/calling.png"),size=(70, 70))
        self.user_regis_image = customtkinter.CTkImage(light_image=Image.open("./image/user_regis.png"),size=(70, 70))
        self.user_profil_image = customtkinter.CTkImage(light_image=Image.open("./image/user_profil.png"),size=(70, 70))
        self.login_logout_image = customtkinter.CTkImage(light_image=Image.open("./image/login_logout.png"),size=(70, 70))
        self.database_setting_image = customtkinter.CTkImage(light_image=Image.open("./image/database_setting.png"),size=(70, 70))
        self.loading_schedule_image = customtkinter.CTkImage(light_image=Image.open("./image/schedule.png"),size=(70, 70))

        # date
        tgl = int(time.strftime("%d"))
        bln = int(time.strftime("%m"))
        tahun = int(time.strftime("%Y"))

        # add some style for treeview
        style = ttk.Style()
        style.theme_use("winnative")
        # configure our treeview colors
        style.configure("Treeview",
            background = "white",
            foreground = "black",
            rowheight = 30,
            fieldbackground = "white"
        )

        # create navigation frame
        self.navigation_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color='light gray')
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(7, weight=1)

        self.navigation_frame_label = customtkinter.CTkLabel(self.navigation_frame, text="  Queuing Management", image=self.logo_image,
                                                             compound="left", font=customtkinter.CTkFont(size=15, weight="bold"))
        self.navigation_frame_label.grid(row=0, column=0, padx=20, pady=20)
        self.navigation_frame_label.bind("<Button-1>", self.show_version)

        self.dashboard_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Dashboard",
                                                   fg_color="transparent", text_color=("black"), hover_color=("light blue"),
                                                   image=self.dashboard_image, anchor="w" ,font=customtkinter.CTkFont(size=13, weight="bold"), state="disabled", command=self.dashboard_button_event)
        self.dashboard_button.grid(row=1, column=0, sticky="ew")

        jumlah_inbox = 0
        self.report_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Report\t\t\t{} Inbox!".format(jumlah_inbox),
                                                      fg_color="transparent", text_color=("black"), hover_color=("light blue"),
                                                      image=self.report_image, anchor="w" ,font=customtkinter.CTkFont(size=13, weight="bold"), state="disabled", command=self.report_button_event)
        self.report_button.grid(row=2, column=0, sticky="ew")

        self.Qsettings_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Queuing Settings",
                                                      fg_color="transparent", text_color=("black"), hover_color=("light blue"),
                                                      image=self.qsettings_image, anchor="w" ,font=customtkinter.CTkFont(size=13, weight="bold"), state="disabled", command=self.qsettings_button_event)
        self.Qsettings_button.grid(row=3, column=0, sticky="ew")

        self.truck_data_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Truck Data",
                                                      fg_color="transparent", text_color=("black"), hover_color=("light blue"),
                                                      image=self.truck_image, anchor="w" ,font=customtkinter.CTkFont(size=13, weight="bold"), state="disabled", command=self.truck_data_button_event)
        self.truck_data_button.grid(row=4, column=0, sticky="ew")

        self.settings_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="System Setting",
                                                      fg_color="transparent", text_color=("black"), hover_color=("light blue"),
                                                      image=self.settings_image, anchor="w" ,font=customtkinter.CTkFont(size=13, weight="bold"), state="disabled", command=self.settings_button_event)
        self.settings_button.grid(row=5, column=0, sticky="ew")

        self.user_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="User Setting",
                                                      fg_color="transparent", text_color=("black"), hover_color=("light blue"),
                                                      image=self.user_image, anchor="w" ,font=customtkinter.CTkFont(size=13, weight="bold"), state="disabled", command=self.user_button_event)
        self.user_button.grid(row=6, column=0, sticky="ew")

        self.location = customtkinter.CTkLabel(self.navigation_frame, text="🚩 Lokasi : ", font=customtkinter.CTkFont(size=15,weight="bold"),anchor="w")
        self.location.grid(row=8, column=0, padx=15, sticky="ews")

        self.user = customtkinter.CTkLabel(self.navigation_frame, text="👷🏼 Silahkan Login", font=customtkinter.CTkFont(size=15,weight="bold"),anchor="w")
        self.user.grid(row=9, column=0, padx=15, pady=10, sticky="ews")

        self.log_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Login / Logout",
                                                      fg_color="transparent", text_color=("black"), hover_color=("#3492eb"), bg_color="#0b5499",
                                                      image=self.log_image, anchor="w" ,font=customtkinter.CTkFont(size=13, weight="bold"), command=self.log_button_event)
        self.log_button.grid(row=10, column=0, sticky="ews")

        # Add a button to toggle the visibility of the navigation sidebar
        self.toggle_button = customtkinter.CTkButton(self, text="≡", fg_color="Dark Blue", text_color=("white"), hover_color=("blue"), width=20, height=68, corner_radius=0, command=self.toggle_navigation)
        self.toggle_button.grid(row=0, column=0, sticky="ne")

        # create dashboard frame #######################################################################
        self.dashboard_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="#1aa8b8")
        self.dashboard_frame.grid_columnconfigure(0, weight=1)
        self.dashboard_frame.grid_rowconfigure(1, weight=1)
        self.dashboard_head_frame = customtkinter.CTkFrame(self.dashboard_frame, corner_radius=10, fg_color="light blue")
        self.dashboard_head_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=5, pady=(5,0))
        self.dashboard_head_label = customtkinter.CTkLabel(self.dashboard_head_frame, text="Dashboard", font=customtkinter.CTkFont(size=17),anchor="w")
        self.dashboard_head_label.grid(row=0,column=0, sticky="ew", padx=10, pady=10)
        # dashboard image/diagram
        self.dashboard_diagram_scrollableFrame = customtkinter.CTkScrollableFrame(self.dashboard_frame, corner_radius=10, fg_color="white", orientation="horizontal")
        self.dashboard_diagram_scrollableFrame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.dashboard_diagram_frame = customtkinter.CTkFrame(self.dashboard_diagram_scrollableFrame, corner_radius=10, fg_color="white")
        self.dashboard_diagram_frame.pack(expand=tk.YES)
        self.dashboard_diagram_frame.grid_columnconfigure(0, weight=1)
        self.dashboard_diagram_frame.grid_rowconfigure(3, weight=1)
        self.akr_dashboard_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=10, fg_color="transparent")
        self.akr_dashboard_frame.grid(row=0, column=0, sticky="nsew", columnspan=11)
        self.akr_dashboard = customtkinter.CTkLabel(self.akr_dashboard_frame, image=self.akr_dashboard_image, text="", compound="top")
        self.akr_dashboard.grid(row=0, column=0)
        self.gatein_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.gatein_frame.grid(row=1, column=0)
        self.gatein_dashboard = customtkinter.CTkLabel(self.gatein_frame, image=self.gatein_image, text="Gate In", compound="top")
        self.gatein_dashboard.grid(row=0, column=0, padx=5, pady=10)
        self.arrow1_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.arrow1_frame.grid(row=1, column=1)
        self.arrow1_dashboard = customtkinter.CTkLabel(self.arrow1_frame, image=self.arrow_image, text="")
        self.arrow1_dashboard.grid(row=0, column=0, padx=5, pady=10)      
        self.wb1_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.wb1_frame.grid(row=1, column=2)
        self.wb1_dashboard = customtkinter.CTkLabel(self.wb1_frame, image=self.wb1_image, text="Weight Bridge 1", compound="top")
        self.wb1_dashboard.grid(row=0, column=0, padx=5, pady=10)
        self.arrow2_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.arrow2_frame.grid(row=1, column=3)
        self.arrow2_dashboard = customtkinter.CTkLabel(self.arrow2_frame, image=self.arrow_image, text="")
        self.arrow2_dashboard.grid(row=0, column=0, padx=5, pady=10)
        self.wh_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.wh_frame.grid(row=1, column=4)
        self.wh_dashboard = customtkinter.CTkLabel(self.wh_frame, image=self.wh_image, text="Warehouse", compound="top")
        self.wh_dashboard.grid(row=0, column=0, padx=5, pady=10)
        self.arrow3_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.arrow3_frame.grid(row=1, column=5)
        self.arrow3_dashboard = customtkinter.CTkLabel(self.arrow3_frame, image=self.arrow_image, text="")
        self.arrow3_dashboard.grid(row=0, column=0, padx=5, pady=10)
        self.wb2_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.wb2_frame.grid(row=1, column=6)
        self.wb2_dashboard = customtkinter.CTkLabel(self.wb2_frame, image=self.wb2_image, text="Weight Bridge 2", compound="top")
        self.wb2_dashboard.grid(row=0, column=0, padx=5, pady=10)
        self.arrow4_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.arrow4_frame.grid(row=1, column=7)
        self.arrow4_dashboard = customtkinter.CTkLabel(self.arrow4_frame, image=self.arrow_image, text="")
        self.arrow4_dashboard.grid(row=0, column=0, padx=5, pady=10)
        self.cc_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.cc_frame.grid(row=1, column=8)
        self.cc_dashboard = customtkinter.CTkLabel(self.cc_frame, image=self.cc_image, text="Cargo Covering", compound="top")
        self.cc_dashboard.grid(row=0, column=0, padx=5, pady=10)
        self.arrow5_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.arrow5_frame.grid(row=1, column=9)
        self.arrow5_dashboard = customtkinter.CTkLabel(self.arrow5_frame, image=self.arrow_image, text="")
        self.arrow5_dashboard.grid(row=0, column=0, padx=5, pady=10)
        self.gateout_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.gateout_frame.grid(row=1, column=10)
        self.gateout_dashboard = customtkinter.CTkLabel(self.gateout_frame, image=self.gateout_image, text="Gate Out", compound="top")
        self.gateout_dashboard.grid(row=0, column=0, padx=5, pady=10)
        self.truck1_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.truck1_frame.grid(row=2, column=0)
        self.truck1_dashboard = customtkinter.CTkLabel(self.truck1_frame, image=self.container_image, text="")
        self.truck1_dashboard.grid(row=2, column=0, padx=5, pady=10)       
        self.truck2_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.truck2_frame.grid(row=2, column=2)
        self.truck2_dashboard = customtkinter.CTkLabel(self.truck2_frame, image=self.container_image, text="")
        self.truck2_dashboard.grid(row=2, column=0, padx=5, pady=10)
        self.truck3_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.truck3_frame.grid(row=2, column=4)
        self.truck3_dashboard = customtkinter.CTkLabel(self.truck3_frame, image=self.container_image, text="")
        self.truck3_dashboard.grid(row=2, column=0, padx=5, pady=10)
        self.truck3_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.truck3_frame.grid(row=2, column=6)
        self.truck3_dashboard = customtkinter.CTkLabel(self.truck3_frame, image=self.container_image, text="")
        self.truck3_dashboard.grid(row=2, column=0, padx=5, pady=10)
        self.truck4_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.truck4_frame.grid(row=2, column=8)
        self.truck4_dashboard = customtkinter.CTkLabel(self.truck4_frame, image=self.container_image, text="")
        self.truck4_dashboard.grid(row=2, column=0, padx=5, pady=10)
        self.truck5_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=0, fg_color="transparent")
        self.truck5_frame.grid(row=2, column=10)
        self.truck5_dashboard = customtkinter.CTkLabel(self.truck5_frame, image=self.container_image, text="")
        self.truck5_dashboard.grid(row=2, column=0, padx=5, pady=10)
        # table dasboard
        self.gatein_table_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=10, fg_color="#1bcf90")
        self.gatein_table_frame.grid(row=3, column=0, padx=5)
        style.map("Treeview",background=[('selected',"#1bcf90")])
        style.configure("Treeview.Heading", font=(None,12))
        self.gatein_table = ttk.Treeview(self.gatein_table_frame, height=5)
        self.gatein_table.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10, padx=10)
        self.gatein_table["columns"] = ("No_Polisi")
        self.gatein_table.column("#0", width=0,stretch=tk.NO)
        self.gatein_table.column("No_Polisi", anchor=tk.CENTER, width=150,minwidth=0,stretch=tk.YES)
        self.gatein_table.heading("#0", text="", anchor=tk.W)
        self.gatein_table.heading("No_Polisi", text="No Polisi", anchor=tk.CENTER)
        self.gatein_table.tag_configure("oddrow",background="yellow",font=(None, 15))
        self.gatein_table.tag_configure("evenrow",background="yellow",font=(None, 15))
        self.wb1_table_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=10, fg_color="#1bcf90")
        self.wb1_table_frame.grid(row=3, column=2, padx=5)
        style.map("Treeview",background=[('selected',"#1bcf90")])
        style.configure("Treeview.Heading", font=(None,12))
        self.wb1_table = ttk.Treeview(self.wb1_table_frame, height=5)
        self.wb1_table.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10, padx=10)
        self.wb1_table["columns"] = ("No_Polisi")
        self.wb1_table.column("#0", width=0,stretch=tk.NO)
        self.wb1_table.column("No_Polisi", anchor=tk.CENTER, width=150,minwidth=0,stretch=tk.YES)
        self.wb1_table.heading("#0", text="", anchor=tk.W)
        self.wb1_table.heading("No_Polisi", text="No Polisi", anchor=tk.CENTER)
        self.wb1_table.tag_configure("oddrow",background="yellow",font=(None, 15))
        self.wb1_table.tag_configure("evenrow",background="yellow",font=(None, 15))
        self.wh_table_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=10, fg_color="#1bcf90")
        self.wh_table_frame.grid(row=3, column=4, padx=5)
        style.map("Treeview",background=[('selected',"#1bcf90")])
        style.configure("Treeview.Heading", font=(None,12))
        self.wh_table = ttk.Treeview(self.wh_table_frame, height=5)
        self.wh_table.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10, padx=10)
        self.wh_table["columns"] = ("No_Polisi","Gudang")
        self.wh_table.column("#0", width=0,stretch=tk.NO)
        self.wh_table.column("No_Polisi", anchor=tk.CENTER, width=100,minwidth=0,stretch=tk.YES)
        self.wh_table.column("Gudang", anchor=tk.CENTER, width=100,minwidth=0,stretch=tk.YES)
        self.wh_table.heading("#0", text="", anchor=tk.W)
        self.wh_table.heading("No_Polisi", text="No Polisi", anchor=tk.CENTER)
        self.wh_table.heading("Gudang", text="Gudang", anchor=tk.CENTER)
        self.wh_table.tag_configure("oddrow",background="yellow",font=(None, 15))
        self.wh_table.tag_configure("evenrow",background="yellow",font=(None, 15))
        self.wb2_table_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=10, fg_color="#1bcf90")
        self.wb2_table_frame.grid(row=3, column=6, padx=5)
        style.map("Treeview",background=[('selected',"#1bcf90")])
        style.configure("Treeview.Heading", font=(None,12))
        self.wb2_table = ttk.Treeview(self.wb2_table_frame, height=5)
        self.wb2_table.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10, padx=10)
        self.wb2_table["columns"] = ("No_Polisi")
        self.wb2_table.column("#0", width=0,stretch=tk.NO)
        self.wb2_table.column("No_Polisi", anchor=tk.CENTER, width=150,minwidth=0,stretch=tk.YES)
        self.wb2_table.heading("#0", text="", anchor=tk.W)
        self.wb2_table.heading("No_Polisi", text="No Polisi", anchor=tk.CENTER)
        self.wb2_table.tag_configure("oddrow",background="yellow",font=(None, 15))
        self.wb2_table.tag_configure("evenrow",background="yellow",font=(None, 15))
        self.cc_table_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=10, fg_color="#1bcf90")
        self.cc_table_frame.grid(row=3, column=8, padx=5)
        style.map("Treeview",background=[('selected',"#1bcf90")])
        style.configure("Treeview.Heading", font=(None,12))
        self.cc_table = ttk.Treeview(self.cc_table_frame, height=5)
        self.cc_table.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10, padx=10)
        self.cc_table["columns"] = ("No_Polisi")
        self.cc_table.column("#0", width=0,stretch=tk.NO)
        self.cc_table.column("No_Polisi", anchor=tk.CENTER, width=150,minwidth=0,stretch=tk.YES)
        self.cc_table.heading("#0", text="", anchor=tk.W)
        self.cc_table.heading("No_Polisi", text="No Polisi", anchor=tk.CENTER)
        self.cc_table.tag_configure("oddrow",background="yellow",font=(None, 15))
        self.cc_table.tag_configure("evenrow",background="yellow",font=(None, 15))
        self.gateout_table_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=10, fg_color="#1bcf90")
        self.gateout_table_frame.grid(row=3, column=10, padx=5)
        style.map("Treeview",background=[('selected',"#1bcf90")])
        style.configure("Treeview.Heading", font=(None,12))
        self.gateout_table = ttk.Treeview(self.gateout_table_frame, height=5)
        self.gateout_table.pack(side="left", expand=tk.YES, fill=tk.BOTH, pady=10, padx=10)
        self.gateout_table["columns"] = ("No_Polisi")
        self.gateout_table.column("#0", width=0,stretch=tk.NO)
        self.gateout_table.column("No_Polisi", anchor=tk.CENTER, width=150,minwidth=0,stretch=tk.YES)
        self.gateout_table.heading("#0", text="", anchor=tk.W)
        self.gateout_table.heading("No_Polisi", text="No Polisi", anchor=tk.CENTER)
        self.gateout_table.tag_configure("oddrow",background="yellow",font=(None, 15))
        self.gateout_table.tag_configure("evenrow",background="yellow",font=(None, 15))
        self.warehouse_status_frame = customtkinter.CTkFrame(self.dashboard_diagram_frame, corner_radius=10, fg_color="transparent")
        self.warehouse_status_frame.grid(row=4, column=0, sticky="nsew", columnspan=11, pady=10)
        self.warehouse_status = customtkinter.CTkLabel(self.warehouse_status_frame, text="", compound="top", fg_color="#0e825a", corner_radius=10, text_color="white")
        self.warehouse_status.pack(pady=(20,0))
        # dashboard alert
        self.dashboard_alert_frame = customtkinter.CTkFrame(self.dashboard_frame, corner_radius=0, fg_color="#0c8491")
        self.dashboard_alert_frame.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)
        self.dashboard_alert_frame.grid_columnconfigure(1, weight=1)
        self.dashboard_alert_frame.grid_rowconfigure(0, weight=1)
        # alert dasboard
        self.dashboard_alert_title_frame = customtkinter.CTkFrame(self.dashboard_alert_frame, corner_radius=0, fg_color="transparent")
        self.dashboard_alert_title_frame.grid(row=0, column=0)
        self.dashboard_alert_title = customtkinter.CTkLabel(self.dashboard_alert_title_frame, text="🚨📢\nAlert/Status", text_color="white", font=customtkinter.CTkFont(size=17),anchor="center")
        self.dashboard_alert_title.pack(padx=10, pady=10)
        # alern scrollframe
        self.dashboard_alert = customtkinter.CTkScrollableFrame(self.dashboard_alert_frame, corner_radius=0, fg_color="white", orientation="horizontal", height=70, scrollbar_button_color="gray80", scrollbar_button_hover_color="#0c8491")
        self.dashboard_alert.grid(row=0, column=1, sticky="nsew")
        # dashboard button
        self.button_zoom_frame = customtkinter.CTkFrame(self.dashboard_alert_frame, corner_radius=0, fg_color="light blue")
        self.button_zoom_frame.grid(row=0, column=2)
        self.zoom_out_button = customtkinter.CTkButton(self.button_zoom_frame, fg_color="#0c8491", text="Zoom Out", hover_color="blue", text_color="white", width=80, command=self.zoom_out_dashboard)
        self.zoom_out_button.pack(side="top", padx=10,pady=10)
        self.zoom_in_button = customtkinter.CTkButton(self.button_zoom_frame, fg_color="#0c8491", text="Zoom In", hover_color="blue", text_color="white", width=80, command=self.zoom_in_dashboard)
        self.zoom_in_button.pack(side="top", padx=10, pady=(0,10))

        # create report frame #######################################################################
        self.report_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.report_frame.grid_columnconfigure(0, weight=1)
        self.report_frame.grid_rowconfigure(0, weight=1)
        # create tabview report
        self.tabview_report = customtkinter.CTkTabview(self.report_frame,fg_color=("gray92"),text_color="White",segmented_button_fg_color="Dark Blue",segmented_button_unselected_color="Dark Blue")
        self.tabview_report.grid(row=0, column=0, padx=(0, 0), pady=(0, 0), sticky="nsew")
        self.tabview_report.add("Report Viewer")
        self.tabview_report.add("Alert Note")
        #>> configure grid of individual tabs (Alert Note) and item
        self.tabview_report.tab("Alert Note").grid_rowconfigure(0, weight=1)
        self.tabview_report.tab("Alert Note").grid_columnconfigure(0, weight=1)
        self.alert_note_frame = customtkinter.CTkFrame(self.tabview_report.tab("Alert Note"), corner_radius=0, fg_color="transparent")
        self.alert_note_frame.grid(row=0, column=0, sticky="nsew")
        self.alert_note_frame.grid_columnconfigure(0, weight=1)
        self.alert_note_frame.grid_rowconfigure(0, weight=1)
        self.tabview_alert_note = customtkinter.CTkTabview(self.alert_note_frame,fg_color=("gray92"),text_color="White",segmented_button_fg_color="Dark Red",segmented_button_unselected_color="Dark Red", segmented_button_selected_color="Dark Blue")
        self.tabview_alert_note.grid(row=0, column=0, padx=(0, 0), pady=(0, 0), sticky="nsew")
        self.tabview_alert_note.add("Inbox")
        self.tabview_alert_note.add("History")
        # inbox tab
        self.tabview_alert_note.tab("Inbox").grid_rowconfigure(0, weight=1)
        self.tabview_alert_note.tab("Inbox").grid_columnconfigure(0, weight=1)
        self.alert_note_inbox_frame = customtkinter.CTkFrame(self.tabview_alert_note.tab("Inbox"), corner_radius=0, fg_color="transparent")
        self.alert_note_inbox_frame.grid(row=0, column=0, sticky="nsew")
        self.alert_note_inbox_frame.grid_columnconfigure(1, weight=1)
        self.alert_note_inbox_frame.grid_rowconfigure(0, weight=1)
        # inbox tab left
        self.alert_note_inbox_left_frame = customtkinter.CTkFrame(self.alert_note_inbox_frame, corner_radius=0, fg_color="transparent")
        self.alert_note_inbox_left_frame.grid(row=0, column=0, sticky="nsew")
        self.alert_note_inbox_left_frame.grid_columnconfigure(0, weight=1)
        self.alert_note_inbox_left_frame.grid_rowconfigure(1, weight=1)
        # title Alert and messages
        self.alert_message_title_label = customtkinter.CTkLabel(self.alert_note_inbox_left_frame, text="Alert and messages", font=customtkinter.CTkFont(size=17),anchor="w")
        self.alert_message_title_label.grid(row=0,column=0, sticky="ew", pady=(0,10))
        # tabel alert
        self.alert_table_frame = customtkinter.CTkFrame(self.alert_note_inbox_left_frame, corner_radius=0, fg_color="gray80")
        self.alert_table_frame.grid(row=1, column=0, sticky="nsew")
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.alert_table = ttk.Treeview(self.alert_table_frame)
        self.alert_table.pack(side="left", expand=tk.YES, fill=tk.BOTH)
        self.alert_table["columns"] = ("Date", "Time", "Status", "Category", "Message", "Notes", "No Polisi", "arrival")
        self.alert_table.column("#0", width=0,stretch=tk.NO)
        self.alert_table.column("Date", anchor=tk.CENTER, width=150,minwidth=0,stretch=tk.YES)
        self.alert_table.column("Time", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.alert_table.column("Status", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.alert_table.column("Category", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.alert_table.column("Message", width=0,stretch=tk.NO)
        self.alert_table.column("Notes", width=0,stretch=tk.NO)
        self.alert_table.column("No Polisi", width=0,stretch=tk.NO)
        self.alert_table.column("arrival", width=0,stretch=tk.NO)
        self.alert_table.heading("#0", text="", anchor=tk.W)
        self.alert_table.heading("Date", text="Date", anchor=tk.CENTER)
        self.alert_table.heading("Time", text="Time", anchor=tk.CENTER)
        self.alert_table.heading("Status", text="Status", anchor=tk.CENTER)
        self.alert_table.heading("Category", text="Category", anchor=tk.CENTER)
        self.alert_table.heading("Message", text="", anchor=tk.W)
        self.alert_table.heading("Notes", text="", anchor=tk.W)
        self.alert_table.heading("No Polisi", text="", anchor=tk.W)
        self.alert_table.heading("arrival", text="", anchor=tk.W)
        self.alert_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.alert_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.alert_table_scrollbar = customtkinter.CTkScrollbar(self.alert_table_frame, hover=True, button_hover_color="dark blue", command=self.alert_table.yview)
        self.alert_table_scrollbar.pack(side="left",fill=tk.Y)
        self.alert_table.configure(yscrollcommand=self.alert_table_scrollbar.set)
        self.alert_table.bind("<<TreeviewSelect>>", self.on_tree_alert_select)
        # inbox tab right
        self.alert_note_inbox_right_frame = customtkinter.CTkFrame(self.alert_note_inbox_frame, corner_radius=0, fg_color="transparent")
        self.alert_note_inbox_right_frame.grid(row=0, column=1, sticky="nsew")
        self.alert_note_inbox_right_frame.grid_columnconfigure(0, weight=1)
        self.alert_note_inbox_right_frame.grid_rowconfigure(3, weight=1)
        self.alertMessage_label = customtkinter.CTkLabel(self.alert_note_inbox_right_frame, text="Alert Message :", font=customtkinter.CTkFont(size=15,weight="bold"),anchor="w")
        self.alertMessage_label.grid(row=0,column=0, sticky="ew", padx=10, pady=(0,10))
        self.alertMessage_textbox = customtkinter.CTkTextbox(self.alert_note_inbox_right_frame, fg_color="gray20", text_color="white", height=120, font=customtkinter.CTkFont(size=15))
        self.alertMessage_textbox.grid(row=1,column=0, padx=(5,0), pady=(0,10), sticky="nsew")
        self.notes_label = customtkinter.CTkLabel(self.alert_note_inbox_right_frame, text="Notes :", font=customtkinter.CTkFont(size=15,weight="bold"),anchor="w")
        self.notes_label.grid(row=2,column=0, sticky="ew", padx=10, pady=(0,10))
        self.notes_textbox = customtkinter.CTkTextbox(self.alert_note_inbox_right_frame, fg_color="gray20", text_color="white", font=customtkinter.CTkFont(size=15))
        self.notes_textbox.grid(row=3,column=0, padx=(5,0), pady=(0,10), sticky="nsew")
        self.button_alert_note_frame = customtkinter.CTkFrame(self.alert_note_inbox_right_frame, corner_radius=0, fg_color="transparent")
        self.button_alert_note_frame.grid(row=4,column=0, padx=(5,0), pady=(0,10), sticky="nsew")
        self.button_alert_note_frame.grid_columnconfigure(1, weight=1)
        self.clear_alert_button = customtkinter.CTkButton(self.button_alert_note_frame, fg_color="dark blue", text="Clear All Remaining Alert", hover_color="blue", command=self.clear_alert)
        self.clear_alert_button.grid(row=0,column=0, sticky="ew", pady=5)
        self.button_alert_note_frame2 = customtkinter.CTkFrame(self.button_alert_note_frame, corner_radius=0, fg_color="transparent")
        self.button_alert_note_frame2.grid(row=0,column=1, padx=(5,0), sticky="nsew")
        self.alert_done_button = customtkinter.CTkButton(self.button_alert_note_frame2, fg_color="dark blue", text="DONE", hover_color="blue", command=self.done_inbox)
        self.alert_done_button.pack(side="right", padx=(10,0))
        self.alert_followup_button = customtkinter.CTkButton(self.button_alert_note_frame2, fg_color="dark blue", text="FOLLOW UP", hover_color="blue", command=self.follow_up_inbox)
        self.alert_followup_button.pack(side="right")
        # history tab
        self.tabview_alert_note.tab("History").grid_rowconfigure(0, weight=1)
        self.tabview_alert_note.tab("History").grid_columnconfigure(0, weight=1)
        self.alert_note_history_frame = customtkinter.CTkFrame(self.tabview_alert_note.tab("History"), corner_radius=0, fg_color="transparent")
        self.alert_note_history_frame.grid(row=0, column=0, sticky="nsew")
        self.alert_note_history_frame.grid_columnconfigure(0, weight=1)
        self.alert_note_history_frame.grid_rowconfigure(1, weight=1)
        # history tab (date search)
        self.alert_note_history_date_frame = customtkinter.CTkFrame(self.alert_note_history_frame, corner_radius=10, fg_color="white")
        self.alert_note_history_date_frame.grid(row=0, column=0, pady=(0,5), sticky="nsew")
        self.alert_note_history_dateStart_label = customtkinter.CTkLabel(self.alert_note_history_date_frame, text="Date Start :", font=customtkinter.CTkFont(size=15),anchor="w")
        self.alert_note_history_dateStart_label.grid(row=0,column=0, padx=10, pady=5, sticky="ew")
        self.alert_note_history_dateStart_calendar = MyDateEntry(self.alert_note_history_date_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.alert_note_history_dateStart_calendar.grid(row=0,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.alert_note_history_dateEnd_label = customtkinter.CTkLabel(self.alert_note_history_date_frame, text="Date End :", font=customtkinter.CTkFont(size=15),anchor="w")
        self.alert_note_history_dateEnd_label.grid(row=0,column=2, padx=10, pady=5, sticky="ew")
        self.alert_note_history_dateEnd_calendar = MyDateEntry(self.alert_note_history_date_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.alert_note_history_dateEnd_calendar.grid(row=0,column=3, padx=(0,10), pady= 5, sticky="ew")
        self.alert_note_history_search_button = customtkinter.CTkButton(self.alert_note_history_date_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search_alert_history)
        self.alert_note_history_search_button.grid(row=0,column=4, padx=(0,10), pady= 5, sticky="ew")
        # history tab (table)
        self.alert_note_history_tabel_frame = customtkinter.CTkFrame(self.alert_note_history_frame, corner_radius=0, fg_color="gray70")
        self.alert_note_history_tabel_frame.grid(row=1, column=0, sticky="nsew")
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.alert_note_history_tabel = ttk.Treeview(self.alert_note_history_tabel_frame)
        self.alert_note_history_tabel.pack(side="left", expand=tk.YES, fill=tk.BOTH)
        self.alert_note_history_tabel["columns"] = ("Date", "Time", "Status", "Category", "Message", "Operator", "warehouse", "notes", "no_polisi", "arrival", "solving_time", "solving_date", "solving_total_time")
        self.alert_note_history_tabel.column("#0", width=0,stretch=tk.NO)
        self.alert_note_history_tabel.column("Date", anchor=tk.CENTER, width=50,minwidth=30,stretch=tk.YES)
        self.alert_note_history_tabel.column("Time", anchor=tk.CENTER, width=50,minwidth=30,stretch=tk.YES)
        self.alert_note_history_tabel.column("Status", anchor=tk.CENTER, width=50, minwidth=0,stretch=tk.YES)
        self.alert_note_history_tabel.column("Category", anchor=tk.CENTER, width=50, minwidth=0,stretch=tk.YES)
        self.alert_note_history_tabel.column("Message", anchor=tk.CENTER, width=300, minwidth=0,stretch=tk.YES)
        self.alert_note_history_tabel.column("Operator", anchor=tk.CENTER, width=50, minwidth=0,stretch=tk.YES)
        self.alert_note_history_tabel.column("warehouse", width=0,stretch=tk.NO)
        self.alert_note_history_tabel.column("notes", width=0,stretch=tk.NO)
        self.alert_note_history_tabel.column("no_polisi", width=0,stretch=tk.NO)
        self.alert_note_history_tabel.column("arrival", width=0,stretch=tk.NO)
        self.alert_note_history_tabel.column("solving_time", width=0,stretch=tk.NO)
        self.alert_note_history_tabel.column("solving_date", width=0,stretch=tk.NO)
        self.alert_note_history_tabel.column("solving_total_time", width=0,stretch=tk.NO)
        self.alert_note_history_tabel.heading("#0", text="", anchor=tk.W)
        self.alert_note_history_tabel.heading("Date", text="Date", anchor=tk.CENTER)
        self.alert_note_history_tabel.heading("Time", text="Time", anchor=tk.CENTER)
        self.alert_note_history_tabel.heading("Status", text="Status", anchor=tk.CENTER)
        self.alert_note_history_tabel.heading("Category", text="Category", anchor=tk.CENTER)
        self.alert_note_history_tabel.heading("Message", text="Message", anchor=tk.CENTER)
        self.alert_note_history_tabel.heading("Operator", text="Operator", anchor=tk.CENTER)
        self.alert_note_history_tabel.heading("warehouse", text="", anchor=tk.W)
        self.alert_note_history_tabel.heading("notes", text="", anchor=tk.W)
        self.alert_note_history_tabel.heading("no_polisi", text="", anchor=tk.W)
        self.alert_note_history_tabel.heading("arrival", text="", anchor=tk.W)
        self.alert_note_history_tabel.heading("solving_time", text="", anchor=tk.W)
        self.alert_note_history_tabel.heading("solving_date", text="", anchor=tk.W)
        self.alert_note_history_tabel.heading("solving_total_time", text="", anchor=tk.W)
        self.alert_note_history_tabel.tag_configure("oddrow",background="white",font=(None, 13))
        self.alert_note_history_tabel.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.alert_note_history_tabel_scrollbar = customtkinter.CTkScrollbar(self.alert_note_history_tabel_frame, hover=True, button_hover_color="dark blue", command=self.alert_note_history_tabel.yview)
        self.alert_note_history_tabel_scrollbar.pack(side="left",fill=tk.Y)
        self.alert_note_history_tabel.configure(yscrollcommand=self.alert_note_history_tabel_scrollbar.set)
        self.alert_note_history_tabel.bind("<<TreeviewSelect>>", self.on_tree_packaging_select)
        # history tab (button)
        self.alert_note_history_button_frame = customtkinter.CTkFrame(self.alert_note_history_frame, corner_radius=0, fg_color="transparent")
        self.alert_note_history_button_frame.grid(row=2, column=0, sticky="nsew", pady=(5,0))
        self.alert_note_history_button = customtkinter.CTkButton(self.alert_note_history_button_frame, fg_color="dark blue", text="Export to Excel", hover_color="blue", command=self.export_excel_alert_note_history)
        self.alert_note_history_button.pack(side="right", padx=(10,0))
        #>> configure grid of individual tabs (Report Viewer) and item
        self.tabview_report.tab("Report Viewer").grid_rowconfigure(0, weight=1)
        self.tabview_report.tab("Report Viewer").grid_columnconfigure(0, weight=1)
        self.report_viewer_frame = customtkinter.CTkFrame(self.tabview_report.tab("Report Viewer"), corner_radius=0, fg_color="transparent")
        self.report_viewer_frame.grid(row=0, column=0, sticky="nsew")
        self.report_viewer_frame.grid_columnconfigure(0, weight=1)
        self.report_viewer_frame.grid_rowconfigure(0, weight=1)
        self.tabview_report_viewer = customtkinter.CTkTabview(self.report_viewer_frame,fg_color=("gray92"),text_color="White",segmented_button_fg_color="Dark Red",segmented_button_unselected_color="Dark Red", segmented_button_selected_color="Dark Blue")
        self.tabview_report_viewer.grid(row=0, column=0, padx=(0, 0), pady=(0, 0), sticky="nsew")
        self.tabview_report_viewer.add("Total Time")
        self.tabview_report_viewer.add("Transaction Time")
        self.tabview_report_viewer.add("Peak Hour")
        self.tabview_report_viewer.add("Total Queuing")
        self.tabview_report_viewer.add("Incident Report")
        self.tabview_report_viewer.add("Ticket History")
        self.tabview_report_viewer.add("SLA")
        self.tabview_report_viewer.add("AuditLog")
        # Total time tab
        self.tabview_report_viewer.tab("Total Time").grid_rowconfigure(0, weight=1)
        self.tabview_report_viewer.tab("Total Time").grid_columnconfigure(0, weight=1)
        self.report_viewer_total_time_frame = customtkinter.CTkFrame(self.tabview_report_viewer.tab("Total Time"), corner_radius=0, fg_color="transparent")
        self.report_viewer_total_time_frame.grid(row=0, column=0, sticky="nsew")
        self.report_viewer_total_time_frame.grid_columnconfigure(0, weight=1)
        self.report_viewer_total_time_frame.grid_rowconfigure(1, weight=1)
        # Total time tab (entry)
        self.report_viewer_total_time_entry_frame = customtkinter.CTkFrame(self.report_viewer_total_time_frame, corner_radius=10, fg_color="light blue")
        self.report_viewer_total_time_entry_frame.grid(row=0, column=0, pady=(0,5), sticky="nsew")
        self.report_viewer_total_time_entry_transaction_label = customtkinter.CTkLabel(self.report_viewer_total_time_entry_frame, text="Transaction :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_total_time_entry_transaction_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_total_time_optionmenu_transaction = customtkinter.CTkOptionMenu(self.report_viewer_total_time_entry_frame, values=["ALL","LOADING","UNLOADING"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_total_time_optionmenu_transaction.grid(row=0,column=1, sticky="ew", columnspan=3)
        self.report_viewer_total_time_entry_packaging_label = customtkinter.CTkLabel(self.report_viewer_total_time_entry_frame, text="Packaging :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_total_time_entry_packaging_label.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_total_time_optionmenu_packaging = customtkinter.CTkOptionMenu(self.report_viewer_total_time_entry_frame, values=[""],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", command=self.optionmenuPackaging_callback)
        self.report_viewer_total_time_optionmenu_packaging.grid(row=1,column=1, sticky="ew", columnspan=3)
        self.report_viewer_total_time_entry_product_label = customtkinter.CTkLabel(self.report_viewer_total_time_entry_frame, text="Product :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_total_time_entry_product_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_total_time_optionmenu_product = customtkinter.CTkOptionMenu(self.report_viewer_total_time_entry_frame, values=["-"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_total_time_optionmenu_product.grid(row=2,column=1, sticky="ew", columnspan=3)
        self.report_viewer_total_time_dateStart_label = customtkinter.CTkLabel(self.report_viewer_total_time_entry_frame, text="Date Start :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_total_time_dateStart_label.grid(row=3,column=0, padx=10, pady=5, sticky="ew")
        self.report_viewer_total_time_dateStart_calendar = MyDateEntry(self.report_viewer_total_time_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_total_time_dateStart_calendar.grid(row=3,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_total_time_dateEnd_label = customtkinter.CTkLabel(self.report_viewer_total_time_entry_frame, text="Date End :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_total_time_dateEnd_label.grid(row=3,column=2, padx=10, pady=5, sticky="ew")
        self.report_viewer_total_time_dateEnd_calendar = MyDateEntry(self.report_viewer_total_time_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_total_time_dateEnd_calendar.grid(row=3,column=3, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_total_time_search_button = customtkinter.CTkButton(self.report_viewer_total_time_entry_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search_report_total_time)
        self.report_viewer_total_time_search_button.grid(row=3,column=4, padx=(0,10), pady= 5, sticky="ew")
        # Total time tab (table)
        self.report_viewer_total_time_table_frame = customtkinter.CTkFrame(self.report_viewer_total_time_frame, corner_radius=0, fg_color="white")
        self.report_viewer_total_time_table_frame.grid(row=1, column=0, pady=(0,5), sticky="nsew")
        self.report_viewer_total_time_table_scrollbarY = customtkinter.CTkScrollbar(self.report_viewer_total_time_table_frame, hover=True, button_hover_color="dark blue", orientation="vertical")
        self.report_viewer_total_time_table_scrollbarY.pack(side="right",fill=tk.Y)
        self.report_viewer_total_time_table_scrollbarX = customtkinter.CTkScrollbar(self.report_viewer_total_time_table_frame, hover=True, button_hover_color="dark blue", orientation="horizontal")
        self.report_viewer_total_time_table_scrollbarX.pack(side="bottom",fill=tk.X)
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.report_viewer_total_time_table = ttk.Treeview(self.report_viewer_total_time_table_frame)
        self.report_viewer_total_time_table.pack(side="top", expand=tk.YES, fill=tk.BOTH)
        self.report_viewer_total_time_table["columns"] = ("Date", "Warehouse", "Truck Number", "Transaction", "Packaging", "Berat", "Product", "Arrival", "Waiting (Minute)", "Weight Bridge 1 (Minute)", "Warehouse (Minute)", "Weight Bidge 2 (Minute)", "Cargo Covering (Minute)", "Finish Time", "Total Time (Minute)", "Company", "Category", "No SO")
        self.report_viewer_total_time_table.column("#0", width=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Date", anchor=tk.CENTER, width=150,minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Warehouse", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Truck Number", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Transaction", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Packaging", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Berat", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Product", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Arrival", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Waiting (Minute)", anchor=tk.CENTER, width=150,minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Weight Bridge 1 (Minute)", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Warehouse (Minute)", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Weight Bidge 2 (Minute)", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Cargo Covering (Minute)", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Finish Time", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Total Time (Minute)", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Company", anchor=tk.CENTER, width=300, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("Category", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.column("No SO", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.report_viewer_total_time_table.heading("#0", text="\n\n", anchor=tk.W)
        self.report_viewer_total_time_table.heading("Date", text="Date", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Warehouse", text="Warehouse", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Truck Number", text="Truck Number", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Transaction", text="Transaction", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Packaging", text="Packaging", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Berat", text="Berat", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Product", text="Product", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Arrival", text="Arrival Time", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Waiting (Minute)", text="Waiting Time\n(Minutes)", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Weight Bridge 1 (Minute)", text="Weight Bridge\n(Minutes)", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Warehouse (Minute)", text="Warehouse\n(Minutes)", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Weight Bidge 2 (Minute)", text="Weight Bidge 2\n(Minutes)", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Cargo Covering (Minute)", text="Cargo Covering\n(Minutes)", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Finish Time", text="Finish Time", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Total Time (Minute)", text="Total Time\n(Minutes)", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Company", text="Company", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("Category", text="Category", anchor=tk.CENTER)
        self.report_viewer_total_time_table.heading("No SO", text="No SO", anchor=tk.CENTER)
        self.report_viewer_total_time_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.report_viewer_total_time_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.report_viewer_total_time_table_scrollbarY.configure(command=self.report_viewer_total_time_table.yview)
        self.report_viewer_total_time_table_scrollbarX.configure(command=self.report_viewer_total_time_table.xview)
        self.report_viewer_total_time_table.configure(yscrollcommand=self.report_viewer_total_time_table_scrollbarY.set)
        self.report_viewer_total_time_table.configure(xscrollcommand=self.report_viewer_total_time_table_scrollbarX.set)
        # Total time tab (button)
        self.report_viewer_total_time_button_frame = customtkinter.CTkFrame(self.report_viewer_total_time_frame, corner_radius=0, fg_color="transparent")
        self.report_viewer_total_time_button_frame.grid(row=2, column=0, sticky="nsew", pady=(5,0))
        self.report_viewer_total_time_button = customtkinter.CTkButton(self.report_viewer_total_time_button_frame, fg_color="dark blue", text="Export to Excel", hover_color="blue", command=self.export_excel_total_time)
        self.report_viewer_total_time_button.pack(side="right", padx=(10,0))
        # Transaction Time tab
        self.tabview_report_viewer.tab("Transaction Time").grid_rowconfigure(0, weight=1)
        self.tabview_report_viewer.tab("Transaction Time").grid_columnconfigure(0, weight=1)
        self.report_viewer_transaction_time_frame = customtkinter.CTkFrame(self.tabview_report_viewer.tab("Transaction Time"), corner_radius=0, fg_color="transparent")
        self.report_viewer_transaction_time_frame.grid(row=0, column=0, sticky="nsew")
        self.report_viewer_transaction_time_frame.grid_columnconfigure(1, weight=1)
        self.report_viewer_transaction_time_frame.grid_rowconfigure(0, weight=1)
        # Transaction Time tab (left side)
        self.report_viewer_transaction_timeLeft_frame = customtkinter.CTkFrame(self.report_viewer_transaction_time_frame, corner_radius=0, fg_color="transparent")
        self.report_viewer_transaction_timeLeft_frame.grid(row=0, column=0, sticky="nsew", padx=(0,5))
        self.report_viewer_transaction_timeLeft_frame.grid_columnconfigure(0, weight=1)
        self.report_viewer_transaction_timeLeft_frame.grid_rowconfigure(1, weight=1)
        # Transaction Time tab (left side) ENTRY
        self.report_viewer_transaction_timeLeft_entry_frame = customtkinter.CTkFrame(self.report_viewer_transaction_timeLeft_frame, corner_radius=10, fg_color="light blue")
        self.report_viewer_transaction_timeLeft_entry_frame.grid(row=0, column=0, sticky="nsew", pady=(0,5))
        self.report_viewer_transaction_time_entry_transaction_label = customtkinter.CTkLabel(self.report_viewer_transaction_timeLeft_entry_frame, text="Transaction :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_transaction_time_entry_transaction_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_transaction_time_optionmenu_transaction = customtkinter.CTkOptionMenu(self.report_viewer_transaction_timeLeft_entry_frame, values=["ALL","LOADING","UNLOADING"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_transaction_time_optionmenu_transaction.grid(row=0,column=1, sticky="ew", columnspan=3, padx=(0,10))
        self.report_viewer_transaction_time_entry_packaging_label = customtkinter.CTkLabel(self.report_viewer_transaction_timeLeft_entry_frame, text="Packaging :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_transaction_time_entry_packaging_label.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_transaction_time_optionmenu_packaging = customtkinter.CTkOptionMenu(self.report_viewer_transaction_timeLeft_entry_frame, values=[""],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", command=self.optionmenuPackaging_callback)
        self.report_viewer_transaction_time_optionmenu_packaging.grid(row=1,column=1, sticky="ew", padx=(0,10))
        self.report_viewer_transaction_time_entry_product_label = customtkinter.CTkLabel(self.report_viewer_transaction_timeLeft_entry_frame, text="Product :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_transaction_time_entry_product_label.grid(row=1,column=2, padx=10, pady= 5, sticky="ew")
        self.report_viewer_transaction_time_optionmenu_product = customtkinter.CTkOptionMenu(self.report_viewer_transaction_timeLeft_entry_frame, values=["-"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_transaction_time_optionmenu_product.grid(row=1,column=3, sticky="ew", padx=(0,10))
        self.report_viewer_transaction_time_entry_tapPoint_label = customtkinter.CTkLabel(self.report_viewer_transaction_timeLeft_entry_frame, text="Tap Point :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_transaction_time_entry_tapPoint_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_transaction_time_optionmenu_tapPoint = customtkinter.CTkOptionMenu(self.report_viewer_transaction_timeLeft_entry_frame, values=["Waiting Time", "Weight Bridge 1", "Warehouse", "Weight Bridge 2", "Cargo Covering"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_transaction_time_optionmenu_tapPoint.grid(row=2,column=1, sticky="ew", columnspan=3, padx=(0,10))
        self.report_viewer_transaction_time_dateStart_label = customtkinter.CTkLabel(self.report_viewer_transaction_timeLeft_entry_frame, text="Date Start :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_transaction_time_dateStart_label.grid(row=3,column=0, padx=10, pady=5, sticky="ew")
        self.report_viewer_transaction_time_dateStart_calendar = MyDateEntry(self.report_viewer_transaction_timeLeft_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_transaction_time_dateStart_calendar.grid(row=3,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_transaction_time_dateEnd_label = customtkinter.CTkLabel(self.report_viewer_transaction_timeLeft_entry_frame, text="Date End :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_transaction_time_dateEnd_label.grid(row=3,column=2, padx=10, pady=5, sticky="ew")
        self.report_viewer_transaction_time_dateEnd_calendar = MyDateEntry(self.report_viewer_transaction_timeLeft_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_transaction_time_dateEnd_calendar.grid(row=3,column=3, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_transaction_time_search_button = customtkinter.CTkButton(self.report_viewer_transaction_timeLeft_entry_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search_report_transaction_time)
        self.report_viewer_transaction_time_search_button.grid(row=3,column=4, padx=(0,10), pady= 5, sticky="ew")
        # Transaction Time tab (left side) TABLE
        self.report_viewer_transaction_timeLeft_table_frame = customtkinter.CTkFrame(self.report_viewer_transaction_timeLeft_frame, corner_radius=0, fg_color="white")
        self.report_viewer_transaction_timeLeft_table_frame.grid(row=1, column=0, sticky="nsew")
        self.report_viewer_transaction_time_table_scrollbarY = customtkinter.CTkScrollbar(self.report_viewer_transaction_timeLeft_table_frame, hover=True, button_hover_color="dark blue", orientation="vertical")
        self.report_viewer_transaction_time_table_scrollbarY.pack(side="right",fill=tk.Y)
        # self.report_viewer_transaction_time_table_scrollbarX = customtkinter.CTkScrollbar(self.report_viewer_transaction_timeLeft_table_frame, hover=True, button_hover_color="dark blue", orientation="horizontal")
        # self.report_viewer_transaction_time_table_scrollbarX.pack(side="bottom",fill=tk.X)
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.report_viewer_transaction_time_table = ttk.Treeview(self.report_viewer_transaction_timeLeft_table_frame)
        self.report_viewer_transaction_time_table.pack(side="top", expand=tk.YES, fill=tk.BOTH)
        self.report_viewer_transaction_time_table["columns"] = ("Warehouse", "Transaction", "Packaging", "Product", "AVG Times (Minute)", "TRN")
        self.report_viewer_transaction_time_table.column("#0", width=0,stretch=tk.NO)
        self.report_viewer_transaction_time_table.column("Warehouse", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.report_viewer_transaction_time_table.column("Transaction", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.report_viewer_transaction_time_table.column("Packaging", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.report_viewer_transaction_time_table.column("Product", anchor=tk.CENTER, width=190, minwidth=0,stretch=tk.YES)
        self.report_viewer_transaction_time_table.column("AVG Times (Minute)", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.report_viewer_transaction_time_table.column("TRN", anchor=tk.CENTER, width=100, minwidth=0,stretch=tk.YES)
        self.report_viewer_transaction_time_table.heading("#0", text="\n\n", anchor=tk.W)
        self.report_viewer_transaction_time_table.heading("Warehouse", text="Warehouse", anchor=tk.CENTER)
        self.report_viewer_transaction_time_table.heading("Transaction", text="Transaction", anchor=tk.CENTER)
        self.report_viewer_transaction_time_table.heading("Packaging", text="Packaging", anchor=tk.CENTER)
        self.report_viewer_transaction_time_table.heading("Product", text="Product", anchor=tk.CENTER)
        self.report_viewer_transaction_time_table.heading("AVG Times (Minute)", text="AVG Times\n(Minute)", anchor=tk.CENTER)
        self.report_viewer_transaction_time_table.heading("TRN", text="TRN", anchor=tk.CENTER)
        self.report_viewer_transaction_time_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.report_viewer_transaction_time_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.report_viewer_transaction_time_table_scrollbarY.configure(command=self.report_viewer_transaction_time_table.yview)
        # self.report_viewer_transaction_time_table_scrollbarX.configure(command=self.report_viewer_transaction_time_table.xview)
        self.report_viewer_transaction_time_table.configure(yscrollcommand=self.report_viewer_transaction_time_table_scrollbarY.set)
        # self.report_viewer_transaction_time_table.configure(xscrollcommand=self.report_viewer_transaction_time_table_scrollbarX.set)
        # Transaction Time tab (left side) BUTTON
        self.report_viewer_transaction_timeLeft_button_frame = customtkinter.CTkFrame(self.report_viewer_transaction_timeLeft_frame, corner_radius=0, fg_color="transparent")
        self.report_viewer_transaction_timeLeft_button_frame.grid(row=2, column=0, sticky="nsew")
        self.report_viewer_transaction_time_button = customtkinter.CTkButton(self.report_viewer_transaction_timeLeft_button_frame, fg_color="dark blue", text="Export to Excel", hover_color="blue", command=self.export_excel_transaction_time)
        self.report_viewer_transaction_time_button.pack(side="left", pady=(10,0))
        # Transaction Time tab (right side)
        self.report_viewer_transaction_timeRight_frame = customtkinter.CTkFrame(self.report_viewer_transaction_time_frame, corner_radius=0, fg_color="red")
        self.report_viewer_transaction_timeRight_frame.grid(row=0, column=1, sticky="nsew")
        self.report_viewer_transaction_timeRight_frame.grid_columnconfigure(0, weight=1)
        self.report_viewer_transaction_timeRight_frame.grid_rowconfigure(0, weight=1)
        # data & plot
        data_product = [""]
        data_avg_time = []
        fig1, self.ax1 = plt.subplots()
        self.ax1.bar(data_product, data_avg_time, color ='maroon', width = 0.4)
        self.ax1.set_xlabel("Product")
        self.ax1.set_ylabel("AVG Time")
        self.ax1.set_title("Average Transaction Time")
        self.bar_data_transaction_time = FigureCanvasTkAgg(fig1, self.report_viewer_transaction_timeRight_frame)
        self.bar_data_transaction_time.draw()
        self.bar_data_transaction_time.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.YES)
        # Peak Hour tab
        self.tabview_report_viewer.tab("Peak Hour").grid_rowconfigure(0, weight=1)
        self.tabview_report_viewer.tab("Peak Hour").grid_columnconfigure(0, weight=1)
        self.report_viewer_peak_hour_frame = customtkinter.CTkFrame(self.tabview_report_viewer.tab("Peak Hour"), corner_radius=0, fg_color="transparent")
        self.report_viewer_peak_hour_frame.grid(row=0, column=0, sticky="nsew")
        self.report_viewer_peak_hour_frame.grid_columnconfigure(1, weight=1)
        self.report_viewer_peak_hour_frame.grid_rowconfigure(0, weight=1)
        # Peak Hour tab (left side)
        self.report_viewer_peak_hourLeft_frame = customtkinter.CTkFrame(self.report_viewer_peak_hour_frame, corner_radius=0, fg_color="transparent")
        self.report_viewer_peak_hourLeft_frame.grid(row=0, column=0, sticky="nsew", padx=(0,5))
        self.report_viewer_peak_hourLeft_frame.grid_columnconfigure(0, weight=1)
        self.report_viewer_peak_hourLeft_frame.grid_rowconfigure(1, weight=1)
        # Peak Hour tab (left side) ENTRY
        self.report_viewer_peak_hour_entry_frame = customtkinter.CTkFrame(self.report_viewer_peak_hourLeft_frame, corner_radius=10, fg_color="light blue")
        self.report_viewer_peak_hour_entry_frame.grid(row=0, column=0, pady=(0,5), sticky="nsew")
        self.report_viewer_peak_hour_entry_transaction_label = customtkinter.CTkLabel(self.report_viewer_peak_hour_entry_frame, text="Transaction :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_peak_hour_entry_transaction_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_peak_hour_optionmenu_transaction = customtkinter.CTkOptionMenu(self.report_viewer_peak_hour_entry_frame, values=["ALL","LOADING","UNLOADING"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_peak_hour_optionmenu_transaction.grid(row=0,column=1, sticky="ew", columnspan=3, padx=(0,10))
        self.report_viewer_peak_hour_entry_packaging_label = customtkinter.CTkLabel(self.report_viewer_peak_hour_entry_frame, text="Packaging :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_peak_hour_entry_packaging_label.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_peak_hour_optionmenu_packaging = customtkinter.CTkOptionMenu(self.report_viewer_peak_hour_entry_frame, values=[""],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", command=self.optionmenuPackaging_callback)
        self.report_viewer_peak_hour_optionmenu_packaging.grid(row=1,column=1, sticky="ew", columnspan=3, padx=(0,10))
        self.report_viewer_peak_hour_entry_product_label = customtkinter.CTkLabel(self.report_viewer_peak_hour_entry_frame, text="Product :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_peak_hour_entry_product_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_peak_hour_optionmenu_product = customtkinter.CTkOptionMenu(self.report_viewer_peak_hour_entry_frame, values=["-"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_peak_hour_optionmenu_product.grid(row=2,column=1, sticky="ew", columnspan=3, padx=(0,10))
        self.report_viewer_peak_hour_dateStart_label = customtkinter.CTkLabel(self.report_viewer_peak_hour_entry_frame, text="Date Start :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_peak_hour_dateStart_label.grid(row=3,column=0, padx=10, pady=5, sticky="ew")
        self.report_viewer_peak_hour_dateStart_calendar = MyDateEntry(self.report_viewer_peak_hour_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_peak_hour_dateStart_calendar.grid(row=3,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_peak_hour_dateEnd_label = customtkinter.CTkLabel(self.report_viewer_peak_hour_entry_frame, text="Date End :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_peak_hour_dateEnd_label.grid(row=3,column=2, padx=10, pady=5, sticky="ew")
        self.report_viewer_peak_hour_dateEnd_calendar = MyDateEntry(self.report_viewer_peak_hour_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_peak_hour_dateEnd_calendar.grid(row=3,column=3, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_peak_hour_search_button = customtkinter.CTkButton(self.report_viewer_peak_hour_entry_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search_report_peak_hour)
        self.report_viewer_peak_hour_search_button.grid(row=3,column=4, padx=(0,10), pady= 5, sticky="ew")
        # Peak Hour tab (left side) TABLE
        self.report_viewer_peak_hourLeft_table_frame = customtkinter.CTkFrame(self.report_viewer_peak_hourLeft_frame, corner_radius=0, fg_color="white")
        self.report_viewer_peak_hourLeft_table_frame.grid(row=1, column=0, sticky="nsew")
        self.report_viewer_peak_hour_table_scrollbarY = customtkinter.CTkScrollbar(self.report_viewer_peak_hourLeft_table_frame, hover=True, button_hover_color="dark blue", orientation="vertical")
        self.report_viewer_peak_hour_table_scrollbarY.pack(side="right",fill=tk.Y)
        # self.report_viewer_peak_hour_table_scrollbarX = customtkinter.CTkScrollbar(self.report_viewer_peak_hourLeft_table_frame, hover=True, button_hover_color="dark blue", orientation="horizontal")
        # self.report_viewer_peak_hour_table_scrollbarX.pack(side="bottom",fill=tk.X)
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.report_viewer_peak_hour_table = ttk.Treeview(self.report_viewer_peak_hourLeft_table_frame)
        self.report_viewer_peak_hour_table.pack(side="top", expand=tk.YES, fill=tk.BOTH)
        self.report_viewer_peak_hour_table["columns"] = ("Warehouse", "Transaction", "Arrival Time", "Total")
        self.report_viewer_peak_hour_table.column("#0", width=0,stretch=tk.NO)
        self.report_viewer_peak_hour_table.column("Warehouse", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.report_viewer_peak_hour_table.column("Transaction", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.report_viewer_peak_hour_table.column("Arrival Time", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.report_viewer_peak_hour_table.column("Total", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.report_viewer_peak_hour_table.heading("#0", text="", anchor=tk.W)
        self.report_viewer_peak_hour_table.heading("Warehouse", text="Warehouse", anchor=tk.CENTER)
        self.report_viewer_peak_hour_table.heading("Transaction", text="Transaction", anchor=tk.CENTER)
        self.report_viewer_peak_hour_table.heading("Arrival Time", text="Arrival Time", anchor=tk.CENTER)
        self.report_viewer_peak_hour_table.heading("Total", text="Total", anchor=tk.CENTER)
        self.report_viewer_peak_hour_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.report_viewer_peak_hour_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.report_viewer_peak_hour_table_scrollbarY.configure(command=self.report_viewer_peak_hour_table.yview)
        # self.report_viewer_peak_hour_table_scrollbarX.configure(command=self.report_viewer_peak_hour_table.xview)
        self.report_viewer_peak_hour_table.configure(yscrollcommand=self.report_viewer_peak_hour_table_scrollbarY.set)
        # self.report_viewer_peak_hour_table.configure(xscrollcommand=self.report_viewer_peak_hour_table_scrollbarX.set)
        # Peak Hour tab (left side) BUTTON
        self.report_viewer_peak_hourLeft_button_frame = customtkinter.CTkFrame(self.report_viewer_peak_hourLeft_frame, corner_radius=0, fg_color="transparent")
        self.report_viewer_peak_hourLeft_button_frame.grid(row=2, column=0, sticky="nsew")
        self.report_viewer_peak_hour_button = customtkinter.CTkButton(self.report_viewer_peak_hourLeft_button_frame, fg_color="dark blue", text="Export to Excel", hover_color="blue", command=self.export_excel_peak_hour)
        self.report_viewer_peak_hour_button.pack(side="left", pady=(10,0))
        # Peak Hour tab (right side)
        self.report_viewer_peak_hourRight_frame = customtkinter.CTkFrame(self.report_viewer_peak_hour_frame, corner_radius=0, fg_color="white")
        self.report_viewer_peak_hourRight_frame.grid(row=0, column=1, sticky="nsew")
        self.report_viewer_peak_hourRight_frame.grid_columnconfigure(0, weight=1)
        self.report_viewer_peak_hourRight_frame.grid_rowconfigure(1, weight=1)
        # data dan plot Peak Hour
        data_arrival_time = []
        data_total_arrival_time = []
        fig2, self.ax2 = plt.subplots()
        self.ax2.plot(data_arrival_time, data_total_arrival_time, color ='blue', marker='o')
        self.ax2.set_xlabel("Arrival Time")
        self.ax2.set_ylabel("Total")
        self.ax2.set_title("Peak Hour")
        self.plt_data_peak_hour = FigureCanvasTkAgg(fig2, self.report_viewer_peak_hourRight_frame)
        self.plt_data_peak_hour.draw()
        self.plt_data_peak_hour.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.YES)
        # Total Queuing tab
        self.tabview_report_viewer.tab("Total Queuing").grid_rowconfigure(0, weight=1)
        self.tabview_report_viewer.tab("Total Queuing").grid_columnconfigure(0, weight=1)
        self.report_viewer_total_queuing_frame = customtkinter.CTkFrame(self.tabview_report_viewer.tab("Total Queuing"), corner_radius=0, fg_color="transparent")
        self.report_viewer_total_queuing_frame.grid(row=0, column=0, sticky="nsew")
        self.report_viewer_total_queuing_frame.grid_columnconfigure(1, weight=1)
        self.report_viewer_total_queuing_frame.grid_rowconfigure(0, weight=1)
        # Total Queuing tab (left side)
        self.report_viewer_total_queuingLeft_frame = customtkinter.CTkFrame(self.report_viewer_total_queuing_frame, corner_radius=0, fg_color="transparent")
        self.report_viewer_total_queuingLeft_frame.grid(row=0, column=0, sticky="nsew", padx=(0,5))
        self.report_viewer_total_queuingLeft_frame.grid_columnconfigure(0, weight=1)
        self.report_viewer_total_queuingLeft_frame.grid_rowconfigure(1, weight=1)
        # Total Queuing tab (left side) ENTRY
        self.report_viewer_total_queuing_entry_frame = customtkinter.CTkFrame(self.report_viewer_total_queuingLeft_frame, corner_radius=10, fg_color="light blue")
        self.report_viewer_total_queuing_entry_frame.grid(row=0, column=0, pady=(0,5), sticky="nsew")
        self.report_viewer_total_queuing_entry_transaction_label = customtkinter.CTkLabel(self.report_viewer_total_queuing_entry_frame, text="Transaction :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_total_queuing_entry_transaction_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_total_queuing_optionmenu_transaction = customtkinter.CTkOptionMenu(self.report_viewer_total_queuing_entry_frame, values=["ALL","LOADING","UNLOADING"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_total_queuing_optionmenu_transaction.grid(row=0,column=1, sticky="ew", columnspan=3, padx=(0,10))
        self.report_viewer_total_queuing_entry_packaging_label = customtkinter.CTkLabel(self.report_viewer_total_queuing_entry_frame, text="Packaging :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_total_queuing_entry_packaging_label.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_total_queuing_optionmenu_packaging = customtkinter.CTkOptionMenu(self.report_viewer_total_queuing_entry_frame, values=[""],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", command=self.optionmenuPackaging_callback)
        self.report_viewer_total_queuing_optionmenu_packaging.grid(row=1,column=1, sticky="ew", columnspan=3, padx=(0,10))
        self.report_viewer_total_queuing_entry_product_label = customtkinter.CTkLabel(self.report_viewer_total_queuing_entry_frame, text="Product :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_total_queuing_entry_product_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_total_queuing_optionmenu_product = customtkinter.CTkOptionMenu(self.report_viewer_total_queuing_entry_frame, values=["-"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_total_queuing_optionmenu_product.grid(row=2,column=1, sticky="ew", columnspan=3, padx=(0,10))
        self.report_viewer_total_queuing_dateStart_label = customtkinter.CTkLabel(self.report_viewer_total_queuing_entry_frame, text="Date Start :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_total_queuing_dateStart_label.grid(row=3,column=0, padx=10, pady=5, sticky="ew")
        self.report_viewer_total_queuing_dateStart_calendar = MyDateEntry(self.report_viewer_total_queuing_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_total_queuing_dateStart_calendar.grid(row=3,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_total_queuing_dateEnd_label = customtkinter.CTkLabel(self.report_viewer_total_queuing_entry_frame, text="Date End :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_total_queuing_dateEnd_label.grid(row=3,column=2, padx=10, pady=5, sticky="ew")
        self.report_viewer_total_queuing_dateEnd_calendar = MyDateEntry(self.report_viewer_total_queuing_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_total_queuing_dateEnd_calendar.grid(row=3,column=3, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_total_queuing_search_button = customtkinter.CTkButton(self.report_viewer_total_queuing_entry_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search_report_total_queuing)
        self.report_viewer_total_queuing_search_button.grid(row=3,column=4, padx=(0,10), pady= 5, sticky="ew")
        # Total Queuing tab (left side) TABLE
        self.report_viewer_total_queuingLeft_table_frame = customtkinter.CTkFrame(self.report_viewer_total_queuingLeft_frame, corner_radius=0, fg_color="white")
        self.report_viewer_total_queuingLeft_table_frame.grid(row=1, column=0, sticky="nsew")
        self.report_viewer_total_queuing_table_scrollbarY = customtkinter.CTkScrollbar(self.report_viewer_total_queuingLeft_table_frame, hover=True, button_hover_color="dark blue", orientation="vertical")
        self.report_viewer_total_queuing_table_scrollbarY.pack(side="right",fill=tk.Y)
        # self.report_viewer_total_queuing_table_scrollbarX = customtkinter.CTkScrollbar(self.report_viewer_total_queuingLeft_table_frame, hover=True, button_hover_color="dark blue", orientation="horizontal")
        # self.report_viewer_total_queuing_table_scrollbarX.pack(side="bottom",fill=tk.X)
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.report_viewer_total_queuing_table = ttk.Treeview(self.report_viewer_total_queuingLeft_table_frame)
        self.report_viewer_total_queuing_table.pack(side="top", expand=tk.YES, fill=tk.BOTH)
        self.report_viewer_total_queuing_table["columns"] = ("Warehouse", "Transaction", "Arrival Date", "Total")
        self.report_viewer_total_queuing_table.column("#0", width=0,stretch=tk.NO)
        self.report_viewer_total_queuing_table.column("Warehouse", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.report_viewer_total_queuing_table.column("Transaction", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.report_viewer_total_queuing_table.column("Arrival Date", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.report_viewer_total_queuing_table.column("Total", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.report_viewer_total_queuing_table.heading("#0", text="", anchor=tk.W)
        self.report_viewer_total_queuing_table.heading("Warehouse", text="Warehouse", anchor=tk.CENTER)
        self.report_viewer_total_queuing_table.heading("Transaction", text="Transaction", anchor=tk.CENTER)
        self.report_viewer_total_queuing_table.heading("Arrival Date", text="Arrival Date", anchor=tk.CENTER)
        self.report_viewer_total_queuing_table.heading("Total", text="Total", anchor=tk.CENTER)
        self.report_viewer_total_queuing_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.report_viewer_total_queuing_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.report_viewer_total_queuing_table_scrollbarY.configure(command=self.report_viewer_total_queuing_table.yview)
        # self.report_viewer_total_queuing_table_scrollbarX.configure(command=self.report_viewer_total_queuing_table.xview)
        self.report_viewer_total_queuing_table.configure(yscrollcommand=self.report_viewer_total_queuing_table_scrollbarY.set)
        # self.report_viewer_total_queuing_table.configure(xscrollcommand=self.report_viewer_total_queuing_table_scrollbarX.set)
        # Total Queuing tab (left side) BUTTON
        self.report_viewer_total_queuingLeft_button_frame = customtkinter.CTkFrame(self.report_viewer_total_queuingLeft_frame, corner_radius=0, fg_color="transparent")
        self.report_viewer_total_queuingLeft_button_frame.grid(row=2, column=0, sticky="nsew")
        self.report_viewer_total_queuing_button = customtkinter.CTkButton(self.report_viewer_total_queuingLeft_button_frame, fg_color="dark blue", text="Export to Excel", hover_color="blue", command=self.export_excel_total_queuing)
        self.report_viewer_total_queuing_button.pack(side="left", pady=(10,0))
        # Total Queuing tab (right side)
        self.report_viewer_total_queuingRight_frame = customtkinter.CTkFrame(self.report_viewer_total_queuing_frame, corner_radius=0, fg_color="white")
        self.report_viewer_total_queuingRight_frame.grid(row=0, column=1, sticky="nsew")
        self.report_viewer_total_queuingRight_frame.grid_columnconfigure(0, weight=1)
        self.report_viewer_total_queuingRight_frame.grid_rowconfigure(1, weight=1)
        # data Total Queuing
        data_date = [""]
        data_total_queuing = []
        fig3, self.ax3 = plt.subplots()
        self.ax3.bar(data_date, data_total_queuing, color ='maroon', width = 0.4)
        self.ax3.set_xlabel("Date")
        self.ax3.set_ylabel("Total")
        self.ax3.set_title("Total Queuing")
        self.bar_data_total_queuing = FigureCanvasTkAgg(fig3, self.report_viewer_total_queuingRight_frame)
        self.bar_data_total_queuing.draw()
        self.bar_data_total_queuing.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.YES)
        # Incident Report tab
        self.tabview_report_viewer.tab("Incident Report").grid_rowconfigure(0, weight=1)
        self.tabview_report_viewer.tab("Incident Report").grid_columnconfigure(0, weight=1)
        self.report_viewer_incident_report_frame = customtkinter.CTkFrame(self.tabview_report_viewer.tab("Incident Report"), corner_radius=0, fg_color="transparent")
        self.report_viewer_incident_report_frame.grid(row=0, column=0, sticky="nsew")
        self.report_viewer_incident_report_frame.grid_columnconfigure(0, weight=1)
        self.report_viewer_incident_report_frame.grid_rowconfigure(0, weight=1)
        # Incident Report tab (left side)
        self.report_viewer_incident_reportLeft_frame = customtkinter.CTkFrame(self.report_viewer_incident_report_frame, corner_radius=0, fg_color="transparent")
        self.report_viewer_incident_reportLeft_frame.grid(row=0, column=0, sticky="nsew", padx=(0,5))
        self.report_viewer_incident_reportLeft_frame.grid_columnconfigure(0, weight=1)
        self.report_viewer_incident_reportLeft_frame.grid_rowconfigure(1, weight=1)
        # Incident Report tab (left side) ENTRY
        self.report_viewer_incident_report_entry_frame = customtkinter.CTkFrame(self.report_viewer_incident_reportLeft_frame, corner_radius=10, fg_color="light blue")
        self.report_viewer_incident_report_entry_frame.grid(row=0, column=0, pady=(0,5), sticky="nsew")
        self.report_viewer_incident_report_entry_category_label = customtkinter.CTkLabel(self.report_viewer_incident_report_entry_frame, text="Category :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_incident_report_entry_category_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_incident_report_optionmenu_category = customtkinter.CTkOptionMenu(self.report_viewer_incident_report_entry_frame, values=["REGISTRATION","ALERT","OVER SLA"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_incident_report_optionmenu_category.grid(row=0,column=1, sticky="ew", columnspan=3)
        self.report_viewer_incident_report_entry_status_label = customtkinter.CTkLabel(self.report_viewer_incident_report_entry_frame, text="Status :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_incident_report_entry_status_label.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_incident_report_optionmenu_status = customtkinter.CTkOptionMenu(self.report_viewer_incident_report_entry_frame, values=["DONE","FOLLOW UP","WAITING"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_incident_report_optionmenu_status.grid(row=1,column=1, sticky="ew", columnspan=3)
        self.report_viewer_incident_report_dateStart_label = customtkinter.CTkLabel(self.report_viewer_incident_report_entry_frame, text="Date Start :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_incident_report_dateStart_label.grid(row=2,column=0, padx=10, pady=5, sticky="ew")
        self.report_viewer_incident_report_dateStart_calendar = MyDateEntry(self.report_viewer_incident_report_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_incident_report_dateStart_calendar.grid(row=2,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_incident_report_dateEnd_label = customtkinter.CTkLabel(self.report_viewer_incident_report_entry_frame, text="Date End :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_incident_report_dateEnd_label.grid(row=2,column=2, padx=10, pady=5, sticky="ew")
        self.report_viewer_incident_report_dateEnd_calendar = MyDateEntry(self.report_viewer_incident_report_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_incident_report_dateEnd_calendar.grid(row=2,column=3, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_incident_report_search_button = customtkinter.CTkButton(self.report_viewer_incident_report_entry_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search_report_incident_report)
        self.report_viewer_incident_report_search_button.grid(row=2,column=4, padx=(0,10), pady= 5, sticky="ew")
        # Incident Report tab (left side) TABLE
        self.report_viewer_incident_reportLeft_table_frame = customtkinter.CTkFrame(self.report_viewer_incident_reportLeft_frame, corner_radius=0, fg_color="WHITE")
        self.report_viewer_incident_reportLeft_table_frame.grid(row=1, column=0, pady=(0,5), sticky="nsew")
        self.report_viewer_incident_report_table_scrollbarY = customtkinter.CTkScrollbar(self.report_viewer_incident_reportLeft_table_frame, hover=True, button_hover_color="dark blue", orientation="vertical")
        self.report_viewer_incident_report_table_scrollbarY.pack(side="right",fill=tk.Y)
        self.report_viewer_incident_report_table_scrollbarX = customtkinter.CTkScrollbar(self.report_viewer_incident_reportLeft_table_frame, hover=True, button_hover_color="dark blue", orientation="horizontal")
        self.report_viewer_incident_report_table_scrollbarX.pack(side="bottom",fill=tk.X)
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.report_viewer_incident_report_table = ttk.Treeview(self.report_viewer_incident_reportLeft_table_frame)
        self.report_viewer_incident_report_table.pack(side="top", expand=tk.YES, fill=tk.BOTH)
        self.report_viewer_incident_report_table["columns"] = ("Warehouse", "Incident Date", "Incident Time", "Transaction", "Operator", "Status", "Solving Date", "Solving Time", "Total (Minute)", "Truck Number", "Arrival Time", "Remark", "Note")
        self.report_viewer_incident_report_table.column("#0", width=0,stretch=tk.NO)
        self.report_viewer_incident_report_table.column("Warehouse", anchor=tk.CENTER, width=130, minwidth=0,stretch=tk.NO)
        self.report_viewer_incident_report_table.column("Incident Date", anchor=tk.CENTER, width=130, minwidth=0,stretch=tk.NO)
        self.report_viewer_incident_report_table.column("Incident Time", anchor=tk.CENTER, width=130, minwidth=0,stretch=tk.NO)
        self.report_viewer_incident_report_table.column("Transaction", anchor=tk.CENTER, width=130, minwidth=0,stretch=tk.NO)
        self.report_viewer_incident_report_table.column("Operator", anchor=tk.CENTER, width=130, minwidth=0,stretch=tk.NO)
        self.report_viewer_incident_report_table.column("Status", anchor=tk.CENTER, width=130, minwidth=0,stretch=tk.NO)
        self.report_viewer_incident_report_table.column("Solving Date", anchor=tk.CENTER, width=130, minwidth=0,stretch=tk.NO)
        self.report_viewer_incident_report_table.column("Solving Time", anchor=tk.CENTER, width=130, minwidth=0,stretch=tk.NO)
        self.report_viewer_incident_report_table.column("Total (Minute)", anchor=tk.CENTER, width=130, minwidth=0,stretch=tk.NO)
        self.report_viewer_incident_report_table.column("Truck Number", anchor=tk.CENTER, width=130, minwidth=0,stretch=tk.NO)
        self.report_viewer_incident_report_table.column("Arrival Time", anchor=tk.CENTER, width=130, minwidth=0,stretch=tk.NO)
        self.report_viewer_incident_report_table.column("Remark", width=0,stretch=tk.NO)
        self.report_viewer_incident_report_table.column("Note", width=0,stretch=tk.NO)
        self.report_viewer_incident_report_table.heading("#0", text="\n\n", anchor=tk.W)
        self.report_viewer_incident_report_table.heading("Warehouse", text="Warehouse", anchor=tk.CENTER)
        self.report_viewer_incident_report_table.heading("Incident Date", text="Incident\nDate", anchor=tk.CENTER)
        self.report_viewer_incident_report_table.heading("Incident Time", text="Incident\nTime", anchor=tk.CENTER)
        self.report_viewer_incident_report_table.heading("Transaction", text="Transaction", anchor=tk.CENTER)
        self.report_viewer_incident_report_table.heading("Operator", text="Operator", anchor=tk.CENTER)
        self.report_viewer_incident_report_table.heading("Status", text="Status", anchor=tk.CENTER)
        self.report_viewer_incident_report_table.heading("Solving Date", text="Solving\nDate", anchor=tk.CENTER)
        self.report_viewer_incident_report_table.heading("Solving Time", text="Solving\nTime", anchor=tk.CENTER)
        self.report_viewer_incident_report_table.heading("Total (Minute)", text="Total\n(Minutes)", anchor=tk.CENTER)
        self.report_viewer_incident_report_table.heading("Truck Number", text="Truck\nNumber", anchor=tk.CENTER)
        self.report_viewer_incident_report_table.heading("Arrival Time", text="Arrival\nTime", anchor=tk.CENTER)
        self.report_viewer_incident_report_table.heading("Remark", text="", anchor=tk.W)
        self.report_viewer_incident_report_table.heading("Note", text="", anchor=tk.W)
        self.report_viewer_incident_report_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.report_viewer_incident_report_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.report_viewer_incident_report_table_scrollbarY.configure(command=self.report_viewer_incident_report_table.yview)
        self.report_viewer_incident_report_table_scrollbarX.configure(command=self.report_viewer_incident_report_table.xview)
        self.report_viewer_incident_report_table.configure(yscrollcommand=self.report_viewer_incident_report_table_scrollbarY.set)
        self.report_viewer_incident_report_table.configure(xscrollcommand=self.report_viewer_incident_report_table_scrollbarX.set)
        self.report_viewer_incident_report_table.bind("<<TreeviewSelect>>", self.on_tree_incident_report_select)
        # Incident Report tab (left side) BUTTON
        self.report_viewer_incident_report_button_frame = customtkinter.CTkFrame(self.report_viewer_incident_reportLeft_frame, corner_radius=10, fg_color="transparent")
        self.report_viewer_incident_report_button_frame.grid(row=2, column=0, sticky="nsew")
        self.report_viewer_incident_report_button = customtkinter.CTkButton(self.report_viewer_incident_report_button_frame, fg_color="dark blue", text="Export to Excel", hover_color="blue", command=self.export_excel_incident_report)
        self.report_viewer_incident_report_button.pack(side="left", pady=(10,0))
        # Incident Report tab (right side)
        self.report_viewer_incident_reportRight_frame = customtkinter.CTkFrame(self.report_viewer_incident_report_frame, corner_radius=0, fg_color="transparent")
        self.report_viewer_incident_reportRight_frame.grid(row=0, column=1, sticky="nsew")
        self.report_viewer_incident_reportRight_frame.grid_columnconfigure(0, weight=1)
        self.report_viewer_incident_reportRight_frame.grid_rowconfigure((1,3), weight=1)
        self.report_viewer_incident_reportRight_label_remark = customtkinter.CTkLabel(self.report_viewer_incident_reportRight_frame, text="Remark :", font=customtkinter.CTkFont(size=15,weight="bold"),anchor="w")
        self.report_viewer_incident_reportRight_label_remark.grid(row=0,column=0, sticky="ew", pady=(0,10))
        self.report_viewer_incident_reportRight_textbox_remark = customtkinter.CTkTextbox(self.report_viewer_incident_reportRight_frame, fg_color="gray20", text_color="white", height=120, font=customtkinter.CTkFont(size=15))
        self.report_viewer_incident_reportRight_textbox_remark.grid(row=1,column=0, padx=(5,0), pady=(0,10), sticky="nsew")
        self.report_viewer_incident_reportRight_label_note = customtkinter.CTkLabel(self.report_viewer_incident_reportRight_frame, text="Note :", font=customtkinter.CTkFont(size=15,weight="bold"),anchor="w")
        self.report_viewer_incident_reportRight_label_note.grid(row=2,column=0, sticky="ew", pady=(0,10))
        self.report_viewer_incident_reportRight_textbox_note = customtkinter.CTkTextbox(self.report_viewer_incident_reportRight_frame, fg_color="gray20", text_color="white", height=120, font=customtkinter.CTkFont(size=15))
        self.report_viewer_incident_reportRight_textbox_note.grid(row=3,column=0, padx=(5,0), pady=(0,10), sticky="nsew")
        # Ticket History tab
        self.tabview_report_viewer.tab("Ticket History").grid_rowconfigure(0, weight=1)
        self.tabview_report_viewer.tab("Ticket History").grid_columnconfigure(0, weight=1)
        self.report_viewer_ticket_history_frame = customtkinter.CTkFrame(self.tabview_report_viewer.tab("Ticket History"), corner_radius=0, fg_color="transparent")
        self.report_viewer_ticket_history_frame.grid(row=0, column=0, sticky="nsew")
        self.report_viewer_ticket_history_frame.grid_columnconfigure(0, weight=1)
        self.report_viewer_ticket_history_frame.grid_rowconfigure(1, weight=1)
        # Ticket History tab (entry)
        self.report_viewer_ticket_history_entry_frame = customtkinter.CTkFrame(self.report_viewer_ticket_history_frame, corner_radius=10, fg_color="light blue")
        self.report_viewer_ticket_history_entry_frame.grid(row=0, column=0, pady=(0,5), sticky="nsew")
        self.report_viewer_ticket_history_entry_transaction_label = customtkinter.CTkLabel(self.report_viewer_ticket_history_entry_frame, text="Transaction :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_ticket_history_entry_transaction_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_ticket_history_optionmenu_transaction = customtkinter.CTkOptionMenu(self.report_viewer_ticket_history_entry_frame, values=["ALL","LOADING","UNLOADING"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_ticket_history_optionmenu_transaction.grid(row=0,column=1, sticky="ew", columnspan=3)
        self.report_viewer_ticket_history_entry_packaging_label = customtkinter.CTkLabel(self.report_viewer_ticket_history_entry_frame, text="Packaging :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_ticket_history_entry_packaging_label.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_ticket_history_optionmenu_packaging = customtkinter.CTkOptionMenu(self.report_viewer_ticket_history_entry_frame, values=[""],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_ticket_history_optionmenu_packaging.grid(row=1,column=1, sticky="ew", columnspan=3)
        self.report_viewer_ticket_history_dateStart_label = customtkinter.CTkLabel(self.report_viewer_ticket_history_entry_frame, text="Date Start :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_ticket_history_dateStart_label.grid(row=2,column=0, padx=10, pady=5, sticky="ew")
        self.report_viewer_ticket_history_dateStart_calendar = MyDateEntry(self.report_viewer_ticket_history_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_ticket_history_dateStart_calendar.grid(row=2,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_ticket_history_dateEnd_label = customtkinter.CTkLabel(self.report_viewer_ticket_history_entry_frame, text="Date End :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_ticket_history_dateEnd_label.grid(row=2,column=2, padx=10, pady=5, sticky="ew")
        self.report_viewer_ticket_history_dateEnd_calendar = MyDateEntry(self.report_viewer_ticket_history_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_ticket_history_dateEnd_calendar.grid(row=2,column=3, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_ticket_history_search_button = customtkinter.CTkButton(self.report_viewer_ticket_history_entry_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search_report_ticket_history)
        self.report_viewer_ticket_history_search_button.grid(row=2,column=4, padx=(0,10), pady= 5, sticky="ew")
        # Ticket History tab (table)
        self.report_viewer_ticket_history_table_frame = customtkinter.CTkFrame(self.report_viewer_ticket_history_frame, corner_radius=0, fg_color="white")
        self.report_viewer_ticket_history_table_frame.grid(row=1, column=0, pady=(0,5), sticky="nsew")
        self.report_viewer_ticket_history_table_scrollbarY = customtkinter.CTkScrollbar(self.report_viewer_ticket_history_table_frame, hover=True, button_hover_color="dark blue", orientation="vertical")
        self.report_viewer_ticket_history_table_scrollbarY.pack(side="right",fill=tk.Y)
        self.report_viewer_ticket_history_table_scrollbarX = customtkinter.CTkScrollbar(self.report_viewer_ticket_history_table_frame, hover=True, button_hover_color="dark blue", orientation="horizontal")
        self.report_viewer_ticket_history_table_scrollbarX.pack(side="bottom",fill=tk.X)
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.report_viewer_ticket_history_table = ttk.Treeview(self.report_viewer_ticket_history_table_frame)
        self.report_viewer_ticket_history_table.pack(side="top", expand=tk.YES, fill=tk.BOTH)
        self.report_viewer_ticket_history_table["columns"] = ("Ticket", "Police No", "Activity", "Date", "Arrival", "Packaging", "Product", "Company", "Status")
        self.report_viewer_ticket_history_table.column("#0", width=0,stretch=tk.NO)
        self.report_viewer_ticket_history_table.column("Ticket", anchor=tk.CENTER, width=150,minwidth=0,stretch=tk.YES)
        self.report_viewer_ticket_history_table.column("Police No", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.report_viewer_ticket_history_table.column("Activity", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.report_viewer_ticket_history_table.column("Date", anchor=tk.CENTER, width=0, minwidth=0,stretch=tk.NO)
        self.report_viewer_ticket_history_table.column("Arrival", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.report_viewer_ticket_history_table.column("Packaging", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.report_viewer_ticket_history_table.column("Product", anchor=tk.CENTER, width=150,minwidth=0,stretch=tk.YES)
        self.report_viewer_ticket_history_table.column("Company", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.report_viewer_ticket_history_table.column("Status", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.report_viewer_ticket_history_table.heading("#0", text="", anchor=tk.W)
        self.report_viewer_ticket_history_table.heading("Ticket", text="Ticket", anchor=tk.CENTER)
        self.report_viewer_ticket_history_table.heading("Police No", text="Police No", anchor=tk.CENTER)
        self.report_viewer_ticket_history_table.heading("Activity", text="Activity", anchor=tk.CENTER)
        self.report_viewer_ticket_history_table.heading("Date", text="Date", anchor=tk.CENTER)
        self.report_viewer_ticket_history_table.heading("Arrival", text="Arrival", anchor=tk.CENTER)
        self.report_viewer_ticket_history_table.heading("Packaging", text="Packaging", anchor=tk.CENTER)
        self.report_viewer_ticket_history_table.heading("Product", text="Product", anchor=tk.CENTER)
        self.report_viewer_ticket_history_table.heading("Company", text="Company", anchor=tk.CENTER)
        self.report_viewer_ticket_history_table.heading("Status", text="Status", anchor=tk.CENTER)
        self.report_viewer_ticket_history_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.report_viewer_ticket_history_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.report_viewer_ticket_history_table_scrollbarY.configure(command=self.report_viewer_ticket_history_table.yview)
        self.report_viewer_ticket_history_table_scrollbarX.configure(command=self.report_viewer_ticket_history_table.xview)
        self.report_viewer_ticket_history_table.configure(yscrollcommand=self.report_viewer_ticket_history_table_scrollbarY.set)
        self.report_viewer_ticket_history_table.configure(xscrollcommand=self.report_viewer_ticket_history_table_scrollbarX.set)
        # Ticket History tab (button)
        self.report_viewer_ticket_history_button_frame = customtkinter.CTkFrame(self.report_viewer_ticket_history_frame, corner_radius=0, fg_color="transparent")
        self.report_viewer_ticket_history_button_frame.grid(row=2, column=0, sticky="nsew", pady=(5,0))
        self.report_viewer_ticket_history_button = customtkinter.CTkButton(self.report_viewer_ticket_history_button_frame, fg_color="dark blue", text="Export to Excel", hover_color="blue", command=self.export_excel_ticket_history)
        self.report_viewer_ticket_history_button.pack(side="right", padx=(10,0))
        # SLA tab
        self.tabview_report_viewer.tab("SLA").grid_rowconfigure(0, weight=1)
        self.tabview_report_viewer.tab("SLA").grid_columnconfigure(0, weight=1)
        self.report_viewer_sla_frame = customtkinter.CTkFrame(self.tabview_report_viewer.tab("SLA"), corner_radius=0, fg_color="transparent")
        self.report_viewer_sla_frame.grid(row=0, column=0, sticky="nsew")
        self.report_viewer_sla_frame.grid_columnconfigure(0, weight=1)
        self.report_viewer_sla_frame.grid_rowconfigure(1, weight=1)
        # SLA tab (entry)
        self.report_viewer_sla_entry_frame = customtkinter.CTkFrame(self.report_viewer_sla_frame, corner_radius=10, fg_color="light blue")
        self.report_viewer_sla_entry_frame.grid(row=0, column=0, pady=(0,5), sticky="nsew")
        self.report_viewer_sla_entry_transaction_label = customtkinter.CTkLabel(self.report_viewer_sla_entry_frame, text="Transaction :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_sla_entry_transaction_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_sla_optionmenu_transaction = customtkinter.CTkOptionMenu(self.report_viewer_sla_entry_frame, values=["ALL","LOADING","UNLOADING"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_sla_optionmenu_transaction.grid(row=0,column=1, sticky="ew", columnspan=3, padx=(0,10))
        self.report_viewer_sla_entry_packaging_label = customtkinter.CTkLabel(self.report_viewer_sla_entry_frame, text="Packaging :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_sla_entry_packaging_label.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_sla_optionmenu_packaging = customtkinter.CTkOptionMenu(self.report_viewer_sla_entry_frame, values=[""],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", command=self.optionmenuPackaging_callback)
        self.report_viewer_sla_optionmenu_packaging.grid(row=1,column=1, sticky="ew", columnspan=3, padx=(0,10))
        self.report_viewer_sla_entry_product_label = customtkinter.CTkLabel(self.report_viewer_sla_entry_frame, text="Product :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_sla_entry_product_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_sla_optionmenu_product = customtkinter.CTkOptionMenu(self.report_viewer_sla_entry_frame, values=["-"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_sla_optionmenu_product.grid(row=2,column=1, sticky="ew", columnspan=3, padx=(0,10))
        self.report_viewer_sla_dateStart_label = customtkinter.CTkLabel(self.report_viewer_sla_entry_frame, text="Date Start :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_sla_dateStart_label.grid(row=3,column=0, padx=10, pady=5, sticky="ew")
        self.report_viewer_sla_dateStart_calendar = MyDateEntry(self.report_viewer_sla_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_sla_dateStart_calendar.grid(row=3,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_sla_dateEnd_label = customtkinter.CTkLabel(self.report_viewer_sla_entry_frame, text="Date End :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_sla_dateEnd_label.grid(row=3,column=2, padx=10, pady=5, sticky="ew")
        self.report_viewer_sla_dateEnd_calendar = MyDateEntry(self.report_viewer_sla_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_sla_dateEnd_calendar.grid(row=3,column=3, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_sla_search_button = customtkinter.CTkButton(self.report_viewer_sla_entry_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search_report_sla)
        self.report_viewer_sla_search_button.grid(row=3,column=4, padx=(0,10), pady= 5, sticky="ew")
        # SLA tab (table)
        self.report_viewer_sla_table_frame = customtkinter.CTkFrame(self.report_viewer_sla_frame, corner_radius=0, fg_color="white")
        self.report_viewer_sla_table_frame.grid(row=1, column=0, pady=(0,5), sticky="nsew")
        self.report_viewer_sla_table_scrollbarY = customtkinter.CTkScrollbar(self.report_viewer_sla_table_frame, hover=True, button_hover_color="dark blue", orientation="vertical")
        self.report_viewer_sla_table_scrollbarY.pack(side="right",fill=tk.Y)
        self.report_viewer_sla_table_scrollbarX = customtkinter.CTkScrollbar(self.report_viewer_sla_table_frame, hover=True, button_hover_color="dark blue", orientation="horizontal")
        self.report_viewer_sla_table_scrollbarX.pack(side="bottom",fill=tk.X)
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.report_viewer_sla_table = ttk.Treeview(self.report_viewer_sla_table_frame)
        self.report_viewer_sla_table.pack(side="top", expand=tk.YES, fill=tk.BOTH)
        self.report_viewer_sla_table["columns"] = ("Product", "Packaging", "Range Tonase(Ton)", "Performance", "Waiting Time(min)", "Weight Bridge 1(min)", "Warehouse (min)", "Weight Bridge 2(min)", "Cargo Covering (min)", "Total Time (min)")
        self.report_viewer_sla_table.column("#0", width=0,stretch=tk.NO)
        self.report_viewer_sla_table.column("Product", anchor=tk.CENTER, width=300,minwidth=0,stretch=tk.YES)
        self.report_viewer_sla_table.column("Packaging", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.report_viewer_sla_table.column("Range Tonase(Ton)", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.report_viewer_sla_table.column("Performance", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.report_viewer_sla_table.column("Waiting Time(min)", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.report_viewer_sla_table.column("Weight Bridge 1(min)", anchor=tk.CENTER, width=150,minwidth=0,stretch=tk.YES)
        self.report_viewer_sla_table.column("Warehouse (min)", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.report_viewer_sla_table.column("Weight Bridge 2(min)", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.report_viewer_sla_table.column("Cargo Covering (min)", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.report_viewer_sla_table.column("Total Time (min)", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.report_viewer_sla_table.heading("#0", text="\n\n", anchor=tk.W)
        self.report_viewer_sla_table.heading("Product", text="Product", anchor=tk.CENTER)
        self.report_viewer_sla_table.heading("Packaging", text="Packaging", anchor=tk.CENTER)
        self.report_viewer_sla_table.heading("Range Tonase(Ton)", text="Range\nTonase(Ton)", anchor=tk.CENTER)
        self.report_viewer_sla_table.heading("Performance", text="Performance", anchor=tk.CENTER)
        self.report_viewer_sla_table.heading("Waiting Time(min)", text="Waiting\nTime(min)", anchor=tk.CENTER)
        self.report_viewer_sla_table.heading("Weight Bridge 1(min)", text="Weight Bridge\n1(min)", anchor=tk.CENTER)
        self.report_viewer_sla_table.heading("Warehouse (min)", text="Warehouse\n(min)", anchor=tk.CENTER)
        self.report_viewer_sla_table.heading("Weight Bridge 2(min)", text="Weight Bridge\n2(min)", anchor=tk.CENTER)
        self.report_viewer_sla_table.heading("Cargo Covering (min)", text="Cargo Covering\n(min)", anchor=tk.CENTER)
        self.report_viewer_sla_table.heading("Total Time (min)", text="Total Time (min)", anchor=tk.CENTER)
        self.report_viewer_sla_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.report_viewer_sla_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.report_viewer_sla_table_scrollbarY.configure(command=self.report_viewer_sla_table.yview)
        self.report_viewer_sla_table_scrollbarX.configure(command=self.report_viewer_sla_table.xview)
        self.report_viewer_sla_table.configure(yscrollcommand=self.report_viewer_sla_table_scrollbarY.set)
        self.report_viewer_sla_table.configure(xscrollcommand=self.report_viewer_sla_table_scrollbarX.set)
        # SLA tab (button)
        self.report_viewer_sla_button_frame = customtkinter.CTkFrame(self.report_viewer_sla_frame, corner_radius=0, fg_color="transparent")
        self.report_viewer_sla_button_frame.grid(row=2, column=0, sticky="nsew", pady=(5,0))
        self.report_viewer_sla_button = customtkinter.CTkButton(self.report_viewer_sla_button_frame, fg_color="dark blue", text="Export to Excel", hover_color="blue", command=self.export_excel_sla)
        self.report_viewer_sla_button.pack(side="right", padx=(10,0))
        # AuditLog tab
        self.tabview_report_viewer.tab("AuditLog").grid_rowconfigure(0, weight=1)
        self.tabview_report_viewer.tab("AuditLog").grid_columnconfigure(0, weight=1)
        self.report_viewer_auditlog_frame = customtkinter.CTkFrame(self.tabview_report_viewer.tab("AuditLog"), corner_radius=0, fg_color="transparent")
        self.report_viewer_auditlog_frame.grid(row=0, column=0, sticky="nsew")
        self.report_viewer_auditlog_frame.grid_columnconfigure(1, weight=1)
        self.report_viewer_auditlog_frame.grid_rowconfigure(0, weight=1)
        # AuditLog tab (left side)
        self.report_viewer_auditlogLeft_frame = customtkinter.CTkFrame(self.report_viewer_auditlog_frame, corner_radius=0, fg_color="transparent")
        self.report_viewer_auditlogLeft_frame.grid(row=0, column=0, sticky="nsew", padx=(0,5))
        self.report_viewer_auditlogLeft_frame.grid_columnconfigure(0, weight=1)
        self.report_viewer_auditlogLeft_frame.grid_rowconfigure(1, weight=1)
        # AuditLog tab (left side) ENTRY
        self.report_viewer_auditlog_entry_frame = customtkinter.CTkFrame(self.report_viewer_auditlogLeft_frame, corner_radius=10, fg_color="light blue")
        self.report_viewer_auditlog_entry_frame.grid(row=0, column=0, pady=(0,5), sticky="nsew")
        self.report_viewer_auditlog_entry_category_label = customtkinter.CTkLabel(self.report_viewer_auditlog_entry_frame, text="Category :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_auditlog_entry_category_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.report_viewer_auditlog_optionmenu_category = customtkinter.CTkOptionMenu(self.report_viewer_auditlog_entry_frame, values=["AUDIT LOG"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.report_viewer_auditlog_optionmenu_category.grid(row=0,column=1, sticky="ew", columnspan=3)
        self.report_viewer_auditlog_dateStart_label = customtkinter.CTkLabel(self.report_viewer_auditlog_entry_frame, text="Date Start :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_auditlog_dateStart_label.grid(row=1,column=0, padx=10, pady=5, sticky="ew")
        self.report_viewer_auditlog_dateStart_calendar = MyDateEntry(self.report_viewer_auditlog_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_auditlog_dateStart_calendar.grid(row=1,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_auditlog_dateEnd_label = customtkinter.CTkLabel(self.report_viewer_auditlog_entry_frame, text="Date End :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.report_viewer_auditlog_dateEnd_label.grid(row=1,column=2, padx=10, pady=5, sticky="ew")
        self.report_viewer_auditlog_dateEnd_calendar = MyDateEntry(self.report_viewer_auditlog_entry_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.report_viewer_auditlog_dateEnd_calendar.grid(row=1,column=3, padx=(0,10), pady= 5, sticky="ew")
        self.report_viewer_auditlog_search_button = customtkinter.CTkButton(self.report_viewer_auditlog_entry_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search_report_auditlog)
        self.report_viewer_auditlog_search_button.grid(row=1,column=4, padx=(0,10), pady= 5, sticky="ew")
        # AuditLog tab (left side) TABLE
        self.report_viewer_auditlogLeft_table_frame = customtkinter.CTkFrame(self.report_viewer_auditlogLeft_frame, corner_radius=0, fg_color="WHITE")
        self.report_viewer_auditlogLeft_table_frame.grid(row=1, column=0, pady=(0,5), sticky="nsew")
        self.report_viewer_auditlog_table_scrollbarY = customtkinter.CTkScrollbar(self.report_viewer_auditlogLeft_table_frame, hover=True, button_hover_color="dark blue", orientation="vertical")
        self.report_viewer_auditlog_table_scrollbarY.pack(side="right",fill=tk.Y)
        self.report_viewer_auditlog_table_scrollbarX = customtkinter.CTkScrollbar(self.report_viewer_auditlogLeft_table_frame, hover=True, button_hover_color="dark blue", orientation="horizontal")
        self.report_viewer_auditlog_table_scrollbarX.pack(side="bottom",fill=tk.X)
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.report_viewer_auditlog_table = ttk.Treeview(self.report_viewer_auditlogLeft_table_frame)
        self.report_viewer_auditlog_table.pack(side="top", expand=tk.YES, fill=tk.BOTH)
        self.report_viewer_auditlog_table["columns"] = ("Warehouse", "Date", "Time", "Transaction", "Operator", "Status", "Remark")
        self.report_viewer_auditlog_table.column("#0", width=0,stretch=tk.NO)
        self.report_viewer_auditlog_table.column("Warehouse", anchor=tk.CENTER, width=120, minwidth=0,stretch=tk.YES)
        self.report_viewer_auditlog_table.column("Date", anchor=tk.CENTER, width=120, minwidth=0,stretch=tk.YES)
        self.report_viewer_auditlog_table.column("Time", anchor=tk.CENTER, width=120, minwidth=0,stretch=tk.YES)
        self.report_viewer_auditlog_table.column("Transaction", anchor=tk.CENTER, width=120, minwidth=0,stretch=tk.YES)
        self.report_viewer_auditlog_table.column("Operator", anchor=tk.CENTER, width=120, minwidth=0,stretch=tk.YES)
        self.report_viewer_auditlog_table.column("Status", anchor=tk.CENTER, width=120, minwidth=0,stretch=tk.YES)
        self.report_viewer_auditlog_table.column("Remark", anchor=tk.CENTER, width=300, minwidth=0,stretch=tk.YES)
        self.report_viewer_auditlog_table.heading("#0", text="", anchor=tk.W)
        self.report_viewer_auditlog_table.heading("Warehouse", text="Warehouse", anchor=tk.CENTER)
        self.report_viewer_auditlog_table.heading("Date", text="Date", anchor=tk.CENTER)
        self.report_viewer_auditlog_table.heading("Time", text="Time", anchor=tk.CENTER)
        self.report_viewer_auditlog_table.heading("Transaction", text="Transaction", anchor=tk.CENTER)
        self.report_viewer_auditlog_table.heading("Operator", text="Operator", anchor=tk.CENTER)
        self.report_viewer_auditlog_table.heading("Status", text="Status", anchor=tk.CENTER)
        self.report_viewer_auditlog_table.heading("Remark", text="Remark", anchor=tk.CENTER)
        self.report_viewer_auditlog_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.report_viewer_auditlog_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.report_viewer_auditlog_table_scrollbarY.configure(command=self.report_viewer_auditlog_table.yview)
        self.report_viewer_auditlog_table_scrollbarX.configure(command=self.report_viewer_auditlog_table.xview)
        self.report_viewer_auditlog_table.configure(yscrollcommand=self.report_viewer_auditlog_table_scrollbarY.set)
        self.report_viewer_auditlog_table.configure(xscrollcommand=self.report_viewer_auditlog_table_scrollbarX.set)
        self.report_viewer_auditlog_table.bind("<<TreeviewSelect>>", self.on_tree_auditlog_select)
        # AuditLog tab (left side) BUTTON
        self.report_viewer_auditlog_button_frame = customtkinter.CTkFrame(self.report_viewer_auditlogLeft_frame, corner_radius=10, fg_color="transparent")
        self.report_viewer_auditlog_button_frame.grid(row=2, column=0, sticky="nsew")
        self.report_viewer_auditlog_button = customtkinter.CTkButton(self.report_viewer_auditlog_button_frame, fg_color="dark blue", text="Export to Excel", hover_color="blue", command=self.export_excel_auditlog)
        self.report_viewer_auditlog_button.pack(side="left", pady=(10,0))
        # AuditLog tab (right side)
        self.report_viewer_auditlogRight_frame = customtkinter.CTkFrame(self.report_viewer_auditlog_frame, corner_radius=0, fg_color="transparent")
        self.report_viewer_auditlogRight_frame.grid(row=0, column=1, sticky="nsew")
        self.report_viewer_auditlogRight_frame.grid_columnconfigure(0, weight=1)
        self.report_viewer_auditlogRight_frame.grid_rowconfigure(1, weight=1)
        self.report_viewer_auditlogRight_label_remark = customtkinter.CTkLabel(self.report_viewer_auditlogRight_frame, text="Remark :", font=customtkinter.CTkFont(size=15,weight="bold"),anchor="w")
        self.report_viewer_auditlogRight_label_remark.grid(row=0,column=0, sticky="ew", pady=(0,10))
        self.report_viewer_auditlogRight_textbox_remark = customtkinter.CTkTextbox(self.report_viewer_auditlogRight_frame, fg_color="gray20", text_color="white", height=120, font=customtkinter.CTkFont(size=15))
        self.report_viewer_auditlogRight_textbox_remark.grid(row=1,column=0, padx=(5,0), pady=(0,10), sticky="nsew")

        # create queuing settings frame #######################################################################
        self.Qsettings_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.Qsettings_frame.grid_columnconfigure(0, weight=1)
        self.Qsettings_frame.grid_rowconfigure(0, weight=1)
        # create tabview queuing settings
        self.tabview_qsettings = customtkinter.CTkTabview(self.Qsettings_frame,fg_color=("gray92"),text_color="White",segmented_button_fg_color="Dark Blue",segmented_button_unselected_color="Dark Blue")
        self.tabview_qsettings.grid(row=0, column=0, padx=(0, 0), pady=(0, 0), sticky="nsew")
        self.tabview_qsettings.add("Packaging")
        self.tabview_qsettings.add("Product")
        self.tabview_qsettings.add("Tapping Point")
        self.tabview_qsettings.add("Warehouse Flow")
        self.tabview_qsettings.add("Queuing Manual Call")
        #>> configure grid of individual tabs (Packaging) and item
        self.tabview_qsettings.tab("Packaging").grid_rowconfigure(0, weight=1)
        self.tabview_qsettings.tab("Packaging").grid_columnconfigure(0, weight=1)
        self.packaging_registration_frame = customtkinter.CTkFrame(self.tabview_qsettings.tab("Packaging"), corner_radius=0, fg_color="transparent")
        self.packaging_registration_frame.grid(row=0, column=0, sticky="nsew")
        self.warehouse_slaSetting_frame = customtkinter.CTkFrame(self.tabview_qsettings.tab("Packaging"), corner_radius=0, fg_color="transparent")
        self.warehouse_slaSetting_frame.grid(row=0, column=1, sticky="nsew")
        # packaging registration
        self.packaging_registration_frame.grid_columnconfigure(0, weight=1)
        self.packaging_registration_frame.grid_rowconfigure(2, weight=1)
        # packaging registration (title)
        self.packaging_registration_title_frame = customtkinter.CTkFrame(self.packaging_registration_frame, corner_radius=0, fg_color="transparent")
        self.packaging_registration_title_frame.grid(row=0, column=0, sticky="nsew")
        self.packaging_registration_label = customtkinter.CTkLabel(self.packaging_registration_title_frame, text="Packaging Registration", font=customtkinter.CTkFont(size=17),anchor="w")
        self.packaging_registration_label.pack(side="left", padx=10)
        # packaging registration (search)
        self.packaging_registration_search_frame = customtkinter.CTkFrame(self.packaging_registration_frame, corner_radius=10, fg_color="white")
        self.packaging_registration_search_frame.grid(row=1, column=0, sticky="nsew")
        self.search_packaging_registration_label = customtkinter.CTkLabel(self.packaging_registration_search_frame, text="Search :", font=customtkinter.CTkFont(size=15),anchor="w")
        self.search_packaging_registration_label.pack(side="left", padx=10, pady=10)
        self.entry_search_packaging_registration = customtkinter.CTkEntry(self.packaging_registration_search_frame, placeholder_text="", width=250)
        self.entry_search_packaging_registration.pack(side="left")
        self.search_button_packaging_registration = customtkinter.CTkButton(self.packaging_registration_search_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search_packaging_registration)
        self.search_button_packaging_registration.pack(side="left")
        # packaging registration (tabel)
        self.packaging_registration_tabel_frame = customtkinter.CTkFrame(self.packaging_registration_frame, corner_radius=0, fg_color="white")
        self.packaging_registration_tabel_frame.grid(row=2, column=0, sticky="nsew")
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.packaging_table = ttk.Treeview(self.packaging_registration_tabel_frame)
        self.packaging_table.pack(side="left", expand=tk.YES, fill=tk.BOTH)
        self.packaging_table["columns"] = ("Packaging", "Prefix", "Remark", "Status", "SLA")
        self.packaging_table.column("#0", width=0,stretch=tk.NO)
        self.packaging_table.column("Packaging", anchor=tk.CENTER, width=150,minwidth=30,stretch=tk.YES)
        self.packaging_table.column("Prefix", anchor=tk.CENTER, width=50, minwidth=0,stretch=tk.YES)
        self.packaging_table.column("Remark", anchor=tk.CENTER, width=300, minwidth=0,stretch=tk.YES)
        self.packaging_table.column("Status", anchor=tk.CENTER, width=50, minwidth=0,stretch=tk.YES)
        self.packaging_table.column("SLA", anchor=tk.CENTER, width=50, minwidth=0,stretch=tk.YES)
        self.packaging_table.heading("#0", text="", anchor=tk.W)
        self.packaging_table.heading("Packaging", text="Packaging", anchor=tk.CENTER)
        self.packaging_table.heading("Prefix", text="Prefix", anchor=tk.CENTER)
        self.packaging_table.heading("Remark", text="Remark", anchor=tk.CENTER)
        self.packaging_table.heading("Status", text="Status", anchor=tk.CENTER)
        self.packaging_table.heading("SLA", text="SLA", anchor=tk.CENTER)
        self.packaging_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.packaging_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.packaging_table_scrollbar = customtkinter.CTkScrollbar(self.packaging_registration_tabel_frame, hover=True, button_hover_color="dark blue", command=self.packaging_table.yview)
        self.packaging_table_scrollbar.pack(side="left",fill=tk.Y)
        self.packaging_table.configure(yscrollcommand=self.packaging_table_scrollbar.set)
        self.packaging_table.bind("<<TreeviewSelect>>", self.on_tree_packaging_select)
        # packaging registration (entry)
        self.packaging_registration_entry_frame = customtkinter.CTkFrame(self.packaging_registration_frame, corner_radius=10, fg_color="grey80")
        self.packaging_registration_entry_frame.grid(row=3, column=0, sticky="nsew")
        self.entry_packaging_label = customtkinter.CTkLabel(self.packaging_registration_entry_frame, text="Packaging :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_packaging_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_packaging = customtkinter.CTkEntry(self.packaging_registration_entry_frame, placeholder_text="")
        self.entry_packaging.grid(row=0,column=1, sticky="ew", columnspan=3)
        self.entry_prefix = customtkinter.CTkLabel(self.packaging_registration_entry_frame, text="Prefix :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_prefix.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_prefix = customtkinter.CTkEntry(self.packaging_registration_entry_frame, placeholder_text="")
        self.entry_prefix.grid(row=1,column=1, sticky="ew", columnspan=3)
        self.entry_remark_packaging_label = customtkinter.CTkLabel(self.packaging_registration_entry_frame, text="Remark :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_remark_packaging_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_remark_packaging = customtkinter.CTkEntry(self.packaging_registration_entry_frame, placeholder_text="")
        self.entry_remark_packaging.grid(row=2,column=1, sticky="ew", columnspan=3)
        self.entry_sla_packaging_label = customtkinter.CTkLabel(self.packaging_registration_entry_frame, text="Default SLA :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_sla_packaging_label.grid(row=3,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_sla_packaging = customtkinter.CTkEntry(self.packaging_registration_entry_frame, placeholder_text="")
        self.entry_sla_packaging.grid(row=3,column=1, sticky="ew")
        self.entry_status_packaging_label = customtkinter.CTkLabel(self.packaging_registration_entry_frame, text="Status :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_status_packaging_label.grid(row=3,column=2, padx=10, pady= 5, sticky="ew")
        self.optionmenu_status_packaging= customtkinter.CTkOptionMenu(self.packaging_registration_entry_frame, values=["Activate","Not Activate"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb",command=self.optionmenuStatus_callback)
        self.optionmenu_status_packaging.grid(row=3,column=3, sticky="ew")
        self.optionmenu_status_packaging.set("Activate")
        # packaging registration (button)
        self.packaging_registration_button_frame = customtkinter.CTkFrame(self.packaging_registration_frame, corner_radius=0, fg_color="transparent")
        self.packaging_registration_button_frame.grid(row=4, column=0)
        self.add_button_packaging = customtkinter.CTkButton(self.packaging_registration_button_frame, fg_color="dark blue", text="ADD", hover_color="blue", command=self.add_packaging)
        self.add_button_packaging.grid(row=0,column=0)
        self.remove_button_packaging = customtkinter.CTkButton(self.packaging_registration_button_frame, fg_color="dark blue", text="Remove", hover_color="blue", command=self.remove_packaging)
        self.remove_button_packaging.grid(row=0,column=1, padx=20, pady= 10)
        self.replace_button_packaging = customtkinter.CTkButton(self.packaging_registration_button_frame, fg_color="dark blue", text="Replace", hover_color="blue", command=self.replace_packaging)
        self.replace_button_packaging.grid(row=0,column=2, padx=(0,20), pady= 10)
        self.cancel_button_packaging = customtkinter.CTkButton(self.packaging_registration_button_frame, fg_color="dark red", text="Cancel", hover_color="red", command=self.cancel_packaging)
        self.cancel_button_packaging.grid(row=0,column=3)
        # warehouse sla setting
        self.warehouse_slaSetting_frame.grid_columnconfigure(0, weight=1)
        self.warehouse_slaSetting_frame.grid_rowconfigure(1, weight=1)
        # warehouse sla setting (title)
        self.warehouse_slaSetting_title_frame = customtkinter.CTkFrame(self.warehouse_slaSetting_frame, corner_radius=0, fg_color="transparent")
        self.warehouse_slaSetting_title_frame.grid(row=0, column=0, sticky="nsew")
        self.warehouse_slaSetting_label = customtkinter.CTkLabel(self.warehouse_slaSetting_title_frame, text="Warehouse SLA Setting", font=customtkinter.CTkFont(size=17),anchor="w")
        self.warehouse_slaSetting_label.pack(side="left", padx=10)
        # warehouse sla setting (table)
        self.warehouse_slaSetting_table_frame = customtkinter.CTkFrame(self.warehouse_slaSetting_frame, corner_radius=0, fg_color="white")
        self.warehouse_slaSetting_table_frame.grid(row=1, column=0, sticky="nsew")
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.warehouse_slaSetting_table = ttk.Treeview(self.warehouse_slaSetting_table_frame)
        self.warehouse_slaSetting_table.pack(side="left", expand=tk.YES, fill=tk.BOTH)
        self.warehouse_slaSetting_table["columns"] = ("Packaging", "Product Grp", "W Min (kg)", "W Max (kg)", "Wh SLA (min)", "Vehicle", "waiting_time", "w_bridge1", "w_bridge2", "covering_time", "oid")
        self.warehouse_slaSetting_table.column("#0", width=0,stretch=tk.NO)
        self.warehouse_slaSetting_table.column("Packaging", anchor=tk.CENTER, width=100,minwidth=30,stretch=tk.YES)
        self.warehouse_slaSetting_table.column("Product Grp", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.warehouse_slaSetting_table.column("W Min (kg)", anchor=tk.CENTER, width=85, minwidth=0,stretch=tk.YES)
        self.warehouse_slaSetting_table.column("W Max (kg)", anchor=tk.CENTER, width=85, minwidth=0,stretch=tk.YES)
        self.warehouse_slaSetting_table.column("Wh SLA (min)", anchor=tk.CENTER, width=103, minwidth=0,stretch=tk.YES)
        self.warehouse_slaSetting_table.column("Vehicle", anchor=tk.CENTER, width=85, minwidth=0,stretch=tk.YES)
        self.warehouse_slaSetting_table.column("waiting_time", width=0,stretch=tk.NO)
        self.warehouse_slaSetting_table.column("w_bridge1", width=0,stretch=tk.NO)
        self.warehouse_slaSetting_table.column("w_bridge2", width=0,stretch=tk.NO)
        self.warehouse_slaSetting_table.column("covering_time", width=0,stretch=tk.NO)
        self.warehouse_slaSetting_table.column("oid", width=0,stretch=tk.NO)
        self.warehouse_slaSetting_table.heading("#0", text="", anchor=tk.W)
        self.warehouse_slaSetting_table.heading("Packaging", text="Packaging", anchor=tk.CENTER)
        self.warehouse_slaSetting_table.heading("Product Grp", text="Product Grp", anchor=tk.CENTER)
        self.warehouse_slaSetting_table.heading("W Min (kg)", text="W Min (kg)", anchor=tk.CENTER)
        self.warehouse_slaSetting_table.heading("W Max (kg)", text="W Max (kg)", anchor=tk.CENTER)
        self.warehouse_slaSetting_table.heading("Wh SLA (min)", text="Wh SLA (min)", anchor=tk.CENTER)
        self.warehouse_slaSetting_table.heading("Vehicle", text="Vehicle", anchor=tk.CENTER)
        self.warehouse_slaSetting_table.heading("waiting_time", text="", anchor=tk.W)
        self.warehouse_slaSetting_table.heading("w_bridge1", text="", anchor=tk.W)
        self.warehouse_slaSetting_table.heading("w_bridge2", text="", anchor=tk.W)
        self.warehouse_slaSetting_table.heading("covering_time", text="", anchor=tk.W)
        self.warehouse_slaSetting_table.heading("oid", text="", anchor=tk.W)
        self.warehouse_slaSetting_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.warehouse_slaSetting_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.warehouse_slaSetting_table_scrollbar = customtkinter.CTkScrollbar(self.warehouse_slaSetting_table_frame, hover=True, button_hover_color="dark blue", command=self.warehouse_slaSetting_table.yview)
        self.warehouse_slaSetting_table_scrollbar.pack(side="left",fill=tk.Y)
        self.warehouse_slaSetting_table.configure(yscrollcommand=self.warehouse_slaSetting_table_scrollbar.set)
        self.warehouse_slaSetting_table.bind("<<TreeviewSelect>>", self.on_tree_warehouse_slaSetting_select)
        # warehouse sla setting (entry)
        self.warehouse_slaSetting_entry_frame = customtkinter.CTkFrame(self.warehouse_slaSetting_frame, corner_radius=10, fg_color="grey60")
        self.warehouse_slaSetting_entry_frame.grid(row=2, column=0, sticky="nsew")
        self.entry_slaSetting_label = customtkinter.CTkLabel(self.warehouse_slaSetting_entry_frame, text="Packaging :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_slaSetting_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_productGrp_label = customtkinter.CTkLabel(self.warehouse_slaSetting_entry_frame, text="Product Grp :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_productGrp_label.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.optionmenu_packaging_sla= customtkinter.CTkOptionMenu(self.warehouse_slaSetting_entry_frame, values=[""],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", command=self.optionmenuPackaging_callback)
        self.optionmenu_packaging_sla.grid(row=0,column=1, padx=(0,10), pady= 5, columnspan=5, sticky="ew")
        self.optionmenu_packaging_sla.set("Pilih Packaging")
        self.optionmenu_product_sla= customtkinter.CTkOptionMenu(self.warehouse_slaSetting_entry_frame, values=[""],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", command=self.optionmenuProduct_callback)
        self.optionmenu_product_sla.grid(row=1,column=1, padx=(0,10), pady= 5, columnspan=3, sticky="ew")
        self.optionmenu_product_sla.set("Pilih Product")
        self.entry_vehicle_label = customtkinter.CTkLabel(self.warehouse_slaSetting_entry_frame, text="Vehicle type :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_vehicle_label.grid(row=1,column=4, padx=(0,10), pady= 5, sticky="ew")
        self.optionmenu_vehicle_sla= customtkinter.CTkOptionMenu(self.warehouse_slaSetting_entry_frame, values=["ALL"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", dynamic_resizing= False, width=60, command=self.optionmenuVehicle_callback)
        self.optionmenu_vehicle_sla.grid(row=1,column=5, padx=(0,10), pady= 5, sticky="ew")
        self.optionmenu_vehicle_sla.set("ALL")
        self.entry_weightMin_label = customtkinter.CTkLabel(self.warehouse_slaSetting_entry_frame, text="Weight Min (kg):", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_weightMin_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_weightMin = customtkinter.CTkEntry(self.warehouse_slaSetting_entry_frame, placeholder_text="", width=50)
        self.entry_weightMin.grid(row=2,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.entry_weightMax_label = customtkinter.CTkLabel(self.warehouse_slaSetting_entry_frame, text="Max (kg):", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_weightMax_label.grid(row=2,column=2, padx=(0,10), pady= 5, sticky="ew")
        self.entry_weightMax = customtkinter.CTkEntry(self.warehouse_slaSetting_entry_frame, placeholder_text="", width=50)
        self.entry_weightMax.grid(row=2,column=3, padx=(0,10), pady= 5, sticky="ew")
        self.entry_waitingTime_label = customtkinter.CTkLabel(self.warehouse_slaSetting_entry_frame, text="Waiting Time 1 (min):", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_waitingTime_label.grid(row=2,column=4, padx=(0,10), pady= 5, sticky="ew")
        self.entry_waitingTime = customtkinter.CTkEntry(self.warehouse_slaSetting_entry_frame, placeholder_text="", width=50)
        self.entry_waitingTime.grid(row=2,column=5, padx=(0,10), pady= 5, sticky="ew")
        self.entry_wBridge1Min_label = customtkinter.CTkLabel(self.warehouse_slaSetting_entry_frame, text="W Bridge 1 (min):", font=customtkinter.CTkFont(size=13),anchor="e")
        self.entry_wBridge1Min_label.grid(row=3,column=1, padx=10, pady= 5, columnspan=2, sticky="ew")
        self.entry_wBridge1Min = customtkinter.CTkEntry(self.warehouse_slaSetting_entry_frame, placeholder_text="", width=50)
        self.entry_wBridge1Min.grid(row=3,column=3, padx=(0,10), pady= 5, sticky="ew")
        self.entry_warehouseMin_label = customtkinter.CTkLabel(self.warehouse_slaSetting_entry_frame, text="Warehouse (min):", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_warehouseMin_label.grid(row=3,column=4, padx=(0,10), pady= 5, sticky="ew")
        self.entry_warehouseMin = customtkinter.CTkEntry(self.warehouse_slaSetting_entry_frame, placeholder_text="", width=50)
        self.entry_warehouseMin.grid(row=3,column=5, padx=(0,10), pady= 5, sticky="ew")
        self.entry_wBridge2Min_label = customtkinter.CTkLabel(self.warehouse_slaSetting_entry_frame, text="W Bridge 2 (min):", font=customtkinter.CTkFont(size=13),anchor="e")
        self.entry_wBridge2Min_label.grid(row=4,column=1, padx=10, pady= 5, columnspan=2, sticky="ew")
        self.entry_wBridge2Min = customtkinter.CTkEntry(self.warehouse_slaSetting_entry_frame, placeholder_text="", width=50)
        self.entry_wBridge2Min.grid(row=4,column=3, padx=(0,10), pady= 5, sticky="ew")
        self.entry_coveringMin_label = customtkinter.CTkLabel(self.warehouse_slaSetting_entry_frame, text="Cargo Covering (min):", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_coveringMin_label.grid(row=4,column=4, padx=(0,10), pady= 5, sticky="ew")
        self.entry_coveringMin = customtkinter.CTkEntry(self.warehouse_slaSetting_entry_frame, placeholder_text="", width=50)
        self.entry_coveringMin.grid(row=4,column=5, padx=(0,10), pady= 5, sticky="ew")
        # warehouse sla setting (button)
        self.warehouse_slaSetting_button_frame = customtkinter.CTkFrame(self.warehouse_slaSetting_frame, corner_radius=10, fg_color="grey70")
        self.warehouse_slaSetting_button_frame.grid(row=3, column=0, sticky="nsew")
        self.cancel_button_warehouse_slaSetting = customtkinter.CTkButton(self.warehouse_slaSetting_button_frame, fg_color="dark red", text="Cancel", hover_color="red", width=70, command=self.cancel_warehouse_slaSetting)
        self.cancel_button_warehouse_slaSetting.pack(side="right", padx=(0,10))
        self.replace_button_warehouse_slaSetting = customtkinter.CTkButton(self.warehouse_slaSetting_button_frame, fg_color="black", text="Replace", hover_color="gray", width=70, command=self.replace_warehouse_slaSetting)
        self.replace_button_warehouse_slaSetting.pack(side="right", padx=(0,10), pady="10")
        self.remove_button_warehouse_slaSetting = customtkinter.CTkButton(self.warehouse_slaSetting_button_frame, fg_color="black", text="Remove", hover_color="gray", width=70, command=self.remove_warehouse_slaSetting)
        self.remove_button_warehouse_slaSetting.pack(side="right", padx="10", pady="10")
        self.add_button_warehouse_slaSetting = customtkinter.CTkButton(self.warehouse_slaSetting_button_frame, fg_color="black", text="ADD", hover_color="gray", width=70, command=self.add_warehouse_slaSetting)
        self.add_button_warehouse_slaSetting.pack(side="right")
        #>> configure grid of individual tabs (Product) and item
        self.tabview_qsettings.tab("Product").grid_rowconfigure(0, weight=1)
        self.tabview_qsettings.tab("Product").grid_columnconfigure(0, weight=1)
        self.product_registration_frame = customtkinter.CTkFrame(self.tabview_qsettings.tab("Product"), corner_radius=0, fg_color="transparent")
        self.product_registration_frame.grid(row=0, column=0, sticky="nsew")
        self.product_registration_frame.grid_columnconfigure(0, weight=1)
        self.product_registration_frame.grid_rowconfigure(2, weight=1)
        # product registration (title)
        self.product_registration_title_frame = customtkinter.CTkFrame(self.product_registration_frame, corner_radius=0, fg_color="transparent")
        self.product_registration_title_frame.grid(row=0, column=0, sticky="nsew")
        self.product_registration_label = customtkinter.CTkLabel(self.product_registration_title_frame, text="Product Registration", font=customtkinter.CTkFont(size=17),anchor="w")
        self.product_registration_label.pack(side="left", padx=10)
        # product registration (optionmenu packaging and search)
        self.product_registration_optnsearch_frame = customtkinter.CTkFrame(self.product_registration_frame, corner_radius=10, fg_color="light blue")
        self.product_registration_optnsearch_frame.grid(row=1, column=0, sticky="nsew")
        self.product_registration_optnsearch_frame.grid_columnconfigure((0,1), weight=1)
        self.product_registration_optnsearch_frame.grid_rowconfigure(0, weight=0)
        self.product_registration_optmenu_frame = customtkinter.CTkFrame(self.product_registration_optnsearch_frame, corner_radius=0, fg_color="transparent")
        self.product_registration_optmenu_frame.grid(row=0, column=0, sticky="ew")
        self.optionmenu_packaging_label_product = customtkinter.CTkLabel(self.product_registration_optmenu_frame, text="Packaging :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.optionmenu_packaging_label_product.pack(side="left", padx= 10)
        self.optionmenu_packaging_product= customtkinter.CTkOptionMenu(self.product_registration_optmenu_frame, values=[""],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", command=self.optionmenuPackaging_product_callback)
        self.optionmenu_packaging_product.pack(side="left")
        self.optionmenu_packaging_product.set("Pilih Packaging")
        self.product_registration_search_frame = customtkinter.CTkFrame(self.product_registration_optnsearch_frame, corner_radius=0, fg_color="transparent")
        self.product_registration_search_frame.grid(row=0, column=1, sticky="ew")
        self.search_button_product_registration = customtkinter.CTkButton(self.product_registration_search_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search_product_registration)
        self.search_button_product_registration.pack(side="right")
        self.entry_search_product_registration = customtkinter.CTkEntry(self.product_registration_search_frame, placeholder_text="", width=250)
        self.entry_search_product_registration.pack(side="right")
        self.search_product_registration_label = customtkinter.CTkLabel(self.product_registration_search_frame, text="Search :", font=customtkinter.CTkFont(size=15),anchor="w")
        self.search_product_registration_label.pack(side="right", padx=10, pady=10)
        # product registration (product table)
        self.product_registration_table_frame = customtkinter.CTkFrame(self.product_registration_frame, corner_radius=0, fg_color="white")
        self.product_registration_table_frame.grid(row=2, column=0, sticky="nsew")
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.product_table = ttk.Treeview(self.product_registration_table_frame)
        self.product_table.pack(side="left", expand=tk.YES, fill=tk.BOTH)
        self.product_table["columns"] = ("Product Group", "Remark", "Status", "oid")
        self.product_table.column("#0", width=0,stretch=tk.NO)
        self.product_table.column("Product Group", anchor=tk.CENTER, width=100,minwidth=30,stretch=tk.YES)
        self.product_table.column("Remark", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.product_table.column("Status", anchor=tk.CENTER, width=85, minwidth=0,stretch=tk.YES)
        self.product_table.column("oid", width=0,stretch=tk.NO)
        self.product_table.heading("#0", text="", anchor=tk.W)
        self.product_table.heading("Product Group", text="Product Group", anchor=tk.CENTER)
        self.product_table.heading("Remark", text="Remark", anchor=tk.CENTER)
        self.product_table.heading("Status", text="Status", anchor=tk.CENTER)
        self.product_table.heading("oid", text="", anchor=tk.W)
        self.product_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.product_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.warehouse_slaSetting_table_scrollbar = customtkinter.CTkScrollbar(self.product_registration_table_frame, hover=True, button_hover_color="dark blue", command=self.product_table.yview)
        self.warehouse_slaSetting_table_scrollbar.pack(side="left",fill=tk.Y)
        self.product_table.configure(yscrollcommand=self.warehouse_slaSetting_table_scrollbar.set)
        self.product_table.bind("<<TreeviewSelect>>", self.on_tree_product_select)
        # product registration (entry)
        self.product_registration_entry_frame = customtkinter.CTkFrame(self.product_registration_frame, corner_radius=10, fg_color="gray80")
        self.product_registration_entry_frame.grid(row=3, column=0, sticky="nsew")
        self.entry_product_group_label = customtkinter.CTkLabel(self.product_registration_entry_frame, text="Product Group :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_product_group_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_product_group = customtkinter.CTkEntry(self.product_registration_entry_frame, placeholder_text="",width=300)
        self.entry_product_group.grid(row=0,column=1, sticky="ew")
        self.entry_product_group_remark = customtkinter.CTkLabel(self.product_registration_entry_frame, text="Remark :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_product_group_remark.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_product_group_remark = customtkinter.CTkEntry(self.product_registration_entry_frame, placeholder_text="",width=300)
        self.entry_product_group_remark.grid(row=1,column=1, sticky="ew")
        self.optionmenu_product_group_status_label = customtkinter.CTkLabel(self.product_registration_entry_frame, text="Status :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.optionmenu_product_group_status_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.optionmenu_product_group_status= customtkinter.CTkOptionMenu(self.product_registration_entry_frame, values=["Activate","Not Activate"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb",command=self.optionmenuStatus_callback)
        self.optionmenu_product_group_status.grid(row=2,column=1, sticky="ew")
        self.optionmenu_product_group_status.set("Activate")
        # product registration (button)
        self.product_registration_button_frame = customtkinter.CTkFrame(self.product_registration_frame, corner_radius=0, fg_color="transparent")
        self.product_registration_button_frame.grid(row=4, column=0)
        self.add_button_product_registration = customtkinter.CTkButton(self.product_registration_button_frame, fg_color="dark blue", text="ADD", hover_color="blue", command=self.add_product_registration)
        self.add_button_product_registration.grid(row=0,column=0)
        self.remove_button_product_registration = customtkinter.CTkButton(self.product_registration_button_frame, fg_color="dark blue", text="Remove", hover_color="blue", command=self.remove_product_registration)
        self.remove_button_product_registration.grid(row=0,column=1, padx=20, pady= 10)
        self.replace_button_product_registration = customtkinter.CTkButton(self.product_registration_button_frame, fg_color="dark blue", text="Replace", hover_color="blue", command=self.replace_product_registration)
        self.replace_button_product_registration.grid(row=0,column=2, padx=(0,20), pady= 10)
        self.cancel_button_product_registration = customtkinter.CTkButton(self.product_registration_button_frame, fg_color="dark red", text="Cancel", hover_color="red", command=self.cancel_product_registration)
        self.cancel_button_product_registration.grid(row=0,column=3)
        #>> configure grid of individual tabs (Tapping Point) and item
        self.tabview_qsettings.tab("Tapping Point").grid_rowconfigure(0, weight=1)
        self.tabview_qsettings.tab("Tapping Point").grid_columnconfigure(0, weight=1)
        self.tapping_pointReg_frame = customtkinter.CTkFrame(self.tabview_qsettings.tab("Tapping Point"), corner_radius=0, fg_color="transparent")
        self.tapping_pointReg_frame.grid(row=0, column=0, sticky="nsew")
        self.tapping_pointReg_frame.grid_columnconfigure(0, weight=1)
        self.tapping_pointReg_frame.grid_rowconfigure(2, weight=1)
        # tapping point (title)
        self.tapping_point_title_frame = customtkinter.CTkFrame(self.tapping_pointReg_frame, corner_radius=0, fg_color="transparent")
        self.tapping_point_title_frame.grid(row=0, column=0, sticky="nsew")
        self.tapping_point_label = customtkinter.CTkLabel(self.tapping_point_title_frame, text="Tapping Point Registration", font=customtkinter.CTkFont(size=17),anchor="w")
        self.tapping_point_label.pack(side="left", padx=10)
        # tapping point (search)
        self.tapping_point_search_frame = customtkinter.CTkFrame(self.tapping_pointReg_frame, corner_radius=10, fg_color="light blue")
        self.tapping_point_search_frame.grid(row=1, column=0, sticky="nsew")
        self.search_tapping_point_label = customtkinter.CTkLabel(self.tapping_point_search_frame, text="Search :", font=customtkinter.CTkFont(size=15),anchor="w")
        self.search_tapping_point_label.pack(side="left", padx=10, pady=10)
        self.entry_search_tapping_point = customtkinter.CTkEntry(self.tapping_point_search_frame, placeholder_text="", width=250)
        self.entry_search_tapping_point.pack(side="left")
        self.search_button_tapping_point = customtkinter.CTkButton(self.tapping_point_search_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search_tapping_point)
        self.search_button_tapping_point.pack(side="left")
        # tapping point (table)
        self.tapping_point_table_frame = customtkinter.CTkFrame(self.tapping_pointReg_frame, corner_radius=0, fg_color="white")
        self.tapping_point_table_frame.grid(row=2, column=0, sticky="nsew")
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.tappingPoint_table = ttk.Treeview(self.tapping_point_table_frame)
        self.tappingPoint_table.pack(side="left", expand=tk.YES, fill=tk.BOTH)
        self.tappingPoint_table["columns"] = ("HW ID", "Tap ID", "Remark", "Packaging", "Product", "Activity", "Block", "MAX", "Function", "Status", "Alias", "oid")
        self.tappingPoint_table.column("#0", width=0,stretch=tk.NO)
        self.tappingPoint_table.column("HW ID", anchor=tk.CENTER, width=20,minwidth=20,stretch=tk.YES)
        self.tappingPoint_table.column("Tap ID", anchor=tk.CENTER, width=20, minwidth=0,stretch=tk.YES)
        self.tappingPoint_table.column("Remark", anchor=tk.CENTER, width=100, minwidth=0,stretch=tk.YES)
        self.tappingPoint_table.column("Packaging", anchor=tk.CENTER, width=200,minwidth=30,stretch=tk.YES)
        self.tappingPoint_table.column("Product", anchor=tk.CENTER, width=80, minwidth=0,stretch=tk.YES)
        self.tappingPoint_table.column("Activity", anchor=tk.CENTER, width=80, minwidth=0,stretch=tk.YES)
        self.tappingPoint_table.column("Block", anchor=tk.CENTER, width=20,minwidth=30,stretch=tk.YES)
        self.tappingPoint_table.column("MAX", anchor=tk.CENTER, width=20, minwidth=0,stretch=tk.YES)
        self.tappingPoint_table.column("Function", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.tappingPoint_table.column("Status", anchor=tk.CENTER, width=80,minwidth=30,stretch=tk.YES)
        self.tappingPoint_table.column("Alias", anchor=tk.CENTER, width=30, minwidth=0,stretch=tk.YES)
        self.tappingPoint_table.column("oid", width=0,stretch=tk.NO)
        self.tappingPoint_table.heading("#0", text="", anchor=tk.W)
        self.tappingPoint_table.heading("HW ID", text="HW ID", anchor=tk.CENTER)
        self.tappingPoint_table.heading("Tap ID", text="Tap ID", anchor=tk.CENTER)
        self.tappingPoint_table.heading("Remark", text="Remark", anchor=tk.CENTER)
        self.tappingPoint_table.heading("Packaging", text="Packaging", anchor=tk.CENTER)
        self.tappingPoint_table.heading("Product", text="Product", anchor=tk.CENTER)
        self.tappingPoint_table.heading("Activity", text="Activity", anchor=tk.CENTER)
        self.tappingPoint_table.heading("Block", text="Block", anchor=tk.CENTER)
        self.tappingPoint_table.heading("MAX", text="MAX", anchor=tk.CENTER)
        self.tappingPoint_table.heading("Function", text="Function", anchor=tk.CENTER)
        self.tappingPoint_table.heading("Status", text="Status", anchor=tk.CENTER)
        self.tappingPoint_table.heading("Alias", text="Alias", anchor=tk.CENTER)
        self.tappingPoint_table.heading("oid", text="", anchor=tk.W)
        self.tappingPoint_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.tappingPoint_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.tappingPoint_table_scrollbar = customtkinter.CTkScrollbar(self.tapping_point_table_frame, hover=True, button_hover_color="dark blue", command=self.tappingPoint_table.yview)
        self.tappingPoint_table_scrollbar.pack(side="left",fill=tk.Y)
        self.tappingPoint_table.configure(yscrollcommand=self.tappingPoint_table_scrollbar.set)
        self.tappingPoint_table.bind("<<TreeviewSelect>>", self.on_tree_tappingPoin_select)
        # tapping point (entry)
        self.tapping_point_entry_frame = customtkinter.CTkFrame(self.tapping_pointReg_frame, corner_radius=10, fg_color="gray70")
        self.tapping_point_entry_frame.grid(row=3, column=0, sticky="nsew")
        self.entry_machineId_label = customtkinter.CTkLabel(self.tapping_point_entry_frame, text="Machine ID :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_machineId_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_machineId = customtkinter.CTkEntry(self.tapping_point_entry_frame, placeholder_text="")
        self.entry_machineId.grid(row=0,column=1, columnspan=3, sticky="ew")
        self.entry_tapId_label = customtkinter.CTkLabel(self.tapping_point_entry_frame, text="Tap ID :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_tapId_label.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_tapId = customtkinter.CTkEntry(self.tapping_point_entry_frame, placeholder_text="")
        self.entry_tapId.grid(row=1,column=1, sticky="ew")
        self.entry_aliasId_label = customtkinter.CTkLabel(self.tapping_point_entry_frame, text="Alias ID :", font=customtkinter.CTkFont(size=13),anchor="center")
        self.entry_aliasId_label.grid(row=1,column=2, padx=10, pady= 5, sticky="ew")
        self.entry_aliasId = customtkinter.CTkEntry(self.tapping_point_entry_frame, placeholder_text="")
        self.entry_aliasId.grid(row=1,column=3, sticky="ew")
        self.packaging_tappingPoint_label = customtkinter.CTkLabel(self.tapping_point_entry_frame, text="Packaging :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.packaging_tappingPoint_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.checkbox_packaging_tappingPoint_frame = customtkinter.CTkFrame(self.tapping_point_entry_frame, corner_radius=10, fg_color="gray70")
        self.checkbox_packaging_tappingPoint_frame.grid(row=2, column=1, padx=(10,10), pady= 5, sticky="ew", columnspan=3)
        self.checkbox_packaging_tappingPoint_all = customtkinter.CTkCheckBox(self.checkbox_packaging_tappingPoint_frame, text="ALL", width=20, onvalue="ALL", command=self.checkBox_All_tappingPoint)
        self.checkbox_packaging_tappingPoint_all.grid(row=0, column=0, padx=(0,10))
        self.checkbox_packaging_tappingPoint_zak = customtkinter.CTkCheckBox(self.checkbox_packaging_tappingPoint_frame, text="ZAK", width=20, onvalue="ZAK", command=self.checkBox_ZAK_tappingPoint)
        self.checkbox_packaging_tappingPoint_zak.grid(row=0, column=1, padx=(0,10))
        self.checkbox_packaging_tappingPoint_curah = customtkinter.CTkCheckBox(self.checkbox_packaging_tappingPoint_frame, text="CURAH", width=20, onvalue="CURAH", command=self.checkBox_CURAH_tappingPoint)
        self.checkbox_packaging_tappingPoint_curah.grid(row=0, column=2, padx=(0,10))
        self.checkbox_packaging_tappingPoint_jumbo = customtkinter.CTkCheckBox(self.checkbox_packaging_tappingPoint_frame, text="JUMBO", width=20, onvalue="JUMBO", command=self.checkBox_JUMBO_tappingPoint)
        self.checkbox_packaging_tappingPoint_jumbo.grid(row=0, column=3, padx=(0,10))
        self.checkbox_packaging_tappingPoint_pallet = customtkinter.CTkCheckBox(self.checkbox_packaging_tappingPoint_frame, text="PALLET", width=20, onvalue="PALLET", command=self.checkBox_PALLET_tappingPoint)
        self.checkbox_packaging_tappingPoint_pallet.grid(row=0, column=4)
        self.product_tappingPoint_label = customtkinter.CTkLabel(self.tapping_point_entry_frame, text="Product :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.product_tappingPoint_label.grid(row=3,column=0, padx=10, pady= 5, sticky="ew")
        self.optionmenu_product_tappingPoint= customtkinter.CTkOptionMenu(self.tapping_point_entry_frame, values=productTapPoint,fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.optionmenu_product_tappingPoint.grid(row=3,column=1, columnspan=3, sticky="ew")
        self.optionmenu_product_tappingPoint.set("-")
        self.activity_tappingPoint_label = customtkinter.CTkLabel(self.tapping_point_entry_frame, text="Activity :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.activity_tappingPoint_label.grid(row=4,column=0, padx=10, pady= 5, sticky="ew")
        self.optionmenu_activity_tappingPoint= customtkinter.CTkOptionMenu(self.tapping_point_entry_frame, values=["ALL","LOADING","UNLOADING"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.optionmenu_activity_tappingPoint.grid(row=4,column=1,sticky="ew")
        self.optionmenu_activity_tappingPoint.set("Pilih Activity")
        self.max_tappingPoint_label = customtkinter.CTkLabel(self.tapping_point_entry_frame, text="Max :", font=customtkinter.CTkFont(size=13),anchor="center")
        self.max_tappingPoint_label.grid(row=4,column=2, padx=10, pady= 5, sticky="ew")
        self.entry_max_tappingPoint = customtkinter.CTkEntry(self.tapping_point_entry_frame, placeholder_text="")
        self.entry_max_tappingPoint.grid(row=4,column=3,sticky="ew")
        self.entry_remark_tappingpoint_label = customtkinter.CTkLabel(self.tapping_point_entry_frame, text="Remark :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_remark_tappingpoint_label.grid(row=0,column=5, padx=10, pady= 5, sticky="ew")
        self.entry_remark_tappingpoint = customtkinter.CTkEntry(self.tapping_point_entry_frame, placeholder_text="")
        self.entry_remark_tappingpoint.grid(row=0,column=6, columnspan=3, sticky="ew")
        self.function_tappingpoint_label = customtkinter.CTkLabel(self.tapping_point_entry_frame, text="Function :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.function_tappingpoint_label.grid(row=1,column=5, padx=10, pady= 5, sticky="ew")
        self.optionmenu_function_tappingPoint= customtkinter.CTkOptionMenu(self.tapping_point_entry_frame, values=["Pairing/Registration","Weight Bridge 1","Warehouse","Weight Bridge 2","Cargo Covering","Unpairing/Unregistration"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.optionmenu_function_tappingPoint.grid(row=1,column=6, columnspan=3, sticky="ew")
        self.optionmenu_function_tappingPoint.set("Pilih Function")
        self.status_tappingpoint_label = customtkinter.CTkLabel(self.tapping_point_entry_frame, text="Status :", font=customtkinter.CTkFont(size=13),anchor="w")
        self.status_tappingpoint_label.grid(row=2,column=5, padx=10, pady= 5, sticky="ew")
        self.optionmenu_status_tappingPoint= customtkinter.CTkOptionMenu(self.tapping_point_entry_frame, values=["Activate","Not Activate"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", command=self.optionmenuStatus_callback)
        self.optionmenu_status_tappingPoint.grid(row=2,column=6, columnspan=3, sticky="ew")
        self.optionmenu_status_tappingPoint.set("Activate")
        self.tapBlocking_tappingpoint_label = customtkinter.CTkLabel(self.tapping_point_entry_frame, text="Tap Blocking (sec):", font=customtkinter.CTkFont(size=13),anchor="w")
        self.tapBlocking_tappingpoint_label.grid(row=3,column=5, columnspan=2, padx=10, pady= 5, sticky="ew")
        self.entry_tapBlocking_tappingpoint = customtkinter.CTkEntry(self.tapping_point_entry_frame, placeholder_text="")
        self.entry_tapBlocking_tappingpoint.grid(row=3,column=7, columnspan=2, sticky="ew")
        self.cancel_button_tappingpoint = customtkinter.CTkButton(self.tapping_point_entry_frame, fg_color="dark red", text="Cancel", hover_color="red", width=60, command=self.cancel_tapping_point)
        self.cancel_button_tappingpoint.grid(row=4,column=8, rowspan=2)
        # tapping point (button)
        self.tapping_point_button_frame = customtkinter.CTkFrame(self.tapping_pointReg_frame, corner_radius=0, fg_color="transparent")
        self.tapping_point_button_frame.grid(row=4, column=0)
        self.add_button_tapping_point = customtkinter.CTkButton(self.tapping_point_button_frame, fg_color="dark blue", text="ADD", hover_color="blue", command=self.add_tapping_point)
        self.add_button_tapping_point.grid(row=0,column=0)
        self.remove_button_tapping_point = customtkinter.CTkButton(self.tapping_point_button_frame, fg_color="dark blue", text="Remove", hover_color="blue", command=self.remove_tapping_point)
        self.remove_button_tapping_point.grid(row=0,column=1, padx=20, pady= 10)
        self.replace_button_tapping_point = customtkinter.CTkButton(self.tapping_point_button_frame, fg_color="dark blue", text="Replace", hover_color="blue", command=self.replace_tapping_point)
        self.replace_button_tapping_point.grid(row=0,column=2)
        #>> configure grid of individual tabs (Warehouse Flow) and item
        self.tabview_qsettings.tab("Warehouse Flow").grid_rowconfigure(0, weight=1)
        self.tabview_qsettings.tab("Warehouse Flow").grid_columnconfigure(0, weight=1)
        self.warehouse_flow_frame = customtkinter.CTkFrame(self.tabview_qsettings.tab("Warehouse Flow"), corner_radius=0, fg_color="transparent")
        self.warehouse_flow_frame.grid(row=0, column=0, sticky="nsew")
        self.warehouse_flow_frame.grid_columnconfigure(0, weight=1)
        self.warehouse_flow_frame.grid_rowconfigure(1, weight=1)
        # warehouse flow setting (title)
        self.warehouse_flow_title_frame = customtkinter.CTkFrame(self.warehouse_flow_frame, corner_radius=0, fg_color="transparent")
        self.warehouse_flow_title_frame.grid(row=0, column=0, sticky="nsew")
        self.warehouse_flow_label = customtkinter.CTkLabel(self.warehouse_flow_title_frame, text="Warehouse Flow Setting", font=customtkinter.CTkFont(size=17),anchor="w")
        self.warehouse_flow_label.pack(side="left", padx=10)
        # warehouse flow setting (table)
        self.warehouse_flow_table_frame = customtkinter.CTkFrame(self.warehouse_flow_frame, corner_radius=0, fg_color="white")
        self.warehouse_flow_table_frame.grid(row=1, column=0, sticky="nsew")
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.warehouse_flow_table = ttk.Treeview(self.warehouse_flow_table_frame)
        self.warehouse_flow_table.pack(side="left", expand=tk.YES, fill=tk.BOTH)
        self.warehouse_flow_table["columns"] = ("Packaging", "Activity", "Registration", "Weight Bridge 1", "Warehouse", "Weight Bridge 2", "Cargo Covering", "Unregistration", "Status", "oid")
        self.warehouse_flow_table.column("#0", width=0,stretch=tk.NO)
        self.warehouse_flow_table.column("Packaging", anchor=tk.CENTER, width=60,minwidth=20,stretch=tk.YES)
        self.warehouse_flow_table.column("Activity", anchor=tk.CENTER, width=60, minwidth=0,stretch=tk.YES)
        self.warehouse_flow_table.column("Registration", anchor=tk.CENTER, width=50, minwidth=0,stretch=tk.YES)
        self.warehouse_flow_table.column("Weight Bridge 1", anchor=tk.CENTER, width=50,minwidth=30,stretch=tk.YES)
        self.warehouse_flow_table.column("Warehouse", anchor=tk.CENTER, width=50, minwidth=0,stretch=tk.YES)
        self.warehouse_flow_table.column("Weight Bridge 2", anchor=tk.CENTER, width=50, minwidth=0,stretch=tk.YES)
        self.warehouse_flow_table.column("Cargo Covering", anchor=tk.CENTER, width=50,minwidth=30,stretch=tk.YES)
        self.warehouse_flow_table.column("Unregistration", anchor=tk.CENTER, width=50, minwidth=0,stretch=tk.YES)
        self.warehouse_flow_table.column("Status", anchor=tk.CENTER, width=30, minwidth=0,stretch=tk.YES)
        self.warehouse_flow_table.column("oid", width=0,stretch=tk.NO)
        self.warehouse_flow_table.heading("#0", text="", anchor=tk.W)
        self.warehouse_flow_table.heading("Packaging", text="Packaging", anchor=tk.CENTER)
        self.warehouse_flow_table.heading("Activity", text="Activity", anchor=tk.CENTER)
        self.warehouse_flow_table.heading("Registration", text="Registration", anchor=tk.CENTER)
        self.warehouse_flow_table.heading("Weight Bridge 1", text="Weight Bridge 1", anchor=tk.CENTER)
        self.warehouse_flow_table.heading("Warehouse", text="Warehouse", anchor=tk.CENTER)
        self.warehouse_flow_table.heading("Weight Bridge 2", text="Weight Bridge 2", anchor=tk.CENTER)
        self.warehouse_flow_table.heading("Cargo Covering", text="Cargo Covering", anchor=tk.CENTER)
        self.warehouse_flow_table.heading("Unregistration", text="Unregistration", anchor=tk.CENTER)
        self.warehouse_flow_table.heading("Status", text="Status", anchor=tk.CENTER)
        self.warehouse_flow_table.heading("oid", text="", anchor=tk.W)
        self.warehouse_flow_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.warehouse_flow_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.warehouse_flow_table_scrollbar = customtkinter.CTkScrollbar(self.warehouse_flow_table_frame, hover=True, button_hover_color="dark blue", command=self.warehouse_flow_table.yview)
        self.warehouse_flow_table_scrollbar.pack(side="left",fill=tk.Y)
        self.warehouse_flow_table.configure(yscrollcommand=self.warehouse_flow_table_scrollbar.set)
        self.warehouse_flow_table.bind("<<TreeviewSelect>>", self.on_tree_warehouseFlow_select)
        # warehouse flow setting (entry)
        self.warehouse_flow_entry_frame = customtkinter.CTkFrame(self.warehouse_flow_frame, corner_radius=10, fg_color="gray70")
        self.warehouse_flow_entry_frame.grid(row=2, column=0, sticky="nsew")
        self.optionmenu_packaging_warehouseflow_label = customtkinter.CTkLabel(self.warehouse_flow_entry_frame, text="Packaging:", font=customtkinter.CTkFont(size=13),anchor="w")
        self.optionmenu_packaging_warehouseflow_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.optionmenu_packaging_warehouseflow= customtkinter.CTkOptionMenu(self.warehouse_flow_entry_frame, values=[""],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", command=self.optionmenuPackaging_callback)
        self.optionmenu_packaging_warehouseflow.grid(row=0,column=1, columnspan=2, pady= 5, sticky="ew")
        self.optionmenu_packaging_warehouseflow.set("Pilih Packaging")
        self.optionmenu_activity_warehouseflow_label = customtkinter.CTkLabel(self.warehouse_flow_entry_frame, text="Activity:", font=customtkinter.CTkFont(size=13),anchor="w")
        self.optionmenu_activity_warehouseflow_label.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.optionmenu_activity_warehouseflow= customtkinter.CTkOptionMenu(self.warehouse_flow_entry_frame, values=["LOADING","UNLOADING"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb")
        self.optionmenu_activity_warehouseflow.grid(row=1,column=1, columnspan=2,sticky="ew")
        self.optionmenu_activity_warehouseflow.set("Pilih Activity")
        self.optionmenu_status_warehouseflow_label = customtkinter.CTkLabel(self.warehouse_flow_entry_frame, text="Status:", font=customtkinter.CTkFont(size=13),anchor="w")
        self.optionmenu_status_warehouseflow_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.optionmenu_status_warehouseflow= customtkinter.CTkOptionMenu(self.warehouse_flow_entry_frame, values=["Activate","Not Activate"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", command=self.optionmenuStatus_callback)
        self.optionmenu_status_warehouseflow.grid(row=2,column=1, columnspan=2, sticky="ew")
        self.optionmenu_status_warehouseflow.set("Activate")
        self.checkbox_tappingPoint_warehouseflow_label = customtkinter.CTkLabel(self.warehouse_flow_entry_frame, text="Tapping Point:", font=customtkinter.CTkFont(size=13),anchor="w")
        self.checkbox_tappingPoint_warehouseflow_label.grid(row=3,column=0, padx=10, pady= 5, sticky="ew")
        self.checkbox_tappingPoint_warehouseflow_registration = customtkinter.CTkCheckBox(self.warehouse_flow_entry_frame, text="Registration", width=20, onvalue="YES")
        self.checkbox_tappingPoint_warehouseflow_registration.grid(row=3, column=1, padx=(0,10), sticky="ew")
        self.checkbox_tappingPoint_warehouseflow_weightbridge2 = customtkinter.CTkCheckBox(self.warehouse_flow_entry_frame, text="Weight Bridge 2", width=20, onvalue="YES")
        self.checkbox_tappingPoint_warehouseflow_weightbridge2.grid(row=3, column=2, sticky="ew")
        self.checkbox_tappingPoint_warehouseflow_weightbridge1 = customtkinter.CTkCheckBox(self.warehouse_flow_entry_frame, text="Weight Bridge 1", width=20, onvalue="YES")
        self.checkbox_tappingPoint_warehouseflow_weightbridge1.grid(row=4, column=1, padx=(0,10), sticky="ew")
        self.checkbox_tappingPoint_warehouseflow_covering = customtkinter.CTkCheckBox(self.warehouse_flow_entry_frame, text="Cargo Covering", width=20, onvalue="YES")
        self.checkbox_tappingPoint_warehouseflow_covering.grid(row=4, column=2, sticky="ew")
        self.checkbox_tappingPoint_warehouseflow_warehouse = customtkinter.CTkCheckBox(self.warehouse_flow_entry_frame, text="Warehouse", width=20, onvalue="YES")
        self.checkbox_tappingPoint_warehouseflow_warehouse.grid(row=5, column=1, padx=(0,10), sticky="ew", pady=(7,5))
        self.checkbox_tappingPoint_warehouseflow_unregistration = customtkinter.CTkCheckBox(self.warehouse_flow_entry_frame, text="Unregistration", width=20, onvalue="YES")
        self.checkbox_tappingPoint_warehouseflow_unregistration.grid(row=5, column=2, sticky="ew", pady=(7,5))
        # warehouse flow setting (button)
        self.warehouse_flow_button_frame = customtkinter.CTkFrame(self.warehouse_flow_frame, corner_radius=0, fg_color="transparent")
        self.warehouse_flow_button_frame.grid(row=3, column=0)
        self.add_button_warehouse_flow = customtkinter.CTkButton(self.warehouse_flow_button_frame, fg_color="dark blue", text="ADD", hover_color="blue", command=self.add_warehouse_flow)
        self.add_button_warehouse_flow.grid(row=0,column=0)
        self.remove_button_warehouse_flow = customtkinter.CTkButton(self.warehouse_flow_button_frame, fg_color="dark blue", text="Remove", hover_color="blue", command=self.remove_warehouse_flow)
        self.remove_button_warehouse_flow.grid(row=0,column=1, padx=20, pady= 10)
        self.replace_button_warehouse_flow = customtkinter.CTkButton(self.warehouse_flow_button_frame, fg_color="dark blue", text="Replace", hover_color="blue", command=self.replace_warehouse_flow)
        self.replace_button_warehouse_flow.grid(row=0,column=2, padx=(0,20), pady= 10)
        self.cancel_button_warehouse_flow = customtkinter.CTkButton(self.warehouse_flow_button_frame, fg_color="dark red", text="Cancel", hover_color="red", command=self.cancel_warehouse_flow)
        self.cancel_button_warehouse_flow.grid(row=0,column=3)
        #>> configure grid of individual tabs (Queuing Manual Call) and item
        self.tabview_qsettings.tab("Queuing Manual Call").grid_rowconfigure(0, weight=1)
        self.tabview_qsettings.tab("Queuing Manual Call").grid_columnconfigure(0, weight=1)
        self.rfidManual_frame = customtkinter.CTkFrame(self.tabview_qsettings.tab("Queuing Manual Call"), corner_radius=0, fg_color="transparent")
        self.rfidManual_frame.grid(row=0, column=0, sticky="nsew")
        self.rfidManual_frame.grid_columnconfigure(0, weight=1)
        self.rfidManual_frame.grid_rowconfigure((1,2,5), weight=1)
        # Queuing Manual Call (title Bluetooth Pairing Status dan Queuing)
        self.rfidPairing_title_frame = customtkinter.CTkFrame(self.rfidManual_frame, corner_radius=0, fg_color="transparent")
        self.rfidPairing_title_frame.grid(row=0, column=0, sticky="nsew")
        self.rfidPairing_label = customtkinter.CTkLabel(self.rfidPairing_title_frame, text="Bluetooth Pairing Status dan Queuing", font=customtkinter.CTkFont(size=17),anchor="w")
        self.rfidPairing_label.pack(side="left")
        # Queuing Manual Call (Tabel Bluetooth Pairing Status dan Queuing)
        self.rfidPairing_table_frame = customtkinter.CTkFrame(self.rfidManual_frame, corner_radius=0, fg_color="gray80")
        self.rfidPairing_table_frame.grid(row=1, column=0, sticky="nsew")
        self.rfidPairing_table_scrollbarY = customtkinter.CTkScrollbar(self.rfidPairing_table_frame, hover=True, button_hover_color="dark blue", orientation="vertical")
        self.rfidPairing_table_scrollbarY.pack(side="right",fill=tk.Y)
        self.rfidPairing_table_scrollbarX = customtkinter.CTkScrollbar(self.rfidPairing_table_frame, hover=True, button_hover_color="dark blue", orientation="horizontal")
        self.rfidPairing_table_scrollbarX.pack(side="bottom",fill=tk.X)
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.rfidPairing_table = ttk.Treeview(self.rfidPairing_table_frame)
        self.rfidPairing_table.pack(side="top", expand=tk.YES, fill=tk.BOTH)
        self.rfidPairing_table["columns"] = ("RFID", "Ticket", "Police No", "Activity", "Packaging", "Product", "Arrival", "Calling", "Pairing", "WB 1-IN", "WB 1-Out", "WH-In", "WH-Out", "WB 2-IN", "WB 2-Out", "CC-In", "CC-Out", "Unpairing", "Total", "Wh ID", "oid")
        self.rfidPairing_table.column("#0", width=0,stretch=tk.NO)
        self.rfidPairing_table.column("RFID", anchor=tk.CENTER, width=150,minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("Ticket", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("Police No", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("Activity", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("Packaging", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("Product", anchor=tk.CENTER, width=0, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("Arrival", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("Calling", anchor=tk.CENTER, width=150,minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("Pairing", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("WB 1-IN", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("WB 1-Out", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("WH-In", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("WH-Out", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("WB 2-IN", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("WB 2-Out", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("CC-In", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("CC-Out", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("Unpairing", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("Total", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("Wh ID", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.NO)
        self.rfidPairing_table.column("oid", width=0,stretch=tk.NO)
        self.rfidPairing_table.heading("#0", text="", anchor=tk.W)
        self.rfidPairing_table.heading("RFID", text="BLE Name", anchor=tk.CENTER)
        self.rfidPairing_table.heading("Ticket", text="Ticket", anchor=tk.CENTER)
        self.rfidPairing_table.heading("Police No", text="Police No", anchor=tk.CENTER)
        self.rfidPairing_table.heading("Activity", text="Activity", anchor=tk.CENTER)
        self.rfidPairing_table.heading("Packaging", text="Packaging", anchor=tk.CENTER)
        self.rfidPairing_table.heading("Product", text="Product", anchor=tk.CENTER)
        self.rfidPairing_table.heading("Arrival", text="Arrival", anchor=tk.CENTER)
        self.rfidPairing_table.heading("Calling", text="Calling", anchor=tk.CENTER)
        self.rfidPairing_table.heading("Pairing", text="Pairing", anchor=tk.CENTER)
        self.rfidPairing_table.heading("WB 1-IN", text="WB 1-IN", anchor=tk.CENTER)
        self.rfidPairing_table.heading("WB 1-Out", text="WB 1-Out", anchor=tk.CENTER)
        self.rfidPairing_table.heading("WH-In", text="WH-In", anchor=tk.CENTER)
        self.rfidPairing_table.heading("WH-Out", text="WH-Out", anchor=tk.CENTER)
        self.rfidPairing_table.heading("WB 2-IN", text="WB 2-IN", anchor=tk.CENTER)
        self.rfidPairing_table.heading("WB 2-Out", text="WB 2-Out", anchor=tk.CENTER)
        self.rfidPairing_table.heading("CC-In", text="CC-In", anchor=tk.CENTER)
        self.rfidPairing_table.heading("CC-Out", text="CC-Out", anchor=tk.CENTER)
        self.rfidPairing_table.heading("Unpairing", text="Unpairing", anchor=tk.CENTER)
        self.rfidPairing_table.heading("Total", text="Total (Minutes)", anchor=tk.CENTER)
        self.rfidPairing_table.heading("Wh ID", text="Wh ID", anchor=tk.CENTER)
        self.rfidPairing_table.heading("oid", text="", anchor=tk.W)
        self.rfidPairing_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.rfidPairing_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.rfidPairing_table_scrollbarY.configure(command=self.rfidPairing_table.yview)
        self.rfidPairing_table_scrollbarX.configure(command=self.rfidPairing_table.xview)
        self.rfidPairing_table.configure(yscrollcommand=self.rfidPairing_table_scrollbarY.set)
        self.rfidPairing_table.configure(xscrollcommand=self.rfidPairing_table_scrollbarX.set)
        self.rfidPairing_table.bind("<<TreeviewSelect>>", self.on_tree_rfidPairing_select)
        # # Queuing Manual Call (textbox for log)
        # self.rfidPairing_log_frame = customtkinter.CTkFrame(self.rfidManual_frame, corner_radius=0, fg_color="transparent")
        # self.rfidPairing_log_frame.grid(row=2, column=0, sticky="nsew")
        # self.rfidPairing_log_textbox = customtkinter.CTkTextbox(self.rfidPairing_log_frame, fg_color="gray20", text_color="white", height=120)
        # self.rfidPairing_log_textbox.pack(expand=tk.YES, fill=tk.BOTH)
        # Queuing Manual Call (entry & button Bluetooth Pairing Status dan Queuing)
        self.rfidPairing_entryButton_frame = customtkinter.CTkFrame(self.rfidManual_frame, corner_radius=0, fg_color="transparent")
        self.rfidPairing_entryButton_frame.grid(row=3, column=0, sticky="nsew", pady=5)
        self.rfidPairing_entryButton_frame.grid_columnconfigure(0, weight=1)
        self.rfidPairing_entryButton_frame.grid_rowconfigure(0, weight=1)
        self.rfidPairing_entryButton_leftFrame = customtkinter.CTkFrame(self.rfidPairing_entryButton_frame, corner_radius=0, fg_color="transparent")
        self.rfidPairing_entryButton_leftFrame.grid(row=0,column=0, sticky="ew")
        self.refresh_button_rfidPairing = customtkinter.CTkButton(self.rfidPairing_entryButton_leftFrame, fg_color="dark blue", text="Refresh", hover_color="blue", width=30, command=self.refresh_pairing_status)
        self.refresh_button_rfidPairing.pack(side="left", padx=(0,10))
        self.rfidPairing_entryButton_rightFrame = customtkinter.CTkFrame(self.rfidPairing_entryButton_frame, corner_radius=0, fg_color="transparent")
        self.rfidPairing_entryButton_rightFrame.grid(row=0,column=1, sticky="ew")
        self.cancel_button_rfidPairing = customtkinter.CTkButton(self.rfidPairing_entryButton_rightFrame, fg_color="dark red", text="Cancel", hover_color="red", width=30, command=self.cancel_pairing_status)
        self.cancel_button_rfidPairing.pack(side="right", padx=(10,0))
        self.set_button_rfidPairing = customtkinter.CTkButton(self.rfidPairing_entryButton_rightFrame, fg_color="dark blue", text="Set", hover_color="blue", width=30, command=self.set_warehouse_pairing_status)
        self.set_button_rfidPairing.pack(side="right", padx=(10,0))
        self.optionmenu_warehouse_rfidPairing = customtkinter.CTkOptionMenu(self.rfidPairing_entryButton_rightFrame, values=[""],fg_color="gray70",text_color="black",button_color="gray60",button_hover_color="#287ceb")
        self.optionmenu_warehouse_rfidPairing.pack(side="right", padx=(10,0))
        self.optionmenu_warehouse_rfidPairing_label = customtkinter.CTkLabel(self.rfidPairing_entryButton_rightFrame, text="Assign to warehouse:", font=customtkinter.CTkFont(size=13),anchor="w")
        self.optionmenu_warehouse_rfidPairing_label.pack(side="right", padx=(10,0))
        self.gateOut_button_rfidPairing = customtkinter.CTkButton(self.rfidPairing_entryButton_rightFrame, fg_color="dark red", text="Gate Out", hover_color="red", width=30, command=self.gate_out)
        self.gateOut_button_rfidPairing.pack(side="right", padx=(10,0))
        # self.gateIn_button_rfidPairing = customtkinter.CTkButton(self.rfidPairing_entryButton_rightFrame, fg_color="dark blue", text="Gate In", hover_color="blue", width=30)
        # self.gateIn_button_rfidPairing.pack(side="right", padx=(10,0))
        self.entry_noTiket_rfidPairing = customtkinter.CTkEntry(self.rfidPairing_entryButton_rightFrame, placeholder_text="", width=60)
        self.entry_noTiket_rfidPairing.pack(side="right", padx=(10,0))
        self.entry_noTiket_rfidPairing_label = customtkinter.CTkLabel(self.rfidPairing_entryButton_rightFrame, text="No Tiket:", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_noTiket_rfidPairing_label.pack(side="right", padx=(10,0))
        # self.entry_rfid_rfidPairing = customtkinter.CTkEntry(self.rfidPairing_entryButton_rightFrame, placeholder_text="", width=60)
        # self.entry_rfid_rfidPairing.pack(side="right", padx=(10,0))
        # self.entry_rfid_rfidPairing_label = customtkinter.CTkLabel(self.rfidPairing_entryButton_rightFrame, text="RFID:", font=customtkinter.CTkFont(size=13),anchor="w")
        # self.entry_rfid_rfidPairing_label.pack(side="right", padx=(10,0))
        # Queuing Manual Call (title manual call waiting list:)
        self.manualCall_title_frame = customtkinter.CTkFrame(self.rfidManual_frame, corner_radius=0, fg_color="transparent")
        self.manualCall_title_frame.grid(row=4, column=0, sticky="nsew")
        self.manualCall_label = customtkinter.CTkLabel(self.manualCall_title_frame, text="Manual Call Waiting List :", font=customtkinter.CTkFont(size=13,weight="bold"),anchor="w")
        self.manualCall_label.pack(side="left")
        # Queuing Manual Call (frame untuk table manual call dan tombol2 tap in out)
        self.manualCall_tableButton_frame = customtkinter.CTkFrame(self.rfidManual_frame, corner_radius=0, fg_color="transparent")
        self.manualCall_tableButton_frame.grid(row=5, column=0, sticky="nsew")
        self.manualCall_tableButton_frame.grid_columnconfigure(0, weight=1)
        self.manualCall_tableButton_frame.grid_rowconfigure(0, weight=1)
        self.manualCall_tableButton_leftframe = customtkinter.CTkFrame(self.manualCall_tableButton_frame, corner_radius=10, fg_color="transparent")
        self.manualCall_tableButton_leftframe.grid(row=0, column=0, sticky="nsew", padx=(0,5))
        self.manualCall_tableButton_leftframe.grid_columnconfigure(0, weight=1)
        self.manualCall_tableButton_leftframe.grid_rowconfigure(0, weight=1)
        self.manualCall_tableFrame_leftframe = customtkinter.CTkFrame(self.manualCall_tableButton_leftframe, corner_radius=0, fg_color="gray80")
        self.manualCall_tableFrame_leftframe.grid(row=0, column=0, sticky="nsew")
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.manualCall_table = ttk.Treeview(self.manualCall_tableFrame_leftframe)
        self.manualCall_table.pack(side="left", expand=tk.YES, fill=tk.BOTH)
        self.manualCall_table["columns"] = ("Ticket", "Police No", "Activity", "Arrival", "Packaging", "Product", "Skip", "Next Call", "Next ID", "oid")
        self.manualCall_table.column("#0", width=0,stretch=tk.NO)
        self.manualCall_table.column("Ticket", anchor=tk.CENTER, width=20,minwidth=20,stretch=tk.YES)
        self.manualCall_table.column("Police No", anchor=tk.CENTER, width=20, minwidth=0,stretch=tk.YES)
        self.manualCall_table.column("Activity", anchor=tk.CENTER, width=100, minwidth=0,stretch=tk.YES)
        self.manualCall_table.column("Arrival", anchor=tk.CENTER, width=20,minwidth=30,stretch=tk.YES)
        self.manualCall_table.column("Packaging", anchor=tk.CENTER, width=20, minwidth=0,stretch=tk.YES)
        self.manualCall_table.column("Product", anchor=tk.CENTER, width=20, minwidth=0,stretch=tk.YES)
        self.manualCall_table.column("Skip", anchor=tk.CENTER, width=0,minwidth=0,stretch=tk.NO)
        self.manualCall_table.column("Next Call", anchor=tk.CENTER, width=0, minwidth=0,stretch=tk.NO)
        self.manualCall_table.column("Next ID", anchor=tk.CENTER, width=0, minwidth=0,stretch=tk.NO)
        self.manualCall_table.column("oid", width=0,stretch=tk.NO)
        self.manualCall_table.heading("#0", text="", anchor=tk.W)
        self.manualCall_table.heading("Ticket", text="Ticket", anchor=tk.CENTER)
        self.manualCall_table.heading("Police No", text="Police No", anchor=tk.CENTER)
        self.manualCall_table.heading("Activity", text="Activity", anchor=tk.CENTER)
        self.manualCall_table.heading("Arrival", text="Arrival", anchor=tk.CENTER)
        self.manualCall_table.heading("Packaging", text="Packaging", anchor=tk.CENTER)
        self.manualCall_table.heading("Product", text="Product", anchor=tk.CENTER)
        self.manualCall_table.heading("Skip", text="Skip", anchor=tk.CENTER)
        self.manualCall_table.heading("Next Call", text="Next Call", anchor=tk.CENTER)
        self.manualCall_table.heading("Next ID", text="Next ID", anchor=tk.CENTER)
        self.manualCall_table.heading("oid", text="", anchor=tk.W)
        self.manualCall_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.manualCall_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.manualCall_table_scrollbar = customtkinter.CTkScrollbar(self.manualCall_tableFrame_leftframe, hover=True, button_hover_color="dark blue", command=self.manualCall_table.yview)
        self.manualCall_table_scrollbar.pack(side="left",fill=tk.Y)
        self.manualCall_table.configure(yscrollcommand=self.manualCall_table_scrollbar.set)
        self.manualCall_table.bind("<<TreeviewSelect>>", self.on_tree_manualCall_select)
        self.manualCall_buttonFrame_leftframe = customtkinter.CTkFrame(self.manualCall_tableButton_leftframe, corner_radius=10, fg_color="light blue")
        self.manualCall_buttonFrame_leftframe.grid(row=1, column=0, sticky="nsew")
        self.entry_truckNo_manualCall_label = customtkinter.CTkLabel(self.manualCall_buttonFrame_leftframe, text="Police No:", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_truckNo_manualCall_label.pack(side="left", padx=10)
        self.entry_truckNo_manualCall = customtkinter.CTkEntry(self.manualCall_buttonFrame_leftframe, placeholder_text="", width=100)
        self.entry_truckNo_manualCall.pack(side="left", padx=(0,10), pady=10)
        self.button_getTicket_manualCall = customtkinter.CTkButton(self.manualCall_buttonFrame_leftframe, fg_color="dark blue", text="PRINT TICKET", hover_color="blue", width=30, command=self.print_ticket)
        self.button_getTicket_manualCall.pack(side="left", padx=(0,10))
        self.button_refresh_manualCall = customtkinter.CTkButton(self.manualCall_buttonFrame_leftframe, fg_color="dark blue", text="REFRESH", hover_color="blue", width=30, command=self.refresh_waiting_list)
        self.button_refresh_manualCall.pack(side="left", padx=(0,10))
        self.button_delete_manualCall = customtkinter.CTkButton(self.manualCall_buttonFrame_leftframe, fg_color="dark blue", text="DELETE", hover_color="blue", width=30, command=self.delete_waiting_list)
        self.button_delete_manualCall.pack(side="left", padx=(0,10))
        self.button_call_manualCall = customtkinter.CTkButton(self.manualCall_buttonFrame_leftframe, fg_color="dark blue", text="CALL", hover_color="blue", width=50, command=self.call_waiting_list)
        self.button_call_manualCall.pack(side="left", padx=(0,10))
        self.button_skip_manualCall = customtkinter.CTkButton(self.manualCall_buttonFrame_leftframe, fg_color="dark blue", text="SKIP", hover_color="blue", width=50, command=self.skip_waiting_list)
        self.button_skip_manualCall.pack(side="left", padx=(0,10))
        self.button_gatein_manualCall = customtkinter.CTkButton(self.manualCall_buttonFrame_leftframe, fg_color="dark green", text="GATE IN", hover_color="green", width=50, command=self.gatein_waiting_list)
        self.button_gatein_manualCall.pack(side="left", padx=(0,10))
        self.button_cancel_manualCall = customtkinter.CTkButton(self.manualCall_buttonFrame_leftframe, fg_color="dark red", text="Cancel", hover_color="red", width=50, command=self.cancel_waiting_list)
        self.button_cancel_manualCall.pack(side="left", padx=(0,10))
        self.manualCall_tableButton_rightframe = customtkinter.CTkFrame(self.manualCall_tableButton_frame, corner_radius=10, fg_color="white")
        self.manualCall_tableButton_rightframe.grid(row=0, column=1, sticky="nsew")
        self.manualCall_tableButton_rightframe.grid_columnconfigure(0, weight=1)
        self.manualCall_tableButton_rightframe.grid_rowconfigure(2, weight=1)
        self.manualTaping_label = customtkinter.CTkLabel(self.manualCall_tableButton_rightframe, text="Manual Taping :", font=customtkinter.CTkFont(size=13,weight="bold"),anchor="w")
        self.manualTaping_label.grid(row=0, column=0, sticky="nsew", padx=(10,0))
        self.optionmenu_manualTaping= customtkinter.CTkOptionMenu(self.manualCall_tableButton_rightframe, values=["Weight Bridge 1","Warehouse","Weight Bridge 2","Cargo Covering"],fg_color="gray20",text_color="white",button_color="gray5",button_hover_color="#287ceb")
        self.optionmenu_manualTaping.grid(row=1, column=0, sticky="nsew", padx=10)
        self.button_manualTaping = customtkinter.CTkButton(self.manualCall_tableButton_rightframe, fg_color="gray0", text="✅\nSCAN MANUAL BLE", hover_color="gray", command=self.tap_manual)
        self.button_manualTaping.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)

        # create truck data frame #######################################################################
        self.truck_data_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.truck_data_frame.grid_columnconfigure(0, weight=1)
        self.truck_data_frame.grid_rowconfigure(0, weight=1)
        # create tabview truck data
        self.tabview_truckData = customtkinter.CTkTabview(self.truck_data_frame,fg_color=("gray92"),text_color="White",segmented_button_fg_color="Dark Blue",segmented_button_unselected_color="Dark Blue")
        self.tabview_truckData.grid(row=0, column=0, padx=(0, 0), pady=(0, 0), sticky="nsew")
        self.tabview_truckData.add("Vetting")
        self.tabview_truckData.add("Unloading")
        self.tabview_truckData.add("Distribution")
        #>> configure grid of individual tabs (Vetting) and item
        self.tabview_truckData.tab("Vetting").grid_columnconfigure(0, weight=1)
        self.tabview_truckData.tab("Vetting").grid_rowconfigure(1, weight=1)
        # Search vetting------------------------------------------------------------------------------------------------------------------------------
        self.search_vetting_frame = customtkinter.CTkFrame(self.tabview_truckData.tab("Vetting"), corner_radius=0, fg_color="transparent")
        self.search_vetting_frame.grid(row=0, pady=(0,10))
        self.search_vetting_label = customtkinter.CTkLabel(self.search_vetting_frame, text="Search :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.search_vetting_label.pack(side="left", padx=10, pady=10)
        self.entry_search_vetting = customtkinter.CTkEntry(self.search_vetting_frame, placeholder_text="")
        self.entry_search_vetting.pack(side="left")
        self.search_button_vetting = customtkinter.CTkButton(self.search_vetting_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search_vetting)
        self.search_button_vetting.pack(side="left")
        # Table vetting-------------------------------------------------------------------------------------------------------------------------------
        self.table_vetting_frame = customtkinter.CTkFrame(self.tabview_truckData.tab("Vetting"), corner_radius=0, fg_color="dark blue")
        self.table_vetting_frame.grid(row=1, sticky="nsew")
        self.table_vetting_frame.grid_columnconfigure(0, weight=1)
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.vetting_table = ttk.Treeview(self.table_vetting_frame)
        self.vetting_table.pack(side="left", expand=tk.YES, fill=tk.BOTH)
        self.vetting_table["columns"] = ("No", "Urutan", "Company", "Police No", "Reg Date", "Expired", "Remark", "Status", "RFID 1", "RFID 2", "Category")
        self.vetting_table.column("#0", width=0,stretch=tk.NO)
        self.vetting_table.column("No", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.vetting_table.column("Urutan", anchor=tk.CENTER, width=50, minwidth=0,stretch=tk.YES)
        self.vetting_table.column("Company", anchor=tk.CENTER, width=250, minwidth=0,stretch=tk.YES)
        self.vetting_table.column("Police No", anchor=tk.CENTER, width=120, minwidth=0,stretch=tk.YES)
        self.vetting_table.column("Reg Date", anchor=tk.CENTER, width=120, minwidth=0,stretch=tk.YES)
        self.vetting_table.column("Expired", anchor=tk.CENTER, width=120, minwidth=0,stretch=tk.YES)
        self.vetting_table.column("Remark", anchor=tk.CENTER, width=200, minwidth=0,stretch=tk.YES)
        self.vetting_table.column("Status", anchor=tk.CENTER, width=100, minwidth=0,stretch=tk.YES)
        self.vetting_table.column("RFID 1", anchor=tk.CENTER, width=160, minwidth=0,stretch=tk.YES)
        self.vetting_table.column("RFID 2", anchor=tk.CENTER, width=160, minwidth=0,stretch=tk.YES)
        self.vetting_table.column("Category", anchor=tk.CENTER, width=120, minwidth=0,stretch=tk.YES)
        self.vetting_table.heading("#0", text="", anchor=tk.W)
        self.vetting_table.heading("No", text="No", anchor=tk.W)
        self.vetting_table.heading("Urutan", text="No", anchor=tk.CENTER)
        self.vetting_table.heading("Company", text="Company", anchor=tk.CENTER)
        self.vetting_table.heading("Police No", text="Police No", anchor=tk.CENTER)
        self.vetting_table.heading("Reg Date", text="Reg Date", anchor=tk.CENTER)
        self.vetting_table.heading("Expired", text="Expired", anchor=tk.CENTER)
        self.vetting_table.heading("Remark", text="Remark", anchor=tk.CENTER)
        self.vetting_table.heading("Status", text="Status", anchor=tk.CENTER)
        self.vetting_table.heading("RFID 1", text="BLE Name 1", anchor=tk.CENTER)
        self.vetting_table.heading("RFID 2", text="BLE Name 2", anchor=tk.CENTER)
        self.vetting_table.heading("Category", text="Category", anchor=tk.CENTER)
        # Create striped row tags
        self.vetting_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.vetting_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        # create vetting table scrollbar
        self.vetting_table_scrollbar = customtkinter.CTkScrollbar(self.table_vetting_frame, hover=True, button_hover_color="yellow", command=self.vetting_table.yview)
        self.vetting_table_scrollbar.pack(side="left",fill=tk.Y)
        # connect self.vetting_table scroll event to self.vetting_table_scrollbar
        self.vetting_table.configure(yscrollcommand=self.vetting_table_scrollbar.set)
        # read and get data when selected in vetting table
        self.vetting_table.bind("<<TreeviewSelect>>", self.on_tree_vetting_select)
        # Entry Vetting -------------------------------------------------------------------------------------------------------------------------------
        self.entry_vetting_frame = customtkinter.CTkFrame(self.tabview_truckData.tab("Vetting"), corner_radius=10, fg_color="#5fa0f5")
        self.entry_vetting_frame.grid(row=2, pady=(10,0), sticky="nesw")
        self.entry_company_label = customtkinter.CTkLabel(self.entry_vetting_frame, text="Company :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_company_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_company = customtkinter.CTkEntry(self.entry_vetting_frame, placeholder_text="", width=200)
        self.entry_company.grid(row=0,column=1, padx=(0,10), pady= 5, columnspan=3, sticky="ew")
        self.entry_policeNo_label = customtkinter.CTkLabel(self.entry_vetting_frame, text="Police No:", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_policeNo_label.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_policeNo = customtkinter.CTkEntry(self.entry_vetting_frame, placeholder_text="", width=150)
        self.entry_policeNo.grid(row=1,column=1, padx=(0,10), pady= 5, columnspan=3, sticky="ew")
        self.entry_remark_label = customtkinter.CTkLabel(self.entry_vetting_frame, text="Remark :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_remark_label.grid(row=0,column=4, padx=10, pady= 5, sticky="ew")
        self.entry_remark = customtkinter.CTkEntry(self.entry_vetting_frame, placeholder_text="", width=150)
        self.entry_remark.grid(row=0,column=5, padx=(0,10), pady= 5, columnspan=3, sticky="ew")
        self.optionmenu_category_label = customtkinter.CTkLabel(self.entry_vetting_frame, text="Category :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.optionmenu_category_label.grid(row=1,column=4, padx=10, pady= 5, sticky="ew")
        self.optionmenu_category= customtkinter.CTkOptionMenu(self.entry_vetting_frame, values=["AKR","Non AKR"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb",command=self.optionmenuCategory_callback)
        self.optionmenu_category.grid(row=1,column=5, padx=(0,10), pady= 5, columnspan=3, sticky="ew")
        self.optionmenu_category.set("AKR")
        self.regdate_label = customtkinter.CTkLabel(self.entry_vetting_frame, text="Reg Date :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.regdate_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.regdate_calendar= MyDateEntry(self.entry_vetting_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.regdate_calendar.grid(row=2,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.expireddate_label = customtkinter.CTkLabel(self.entry_vetting_frame, text="Expired :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.expireddate_label.grid(row=2,column=2, padx=(0,10), pady= 5, sticky="ew")
        self.expireddate_calendar= MyDateEntry(self.entry_vetting_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.expireddate_calendar.grid(row=2,column=3, padx=(0,10), pady= 5, sticky="ew")
        self.entry_rfid1_label = customtkinter.CTkLabel(self.entry_vetting_frame, text="BLE Name 1:", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_rfid1_label.grid(row=2,column=4, padx=10, pady= 5, sticky="ew")
        self.entry_rfid1 = customtkinter.CTkEntry(self.entry_vetting_frame, placeholder_text="", width=80)
        self.entry_rfid1.grid(row=2,column=5, padx=(0,10), pady= 5, sticky="ew")
        self.entry_rfid2_label = customtkinter.CTkLabel(self.entry_vetting_frame, text="BLE Name 2:", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_rfid2_label.grid(row=2,column=6, padx=(0,10), pady= 5, sticky="ew")
        self.entry_rfid2 = customtkinter.CTkEntry(self.entry_vetting_frame, placeholder_text="", width=80)
        self.entry_rfid2.grid(row=2,column=7, padx=(0,10), pady= 5, sticky="ew")
        self.optionmenu_status_label = customtkinter.CTkLabel(self.entry_vetting_frame, text="Status :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.optionmenu_status_label.grid(row=3,column=0, padx=10, pady= 5, sticky="ew")
        self.optionmenu_status= customtkinter.CTkOptionMenu(self.entry_vetting_frame, values=["Activate","Not Activate"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb",command=self.optionmenuStatus_callback)
        self.optionmenu_status.grid(row=3,column=1, padx=(0,10), pady= 5, columnspan=3, sticky="ew")
        self.optionmenu_status.set("Activate")
        self.cancel_button_vetting = customtkinter.CTkButton(self.entry_vetting_frame, fg_color="dark red", text="Cancel", hover_color="red", width=60, command=self.cancel_modify_vetting)
        self.cancel_button_vetting.grid(row=3,column=7, rowspan=2)
        # Button Vetting -------------------------------------------------------------------------------------------------------------------------------
        self.button_vetting_frame = customtkinter.CTkFrame(self.tabview_truckData.tab("Vetting"), corner_radius=0, fg_color="transparent")
        self.button_vetting_frame.grid(row=3, pady=10)
        self.add_button_vetting = customtkinter.CTkButton(self.button_vetting_frame, fg_color="dark blue", text="ADD", hover_color="blue", command=self.add_vetting)
        self.add_button_vetting.grid(row=0,column=0)
        self.remove_button_vetting = customtkinter.CTkButton(self.button_vetting_frame, fg_color="dark blue", text="Remove", hover_color="blue", command=self.remove_vetting)
        self.remove_button_vetting.grid(row=0,column=1, padx=20, pady= 10)
        self.replace_button_vetting = customtkinter.CTkButton(self.button_vetting_frame, fg_color="dark blue", text="Replace", hover_color="blue", command=self.replace_vetting)
        self.replace_button_vetting.grid(row=0,column=2)
        self.upload_excel_button_vetting = customtkinter.CTkButton(self.button_vetting_frame, fg_color="dark green", text="Upload Excel", hover_color="green", command=self.upload_excel_vetting)
        self.upload_excel_button_vetting.grid(row=0,column=3, padx=20, pady= 10)
        #>> configure grid of individual tabs (Unloading) and item
        self.tabview_truckData.tab("Unloading").grid_columnconfigure(0, weight=1)
        self.tabview_truckData.tab("Unloading").grid_rowconfigure(1, weight=1)
        # Search unloading------------------------------------------------------------------------------------------------------------------------------
        self.search_unloading_frame = customtkinter.CTkFrame(self.tabview_truckData.tab("Unloading"), corner_radius=0, fg_color="transparent")
        self.search_unloading_frame.grid(row=0, pady=(0,10))
        self.search_unloading_label = customtkinter.CTkLabel(self.search_unloading_frame, text="Search :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.search_unloading_label.pack(side="left", padx=10, pady=10)
        self.entry_search_unloading = customtkinter.CTkEntry(self.search_unloading_frame, placeholder_text="")
        self.entry_search_unloading.pack(side="left")
        self.search_button_unloading = customtkinter.CTkButton(self.search_unloading_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search_unloading)
        self.search_button_unloading.pack(side="left")
        # Table unloading-------------------------------------------------------------------------------------------------------------------------------
        self.table_unloading_frame = customtkinter.CTkFrame(self.tabview_truckData.tab("Unloading"), corner_radius=0, fg_color="dark blue")
        self.table_unloading_frame.grid(row=1, sticky="nsew")
        self.table_unloading_frame.grid_columnconfigure(0, weight=1)
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.unloading_table = ttk.Treeview(self.table_unloading_frame)
        self.unloading_table.pack(side="left", expand=tk.YES, fill=tk.BOTH)
        self.unloading_table["columns"] = ("No", "Urutan", "Company", "Police No", "Packaging", "Product", "Date", "Time", "Remark", "Status", "Quota", "Berat", "WH ID", "No SO")
        self.unloading_table.column("#0", width=0,stretch=tk.NO)
        self.unloading_table.column("No", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.unloading_table.column("Urutan", anchor=tk.CENTER, width=50, minwidth=0,stretch=tk.YES)
        self.unloading_table.column("Company", anchor=tk.CENTER, width=250, minwidth=0,stretch=tk.YES)
        self.unloading_table.column("Police No", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.unloading_table.column("Packaging", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.unloading_table.column("Product", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.unloading_table.column("Date", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.unloading_table.column("Time", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.unloading_table.column("Remark", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.unloading_table.column("Status", anchor=tk.CENTER, width=100, minwidth=0,stretch=tk.YES)
        self.unloading_table.column("Quota", anchor=tk.CENTER, width=70, minwidth=0,stretch=tk.YES)
        self.unloading_table.column("Berat", anchor=tk.CENTER, width=70, minwidth=0,stretch=tk.YES)
        self.unloading_table.column("WH ID", anchor=tk.CENTER, width=70, minwidth=0,stretch=tk.YES)
        self.unloading_table.column("No SO", anchor=tk.CENTER, width=70, minwidth=0,stretch=tk.YES)
        self.unloading_table.heading("#0", text="", anchor=tk.W)
        self.unloading_table.heading("No", text="", anchor=tk.W)
        self.unloading_table.heading("Urutan", text="No", anchor=tk.CENTER)
        self.unloading_table.heading("Company", text="Company", anchor=tk.CENTER)
        self.unloading_table.heading("Police No", text="Police No", anchor=tk.CENTER)
        self.unloading_table.heading("Packaging", text="Packaging", anchor=tk.CENTER)
        self.unloading_table.heading("Product", text="Product", anchor=tk.CENTER)
        self.unloading_table.heading("Date", text="Date", anchor=tk.CENTER)
        self.unloading_table.heading("Time", text="Time", anchor=tk.CENTER)
        self.unloading_table.heading("Remark", text="Remark", anchor=tk.CENTER)
        self.unloading_table.heading("Status", text="Status", anchor=tk.CENTER)
        self.unloading_table.heading("Quota", text="Quota", anchor=tk.CENTER)
        self.unloading_table.heading("Berat", text="Berat", anchor=tk.CENTER)
        self.unloading_table.heading("WH ID", text="WH ID", anchor=tk.CENTER)
        self.unloading_table.heading("No SO", text="No SO", anchor=tk.CENTER)
        # Create striped row tags
        self.unloading_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.unloading_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        # create unloading table scrollbar
        self.unloading_table_scrollbar = customtkinter.CTkScrollbar(self.table_unloading_frame, hover=True, button_hover_color="yellow", command=self.unloading_table.yview)
        self.unloading_table_scrollbar.pack(side="left",fill=tk.Y)
        # connect self.unloading_table scroll event to self.unloading_table_scrollbar
        self.unloading_table.configure(yscrollcommand=self.unloading_table_scrollbar.set)
        # read and get data when selected in unloading table
        self.unloading_table.bind("<<TreeviewSelect>>", self.on_tree_unloading_select)
        # Entry unloading -------------------------------------------------------------------------------------------------------------------------------
        self.entry_unloading_frame = customtkinter.CTkFrame(self.tabview_truckData.tab("Unloading"), corner_radius=10, fg_color="#5fa0f5")
        self.entry_unloading_frame.grid(row=2, pady=(10,0), sticky="nesw")
        self.entry_company_unloading_label = customtkinter.CTkLabel(self.entry_unloading_frame, text="Company :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_company_unloading_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_company_unloading = customtkinter.CTkEntry(self.entry_unloading_frame, placeholder_text="", width=200)
        self.entry_company_unloading.grid(row=0,column=1, padx=(0,10), pady= 5, sticky="ew", columnspan=5)
        self.entry_policeNo_unloading_label = customtkinter.CTkLabel(self.entry_unloading_frame, text="Police No:", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_policeNo_unloading_label.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_policeNo_unloading = customtkinter.CTkEntry(self.entry_unloading_frame, placeholder_text="", width=200)
        self.entry_policeNo_unloading.grid(row=1,column=1, padx=(0,10), pady= 5, sticky="ew", columnspan=5)
        self.optionmenu_packaging_label = customtkinter.CTkLabel(self.entry_unloading_frame, text="Packaging :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.optionmenu_packaging_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.optionmenu_packaging= customtkinter.CTkOptionMenu(self.entry_unloading_frame, values=[""],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", dynamic_resizing= False, width=188, command=self.optionmenuPackaging_callback)
        self.optionmenu_packaging.grid(row=2,column=1, padx=(0,10), pady= 5, sticky="ew", columnspan=5)
        self.optionmenu_packaging.set("Pilih Packaging")
        self.optionmenu_product_label = customtkinter.CTkLabel(self.entry_unloading_frame, text="Product :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.optionmenu_product_label.grid(row=3,column=0, padx=10, pady= 5, sticky="ew")
        self.optionmenu_product= customtkinter.CTkOptionMenu(self.entry_unloading_frame, values=[""],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", dynamic_resizing= False, width=205, command=self.optionmenuProduct_callback)
        self.optionmenu_product.grid(row=3,column=1, padx=(0,10), pady= 5, sticky="ew", columnspan=5)
        self.optionmenu_product.set("Pilih Product")
        self.date_unloading_label = customtkinter.CTkLabel(self.entry_unloading_frame, text="Date :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.date_unloading_label.grid(row=0,column=6, padx=10, pady= 5, sticky="ew")
        self.date_unloading_calendar= MyDateEntry(self.entry_unloading_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.date_unloading_calendar.grid(row=0,column=7, padx=(0,10), pady= 5, sticky="ew")
        self.date_unloading_label = customtkinter.CTkLabel(self.entry_unloading_frame, text="Time :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="e")
        self.date_unloading_label.grid(row=0,column=8, padx=(0,10), pady= 5, sticky="ew")
        self.spbox_hour = tk.Spinbox(self.entry_unloading_frame,from_=00,to=23,bg="white",justify=tk.CENTER,width=3,format="%02.0f",font=23)
        self.spbox_hour.grid(row=0,column=9, padx=(0,10), pady= 5, sticky="ew")
        self.date_unloading_label_pemisah = customtkinter.CTkLabel(self.entry_unloading_frame, text=":", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.date_unloading_label_pemisah.grid(row=0,column=10, padx=(0,10), pady= 5, sticky="ew")
        self.spbox_minute = tk.Spinbox(self.entry_unloading_frame,from_=00,to=59,bg="white",justify=tk.CENTER,width=3,format="%02.0f",font=23)
        self.spbox_minute.grid(row=0,column=11, padx=(0,10), pady= 5, sticky="ew")
        self.entry_quota_unloading_label = customtkinter.CTkLabel(self.entry_unloading_frame, text="Quota :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_quota_unloading_label.grid(row=1,column=6, padx=10, pady= 5, sticky="ew")
        self.entry_quota_unloading = customtkinter.CTkEntry(self.entry_unloading_frame, placeholder_text="", width=80)
        self.entry_quota_unloading.insert(0,"0")
        self.entry_quota_unloading.grid(row=1,column=7, padx=(0,10), pady= 5, sticky="ew")
        self.entry_berat_unloading_label = customtkinter.CTkLabel(self.entry_unloading_frame, text="Berat (KG):", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_berat_unloading_label.grid(row=1,column=8, padx=(0,10), pady= 5, sticky="ew")
        self.entry_berat_unloading = customtkinter.CTkEntry(self.entry_unloading_frame, placeholder_text="", width=67)
        self.entry_berat_unloading.grid(row=1,column=9, columnspan=3, padx=(0,10), pady= 5, sticky="ew")
        self.entry_remark_unloading_label = customtkinter.CTkLabel(self.entry_unloading_frame, text="Remark :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_remark_unloading_label.grid(row=2,column=6, padx=10, pady= 5, sticky="ew")
        self.entry_remark_unloading = customtkinter.CTkEntry(self.entry_unloading_frame, placeholder_text="", width=238)
        self.entry_remark_unloading.insert(0,"UNLOADING")
        self.entry_remark_unloading.configure(state="disable")
        self.entry_remark_unloading.grid(row=2,column=7, padx=(0,10), pady= 5, sticky="ew", columnspan=5)
        self.optionmenu_status_unloading_label = customtkinter.CTkLabel(self.entry_unloading_frame, text="Status :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.optionmenu_status_unloading_label.grid(row=3,column=6, padx=10, pady= 5, sticky="ew")
        self.optionmenu_unloading_status= customtkinter.CTkOptionMenu(self.entry_unloading_frame, values=["Activate","Not Activate"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb",dynamic_resizing= False, width=240,command=self.optionmenuStatus_callback)
        self.optionmenu_unloading_status.grid(row=3,column=7, padx=(0,10), pady= 5, sticky="ew", columnspan=5)
        self.optionmenu_unloading_status.set("Activate")
        self.optionmenu_warehouse_label = customtkinter.CTkLabel(self.entry_unloading_frame, text="Assign to warehouse :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.optionmenu_warehouse_label.grid(row=0,column=12, padx=10, pady= 5, sticky="ew")
        self.optionmenu_warehouse= customtkinter.CTkOptionMenu(self.entry_unloading_frame, values=["Otomatis"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb",dynamic_resizing= False, command=self.optionmenuStatus_callback)
        self.optionmenu_warehouse.grid(row=0,column=13, padx=10, pady= 5, sticky="ew")
        self.optionmenu_noSo_label = customtkinter.CTkLabel(self.entry_unloading_frame, text="No SO :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.optionmenu_noSo_label.grid(row=1,column=12, padx=10, pady= 5, sticky="ew")
        self.optionmenu_noSo= customtkinter.CTkEntry(self.entry_unloading_frame, placeholder_text="")
        self.optionmenu_noSo.grid(row=1,column=13, padx=10, pady= 5, sticky="ew")
        self.cancel_button_unloading = customtkinter.CTkButton(self.entry_unloading_frame, fg_color="dark red", text="Cancel", hover_color="red", width=60, command=self.cancel_modify_unloading)
        self.cancel_button_unloading.grid(row=2,column=13, rowspan=2)
        # Button unloading -------------------------------------------------------------------------------------------------------------------------------
        self.button_unloading_frame = customtkinter.CTkFrame(self.tabview_truckData.tab("Unloading"), corner_radius=0, fg_color="transparent")
        self.button_unloading_frame.grid(row=3, pady=10)
        self.add_button_unloading = customtkinter.CTkButton(self.button_unloading_frame, fg_color="dark blue", text="ADD", hover_color="blue", command=self.add_unloading)
        self.add_button_unloading.grid(row=0,column=0)
        self.remove_button_unloading = customtkinter.CTkButton(self.button_unloading_frame, fg_color="dark blue", text="Remove", hover_color="blue", command=self.remove_unloading)
        self.remove_button_unloading.grid(row=0,column=1, padx=20, pady= 10)
        self.replace_button_unloading = customtkinter.CTkButton(self.button_unloading_frame, fg_color="dark blue", text="Replace", hover_color="blue", command=self.replace_unloading)
        self.replace_button_unloading.grid(row=0,column=2)
        self.clear_button_unloading = customtkinter.CTkButton(self.button_unloading_frame, fg_color="dark blue", text="Clear", hover_color="blue", command=self.clear_unloading)
        self.clear_button_unloading.grid(row=0,column=3, padx=20, pady= 10)
        self.upload_excel_button_unloading = customtkinter.CTkButton(self.button_unloading_frame, fg_color="dark green", text="Upload Excel", hover_color="green", command=self.upload_excel_unloading)
        self.upload_excel_button_unloading.grid(row=0,column=4)
        #>> configure grid of individual tabs (Distribution) and item
        self.tabview_truckData.tab("Distribution").grid_columnconfigure(0, weight=1)
        self.tabview_truckData.tab("Distribution").grid_rowconfigure(1, weight=1)
        # Search unloading------------------------------------------------------------------------------------------------------------------------------
        self.search_distribution_frame = customtkinter.CTkFrame(self.tabview_truckData.tab("Distribution"), corner_radius=0, fg_color="transparent")
        self.search_distribution_frame.grid(row=0)
        self.search_distribution_label = customtkinter.CTkLabel(self.search_distribution_frame, text="Search :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.search_distribution_label.pack(side="left", padx=10, pady=10)
        self.entry_search_distribution = customtkinter.CTkEntry(self.search_distribution_frame, placeholder_text="")
        self.entry_search_distribution.pack(side="left")
        self.search_button_distribution = customtkinter.CTkButton(self.search_distribution_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search_distribution_loadingList)
        self.search_button_distribution.pack(side="left")
        # create tabview distribution
        self.tabview_in_distribution = customtkinter.CTkTabview(self.tabview_truckData.tab("Distribution"),fg_color=("gray92"),text_color="White",segmented_button_fg_color="Dark Red",segmented_button_unselected_color="Dark Red", segmented_button_selected_color="Dark Blue",command=self.tabview_distribution_button)
        self.tabview_in_distribution.grid(row=1, sticky="nsew")
        self.tabview_in_distribution.add("Loading List")
        self.tabview_in_distribution.tab("Loading List").grid_columnconfigure(0, weight=1)
        self.tabview_in_distribution.tab("Loading List").grid_rowconfigure(0, weight=1)
        self.tabview_in_distribution.add("Displan Data")
        self.tabview_in_distribution.tab("Displan Data").grid_columnconfigure(0, weight=1)
        self.tabview_in_distribution.tab("Displan Data").grid_rowconfigure(0, weight=1)
        # Table loading list-------------------------------------------------------------------------------------------------------------------------------
        self.table_loadingList_frame = customtkinter.CTkFrame(self.tabview_in_distribution.tab("Loading List"), corner_radius=0, fg_color="dark blue")
        self.table_loadingList_frame.grid(row=0, sticky="nsew")
        self.table_loadingList_frame.grid_columnconfigure(0, weight=1)
        self.loadingList_table = ttk.Treeview(self.table_loadingList_frame)
        self.loadingList_table.pack(side="left", expand=tk.YES, fill=tk.BOTH)
        self.loadingList_table["columns"] = ("No", "Urutan", "Company", "Police No", "Packaging", "Product", "Date", "Time", "Remark", "Status", "Quota", "No SO", "Satuan", "Berat", "WH ID")
        self.loadingList_table.column("#0", width=0,stretch=tk.NO)
        self.loadingList_table.column("No", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.loadingList_table.column("Urutan", anchor=tk.CENTER, width=50, minwidth=0,stretch=tk.YES)
        self.loadingList_table.column("Company", anchor=tk.CENTER, width=170, minwidth=0,stretch=tk.YES)
        self.loadingList_table.column("Police No", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.loadingList_table.column("Packaging", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.loadingList_table.column("Product", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.loadingList_table.column("Date", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.loadingList_table.column("Time", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.loadingList_table.column("Remark", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.loadingList_table.column("Status", anchor=tk.CENTER, width=100, minwidth=0,stretch=tk.YES)
        self.loadingList_table.column("Quota", anchor=tk.CENTER, width=70, minwidth=0,stretch=tk.YES)
        self.loadingList_table.column("No SO", anchor=tk.CENTER, width=70, minwidth=0,stretch=tk.YES)
        self.loadingList_table.column("Satuan", anchor=tk.CENTER, width=70, minwidth=0,stretch=tk.YES)
        self.loadingList_table.column("Berat", anchor=tk.CENTER, width=70, minwidth=0,stretch=tk.YES)
        self.loadingList_table.column("WH ID", anchor=tk.CENTER, width=60, minwidth=0,stretch=tk.YES)
        self.loadingList_table.heading("#0", text="", anchor=tk.W)
        self.loadingList_table.heading("No", text="", anchor=tk.W)
        self.loadingList_table.heading("Urutan", text="No", anchor=tk.CENTER)
        self.loadingList_table.heading("Company", text="Company", anchor=tk.CENTER)
        self.loadingList_table.heading("Police No", text="Police No", anchor=tk.CENTER)
        self.loadingList_table.heading("Packaging", text="Packaging", anchor=tk.CENTER)
        self.loadingList_table.heading("Product", text="Product", anchor=tk.CENTER)
        self.loadingList_table.heading("Date", text="Date", anchor=tk.CENTER)
        self.loadingList_table.heading("Time", text="Time", anchor=tk.CENTER)
        self.loadingList_table.heading("Remark", text="Remark", anchor=tk.CENTER)
        self.loadingList_table.heading("Status", text="Status", anchor=tk.CENTER)
        self.loadingList_table.heading("Quota", text="Quota", anchor=tk.CENTER)
        self.loadingList_table.heading("No SO", text="No SO", anchor=tk.CENTER)
        self.loadingList_table.heading("Satuan", text="Satuan", anchor=tk.CENTER)
        self.loadingList_table.heading("Berat", text="Berat", anchor=tk.CENTER)
        self.loadingList_table.heading("WH ID", text="WH ID", anchor=tk.CENTER)
        # Create striped row tags
        self.loadingList_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.loadingList_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        # create loadingList table scrollbar
        self.loadingList_table_scrollbar = customtkinter.CTkScrollbar(self.table_loadingList_frame, hover=True, button_hover_color="yellow", command=self.loadingList_table.yview)
        self.loadingList_table_scrollbar.pack(side="left",fill=tk.Y)
        # connect self.loadingList_table scroll event to self.loadingList_table_scrollbar
        self.loadingList_table.configure(yscrollcommand=self.loadingList_table_scrollbar.set)
        # read and get data when selected in loadingList table
        self.loadingList_table.bind("<<TreeviewSelect>>", self.on_tree_loadingList_select)
        # Table displan data-------------------------------------------------------------------------------------------------------------------------------
        self.table_displanData_frame = customtkinter.CTkFrame(self.tabview_in_distribution.tab("Displan Data"), corner_radius=0, fg_color="dark blue")
        self.table_displanData_frame.grid(row=0, sticky="nsew")
        self.table_displanData_frame.grid_columnconfigure(0, weight=1)
        self.displanData_table = ttk.Treeview(self.table_displanData_frame)
        self.displanData_table.pack(side="left", expand=tk.YES, fill=tk.BOTH)
        self.displanData_table["columns"] = ("Company", "Truck", "Packaging", "Product", "Date", "Time", "Pkg remark", "Remark", "No SO", "TR Part", "Satuan", "Berat")
        self.displanData_table.column("#0", width=0,stretch=tk.NO)
        self.displanData_table.column("Company", anchor=tk.CENTER, width=250, minwidth=0,stretch=tk.YES)
        self.displanData_table.column("Truck", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.displanData_table.column("Packaging", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.displanData_table.column("Product", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.displanData_table.column("Date", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.displanData_table.column("Time", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.displanData_table.column("Pkg remark", anchor=tk.CENTER, width=150, minwidth=0,stretch=tk.YES)
        self.displanData_table.column("Remark", anchor=tk.CENTER, width=100, minwidth=0,stretch=tk.YES)
        self.displanData_table.column("No SO", anchor=tk.CENTER, width=70, minwidth=0,stretch=tk.YES)
        self.displanData_table.column("TR Part", anchor=tk.CENTER, width=70, minwidth=0,stretch=tk.YES)
        self.displanData_table.column("Satuan", anchor=tk.CENTER, width=70, minwidth=0,stretch=tk.YES)
        self.displanData_table.column("Berat", anchor=tk.CENTER, width=70, minwidth=0,stretch=tk.YES)
        self.displanData_table.heading("#0", text="", anchor=tk.W)
        self.displanData_table.heading("Company", text="Company", anchor=tk.CENTER)
        self.displanData_table.heading("Truck", text="Truck", anchor=tk.CENTER)
        self.displanData_table.heading("Packaging", text="Packaging", anchor=tk.CENTER)
        self.displanData_table.heading("Product", text="Product", anchor=tk.CENTER)
        self.displanData_table.heading("Date", text="Date", anchor=tk.CENTER)
        self.displanData_table.heading("Time", text="Time", anchor=tk.CENTER)
        self.displanData_table.heading("Pkg remark", text="Pkg remark", anchor=tk.CENTER)
        self.displanData_table.heading("Remark", text="Remark", anchor=tk.CENTER)
        self.displanData_table.heading("No SO", text="No SO", anchor=tk.CENTER)
        self.displanData_table.heading("TR Part", text="TR Part", anchor=tk.CENTER)
        self.displanData_table.heading("Satuan", text="Satuan", anchor=tk.CENTER)
        self.displanData_table.heading("Berat", text="Berat", anchor=tk.CENTER)
        # Create striped row tags
        self.displanData_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.displanData_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        # create displanData table scrollbar
        self.displanData_table_scrollbar = customtkinter.CTkScrollbar(self.table_displanData_frame, hover=True, button_hover_color="yellow", command=self.displanData_table.yview)
        self.displanData_table_scrollbar.pack(side="left",fill=tk.Y)
        # connect self.displanData_table scroll event to self.displanData_table_scrollbar
        self.displanData_table.configure(yscrollcommand=self.displanData_table_scrollbar.set)
        # read and get data when selected in displanData table
        self.displanData_table.bind("<<TreeviewSelect>>", self.on_tree_displanData_select)
        # Entry distribution --------------------------------------------------------------------------------------------------------------------------------
        self.entry_distribution_frame = customtkinter.CTkFrame(self.tabview_truckData.tab("Distribution"), corner_radius=0, fg_color="#5fa0f5")
        self.entry_distribution_frame.grid(row=2, pady=(10,0), sticky="nesw")
        self.entry_company_distribution_label = customtkinter.CTkLabel(self.entry_distribution_frame, text="Company :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_company_distribution_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_company_distribution = customtkinter.CTkEntry(self.entry_distribution_frame, placeholder_text="", width=200)
        self.entry_company_distribution.grid(row=0,column=1, padx=(0,10), pady= 5, sticky="ew", columnspan=5)
        self.entry_policeNo_distribution_label = customtkinter.CTkLabel(self.entry_distribution_frame, text="Police No:", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_policeNo_distribution_label.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_policeNo_distribution = customtkinter.CTkEntry(self.entry_distribution_frame, placeholder_text="", width=200)
        self.entry_policeNo_distribution.grid(row=1,column=1, padx=(0,10), pady= 5, sticky="ew", columnspan=5)
        self.optionmenu_packaging_distribution_label = customtkinter.CTkLabel(self.entry_distribution_frame, text="Packaging :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.optionmenu_packaging_distribution_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.optionmenu_packaging_distribution= customtkinter.CTkOptionMenu(self.entry_distribution_frame, values=[""],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", dynamic_resizing= False, width=188, command=self.optionmenuPackaging_callback)
        self.optionmenu_packaging_distribution.grid(row=2,column=1, padx=(0,10), pady= 5, sticky="ew", columnspan=5)
        self.optionmenu_packaging_distribution.set("Pilih Packaging")
        self.optionmenu_product_distribution_label = customtkinter.CTkLabel(self.entry_distribution_frame, text="Product :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.optionmenu_product_distribution_label.grid(row=3,column=0, padx=10, pady= 5, sticky="ew")
        self.optionmenu_product_distribution= customtkinter.CTkOptionMenu(self.entry_distribution_frame, values=[""],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb", dynamic_resizing= False, width=205, command=self.optionmenuProduct_callback)
        self.optionmenu_product_distribution.grid(row=3,column=1, padx=(0,10), pady= 5, sticky="ew", columnspan=5)
        self.optionmenu_product_distribution.set("Pilih Product")
        self.date_distribution_label = customtkinter.CTkLabel(self.entry_distribution_frame, text="Date :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.date_distribution_label.grid(row=0,column=6, padx=10, pady= 5, sticky="ew")
        self.date_distribution_calendar= MyDateEntry(self.entry_distribution_frame, width=10, year=tahun, month=bln, day=tgl, background='darkblue', foreground='white', borderwidth=10, font=(None, 12))
        self.date_distribution_calendar.grid(row=0,column=7, padx=(0,10), pady= 5, sticky="ew")
        self.date_distribution_label = customtkinter.CTkLabel(self.entry_distribution_frame, text="Time :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="e")
        self.date_distribution_label.grid(row=0,column=8, padx=(0,10), pady= 5, sticky="ew")
        self.spbox_hour_distribution = tk.Spinbox(self.entry_distribution_frame,from_=00,to=23,bg="white",justify=tk.CENTER,width=3,format="%02.0f",font=23)
        self.spbox_hour_distribution.grid(row=0,column=9, padx=(0,10), pady= 5, sticky="ew")
        self.date_distribution_label_pemisah = customtkinter.CTkLabel(self.entry_distribution_frame, text=":", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.date_distribution_label_pemisah.grid(row=0,column=10, padx=(0,10), pady= 5, sticky="ew")
        self.spbox_minute_distribution = tk.Spinbox(self.entry_distribution_frame,from_=00,to=59,bg="white",justify=tk.CENTER,width=3,format="%02.0f",font=23)
        self.spbox_minute_distribution.grid(row=0,column=11, padx=(0,10), pady= 5, sticky="ew")
        self.entry_quota_distribution_label = customtkinter.CTkLabel(self.entry_distribution_frame, text="Quota :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_quota_distribution_label.grid(row=1,column=6, padx=10, pady= 5, sticky="ew")
        self.entry_quota_distribution = customtkinter.CTkEntry(self.entry_distribution_frame, placeholder_text="", width=80)
        self.entry_quota_distribution.insert(0,"0")
        self.entry_quota_distribution.grid(row=1,column=7, padx=(0,10), pady= 5, sticky="ew")
        self.entry_berat_distribution_label = customtkinter.CTkLabel(self.entry_distribution_frame, text="Berat (KG):", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_berat_distribution_label.grid(row=1,column=8, padx=(0,10), pady= 5, sticky="ew")
        self.entry_berat_distribution = customtkinter.CTkEntry(self.entry_distribution_frame, placeholder_text="", width=67)
        self.entry_berat_distribution.grid(row=1,column=9, columnspan=3, padx=(0,10), pady= 5, sticky="ew")
        self.entry_remark_distribution_label = customtkinter.CTkLabel(self.entry_distribution_frame, text="Remark :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_remark_distribution_label.grid(row=2,column=6, padx=10, pady= 5, sticky="ew")
        self.entry_remark_distribution = customtkinter.CTkEntry(self.entry_distribution_frame, placeholder_text="", width=238)
        self.entry_remark_distribution.insert(0,"LOADING")
        self.entry_remark_distribution.configure(state="disable")
        self.entry_remark_distribution.grid(row=2,column=7, padx=(0,10), pady= 5, sticky="ew", columnspan=5)
        self.optionmenu_status_distribution_label = customtkinter.CTkLabel(self.entry_distribution_frame, text="Status :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.optionmenu_status_distribution_label.grid(row=3,column=6, padx=10, pady= 5, sticky="ew")
        self.optionmenu_distribution_status= customtkinter.CTkOptionMenu(self.entry_distribution_frame, values=["Activate","Not Activate"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb",dynamic_resizing= False, width=90,command=self.optionmenuStatus_callback)
        self.optionmenu_distribution_status.grid(row=3,column=7, padx=(0,10), pady= 5, sticky="ew")
        self.optionmenu_distribution_status.set("Activate")
        self.entry_noSO_distribution_label = customtkinter.CTkLabel(self.entry_distribution_frame, text="No SO :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="e")
        self.entry_noSO_distribution_label.grid(row=3,column=8, padx=(0,10), pady= 5, sticky="ew")
        self.entry_noSO_distribution = customtkinter.CTkEntry(self.entry_distribution_frame, placeholder_text="", width=80)
        self.entry_noSO_distribution.grid(row=3,column=9, columnspan=3, padx=(0,10), pady= 5, sticky="ew")
        self.optionmenu_warehouse_distribution_label = customtkinter.CTkLabel(self.entry_distribution_frame, text="Assign to warehouse :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.optionmenu_warehouse_distribution_label.grid(row=0,column=12, padx=10, pady= 5, sticky="ew")
        self.optionmenu_warehouse_distribution= customtkinter.CTkOptionMenu(self.entry_distribution_frame, values=["Otomatis"],fg_color="white",text_color="black",button_color="white",button_hover_color="#287ceb",command=self.optionmenuStatus_callback)
        self.optionmenu_warehouse_distribution.grid(row=0,column=13, padx=10, pady= 5, sticky="ew")
        self.entry_pkgRemark_distribution_label = customtkinter.CTkLabel(self.entry_distribution_frame, text="Pkg Remark :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_pkgRemark_distribution_label.grid(row=1,column=12, padx=10, pady= 5, sticky="ew")
        self.entry_pkgRemark_distribution = customtkinter.CTkEntry(self.entry_distribution_frame, placeholder_text="", width=200)
        self.entry_pkgRemark_distribution.grid(row=1,column=13, padx=10, pady= 5, sticky="ew")
        self.entry_pkgRemark_distribution.configure(state="disabled",fg_color="gray")
        self.entry_trPart_distribution_label = customtkinter.CTkLabel(self.entry_distribution_frame, text="TR Part :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_trPart_distribution_label.grid(row=2,column=12, padx=10, pady= 5, sticky="ew")
        self.entry_trPart_distribution = customtkinter.CTkEntry(self.entry_distribution_frame, placeholder_text="", width=200)
        self.entry_trPart_distribution.grid(row=2,column=13, padx=10, pady= 5, sticky="ew")
        self.entry_trPart_distribution.configure(state="disabled",fg_color="gray")
        self.cancel_button_distribution = customtkinter.CTkButton(self.entry_distribution_frame, fg_color="dark red", text="Cancel", hover_color="red", width=60, command=self.cancel_modify_distribution)
        self.cancel_button_distribution.grid(row=3,column=13, rowspan=2)
        # Button distribution -------------------------------------------------------------------------------------------------------------------------------
        self.button_distribution_frame = customtkinter.CTkFrame(self.tabview_truckData.tab("Distribution"), corner_radius=0, fg_color="transparent")
        self.button_distribution_frame.grid(row=3, pady=10)
        self.add_button_distribution = customtkinter.CTkButton(self.button_distribution_frame, fg_color="dark blue", text="ADD", hover_color="blue", command=self.add_distribution_loadingList)
        self.add_button_distribution.grid(row=0,column=0)
        self.remove_button_distribution = customtkinter.CTkButton(self.button_distribution_frame, fg_color="dark blue", text="Remove", hover_color="blue", command=self.remove_distribution)
        self.remove_button_distribution.grid(row=0,column=1, padx=20, pady= 10)
        self.replace_button_distribution = customtkinter.CTkButton(self.button_distribution_frame, fg_color="dark blue", text="Replace", hover_color="blue", command=self.replace_distribution)
        self.replace_button_distribution.grid(row=0,column=2)
        self.clear_button_distribution = customtkinter.CTkButton(self.button_distribution_frame, fg_color="dark blue", text="Clear", hover_color="blue", command=self.clear_distribution)
        self.clear_button_distribution.grid(row=0,column=3, padx=20, pady= 10)
        self.upload_excel_button_distribution = customtkinter.CTkButton(self.button_distribution_frame, fg_color="dark green", text="Upload Excel", hover_color="green", command=self.upload_excel_distribution)
        self.upload_excel_button_distribution.grid(row=0,column=4)

        # create settings frame #######################################################################
        self.settings_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.settings_frame.grid_columnconfigure(0, weight=1)
        self.settings_frame.grid_rowconfigure(0, weight=1)
        # create tabview settings
        self.tabview_settings = customtkinter.CTkTabview(self.settings_frame,fg_color=("gray92"),text_color="White",segmented_button_fg_color="Dark Blue",segmented_button_unselected_color="Dark Blue")
        self.tabview_settings.grid(row=0, column=0, padx=(0, 0), pady=(0, 0), sticky="nsew")
        self.tabview_settings.add("Database Setting")
        self.tabview_settings.add("Connection Displan API")
        self.tabview_settings.add("Other Settings")
        # Database Setting tab
        self.tabview_settings.tab("Database Setting").grid_rowconfigure(0, weight=1)
        self.tabview_settings.tab("Database Setting").grid_columnconfigure(0, weight=1)
        self.db_setting_frame = customtkinter.CTkFrame(self.tabview_settings.tab("Database Setting"), corner_radius=20, fg_color="light blue")
        self.db_setting_frame.grid(row=0, column=0)
        self.db_setting_frame.grid_columnconfigure(0, weight=1)
        self.db_setting_frame.grid_rowconfigure(0, weight=1)
        self.db_setting_icon_label = customtkinter.CTkLabel(self.db_setting_frame, image=self.database_setting_image, text="")
        self.db_setting_icon_label.grid(row=0, column=0, columnspan=2, pady=(10,0))
        self.entry_warehouse_name_label = customtkinter.CTkLabel(self.db_setting_frame, text="Warehouse Name :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_warehouse_name_label.grid(row=1,column=0, padx=10, pady= (10,5), sticky="ew")
        self.entry_warehouse_name = customtkinter.CTkEntry(self.db_setting_frame, placeholder_text="", width=200)
        self.entry_warehouse_name.grid(row=1,column=1, padx=(0,10), pady= (10,5), sticky="ew")
        self.entry_server_label = customtkinter.CTkLabel(self.db_setting_frame, text="Server :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_server_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_server = customtkinter.CTkEntry(self.db_setting_frame, placeholder_text="", width=200)
        self.entry_server.grid(row=2,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.entry_database_name_label = customtkinter.CTkLabel(self.db_setting_frame, text="Database Name :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_database_name_label.grid(row=3,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_database_name = customtkinter.CTkEntry(self.db_setting_frame, placeholder_text="", width=200)
        self.entry_database_name.grid(row=3,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.entry_username_label = customtkinter.CTkLabel(self.db_setting_frame, text="Username :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_username_label.grid(row=4,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_username = customtkinter.CTkEntry(self.db_setting_frame, placeholder_text="", width=200)
        self.entry_username.grid(row=4,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.entry_password_label = customtkinter.CTkLabel(self.db_setting_frame, text="Password :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_password_label.grid(row=5,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_password = customtkinter.CTkEntry(self.db_setting_frame, placeholder_text="", width=200, show="*")
        self.entry_password.grid(row=5,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.connect_database_button = customtkinter.CTkButton(self.db_setting_frame, fg_color="dark blue", text="Connect Database", hover_color="blue", command=self.connect_database)
        self.connect_database_button.grid(row=6,column=0, columnspan=2, pady=10)
        self.reset_frame = customtkinter.CTkFrame(self.db_setting_frame, corner_radius=0, fg_color="transparent")
        self.reset_frame.grid(row=7,column=0, columnspan=2, pady=(0,10))
        self.reset_queuing_button = customtkinter.CTkButton(self.reset_frame, fg_color="dark red", text="Reset Queuing", hover_color="blue", command=self.reset_queuing)
        self.reset_queuing_button.pack(side="left", padx=10)
        self.reset_tappingSound_button = customtkinter.CTkButton(self.reset_frame, fg_color="dark red", text="Reset Tapping Sound", hover_color="blue", command=self.reset_tappingSound)
        self.reset_tappingSound_button.pack(side="left", padx=10)
        # self.delete_report_button = customtkinter.CTkButton(self.reset_frame, fg_color="dark red", text="Delete Report", hover_color="blue", command=self.delete_report)
        # self.delete_report_button.pack(side="left", padx=10)
        # Connection Displan API tab
        self.tabview_settings.tab("Connection Displan API").grid_rowconfigure(0, weight=1)
        self.tabview_settings.tab("Connection Displan API").grid_columnconfigure(0, weight=1)
        self.displan_api_frame = customtkinter.CTkFrame(self.tabview_settings.tab("Connection Displan API"), corner_radius=20, fg_color="light blue")
        self.displan_api_frame.grid(row=0, column=0)
        self.displan_api_frame.grid_columnconfigure(0, weight=1)
        self.displan_api_frame.grid_rowconfigure(0, weight=1)
        self.displan_api_icon_label = customtkinter.CTkLabel(self.displan_api_frame, image=self.api_image, text="")
        self.displan_api_icon_label.grid(row=0, column=0, columnspan=2, pady=(10,0))
        self.entry_warehouse_name_api_label = customtkinter.CTkLabel(self.displan_api_frame, text="Warehouse Name :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_warehouse_name_api_label.grid(row=1,column=0, padx=10, pady= (10,5), sticky="ew")
        self.entry_warehouse_name_api = customtkinter.CTkEntry(self.displan_api_frame, placeholder_text="", width=200)
        self.entry_warehouse_name_api.grid(row=1,column=1, padx=(0,10), pady= (10,5), sticky="ew")
        self.entry_warehouse_id_label = customtkinter.CTkLabel(self.displan_api_frame, text="Warehouse ID :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_warehouse_id_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_warehouse_id = customtkinter.CTkEntry(self.displan_api_frame, placeholder_text="", width=200)
        self.entry_warehouse_id.grid(row=2,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.entry_token_api_label = customtkinter.CTkLabel(self.displan_api_frame, text="Token :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_token_api_label.grid(row=3,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_token_api = customtkinter.CTkEntry(self.displan_api_frame, placeholder_text="", width=200)
        self.entry_token_api.grid(row=3,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.entry_domain_api_label = customtkinter.CTkLabel(self.displan_api_frame, text="Domain :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_domain_api_label.grid(row=4,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_domain_api = customtkinter.CTkEntry(self.displan_api_frame, placeholder_text="", width=200)
        self.entry_domain_api.grid(row=4,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.entry_url_api_label = customtkinter.CTkLabel(self.displan_api_frame, text="URL :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_url_api_label.grid(row=5,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_url_api = customtkinter.CTkEntry(self.displan_api_frame, placeholder_text="", width=200)
        self.entry_url_api.grid(row=5,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.entry_download_displan_api_label = customtkinter.CTkLabel(self.displan_api_frame, text="Download Displan\nEvery ? (Minute) : ", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_download_displan_api_label.grid(row=6,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_download_displan_api = customtkinter.CTkEntry(self.displan_api_frame, placeholder_text="", width=200)
        self.entry_download_displan_api.grid(row=6,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.get_displan_api_button = customtkinter.CTkButton(self.displan_api_frame, fg_color="dark blue", text="GET API DATA", hover_color="blue")
        self.get_displan_api_button.grid(row=7,column=0, columnspan=2, pady=10)
        # Other Settings tab
        self.tabview_settings.tab("Other Settings").grid_rowconfigure(0, weight=1)
        self.tabview_settings.tab("Other Settings").grid_columnconfigure(0, weight=1)
        self.other_settings_frame = customtkinter.CTkFrame(self.tabview_settings.tab("Other Settings"), corner_radius=20, fg_color="light blue")
        self.other_settings_frame.grid(row=0, column=0)
        self.other_settings_frame.grid_columnconfigure(0, weight=1)
        self.other_settings_frame.grid_rowconfigure(0, weight=1)
        self.call_rules_icon_label = customtkinter.CTkLabel(self.other_settings_frame, image=self.calling_image, text="")
        self.call_rules_icon_label.grid(row=0, column=0, columnspan=2, pady=(10,0))
        self.entry_call_rules_repeat_label = customtkinter.CTkLabel(self.other_settings_frame, text="Delay for repeat calling (Second):", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_call_rules_repeat_label.grid(row=1,column=0, padx=10, pady= (10,5), sticky="ew")
        self.entry_call_rules_repeat = customtkinter.CTkEntry(self.other_settings_frame, placeholder_text="", width=50)
        self.entry_call_rules_repeat.grid(row=1,column=1, padx=(0,10), pady= (10,5), sticky="ew")
        self.entry_call_rules_skip_label = customtkinter.CTkLabel(self.other_settings_frame, text="Skip queuing if called more than (times) :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_call_rules_skip_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_call_rules_skip = customtkinter.CTkEntry(self.other_settings_frame, placeholder_text="", width=50)
        self.entry_call_rules_skip.grid(row=2,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.entry_call_rules_delete_label = customtkinter.CTkLabel(self.other_settings_frame, text="Delete queuing after two session?", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_call_rules_delete_label.grid(row=3,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_call_rules_delete = customtkinter.CTkCheckBox(self.other_settings_frame, text="", width=20, onvalue="1", offvalue="0")
        self.entry_call_rules_delete.grid(row=3,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.set_call_rules_button = customtkinter.CTkButton(self.other_settings_frame, fg_color="dark blue", text="Set Call Rules", hover_color="blue", command=self.set_call_rules)
        self.set_call_rules_button.grid(row=4,column=0, columnspan=2, pady=(10,20))
        self.loading_schedule_frame = customtkinter.CTkFrame(self.other_settings_frame, corner_radius=20, fg_color="light blue")
        self.loading_schedule_frame.grid(row=5, column=0, columnspan=2, pady=(5,20))
        self.loading_schedule_label = customtkinter.CTkLabel(self.loading_schedule_frame, image=self.loading_schedule_image, text="")
        self.loading_schedule_label.grid(row=0, column=0, columnspan=2, pady=(5,0))
        self.call_loading_schedule_title = customtkinter.CTkLabel(self.loading_schedule_frame, text="Loading Schedule Setting", font=customtkinter.CTkFont(size=15, weight="bold"))
        self.call_loading_schedule_title.grid(row=1, column=0, columnspan=2, pady=10)
        self.checkbox_loading_schedule_hmin1 = customtkinter.CTkCheckBox(self.loading_schedule_frame, text="Allow H-1 Loading Schedule", width=20, onvalue="1", offvalue="0", command=self.update_date_range)
        self.checkbox_loading_schedule_hmin1.grid(row=2, column=0, padx=10, sticky="ew")
        self.checkbox_loading_schedule_hmin2 = customtkinter.CTkCheckBox(self.loading_schedule_frame, text="Allow H-2 Loading Schedule", width=20, onvalue="1", offvalue="0", command=self.update_date_range)
        self.checkbox_loading_schedule_hmin2.grid(row=3, column=0, padx=10, pady=10, sticky="ew")
        self.checkbox_loading_schedule_hmin3 = customtkinter.CTkCheckBox(self.loading_schedule_frame, text="Allow H-3 Loading Schedule", width=20, onvalue="1", offvalue="0", command=self.update_date_range)
        self.checkbox_loading_schedule_hmin3.grid(row=4, column=0, padx=10, pady=(0,10), sticky="ew")
        self.checkbox_loading_schedule_hplus1 = customtkinter.CTkCheckBox(self.loading_schedule_frame, text="Allow H+1 Loading Schedule", width=20, onvalue="1", offvalue="0", command=self.update_date_range)
        self.checkbox_loading_schedule_hplus1.grid(row=2, column=1, padx=10, sticky="ew")
        self.checkbox_loading_schedule_hplus2 = customtkinter.CTkCheckBox(self.loading_schedule_frame, text="Allow H+2 Loading Schedule", width=20, onvalue="1", offvalue="0", command=self.update_date_range)
        self.checkbox_loading_schedule_hplus2.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        self.checkbox_loading_schedule_hplus3 = customtkinter.CTkCheckBox(self.loading_schedule_frame, text="Allow H+3 Loading Schedule", width=20, onvalue="1", offvalue="0", command=self.update_date_range)
        self.checkbox_loading_schedule_hplus3.grid(row=4, column=1, padx=10, pady=(0,10), sticky="ew")

        # create user setting frame #######################################################################
        self.user_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.user_frame.grid_columnconfigure(0, weight=1)
        self.user_frame.grid_rowconfigure(0, weight=1)
        # create tabview user settings
        self.tabview_user_settings = customtkinter.CTkTabview(self.user_frame,fg_color=("gray92"),text_color="White",segmented_button_fg_color="Dark Blue",segmented_button_unselected_color="Dark Blue")
        self.tabview_user_settings.grid(row=0, column=0, padx=(0, 0), pady=(0, 0), sticky="nsew")
        self.tabview_user_settings.add("User Profile")
        self.tabview_user_settings.add("User Registration")
        # User Profile tab
        self.tabview_user_settings.tab("User Profile").grid_rowconfigure(0, weight=1)
        self.tabview_user_settings.tab("User Profile").grid_columnconfigure(0, weight=1)
        self.user_profile_frame = customtkinter.CTkFrame(self.tabview_user_settings.tab("User Profile"), corner_radius=20, fg_color="light blue")
        self.user_profile_frame.grid(row=0, column=0)
        self.user_profile_frame.grid_columnconfigure(0, weight=1)
        self.user_profile_frame.grid_rowconfigure(0, weight=1)
        self.user_profile_icon_label = customtkinter.CTkLabel(self.user_profile_frame, image=self.user_profil_image, text="")
        self.user_profile_icon_label.grid(row=0, column=0, columnspan=2, pady=(10,0))
        self.entry_username_user_label = customtkinter.CTkLabel(self.user_profile_frame, text="Username :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_username_user_label.grid(row=1,column=0, padx=10, pady= (10,5), sticky="ew")
        self.entry_username_user = customtkinter.CTkEntry(self.user_profile_frame, placeholder_text="", width=150)
        self.entry_username_user.grid(row=1,column=1, padx=(0,10), pady= (10,5), sticky="ew")
        self.entry_password_user_label = customtkinter.CTkLabel(self.user_profile_frame, text="Password :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_password_user_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_password_user = customtkinter.CTkEntry(self.user_profile_frame, placeholder_text="", show="*", width=150)
        self.entry_password_user.grid(row=2,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.entry_new_password_label = customtkinter.CTkLabel(self.user_profile_frame, text="New Password :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_new_password_label.grid(row=3,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_new_password = customtkinter.CTkEntry(self.user_profile_frame, placeholder_text="", show="*", width=150)
        self.entry_new_password.grid(row=3,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.entry_retype_new_password_label = customtkinter.CTkLabel(self.user_profile_frame, text="Retype Password :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_retype_new_password_label.grid(row=4,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_retype_new_password = customtkinter.CTkEntry(self.user_profile_frame, placeholder_text="", show="*", width=150)
        self.entry_retype_new_password.grid(row=4,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.change_password_button = customtkinter.CTkButton(self.user_profile_frame, fg_color="dark blue", text="Change Password", hover_color="blue", command=self.change_password)
        self.change_password_button.grid(row=5,column=0, columnspan=2, pady=10)
        # User Registration tab
        self.tabview_user_settings.tab("User Registration").grid_rowconfigure(0, weight=1)
        self.tabview_user_settings.tab("User Registration").grid_columnconfigure(0, weight=1)
        self.user_registration_frame = customtkinter.CTkFrame(self.tabview_user_settings.tab("User Registration"), corner_radius=0, fg_color="transparent")
        self.user_registration_frame.grid(row=0, column=0 , sticky="nesw")
        self.user_registration_frame.grid_columnconfigure(0, weight=1)
        self.user_registration_frame.grid_rowconfigure(0, weight=1)
        # user table
        self.user_registration_table_frame = customtkinter.CTkFrame(self.user_registration_frame, corner_radius=0, fg_color="blue")
        self.user_registration_table_frame.grid(row=0, column=0 , sticky="nesw")
        self.userRegistration_table = ttk.Treeview(self.user_registration_table_frame)
        self.userRegistration_table.pack(side="left", expand=tk.YES, fill=tk.BOTH)
        self.userRegistration_table["columns"] = ("No", "User Name", "Email", "Remark", "Password", "alertNote", "reportViewer", "packaging", "productReg", "tappingPoint", "warehouseFlow", "rfidManualCall", "vetting", "unloading", "distribution", "databaseSetting", "displanApi", "otherSetting", "userProfile", "userRegistration")
        self.userRegistration_table.column("#0", width=0,stretch=tk.NO)
        self.userRegistration_table.column("No", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("User Name", anchor=tk.CENTER, width=30, minwidth=0,stretch=tk.YES)
        self.userRegistration_table.column("Email", anchor=tk.CENTER, width=110, minwidth=0,stretch=tk.YES)
        self.userRegistration_table.column("Remark", anchor=tk.CENTER, width=200, minwidth=0,stretch=tk.YES)
        self.userRegistration_table.column("Password", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("alertNote", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("reportViewer", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("packaging", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("productReg", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("tappingPoint", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("warehouseFlow", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("rfidManualCall", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("vetting", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("unloading", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("distribution", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("databaseSetting", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("displanApi", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("otherSetting", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("userProfile", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.column("userRegistration", anchor=tk.CENTER, width=0,stretch=tk.NO)
        self.userRegistration_table.heading("#0", text="", anchor=tk.W)
        self.userRegistration_table.heading("No", text="", anchor=tk.W)
        self.userRegistration_table.heading("User Name", text="User Name", anchor=tk.CENTER)
        self.userRegistration_table.heading("Email", text="Email", anchor=tk.CENTER)
        self.userRegistration_table.heading("Remark", text="Remark", anchor=tk.CENTER)
        self.userRegistration_table.heading("Password", text="", anchor=tk.W)
        self.userRegistration_table.heading("alertNote", text="", anchor=tk.W)
        self.userRegistration_table.heading("reportViewer", text="", anchor=tk.W)
        self.userRegistration_table.heading("packaging", text="", anchor=tk.W)
        self.userRegistration_table.heading("productReg", text="", anchor=tk.W)
        self.userRegistration_table.heading("tappingPoint", text="", anchor=tk.W)
        self.userRegistration_table.heading("warehouseFlow", text="", anchor=tk.W)
        self.userRegistration_table.heading("rfidManualCall", text="", anchor=tk.W)
        self.userRegistration_table.heading("vetting", text="", anchor=tk.W)
        self.userRegistration_table.heading("unloading", text="", anchor=tk.W)
        self.userRegistration_table.heading("distribution", text="", anchor=tk.W)
        self.userRegistration_table.heading("databaseSetting", text="", anchor=tk.W)
        self.userRegistration_table.heading("displanApi", text="", anchor=tk.W)
        self.userRegistration_table.heading("otherSetting", text="", anchor=tk.W)
        self.userRegistration_table.heading("userProfile", text="", anchor=tk.W)
        self.userRegistration_table.heading("userRegistration", text="", anchor=tk.W)
        self.userRegistration_table.tag_configure("oddrow",background="white",font=(None, 13))
        self.userRegistration_table.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.userRegistration_table_scrollbar = customtkinter.CTkScrollbar(self.user_registration_table_frame, hover=True, button_hover_color="yellow", command=self.userRegistration_table.yview)
        self.userRegistration_table_scrollbar.pack(side="left",fill=tk.Y)
        self.userRegistration_table.configure(yscrollcommand=self.userRegistration_table_scrollbar.set)
        self.userRegistration_table.bind("<<TreeviewSelect>>", self.on_tree_userRegistration_select)
        # user form
        self.user_registration_entry_frame = customtkinter.CTkFrame(self.user_registration_frame, corner_radius=10, fg_color="light blue")
        self.user_registration_entry_frame.grid(row=1, column=0 , sticky="nesw")
        self.entry_regis_username_label = customtkinter.CTkLabel(self.user_registration_entry_frame, text="Username :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_regis_username_label.grid(row=0,column=0, padx=10, pady= (10,5), sticky="ew")
        self.entry_regis_username = customtkinter.CTkEntry(self.user_registration_entry_frame, placeholder_text="", width=150)
        self.entry_regis_username.grid(row=0,column=1, padx=(0,10), pady= (10,5), sticky="ew")
        self.entry_regis_password_label = customtkinter.CTkLabel(self.user_registration_entry_frame, text="Password :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_regis_password_label.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_regis_password = customtkinter.CTkEntry(self.user_registration_entry_frame, placeholder_text="", show="*", width=150)
        self.entry_regis_password.grid(row=1,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.entry_regis_retypePassword_label = customtkinter.CTkLabel(self.user_registration_entry_frame, text="Retype Password :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_regis_retypePassword_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_regis_retypePassword = customtkinter.CTkEntry(self.user_registration_entry_frame, placeholder_text="", show="*", width=150)
        self.entry_regis_retypePassword.grid(row=2,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.entry_regis_email_label = customtkinter.CTkLabel(self.user_registration_entry_frame, text="Email :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_regis_email_label.grid(row=3,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_regis_email = customtkinter.CTkEntry(self.user_registration_entry_frame, placeholder_text="", width=150)
        self.entry_regis_email.grid(row=3,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.entry_regis_remark_label = customtkinter.CTkLabel(self.user_registration_entry_frame, text="Remark :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_regis_remark_label.grid(row=4,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_regis_remark = customtkinter.CTkEntry(self.user_registration_entry_frame, placeholder_text="", width=150)
        self.entry_regis_remark.grid(row=4,column=1, padx=(0,10), pady= (5,10), sticky="ew")
        self.checkbox_report_label = customtkinter.CTkLabel(self.user_registration_entry_frame, text="Report :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.checkbox_report_label.grid(row=0,column=2, padx=10, pady= 10, sticky="ew")
        self.checkbox_alertNoteMenu = customtkinter.CTkCheckBox(self.user_registration_entry_frame, text="Alert Note", width=20, onvalue="1", offvalue="0")
        self.checkbox_alertNoteMenu.grid(row=0, column=3, padx=5, sticky="ew")
        self.checkbox_reportViewerMenu = customtkinter.CTkCheckBox(self.user_registration_entry_frame, text="Report Viewer", width=20, onvalue="1", offvalue="0")
        self.checkbox_reportViewerMenu.grid(row=0, column=4, padx=5, sticky="ew")
        self.checkbox_queuingSettings_label = customtkinter.CTkLabel(self.user_registration_entry_frame, text="Queuing Settings :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.checkbox_queuingSettings_label.grid(row=1,column=2, padx=10, pady= 10, sticky="ew")
        self.checkbox_packagingMenu = customtkinter.CTkCheckBox(self.user_registration_entry_frame, text="Packaging", width=20, onvalue="1", offvalue="0")
        self.checkbox_packagingMenu.grid(row=1, column=3, padx=5, sticky="ew")
        self.checkbox_productMenu = customtkinter.CTkCheckBox(self.user_registration_entry_frame, text="Product Reg", width=20, onvalue="1", offvalue="0")
        self.checkbox_productMenu.grid(row=1, column=4, padx=5, sticky="ew")
        self.checkbox_tappingPointMenu = customtkinter.CTkCheckBox(self.user_registration_entry_frame, text="Tapping Point", width=20, onvalue="1", offvalue="0")
        self.checkbox_tappingPointMenu.grid(row=1, column=5, padx=5, sticky="ew")
        self.checkbox_warehouseFlowMenu = customtkinter.CTkCheckBox(self.user_registration_entry_frame, text="Warehouse Flow", width=20, onvalue="1", offvalue="0")
        self.checkbox_warehouseFlowMenu.grid(row=1, column=6, padx=5, sticky="ew")
        self.checkbox_rfidManualCallMenu = customtkinter.CTkCheckBox(self.user_registration_entry_frame, text="Queuing Manual Call", width=20, onvalue="1", offvalue="0")
        self.checkbox_rfidManualCallMenu.grid(row=1, column=7, padx=5, sticky="ew")
        self.checkbox_truckData_label = customtkinter.CTkLabel(self.user_registration_entry_frame, text="Truck Data :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.checkbox_truckData_label.grid(row=2,column=2, padx=10, pady=10, sticky="ew")
        self.checkbox_vettingMenu = customtkinter.CTkCheckBox(self.user_registration_entry_frame, text="Vetting", width=20, onvalue="1", offvalue="0")
        self.checkbox_vettingMenu.grid(row=2, column=3, padx=5, sticky="ew")
        self.checkbox_unloadingMenu = customtkinter.CTkCheckBox(self.user_registration_entry_frame, text="Unloading", width=20, onvalue="1", offvalue="0")
        self.checkbox_unloadingMenu.grid(row=2, column=4, padx=5, sticky="ew")
        self.checkbox_distributionMenu = customtkinter.CTkCheckBox(self.user_registration_entry_frame, text="Distribution", width=20, onvalue="1", offvalue="0")
        self.checkbox_distributionMenu.grid(row=2, column=5, padx=5, sticky="ew")
        self.checkbox_systemSetting_label = customtkinter.CTkLabel(self.user_registration_entry_frame, text="System Setting :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.checkbox_systemSetting_label.grid(row=3,column=2, padx=10, pady=10, sticky="ew")
        self.checkbox_databaseSettingMenu = customtkinter.CTkCheckBox(self.user_registration_entry_frame, text="Database", width=20, onvalue="1", offvalue="0")
        self.checkbox_databaseSettingMenu.grid(row=3, column=3, padx=5, sticky="ew")
        self.checkbox_displanApiMenu = customtkinter.CTkCheckBox(self.user_registration_entry_frame, text="Displan API", width=20, onvalue="1", offvalue="0")
        self.checkbox_displanApiMenu.grid(row=3, column=4, padx=5, sticky="ew")
        self.checkbox_otherSettingMenu = customtkinter.CTkCheckBox(self.user_registration_entry_frame, text="Other Setting", width=20, onvalue="1", offvalue="0")
        self.checkbox_otherSettingMenu.grid(row=3, column=5, padx=5, sticky="ew")
        self.checkbox_userSetting_label = customtkinter.CTkLabel(self.user_registration_entry_frame, text="User Setting :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.checkbox_userSetting_label.grid(row=4,column=2, padx=10, pady=10, sticky="ew")
        self.checkbox_userProfileMenu = customtkinter.CTkCheckBox(self.user_registration_entry_frame, text="User Profile", width=20, onvalue="1", offvalue="0")
        self.checkbox_userProfileMenu.grid(row=4, column=3, padx=5, sticky="ew")
        self.checkbox_userRegistrationMenu = customtkinter.CTkCheckBox(self.user_registration_entry_frame, text="User Registration", width=20, onvalue="1", offvalue="0")
        self.checkbox_userRegistrationMenu.grid(row=4, column=4, padx=5, sticky="ew")
        # user button
        self.user_registration_button_frame = customtkinter.CTkFrame(self.user_registration_frame, corner_radius=0, fg_color="transparent")
        self.user_registration_button_frame.grid(row=2, column=0)
        self.add_button_user_registration = customtkinter.CTkButton(self.user_registration_button_frame, fg_color="dark blue", text="ADD", hover_color="blue", command=self.add_user_registration)
        self.add_button_user_registration.grid(row=0,column=0)
        self.remove_button_user_registration = customtkinter.CTkButton(self.user_registration_button_frame, fg_color="dark blue", text="Remove", hover_color="blue", command=self.remove_user_registration)
        self.remove_button_user_registration.grid(row=0,column=1, padx=20, pady= 10)
        self.replace_button_user_registration = customtkinter.CTkButton(self.user_registration_button_frame, fg_color="dark blue", text="Replace", hover_color="blue", command=self.replace_user_registration)
        self.replace_button_user_registration.grid(row=0,column=2)

        # create log frame #######################################################################
        self.log_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(0, weight=1)
        self.login_logout_frame = customtkinter.CTkFrame(self.log_frame, corner_radius=20, fg_color="light blue")
        self.login_logout_frame.grid(row=0, column=0)
        self.login_logout_frame.grid_columnconfigure(0, weight=1)
        self.login_logout_frame.grid_rowconfigure(0, weight=1)
        self.login_logout_icon_label = customtkinter.CTkLabel(self.login_logout_frame, image=self.login_logout_image, text="")
        self.login_logout_icon_label.grid(row=0, column=0, columnspan=2, pady=(10,0))
        self.entry_username_log_label = customtkinter.CTkLabel(self.login_logout_frame, text="Username :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_username_log_label.grid(row=1,column=0, padx=10, pady= (10,5), sticky="ew")
        self.entry_username_log = customtkinter.CTkEntry(self.login_logout_frame, placeholder_text="", width=150)
        self.entry_username_log.grid(row=1,column=1, padx=(0,10), pady= (10,5), sticky="ew")
        self.entry_password_log_label = customtkinter.CTkLabel(self.login_logout_frame, text="Password :", font=customtkinter.CTkFont(size=15, weight="bold"),anchor="w")
        self.entry_password_log_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_password_log = customtkinter.CTkEntry(self.login_logout_frame, placeholder_text="", show="*", width=150)
        self.entry_password_log.grid(row=2,column=1, padx=(0,10), pady= 5, sticky="ew")
        self.entry_password_log.bind('<Return>', self.login)
        self.login_button = customtkinter.CTkButton(self.login_logout_frame, fg_color="dark green", text="Login", hover_color="green", command=self.login)
        self.login_button.grid(row=3,column=0, columnspan=2, pady=10)
        self.logout_button = customtkinter.CTkButton(self.login_logout_frame, fg_color="gray", text="Logout / Exit", hover_color="red", state="disabled", command=self.logout)
        self.logout_button.grid(row=4,column=0, columnspan=2, pady=(0,10))

        # select default frame
        self.select_frame_by_name("log")

        # database entry auto fill
        self.database_entry_auto_fill()

        # api entry auto fill
        self.api_entry_auto_fill()

        # thread for refresh table
        global stop_event, thread
        stop_event = threading.Event()
        thread = threading.Thread(target=self.thread1, args=(stop_event,))

        # database
        self.try_connect_db()

        # thread for refresh table
        global stop_event2, thread2
        stop_event2 = threading.Event()
        thread2 = threading.Thread(target=self.get_displan_api, args=(stop_event2,))
        thread2.start()

        # set checkbox loadingList schedule
        self.set_schedule_loadingList()

        # saat aplikasi di tutup fungsi stop menghentika semua thread
        self.wm_protocol("WM_DELETE_WINDOW", self.stop)

    def database_entry_auto_fill(self):
        global wh_name, driver_db, server_db, dbname, username, password
        # local database (SQLITE)
        file = sq.connect("./dataBase/dataBase.db")
        c = file.cursor()
        # tb data database sql express
        c.execute("""CREATE TABLE IF NOT EXISTS tb_database(
            wh_name text,
            server text,
            dbname text,
            username text,
            password text
        )""")
        file.commit()
        file.close()
        try:
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            c.execute("SELECT * FROM tb_database")
            rec = c.fetchall()
            if len(rec) != 0 :
                for i in rec:
                    driver_db = "ODBC Driver 17 for SQL Server"
                    wh_name = i[0]
                    server_db = i[1]
                    dbname = i[2]
                    username = i[3]
                    password = i[4]
                    self.entry_warehouse_name.insert(0, i[0])
                    self.entry_server.insert(0, i[1])
                    self.entry_database_name.insert(0, i[2])
                    self.entry_username.insert(0, i[3])
                    self.entry_password.insert(0, i[4])
            else:
                self.entry_warehouse_name.insert(0, "")
                self.entry_server.insert(0, "")
                self.entry_database_name.insert(0, "")
                self.entry_username.insert(0, "")
                self.entry_password.insert(0, "")
                driver_db = "ODBC Driver 17 for SQL Server"
                server_db = ""
                dbname = ""
                username = ""
                password = ""
            file.commit()
            file.close()
        except:
            pass

    def api_entry_auto_fill(self):
        # local database (SQLITE)
        file = sq.connect("./dataBase/dataBase.db")
        c = file.cursor()
        # tb data database sql express
        c.execute("""CREATE TABLE IF NOT EXISTS tb_api_setting(
            wh_id text,
            token text,
            domain text,
            url text,
            schedule_api text
        )""")
        file.commit()
        file.close()
        try:
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            c.execute("SELECT * FROM tb_api_setting")
            rec = c.fetchall()
            if len(rec) != 0 :
                for i in rec:
                    wh_id = i[0]
                    token = i[1]
                    domain = i[2]
                    url = i[3]
                    schedule = i[4]
                    self.entry_warehouse_name_api.insert(0, wh_name)
                    self.entry_warehouse_name_api.configure(state="disable")
                    self.entry_warehouse_id.insert(0, wh_id)
                    self.entry_token_api.insert(0, token)
                    self.entry_domain_api.insert(0, domain)
                    self.entry_url_api.insert(0, url)
                    self.entry_download_displan_api.insert(0, schedule)
            else:
                self.entry_warehouse_name_api.insert(0, wh_name)
                self.entry_warehouse_name_api.configure(state="disable")
                self.entry_warehouse_name_api.insert(0, "")
                self.entry_warehouse_id.insert(0, "")
                self.entry_token_api.insert(0, "")
                self.entry_domain_api.insert(0, "")
                self.entry_url_api.insert(0, "")
                self.entry_download_displan_api.insert(0, "")
            file.commit()
            file.close()
        except:
            pass

    def try_connect_db(self):
        global driver_db, server_db, dbname, username, password
        timeout = 1
        try:
            connection = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password};Connect Timeout={timeout}', timeout=1)
            connection.close()
            self.database()
            self.optionMenu_packaging()
            self.optionMenu_warehouseAssign()
            thread.start() # THREAD UNTUK TUGAS2 YANG REAL TIME (self.thread1)
            self.wm_protocol("WM_DELETE_WINDOW", self.stop)
        except:
            self.select_frame_by_name("settings")
            error_message = "Unable to connect to the database"
            messagebox.showerror("Error", error_message)
            self.dashboard_button.configure(state="disabled", text_color="gray")
            self.report_button.configure(state="disabled", text_color="gray")
            self.Qsettings_button.configure(state="disabled", text_color="gray")
            self.truck_data_button.configure(state="disabled", text_color="gray")
            self.user_button.configure(state="disabled", text_color="gray")
            self.log_button.configure(state="disabled", text_color="gray")
            self.reset_queuing_button.configure(state="disabled", fg_color="gray")
            # self.delete_report_button.configure(state="disabled", fg_color="gray")

    def database(self):
        # Create a thread to run the database connection code
        thread_database_setup = threading.Thread(target=self.database_setup)
        thread_database_setup.start()

    def database_setup(self):
        global driver_db, server_db, dbname, username, password

        # network database (SQL EXPRESS)
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

    def connect_database(self):
        # Create a thread to run the database connection code
        thread_connect_to_database = threading.Thread(target=self.connect_to_database)
        thread_connect_to_database.start()

    def connect_to_database(self):
        global wh_name, driver_db, server_db, dbname, username, password

        driver_db = "ODBC Driver 17 for SQL Server"
        wh_name = self.entry_warehouse_name.get()
        server_db = self.entry_server.get()
        dbname = self.entry_database_name.get()
        username = self.entry_username.get()
        password = self.entry_password.get()
        timeout = 1
        try:
            self.connect_database_button.configure(state="disabled", text="Please Wait", fg_color="gray")
            connection = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password};Connect Timeout={timeout}')
            connection.close()
            messagebox.showinfo("Connection", "Koneksi Berhasil")
            self.connect_database_button.configure(state="normal", text="Connect Database", fg_color="dark blue")
            self.dashboard_button.configure(state="normal", text_color="black")
            self.report_button.configure(state="normal", text_color="black")
            self.Qsettings_button.configure(state="normal", text_color="black")
            self.truck_data_button.configure(state="normal", text_color="black")
            self.user_button.configure(state="normal", text_color="black")
            self.log_button.configure(state="normal", text_color="black")
            self.reset_queuing_button.configure(state="normal", fg_color="dark red")
            self.delete_report_button.configure(state="normal", fg_color="dark red")
            # delete data in tb_database
            try:
                file = sq.connect("./dataBase/dataBase.db")
                c = file.cursor()
                c.execute("DELETE from tb_database")
                file.commit()
                file.close()
            except:
                pass
            # save data connection to sqlite table
            try:
                file = sq.connect("./dataBase/dataBase.db")
                c = file.cursor()
                c.execute("INSERT INTO tb_database VALUEs(:wh_name,:server,:dbname,:username,:password)",{
                    "wh_name": wh_name,
                    "server": server_db,
                    "dbname": dbname,
                    "username": username,
                    "password": password
                })
                file.commit()
                file.close()
            except:
                pass
            self.database()
            self.optionMenu_packaging()
            self.optionMenu_warehouseAssign()
            self.wm_protocol("WM_DELETE_WINDOW", self.stop)
        except:
            error_message = "Unable to connect to the database"
            messagebox.showerror("Error", error_message)
            self.connect_database_button.configure(state="normal", text="Connect Database", fg_color="dark blue")
            self.dashboard_button.configure(state="disabled", text_color="gray")
            self.report_button.configure(state="disabled", text_color="gray")
            self.Qsettings_button.configure(state="disabled", text_color="gray")
            self.truck_data_button.configure(state="disabled", text_color="gray")
            self.user_button.configure(state="disabled", text_color="gray")
            self.log_button.configure(state="disabled", text_color="gray")
            self.reset_queuing_button.configure(state="disabled", fg_color="gray")
            # self.delete_report_button.configure(state="disabled", fg_color="gray")

    def select_frame_by_name(self, name):
        # set button color for selected button
        self.dashboard_button.configure(fg_color=("#1aa8b8") if name == "dashboard" else "transparent")
        self.report_button.configure(fg_color=("#1aa8b8") if name == "report" else "transparent")
        self.Qsettings_button.configure(fg_color=("#1aa8b8") if name == "Qsettings" else "transparent")
        self.truck_data_button.configure(fg_color=("#1aa8b8") if name == "truck_data" else "transparent")
        self.settings_button.configure(fg_color=("#1aa8b8") if name == "settings" else "transparent")
        self.user_button.configure(fg_color=("#1aa8b8") if name == "user" else "transparent")
        self.log_button.configure(fg_color=("#1aa8b8") if name == "log" else "transparent")

        # show selected frame
        if name == "dashboard":
            self.dashboard_frame.grid(row=0, column=1, sticky="nsew")
            functions = "Warehouse"
            status = "Activate"
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT remark FROM tb_tappingPointReg WHERE functions = ?  AND status = ?",(functions, status))
            rec = c.fetchall()
            status_gudang=" "
            # input to tabel GUI
            for i in rec:
                status_gudang=status_gudang+i[0]+","+" "
            file.commit()
            file.close()
            self.warehouse_status.configure(text=f"Gudang Activate:{status_gudang}")
        else:
            self.dashboard_frame.grid_forget()
        if name == "report":
            self.report_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.report_frame.grid_forget()
        if name == "Qsettings":
            self.Qsettings_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.Qsettings_frame.grid_forget()
        if name == "truck_data":
            self.truck_data_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.truck_data_frame.grid_forget()
        if name == "settings":
            self.settings_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.settings_frame.grid_forget()
        if name == "user":
            self.user_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.user_frame.grid_forget()
        if name == "log":
            self.log_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.log_frame.grid_forget()

    def dashboard_button_event(self):
        self.select_frame_by_name("dashboard")

        functions = "Warehouse"
        status = "Activate"
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT remark FROM tb_tappingPointReg WHERE functions = ?  AND status = ?",(functions, status))
        rec = c.fetchall()
        status_gudang=" "
        # input to tabel GUI
        for i in rec:
            status_gudang=status_gudang+i[0]+","+" "
        file.commit()
        file.close()
        self.warehouse_status.configure(text=f"Gudang Activate:{status_gudang}")

    def report_button_event(self):
        self.select_frame_by_name("report")

        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # data alert inbox
        for record in self.alert_table.get_children():
            self.alert_table.delete(record)
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_alert")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.alert_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=("evenrow",))
            else:
                self.alert_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=("oddrow",))
            count+=1
        file.commit()
        file.close()

    def qsettings_button_event(self):
        self.select_frame_by_name("Qsettings")
        # real all data in all tab and show in each tab
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()

        # data packaging registration
        for record in self.packaging_table.get_children():
            self.packaging_table.delete(record)
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_packagingReg")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.packaging_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4]), tags=("evenrow",))
            else:
                self.packaging_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4]), tags=("oddrow",))
            count+=1

        # data warehouse sla setting
        for record in self.warehouse_slaSetting_table.get_children():
            self.warehouse_slaSetting_table.delete(record)
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_warehouseSla")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.warehouse_slaSetting_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10]), tags=("evenrow",))
            else:
                self.warehouse_slaSetting_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10]), tags=("oddrow",))
            count+=1

        # data tapping point registration
        for record in self.tappingPoint_table.get_children():
            self.tappingPoint_table.delete(record)
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_tappingPointReg")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.tappingPoint_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("evenrow",))
            else:
                self.tappingPoint_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("oddrow",))
            count+=1

        # data warehouse flow registration
        for record in self.warehouse_flow_table.get_children():
            self.warehouse_flow_table.delete(record)
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_warehouseFlow")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.warehouse_flow_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("evenrow",))
            else:
                self.warehouse_flow_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("oddrow",))
            count+=1

        # data waiting list
        for record in self.manualCall_table.get_children():
            self.manualCall_table.delete(record)
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_waitinglist")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.manualCall_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("evenrow",))
            else:
                self.manualCall_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("oddrow",))
            count+=1

        # data alert inbox
        for record in self.rfidPairing_table.get_children():
            self.rfidPairing_table.delete(record)
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_queuing_status")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.rfidPairing_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18], i[19]), tags=("evenrow",))
            else:
                self.rfidPairing_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18], i[19]), tags=("oddrow",))
            count+=1

    def truck_data_button_event(self):
        self.select_frame_by_name("truck_data")
        # real all data in all tab and show in each tab
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()

        # data tab vetting
        for record in self.vetting_table.get_children():
            self.vetting_table.delete(record)
        count = 0
        c.execute("SELECT * FROM tb_vetting")
        rec = c.fetchall()
        urutan = 1
        for i in rec:
            if count % 2 == 0:
                self.vetting_table.insert(parent="",index="end", iid=count, text="", values=(i[9], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
            else:
                self.vetting_table.insert(parent="",index="end", iid=count, text="", values=(i[9], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
            count+=1
            urutan+=1

        # data tab unloading
        for record in self.unloading_table.get_children():
            self.unloading_table.delete(record)
        count = 0
        c.execute("SELECT * FROM tb_unloading")
        rec = c.fetchall()
        urutan = 1
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.unloading_table.insert(parent="",index="end", iid=count, text="", values=(i[12], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("evenrow",))
            else:
                self.unloading_table.insert(parent="",index="end", iid=count, text="", values=(i[12], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("oddrow",))
            count+=1
            urutan+=1

        # data tab distribution -> loading list
        for record in self.loadingList_table.get_children():
            self.loadingList_table.delete(record)
        count = 0
        c.execute("SELECT * FROM tb_loadingList")
        rec = c.fetchall()
        # input to tabel GUI
        urutan = 1
        for i in rec:
            if count % 2 == 0:
                self.loadingList_table.insert(parent="",index="end", iid=count, text="", values=(i[13], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12]), tags=("evenrow",))
            else:
                self.loadingList_table.insert(parent="",index="end", iid=count, text="", values=(i[13], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12]), tags=("oddrow",))
            count+=1
            urutan+=1

        # data tab distribution -> displan data
        for record in self.displanData_table.get_children():
            self.displanData_table.delete(record)
        count = 0
        c.execute("SELECT * FROM tb_displanData")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.displanData_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("evenrow",))
            else:
                self.displanData_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("oddrow",))
            count+=1

    def settings_button_event(self):
        self.select_frame_by_name("settings")

        # input entry database
        self.entry_warehouse_name.delete(0,tk.END)
        self.entry_server.delete(0,tk.END)
        self.entry_database_name.delete(0,tk.END)
        self.entry_username.delete(0,tk.END)
        self.entry_password.delete(0,tk.END)
        file = sq.connect("./dataBase/dataBase.db")
        c = file.cursor()
        c.execute("SELECT * FROM tb_database")
        rec = c.fetchall()
        if len(rec) != 0 :
            for i in rec:
                self.entry_warehouse_name.insert(0, i[0])
                self.entry_server.insert(0, i[1])
                self.entry_database_name.insert(0, i[2])
                self.entry_username.insert(0, i[3])
                self.entry_password.insert(0, i[4])
        else:
            self.entry_warehouse_name.insert(0, "")
            self.entry_server.insert(0, "")
            self.entry_database_name.insert(0, "")
            self.entry_username.insert(0, "")
            self.entry_password.insert(0, "")
        file.commit()
        file.close()

        # input entry database
        self.entry_call_rules_repeat.delete(0,tk.END)
        self.entry_call_rules_skip.delete(0,tk.END)
        self.entry_call_rules_delete.deselect()

        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT * FROM tb_call_rules")
        rec = c.fetchall()
        if len(rec) != 0 :
            for i in rec:
                self.entry_call_rules_repeat.insert(0, i[0])
                self.entry_call_rules_skip.insert(0, i[1])
                if i[2] == True:
                    self.entry_call_rules_delete.select()
        else:
            self.entry_call_rules_repeat.insert(0, "")
            self.entry_call_rules_skip.insert(0, "")
            self.entry_call_rules_delete.deselect()
        file.commit()
        file.close()

    def user_button_event(self):
        self.select_frame_by_name("user")

        # hapus item di tabel GUI
        for record in self.userRegistration_table.get_children():
            self.userRegistration_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_userRegistration")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.userRegistration_table.insert(parent="",index="end", iid=count, text="", values=(i[19], i[0], i[2], i[3], i[1], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18]), tags=("evenrow",))
            else:
                self.userRegistration_table.insert(parent="",index="end", iid=count, text="", values=(i[19], i[0], i[2], i[3], i[1], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18]), tags=("oddrow",))
            count+=1
        file.commit()
        file.close()

    def log_button_event(self):
        self.select_frame_by_name("log")

    def optionmenuStatus_callback(self, choice: str):
        global status
        status = choice

    def optionmenuCategory_callback(self, choice: str):
        global category
        category = choice

    # optionmenuPackaging_callback juga untuk set product optionmenu yang tergantung packaging
    def optionmenuPackaging_callback(self, choice: str):
        global packaging
        packaging = choice

        # set tupple for packaging name to apply to optionmenu
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()

        # set product optionmenu
        c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?", (self.optionmenu_packaging_product.get(),))
        rec = c.fetchall()
        productlist_productreg = []
        for i in rec:
            productlist_productreg.append(i[0])
        self.optionmenu_product.configure(values=productlist_productreg)

        c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",(self.optionmenu_packaging_distribution.get(),))
        rec = c.fetchall()
        productlist_dist = []
        for i in rec:
            productlist_dist.append(i[0])
        self.optionmenu_product_distribution.configure(values=productlist_dist)

        c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",(self.optionmenu_packaging_sla.get(),))
        rec = c.fetchall()
        productlist_sla = []
        for i in rec:
            productlist_sla.append(i[0])
        self.optionmenu_product_sla.configure(values=productlist_sla)

        c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",(self.optionmenu_packaging.get(),))
        rec = c.fetchall()
        productlist_unloading = []
        for i in rec:
            productlist_unloading.append(i[0])
        self.optionmenu_product.configure(values=productlist_unloading)

        c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",(self.report_viewer_total_time_optionmenu_packaging.get(),))
        rec = c.fetchall()
        productlist_report_total_time = ["-"]
        for i in rec:
            productlist_report_total_time.append(i[0])
        self.report_viewer_total_time_optionmenu_product.configure(values=productlist_report_total_time)

        c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",(self.report_viewer_transaction_time_optionmenu_packaging.get(),))
        rec = c.fetchall()
        productlist_report_transaction_time = ["-"]
        for i in rec:
            productlist_report_transaction_time.append(i[0])
        self.report_viewer_transaction_time_optionmenu_product.configure(values=productlist_report_transaction_time)

        c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",(self.report_viewer_peak_hour_optionmenu_packaging.get(),))
        rec = c.fetchall()
        productlist_report_peak_hour = ["-"]
        for i in rec:
            productlist_report_peak_hour.append(i[0])
        self.report_viewer_peak_hour_optionmenu_product.configure(values=productlist_report_peak_hour)

        c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",(self.report_viewer_total_queuing_optionmenu_packaging.get(),))
        rec = c.fetchall()
        productlist_report_total_queuing = ["-"]
        for i in rec:
            productlist_report_total_queuing.append(i[0])
        self.report_viewer_total_queuing_optionmenu_product.configure(values=productlist_report_total_queuing)

        c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",(self.report_viewer_sla_optionmenu_packaging.get(),))
        rec = c.fetchall()
        productlist_report_sla = ["-"]
        for i in rec:
            productlist_report_sla.append(i[0])
        self.report_viewer_sla_optionmenu_product.configure(values=productlist_report_sla)

        file.commit()
        file.close()

    def optionmenuProduct_callback(self, choice: str):
        global product
        product = choice

    def optionmenuVehicle_callback(self, choice: str):
        global vehicle
        vehicle = choice

    def optionmenuPackaging_product_callback(self, choice: str):
        # global packaging_product
        packaging_product = choice

        # hapus item entry product tabel
        for record in self.product_table.get_children():
            self.product_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_productReg WHERE packaging = ?",(packaging_product,))
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.product_table.insert(parent="",index="end", iid=count, text="", values=(i[1], i[2], i[3], i[4]), tags=("evenrow",))
            else:
                self.product_table.insert(parent="",index="end", iid=count, text="", values=(i[1], i[2], i[3], i[4]), tags=("oddrow",))
            count+=1
        # reset input
        self.entry_product_group.delete(0,tk.END)
        self.entry_product_group_remark.delete(0,tk.END)

        file.commit()
        file.close()

    # FUNGSI PENCARIAN

    def search_vetting(self):
        query = self.entry_search_vetting.get()
        selections = []
        for child in self.vetting_table.get_children():
            if query.capitalize() in self.vetting_table.item(child)['values'] or query.lower() in self.vetting_table.item(child)['values'] or query.upper() in self.vetting_table.item(child)['values'] or query.lower() in self.vetting_table.item(child)['values'] or query.swapcase() in self.vetting_table.item(child)['values'] or query.title() in self.vetting_table.item(child)['values']:
                selections.append(child)
        self.vetting_table.selection_set(selections)

    def search_unloading(self):
        query = self.entry_search_unloading.get()
        selections = []
        for child in self.unloading_table.get_children():
            if query.capitalize() in self.unloading_table.item(child)['values'] or query.lower() in self.unloading_table.item(child)['values'] or query.upper() in self.unloading_table.item(child)['values'] or query.lower() in self.unloading_table.item(child)['values'] or query.swapcase() in self.unloading_table.item(child)['values'] or query.title() in self.unloading_table.item(child)['values']:
                selections.append(child)
        self.unloading_table.selection_set(selections)

    def search_packaging_registration(self):
        query = self.entry_search_packaging_registration.get()
        selections = []
        for child in self.packaging_table.get_children():
            if query.capitalize() in self.packaging_table.item(child)['values'] or query.lower() in self.packaging_table.item(child)['values'] or query.upper() in self.packaging_table.item(child)['values'] or query.lower() in self.packaging_table.item(child)['values'] or query.swapcase() in self.packaging_table.item(child)['values'] or query.title() in self.packaging_table.item(child)['values']:
                selections.append(child)
        self.packaging_table.selection_set(selections)

    def search_product_registration(self):
        query = self.entry_search_product_registration.get()
        selections = []
        for child in self.product_table.get_children():
            if query.capitalize() in self.product_table.item(child)['values'] or query.lower() in self.product_table.item(child)['values'] or query.upper() in self.product_table.item(child)['values'] or query.lower() in self.product_table.item(child)['values'] or query.swapcase() in self.product_table.item(child)['values'] or query.title() in self.product_table.item(child)['values']:
                selections.append(child)
        self.product_table.selection_set(selections)

    def search_distribution_loadingList(self):
        query = self.entry_search_distribution.get()
        selections = []
        for child in self.loadingList_table.get_children():
            if query.capitalize() in self.loadingList_table.item(child)['values'] or query.lower() in self.loadingList_table.item(child)['values'] or query.upper() in self.loadingList_table.item(child)['values'] or query.lower() in self.loadingList_table.item(child)['values'] or query.swapcase() in self.loadingList_table.item(child)['values'] or query.title() in self.loadingList_table.item(child)['values']:
                selections.append(child)
        self.loadingList_table.selection_set(selections)

    def search_distribution_displanData(self):
        query = self.entry_search_distribution.get()
        selections = []
        for child in self.displanData_table.get_children():
            if query.capitalize() in self.displanData_table.item(child)['values'] or query.lower() in self.displanData_table.item(child)['values'] or query.upper() in self.displanData_table.item(child)['values'] or query.lower() in self.displanData_table.item(child)['values'] or query.swapcase() in self.displanData_table.item(child)['values'] or query.title() in self.displanData_table.item(child)['values']:
                selections.append(child)
        self.displanData_table.selection_set(selections)

    def search_tapping_point(self):
        query = self.entry_search_tapping_point.get()
        selections = []
        for child in self.tappingPoint_table.get_children():
            if query.capitalize() in self.tappingPoint_table.item(child)['values'] or query.lower() in self.tappingPoint_table.item(child)['values'] or query.upper() in self.tappingPoint_table.item(child)['values'] or query.lower() in self.tappingPoint_table.item(child)['values'] or query.swapcase() in self.tappingPoint_table.item(child)['values'] or query.title() in self.tappingPoint_table.item(child)['values']:
                selections.append(child)
        self.tappingPoint_table.selection_set(selections)

    def search_alert_history(self):
        # hapus item di tabel GUI
        for record in self.alert_note_history_tabel.get_children():
            self.alert_note_history_tabel.delete(record)

        date_start = self.alert_note_history_dateStart_calendar.get()
        date_start = date_start.split("/")
        date_end = self.alert_note_history_dateEnd_calendar.get()
        date_end = date_end.split("/")

        if int(date_start[1]) < 10 :
            tgl_start = "0" + date_start[1]
        else:
            tgl_start = date_start[1]
        if int(date_start[0]) < 10 :
            bln_start = "0" + date_start[0]
        else:
            bln_start = date_start[0]
        tahun_start = "20" + date_start[2]

        if int(date_end[1]) < 10 :
            tgl_end = "0" + date_end[1]
        else:
            tgl_end = date_end[1]
        if int(date_end[0]) < 10 :
            bln_end = "0" + date_end[0]
        else:
            bln_end = date_end[0]
        tahun_end = "20" + date_end[2]

        # delete all from tree
        for record in self.alert_note_history_tabel.get_children():
            self.alert_note_history_tabel.delete(record)

        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT * FROM tb_history")
        rec = c.fetchall()

        # input to tabel GUI
        count = 0
        for i in rec:
            date_history = i[0]
            date_history = date_history.split("/")
            if int(date_history[2]) >= int(tahun_start) and int(date_history[2]) <= int(tahun_end):
                if int(date_history[1]) >= int(bln_start) and int(date_history[1]) <= int(bln_end):
                    if int(date_history[0]) >= int(tgl_start) and int(date_history[0]) <= int(tgl_end):
                        if count % 2 == 0:
                            self.alert_note_history_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12]), tags=("evenrow",))
                        else:
                            self.alert_note_history_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12]), tags=("oddrow",))
                        count+=1
        file.commit()
        file.close()

    def search_report_total_time(self):
        # hapus item di tabel GUI
        for record in self.report_viewer_total_time_table.get_children():
            self.report_viewer_total_time_table.delete(record)

        transaction = self.report_viewer_total_time_optionmenu_transaction.get()
        packaging_total_time = self.report_viewer_total_time_optionmenu_packaging.get()
        product_total_time = self.report_viewer_total_time_optionmenu_product.get()
        datestart_total_time = self.report_viewer_total_time_dateStart_calendar.get()
        datestart_total_time = datestart_total_time.split("/")
        dateend_total_time = self.report_viewer_total_time_dateEnd_calendar.get()
        dateend_total_time = dateend_total_time.split("/")

        # date treatment
        if int(datestart_total_time[1]) < 10 :
            tgl_start = "0" + datestart_total_time[1]
        else:
            tgl_start = datestart_total_time[1]
        if int(datestart_total_time[0]) < 10 :
            bln_start = "0" + datestart_total_time[0]
        else:
            bln_start = datestart_total_time[0]
        tahun_start = "20" + datestart_total_time[2]

        if int(dateend_total_time[1]) < 10 :
            tgl_end = "0" + dateend_total_time[1]
        else:
            tgl_end = dateend_total_time[1]
        if int(dateend_total_time[0]) < 10 :
            bln_end = "0" + dateend_total_time[0]
        else:
            bln_end = dateend_total_time[0]
        tahun_end = "20" + dateend_total_time[2]


        if product_total_time != "-":
            if transaction == "ALL":
                # get data
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                c.execute("SELECT * FROM tb_queuing_history where packaging = ? AND product = ?",(packaging_total_time, product_total_time))
                rec = c.fetchall()

                count = 0
                for i in rec:
                    date_total_time = i[0]
                    date_total_time = date_total_time.split("/")
                    if int(date_total_time[2]) >= int(tahun_start) and int(date_total_time[2]) <= int(tahun_end):
                        if int(date_total_time[1]) >= int(bln_start) and int(date_total_time[1]) <= int(bln_end):
                            if int(bln_start) < int(date_total_time[1]):
                                if int(date_total_time[0]) <= int(tgl_end):
                                    if count % 2 == 0:
                                        self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("evenrow",))
                                    else:
                                        self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("oddrow",))
                                    count+=1
                            else:
                                if int(date_total_time[0]) >= int(tgl_start) and int(date_total_time[0]) <= int(tgl_end):
                                    if count % 2 == 0:
                                        self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("evenrow",))
                                    else:
                                        self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("oddrow",))
                                    count+=1

                file.commit()
                file.close()
            else:
                # get data
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                c.execute("SELECT * FROM tb_queuing_history where activity = ? AND packaging = ? AND product = ?",(transaction, packaging_total_time, product_total_time))
                rec = c.fetchall()

                count = 0
                for i in rec:
                    date_total_time = i[0]
                    date_total_time = date_total_time.split("/")
                    if int(date_total_time[2]) >= int(tahun_start) and int(date_total_time[2]) <= int(tahun_end):
                        if int(date_total_time[1]) >= int(bln_start) and int(date_total_time[1]) <= int(bln_end):
                            if int(bln_start) < int(date_total_time[1]):
                                if int(date_total_time[0]) <= int(tgl_end):
                                    if count % 2 == 0:
                                        self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("evenrow",))
                                    else:
                                        self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29]), tags=("oddrow",))
                                    count+=1
                            else:
                                if int(date_total_time[0]) >= int(tgl_start) and int(date_total_time[0]) <= int(tgl_end):
                                    if count % 2 == 0:
                                        self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("evenrow",))
                                    else:
                                        self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("oddrow",))
                                    count+=1

                file.commit()
                file.close()
        else:
            if transaction == "ALL":
                if packaging_total_time == "ALL":
                    # get data
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT * FROM tb_queuing_history")
                    rec = c.fetchall()

                    count = 0
                    for i in rec:
                        date_total_time = i[0]
                        date_total_time = date_total_time.split("/")
                        if int(date_total_time[2]) >= int(tahun_start) and int(date_total_time[2]) <= int(tahun_end):
                            if int(date_total_time[1]) >= int(bln_start) and int(date_total_time[1]) <= int(bln_end):
                                if int(bln_start) < int(date_total_time[1]):
                                    if int(date_total_time[0]) <= int(tgl_end):
                                        if count % 2 == 0:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("evenrow",))
                                        else:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("oddrow",))
                                        count+=1
                                else:
                                    if int(date_total_time[0]) >= int(tgl_start) and int(date_total_time[0]) <= int(tgl_end):
                                        if count % 2 == 0:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("evenrow",))
                                        else:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("oddrow",))
                                        count+=1

                    file.commit()
                    file.close()
                else:
                    # get data
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT * FROM tb_queuing_history where packaging = ?",(packaging_total_time))
                    rec = c.fetchall()

                    count = 0
                    for i in rec:
                        date_total_time = i[0]
                        date_total_time = date_total_time.split("/")
                        if int(date_total_time[2]) >= int(tahun_start) and int(date_total_time[2]) <= int(tahun_end):
                            if int(date_total_time[1]) >= int(bln_start) and int(date_total_time[1]) <= int(bln_end):
                                if int(bln_start) < int(date_total_time[1]):
                                    if int(date_total_time[0]) <= int(tgl_end):
                                        if count % 2 == 0:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("evenrow",))
                                        else:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("oddrow",))
                                        count+=1
                                else:
                                    if int(date_total_time[0]) >= int(tgl_start) and int(date_total_time[0]) <= int(tgl_end):
                                        if count % 2 == 0:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("evenrow",))
                                        else:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("oddrow",))
                                        count+=1

                    file.commit()
                    file.close()
            else:
                if packaging_total_time == "ALL":
                    # get data
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT * FROM tb_queuing_history where activity = ?",(transaction))
                    rec = c.fetchall()

                    count = 0
                    for i in rec:
                        date_total_time = i[0]
                        date_total_time = date_total_time.split("/")
                        if int(date_total_time[2]) >= int(tahun_start) and int(date_total_time[2]) <= int(tahun_end):
                            if int(date_total_time[1]) >= int(bln_start) and int(date_total_time[1]) <= int(bln_end):
                                if int(bln_start) < int(date_total_time[1]):
                                    if int(date_total_time[0]) <= int(tgl_end):
                                        if count % 2 == 0:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("evenrow",))
                                        else:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("oddrow",))
                                        count+=1
                                else:
                                    if int(date_total_time[0]) >= int(tgl_start) and int(date_total_time[0]) <= int(tgl_end):
                                        if count % 2 == 0:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("evenrow",))
                                        else:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("oddrow",))
                                        count+=1

                    file.commit()
                    file.close()
                else:
                    # get data
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT * FROM tb_queuing_history where activity = ? AND packaging = ?",(transaction, packaging_total_time))
                    rec = c.fetchall()

                    count = 0
                    for i in rec:
                        date_total_time = i[0]
                        date_total_time = date_total_time.split("/")
                        if int(date_total_time[2]) >= int(tahun_start) and int(date_total_time[2]) <= int(tahun_end):
                            if int(date_total_time[1]) >= int(bln_start) and int(date_total_time[1]) <= int(bln_end):
                                if int(bln_start) < int(date_total_time[1]):
                                    if int(date_total_time[0]) <= int(tgl_end):
                                        if count % 2 == 0:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("evenrow",))
                                        else:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("oddrow",))
                                        count+=1
                                else:
                                    if int(date_total_time[0]) >= int(tgl_start) and int(date_total_time[0]) <= int(tgl_end):
                                        if count % 2 == 0:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("evenrow",))
                                        else:
                                            self.report_viewer_total_time_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[21], i[3], i[4], i[5], i[22], i[6], i[7], i[23], i[24], i[25], i[26], i[27], i[18], i[19], i[28], i[29], i[30]), tags=("oddrow",))
                                        count+=1

                    file.commit()
                    file.close()
         
    def search_report_transaction_time(self):
        global data_product, data_avg_time

        # hapus item di tabel GUI
        for record in self.report_viewer_transaction_time_table.get_children():
            self.report_viewer_transaction_time_table.delete(record)

        transaction = self.report_viewer_transaction_time_optionmenu_transaction.get()
        packaging_transaction_time = self.report_viewer_transaction_time_optionmenu_packaging.get()
        product_transaction_time = self.report_viewer_transaction_time_optionmenu_product.get()
        tappoint_transaction_time = self.report_viewer_transaction_time_optionmenu_tapPoint.get()
        datestart_transaction_time = self.report_viewer_transaction_time_dateStart_calendar.get()
        datestart_transaction_time = datestart_transaction_time.split("/")
        dateend_transaction_time = self.report_viewer_transaction_time_dateEnd_calendar.get()
        dateend_transaction_time = dateend_transaction_time.split("/")

        # date treatment
        if int(datestart_transaction_time[1]) < 10 :
            tgl_start = "0" + datestart_transaction_time[1]
        else:
            tgl_start = datestart_transaction_time[1]
        if int(datestart_transaction_time[0]) < 10 :
            bln_start = "0" + datestart_transaction_time[0]
        else:
            bln_start = datestart_transaction_time[0]
        tahun_start = "20" + datestart_transaction_time[2]

        if int(dateend_transaction_time[1]) < 10 :
            tgl_end = "0" + dateend_transaction_time[1]
        else:
            tgl_end = dateend_transaction_time[1]
        if int(dateend_transaction_time[0]) < 10 :
            bln_end = "0" + dateend_transaction_time[0]
        else:
            bln_end = dateend_transaction_time[0]
        tahun_end = "20" + dateend_transaction_time[2]

        formatted_start_date = tgl_start+"/"+bln_start+"/"+tahun_start
        formatted_end_date = tgl_end+"/"+bln_end+"/"+tahun_end

        product_list = []
        avg_list = []

        if product_transaction_time != "-":
            if transaction == "ALL":
                if tappoint_transaction_time == "Waiting Time":
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()

                    # SQL query to fetch data from the database and group it
                    query = f"""
                            SELECT [packaging], [product], [activity], 
                                COUNT(*) AS TotalCount, 
                                AVG(CAST([waiting_minute] AS FLOAT)) AS AvgWaitingMinute 
                            FROM tb_queuing_history
                            WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}' AND [product]='{product_transaction_time}'
                            GROUP BY [packaging], [product], [activity]
                        """

                    # Execute the query and fetch the results
                    c.execute(query)
                    results = c.fetchall()

                    # Populate the Treeview with the fetched data
                    for row in results:
                        packaging, product, activity, total_count, avg_waiting_minute = row
                        self.report_viewer_transaction_time_table.insert("", "end", values=[
                            wh_name,  # Warehouse
                            activity,    # Transaction
                            packaging,   # Packaging
                            product,     # Product
                            f"{avg_waiting_minute:.2f}",  # AVG Times (Minute)
                            total_count  # TRN
                        ])

                        # Append product and avg_waiting_minute to the respective lists
                        product_list.append(product)
                        avg_list.append(avg_waiting_minute)

                    # Commit and close the database connection
                    file.commit()
                    file.close()
                if tappoint_transaction_time == "Weight Bridge 1":
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()

                    # SQL query to fetch data from the database and group it
                    query = f"""
                            SELECT [packaging], [product], [activity], 
                                COUNT(*) AS TotalCount, 
                                AVG(CAST([wb1_minute] AS FLOAT)) AS AvgWaitingMinute 
                            FROM tb_queuing_history
                            WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}' AND [product]='{product_transaction_time}'
                            GROUP BY [packaging], [product], [activity]
                        """

                    # Execute the query and fetch the results
                    c.execute(query)
                    results = c.fetchall()

                    # Populate the Treeview with the fetched data
                    for row in results:
                        packaging, product, activity, total_count, avg_waiting_minute = row
                        self.report_viewer_transaction_time_table.insert("", "end", values=[
                            wh_name,  # Warehouse
                            activity,    # Transaction
                            packaging,   # Packaging
                            product,     # Product
                            f"{avg_waiting_minute:.2f}",  # AVG Times (Minute)
                            total_count  # TRN
                        ])

                        # Append product and avg_waiting_minute to the respective lists
                        product_list.append(product)
                        avg_list.append(avg_waiting_minute)

                    # Commit and close the database connection
                    file.commit()
                    file.close()
                if tappoint_transaction_time == "Warehouse":
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()

                    # SQL query to fetch data from the database and group it
                    query = f"""
                            SELECT [packaging], [product], [activity], 
                                COUNT(*) AS TotalCount, 
                                AVG(CAST([wh_minute] AS FLOAT)) AS AvgWaitingMinute 
                            FROM tb_queuing_history
                            WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}' AND [product]='{product_transaction_time}'
                            GROUP BY [packaging], [product], [activity]
                        """

                    # Execute the query and fetch the results
                    c.execute(query)
                    results = c.fetchall()

                    # Populate the Treeview with the fetched data
                    for row in results:
                        packaging, product, activity, total_count, avg_waiting_minute = row
                        self.report_viewer_transaction_time_table.insert("", "end", values=[
                            wh_name,  # Warehouse
                            activity,    # Transaction
                            packaging,   # Packaging
                            product,     # Product
                            f"{avg_waiting_minute:.2f}",  # AVG Times (Minute)
                            total_count  # TRN
                        ])

                        # Append product and avg_waiting_minute to the respective lists
                        product_list.append(product)
                        avg_list.append(avg_waiting_minute)

                    # Commit and close the database connection
                    file.commit()
                    file.close()
                if tappoint_transaction_time == "Weight Bridge 2":
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()

                    # SQL query to fetch data from the database and group it
                    query = f"""
                            SELECT [packaging], [product], [activity], 
                                COUNT(*) AS TotalCount, 
                                AVG(CAST([wb2_minute] AS FLOAT)) AS AvgWaitingMinute 
                            FROM tb_queuing_history
                            WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}' AND [product]='{product_transaction_time}'
                            GROUP BY [packaging], [product], [activity]
                        """

                    # Execute the query and fetch the results
                    c.execute(query)
                    results = c.fetchall()

                    # Populate the Treeview with the fetched data
                    for row in results:
                        packaging, product, activity, total_count, avg_waiting_minute = row
                        self.report_viewer_transaction_time_table.insert("", "end", values=[
                            wh_name,  # Warehouse
                            activity,    # Transaction
                            packaging,   # Packaging
                            product,     # Product
                            f"{avg_waiting_minute:.2f}",  # AVG Times (Minute)
                            total_count  # TRN
                        ])

                        # Append product and avg_waiting_minute to the respective lists
                        product_list.append(product)
                        avg_list.append(avg_waiting_minute)

                    # Commit and close the database connection
                    file.commit()
                    file.close()
                if tappoint_transaction_time == "Cargo Covering":
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()

                    # SQL query to fetch data from the database and group it
                    query = f"""
                            SELECT [packaging], [product], [activity], 
                                COUNT(*) AS TotalCount, 
                                AVG(CAST([cc_minute] AS FLOAT)) AS AvgWaitingMinute 
                            FROM tb_queuing_history
                            WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}' AND [product]='{product_transaction_time}'
                            GROUP BY [packaging], [product], [activity]
                        """

                    # Execute the query and fetch the results
                    c.execute(query)
                    results = c.fetchall()

                    # Populate the Treeview with the fetched data
                    for row in results:
                        packaging, product, activity, total_count, avg_waiting_minute = row
                        self.report_viewer_transaction_time_table.insert("", "end", values=[
                            wh_name,  # Warehouse
                            activity,    # Transaction
                            packaging,   # Packaging
                            product,     # Product
                            f"{avg_waiting_minute:.2f}",  # AVG Times (Minute)
                            total_count  # TRN
                        ])

                        # Append product and avg_waiting_minute to the respective lists
                        product_list.append(product)
                        avg_list.append(avg_waiting_minute)

                    # Commit and close the database connection
                    file.commit()
                    file.close()
            else:
                if tappoint_transaction_time == "Waiting Time":
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()

                    # SQL query to fetch data from the database and group it
                    query = f"""
                            SELECT [packaging], [product], [activity], 
                                COUNT(*) AS TotalCount, 
                                AVG(CAST([waiting_minute] AS FLOAT)) AS AvgWaitingMinute 
                            FROM tb_queuing_history
                            WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}' AND [product]='{product_transaction_time}' AND [activity]='{transaction}'
                            GROUP BY [packaging], [product], [activity]
                        """

                    # Execute the query and fetch the results
                    c.execute(query)
                    results = c.fetchall()

                    # Populate the Treeview with the fetched data
                    for row in results:
                        packaging, product, activity, total_count, avg_waiting_minute = row
                        self.report_viewer_transaction_time_table.insert("", "end", values=[
                            wh_name,  # Warehouse
                            activity,    # Transaction
                            packaging,   # Packaging
                            product,     # Product
                            f"{avg_waiting_minute:.2f}",  # AVG Times (Minute)
                            total_count  # TRN
                        ])

                        # Append product and avg_waiting_minute to the respective lists
                        product_list.append(product)
                        avg_list.append(avg_waiting_minute)

                    # Commit and close the database connection
                    file.commit()
                    file.close()
                if tappoint_transaction_time == "Weight Bridge 1":
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()

                    # SQL query to fetch data from the database and group it
                    query = f"""
                            SELECT [packaging], [product], [activity], 
                                COUNT(*) AS TotalCount, 
                                AVG(CAST([wb1_minute] AS FLOAT)) AS AvgWaitingMinute 
                            FROM tb_queuing_history
                            WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}' AND [product]='{product_transaction_time}' AND [activity]='{transaction}'
                            GROUP BY [packaging], [product], [activity]
                        """

                    # Execute the query and fetch the results
                    c.execute(query)
                    results = c.fetchall()

                    # Populate the Treeview with the fetched data
                    for row in results:
                        packaging, product, activity, total_count, avg_waiting_minute = row
                        self.report_viewer_transaction_time_table.insert("", "end", values=[
                            wh_name,  # Warehouse
                            activity,    # Transaction
                            packaging,   # Packaging
                            product,     # Product
                            f"{avg_waiting_minute:.2f}",  # AVG Times (Minute)
                            total_count  # TRN
                        ])

                        # Append product and avg_waiting_minute to the respective lists
                        product_list.append(product)
                        avg_list.append(avg_waiting_minute)

                    # Commit and close the database connection
                    file.commit()
                    file.close()
                if tappoint_transaction_time == "Warehouse":
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()

                    # SQL query to fetch data from the database and group it
                    query = f"""
                            SELECT [packaging], [product], [activity], 
                                COUNT(*) AS TotalCount, 
                                AVG(CAST([wh_minute] AS FLOAT)) AS AvgWaitingMinute 
                            FROM tb_queuing_history
                            WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}' AND [product]='{product_transaction_time}' AND [activity]='{transaction}'
                            GROUP BY [packaging], [product], [activity]
                        """

                    # Execute the query and fetch the results
                    c.execute(query)
                    results = c.fetchall()

                    # Populate the Treeview with the fetched data
                    for row in results:
                        packaging, product, activity, total_count, avg_waiting_minute = row
                        self.report_viewer_transaction_time_table.insert("", "end", values=[
                            wh_name,  # Warehouse
                            activity,    # Transaction
                            packaging,   # Packaging
                            product,     # Product
                            f"{avg_waiting_minute:.2f}",  # AVG Times (Minute)
                            total_count  # TRN
                        ])

                        # Append product and avg_waiting_minute to the respective lists
                        product_list.append(product)
                        avg_list.append(avg_waiting_minute)

                    # Commit and close the database connection
                    file.commit()
                    file.close()
                if tappoint_transaction_time == "Weight Bridge 2":
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()

                    # SQL query to fetch data from the database and group it
                    query = f"""
                            SELECT [packaging], [product], [activity], 
                                COUNT(*) AS TotalCount, 
                                AVG(CAST([wb2_minute] AS FLOAT)) AS AvgWaitingMinute 
                            FROM tb_queuing_history
                            WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}' AND [product]='{product_transaction_time}' AND [activity]='{transaction}'
                            GROUP BY [packaging], [product], [activity]
                        """

                    # Execute the query and fetch the results
                    c.execute(query)
                    results = c.fetchall()

                    # Populate the Treeview with the fetched data
                    for row in results:
                        packaging, product, activity, total_count, avg_waiting_minute = row
                        self.report_viewer_transaction_time_table.insert("", "end", values=[
                            wh_name,  # Warehouse
                            activity,    # Transaction
                            packaging,   # Packaging
                            product,     # Product
                            f"{avg_waiting_minute:.2f}",  # AVG Times (Minute)
                            total_count  # TRN
                        ])

                        # Append product and avg_waiting_minute to the respective lists
                        product_list.append(product)
                        avg_list.append(avg_waiting_minute)

                    # Commit and close the database connection
                    file.commit()
                    file.close()
                if tappoint_transaction_time == "Cargo Covering":
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()

                    # SQL query to fetch data from the database and group it
                    query = f"""
                            SELECT [packaging], [product], [activity], 
                                COUNT(*) AS TotalCount, 
                                AVG(CAST([cc_minute] AS FLOAT)) AS AvgWaitingMinute 
                            FROM tb_queuing_history
                            WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}' AND [product]='{product_transaction_time}' AND [activity]='{transaction}'
                            GROUP BY [packaging], [product], [activity]
                        """

                    # Execute the query and fetch the results
                    c.execute(query)
                    results = c.fetchall()

                    # Populate the Treeview with the fetched data
                    for row in results:
                        packaging, product, activity, total_count, avg_waiting_minute = row
                        self.report_viewer_transaction_time_table.insert("", "end", values=[
                            wh_name,  # Warehouse
                            activity,    # Transaction
                            packaging,   # Packaging
                            product,     # Product
                            f"{avg_waiting_minute:.2f}",  # AVG Times (Minute)
                            total_count  # TRN
                        ])

                        # Append product and avg_waiting_minute to the respective lists
                        product_list.append(product)
                        avg_list.append(avg_waiting_minute)

                    # Commit and close the database connection
                    file.commit()
                    file.close()
        else:
            if transaction == "ALL":
                if packaging_transaction_time == "ALL":
                    if tappoint_transaction_time == "Waiting Time":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([waiting_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103)
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_waiting_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_waiting_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_waiting_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_waiting_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Weight Bridge 1":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([wb1_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103)
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Warehouse":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([wh_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103)
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Weight Bridge 2":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([wb2_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103)
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Cargo Covering":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([cc_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103)
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                else:
                    if tappoint_transaction_time == "Waiting Time":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([waiting_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}'
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Weight Bridge 1":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([wb1_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}'
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Warehouse":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([wh_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}'
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Weight Bridge 2":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([wb2_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}'
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Cargo Covering":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([cc_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}'
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
            else:
                if packaging_transaction_time == "ALL":
                    if tappoint_transaction_time == "Waiting Time":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([waiting_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [activity]='{transaction}'
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Weight Bridge 1":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([wb1_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [activity]='{transaction}'
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Warehouse":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([wh_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [activity]='{transaction}'
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Weight Bridge 2":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([wb2_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [activity]='{transaction}'
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Cargo Covering":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([cc_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [activity]='{transaction}'
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                else:
                    if tappoint_transaction_time == "Waiting Time":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([waiting_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}' AND [activity]='{transaction}'
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Weight Bridge 1":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([wb1_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}' AND [activity]='{transaction}'
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Warehouse":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([wh_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}' AND [activity]='{transaction}'
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Weight Bridge 2":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([wb2_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}' AND [activity]='{transaction}'
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
                    if tappoint_transaction_time == "Cargo Covering":
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()

                        # SQL query to fetch data from the database and group it
                        query = f"""
                                SELECT [packaging], [product], [activity], 
                                    COUNT(*) AS TotalCount, 
                                    AVG(CAST([cc_minute] AS FLOAT)) AS AvgWaitingMinute 
                                FROM tb_queuing_history
                                WHERE CONVERT(DATE, [date], 103) BETWEEN CONVERT(DATE, '{formatted_start_date}', 103) AND CONVERT(DATE, '{formatted_end_date}', 103) AND [packaging]='{packaging_transaction_time}' AND [activity]='{transaction}'
                                GROUP BY [packaging], [product], [activity]
                            """

                        # Execute the query and fetch the results
                        c.execute(query)
                        results = c.fetchall()

                        # Populate the Treeview with the fetched data
                        for row in results:
                            packaging, product, activity, total_count, avg_minute = row
                            self.report_viewer_transaction_time_table.insert("", "end", values=[
                                wh_name,  # Warehouse
                                activity,    # Transaction
                                packaging,   # Packaging
                                product,     # Product
                                f"{avg_minute:.2f}",  # AVG Times (Minute)
                                total_count  # TRN
                            ])

                            # Append product and avg_minute to the respective lists
                            product_list.append(product)
                            avg_list.append(avg_minute)

                        # Commit and close the database connection
                        file.commit()
                        file.close()
        
        self.ax1.cla()
        self.ax1.bar(product_list, avg_list, color ='maroon', width = 0.4)
        self.ax1.set_xlabel("Product")
        self.ax1.set_ylabel("AVG Time")
        self.ax1.set_title("Average Transaction Time")
        # Rotate the tick labels vertically
        self.ax1.set_xticks(product_list)
        self.ax1.set_xticklabels(product_list, rotation=20)
        self.bar_data_transaction_time.draw()
        self.bar_data_transaction_time.get_tk_widget().configure()

    def search_report_peak_hour(self):
        global data_arrival_time, data_total_arrival_time

        # hapus item di tabel GUI
        for record in self.report_viewer_peak_hour_table.get_children():
            self.report_viewer_peak_hour_table.delete(record)

        transaction = self.report_viewer_peak_hour_optionmenu_transaction.get()
        packaging_peak_hour = self.report_viewer_peak_hour_optionmenu_packaging.get()
        product_peak_hour = self.report_viewer_peak_hour_optionmenu_product.get()
        datestart_peak_hour = self.report_viewer_peak_hour_dateStart_calendar.get()
        datestart_peak_hour = datestart_peak_hour.split("/")
        dateend_peak_hour = self.report_viewer_peak_hour_dateEnd_calendar.get()
        dateend_peak_hour = dateend_peak_hour.split("/")

        # date treatment
        if int(datestart_peak_hour[1]) < 10 :
            tgl_start = "0" + datestart_peak_hour[1]
        else:
            tgl_start = datestart_peak_hour[1]
        if int(datestart_peak_hour[0]) < 10 :
            bln_start = "0" + datestart_peak_hour[0]
        else:
            bln_start = datestart_peak_hour[0]
        tahun_start = "20" + datestart_peak_hour[2]

        if int(dateend_peak_hour[1]) < 10 :
            tgl_end = "0" + dateend_peak_hour[1]
        else:
            tgl_end = dateend_peak_hour[1]
        if int(dateend_peak_hour[0]) < 10 :
            bln_end = "0" + dateend_peak_hour[0]
        else:
            bln_end = dateend_peak_hour[0]
        tahun_end = "20" + dateend_peak_hour[2]

        data_arrival_time[:] = []
        data_total_arrival_time[:] = []
        count = 0
        dict_data = {}

        if product_peak_hour != "-":
            if transaction == "ALL":
                # get data
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                c.execute("SELECT date, arrival FROM tb_queuing_history where packaging = ? AND product = ?",(packaging_peak_hour, product_peak_hour))
                rec = c.fetchall()
                file.commit()
                file.close()
                for i in rec:
                    date_peak_hour = i[0]
                    date_peak_hour = date_peak_hour.split("/")
                    if int(date_peak_hour[2]) >= int(tahun_start) and int(date_peak_hour[2]) <= int(tahun_end):
                        if int(date_peak_hour[1]) >= int(bln_start) and int(date_peak_hour[1]) <= int(bln_end):
                            if int(bln_start) < int(date_peak_hour[1]):
                                if int(date_peak_hour[0]) <= int(tgl_end):
                                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                    c = file.cursor()
                                    c.execute("SELECT arrival FROM tb_queuing_history where date = ? AND packaging = ? AND product = ?",(i[0], packaging_peak_hour, product_peak_hour))
                                    rec_arrival = c.fetchall()
                                    file.commit()
                                    file.close()
                                    total = 0
                                    for j in rec_arrival:
                                        if j[0] == i[1]:
                                            arrival_time = i[1][0:2] + ".00"
                                            total = total + 1
                                            try:
                                                value = dict_data[arrival_time]
                                                total = total + value
                                                new_dict_data = {arrival_time : total}
                                                dict_data.update(new_dict_data)
                                                new_dict_data.clear()
                                            except:
                                                new_dict_data = {arrival_time : total}
                                                dict_data.update(new_dict_data)
                                                new_dict_data.clear()
                            else:
                                if int(date_peak_hour[0]) >= int(tgl_start) and int(date_peak_hour[0]) <= int(tgl_end):
                                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                    c = file.cursor()
                                    c.execute("SELECT arrival FROM tb_queuing_history where date = ? AND packaging = ? AND product = ?",(i[0], packaging_peak_hour, product_peak_hour))
                                    rec_arrival = c.fetchall()
                                    file.commit()
                                    file.close()
                                    total = 0
                                    for j in rec_arrival:
                                        if j[0] == i[1]:
                                            arrival_time = i[1][0:2] + ".00"
                                            total = total + 1
                                            try:
                                                value = dict_data[arrival_time]
                                                total = total + value
                                                new_dict_data = {arrival_time : total}
                                                dict_data.update(new_dict_data)
                                                new_dict_data.clear()
                                            except:
                                                new_dict_data = {arrival_time : total}
                                                dict_data.update(new_dict_data)
                                                new_dict_data.clear()
            else:
                # get data
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                c.execute("SELECT date, arrival FROM tb_queuing_history where packaging = ? AND product = ? AND activity = ?",(packaging_peak_hour, product_peak_hour, transaction))
                rec = c.fetchall()
                file.commit()
                file.close()
                for i in rec:
                    date_peak_hour = i[0]
                    date_peak_hour = date_peak_hour.split("/")
                    if int(date_peak_hour[2]) >= int(tahun_start) and int(date_peak_hour[2]) <= int(tahun_end):
                        if int(date_peak_hour[1]) >= int(bln_start) and int(date_peak_hour[1]) <= int(bln_end):
                            if int(bln_start) < int(date_peak_hour[1]):
                                if int(date_peak_hour[0]) <= int(tgl_end):
                                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                    c = file.cursor()
                                    c.execute("SELECT arrival FROM tb_queuing_history where date = ? AND packaging = ? AND product = ? AND activity = ?",(i[0], packaging_peak_hour, product_peak_hour, transaction))
                                    rec_arrival = c.fetchall()
                                    file.commit()
                                    file.close()
                                    total = 0
                                    for j in rec_arrival:
                                        if j[0] == i[1]:
                                            arrival_time = i[1][0:2] + ".00"
                                            total = total + 1
                                            try:
                                                value = dict_data[arrival_time]
                                                total = total + value
                                                new_dict_data = {arrival_time : total}
                                                dict_data.update(new_dict_data)
                                                new_dict_data.clear()
                                            except:
                                                new_dict_data = {arrival_time : total}
                                                dict_data.update(new_dict_data)
                                                new_dict_data.clear()
                            else:
                                if int(date_peak_hour[0]) >= int(tgl_start) and int(date_peak_hour[0]) <= int(tgl_end):
                                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                    c = file.cursor()
                                    c.execute("SELECT arrival FROM tb_queuing_history where date = ? AND packaging = ? AND product = ? AND activity = ?",(i[0], packaging_peak_hour, product_peak_hour, transaction))
                                    rec_arrival = c.fetchall()
                                    file.commit()
                                    file.close()
                                    total = 0
                                    for j in rec_arrival:
                                        if j[0] == i[1]:
                                            arrival_time = i[1][0:2] + ".00"
                                            total = total + 1
                                            try:
                                                value = dict_data[arrival_time]
                                                total = total + value
                                                new_dict_data = {arrival_time : total}
                                                dict_data.update(new_dict_data)
                                                new_dict_data.clear()
                                            except:
                                                new_dict_data = {arrival_time : total}
                                                dict_data.update(new_dict_data)
                                                new_dict_data.clear()
        else:
            if transaction == "ALL":
                if packaging_peak_hour == "ALL":
                     # get data
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT date, arrival FROM tb_queuing_history")
                    rec = c.fetchall()
                    file.commit()
                    file.close()
                    for i in rec:
                        date_peak_hour = i[0]
                        date_peak_hour = date_peak_hour.split("/")
                        if int(date_peak_hour[2]) >= int(tahun_start) and int(date_peak_hour[2]) <= int(tahun_end):
                            if int(date_peak_hour[1]) >= int(bln_start) and int(date_peak_hour[1]) <= int(bln_end):
                                if int(bln_start) < int(date_peak_hour[1]):
                                    if int(date_peak_hour[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT arrival FROM tb_queuing_history where date = ?",(i[0]))
                                        rec_arrival = c.fetchall()
                                        file.commit()
                                        file.close()
                                        total = 0
                                        for j in rec_arrival:
                                            if j[0] == i[1]:
                                                arrival_time = i[1][0:2] + ".00"
                                                total = total + 1
                                                try:
                                                    value = dict_data[arrival_time]
                                                    total = total + value
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()
                                                except:
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()
                                else:
                                    if int(date_peak_hour[0]) >= int(tgl_start) and int(date_peak_hour[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT arrival FROM tb_queuing_history where date = ?",(i[0]))
                                        rec_arrival = c.fetchall()
                                        file.commit()
                                        file.close()
                                        total = 0
                                        for j in rec_arrival:
                                            if j[0] == i[1]:
                                                arrival_time = i[1][0:2] + ".00"
                                                total = total + 1
                                                try:
                                                    value = dict_data[arrival_time]
                                                    total = total + value
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()
                                                except:
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()
                else:
                    # get data
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT date, arrival FROM tb_queuing_history where packaging = ?",(packaging_peak_hour,))
                    rec = c.fetchall()
                    file.commit()
                    file.close()
                    for i in rec:
                        date_peak_hour = i[0]
                        date_peak_hour = date_peak_hour.split("/")
                        if int(date_peak_hour[2]) >= int(tahun_start) and int(date_peak_hour[2]) <= int(tahun_end):
                            if int(date_peak_hour[1]) >= int(bln_start) and int(date_peak_hour[1]) <= int(bln_end):
                                if int(bln_start) < int(date_peak_hour[1]):
                                    if int(date_peak_hour[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT arrival FROM tb_queuing_history where date = ? AND packaging = ?",(i[0], packaging_peak_hour))
                                        rec_arrival = c.fetchall()
                                        file.commit()
                                        file.close()
                                        total = 0
                                        for j in rec_arrival:
                                            if j[0] == i[1]:
                                                arrival_time = i[1][0:2] + ".00"
                                                total = total + 1
                                                try:
                                                    value = dict_data[arrival_time]
                                                    total = total + value
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()
                                                except:
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()
                                else:
                                    if int(date_peak_hour[0]) >= int(tgl_start) and int(date_peak_hour[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT arrival FROM tb_queuing_history where date = ? AND packaging = ?",(i[0], packaging_peak_hour))
                                        rec_arrival = c.fetchall()
                                        file.commit()
                                        file.close()
                                        total = 0
                                        for j in rec_arrival:
                                            if j[0] == i[1]:
                                                arrival_time = i[1][0:2] + ".00"
                                                total = total + 1
                                                try:
                                                    value = dict_data[arrival_time]
                                                    total = total + value
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()
                                                except:
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()
            else:
                if packaging_peak_hour == "ALL":
                    # get data
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT date, arrival FROM tb_queuing_history where activity = ?",(transaction))
                    rec = c.fetchall()
                    file.commit()
                    file.close()
                    for i in rec:
                        date_peak_hour = i[0]
                        date_peak_hour = date_peak_hour.split("/")
                        if int(date_peak_hour[2]) >= int(tahun_start) and int(date_peak_hour[2]) <= int(tahun_end):
                            if int(date_peak_hour[1]) >= int(bln_start) and int(date_peak_hour[1]) <= int(bln_end):
                                if int(bln_start) < int(date_peak_hour[1]):
                                    if int(date_peak_hour[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT arrival FROM tb_queuing_history where date = ? AND activity = ?",(i[0], transaction))
                                        rec_arrival = c.fetchall()
                                        file.commit()
                                        file.close()
                                        total = 0
                                        for j in rec_arrival:
                                            if j[0] == i[1]:
                                                arrival_time = i[1][0:2] + ".00"
                                                total = total + 1
                                                try:
                                                    value = dict_data[arrival_time]
                                                    total = total + value
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()
                                                except:
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()
                                else:
                                    if int(date_peak_hour[0]) >= int(tgl_start) and int(date_peak_hour[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT arrival FROM tb_queuing_history where date = ? AND activity = ?",(i[0], transaction))
                                        rec_arrival = c.fetchall()
                                        file.commit()
                                        file.close()
                                        total = 0
                                        for j in rec_arrival:
                                            if j[0] == i[1]:
                                                arrival_time = i[1][0:2] + ".00"
                                                total = total + 1
                                                try:
                                                    value = dict_data[arrival_time]
                                                    total = total + value
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()
                                                except:
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()
                else:
                    # get data
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT date, arrival FROM tb_queuing_history where packaging = ? AND activity = ?",(packaging_peak_hour, transaction))
                    rec = c.fetchall()
                    file.commit()
                    file.close()
                    for i in rec:
                        date_peak_hour = i[0]
                        date_peak_hour = date_peak_hour.split("/")
                        if int(date_peak_hour[2]) >= int(tahun_start) and int(date_peak_hour[2]) <= int(tahun_end):
                            if int(date_peak_hour[1]) >= int(bln_start) and int(date_peak_hour[1]) <= int(bln_end):
                                if int(bln_start) < int(date_peak_hour[1]):
                                    if int(date_peak_hour[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT arrival FROM tb_queuing_history where date = ? AND packaging = ? AND activity = ?",(i[0], packaging_peak_hour, transaction))
                                        rec_arrival = c.fetchall()
                                        file.commit()
                                        file.close()
                                        total = 0
                                        for j in rec_arrival:
                                            if j[0] == i[1]:
                                                arrival_time = i[1][0:2] + ".00"
                                                total = total + 1
                                                try:
                                                    value = dict_data[arrival_time]
                                                    total = total + value
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()
                                                except:
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()
                                else:
                                    if int(date_peak_hour[0]) >= int(tgl_start) and int(date_peak_hour[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT arrival FROM tb_queuing_history where date = ? AND packaging = ? AND activity = ?",(i[0], packaging_peak_hour, transaction))
                                        rec_arrival = c.fetchall()
                                        file.commit()
                                        file.close()
                                        total = 0
                                        for j in rec_arrival:
                                            if j[0] == i[1]:
                                                arrival_time = i[1][0:2] + ".00"
                                                total = total + 1
                                                try:
                                                    value = dict_data[arrival_time]
                                                    total = total + value
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()
                                                except:
                                                    new_dict_data = {arrival_time : total}
                                                    dict_data.update(new_dict_data)
                                                    new_dict_data.clear()

        # data
        dict_data = dict(sorted(dict_data.items()))
        for i in dict_data:
            if count % 2 == 0:
                self.report_viewer_peak_hour_table.insert(parent="",index="end", iid=count, text="", values=(wh_name, transaction, i, dict_data[i]), tags=("evenrow",))
            else:
                self.report_viewer_peak_hour_table.insert(parent="",index="end", iid=count, text="", values=(wh_name, transaction, i, dict_data[i]), tags=("oddrow",))
            count+=1
        data_arrival_time = list(dict_data.keys())
        data_total_arrival_time = list(dict_data.values())
        self.ax2.cla()
        self.ax2.plot(data_arrival_time, data_total_arrival_time, color ='blue', marker='o')
        self.ax2.set_xlabel("Arrival Time")
        self.ax2.set_ylabel("Total")
        self.ax2.set_title("Peak Hour")
        self.plt_data_peak_hour.draw()
        self.plt_data_peak_hour.get_tk_widget().configure()

    def date_key(self,date_str):
        return datetime.strptime(date_str, '%d/%m/%Y')
    def search_report_total_queuing(self):
        global data_date, data_total_queuing

        # hapus item di tabel GUI
        for record in self.report_viewer_total_queuing_table.get_children():
            self.report_viewer_total_queuing_table.delete(record)

        transaction = self.report_viewer_total_queuing_optionmenu_transaction.get()
        packaging_total_queuing = self.report_viewer_total_queuing_optionmenu_packaging.get()
        product_total_queuing = self.report_viewer_total_queuing_optionmenu_product.get()
        datestart_total_queuing = self.report_viewer_total_queuing_dateStart_calendar.get()
        datestart_total_queuing = datestart_total_queuing.split("/")
        dateend_total_queuing = self.report_viewer_total_queuing_dateEnd_calendar.get()
        dateend_total_queuing = dateend_total_queuing.split("/")

        # date treatment
        if int(datestart_total_queuing[1]) < 10 :
            tgl_start = "0" + datestart_total_queuing[1]
        else:
            tgl_start = datestart_total_queuing[1]
        if int(datestart_total_queuing[0]) < 10 :
            bln_start = "0" + datestart_total_queuing[0]
        else:
            bln_start = datestart_total_queuing[0]
        tahun_start = "20" + datestart_total_queuing[2]

        if int(dateend_total_queuing[1]) < 10 :
            tgl_end = "0" + dateend_total_queuing[1]
        else:
            tgl_end = dateend_total_queuing[1]
        if int(dateend_total_queuing[0]) < 10 :
            bln_end = "0" + dateend_total_queuing[0]
        else:
            bln_end = dateend_total_queuing[0]
        tahun_end = "20" + dateend_total_queuing[2]

        data_date[:] = []
        data_total_queuing[:] = []
        count = 0
        dict_data = {}

        if product_total_queuing != "-":
            if transaction == "ALL":
                # get data
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                c.execute("SELECT date FROM tb_queuing_history where packaging = ? AND product = ?",(packaging_total_queuing, product_total_queuing))
                rec = c.fetchall()
                file.commit()
                file.close()
                for i in rec:
                    date_total_queuing = i[0]
                    date_total_queuing = date_total_queuing.split("/")
                    total = 0
                    if int(date_total_queuing[2]) >= int(tahun_start) and int(date_total_queuing[2]) <= int(tahun_end):
                        if int(date_total_queuing[1]) >= int(bln_start) and int(date_total_queuing[1]) <= int(bln_end):
                            if int(bln_start) < int(date_total_queuing[1]):
                                if int(date_total_queuing[0]) <= int(tgl_end):
                                    total = total + 1
                                    try:
                                        value = dict_data[i[0]]
                                        total = total + value
                                        new_dict_data = {i[0] : total}
                                        dict_data.update(new_dict_data)
                                        new_dict_data.clear()
                                    except:
                                        new_dict_data = {i[0] : total}
                                        dict_data.update(new_dict_data)
                                        new_dict_data.clear()
                            else:
                                if int(date_total_queuing[0]) >= int(tgl_start) and int(date_total_queuing[0]) <= int(tgl_end):
                                    total = total + 1
                                    try:
                                        value = dict_data[i[0]]
                                        total = total + value
                                        new_dict_data = {i[0] : total}
                                        dict_data.update(new_dict_data)
                                        new_dict_data.clear()
                                    except:
                                        new_dict_data = {i[0] : total}
                                        dict_data.update(new_dict_data)
                                        new_dict_data.clear()
            else:
                # get data
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                c.execute("SELECT date FROM tb_queuing_history where packaging = ? AND product = ? AND activity = ?",(packaging_total_queuing, product_total_queuing, transaction))
                rec = c.fetchall()
                file.commit()
                file.close()
                for i in rec:
                    date_total_queuing = i[0]
                    date_total_queuing = date_total_queuing.split("/")
                    total = 0
                    if int(date_total_queuing[2]) >= int(tahun_start) and int(date_total_queuing[2]) <= int(tahun_end):
                        if int(date_total_queuing[1]) >= int(bln_start) and int(date_total_queuing[1]) <= int(bln_end):
                            if int(bln_start) < int(date_total_queuing[1]):
                                if int(date_total_queuing[0]) <= int(tgl_end):
                                    total = total + 1
                                    try:
                                        value = dict_data[i[0]]
                                        total = total + value
                                        new_dict_data = {i[0] : total}
                                        dict_data.update(new_dict_data)
                                        new_dict_data.clear()
                                    except:
                                        new_dict_data = {i[0] : total}
                                        dict_data.update(new_dict_data)
                                        new_dict_data.clear()
                            else:
                                if int(date_total_queuing[0]) >= int(tgl_start) and int(date_total_queuing[0]) <= int(tgl_end):
                                    total = total + 1
                                    try:
                                        value = dict_data[i[0]]
                                        total = total + value
                                        new_dict_data = {i[0] : total}
                                        dict_data.update(new_dict_data)
                                        new_dict_data.clear()
                                    except:
                                        new_dict_data = {i[0] : total}
                                        dict_data.update(new_dict_data)
                                        new_dict_data.clear()
        else:
            if transaction == "ALL":
                if packaging_total_queuing == "ALL":
                    # get data
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT date FROM tb_queuing_history")
                    rec = c.fetchall()
                    file.commit()
                    file.close()
                    for i in rec:
                        date_total_queuing = i[0]
                        date_total_queuing = date_total_queuing.split("/")
                        total = 0
                        if int(date_total_queuing[2]) >= int(tahun_start) and int(date_total_queuing[2]) <= int(tahun_end):
                            if int(date_total_queuing[1]) >= int(bln_start) and int(date_total_queuing[1]) <= int(bln_end):
                                if int(bln_start) < int(date_total_queuing[1]):
                                    if int(date_total_queuing[0]) <= int(tgl_end):
                                        total = total + 1
                                        try:
                                            value = dict_data[i[0]]
                                            total = total + value
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()
                                        except:
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()
                                else:
                                    if int(date_total_queuing[0]) >= int(tgl_start) and int(date_total_queuing[0]) <= int(tgl_end):
                                        total = total + 1
                                        try:
                                            value = dict_data[i[0]]
                                            total = total + value
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()
                                        except:
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()
                else:
                    # get data
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT date FROM tb_queuing_history where packaging = ?",(packaging_total_queuing))
                    rec = c.fetchall()
                    file.commit()
                    file.close()
                    for i in rec:
                        date_total_queuing = i[0]
                        date_total_queuing = date_total_queuing.split("/")
                        total = 0
                        if int(date_total_queuing[2]) >= int(tahun_start) and int(date_total_queuing[2]) <= int(tahun_end):
                            if int(date_total_queuing[1]) >= int(bln_start) and int(date_total_queuing[1]) <= int(bln_end):
                                if int(bln_start) < int(date_total_queuing[1]):
                                    if int(date_total_queuing[0]) <= int(tgl_end):
                                        total = total + 1
                                        try:
                                            value = dict_data[i[0]]
                                            total = total + value
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()
                                        except:
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()
                                else:
                                    if int(date_total_queuing[0]) >= int(tgl_start) and int(date_total_queuing[0]) <= int(tgl_end):
                                        total = total + 1
                                        try:
                                            value = dict_data[i[0]]
                                            total = total + value
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()
                                        except:
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()
            else:
                if packaging_total_queuing == "ALL":
                    # get data
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT date FROM tb_queuing_history where activity = ?",(transaction))
                    rec = c.fetchall()
                    file.commit()
                    file.close()
                    for i in rec:
                        date_total_queuing = i[0]
                        date_total_queuing = date_total_queuing.split("/")
                        total = 0
                        if int(date_total_queuing[2]) >= int(tahun_start) and int(date_total_queuing[2]) <= int(tahun_end):
                            if int(date_total_queuing[1]) >= int(bln_start) and int(date_total_queuing[1]) <= int(bln_end):
                                if int(bln_start) < int(date_total_queuing[1]):
                                    if int(date_total_queuing[0]) <= int(tgl_end):
                                        total = total + 1
                                        try:
                                            value = dict_data[i[0]]
                                            total = total + value
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()
                                        except:
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()
                                else:
                                    if int(date_total_queuing[0]) >= int(tgl_start) and int(date_total_queuing[0]) <= int(tgl_end):
                                        total = total + 1
                                        try:
                                            value = dict_data[i[0]]
                                            total = total + value
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()
                                        except:
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()
                else:
                    # get data
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT date FROM tb_queuing_history where packaging = ? AND activity = ?",(packaging_total_queuing, transaction))
                    rec = c.fetchall()
                    file.commit()
                    file.close()
                    for i in rec:
                        date_total_queuing = i[0]
                        date_total_queuing = date_total_queuing.split("/")
                        total = 0
                        if int(date_total_queuing[2]) >= int(tahun_start) and int(date_total_queuing[2]) <= int(tahun_end):
                            if int(date_total_queuing[1]) >= int(bln_start) and int(date_total_queuing[1]) <= int(bln_end):
                                if int(bln_start) < int(date_total_queuing[1]):
                                    if int(date_total_queuing[0]) <= int(tgl_end):
                                        total = total + 1
                                        try:
                                            value = dict_data[i[0]]
                                            total = total + value
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()
                                        except:
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()
                                else:
                                    if int(date_total_queuing[0]) >= int(tgl_start) and int(date_total_queuing[0]) <= int(tgl_end):
                                        total = total + 1
                                        try:
                                            value = dict_data[i[0]]
                                            total = total + value
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()
                                        except:
                                            new_dict_data = {i[0] : total}
                                            dict_data.update(new_dict_data)
                                            new_dict_data.clear()

        # data
        dict_data = dict(sorted(dict_data.items(), key=lambda x: self.date_key(x[0])))
        for i in dict_data: 
            if count % 2 == 0:
                self.report_viewer_total_queuing_table.insert(parent="",index="end", iid=count, text="", values=(wh_name, transaction, i, dict_data[i]), tags=("evenrow",))
            else:
                self.report_viewer_total_queuing_table.insert(parent="",index="end", iid=count, text="", values=(wh_name, transaction, i, dict_data[i]), tags=("oddrow",))
            count+=1
        data_date = list(dict_data.keys())
        data_total_queuing = list(dict_data.values())
        self.ax3.cla()
        self.ax3.bar(data_date, data_total_queuing, color ='maroon', width = 0.4)
        self.ax3.set_xlabel("Date")
        self.ax3.set_ylabel("Total")
        self.ax3.set_title("Total Queuing")
        # Rotate the tick labels vertically
        self.ax3.set_xticks(data_date)
        self.ax3.set_xticklabels(data_date, rotation=90)
        self.bar_data_total_queuing.draw()
        self.bar_data_total_queuing.get_tk_widget().configure()

    def search_report_incident_report(self):
        category_incident = self.report_viewer_incident_report_optionmenu_category.get()
        status_incident = self.report_viewer_incident_report_optionmenu_status.get()
        datestart_incident = self.report_viewer_incident_report_dateStart_calendar.get()
        datestart_incident = datestart_incident.split("/")
        dateend_incident = self.report_viewer_incident_report_dateEnd_calendar.get()
        dateend_incident =dateend_incident.split("/")

        # date treatment
        if int(datestart_incident[1]) < 10 :
            tgl_start = "0" + datestart_incident[1]
        else:
            tgl_start = datestart_incident[1]
        if int(datestart_incident[0]) < 10 :
            bln_start = "0" + datestart_incident[0]
        else:
            bln_start = datestart_incident[0]
        tahun_start = "20" + datestart_incident[2]

        if int(dateend_incident[1]) < 10 :
            tgl_end = "0" + dateend_incident[1]
        else:
            tgl_end = dateend_incident[1]
        if int(dateend_incident[0]) < 10 :
            bln_end = "0" + dateend_incident[0]
        else:
            bln_end = dateend_incident[0]
        tahun_end = "20" + dateend_incident[2]

        # delete all from tree
        for record in self.report_viewer_incident_report_table.get_children():
            self.report_viewer_incident_report_table.delete(record)

        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT * FROM tb_history where category = ? AND status = ?",(category_incident, status_incident))
        rec = c.fetchall()

        # input to tabel GUI
        count = 0
        for i in rec:
            date_incident = i[0]
            date_incident = date_incident.split("/")
            if int(date_incident[2]) >= int(tahun_start) and int(date_incident[2]) <= int(tahun_end):
                if int(date_incident[1]) >= int(bln_start) and int(date_incident[1]) <= int(bln_end):
                    if int(bln_start) < int(date_incident[1]):
                        if int(date_incident[0]) <= int(tgl_end):
                            if count % 2 == 0:
                                self.report_viewer_incident_report_table.insert(parent="",index="end", iid=count, text="", values=(i[6], i[0], i[1], i[3], i[5], i[2], i[11], i[10], i[12], i[8], i[9], i[4], i[7]), tags=("evenrow",))
                            else:
                                self.report_viewer_incident_report_table.insert(parent="",index="end", iid=count, text="", values=(i[6], i[0], i[1], i[3], i[5], i[2], i[11], i[10], i[12], i[8], i[9], i[4], i[7]), tags=("oddrow",))
                            count+=1
                    else:
                        if int(date_incident[0]) >= int(tgl_start) and int(date_incident[0]) <= int(tgl_end):
                            if count % 2 == 0:
                                self.report_viewer_incident_report_table.insert(parent="",index="end", iid=count, text="", values=(i[6], i[0], i[1], i[3], i[5], i[2], i[11], i[10], i[12], i[8], i[9], i[4], i[7]), tags=("evenrow",))
                            else:
                                self.report_viewer_incident_report_table.insert(parent="",index="end", iid=count, text="", values=(i[6], i[0], i[1], i[3], i[5], i[2], i[11], i[10], i[12], i[8], i[9], i[4], i[7]), tags=("oddrow",))
                            count+=1
        file.commit()
        file.close()

    def search_report_ticket_history(self):
        transaction = self.report_viewer_ticket_history_optionmenu_transaction.get()
        packaging_ticket_history = self.report_viewer_ticket_history_optionmenu_packaging.get()
        datestart_ticket_history = self.report_viewer_ticket_history_dateStart_calendar.get()
        datestart_ticket_history = datestart_ticket_history.split("/")
        dateend_ticket_history = self.report_viewer_ticket_history_dateEnd_calendar.get()
        dateend_ticket_history = dateend_ticket_history.split("/")

        # date treatment
        if int(datestart_ticket_history[1]) < 10 :
            tgl_start = "0" + datestart_ticket_history[1]
        else:
            tgl_start = datestart_ticket_history[1]
        if int(datestart_ticket_history[0]) < 10 :
            bln_start = "0" + datestart_ticket_history[0]
        else:
            bln_start = datestart_ticket_history[0]
        tahun_start = "20" + datestart_ticket_history[2]

        if int(dateend_ticket_history[1]) < 10 :
            tgl_end = "0" + dateend_ticket_history[1]
        else:
            tgl_end = dateend_ticket_history[1]
        if int(dateend_ticket_history[0]) < 10 :
            bln_end = "0" + dateend_ticket_history[0]
        else:
            bln_end = dateend_ticket_history[0]
        tahun_end = "20" + dateend_ticket_history[2]

        # delete all from tree
        for record in self.report_viewer_ticket_history_table.get_children():
            self.report_viewer_ticket_history_table.delete(record)

        if transaction == "ALL":
            # get data from database
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT * FROM tb_ticket_history where packaging = ?",(packaging_ticket_history))
            rec = c.fetchall()
            # input to tabel GUI
            count = 0
            for i in rec:
                date_ticket = i[3]
                date_ticket = date_ticket.split("/")
                if int(date_ticket[2]) >= int(tahun_start) and int(date_ticket[2]) <= int(tahun_end):
                    if int(date_ticket[1]) >= int(bln_start) and int(date_ticket[1]) <= int(bln_end):
                        if int(bln_start) < int(date_ticket[1]):
                            if int(date_ticket[0]) <= int(tgl_end):
                                if count % 2 == 0:
                                    self.report_viewer_ticket_history_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
                                else:
                                    self.report_viewer_ticket_history_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
                                count+=1
                        else:
                            if int(date_ticket[0]) >= int(tgl_start) and int(date_ticket[0]) <= int(tgl_end):
                                if count % 2 == 0:
                                    self.report_viewer_ticket_history_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
                                else:
                                    self.report_viewer_ticket_history_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
                                count+=1
            file.commit()
            file.close()
        else:
            # get data from database
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT * FROM tb_ticket_history where activity = ? AND packaging = ?",(transaction, packaging_ticket_history))
            rec = c.fetchall()
            # input to tabel GUI
            count = 0
            for i in rec:
                date_ticket = i[3]
                date_ticket = date_ticket.split("/")
                if int(date_ticket[2]) >= int(tahun_start) and int(date_ticket[2]) <= int(tahun_end):
                    if int(date_ticket[1]) >= int(bln_start) and int(date_ticket[1]) <= int(bln_end):
                        if int(bln_start) < int(date_ticket[1]):
                            if int(date_ticket[0]) <= int(tgl_end):
                                if count % 2 == 0:
                                    self.report_viewer_ticket_history_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
                                else:
                                    self.report_viewer_ticket_history_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
                                count+=1
                        else:
                            if int(date_ticket[0]) >= int(tgl_start) and int(date_ticket[0]) <= int(tgl_end):
                                if count % 2 == 0:
                                    self.report_viewer_ticket_history_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
                                else:
                                    self.report_viewer_ticket_history_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
                                count+=1
            file.commit()
            file.close()

    def search_report_sla(self):
        transaction = self.report_viewer_sla_optionmenu_transaction.get()
        packaging_sla = self.report_viewer_sla_optionmenu_packaging.get()
        product_sla = self.report_viewer_sla_optionmenu_product.get()
        datestart_sla = self.report_viewer_sla_dateStart_calendar.get()
        datestart_sla = datestart_sla.split("/")
        dateend_sla = self.report_viewer_sla_dateEnd_calendar.get()
        dateend_sla = dateend_sla.split("/")

        # date treatment
        if int(datestart_sla[1]) < 10 :
            tgl_start = "0" + datestart_sla[1]
        else:
            tgl_start = datestart_sla[1]
        if int(datestart_sla[0]) < 10 :
            bln_start = "0" + datestart_sla[0]
        else:
            bln_start = datestart_sla[0]
        tahun_start = "20" + datestart_sla[2]

        if int(dateend_sla[1]) < 10 :
            tgl_end = "0" + dateend_sla[1]
        else:
            tgl_end = dateend_sla[1]
        if int(dateend_sla[0]) < 10 :
            bln_end = "0" + dateend_sla[0]
        else:
            bln_end = dateend_sla[0]
        tahun_end = "20" + dateend_sla[2]

        # delete all from tree
        for record in self.report_viewer_sla_table.get_children():
            self.report_viewer_sla_table.delete(record)

        count = 1
        dict_avg = {"wt":0,"wb1":0,"wh":0,"wb2":0,"cc":0,"total":0}

        if product_sla != "-":
            if transaction == "ALL":
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                c.execute("SELECT date, product, berat FROM tb_queuing_history where packaging = ? AND product = ?",(packaging_sla, product_sla))
                rec_date = c.fetchall()
                file.commit()
                file.close()
                for i in rec_date:
                    product = i[1]
                    date_sla = i[0]
                    date_sla = date_sla.split("/")
                    if int(date_sla[2]) >= int(tahun_start) and int(date_sla[2]) <= int(tahun_end):
                        if int(date_sla[1]) >= int(bln_start) and int(date_sla[1]) <= int(bln_end):
                            if int(bln_start) < int(date_sla[1]):
                                if int(date_sla[0]) <= int(tgl_end):
                                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                    c = file.cursor()
                                    c.execute("SELECT product, packaging, waiting_minute, wb1_minute, wh_minute, wb2_minute, cc_minute FROM tb_queuing_history where date = ? AND packaging = ? AND product = ?",(i[0], packaging_sla, product_sla))
                                    rec = c.fetchall()
                                    file.commit()
                                    file.close()
                                    pembagi = len(rec)
                                    for j in rec:
                                        if j[0] == i[1]:
                                            dict_avg['wt'] = round((dict_avg['wt'] + float(j[2]))/pembagi,2)
                                            dict_avg['wb1'] = round((dict_avg['wb1'] + float(j[3]))/pembagi,2)
                                            dict_avg['wh'] = round((dict_avg['wh'] + float(j[4]))/pembagi,2)
                                            dict_avg['wb2'] = round((dict_avg['wb2'] + float(j[5]))/pembagi,2)
                                            dict_avg['cc'] = round((dict_avg['cc'] + float(j[6]))/pembagi,2)
                                            # dict_avg['total'] = dict_avg['total'] + float(total_minute)
                                            dict_avg['total'] = round(dict_avg['wt']+dict_avg['wb1']+dict_avg['wh']+dict_avg['wb2']+dict_avg['cc'],2)
                                    # treatment berat (i[2])
                                    berat = i[2]
                                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                    c = file.cursor()
                                    c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla where packaging = ? AND product_grp = ?",(packaging_sla, product_sla))
                                    rec_berat = c.fetchall()
                                    file.commit()
                                    file.close()
                                    for y in rec_berat:
                                        weight_min = y[0]
                                        weight_max = y[1]
                                        # get SLA wb1 where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                                        if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                            c = file.cursor()
                                            c.execute("SELECT waiting_time, w_bridge1, wh_sla, w_bridge2, covering_time FROM tb_warehouseSla where packaging = ? AND product_grp = ? AND weight_min = ? AND weight_max = ?",(packaging_sla, product_sla, weight_min, weight_max))
                                            rec = c.fetchone()
                                            kpi_wt = rec[0]
                                            kpi_wb1 = rec[1]
                                            kpi_wh = rec[2]
                                            kpi_wb2 = rec[3]
                                            kpi_cc = rec[4]
                                            file.commit()
                                            file.close()
                                            weight_min = int(int(weight_min)/1000)
                                            weight_max = int(int(weight_max)/1000)
                                            range_ton = str(weight_min)+"-"+str(weight_max)
                                    # input to treeview based on count
                                    for x in range(2):
                                        if count % 2 == 0: # genap ACTUAL
                                            self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=("", "", "", "Actual", dict_avg['wt'], dict_avg['wb1'], dict_avg['wh'], dict_avg['wb2'], dict_avg['cc'], dict_avg['total']), tags=("evenrow",))
                                        else: # ganjil ACTUAL
                                            total = str(int(kpi_wt)+int(kpi_wb1)+int(kpi_wh)+int(kpi_wb2)+int(kpi_cc))
                                            self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=(product, packaging_sla, range_ton, "KPI", kpi_wt, kpi_wb1, kpi_wh, kpi_wb2, kpi_cc, total), tags=("oddrow",))
                                        count+=1
                                    # reset dict_avg
                                    dict_avg['wt'] = 0
                                    dict_avg['wb1'] = 0
                                    dict_avg['wh'] = 0
                                    dict_avg['wb2'] = 0
                                    dict_avg['cc'] = 0
                                    dict_avg['total'] = 0
                            else:
                                if int(date_sla[0]) >= int(tgl_start) and int(date_sla[0]) <= int(tgl_end):
                                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                    c = file.cursor()
                                    c.execute("SELECT product, packaging, waiting_minute, wb1_minute, wh_minute, wb2_minute, cc_minute FROM tb_queuing_history where date = ? AND packaging = ? AND product = ?",(i[0], packaging_sla, product_sla))
                                    rec = c.fetchall()
                                    file.commit()
                                    file.close()
                                    pembagi = len(rec)
                                    for j in rec:
                                        if j[0] == i[1]:
                                            dict_avg['wt'] = round((dict_avg['wt'] + float(j[2]))/pembagi,2)
                                            dict_avg['wb1'] = round((dict_avg['wb1'] + float(j[3]))/pembagi,2)
                                            dict_avg['wh'] = round((dict_avg['wh'] + float(j[4]))/pembagi,2)
                                            dict_avg['wb2'] = round((dict_avg['wb2'] + float(j[5]))/pembagi,2)
                                            dict_avg['cc'] = round((dict_avg['cc'] + float(j[6]))/pembagi,2)
                                            # dict_avg['total'] = dict_avg['total'] + float(total_minute)
                                            dict_avg['total'] = round(dict_avg['wt']+dict_avg['wb1']+dict_avg['wh']+dict_avg['wb2']+dict_avg['cc'],2)
                                    # treatment berat (i[2])
                                    berat = i[2]
                                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                    c = file.cursor()
                                    c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla where packaging = ? AND product_grp = ?",(packaging_sla, product_sla))
                                    rec_berat = c.fetchall()
                                    file.commit()
                                    file.close()
                                    for y in rec_berat:
                                        weight_min = y[0]
                                        weight_max = y[1]
                                        # get SLA wb1 where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                                        if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                            c = file.cursor()
                                            c.execute("SELECT waiting_time, w_bridge1, wh_sla, w_bridge2, covering_time FROM tb_warehouseSla where packaging = ? AND product_grp = ? AND weight_min = ? AND weight_max = ?",(packaging_sla, product_sla, weight_min, weight_max))
                                            rec = c.fetchone()
                                            kpi_wt = rec[0]
                                            kpi_wb1 = rec[1]
                                            kpi_wh = rec[2]
                                            kpi_wb2 = rec[3]
                                            kpi_cc = rec[4]
                                            file.commit()
                                            file.close()
                                            weight_min = int(int(weight_min)/1000)
                                            weight_max = int(int(weight_max)/1000)
                                            range_ton = str(weight_min)+"-"+str(weight_max)
                                    # input to treeview based on count
                                    for x in range(2):
                                        if count % 2 == 0: # genap ACTUAL
                                            self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=("", "", "", "Actual", dict_avg['wt'], dict_avg['wb1'], dict_avg['wh'], dict_avg['wb2'], dict_avg['cc'], dict_avg['total']), tags=("evenrow",))
                                        else: # ganjil ACTUAL
                                            total = str(int(kpi_wt)+int(kpi_wb1)+int(kpi_wh)+int(kpi_wb2)+int(kpi_cc))
                                            self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=(product, packaging_sla, range_ton, "KPI", kpi_wt, kpi_wb1, kpi_wh, kpi_wb2, kpi_cc, total), tags=("oddrow",))
                                        count+=1
                                    # reset dict_avg
                                    dict_avg['wt'] = 0
                                    dict_avg['wb1'] = 0
                                    dict_avg['wh'] = 0
                                    dict_avg['wb2'] = 0
                                    dict_avg['cc'] = 0
                                    dict_avg['total'] = 0
            else:
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                c.execute("SELECT date, product, berat FROM tb_queuing_history where packaging = ? AND product = ? AND activity = ?",(packaging_sla, product_sla, transaction))
                rec_date = c.fetchall()
                file.commit()
                file.close()
                for i in rec_date:
                    product = i[1]
                    date_sla = i[0]
                    date_sla = date_sla.split("/")
                    if int(date_sla[2]) >= int(tahun_start) and int(date_sla[2]) <= int(tahun_end):
                        if int(date_sla[1]) >= int(bln_start) and int(date_sla[1]) <= int(bln_end):
                            if int(bln_start) < int(date_sla[1]):
                                if int(date_sla[0]) <= int(tgl_end):
                                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                    c = file.cursor()
                                    c.execute("SELECT product, packaging, waiting_minute, wb1_minute, wh_minute, wb2_minute, cc_minute FROM tb_queuing_history where date = ? AND packaging = ? AND product = ? AND activity = ?",(i[0], packaging_sla, product_sla, transaction))
                                    rec = c.fetchall()
                                    file.commit()
                                    file.close()
                                    pembagi = len(rec)
                                    for j in rec:
                                        if j[0] == i[1]:
                                            dict_avg['wt'] = round((dict_avg['wt'] + float(j[2]))/pembagi,2)
                                            dict_avg['wb1'] = round((dict_avg['wb1'] + float(j[3]))/pembagi,2)
                                            dict_avg['wh'] = round((dict_avg['wh'] + float(j[4]))/pembagi,2)
                                            dict_avg['wb2'] = round((dict_avg['wb2'] + float(j[5]))/pembagi,2)
                                            dict_avg['cc'] = round((dict_avg['cc'] + float(j[6]))/pembagi,2)
                                            # dict_avg['total'] = dict_avg['total'] + float(total_minute)
                                            dict_avg['total'] = round(dict_avg['wt']+dict_avg['wb1']+dict_avg['wh']+dict_avg['wb2']+dict_avg['cc'],2)
                                    # treatment berat (i[2])
                                    berat = i[2]
                                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                    c = file.cursor()
                                    c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla where packaging = ? AND product_grp = ?",(packaging_sla, product_sla))
                                    rec_berat = c.fetchall()
                                    file.commit()
                                    file.close()
                                    for y in rec_berat:
                                        weight_min = y[0]
                                        weight_max = y[1]
                                        # get SLA wb1 where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                                        if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                            c = file.cursor()
                                            c.execute("SELECT waiting_time, w_bridge1, wh_sla, w_bridge2, covering_time FROM tb_warehouseSla where packaging = ? AND product_grp = ? AND weight_min = ? AND weight_max = ?",(packaging_sla, product_sla, weight_min, weight_max))
                                            rec = c.fetchone()
                                            kpi_wt = rec[0]
                                            kpi_wb1 = rec[1]
                                            kpi_wh = rec[2]
                                            kpi_wb2 = rec[3]
                                            kpi_cc = rec[4]
                                            file.commit()
                                            file.close()
                                            weight_min = int(int(weight_min)/1000)
                                            weight_max = int(int(weight_max)/1000)
                                            range_ton = str(weight_min)+"-"+str(weight_max)
                                    # input to treeview based on count
                                    for x in range(2):
                                        if count % 2 == 0: # genap ACTUAL
                                            self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=("", "", "", "Actual", dict_avg['wt'], dict_avg['wb1'], dict_avg['wh'], dict_avg['wb2'], dict_avg['cc'], dict_avg['total']), tags=("evenrow",))
                                        else: # ganjil ACTUAL
                                            total = str(int(kpi_wt)+int(kpi_wb1)+int(kpi_wh)+int(kpi_wb2)+int(kpi_cc))
                                            self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=(product, packaging_sla, range_ton, "KPI", kpi_wt, kpi_wb1, kpi_wh, kpi_wb2, kpi_cc, total), tags=("oddrow",))
                                        count+=1
                                    # reset dict_avg
                                    dict_avg['wt'] = 0
                                    dict_avg['wb1'] = 0
                                    dict_avg['wh'] = 0
                                    dict_avg['wb2'] = 0
                                    dict_avg['cc'] = 0
                                    dict_avg['total'] = 0
                            else:
                                if int(date_sla[0]) >= int(tgl_start) and int(date_sla[0]) <= int(tgl_end):
                                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                    c = file.cursor()
                                    c.execute("SELECT product, packaging, waiting_minute, wb1_minute, wh_minute, wb2_minute, cc_minute FROM tb_queuing_history where date = ? AND packaging = ? AND product = ? AND activity = ?",(i[0], packaging_sla, product_sla, transaction))
                                    rec = c.fetchall()
                                    file.commit()
                                    file.close()
                                    pembagi = len(rec)
                                    for j in rec:
                                        if j[0] == i[1]:
                                            dict_avg['wt'] = round((dict_avg['wt'] + float(j[2]))/pembagi,2)
                                            dict_avg['wb1'] = round((dict_avg['wb1'] + float(j[3]))/pembagi,2)
                                            dict_avg['wh'] = round((dict_avg['wh'] + float(j[4]))/pembagi,2)
                                            dict_avg['wb2'] = round((dict_avg['wb2'] + float(j[5]))/pembagi,2)
                                            dict_avg['cc'] = round((dict_avg['cc'] + float(j[6]))/pembagi,2)
                                            # dict_avg['total'] = dict_avg['total'] + float(total_minute)
                                            dict_avg['total'] = round(dict_avg['wt']+dict_avg['wb1']+dict_avg['wh']+dict_avg['wb2']+dict_avg['cc'],2)
                                    # treatment berat (i[2])
                                    berat = i[2]
                                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                    c = file.cursor()
                                    c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla where packaging = ? AND product_grp = ?",(packaging_sla, product_sla))
                                    rec_berat = c.fetchall()
                                    file.commit()
                                    file.close()
                                    for y in rec_berat:
                                        weight_min = y[0]
                                        weight_max = y[1]
                                        # get SLA wb1 where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                                        if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                            c = file.cursor()
                                            c.execute("SELECT waiting_time, w_bridge1, wh_sla, w_bridge2, covering_time FROM tb_warehouseSla where packaging = ? AND product_grp = ? AND weight_min = ? AND weight_max = ?",(packaging_sla, product_sla, weight_min, weight_max))
                                            rec = c.fetchone()
                                            kpi_wt = rec[0]
                                            kpi_wb1 = rec[1]
                                            kpi_wh = rec[2]
                                            kpi_wb2 = rec[3]
                                            kpi_cc = rec[4]
                                            file.commit()
                                            file.close()
                                            weight_min = int(int(weight_min)/1000)
                                            weight_max = int(int(weight_max)/1000)
                                            range_ton = str(weight_min)+"-"+str(weight_max)
                                    # input to treeview based on count
                                    for x in range(2):
                                        if count % 2 == 0: # genap ACTUAL
                                            self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=("", "", "", "Actual", dict_avg['wt'], dict_avg['wb1'], dict_avg['wh'], dict_avg['wb2'], dict_avg['cc'], dict_avg['total']), tags=("evenrow",))
                                        else: # ganjil ACTUAL
                                            total = str(int(kpi_wt)+int(kpi_wb1)+int(kpi_wh)+int(kpi_wb2)+int(kpi_cc))
                                            self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=(product, packaging_sla, range_ton, "KPI", kpi_wt, kpi_wb1, kpi_wh, kpi_wb2, kpi_cc, total), tags=("oddrow",))
                                        count+=1
                                    # reset dict_avg
                                    dict_avg['wt'] = 0
                                    dict_avg['wb1'] = 0
                                    dict_avg['wh'] = 0
                                    dict_avg['wb2'] = 0
                                    dict_avg['cc'] = 0
                                    dict_avg['total'] = 0
        else:
            if transaction == "ALL":
                if packaging_sla == "ALL":
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT date, product, berat FROM tb_queuing_history")
                    rec_date = c.fetchall()
                    file.commit()
                    file.close()
                    for i in rec_date:
                        product = i[1]
                        date_sla = i[0]
                        date_sla = date_sla.split("/")
                        if int(date_sla[2]) >= int(tahun_start) and int(date_sla[2]) <= int(tahun_end):
                            if int(date_sla[1]) >= int(bln_start) and int(date_sla[1]) <= int(bln_end):
                                if int(bln_start) < int(date_sla[1]):
                                    if int(date_sla[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT product, packaging, waiting_minute, wb1_minute, wh_minute, wb2_minute, cc_minute FROM tb_queuing_history where date = ?",(i[0]))
                                        rec = c.fetchall()
                                        file.commit()
                                        file.close()
                                        pembagi = len(rec)
                                        for j in rec:
                                            if j[0] == i[1]:
                                                dict_avg['wt'] = round((dict_avg['wt'] + float(j[2]))/pembagi,2)
                                                dict_avg['wb1'] = round((dict_avg['wb1'] + float(j[3]))/pembagi,2)
                                                dict_avg['wh'] = round((dict_avg['wh'] + float(j[4]))/pembagi,2)
                                                dict_avg['wb2'] = round((dict_avg['wb2'] + float(j[5]))/pembagi,2)
                                                dict_avg['cc'] = round((dict_avg['cc'] + float(j[6]))/pembagi,2)
                                                # dict_avg['total'] = dict_avg['total'] + float(total_minute)
                                                dict_avg['total'] = round(dict_avg['wt']+dict_avg['wb1']+dict_avg['wh']+dict_avg['wb2']+dict_avg['cc'],2)
                                        # treatment berat (i[2])
                                        berat = i[2]
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla")
                                        rec_berat = c.fetchall()
                                        file.commit()
                                        file.close()
                                        for y in rec_berat:
                                            weight_min = y[0]
                                            weight_max = y[1]
                                            # get SLA wb1 where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                                            if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                c = file.cursor()
                                                c.execute("SELECT waiting_time, w_bridge1, wh_sla, w_bridge2, covering_time FROM tb_warehouseSla where weight_min = ? AND weight_max = ?",(weight_min, weight_max))
                                                rec = c.fetchone()
                                                kpi_wt = rec[0]
                                                kpi_wb1 = rec[1]
                                                kpi_wh = rec[2]
                                                kpi_wb2 = rec[3]
                                                kpi_cc = rec[4]
                                                file.commit()
                                                file.close()
                                                weight_min = int(int(weight_min)/1000)
                                                weight_max = int(int(weight_max)/1000)
                                                range_ton = str(weight_min)+"-"+str(weight_max)
                                        # input to treeview based on count
                                        for x in range(2):
                                            if count % 2 == 0: # genap ACTUAL
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=("", "", "", "Actual", dict_avg['wt'], dict_avg['wb1'], dict_avg['wh'], dict_avg['wb2'], dict_avg['cc'], dict_avg['total']), tags=("evenrow",))
                                            else: # ganjil ACTUAL
                                                total = str(int(kpi_wt)+int(kpi_wb1)+int(kpi_wh)+int(kpi_wb2)+int(kpi_cc))
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=(product, packaging_sla, range_ton, "KPI", kpi_wt, kpi_wb1, kpi_wh, kpi_wb2, kpi_cc, total), tags=("oddrow",))
                                            count+=1
                                        # reset dict_avg
                                        dict_avg['wt'] = 0
                                        dict_avg['wb1'] = 0
                                        dict_avg['wh'] = 0
                                        dict_avg['wb2'] = 0
                                        dict_avg['cc'] = 0
                                        dict_avg['total'] = 0
                                else:
                                    if int(date_sla[0]) >= int(tgl_start) and int(date_sla[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT product, packaging, waiting_minute, wb1_minute, wh_minute, wb2_minute, cc_minute FROM tb_queuing_history where date = ?",(i[0]))
                                        rec = c.fetchall()
                                        file.commit()
                                        file.close()
                                        pembagi = len(rec)
                                        for j in rec:
                                            if j[0] == i[1]:
                                                dict_avg['wt'] = round((dict_avg['wt'] + float(j[2]))/pembagi,2)
                                                dict_avg['wb1'] = round((dict_avg['wb1'] + float(j[3]))/pembagi,2)
                                                dict_avg['wh'] = round((dict_avg['wh'] + float(j[4]))/pembagi,2)
                                                dict_avg['wb2'] = round((dict_avg['wb2'] + float(j[5]))/pembagi,2)
                                                dict_avg['cc'] = round((dict_avg['cc'] + float(j[6]))/pembagi,2)
                                                # dict_avg['total'] = dict_avg['total'] + float(total_minute)
                                                dict_avg['total'] = round(dict_avg['wt']+dict_avg['wb1']+dict_avg['wh']+dict_avg['wb2']+dict_avg['cc'],2)
                                        # treatment berat (i[2])
                                        berat = i[2]
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla")
                                        rec_berat = c.fetchall()
                                        file.commit()
                                        file.close()
                                        for y in rec_berat:
                                            weight_min = y[0]
                                            weight_max = y[1]
                                            # get SLA wb1 where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                                            if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                c = file.cursor()
                                                c.execute("SELECT waiting_time, w_bridge1, wh_sla, w_bridge2, covering_time FROM tb_warehouseSla where weight_min = ? AND weight_max = ?",(weight_min, weight_max))
                                                rec = c.fetchone()
                                                kpi_wt = rec[0]
                                                kpi_wb1 = rec[1]
                                                kpi_wh = rec[2]
                                                kpi_wb2 = rec[3]
                                                kpi_cc = rec[4]
                                                file.commit()
                                                file.close()
                                                weight_min = int(int(weight_min)/1000)
                                                weight_max = int(int(weight_max)/1000)
                                                range_ton = str(weight_min)+"-"+str(weight_max)
                                        # input to treeview based on count
                                        for x in range(2):
                                            if count % 2 == 0: # genap ACTUAL
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=("", "", "", "Actual", dict_avg['wt'], dict_avg['wb1'], dict_avg['wh'], dict_avg['wb2'], dict_avg['cc'], dict_avg['total']), tags=("evenrow",))
                                            else: # ganjil ACTUAL
                                                total = str(int(kpi_wt)+int(kpi_wb1)+int(kpi_wh)+int(kpi_wb2)+int(kpi_cc))
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=(product, packaging_sla, range_ton, "KPI", kpi_wt, kpi_wb1, kpi_wh, kpi_wb2, kpi_cc, total), tags=("oddrow",))
                                            count+=1
                                        # reset dict_avg
                                        dict_avg['wt'] = 0
                                        dict_avg['wb1'] = 0
                                        dict_avg['wh'] = 0
                                        dict_avg['wb2'] = 0
                                        dict_avg['cc'] = 0
                                        dict_avg['total'] = 0
                else:
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT date, product, berat FROM tb_queuing_history where packaging = ?",(packaging_sla,))
                    rec_date = c.fetchall()
                    file.commit()
                    file.close()
                    for i in rec_date:
                        product = i[1]
                        date_sla = i[0]
                        date_sla = date_sla.split("/")
                        if int(date_sla[2]) >= int(tahun_start) and int(date_sla[2]) <= int(tahun_end):
                            if int(date_sla[1]) >= int(bln_start) and int(date_sla[1]) <= int(bln_end):
                                if int(bln_start) < int(date_sla[1]):
                                    if int(date_sla[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT product, packaging, waiting_minute, wb1_minute, wh_minute, wb2_minute, cc_minute FROM tb_queuing_history where date = ? AND packaging = ?",(i[0], packaging_sla))
                                        rec = c.fetchall()
                                        file.commit()
                                        file.close()
                                        pembagi = len(rec)
                                        for j in rec:
                                            if j[0] == i[1]:
                                                dict_avg['wt'] = round((dict_avg['wt'] + float(j[2]))/pembagi,2)
                                                dict_avg['wb1'] = round((dict_avg['wb1'] + float(j[3]))/pembagi,2)
                                                dict_avg['wh'] = round((dict_avg['wh'] + float(j[4]))/pembagi,2)
                                                dict_avg['wb2'] = round((dict_avg['wb2'] + float(j[5]))/pembagi,2)
                                                dict_avg['cc'] = round((dict_avg['cc'] + float(j[6]))/pembagi,2)
                                                # dict_avg['total'] = dict_avg['total'] + float(total_minute)
                                                dict_avg['total'] = round(dict_avg['wt']+dict_avg['wb1']+dict_avg['wh']+dict_avg['wb2']+dict_avg['cc'],2)
                                        # treatment berat (i[2])
                                        berat = i[2]
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla where packaging = ?",(packaging_sla))
                                        rec_berat = c.fetchall()
                                        file.commit()
                                        file.close()
                                        for y in rec_berat:
                                            weight_min = y[0]
                                            weight_max = y[1]
                                            # get SLA wb1 where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                                            if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                c = file.cursor()
                                                c.execute("SELECT waiting_time, w_bridge1, wh_sla, w_bridge2, covering_time FROM tb_warehouseSla where packaging = ? AND weight_min = ? AND weight_max = ?",(packaging_sla, weight_min, weight_max))
                                                rec = c.fetchone()
                                                kpi_wt = rec[0]
                                                kpi_wb1 = rec[1]
                                                kpi_wh = rec[2]
                                                kpi_wb2 = rec[3]
                                                kpi_cc = rec[4]
                                                file.commit()
                                                file.close()
                                                weight_min = int(int(weight_min)/1000)
                                                weight_max = int(int(weight_max)/1000)
                                                range_ton = str(weight_min)+"-"+str(weight_max)
                                        # input to treeview based on count
                                        for x in range(2):
                                            if count % 2 == 0: # genap ACTUAL
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=("", "", "", "Actual", dict_avg['wt'], dict_avg['wb1'], dict_avg['wh'], dict_avg['wb2'], dict_avg['cc'], dict_avg['total']), tags=("evenrow",))
                                            else: # ganjil ACTUAL
                                                total = str(int(kpi_wt)+int(kpi_wb1)+int(kpi_wh)+int(kpi_wb2)+int(kpi_cc))
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=(product, packaging_sla, range_ton, "KPI", kpi_wt, kpi_wb1, kpi_wh, kpi_wb2, kpi_cc, total), tags=("oddrow",))
                                            count+=1
                                        # reset dict_avg
                                        dict_avg['wt'] = 0
                                        dict_avg['wb1'] = 0
                                        dict_avg['wh'] = 0
                                        dict_avg['wb2'] = 0
                                        dict_avg['cc'] = 0
                                        dict_avg['total'] = 0
                                else:
                                    if int(date_sla[0]) >= int(tgl_start) and int(date_sla[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT product, packaging, waiting_minute, wb1_minute, wh_minute, wb2_minute, cc_minute FROM tb_queuing_history where date = ? AND packaging = ?",(i[0], packaging_sla))
                                        rec = c.fetchall()
                                        file.commit()
                                        file.close()
                                        pembagi = len(rec)
                                        for j in rec:
                                            if j[0] == i[1]:
                                                dict_avg['wt'] = round((dict_avg['wt'] + float(j[2]))/pembagi,2)
                                                dict_avg['wb1'] = round((dict_avg['wb1'] + float(j[3]))/pembagi,2)
                                                dict_avg['wh'] = round((dict_avg['wh'] + float(j[4]))/pembagi,2)
                                                dict_avg['wb2'] = round((dict_avg['wb2'] + float(j[5]))/pembagi,2)
                                                dict_avg['cc'] = round((dict_avg['cc'] + float(j[6]))/pembagi,2)
                                                # dict_avg['total'] = dict_avg['total'] + float(total_minute)
                                                dict_avg['total'] = round(dict_avg['wt']+dict_avg['wb1']+dict_avg['wh']+dict_avg['wb2']+dict_avg['cc'],2)
                                        # treatment berat (i[2])
                                        berat = i[2]
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla where packaging = ?",(packaging_sla,))
                                        rec_berat = c.fetchall()
                                        file.commit()
                                        file.close()
                                        for y in rec_berat:
                                            weight_min = y[0]
                                            weight_max = y[1]
                                            # get SLA wb1 where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                                            if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                c = file.cursor()
                                                c.execute("SELECT waiting_time, w_bridge1, wh_sla, w_bridge2, covering_time FROM tb_warehouseSla where packaging = ? AND weight_min = ? AND weight_max = ?",(packaging_sla, weight_min, weight_max))
                                                rec = c.fetchone()
                                                kpi_wt = rec[0]
                                                kpi_wb1 = rec[1]
                                                kpi_wh = rec[2]
                                                kpi_wb2 = rec[3]
                                                kpi_cc = rec[4]
                                                file.commit()
                                                file.close()
                                                weight_min = int(int(weight_min)/1000)
                                                weight_max = int(int(weight_max)/1000)
                                                range_ton = str(weight_min)+"-"+str(weight_max)
                                        # input to treeview based on count
                                        for x in range(2):
                                            if count % 2 == 0: # genap ACTUAL
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=("", "", "", "Actual", dict_avg['wt'], dict_avg['wb1'], dict_avg['wh'], dict_avg['wb2'], dict_avg['cc'], dict_avg['total']), tags=("evenrow",))
                                            else: # ganjil ACTUAL
                                                total = str(int(kpi_wt)+int(kpi_wb1)+int(kpi_wh)+int(kpi_wb2)+int(kpi_cc))
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=(product, packaging_sla, range_ton, "KPI", kpi_wt, kpi_wb1, kpi_wh, kpi_wb2, kpi_cc, total), tags=("oddrow",))
                                            count+=1
                                        # reset dict_avg
                                        dict_avg['wt'] = 0
                                        dict_avg['wb1'] = 0
                                        dict_avg['wh'] = 0
                                        dict_avg['wb2'] = 0
                                        dict_avg['cc'] = 0
                                        dict_avg['total'] = 0
            else:
                if packaging_sla == "ALL":
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT date, product, berat FROM tb_queuing_history where activity = ?",(transaction))
                    rec_date = c.fetchall()
                    file.commit()
                    file.close()
                    for i in rec_date:
                        product = i[1]
                        date_sla = i[0]
                        date_sla = date_sla.split("/")
                        if int(date_sla[2]) >= int(tahun_start) and int(date_sla[2]) <= int(tahun_end):
                            if int(date_sla[1]) >= int(bln_start) and int(date_sla[1]) <= int(bln_end):
                                if int(bln_start) < int(date_sla[1]):
                                    if int(date_sla[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT product, packaging, waiting_minute, wb1_minute, wh_minute, wb2_minute, cc_minute FROM tb_queuing_history where date = ? AND activity = ?",(i[0], transaction))
                                        rec = c.fetchall()
                                        file.commit()
                                        file.close()
                                        pembagi = len(rec)
                                        for j in rec:
                                            if j[0] == i[1]:
                                                dict_avg['wt'] = round((dict_avg['wt'] + float(j[2]))/pembagi,2)
                                                dict_avg['wb1'] = round((dict_avg['wb1'] + float(j[3]))/pembagi,2)
                                                dict_avg['wh'] = round((dict_avg['wh'] + float(j[4]))/pembagi,2)
                                                dict_avg['wb2'] = round((dict_avg['wb2'] + float(j[5]))/pembagi,2)
                                                dict_avg['cc'] = round((dict_avg['cc'] + float(j[6]))/pembagi,2)
                                                # dict_avg['total'] = dict_avg['total'] + float(total_minute)
                                                dict_avg['total'] = round(dict_avg['wt']+dict_avg['wb1']+dict_avg['wh']+dict_avg['wb2']+dict_avg['cc'],2)
                                        # treatment berat (i[2])
                                        berat = i[2]
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla")
                                        rec_berat = c.fetchall()
                                        file.commit()
                                        file.close()
                                        for y in rec_berat:
                                            weight_min = y[0]
                                            weight_max = y[1]
                                            # get SLA wb1 where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                                            if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                c = file.cursor()
                                                c.execute("SELECT waiting_time, w_bridge1, wh_sla, w_bridge2, covering_time FROM tb_warehouseSla where weight_min = ? AND weight_max = ?",(weight_min, weight_max))
                                                rec = c.fetchone()
                                                kpi_wt = rec[0]
                                                kpi_wb1 = rec[1]
                                                kpi_wh = rec[2]
                                                kpi_wb2 = rec[3]
                                                kpi_cc = rec[4]
                                                file.commit()
                                                file.close()
                                                weight_min = int(int(weight_min)/1000)
                                                weight_max = int(int(weight_max)/1000)
                                                range_ton = str(weight_min)+"-"+str(weight_max)
                                        # input to treeview based on count
                                        for x in range(2):
                                            if count % 2 == 0: # genap ACTUAL
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=("", "", "", "Actual", dict_avg['wt'], dict_avg['wb1'], dict_avg['wh'], dict_avg['wb2'], dict_avg['cc'], dict_avg['total']), tags=("evenrow",))
                                            else: # ganjil ACTUAL
                                                total = str(int(kpi_wt)+int(kpi_wb1)+int(kpi_wh)+int(kpi_wb2)+int(kpi_cc))
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=(product, packaging_sla, range_ton, "KPI", kpi_wt, kpi_wb1, kpi_wh, kpi_wb2, kpi_cc, total), tags=("oddrow",))
                                            count+=1
                                        # reset dict_avg
                                        dict_avg['wt'] = 0
                                        dict_avg['wb1'] = 0
                                        dict_avg['wh'] = 0
                                        dict_avg['wb2'] = 0
                                        dict_avg['cc'] = 0
                                        dict_avg['total'] = 0
                                else:
                                    if int(date_sla[0]) >= int(tgl_start) and int(date_sla[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT product, packaging, waiting_minute, wb1_minute, wh_minute, wb2_minute, cc_minute FROM tb_queuing_history where date = ? AND activity = ?",(i[0], transaction))
                                        rec = c.fetchall()
                                        file.commit()
                                        file.close()
                                        pembagi = len(rec)
                                        for j in rec:
                                            if j[0] == i[1]:
                                                dict_avg['wt'] = round((dict_avg['wt'] + float(j[2]))/pembagi,2)
                                                dict_avg['wb1'] = round((dict_avg['wb1'] + float(j[3]))/pembagi,2)
                                                dict_avg['wh'] = round((dict_avg['wh'] + float(j[4]))/pembagi,2)
                                                dict_avg['wb2'] = round((dict_avg['wb2'] + float(j[5]))/pembagi,2)
                                                dict_avg['cc'] = round((dict_avg['cc'] + float(j[6]))/pembagi,2)
                                                # dict_avg['total'] = dict_avg['total'] + float(total_minute)
                                                dict_avg['total'] = round(dict_avg['wt']+dict_avg['wb1']+dict_avg['wh']+dict_avg['wb2']+dict_avg['cc'],2)
                                        # treatment berat (i[2])
                                        berat = i[2]
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla")
                                        rec_berat = c.fetchall()
                                        file.commit()
                                        file.close()
                                        for y in rec_berat:
                                            weight_min = y[0]
                                            weight_max = y[1]
                                            # get SLA wb1 where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                                            if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                c = file.cursor()
                                                c.execute("SELECT waiting_time, w_bridge1, wh_sla, w_bridge2, covering_time FROM tb_warehouseSla where weight_min = ? AND weight_max = ?",(weight_min, weight_max))
                                                rec = c.fetchone()
                                                kpi_wt = rec[0]
                                                kpi_wb1 = rec[1]
                                                kpi_wh = rec[2]
                                                kpi_wb2 = rec[3]
                                                kpi_cc = rec[4]
                                                file.commit()
                                                file.close()
                                                weight_min = int(int(weight_min)/1000)
                                                weight_max = int(int(weight_max)/1000)
                                                range_ton = str(weight_min)+"-"+str(weight_max)
                                        # input to treeview based on count
                                        for x in range(2):
                                            if count % 2 == 0: # genap ACTUAL
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=("", "", "", "Actual", dict_avg['wt'], dict_avg['wb1'], dict_avg['wh'], dict_avg['wb2'], dict_avg['cc'], dict_avg['total']), tags=("evenrow",))
                                            else: # ganjil ACTUAL
                                                total = str(int(kpi_wt)+int(kpi_wb1)+int(kpi_wh)+int(kpi_wb2)+int(kpi_cc))
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=(product, packaging_sla, range_ton, "KPI", kpi_wt, kpi_wb1, kpi_wh, kpi_wb2, kpi_cc, total), tags=("oddrow",))
                                            count+=1
                                        # reset dict_avg
                                        dict_avg['wt'] = 0
                                        dict_avg['wb1'] = 0
                                        dict_avg['wh'] = 0
                                        dict_avg['wb2'] = 0
                                        dict_avg['cc'] = 0
                                        dict_avg['total'] = 0
                else:
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT date, product, berat FROM tb_queuing_history where packaging = ? AND activity = ?",(packaging_sla, transaction))
                    rec_date = c.fetchall()
                    file.commit()
                    file.close()
                    for i in rec_date:
                        product = i[1]
                        date_sla = i[0]
                        date_sla = date_sla.split("/")
                        if int(date_sla[2]) >= int(tahun_start) and int(date_sla[2]) <= int(tahun_end):
                            if int(date_sla[1]) >= int(bln_start) and int(date_sla[1]) <= int(bln_end):
                                if int(bln_start) < int(date_sla[1]):
                                    if int(date_sla[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT product, packaging, waiting_minute, wb1_minute, wh_minute, wb2_minute, cc_minute FROM tb_queuing_history where date = ? AND packaging = ? AND activity = ?",(i[0], packaging_sla, transaction))
                                        rec = c.fetchall()
                                        file.commit()
                                        file.close()
                                        pembagi = len(rec)
                                        for j in rec:
                                            if j[0] == i[1]:
                                                dict_avg['wt'] = round((dict_avg['wt'] + float(j[2]))/pembagi,2)
                                                dict_avg['wb1'] = round((dict_avg['wb1'] + float(j[3]))/pembagi,2)
                                                dict_avg['wh'] = round((dict_avg['wh'] + float(j[4]))/pembagi,2)
                                                dict_avg['wb2'] = round((dict_avg['wb2'] + float(j[5]))/pembagi,2)
                                                dict_avg['cc'] = round((dict_avg['cc'] + float(j[6]))/pembagi,2)
                                                # dict_avg['total'] = dict_avg['total'] + float(total_minute)
                                                dict_avg['total'] = round(dict_avg['wt']+dict_avg['wb1']+dict_avg['wh']+dict_avg['wb2']+dict_avg['cc'],2)
                                        # treatment berat (i[2])
                                        berat = i[2]
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla where packaging = ?",(packaging_sla,))
                                        rec_berat = c.fetchall()
                                        file.commit()
                                        file.close()
                                        for y in rec_berat:
                                            weight_min = y[0]
                                            weight_max = y[1]
                                            # get SLA wb1 where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                                            if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                c = file.cursor()
                                                c.execute("SELECT waiting_time, w_bridge1, wh_sla, w_bridge2, covering_time FROM tb_warehouseSla where packaging = ? AND weight_min = ? AND weight_max = ?",(packaging_sla, weight_min, weight_max))
                                                rec = c.fetchone()
                                                kpi_wt = rec[0]
                                                kpi_wb1 = rec[1]
                                                kpi_wh = rec[2]
                                                kpi_wb2 = rec[3]
                                                kpi_cc = rec[4]
                                                file.commit()
                                                file.close()
                                                weight_min = int(int(weight_min)/1000)
                                                weight_max = int(int(weight_max)/1000)
                                                range_ton = str(weight_min)+"-"+str(weight_max)
                                        # input to treeview based on count
                                        for x in range(2):
                                            if count % 2 == 0: # genap ACTUAL
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=("", "", "", "Actual", dict_avg['wt'], dict_avg['wb1'], dict_avg['wh'], dict_avg['wb2'], dict_avg['cc'], dict_avg['total']), tags=("evenrow",))
                                            else: # ganjil ACTUAL
                                                total = str(int(kpi_wt)+int(kpi_wb1)+int(kpi_wh)+int(kpi_wb2)+int(kpi_cc))
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=(product, packaging_sla, range_ton, "KPI", kpi_wt, kpi_wb1, kpi_wh, kpi_wb2, kpi_cc, total), tags=("oddrow",))
                                            count+=1
                                        # reset dict_avg
                                        dict_avg['wt'] = 0
                                        dict_avg['wb1'] = 0
                                        dict_avg['wh'] = 0
                                        dict_avg['wb2'] = 0
                                        dict_avg['cc'] = 0
                                        dict_avg['total'] = 0
                                else:
                                    if int(date_sla[0]) >= int(tgl_start) and int(date_sla[0]) <= int(tgl_end):
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT product, packaging, waiting_minute, wb1_minute, wh_minute, wb2_minute, cc_minute FROM tb_queuing_history where date = ? AND packaging = ? AND activity = ?",(i[0], packaging_sla, transaction))
                                        rec = c.fetchall()
                                        file.commit()
                                        file.close()
                                        pembagi = len(rec)
                                        for j in rec:
                                            if j[0] == i[1]:
                                                dict_avg['wt'] = round((dict_avg['wt'] + float(j[2]))/pembagi,2)
                                                dict_avg['wb1'] = round((dict_avg['wb1'] + float(j[3]))/pembagi,2)
                                                dict_avg['wh'] = round((dict_avg['wh'] + float(j[4]))/pembagi,2)
                                                dict_avg['wb2'] = round((dict_avg['wb2'] + float(j[5]))/pembagi,2)
                                                dict_avg['cc'] = round((dict_avg['cc'] + float(j[6]))/pembagi,2)
                                                # dict_avg['total'] = dict_avg['total'] + float(total_minute)
                                                dict_avg['total'] = round(dict_avg['wt']+dict_avg['wb1']+dict_avg['wh']+dict_avg['wb2']+dict_avg['cc'],2)
                                        # treatment berat (i[2])
                                        berat = i[2]
                                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                        c = file.cursor()
                                        c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla where packaging = ?",(packaging_sla,))
                                        rec_berat = c.fetchall()
                                        file.commit()
                                        file.close()
                                        for y in rec_berat:
                                            weight_min = y[0]
                                            weight_max = y[1]
                                            # get SLA wb1 where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                                            if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                                c = file.cursor()
                                                c.execute("SELECT waiting_time, w_bridge1, wh_sla, w_bridge2, covering_time FROM tb_warehouseSla where packaging = ? AND weight_min = ? AND weight_max = ?",(packaging_sla, weight_min, weight_max))
                                                rec = c.fetchone()
                                                kpi_wt = rec[0]
                                                kpi_wb1 = rec[1]
                                                kpi_wh = rec[2]
                                                kpi_wb2 = rec[3]
                                                kpi_cc = rec[4]
                                                file.commit()
                                                file.close()
                                                weight_min = int(int(weight_min)/1000)
                                                weight_max = int(int(weight_max)/1000)
                                                range_ton = str(weight_min)+"-"+str(weight_max)
                                        # input to treeview based on count
                                        for x in range(2):
                                            if count % 2 == 0: # genap ACTUAL
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=("", "", "", "Actual", dict_avg['wt'], dict_avg['wb1'], dict_avg['wh'], dict_avg['wb2'], dict_avg['cc'], dict_avg['total']), tags=("evenrow",))
                                            else: # ganjil ACTUAL
                                                total = str(int(kpi_wt)+int(kpi_wb1)+int(kpi_wh)+int(kpi_wb2)+int(kpi_cc))
                                                self.report_viewer_sla_table.insert(parent="",index="end", iid=count, text="", values=(product, packaging_sla, range_ton, "KPI", kpi_wt, kpi_wb1, kpi_wh, kpi_wb2, kpi_cc, total), tags=("oddrow",))
                                            count+=1
                                        # reset dict_avg
                                        dict_avg['wt'] = 0
                                        dict_avg['wb1'] = 0
                                        dict_avg['wh'] = 0
                                        dict_avg['wb2'] = 0
                                        dict_avg['cc'] = 0
                                        dict_avg['total'] = 0

    def search_report_auditlog(self):
        category = self.report_viewer_auditlog_optionmenu_category.get()
        datestart_log_history = self.report_viewer_auditlog_dateStart_calendar.get()
        datestart_log_history = datestart_log_history.split("/")
        dateend_log_history = self.report_viewer_auditlog_dateEnd_calendar.get()
        dateend_log_history = dateend_log_history.split("/")

        # date treatment
        if int(datestart_log_history[1]) < 10 :
            tgl_start = "0" + datestart_log_history[1]
        else:
            tgl_start = datestart_log_history[1]
        if int(datestart_log_history[0]) < 10 :
            bln_start = "0" + datestart_log_history[0]
        else:
            bln_start = datestart_log_history[0]
        tahun_start = "20" + datestart_log_history[2]

        if int(dateend_log_history[1]) < 10 :
            tgl_end = "0" + dateend_log_history[1]
        else:
            tgl_end = dateend_log_history[1]
        if int(dateend_log_history[0]) < 10 :
            bln_end = "0" + dateend_log_history[0]
        else:
            bln_end = dateend_log_history[0]
        tahun_end = "20" + dateend_log_history[2]

        # delete all from tree
        for record in self.report_viewer_auditlog_table.get_children():
            self.report_viewer_auditlog_table.delete(record)

        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT * FROM tb_log_report where transaction_log = ?",(category))
        rec = c.fetchall()

        # input to tabel GUI
        count = 0
        for i in rec:
            date_log = i[1]
            date_log = date_log.split("/")
            if int(date_log[2]) >= int(tahun_start) and int(date_log[2]) <= int(tahun_end):
                if int(date_log[1]) >= int(bln_start) and int(date_log[1]) <= int(bln_end):
                    if int(bln_start) < int(date_log[1]):
                        if int(date_log[0]) <= int(tgl_end):
                            if count % 2 == 0:
                                self.report_viewer_auditlog_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6]), tags=("evenrow",))
                            else:
                                self.report_viewer_auditlog_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6]), tags=("oddrow",))
                            count+=1
                    else:
                        if int(date_log[0]) >= int(tgl_start) and int(date_log[0]) <= int(tgl_end):
                            if count % 2 == 0:
                                self.report_viewer_auditlog_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6]), tags=("evenrow",))
                            else:
                                self.report_viewer_auditlog_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6]), tags=("oddrow",))
                            count+=1
        file.commit()
        file.close()

    # FUNGSI SAAT MEMILIH DATA PADA TABEL

    def on_tree_vetting_select(self, event):
        global no

        self.entry_company.delete(0,tk.END)
        self.entry_policeNo.delete(0,tk.END)
        self.entry_remark.delete(0,tk.END)
        self.entry_rfid1.delete(0,tk.END)
        self.entry_rfid2.delete(0,tk.END)

        # get data and set all entry with it
        try:
            selectedItem = self.vetting_table.selection()[0]
            no = self.vetting_table.item(selectedItem)['values'][0]
            company = self.vetting_table.item(selectedItem)['values'][2]
            select_police_no = self.vetting_table.item(selectedItem)['values'][3]
            reg_date = self.vetting_table.item(selectedItem)['values'][4]
            expired_date = self.vetting_table.item(selectedItem)['values'][5]
            remark = self.vetting_table.item(selectedItem)['values'][6]
            status = self.vetting_table.item(selectedItem)['values'][7]
            rfid1 = self.vetting_table.item(selectedItem)['values'][8]
            rfid2 = self.vetting_table.item(selectedItem)['values'][9]
            category = self.vetting_table.item(selectedItem)['values'][10]
        except:
            pass
        # insert to entry
        try:
            self.entry_company.insert(0, company)
            self.entry_policeNo.insert(0, select_police_no)
            self.regdate_calendar.set_date(date(int(reg_date[0:4]),int(reg_date[5:7]),int(reg_date[8:])))
            self.expireddate_calendar.set_date(date(int(expired_date[0:4]),int(expired_date[5:7]),int(expired_date[8:])))
            self.entry_remark.insert(0, remark)
            self.optionmenu_status.set(status)
            self.entry_rfid1.insert(0, rfid1)
            self.entry_rfid2.insert(0, rfid2)
            # self.entry_rfid1.insert(0, "{:04d}".format(int(rfid1)))
            # self.entry_rfid2.insert(0, "{:04d}".format(int(rfid2)))
            self.optionmenu_category.set(category)
        except:
            pass

    def on_tree_unloading_select(self, event):
        global no_unloading

        self.entry_company_unloading.delete(0,tk.END)
        self.entry_policeNo_unloading.delete(0,tk.END)
        self.spbox_hour.delete(0,tk.END)
        self.spbox_minute.delete(0,tk.END)
        self.entry_quota_unloading.delete(0,tk.END)
        self.entry_remark_unloading.delete(0,tk.END)
        self.entry_berat_unloading.delete(0,tk.END)
        self.optionmenu_noSo.delete(0,tk.END)

        # get data and set all entry with it
        try:
            selectedItem = self.unloading_table.selection()[0]
            no_unloading = self.unloading_table.item(selectedItem)['values'][0]
            company = self.unloading_table.item(selectedItem)['values'][2]
            select_police_no = self.unloading_table.item(selectedItem)['values'][3]
            packaging = self.unloading_table.item(selectedItem)['values'][4]
            product = self.unloading_table.item(selectedItem)['values'][5]
            date_value = self.unloading_table.item(selectedItem)['values'][6]
            time = self.unloading_table.item(selectedItem)['values'][7]
            remark = self.unloading_table.item(selectedItem)['values'][8]
            status = self.unloading_table.item(selectedItem)['values'][9]
            quota = self.unloading_table.item(selectedItem)['values'][10]
            berat = self.unloading_table.item(selectedItem)['values'][11]
            wh_id = self.unloading_table.item(selectedItem)['values'][12]
            no_so = self.unloading_table.item(selectedItem)['values'][13]
        except:
            pass
        # insert to entry
        if wh_id != "Auto":
            try:
                self.entry_company_unloading.insert(0, company)
                self.entry_policeNo_unloading.insert(0, select_police_no)
                self.optionmenu_packaging.set(packaging)
                self.optionmenu_product.set(product)
                self.date_unloading_calendar.set_date(date(int(date_value[0:4]),int(date_value[5:7]),int(date_value[8:])))
                self.spbox_hour.insert(0,time[0:2])
                self.spbox_minute.insert(0,time[3:])
                self.entry_quota_unloading.insert(0, quota)
                self.entry_berat_unloading.insert(0, berat)
                self.entry_remark_unloading.insert(0, remark)
                self.optionmenu_unloading_status.set(status)
                self.optionmenu_warehouse.set("Gudang {}".format(wh_id))
                self.optionmenu_noSo.insert(0, no_so)
            except:
                pass
        else:
            try:
                self.entry_company_unloading.insert(0, company)
                self.entry_policeNo_unloading.insert(0, select_police_no)
                self.optionmenu_packaging.set(packaging)
                self.optionmenu_product.set(product)
                self.date_unloading_calendar.set_date(date(int(date_value[0:4]),int(date_value[5:7]),int(date_value[8:])))
                self.spbox_hour.insert(0,time[0:2])
                self.spbox_minute.insert(0,time[3:])
                self.entry_quota_unloading.insert(0, quota)
                self.entry_berat_unloading.insert(0, berat)
                self.entry_remark_unloading.insert(0, remark)
                self.optionmenu_unloading_status.set(status)
                self.optionmenu_warehouse.set("Otomatis")
                self.optionmenu_noSo.insert(0, no_so)
            except:
                pass

    def on_tree_loadingList_select(self, event):
        global no_loadingList

        self.entry_company_distribution.delete(0,tk.END)
        self.entry_policeNo_distribution.delete(0,tk.END)
        self.spbox_hour_distribution.delete(0,tk.END)
        self.spbox_minute_distribution.delete(0,tk.END)
        self.entry_quota_distribution.delete(0,tk.END)
        self.entry_remark_distribution.delete(0,tk.END)
        self.entry_berat_distribution.delete(0,tk.END)
        self.entry_noSO_distribution.delete(0,tk.END)

        # get data and set all entry with it
        try:
            selectedItem = self.loadingList_table.selection()[0]
            no_loadingList = self.loadingList_table.item(selectedItem)['values'][0]
            company = self.loadingList_table.item(selectedItem)['values'][2]
            select_police_no = self.loadingList_table.item(selectedItem)['values'][3]
            packaging = self.loadingList_table.item(selectedItem)['values'][4]
            product = self.loadingList_table.item(selectedItem)['values'][5]
            date_value = self.loadingList_table.item(selectedItem)['values'][6]
            time = self.loadingList_table.item(selectedItem)['values'][7]
            remark = self.loadingList_table.item(selectedItem)['values'][8]
            status = self.loadingList_table.item(selectedItem)['values'][9]
            quota = self.loadingList_table.item(selectedItem)['values'][10]
            no_so = self.loadingList_table.item(selectedItem)['values'][11]
            satuan = self.loadingList_table.item(selectedItem)['values'][12]
            berat = self.loadingList_table.item(selectedItem)['values'][13]
            wh_id = self.loadingList_table.item(selectedItem)['values'][14]
        except:
            pass
        # insert to entry
        if wh_id != "Auto":
            try:
                self.entry_company_distribution.insert(0, company)
                self.entry_policeNo_distribution.insert(0, select_police_no)
                self.optionmenu_packaging_distribution.set(packaging)
                self.optionmenu_product_distribution.set(product)
                self.date_distribution_calendar.set_date(date(int(date_value[0:4]),int(date_value[5:7]),int(date_value[8:])))
                self.spbox_hour_distribution.insert(0,time[0:2])
                self.spbox_minute_distribution.insert(0,time[3:])
                self.entry_quota_distribution.insert(0, quota)
                self.entry_remark_distribution.insert(0, remark)
                self.optionmenu_distribution_status.set(status)
                self.optionmenu_warehouse_distribution.set("Gudang {}".format(wh_id))
                self.entry_noSO_distribution.insert(0, no_so)
                self.entry_berat_distribution.insert(0, berat)
            except:
                pass
        else:
            try:
                self.entry_company_distribution.insert(0, company)
                self.entry_policeNo_distribution.insert(0, select_police_no)
                self.optionmenu_packaging_distribution.set(packaging)
                self.optionmenu_product_distribution.set(product)
                self.date_distribution_calendar.set_date(date(int(date_value[0:4]),int(date_value[5:7]),int(date_value[8:])))
                self.spbox_hour_distribution.insert(0,time[0:2])
                self.spbox_minute_distribution.insert(0,time[3:])
                self.entry_quota_distribution.insert(0, quota)
                self.entry_remark_distribution.insert(0, remark)
                self.optionmenu_distribution_status.set(status)
                self.optionmenu_warehouse_distribution.set("Otomatis")
                self.entry_noSO_distribution.insert(0, no_so)
                self.entry_berat_distribution.insert(0, berat)
            except:
                pass

    def on_tree_displanData_select(self, event):
        self.entry_company_distribution.delete(0,tk.END)
        self.entry_policeNo_distribution.delete(0,tk.END)
        self.spbox_hour_distribution.delete(0,tk.END)
        self.spbox_minute_distribution.delete(0,tk.END)
        self.entry_quota_distribution.delete(0,tk.END)
        self.entry_remark_distribution.delete(0,tk.END)
        self.entry_berat_distribution.delete(0,tk.END)
        self.entry_noSO_distribution.delete(0,tk.END)
        self.entry_pkgRemark_distribution.delete(0,tk.END)
        self.entry_trPart_distribution.delete(0,tk.END)

        # get data and set all entry with it
        try:
            selectedItem = self.displanData_table.selection()[0]
            company = self.displanData_table.item(selectedItem)['values'][0]
            select_truck = self.displanData_table.item(selectedItem)['values'][1]
            packaging = self.displanData_table.item(selectedItem)['values'][2]
            product = self.displanData_table.item(selectedItem)['values'][3]
            date_value = self.displanData_table.item(selectedItem)['values'][4]
            time = self.displanData_table.item(selectedItem)['values'][5]
            pkg_remark = self.displanData_table.item(selectedItem)['values'][6]
            remark = self.displanData_table.item(selectedItem)['values'][7]
            no_so = self.displanData_table.item(selectedItem)['values'][8]
            tr_part = self.displanData_table.item(selectedItem)['values'][9]
            satuan = self.displanData_table.item(selectedItem)['values'][10]
            berat = self.displanData_table.item(selectedItem)['values'][11]
        except:
            pass
        # insert to entry
        try:
            self.entry_company_distribution.insert(0, company)
            self.entry_policeNo_distribution.insert(0, select_truck)
            self.optionmenu_packaging_distribution.set(packaging)
            self.optionmenu_product_distribution.set(product)
            self.date_distribution_calendar.set_date(date(int(date_value[0:4]),int(date_value[5:7]),int(date_value[8:])))
            self.spbox_hour_distribution.insert(0,time[0:2])
            self.spbox_minute_distribution.insert(0,time[3:5])
            self.entry_remark_distribution.insert(0, remark)
            self.optionmenu_warehouse_distribution.set("")
            self.entry_noSO_distribution.insert(0, no_so)
            self.entry_berat_distribution.insert(0, berat)
            self.entry_pkgRemark_distribution.insert(0, pkg_remark)
            self.entry_trPart_distribution.insert(0, tr_part)
        except:
            pass

    def on_tree_packaging_select(self, event):
        global packaging_reg
        
        self.entry_packaging.delete(0,tk.END)
        self.entry_prefix.delete(0,tk.END)
        self.entry_remark_packaging.delete(0,tk.END)
        self.entry_sla_packaging.delete(0,tk.END)

        # get data and set all entry with it
        try:
            selectedItem = self.packaging_table.selection()[0]
            packaging = self.packaging_table.item(selectedItem)['values'][0]
            packaging_reg = packaging
            prefix = self.packaging_table.item(selectedItem)['values'][1]
            remark = self.packaging_table.item(selectedItem)['values'][2]
            status = self.packaging_table.item(selectedItem)['values'][3]
            sla = self.packaging_table.item(selectedItem)['values'][4]
        except:
            pass
        # insert to entry
        try:
            self.entry_packaging.insert(0, packaging)
            self.entry_prefix.insert(0, prefix)
            self.entry_remark_packaging.insert(0, remark)
            self.entry_sla_packaging.insert(0, sla)
            self.optionmenu_status_packaging.set(status)
        except:
            pass

        self.optionmenu_packaging_sla.set(packaging)
        self.optionmenuPackaging_callback(packaging)

        # get all data where packaging selected and put to warehouse_slaSetting_table
        for record in self.warehouse_slaSetting_table.get_children():
            self.warehouse_slaSetting_table.delete(record)
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT * FROM tb_warehouseSla WHERE packaging = ?",(packaging,))
        rec = c.fetchall()
        # input to tabel GUI
        count = 0
        for i in rec:
            if count % 2 == 0:
                self.warehouse_slaSetting_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10]), tags=("evenrow",))
            else:
                self.warehouse_slaSetting_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10]), tags=("oddrow",))
            count+=1

        file.commit()
        file.close()

    def on_tree_warehouse_slaSetting_select(self, event):
        global no_wh_sla

        self.entry_weightMin.delete(0,tk.END)
        self.entry_weightMax.delete(0,tk.END)
        self.entry_warehouseMin.delete(0,tk.END)
        self.entry_waitingTime.delete(0,tk.END)
        self.entry_wBridge1Min.delete(0,tk.END)
        self.entry_wBridge2Min.delete(0,tk.END)
        self.entry_coveringMin.delete(0,tk.END)

        # get data and set all entry with it
        try:
            selectedItem = self.warehouse_slaSetting_table.selection()[0]
            packaging = self.warehouse_slaSetting_table.item(selectedItem)['values'][0]
            product_grp = self.warehouse_slaSetting_table.item(selectedItem)['values'][1]
            w_min = self.warehouse_slaSetting_table.item(selectedItem)['values'][2]
            w_max = self.warehouse_slaSetting_table.item(selectedItem)['values'][3]
            wh_sla = self.warehouse_slaSetting_table.item(selectedItem)['values'][4]
            vehicle = self.warehouse_slaSetting_table.item(selectedItem)['values'][5]
            waiting_time = self.warehouse_slaSetting_table.item(selectedItem)['values'][6]
            w_bridge1 = self.warehouse_slaSetting_table.item(selectedItem)['values'][7]
            w_bridge2 = self.warehouse_slaSetting_table.item(selectedItem)['values'][8]
            covering_time = self.warehouse_slaSetting_table.item(selectedItem)['values'][9]
            no_wh_sla = self.warehouse_slaSetting_table.item(selectedItem)['values'][10]
        except:
            pass
        # insert to entry
        try:
            self.optionmenu_packaging_sla.set(packaging)
            self.optionmenu_product_sla.set(product_grp)
            self.optionmenu_vehicle_sla.set(vehicle)
            self.entry_weightMin.insert(0, w_min)
            self.entry_weightMax.insert(0, w_max)
            self.entry_warehouseMin.insert(0, wh_sla)
            self.entry_waitingTime.insert(0, waiting_time)
            self.entry_wBridge1Min.insert(0, w_bridge1)
            self.entry_wBridge2Min.insert(0, w_bridge2)
            self.entry_coveringMin.insert(0, covering_time)
        except:
            pass

    def on_tree_product_select(self, event):
        global no_product

        self.entry_product_group.delete(0,tk.END)
        self.entry_product_group_remark.delete(0,tk.END)

        # get data and set all entry with it
        try:
            selectedItem = self.product_table.selection()[0]
            product_grp = self.product_table.item(selectedItem)['values'][0]
            remark = self.product_table.item(selectedItem)['values'][1]
            status = self.product_table.item(selectedItem)['values'][2]
            no_product = self.product_table.item(selectedItem)['values'][3]
        except:
            pass
        # insert to entry
        try:
            self.entry_product_group.insert(0, product_grp)
            self.entry_product_group_remark.insert(0, remark)
            self.optionmenu_product_group_status.set(status)
        except:
            pass

    def on_tree_tappingPoin_select(self, event):
        global no_tappingPoint, check_all, check_zak, check_curah, check_jumbo, check_pallet

        self.entry_machineId.delete(0,tk.END)
        self.entry_tapId.delete(0,tk.END)
        self.entry_remark_tappingpoint.delete(0,tk.END)
        self.entry_tapBlocking_tappingpoint.delete(0,tk.END)
        self.entry_max_tappingPoint.delete(0,tk.END)
        self.entry_aliasId.delete(0,tk.END)
        self.checkbox_packaging_tappingPoint_all.deselect()
        self.checkbox_packaging_tappingPoint_all.configure(state="normal")
        self.checkbox_packaging_tappingPoint_zak.deselect()
        self.checkbox_packaging_tappingPoint_zak.configure(state="normal")
        self.checkbox_packaging_tappingPoint_curah.deselect()
        self.checkbox_packaging_tappingPoint_curah.configure(state="normal")
        self.checkbox_packaging_tappingPoint_jumbo.deselect()
        self.checkbox_packaging_tappingPoint_jumbo.configure(state="normal")
        self.checkbox_packaging_tappingPoint_pallet.deselect()

        # get data and set all entry with it
        try:
            selectedItem = self.tappingPoint_table.selection()[0]
            hw_id = self.tappingPoint_table.item(selectedItem)['values'][0]
            tap_id = self.tappingPoint_table.item(selectedItem)['values'][1]
            remark = self.tappingPoint_table.item(selectedItem)['values'][2]
            packaging = self.tappingPoint_table.item(selectedItem)['values'][3]
            product = self.tappingPoint_table.item(selectedItem)['values'][4]
            activity = self.tappingPoint_table.item(selectedItem)['values'][5]
            block = self.tappingPoint_table.item(selectedItem)['values'][6]
            max = self.tappingPoint_table.item(selectedItem)['values'][7]
            function = self.tappingPoint_table.item(selectedItem)['values'][8]
            status = self.tappingPoint_table.item(selectedItem)['values'][9]
            alias = self.tappingPoint_table.item(selectedItem)['values'][10]
            no_tappingPoint = self.tappingPoint_table.item(selectedItem)['values'][11]
        except:
            pass
        # insert to entry
        try:
            self.entry_machineId.insert(0, "{:02d}".format(int(hw_id)))
            self.entry_tapId.insert(0, "{:02d}".format(int(tap_id)))
            self.entry_remark_tappingpoint.insert(0, remark)
            self.optionmenu_product_tappingPoint.set(product)
            self.entry_aliasId.insert(0, alias)
            self.optionmenu_activity_tappingPoint.set(activity)
            self.entry_tapBlocking_tappingpoint.insert(0, block)
            self.entry_max_tappingPoint.insert(0, max)
            self.optionmenu_function_tappingPoint.set(function)
            self.optionmenu_status_tappingPoint.set(status)
            if packaging == "ALL":
                self.checkbox_packaging_tappingPoint_all.select()
                self.checkBox_All_tappingPoint()
                check_all = False
            else:
                packaging = packaging.split('/')
                for i in packaging:
                    if i == "ZAK":
                        self.checkbox_packaging_tappingPoint_zak.select()
                        self.checkBox_ZAK_tappingPoint()
                        check_zak = False
                    elif i == "CURAH":
                        self.checkbox_packaging_tappingPoint_curah.select()
                        self.checkBox_CURAH_tappingPoint()
                        check_curah = False
                    elif i == "JUMBO":
                        self.checkbox_packaging_tappingPoint_jumbo.select()
                        self.checkBox_JUMBO_tappingPoint()
                        check_jumbo = False
                    elif i == "PALLET":
                        self.checkbox_packaging_tappingPoint_pallet.select()
                        self.checkBox_PALLET_tappingPoint()
                        check_pallet = False
                        self.checkbox_packaging_tappingPoint_pallet.configure(state="normal")
        except:
            pass

    def on_tree_warehouseFlow_select(self, event):
        global no_warehouseFlow

        self.checkbox_tappingPoint_warehouseflow_registration.deselect()
        self.checkbox_tappingPoint_warehouseflow_weightbridge1.deselect()
        self.checkbox_tappingPoint_warehouseflow_warehouse.deselect()
        self.checkbox_tappingPoint_warehouseflow_weightbridge2.deselect()
        self.checkbox_tappingPoint_warehouseflow_covering.deselect()
        self.checkbox_tappingPoint_warehouseflow_unregistration.deselect()

        # get data and set all entry with it
        try:
            selectedItem = self.warehouse_flow_table.selection()[0]
            packaging = self.warehouse_flow_table.item(selectedItem)['values'][0]
            activity = self.warehouse_flow_table.item(selectedItem)['values'][1]
            registration = self.warehouse_flow_table.item(selectedItem)['values'][2]
            weightbridge1 = self.warehouse_flow_table.item(selectedItem)['values'][3]
            warehouse = self.warehouse_flow_table.item(selectedItem)['values'][4]
            weightbridge2 = self.warehouse_flow_table.item(selectedItem)['values'][5]
            covering = self.warehouse_flow_table.item(selectedItem)['values'][6]
            unregistration = self.warehouse_flow_table.item(selectedItem)['values'][7]
            status = self.warehouse_flow_table.item(selectedItem)['values'][8]
            no_warehouseFlow = self.warehouse_flow_table.item(selectedItem)['values'][9]
        except:
            pass
        # insert to entry
        try:
            self.optionmenu_packaging_warehouseflow.set(packaging)
            self.optionmenu_activity_warehouseflow.set(activity)
            self.optionmenu_status_warehouseflow.set(status)
            if registration == "YES":
                self.checkbox_tappingPoint_warehouseflow_registration.select()
            if weightbridge1 == "YES":
                self.checkbox_tappingPoint_warehouseflow_weightbridge1.select()
            if warehouse == "YES":
                self.checkbox_tappingPoint_warehouseflow_warehouse.select()
            if weightbridge2 == "YES":
                self.checkbox_tappingPoint_warehouseflow_weightbridge2.select()
            if covering == "YES":
                self.checkbox_tappingPoint_warehouseflow_covering.select()
            if unregistration == "YES":
                self.checkbox_tappingPoint_warehouseflow_unregistration.select()
        except:
            pass

    def on_tree_rfidPairing_select(self, event):
        self.entry_noTiket_rfidPairing.delete(0,tk.END)
        # get data and set all entry with it
        try:
            selectedItem = self.rfidPairing_table.selection()[0]
            no_ticket = self.rfidPairing_table.item(selectedItem)['values'][1]
            warehouse = self.rfidPairing_table.item(selectedItem)['values'][19]
        except:
            pass
        # insert to entry
        try:
            self.entry_noTiket_rfidPairing.insert(0, no_ticket)
            self.optionmenu_warehouse_rfidPairing.set("Gudang " + str(warehouse))
        except:
            pass

    def on_tree_manualCall_select(self, event):
        self.entry_truckNo_manualCall.delete(0,tk.END)
        # get data and set all entry with it
        try:
            selectedItem = self.manualCall_table.selection()[0]
            no_polisi = self.manualCall_table.item(selectedItem)['values'][1]
        except:
            pass
        # insert to entry
        try:
            self.entry_truckNo_manualCall.insert(0, no_polisi)
        except:
            pass

    def on_tree_alert_select(self, event):
        self.alertMessage_textbox.delete('1.0', tk.END)
        self.notes_textbox.delete('1.0', tk.END)
        try:
            selectedItem = self.alert_table.selection()[0]
            message_text = self.alert_table.item(selectedItem)['values'][4]
            notes_text = self.alert_table.item(selectedItem)['values'][5]
        except:
            pass
        self.alertMessage_textbox.insert("insert", text=message_text, tags=None)
        self.notes_textbox.insert("insert", text=notes_text, tags=None)

    def on_tree_incident_report_select(self, event):
        self.report_viewer_incident_reportRight_textbox_remark.delete('1.0', tk.END)
        self.report_viewer_incident_reportRight_textbox_note.delete('1.0', tk.END)
        try:
            selectedItem = self.report_viewer_incident_report_table.selection()[0]
            remark_text = self.report_viewer_incident_report_table.item(selectedItem)['values'][11]
            notes_text = self.report_viewer_incident_report_table.item(selectedItem)['values'][12]
        except:
            pass
        self.report_viewer_incident_reportRight_textbox_remark.insert("insert", text=remark_text, tags=None)
        self.report_viewer_incident_reportRight_textbox_note.insert("insert", text=notes_text, tags=None)

    def on_tree_userRegistration_select(self, event):
        global no_user

        self.entry_regis_username.delete(0,tk.END)
        self.entry_regis_password.delete(0,tk.END)
        self.entry_regis_retypePassword.delete(0,tk.END)
        self.entry_regis_email.delete(0,tk.END)
        self.entry_regis_remark.delete(0,tk.END)

        self.checkbox_alertNoteMenu.deselect()
        self.checkbox_reportViewerMenu.deselect()
        self.checkbox_packagingMenu.deselect()
        self.checkbox_productMenu.deselect()
        self.checkbox_tappingPointMenu.deselect()
        self.checkbox_warehouseFlowMenu.deselect()
        self.checkbox_rfidManualCallMenu.deselect()
        self.checkbox_vettingMenu.deselect()
        self.checkbox_unloadingMenu.deselect()
        self.checkbox_distributionMenu.deselect()
        self.checkbox_databaseSettingMenu.deselect()
        self.checkbox_displanApiMenu.deselect()
        self.checkbox_otherSettingMenu.deselect()
        self.checkbox_userProfileMenu.deselect()
        self.checkbox_userRegistrationMenu.deselect()

        # get data and set all entry with it
        try:
            selectedItem = self.userRegistration_table.selection()[0]
            no_user = self.userRegistration_table.item(selectedItem)['values'][0]
            username_app = self.userRegistration_table.item(selectedItem)['values'][1]
            password_app = self.userRegistration_table.item(selectedItem)['values'][4]
            email = self.userRegistration_table.item(selectedItem)['values'][2]
            remark = self.userRegistration_table.item(selectedItem)['values'][3]

            alertNoteMenu = self.userRegistration_table.item(selectedItem)['values'][5]
            reportViewerMenu = self.userRegistration_table.item(selectedItem)['values'][6]
            packagingMenu = self.userRegistration_table.item(selectedItem)['values'][7]
            productMenu = self.userRegistration_table.item(selectedItem)['values'][8]
            tappingPointMenu = self.userRegistration_table.item(selectedItem)['values'][9]
            warehouseFlowMenu = self.userRegistration_table.item(selectedItem)['values'][10]
            rfidManualCallMenu = self.userRegistration_table.item(selectedItem)['values'][11]
            vettingMenu = self.userRegistration_table.item(selectedItem)['values'][12]
            unloadingMenu = self.userRegistration_table.item(selectedItem)['values'][13]
            distributionMenu = self.userRegistration_table.item(selectedItem)['values'][14]
            databaseSettingMenu = self.userRegistration_table.item(selectedItem)['values'][15]
            displanApiMenu = self.userRegistration_table.item(selectedItem)['values'][16]
            otherSettingMenu = self.userRegistration_table.item(selectedItem)['values'][17]
            userProfileMenu = self.userRegistration_table.item(selectedItem)['values'][18]
            userRegistrationMenu = self.userRegistration_table.item(selectedItem)['values'][19]
        except:
            pass
        # insert to entry
        try:
            self.entry_regis_username.insert(0, username_app)
            self.entry_regis_password.insert(0, password_app)
            self.entry_regis_retypePassword.insert(0, password_app)
            self.entry_regis_email.insert(0, email)
            self.entry_regis_remark.insert(0, remark)

            if alertNoteMenu == "True":
                self.checkbox_alertNoteMenu.select()
            if reportViewerMenu == "True":
                self.checkbox_reportViewerMenu.select()
            if packagingMenu == "True":
                self.checkbox_packagingMenu.select()
            if productMenu == "True":
                self.checkbox_productMenu.select()
            if tappingPointMenu == "True":
                self.checkbox_tappingPointMenu.select()
            if warehouseFlowMenu == "True":
                self.checkbox_warehouseFlowMenu.select()
            if rfidManualCallMenu == "True":
                self.checkbox_rfidManualCallMenu.select()
            if vettingMenu == "True":
                self.checkbox_vettingMenu.select()
            if unloadingMenu == "True":
                self.checkbox_unloadingMenu.select()
            if distributionMenu == "True":
                self.checkbox_distributionMenu.select()
            if databaseSettingMenu == "True":
                self.checkbox_databaseSettingMenu.select()
            if displanApiMenu == "True":
                self.checkbox_displanApiMenu.select()
            if otherSettingMenu == "True":
                self.checkbox_otherSettingMenu.select()
            if userProfileMenu == "True":
                self.checkbox_userProfileMenu.select()
            if userRegistrationMenu == "True":
                self.checkbox_userRegistrationMenu.select()
        except:
            pass

    def on_tree_auditlog_select(self, event):
        self.report_viewer_auditlogRight_textbox_remark.delete('1.0', tk.END)
        try:
            selectedItem = self.report_viewer_auditlog_table.selection()[0]
            remark_text = self.report_viewer_auditlog_table.item(selectedItem)['values'][6]
        except:
            pass
        self.report_viewer_auditlogRight_textbox_remark.insert("insert", text=remark_text, tags=None)

    # FUNGSI TOMBOL CRUD

    def add_vetting(self):
        # input to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # get data from input
        company = self.entry_company.get()
        no_polisi = self.entry_policeNo.get()
        reg_date = self.regdate_calendar.get_date()
        expired_date = self.expireddate_calendar.get_date()
        remark = self.entry_remark.get()
        status = self.optionmenu_status.get()
        rfid1 = self.entry_rfid1.get()
        rfid2 = self.entry_rfid2.get()
        category = self.optionmenu_category.get()
        # write data to database
        c.execute("INSERT INTO tb_vetting (company,no_polisi,reg_date,expired_date,remark,status,rfid1,rfid2,category) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            company,
            no_polisi,
            reg_date,
            expired_date,
            remark,
            status,
            rfid1,
            rfid2,
            category
        )
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in self.vetting_table.get_children():
            self.vetting_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_vetting")
        rec = c.fetchall()
        urutan = 1
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.vetting_table.insert(parent="",index="end", iid=count, text="", values=(i[9], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
            else:
                self.vetting_table.insert(parent="",index="end", iid=count, text="", values=(i[9], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
            count+=1
            urutan+=1
        # reset input
        self.entry_company.delete(0,tk.END)
        self.entry_policeNo.delete(0,tk.END)
        self.entry_remark.delete(0,tk.END)
        self.entry_rfid1.delete(0,tk.END)
        self.entry_rfid2.delete(0,tk.END)

        file.commit()
        file.close()

    def add_packaging(self):
        confirmation = messagebox.askyesno("Confirmation", "Are you sure you want to add the data?")

        if confirmation:
            # input to database
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # get data from input
            packaging = self.entry_packaging.get()
            prefix = self.entry_prefix.get()
            remark = self.entry_remark_packaging.get()
            status = self.optionmenu_status_packaging.get()
            sla = self.entry_sla_packaging.get()
            # write data to database
            c.execute("INSERT INTO tb_packagingReg (packaging,prefix,remark,status,sla) VALUES (?,?,?,?,?)",
                packaging,
                prefix,
                remark,
                status,
                sla
            )
            file.commit()
            file.close()
            # hapus item di tabel GUI
            for record in self.packaging_table.get_children():
                self.packaging_table.delete(record)
            # get data from database
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            count = 0
            c.execute("SELECT * FROM tb_packagingReg")
            rec = c.fetchall()
            # input to tabel GUI
            for i in rec:
                if count % 2 == 0:
                    self.packaging_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4]), tags=("evenrow",))
                else:
                    self.packaging_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4]), tags=("oddrow",))
                count+=1
            # reset input
            self.entry_packaging.delete(0,tk.END)
            self.entry_prefix.delete(0,tk.END)
            self.entry_remark_packaging.delete(0,tk.END)
            self.entry_sla_packaging.delete(0,tk.END)

            file.commit()
            file.close()

            self.optionMenu_packaging()

            # save log
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            date_log = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
            time_log = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
            transaction_log = "AUDIT LOG"
            status = "DONE"
            remark = f"Add item in packaging registration: {packaging}"
            # write data to database
            c.execute("INSERT INTO tb_log_report (warehouse,date_log,time_log,transaction_log,operator,status,remark) VALUES (?,?,?,?,?,?,?)",
                wh_name,
                date_log,
                time_log,
                transaction_log,
                username_log,
                status,
                remark
            )
            file.commit()
            file.close()

    def add_warehouse_slaSetting(self):
        confirmation = messagebox.askyesno("Confirmation", "Are you sure you want to add the data?")

        if confirmation:
            # input to database
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # get data from input
            packaging = self.optionmenu_packaging_sla.get()
            product_grp = self.optionmenu_product_sla.get()
            weight_min = self.entry_weightMin.get()
            weight_max = self.entry_weightMax.get()
            wh_sla = self.entry_warehouseMin.get()
            vehicle = self.optionmenu_vehicle_sla.get()
            waiting_time = self.entry_waitingTime.get()
            w_bridge1 = self.entry_wBridge1Min.get()
            w_bridge2 = self.entry_wBridge2Min.get()
            covering_time = self.entry_coveringMin.get()
            # write data to database
            c.execute("INSERT INTO tb_warehouseSla (packaging,product_grp,weight_min,weight_max,wh_sla,vehicle,waiting_time,w_bridge1,w_bridge2,covering_time) VALUES (?,?,?,?,?,?,?,?,?,?)",
                packaging,
                product_grp,
                weight_min,
                weight_max,
                wh_sla,
                vehicle,
                waiting_time,
                w_bridge1,
                w_bridge2,
                covering_time
            )
            file.commit()
            file.close()
            # hapus item di tabel GUI
            for record in self.warehouse_slaSetting_table.get_children():
                self.warehouse_slaSetting_table.delete(record)
            # get data from database
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            count = 0
            c.execute("SELECT * FROM tb_warehouseSla WHERE packaging = ?",(packaging,))
            rec = c.fetchall()
            # input to tabel GUI
            for i in rec:
                if count % 2 == 0:
                    self.warehouse_slaSetting_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10]), tags=("evenrow",))
                else:
                    self.warehouse_slaSetting_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10]), tags=("oddrow",))
                count+=1
            # reset input
            self.entry_weightMin.delete(0,tk.END)
            self.entry_weightMax.delete(0,tk.END)
            self.entry_warehouseMin.delete(0,tk.END)
            self.entry_waitingTime.delete(0,tk.END)
            self.entry_wBridge1Min.delete(0,tk.END)
            self.entry_wBridge2Min.delete(0,tk.END)
            self.entry_coveringMin.delete(0,tk.END)

            file.commit()
            file.close()

            # save log
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            date_log = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
            time_log = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
            transaction_log = "AUDIT LOG"
            status = "DONE"
            remark = f"Add data in warehouse SLA Setting: {packaging}, {product_grp}, Weight Min:{weight_min} ,Weight Max:{weight_max}"
            # write data to database
            c.execute("INSERT INTO tb_log_report (warehouse,date_log,time_log,transaction_log,operator,status,remark) VALUES (?,?,?,?,?,?,?)",
                wh_name,
                date_log,
                time_log,
                transaction_log,
                username_log,
                status,
                remark
            )
            file.commit()
            file.close()

    def add_distribution_loadingList(self):
        # input to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # get data from input
        company = self.entry_company_distribution.get()
        no_polisi = self.entry_policeNo_distribution.get()
        packaging = self.optionmenu_packaging_distribution.get()
        product = self.optionmenu_product_distribution.get()
        date = self.date_distribution_calendar.get_date()
        time_hour = self.spbox_hour_distribution.get()
        time_minute = self.spbox_minute_distribution.get()
        time = f"{time_hour}:{time_minute}"
        remark = self.entry_remark_distribution.get()
        status = self.optionmenu_distribution_status.get()
        quota = self.entry_quota_distribution.get()
        no_so = self.entry_noSO_distribution.get()
        satuan = "KG"
        berat = self.entry_berat_distribution.get()
        wh_id = self.optionmenu_warehouse_distribution.get()
        # write data to database
        if wh_id != "Otomatis":
            c.execute("INSERT INTO tb_loadingList (company,no_polisi,packaging,product,date,time,remark,status,quota,no_so,satuan,berat,wh_id) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                company,
                no_polisi,
                packaging,
                product,
                date,
                time,
                remark,
                status,
                quota,
                no_so,
                satuan,
                berat,
                "0"+wh_id[7]
            )
        else:
            c.execute("INSERT INTO tb_loadingList (company,no_polisi,packaging,product,date,time,remark,status,quota,no_so,satuan,berat,wh_id) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                company,
                no_polisi,
                packaging,
                product,
                date,
                time,
                remark,
                status,
                quota,
                no_so,
                satuan,
                berat,
                "Auto"
            )
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in self.loadingList_table.get_children():
            self.loadingList_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_loadingList")
        rec = c.fetchall()
        # input to tabel GUI
        urutan = 1
        for i in rec:
            if count % 2 == 0:
                self.loadingList_table.insert(parent="",index="end", iid=count, text="", values=(i[13], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12]), tags=("evenrow",))
            else:
                self.loadingList_table.insert(parent="",index="end", iid=count, text="", values=(i[13], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12]), tags=("oddrow",))
            count+=1
            urutan+=1
        # reset input
        self.entry_company_distribution.delete(0,tk.END)
        self.entry_policeNo_distribution.delete(0,tk.END)
        self.optionmenu_packaging_distribution.set("Pilih Packaging")
        self.optionmenu_product_distribution.set("Pilih Product")
        self.entry_quota_distribution.delete(0,tk.END)
        self.entry_quota_distribution.insert(0,"0")
        self.entry_remark_distribution.delete(0,tk.END)
        self.entry_remark_distribution.insert(0,"LOADING")
        self.entry_berat_distribution.delete(0,tk.END)
        self.entry_noSO_distribution.delete(0,tk.END)

        file.commit()
        file.close()

    def add_unloading(self):
        # input to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # get data from input
        company = self.entry_company_unloading.get()
        no_polisi = self.entry_policeNo_unloading.get()
        packaging = self.optionmenu_packaging.get()
        product = self.optionmenu_product.get()
        date = self.date_unloading_calendar.get_date()
        time_hour = self.spbox_hour.get()
        time_minute = self.spbox_minute.get()
        time = f"{time_hour}:{time_minute}"
        remark = self.entry_remark_unloading.get()
        status = self.optionmenu_unloading_status.get()
        quota = self.entry_quota_unloading.get()
        berat = self.entry_berat_unloading.get()
        wh_id = self.optionmenu_warehouse.get()
        no_so = self.optionmenu_noSo.get()
        # write data to database
        if wh_id != "Otomatis":
            c.execute("INSERT INTO tb_unloading (company,no_polisi,packaging,product,date,time,remark,status,quota,berat,wh_id,no_so) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                company,
                no_polisi,
                packaging,
                product,
                date,
                time,
                remark,
                status,
                quota,
                berat,
                "0"+wh_id[7],
                no_so
            )
        else:
            c.execute("INSERT INTO tb_unloading (company,no_polisi,packaging,product,date,time,remark,status,quota,berat,wh_id,no_so) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                company,
                no_polisi,
                packaging,
                product,
                date,
                time,
                remark,
                status,
                quota,
                berat,
                "Auto",
                no_so
            )
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in self.unloading_table.get_children():
            self.unloading_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_unloading")
        rec = c.fetchall()
        # input to tabel GUI
        urutan = 1
        for i in rec:
            if count % 2 == 0:
                self.unloading_table.insert(parent="",index="end", iid=count, text="", values=(i[12], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("evenrow",))
            else:
                self.unloading_table.insert(parent="",index="end", iid=count, text="", values=(i[12], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("oddrow",))
            count+=1
            urutan+=1
        # reset input
        self.entry_company_unloading.delete(0,tk.END)
        self.entry_policeNo_unloading.delete(0,tk.END)
        self.optionmenu_packaging.set("Pilih Packaging")
        self.optionmenu_product.set("Pilih Product")
        self.entry_quota_unloading.delete(0,tk.END)
        self.entry_quota_unloading.insert(0,"0")
        self.entry_remark_unloading.delete(0,tk.END)
        self.entry_berat_unloading.delete(0,tk.END)
        self.entry_remark_unloading.insert(0,"UNLOADING")
        self.optionmenu_noSo.delete(0,tk.END)

        file.commit()
        file.close()

    def add_distribution_displanData(self):
        # input to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # get data from input
        company = self.entry_company_distribution.get()
        truck = self.entry_policeNo_distribution.get()
        packaging = self.optionmenu_packaging_distribution.get()
        product = self.optionmenu_product_distribution.get()
        date = self.date_distribution_calendar.get_date()
        time_hour = self.spbox_hour_distribution.get()
        time_minute = self.spbox_minute_distribution.get()
        time = f"{time_hour}:{time_minute}"
        remark = self.entry_remark_distribution.get()
        no_so = self.entry_noSO_distribution.get()
        satuan = "KG"
        berat = self.entry_berat_distribution.get()
        pkg_remark = self.entry_pkgRemark_distribution.get()
        tr_part = self.entry_trPart_distribution.get()
        # write data to database
        c.execute("INSERT INTO tb_displanData (company,truck,packaging,product,date,time,pkg_remark,remark,no_so,tr_part,satuan,berat) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
            company,
            truck,
            packaging,
            product,
            date,
            time,
            pkg_remark,
            remark,
            no_so,
            tr_part,
            satuan,
            berat,
        )
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in self.displanData_table.get_children():
            self.displanData_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_displanData")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.displanData_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("evenrow",))
            else:
                self.displanData_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("oddrow",))
            count+=1
        # reset input
        self.entry_company_distribution.delete(0,tk.END)
        self.entry_policeNo_distribution.delete(0,tk.END)
        self.optionmenu_packaging_distribution.set("Pilih Packaging")
        self.optionmenu_product_distribution.set("Pilih Product")
        self.entry_quota_distribution.delete(0,tk.END)
        self.entry_quota_distribution.insert(0,"0")
        self.entry_remark_distribution.delete(0,tk.END)
        self.entry_remark_distribution.insert(0,"LOADING")
        self.entry_berat_distribution.delete(0,tk.END)
        self.entry_noSO_distribution.delete(0,tk.END)
        self.entry_pkgRemark_distribution.delete(0,tk.END)
        self.entry_trPart_distribution.delete(0,tk.END)

        file.commit()
        file.close()

        # add data to tb_displan_data_container also
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("INSERT INTO tb_displan_data_container (company,truck,packaging,product,date,time,pkg_remark,remark,no_so,tr_part,satuan,berat) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
            company,
            truck,
            packaging,
            product,
            date,
            time,
            pkg_remark,
            remark,
            no_so,
            tr_part,
            satuan,
            berat,
        )
        file.commit()
        file.close()

    def add_product_registration(self):
        # input to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # get data from input
        packaging = self.optionmenu_packaging_product.get()
        product_grp = self.entry_product_group.get()
        remark =  self.entry_product_group_remark.get()
        status = self.optionmenu_product_group_status.get()
        # write data to database
        c.execute("INSERT INTO tb_productReg (packaging,product_grp,remark,status) VALUES (?,?,?,?)",
            packaging,
            product_grp,
            remark,
            status
        )
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in self.product_table.get_children():
            self.product_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_productReg WHERE packaging = ?",(packaging))
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.product_table.insert(parent="",index="end", iid=count, text="", values=(i[1], i[2], i[3], i[4]), tags=("evenrow",))
            else:
                self.product_table.insert(parent="",index="end", iid=count, text="", values=(i[1], i[2], i[3], i[4]), tags=("oddrow",))
            count+=1
        # reset input
        self.entry_product_group.delete(0,tk.END)
        self.entry_product_group_remark.delete(0,tk.END)

        file.commit()
        file.close()

    def add_tapping_point(self):
        global check_all
        # input to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # get data from input
        hw_id = self.entry_machineId.get()
        tap_id = self.entry_tapId.get()
        remark = self.entry_remark_tappingpoint.get()
        packaging = self.checkBox_get_tappingPoint()
        product = self.optionmenu_product_tappingPoint.get()
        activity = self.optionmenu_activity_tappingPoint.get()
        block = self.entry_tapBlocking_tappingpoint.get()
        max_value = self.entry_max_tappingPoint.get()
        function = self.optionmenu_function_tappingPoint.get()
        status = self.optionmenu_status_tappingPoint.get()
        alias = self.entry_aliasId.get()
        # write data to database
        c.execute("INSERT INTO tb_tappingPointReg (hw_id,tap_id,remark,packaging,product,activity,block,max,functions,status,alias) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            hw_id,
            tap_id,
            remark,
            packaging,
            product,
            activity,
            block,
            max_value,
            function,
            status,
            alias
        )
        file.commit()
        file.close()

        # hapus item di tabel GUI
        for record in self.tappingPoint_table.get_children():
            self.tappingPoint_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_tappingPointReg")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.tappingPoint_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("evenrow",))
            else:
                self.tappingPoint_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("oddrow",))
            count+=1
        # reset input
        self.entry_machineId.delete(0,tk.END)
        self.entry_tapId.delete(0,tk.END)
        self.entry_remark_tappingpoint.delete(0,tk.END)
        self.entry_tapBlocking_tappingpoint.delete(0,tk.END)
        self.entry_max_tappingPoint.delete(0,tk.END)
        self.entry_aliasId.delete(0,tk.END)
        self.checkbox_packaging_tappingPoint_all.deselect()
        self.checkbox_packaging_tappingPoint_all.configure(state="normal")
        self.checkbox_packaging_tappingPoint_zak.deselect()
        self.checkbox_packaging_tappingPoint_zak.configure(state="normal")
        self.checkbox_packaging_tappingPoint_curah.deselect()
        self.checkbox_packaging_tappingPoint_curah.configure(state="normal")
        self.checkbox_packaging_tappingPoint_jumbo.deselect()
        self.checkbox_packaging_tappingPoint_jumbo.configure(state="normal")
        self.checkbox_packaging_tappingPoint_pallet.deselect()
        self.checkbox_packaging_tappingPoint_pallet.configure(state="normal")
        check_all = False

        file.commit()
        file.close()

        self.optionMenu_warehouseAssign()

    def add_warehouse_flow(self):
        # input to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # get data from input
        packaging = self.optionmenu_packaging_warehouseflow.get()
        activity = self.optionmenu_activity_warehouseflow.get()
        status = self.optionmenu_status_warehouseflow.get()
        registration = self.checkbox_tappingPoint_warehouseflow_registration.get()
        if registration == 0:
            registration = "NO"
        weightbridge1 = self.checkbox_tappingPoint_warehouseflow_weightbridge1.get()
        if weightbridge1 == 0:
            weightbridge1 = "NO"
        warehouse = self.checkbox_tappingPoint_warehouseflow_warehouse.get()
        if warehouse == 0:
            warehouse = "NO"
        weightbridge2 = self.checkbox_tappingPoint_warehouseflow_weightbridge2.get()
        if weightbridge2 == 0:
            weightbridge2 = "NO"
        covering = self.checkbox_tappingPoint_warehouseflow_covering.get()
        if covering == 0:
            covering = "NO"
        unregistration = self.checkbox_tappingPoint_warehouseflow_unregistration.get()
        if unregistration == 0:
            unregistration = "NO"
        # write data to database
        c.execute("INSERT INTO tb_warehouseFlow (packaging,activity,registration,weightbridge1,warehouse,weightbridge2,covering,unregistration,status) VALUES (?,?,?,?,?,?,?,?,?)",
            packaging,
            activity,
            registration,
            weightbridge1,
            warehouse,
            weightbridge2,
            covering,
            unregistration,
            status
        )
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in self.warehouse_flow_table.get_children():
            self.warehouse_flow_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_warehouseFlow")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.warehouse_flow_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("evenrow",))
            else:
                self.warehouse_flow_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("oddrow",))
            count+=1
        # reset input
        self.checkbox_tappingPoint_warehouseflow_registration.deselect()
        self.checkbox_tappingPoint_warehouseflow_weightbridge1.deselect()
        self.checkbox_tappingPoint_warehouseflow_warehouse.deselect()
        self.checkbox_tappingPoint_warehouseflow_weightbridge2.deselect()
        self.checkbox_tappingPoint_warehouseflow_covering.deselect()
        self.checkbox_tappingPoint_warehouseflow_unregistration.deselect()

        file.commit()
        file.close()

    def add_user_registration(self):
        username_app = self.entry_regis_username.get()
        password_app = self.entry_regis_password.get()
        retypePassword = self.entry_regis_retypePassword.get()
        email = self.entry_regis_email.get()
        remark = self.entry_regis_remark.get()

        alertNoteMenu = self.checkbox_alertNoteMenu.get()
        reportViewerMenu = self.checkbox_reportViewerMenu.get()
        packagingMenu = self.checkbox_packagingMenu.get()
        productMenu = self.checkbox_productMenu.get()
        tappingPointMenu = self.checkbox_tappingPointMenu.get()
        warehouseFlowMenu = self.checkbox_warehouseFlowMenu.get()
        rfidManualCallMenu = self.checkbox_rfidManualCallMenu.get()
        vettingMenu = self.checkbox_vettingMenu.get()
        unloadingMenu = self.checkbox_unloadingMenu.get()
        distributionMenu = self.checkbox_distributionMenu.get()
        databaseSettingMenu = self.checkbox_databaseSettingMenu.get()
        displanApiMenu = self.checkbox_displanApiMenu.get()
        otherSettingMenu = self.checkbox_otherSettingMenu.get()
        userProfileMenu = self.checkbox_userProfileMenu.get()
        userRegistrationMenu = self.checkbox_userRegistrationMenu.get()

        if password_app == retypePassword:
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()

            # check apakah username sudah ada atau belum
            c.execute("SELECT username FROM tb_userRegistration where username = ?",(username_app,))
            existing_username = c.fetchone()
            # save data
            if existing_username == None:
                c.execute("INSERT INTO tb_userRegistration (username,password,email,remark,alertNote,reportViewer,packaging,productReg,tappingPoint,warehouseFlow,rfidManualCall,vetting, unloading, distribution, databaseSetting, displanApi, otherSetting, userProfile, userRegistration) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    username_app,
                    password_app,
                    email,
                    remark,
                    alertNoteMenu,
                    reportViewerMenu,
                    packagingMenu,
                    productMenu,
                    tappingPointMenu,
                    warehouseFlowMenu,
                    rfidManualCallMenu,
                    vettingMenu,
                    unloadingMenu,
                    distributionMenu,
                    databaseSettingMenu,
                    displanApiMenu,
                    otherSettingMenu,
                    userProfileMenu,
                    userRegistrationMenu,
                )
                file.commit()
                file.close()

                self.entry_regis_username.delete(0,tk.END)
                self.entry_regis_password.delete(0,tk.END)
                self.entry_regis_retypePassword.delete(0,tk.END)
                self.entry_regis_email.delete(0,tk.END)
                self.entry_regis_remark.delete(0,tk.END)

                self.checkbox_alertNoteMenu.deselect()
                self.checkbox_reportViewerMenu.deselect()
                self.checkbox_packagingMenu.deselect()
                self.checkbox_productMenu.deselect()
                self.checkbox_tappingPointMenu.deselect()
                self.checkbox_warehouseFlowMenu.deselect()
                self.checkbox_rfidManualCallMenu.deselect()
                self.checkbox_vettingMenu.deselect()
                self.checkbox_unloadingMenu.deselect()
                self.checkbox_distributionMenu.deselect()
                self.checkbox_databaseSettingMenu.deselect()
                self.checkbox_displanApiMenu.deselect()
                self.checkbox_otherSettingMenu.deselect()
                self.checkbox_userProfileMenu.deselect()
                self.checkbox_userRegistrationMenu.deselect()

                # hapus item di tabel GUI
                for record in self.userRegistration_table.get_children():
                    self.userRegistration_table.delete(record)
                # get data from database
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                count = 0
                c.execute("SELECT * FROM tb_userRegistration")
                rec = c.fetchall()
                # input to tabel GUI
                for i in rec:
                    if count % 2 == 0:
                        self.userRegistration_table.insert(parent="",index="end", iid=count, text="", values=(i[19], i[0], i[2], i[3], i[1], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18]), tags=("evenrow",))
                    else:
                        self.userRegistration_table.insert(parent="",index="end", iid=count, text="", values=(i[19], i[0], i[2], i[3], i[1], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18]), tags=("oddrow",))
                    count+=1
                file.commit()
                file.close()

                messagebox.showinfo("Sukses", "Data telah tersimpan")
            else:
                messagebox.showerror("Error", "Username sudah tersedia\nSilahkan menggunakan username lain")
        else:
            messagebox.showerror("Error", "Retype password is incorrect.")

    def remove_vetting(self):
        # remove to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("DELETE from tb_vetting where oid = ?",(no,))
        file.commit()
        file.close()
        # hapus item di tabel GUI
        x = self.vetting_table.selection()
        for record in x:
            self.vetting_table.delete(record)
        # read data from database
        for record in self.vetting_table.get_children():
            self.vetting_table.delete(record)
        count = 0
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT * FROM tb_vetting")
        # input to tabel GUI
        rec = c.fetchall()
        urutan = 1
        for i in rec:
            if count % 2 == 0:
                self.vetting_table.insert(parent="",index="end", iid=count, text="", values=(i[9], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
            else:
                self.vetting_table.insert(parent="",index="end", iid=count, text="", values=(i[9], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
            count+=1
            urutan+=1

        file.commit()
        file.close()

        # reset input
        self.entry_company.delete(0,tk.END)
        self.entry_policeNo.delete(0,tk.END)
        self.entry_remark.delete(0,tk.END)
        self.entry_rfid1.delete(0,tk.END)
        self.entry_rfid2.delete(0,tk.END)

    def remove_packaging(self):
        # Display confirmation popup
        confirmation = messagebox.askyesno("Confirmation", "Are you sure you want to delete the data?")

        if confirmation:
            # remove to database
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("DELETE from tb_packagingReg where packaging = ?",(packaging_reg,))
            file.commit()
            file.close()
            # hapus item di tabel GUI
            for record in self.packaging_table.get_children():
                self.packaging_table.delete(record)
            # get data from database
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            count = 0
            c.execute("SELECT * FROM tb_packagingReg")
            rec = c.fetchall()
            # input to tabel GUI
            for i in rec:
                if count % 2 == 0:
                    self.packaging_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4]), tags=("evenrow",))
                else:
                    self.packaging_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4]), tags=("oddrow",))
                count+=1
            # reset input
            self.entry_packaging.delete(0,tk.END)
            self.entry_prefix.delete(0,tk.END)
            self.entry_remark_packaging.delete(0,tk.END)
            self.entry_sla_packaging.delete(0,tk.END)

            file.commit()
            file.close()

            # save log
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            date_log = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
            time_log = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
            transaction_log = "AUDIT LOG"
            status = "DONE"
            remark = f"Remove item in packaging registration: {packaging_reg}"
            # write data to database
            c.execute("INSERT INTO tb_log_report (warehouse,date_log,time_log,transaction_log,operator,status,remark) VALUES (?,?,?,?,?,?,?)",
                wh_name,
                date_log,
                time_log,
                transaction_log,
                username_log,
                status,
                remark
            )
            file.commit()
            file.close()

    def remove_warehouse_slaSetting(self):
        confirmation = messagebox.askyesno("Confirmation", "Are you sure you want to remove the data?")

        if confirmation:
            packaging = self.optionmenu_packaging_sla.get()
            # remove to database
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("DELETE from tb_warehouseSla where oid = ?",(no_wh_sla,))
            file.commit()
            file.close()
            # hapus item di tabel GUI
            for record in self.warehouse_slaSetting_table.get_children():
                self.warehouse_slaSetting_table.delete(record)
            # get data from database
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            count = 0
            c.execute("SELECT * FROM tb_warehouseSla WHERE packaging = ?",(packaging,))
            rec = c.fetchall()
            # input to tabel GUI
            for i in rec:
                if count % 2 == 0:
                    self.warehouse_slaSetting_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10]), tags=("evenrow",))
                else:
                    self.warehouse_slaSetting_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10]), tags=("oddrow",))
                count+=1
            # reset input
            self.entry_weightMin.delete(0,tk.END)
            self.entry_weightMax.delete(0,tk.END)
            self.entry_warehouseMin.delete(0,tk.END)
            self.entry_waitingTime.delete(0,tk.END)
            self.entry_wBridge1Min.delete(0,tk.END)
            self.entry_wBridge2Min.delete(0,tk.END)
            self.entry_coveringMin.delete(0,tk.END)

            file.commit()
            file.close()

            # save log
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            date_log = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
            time_log = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
            transaction_log = "AUDIT LOG"
            status = "DONE"
            remark = f"Remove data in warehouse SLA setting"
            # write data to database
            c.execute("INSERT INTO tb_log_report (warehouse,date_log,time_log,transaction_log,operator,status,remark) VALUES (?,?,?,?,?,?,?)",
                wh_name,
                date_log,
                time_log,
                transaction_log,
                username_log,
                status,
                remark
            )
            file.commit()
            file.close()

    def remove_unloading(self):
        # remove to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("DELETE from tb_unloading where oid = ?",(no_unloading,))
        file.commit()
        file.close()
        # hapus item di tabel GUI
        x = self.unloading_table.selection()
        for record in x:
            self.unloading_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_unloading")
        rec = c.fetchall()
        # input to tabel GUI
        urutan = 1
        for i in rec:
            if count % 2 == 0:
                self.unloading_table.insert(parent="",index="end", iid=count, text="", values=(i[12], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("evenrow",))
            else:
                self.unloading_table.insert(parent="",index="end", iid=count, text="", values=(i[12], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("oddrow",))
            count+=1
            urutan+=1

        file.commit()
        file.close()

        # reset input
        self.entry_company_unloading.delete(0,tk.END)
        self.entry_policeNo_unloading.delete(0,tk.END)
        self.optionmenu_packaging.set("Pilih Packaging")
        self.optionmenu_product.set("Pilih Product")
        self.entry_quota_unloading.delete(0,tk.END)
        self.entry_quota_unloading.insert(0,"0")
        self.entry_remark_unloading.delete(0,tk.END)
        self.entry_berat_unloading.delete(0,tk.END)
        self.entry_remark_unloading.insert(0,"UNLOADING")
        self.optionmenu_noSo.delete(0,tk.END)

    def remove_distribution(self):
        # remove to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("DELETE from tb_loadingList where oid = ?",(no_loadingList,))
        file.commit()
        file.close()
        # hapus item di tabel GUI
        x = self.loadingList_table.selection()
        for record in x:
            self.loadingList_table.delete(record)
        # data tab distribution -> loading list
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_loadingList")
        rec = c.fetchall()
        # input to tabel GUI
        urutan = 1
        for i in rec:
            if count % 2 == 0:
                self.loadingList_table.insert(parent="",index="end", iid=count, text="", values=(i[13], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12]), tags=("evenrow",))
            else:
                self.loadingList_table.insert(parent="",index="end", iid=count, text="", values=(i[13], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12]), tags=("oddrow",))
            count+=1
            urutan+=1

        file.commit()
        file.close()

        # reset input
        self.entry_company_distribution.delete(0,tk.END)
        self.entry_policeNo_distribution.delete(0,tk.END)
        self.optionmenu_packaging_distribution.set("Pilih Packaging")
        self.optionmenu_product_distribution.set("Pilih Product")
        self.entry_quota_distribution.delete(0,tk.END)
        self.entry_quota_distribution.insert(0,"0")
        self.entry_remark_distribution.delete(0,tk.END)
        self.entry_remark_distribution.insert(0,"LOADING")

    def remove_product_registration(self):
        # remove to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("DELETE from tb_productReg where oid = ?",(no_product,))
        file.commit()
        file.close()

        # global packaging_product
        packaging_product = self.optionmenu_packaging_product.get()

        # hapus item entry product tabel
        for record in self.product_table.get_children():
            self.product_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_productReg WHERE packaging = ?",(packaging_product,))
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.product_table.insert(parent="",index="end", iid=count, text="", values=(i[1], i[2], i[3], i[4]), tags=("evenrow",))
            else:
                self.product_table.insert(parent="",index="end", iid=count, text="", values=(i[1], i[2], i[3], i[4]), tags=("oddrow",))
            count+=1
        # reset input
        self.entry_product_group.delete(0,tk.END)
        self.entry_product_group_remark.delete(0,tk.END)

        file.commit()
        file.close()

    def remove_tapping_point(self):
        # remove to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("DELETE from tb_tappingPointReg where oid = ?",(no_tappingPoint,))
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in self.tappingPoint_table.get_children():
            self.tappingPoint_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_tappingPointReg")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.tappingPoint_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("evenrow",))
            else:
                self.tappingPoint_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("oddrow",))
            count+=1
        # reset input
        self.entry_machineId.delete(0,tk.END)
        self.entry_tapId.delete(0,tk.END)
        self.entry_remark_tappingpoint.delete(0,tk.END)
        self.entry_tapBlocking_tappingpoint.delete(0,tk.END)
        self.entry_max_tappingPoint.delete(0,tk.END)
        self.entry_aliasId.delete(0,tk.END)
        self.checkbox_packaging_tappingPoint_all.deselect()
        self.checkbox_packaging_tappingPoint_all.configure(state="normal")
        self.checkbox_packaging_tappingPoint_zak.deselect()
        self.checkbox_packaging_tappingPoint_zak.configure(state="normal")
        self.checkbox_packaging_tappingPoint_curah.deselect()
        self.checkbox_packaging_tappingPoint_curah.configure(state="normal")
        self.checkbox_packaging_tappingPoint_jumbo.deselect()
        self.checkbox_packaging_tappingPoint_jumbo.configure(state="normal")
        self.checkbox_packaging_tappingPoint_pallet.deselect()
        self.checkbox_packaging_tappingPoint_pallet.configure(state="normal")

        file.commit()
        file.close()

    def remove_warehouse_flow(self):
        # remove to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("DELETE from tb_warehouseFlow where oid = ?",(no_warehouseFlow,))
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in self.warehouse_flow_table.get_children():
            self.warehouse_flow_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_warehouseFlow")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.warehouse_flow_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("evenrow",))
            else:
                self.warehouse_flow_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("oddrow",))
            count+=1
        # reset input
        self.checkbox_tappingPoint_warehouseflow_registration.deselect()
        self.checkbox_tappingPoint_warehouseflow_weightbridge1.deselect()
        self.checkbox_tappingPoint_warehouseflow_warehouse.deselect()
        self.checkbox_tappingPoint_warehouseflow_weightbridge2.deselect()
        self.checkbox_tappingPoint_warehouseflow_covering.deselect()
        self.checkbox_tappingPoint_warehouseflow_unregistration.deselect()

        file.commit()
        file.close()

    def remove_user_registration(self):
        # remove to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("DELETE from tb_userRegistration where oid = ?",(no_user,))
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in  self.userRegistration_table.get_children():
             self.userRegistration_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_userRegistration")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.userRegistration_table.insert(parent="",index="end", iid=count, text="", values=(i[19], i[0], i[2], i[3], i[1], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18]), tags=("evenrow",))
            else:
                self.userRegistration_table.insert(parent="",index="end", iid=count, text="", values=(i[19], i[0], i[2], i[3], i[1], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18]), tags=("oddrow",))
            count+=1
        # reset input
        self.entry_regis_username.delete(0,tk.END)
        self.entry_regis_password.delete(0,tk.END)
        self.entry_regis_retypePassword.delete(0,tk.END)
        self.entry_regis_email.delete(0,tk.END)
        self.entry_regis_remark.delete(0,tk.END)

        self.checkbox_alertNoteMenu.deselect()
        self.checkbox_reportViewerMenu.deselect()
        self.checkbox_packagingMenu.deselect()
        self.checkbox_productMenu.deselect()
        self.checkbox_tappingPointMenu.deselect()
        self.checkbox_warehouseFlowMenu.deselect()
        self.checkbox_rfidManualCallMenu.deselect()
        self.checkbox_vettingMenu.deselect()
        self.checkbox_unloadingMenu.deselect()
        self.checkbox_distributionMenu.deselect()
        self.checkbox_databaseSettingMenu.deselect()
        self.checkbox_displanApiMenu.deselect()
        self.checkbox_otherSettingMenu.deselect()
        self.checkbox_userProfileMenu.deselect()
        self.checkbox_userRegistrationMenu.deselect()

        file.commit()
        file.close()

    def replace_vetting(self):
        # update to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # get data from input
        company = self.entry_company.get()
        no_polisi = self.entry_policeNo.get()
        reg_date = self.regdate_calendar.get_date()
        expired_date = self.expireddate_calendar.get_date()
        remark = self.entry_remark.get()
        status = self.optionmenu_status.get()
        rfid1 = self.entry_rfid1.get()
        rfid2 = self.entry_rfid2.get()
        category = self.optionmenu_category.get()
        # write -> update data to database
        c.execute("UPDATE tb_vetting SET company=?, no_polisi=?, reg_date=?, expired_date=?, remark=?, status=?, rfid1=?, rfid2=?, category=? WHERE oid = ?",
            company,
            no_polisi,
            reg_date,
            expired_date,
            remark,
            status,
            rfid1,
            rfid2,
            category,
            no
        )
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in self.vetting_table.get_children():
            self.vetting_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_vetting")
        rec = c.fetchall()
        urutan = 1
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.vetting_table.insert(parent="",index="end", iid=count, text="", values=(i[9], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
            else:
                self.vetting_table.insert(parent="",index="end", iid=count, text="", values=(i[9], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
            count+=1
            urutan+=1
        # reset input
        self.entry_company.delete(0,tk.END)
        self.entry_policeNo.delete(0,tk.END)
        self.entry_remark.delete(0,tk.END)
        self.entry_rfid1.delete(0,tk.END)
        self.entry_rfid2.delete(0,tk.END)

        file.commit()
        file.close()

    def replace_packaging(self):
        confirmation = messagebox.askyesno("Confirmation", "Are you sure you want to replace the data?")

        if confirmation:
            # input to database
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # get data from input
            packaging = self.entry_packaging.get()
            prefix = self.entry_prefix.get()
            remark = self.entry_remark_packaging.get()
            status = self.optionmenu_status_packaging.get()
            sla = self.entry_sla_packaging.get()
            # write data to database
            c.execute("UPDATE tb_packagingReg SET packaging=?, prefix=?, remark=?, status=?, sla=? WHERE packaging = ?",
                packaging,
                prefix,
                remark,
                status,
                sla,
                packaging_reg
            )
            file.commit()
            file.close()
            # hapus item di tabel GUI
            for record in self.packaging_table.get_children():
                self.packaging_table.delete(record)
            # get data from database
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            count = 0
            c.execute("SELECT * FROM tb_packagingReg")
            rec = c.fetchall()
            # input to tabel GUI
            for i in rec:
                if count % 2 == 0:
                    self.packaging_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4]), tags=("evenrow",))
                else:
                    self.packaging_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4]), tags=("oddrow",))
                count+=1
            # reset input
            self.entry_packaging.delete(0,tk.END)
            self.entry_prefix.delete(0,tk.END)
            self.entry_remark_packaging.delete(0,tk.END)
            self.entry_sla_packaging.delete(0,tk.END)

            file.commit()
            file.close()

            # save log
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            date_log = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
            time_log = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
            transaction_log = "AUDIT LOG"
            status = "DONE"
            remark = f"Replace Item in packaging registration: {packaging}"
            # write data to database
            c.execute("INSERT INTO tb_log_report (warehouse,date_log,time_log,transaction_log,operator,status,remark) VALUES (?,?,?,?,?,?,?)",
                wh_name,
                date_log,
                time_log,
                transaction_log,
                username_log,
                status,
                remark
            )
            file.commit()
            file.close()

    def replace_warehouse_slaSetting(self):
        confirmation = messagebox.askyesno("Confirmation", "Are you sure you want to delete the data?")

        if confirmation:
            # input to database
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # get data from input
            packaging = self.optionmenu_packaging_sla.get()
            product_grp = self.optionmenu_product_sla.get()
            weight_min = self.entry_weightMin.get()
            weight_max = self.entry_weightMax.get()
            wh_sla = self.entry_warehouseMin.get()
            vehicle = self.optionmenu_vehicle_sla.get()
            waiting_time = self.entry_waitingTime.get()
            w_bridge1 = self.entry_wBridge1Min.get()
            w_bridge2 = self.entry_wBridge2Min.get()
            covering_time = self.entry_coveringMin.get()
            # write data to database
            c.execute("UPDATE tb_warehouseSla SET packaging=?, product_grp=?, weight_min=?, weight_max=?, wh_sla=?, vehicle=?, waiting_time=?, w_bridge1=?, w_bridge2=?, covering_time=? WHERE oid = ?",
                packaging,
                product_grp,
                weight_min,
                weight_max,
                wh_sla,
                vehicle,
                waiting_time,
                w_bridge1,
                w_bridge2,
                covering_time,
                no_wh_sla
            )
            file.commit()
            file.close()
            # hapus item di tabel GUI
            for record in self.warehouse_slaSetting_table.get_children():
                self.warehouse_slaSetting_table.delete(record)
            # get data from database
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            count = 0
            c.execute("SELECT * FROM tb_warehouseSla WHERE packaging = ?",(packaging,))
            rec = c.fetchall()
            # input to tabel GUI
            for i in rec:
                if count % 2 == 0:
                    self.warehouse_slaSetting_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10]), tags=("evenrow",))
                else:
                    self.warehouse_slaSetting_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10]), tags=("oddrow",))
                count+=1
            # reset input
            self.entry_weightMin.delete(0,tk.END)
            self.entry_weightMax.delete(0,tk.END)
            self.entry_warehouseMin.delete(0,tk.END)
            self.entry_waitingTime.delete(0,tk.END)
            self.entry_wBridge1Min.delete(0,tk.END)
            self.entry_wBridge2Min.delete(0,tk.END)
            self.entry_coveringMin.delete(0,tk.END)

            file.commit()
            file.close()

            # save log
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            date_log = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
            time_log = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
            transaction_log = "AUDIT LOG"
            status = "DONE"
            remark = f"Replace data warehouse SLA setting: {packaging}, {product_grp}"
            # write data to database
            c.execute("INSERT INTO tb_log_report (warehouse,date_log,time_log,transaction_log,operator,status,remark) VALUES (?,?,?,?,?,?,?)",
                wh_name,
                date_log,
                time_log,
                transaction_log,
                username_log,
                status,
                remark
            )
            file.commit()
            file.close()

    def replace_distribution(self):
        # update to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # get data from input
        company = self.entry_company_distribution.get()
        no_polisi = self.entry_policeNo_distribution.get()
        packaging = self.optionmenu_packaging_distribution.get()
        product = self.optionmenu_product_distribution.get()
        date = self.date_distribution_calendar.get_date()
        time_hour = self.spbox_hour_distribution.get()
        time_minute = self.spbox_minute_distribution.get()
        time = f"{time_hour}:{time_minute}"
        remark = self.entry_remark_distribution.get()
        status = self.optionmenu_distribution_status.get()
        quota = self.entry_quota_distribution.get()
        no_so = self.entry_noSO_distribution.get()
        satuan = "KG"
        berat = self.entry_berat_distribution.get()
        wh_id = self.optionmenu_warehouse_distribution.get()
        # write -> update data to database
        if wh_id != "Otomatis":
            c.execute("UPDATE tb_loadingList SET company=?, no_polisi=?, packaging=?, product=?, date=?, time=?, remark=?, status=?, quota=?, no_so=?, satuan=?, berat=?, wh_id=? WHERE oid = ?",
                company,
                no_polisi,
                packaging,
                product,
                date,
                time,
                remark,
                status,
                quota,
                no_so,
                satuan,
                berat,
                "0"+wh_id[7],
                no_loadingList
            )
        else:
            c.execute("UPDATE tb_loadingList SET company=?, no_polisi=?, packaging=?, product=?, date=?, time=?, remark=?, status=?, quota=?, no_so=?, satuan=?, berat=?, wh_id=? WHERE oid = ?",
                company,
                no_polisi,
                packaging,
                product,
                date,
                time,
                remark,
                status,
                quota,
                no_so,
                satuan,
                berat,
                "Auto",
                no_loadingList
            )
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in self.loadingList_table.get_children():
            self.loadingList_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_loadingList")
        rec = c.fetchall()
        # input to tabel GUI
        urutan = 1
        for i in rec:
            if count % 2 == 0:
                self.loadingList_table.insert(parent="",index="end", iid=count, text="", values=(i[13], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12]), tags=("evenrow",))
            else:
                self.loadingList_table.insert(parent="",index="end", iid=count, text="", values=(i[13], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12]), tags=("oddrow",))
            count+=1
            urutan+=1
        # reset input
        self.entry_company_distribution.delete(0,tk.END)
        self.entry_policeNo_distribution.delete(0,tk.END)
        self.optionmenu_packaging_distribution.set("Pilih Packaging")
        self.optionmenu_product_distribution.set("Pilih Product")
        self.entry_quota_distribution.delete(0,tk.END)
        self.entry_quota_distribution.insert(0,"0")
        self.entry_remark_distribution.delete(0,tk.END)
        self.entry_remark_distribution.insert(0,"LOADING")
        self.entry_berat_distribution.delete(0,tk.END)
        self.entry_noSO_distribution.delete(0,tk.END)

        file.commit()
        file.close()

    def replace_product_registration(self):
        # input to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # get data from input
        packaging = self.optionmenu_packaging_product.get()
        product_grp = self.entry_product_group.get()
        remark =  self.entry_product_group_remark.get()
        status = self.optionmenu_product_group_status.get()
        # write data to database
        c.execute("UPDATE tb_productReg SET packaging=?, product_grp=?, remark=?, status=? WHERE oid = ?",
            packaging,
            product_grp,
            remark,
            status,
            no_product
        )
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in self.product_table.get_children():
            self.product_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_productReg WHERE packaging = ?",(packaging,))
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.product_table.insert(parent="",index="end", iid=count, text="", values=(i[1], i[2], i[3], i[4]), tags=("evenrow",))
            else:
                self.product_table.insert(parent="",index="end", iid=count, text="", values=(i[1], i[2], i[3], i[4]), tags=("oddrow",))
            count+=1
        # reset input
        self.entry_product_group.delete(0,tk.END)
        self.entry_product_group_remark.delete(0,tk.END)

        file.commit()
        file.close()

    def replace_tapping_point(self):
        # input to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # get data from input
        hw_id = self.entry_machineId.get()
        tap_id = self.entry_tapId.get()
        remark = self.entry_remark_tappingpoint.get()
        packaging = self.checkBox_get_tappingPoint()
        product = self.optionmenu_product_tappingPoint.get()
        activity = self.optionmenu_activity_tappingPoint.get()
        block = self.entry_tapBlocking_tappingpoint.get()
        max_value = self.entry_max_tappingPoint.get()
        function = self.optionmenu_function_tappingPoint.get()
        status = self.optionmenu_status_tappingPoint.get()
        alias = self.entry_aliasId.get()
        # write data to database
        c.execute("UPDATE tb_tappingPointReg SET hw_id=?, tap_id=?, remark=?, packaging=?, product=?, activity=?, block=?, max=?, functions=?, status=?, alias=? WHERE oid = ?",
            hw_id,
            tap_id,
            remark,
            packaging,
            product,
            activity,
            block,
            max_value,
            function,
            status,
            alias,
            no_tappingPoint
        )
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in self.tappingPoint_table.get_children():
            self.tappingPoint_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_tappingPointReg")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.tappingPoint_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("evenrow",))
            else:
                self.tappingPoint_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("oddrow",))
            count+=1
        # reset input
        self.entry_machineId.delete(0,tk.END)
        self.entry_tapId.delete(0,tk.END)
        self.entry_remark_tappingpoint.delete(0,tk.END)
        self.entry_tapBlocking_tappingpoint.delete(0,tk.END)
        self.entry_max_tappingPoint.delete(0,tk.END)
        self.entry_aliasId.delete(0,tk.END)
        self.checkbox_packaging_tappingPoint_all.deselect()
        self.checkbox_packaging_tappingPoint_all.configure(state="normal")
        self.checkbox_packaging_tappingPoint_zak.deselect()
        self.checkbox_packaging_tappingPoint_zak.configure(state="normal")
        self.checkbox_packaging_tappingPoint_curah.deselect()
        self.checkbox_packaging_tappingPoint_curah.configure(state="normal")
        self.checkbox_packaging_tappingPoint_jumbo.deselect()
        self.checkbox_packaging_tappingPoint_jumbo.configure(state="normal")
        self.checkbox_packaging_tappingPoint_pallet.deselect()
        self.checkbox_packaging_tappingPoint_pallet.configure(state="normal")

        file.commit()
        file.close()

    def replace_warehouse_flow(self):
        # input to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # get data from input
        packaging = self.optionmenu_packaging_warehouseflow.get()
        activity = self.optionmenu_activity_warehouseflow.get()
        status = self.optionmenu_status_warehouseflow.get()
        registration = self.checkbox_tappingPoint_warehouseflow_registration.get()
        if registration == 0:
            registration = "NO"
        weightbridge1 = self.checkbox_tappingPoint_warehouseflow_weightbridge1.get()
        if weightbridge1 == 0:
            weightbridge1 = "NO"
        warehouse = self.checkbox_tappingPoint_warehouseflow_warehouse.get()
        if warehouse == 0:
            warehouse = "NO"
        weightbridge2 = self.checkbox_tappingPoint_warehouseflow_weightbridge2.get()
        if weightbridge2 == 0:
            weightbridge2 = "NO"
        covering = self.checkbox_tappingPoint_warehouseflow_covering.get()
        if covering == 0:
            covering = "NO"
        unregistration = self.checkbox_tappingPoint_warehouseflow_unregistration.get()
        if unregistration == 0:
            unregistration = "NO"
        # write data to database
        c.execute("UPDATE tb_warehouseFlow SET packaging=?, activity=?, registration=?, weightbridge1=?, warehouse=?, weightbridge2=?, covering=?, unregistration=?, status=? WHERE oid = ?",
            packaging,
            activity,
            registration,
            weightbridge1,
            warehouse,
            weightbridge2,
            covering,
            unregistration,
            status,
            no_warehouseFlow
        )
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in self.warehouse_flow_table.get_children():
            self.warehouse_flow_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_warehouseFlow")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.warehouse_flow_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("evenrow",))
            else:
                self.warehouse_flow_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("oddrow",))
            count+=1
        # reset input
        self.checkbox_tappingPoint_warehouseflow_registration.deselect()
        self.checkbox_tappingPoint_warehouseflow_weightbridge1.deselect()
        self.checkbox_tappingPoint_warehouseflow_warehouse.deselect()
        self.checkbox_tappingPoint_warehouseflow_weightbridge2.deselect()
        self.checkbox_tappingPoint_warehouseflow_covering.deselect()
        self.checkbox_tappingPoint_warehouseflow_unregistration.deselect()

        file.commit()
        file.close()

    def replace_unloading(self):
        # update to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # get data from input
        company = self.entry_company_unloading.get()
        no_polisi = self.entry_policeNo_unloading.get()
        packaging = self.optionmenu_packaging.get()
        product = self.optionmenu_product.get()
        date = self.date_unloading_calendar.get_date()
        time_hour = self.spbox_hour.get()
        time_minute = self.spbox_minute.get()
        time = f"{time_hour}:{time_minute}"
        remark = self.entry_remark_unloading.get()
        status = self.optionmenu_unloading_status.get()
        quota = self.entry_quota_unloading.get()
        berat = self.entry_berat_unloading.get()
        wh_id = self.optionmenu_warehouse.get()
        no_so = self.optionmenu_noSo.get()
        # write -> update data to database
        if wh_id != "Otomatis":
            c.execute("UPDATE tb_unloading SET company=?, no_polisi=?, packaging=?, product=?, date=?, time=?, remark=?, status=?, quota=?, berat=?, wh_id=?, no_so=? WHERE oid = ?",
                company,
                no_polisi,
                packaging,
                product,
                date,
                time,
                remark,
                status,
                quota,
                berat,
                "0"+wh_id[7],
                no_so,
                no_unloading
            )
        else:
            c.execute("UPDATE tb_unloading SET company=?, no_polisi=?, packaging=?, product=?, date=?, time=?, remark=?, status=?, quota=?, berat=?, wh_id=?, no_so=? WHERE oid = ?",
                company,
                no_polisi,
                packaging,
                product,
                date,
                time,
                remark,
                status,
                quota,
                berat,
                "Auto",
                no_so,
                no_unloading
            )
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in self.unloading_table.get_children():
            self.unloading_table.delete(record)
        # get data from database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        count = 0
        c.execute("SELECT * FROM tb_unloading")
        rec = c.fetchall()
        # input to tabel GUI
        urutan = 1
        for i in rec:
            if count % 2 == 0:
                self.unloading_table.insert(parent="",index="end", iid=count, text="", values=(i[12], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("evenrow",))
            else:
                self.unloading_table.insert(parent="",index="end", iid=count, text="", values=(i[12], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("oddrow",))
            count+=1
            urutan+=1

        file.commit()
        file.close()

        # reset input
        self.entry_company_unloading.delete(0,tk.END)
        self.entry_policeNo_unloading.delete(0,tk.END)
        self.optionmenu_packaging.set("Pilih Packaging")
        self.optionmenu_product.set("Pilih Product")
        self.entry_quota_unloading.delete(0,tk.END)
        self.entry_quota_unloading.insert(0,"0")
        self.entry_remark_unloading.delete(0,tk.END)
        self.entry_berat_unloading.delete(0,tk.END)
        self.entry_remark_unloading.insert(0,"UNLOADING")
        self.optionmenu_noSo.delete(0,tk.END)

    def replace_user_registration(self):
        username_app = self.entry_regis_username.get()
        password_app = self.entry_regis_password.get()
        retypePassword = self.entry_regis_retypePassword.get()
        email = self.entry_regis_email.get()
        remark = self.entry_regis_remark.get()

        alertNoteMenu = self.checkbox_alertNoteMenu.get()
        reportViewerMenu = self.checkbox_reportViewerMenu.get()
        packagingMenu = self.checkbox_packagingMenu.get()
        productMenu = self.checkbox_productMenu.get()
        tappingPointMenu = self.checkbox_tappingPointMenu.get()
        warehouseFlowMenu = self.checkbox_warehouseFlowMenu.get()
        rfidManualCallMenu = self.checkbox_rfidManualCallMenu.get()
        vettingMenu = self.checkbox_vettingMenu.get()
        unloadingMenu = self.checkbox_unloadingMenu.get()
        distributionMenu = self.checkbox_distributionMenu.get()
        databaseSettingMenu = self.checkbox_databaseSettingMenu.get()
        displanApiMenu = self.checkbox_displanApiMenu.get()
        otherSettingMenu = self.checkbox_otherSettingMenu.get()
        userProfileMenu = self.checkbox_userProfileMenu.get()
        userRegistrationMenu = self.checkbox_userRegistrationMenu.get()

        if password_app == retypePassword:
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()

            c.execute("UPDATE tb_userRegistration SET username=?, password=?, email=?, remark=?, alertNote=?, reportViewer=?, packaging=?, productReg=?, tappingPoint=?, warehouseFlow=?, rfidManualCall=?, vetting=?, unloading=?, distribution=?, databaseSetting=?, displanApi=?, otherSetting=?, userProfile=?, userRegistration=? WHERE oid=?",
                username_app,
                password_app,
                email,
                remark,
                alertNoteMenu,
                reportViewerMenu,
                packagingMenu,
                productMenu,
                tappingPointMenu,
                warehouseFlowMenu,
                rfidManualCallMenu,
                vettingMenu,
                unloadingMenu,
                distributionMenu,
                databaseSettingMenu,
                displanApiMenu,
                otherSettingMenu,
                userProfileMenu,
                userRegistrationMenu,
                no_user
            )
            file.commit()
            file.close()

            self.entry_regis_username.delete(0,tk.END)
            self.entry_regis_password.delete(0,tk.END)
            self.entry_regis_retypePassword.delete(0,tk.END)
            self.entry_regis_email.delete(0,tk.END)
            self.entry_regis_remark.delete(0,tk.END)

            self.checkbox_alertNoteMenu.deselect()
            self.checkbox_reportViewerMenu.deselect()
            self.checkbox_packagingMenu.deselect()
            self.checkbox_productMenu.deselect()
            self.checkbox_tappingPointMenu.deselect()
            self.checkbox_warehouseFlowMenu.deselect()
            self.checkbox_rfidManualCallMenu.deselect()
            self.checkbox_vettingMenu.deselect()
            self.checkbox_unloadingMenu.deselect()
            self.checkbox_distributionMenu.deselect()
            self.checkbox_databaseSettingMenu.deselect()
            self.checkbox_displanApiMenu.deselect()
            self.checkbox_otherSettingMenu.deselect()
            self.checkbox_userProfileMenu.deselect()
            self.checkbox_userRegistrationMenu.deselect()

            # hapus item di tabel GUI
            for record in self.userRegistration_table.get_children():
                self.userRegistration_table.delete(record)
            # get data from database
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            count = 0
            c.execute("SELECT * FROM tb_userRegistration")
            rec = c.fetchall()
            # input to tabel GUI
            for i in rec:
                if count % 2 == 0:
                    self.userRegistration_table.insert(parent="",index="end", iid=count, text="", values=(i[19], i[0], i[2], i[3], i[1], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18]), tags=("evenrow",))
                else:
                    self.userRegistration_table.insert(parent="",index="end", iid=count, text="", values=(i[19], i[0], i[2], i[3], i[1], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18]), tags=("oddrow",))
                count+=1
            file.commit()
            file.close()

            messagebox.showinfo("Sukses", "Data telah diperbarui")
        else:
            messagebox.showerror("Error", "Retype password is incorrect.")

    def clear_unloading(self):
        # remove to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("DELETE from tb_unloading")
        # hapus item di tabel GUI
        for record in self.unloading_table.get_children():
            self.unloading_table.delete(record)

        file.commit()
        file.close()

        # reset input
        self.entry_company_unloading.delete(0,tk.END)
        self.entry_policeNo_unloading.delete(0,tk.END)
        self.optionmenu_packaging.set("Pilih Packaging")
        self.optionmenu_product.set("Pilih Product")
        self.entry_quota_unloading.delete(0,tk.END)
        self.entry_quota_unloading.insert(0,"0")
        self.entry_remark_unloading.delete(0,tk.END)
        self.optionmenu_warehouse.set("Otomatis")
        self.entry_remark_unloading.insert(0,"UNLOADING")
        self.optionmenu_noSo.delete(0,tk.END)

    def clear_distribution(self):
        # remove to database
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("DELETE from tb_loadingList")
        # hapus item di tabel GUI
        for record in self.loadingList_table.get_children():
            self.loadingList_table.delete(record)

        file.commit()
        file.close()

        # reset input
        self.entry_company_distribution.delete(0,tk.END)
        self.entry_policeNo_distribution.delete(0,tk.END)
        self.optionmenu_packaging_distribution.set("Pilih Packaging")
        self.optionmenu_product_distribution.set("Pilih Product")
        self.entry_quota_distribution.delete(0,tk.END)
        self.entry_quota_distribution.insert(0,"0")
        self.entry_remark_distribution.delete(0,tk.END)
        self.entry_remark_distribution.insert(0,"LOADING")
        self.entry_berat_distribution.delete(0,tk.END)
        self.entry_noSO_distribution.delete(0,tk.END)

    def tabview_distribution_button(self):
        tab_name = self.tabview_in_distribution.get()
        if tab_name == "Displan Data":
            self.search_button_distribution.configure(command=self.search_distribution_displanData)
            self.add_button_distribution.configure(state="normal",fg_color="dark blue",command=self.add_distribution_displanData)
            self.remove_button_distribution.configure(state="disabled",fg_color="gray")
            self.replace_button_distribution.configure(state="disabled",fg_color="gray")
            self.clear_button_distribution.configure(state="disabled",fg_color="gray")
            self.upload_excel_button_distribution.configure(state="disabled",fg_color="gray")
            self.entry_quota_distribution.delete(0,tk.END)
            self.entry_quota_distribution.configure(state="disabled",fg_color="gray")
            self.optionmenu_distribution_status.configure(state="disabled",fg_color="gray")
            self.optionmenu_warehouse_distribution.configure(state="disabled",fg_color="gray")
            self.entry_remark_distribution.configure(state="normal")
            self.entry_remark_distribution.delete(0,tk.END)
            self.entry_pkgRemark_distribution.configure(state="normal",fg_color="white")
            self.entry_trPart_distribution.configure(state="normal",fg_color="white")

            self.entry_company_distribution.delete(0,tk.END)
            self.entry_policeNo_distribution.delete(0,tk.END)
            self.spbox_hour_distribution.delete(0,tk.END)
            self.spbox_minute_distribution.delete(0,tk.END)
            self.entry_berat_distribution.delete(0,tk.END)
            self.entry_noSO_distribution.delete(0,tk.END)
            self.optionmenu_packaging_distribution.set("Pilih Packaging")
            self.optionmenu_product_distribution.set("Pilih Product")
        else:
            self.search_button_distribution.configure(command=self.search_distribution_loadingList)
            self.add_button_distribution.configure(state="normal",fg_color="dark blue",command=self.add_distribution_loadingList)
            self.remove_button_distribution.configure(state="normal",fg_color="dark blue")
            self.replace_button_distribution.configure(state="normal",fg_color="dark blue")
            self.clear_button_distribution.configure(state="normal",fg_color="dark blue")
            self.upload_excel_button_distribution.configure(state="normal",fg_color="dark green")
            self.entry_quota_distribution.configure(state="normal",fg_color="white")
            self.optionmenu_distribution_status.configure(state="normal",fg_color="white")
            self.optionmenu_warehouse_distribution.configure(state="normal",fg_color="white")
            self.entry_remark_distribution.delete(0,tk.END)
            self.entry_remark_distribution.insert(0,"LOADING")
            self.entry_remark_distribution.configure(state="disabled")
            self.entry_pkgRemark_distribution.delete(0,tk.END)
            self.entry_pkgRemark_distribution.configure(state="disabled",fg_color="gray")
            self.entry_trPart_distribution.delete(0,tk.END)
            self.entry_trPart_distribution.configure(state="disabled",fg_color="gray")

            self.entry_company_distribution.delete(0,tk.END)
            self.entry_policeNo_distribution.delete(0,tk.END)
            self.spbox_hour_distribution.delete(0,tk.END)
            self.spbox_minute_distribution.delete(0,tk.END)
            self.entry_berat_distribution.delete(0,tk.END)
            self.entry_noSO_distribution.delete(0,tk.END)
            self.optionmenu_packaging_distribution.set("Pilih Packaging")
            self.optionmenu_product_distribution.set("Pilih Product")

    # untuk mengatur isi option menu packaging sesuai master
    def optionMenu_packaging(self):
        # set tupple for packaging name to apply to optionmenu
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT packaging FROM tb_packagingReg")
        rec = c.fetchall()
        pckglist = []
        pckglist_report = ["ALL"]
        for i in rec:
            pckglist.append(i[0])
            pckglist_report.append(i[0])

        # set package optionmenu
        self.optionmenu_packaging.configure(values=pckglist)
        self.optionmenu_packaging_distribution.configure(values=pckglist)
        self.optionmenu_packaging_sla.configure(values=pckglist)
        self.optionmenu_packaging_product.configure(values=pckglist)
        self.optionmenu_packaging_warehouseflow.configure(values=pckglist)

        self.report_viewer_total_time_optionmenu_packaging.configure(values=pckglist_report)
        self.report_viewer_transaction_time_optionmenu_packaging.configure(values=pckglist_report)
        self.report_viewer_peak_hour_optionmenu_packaging.configure(values=pckglist_report)
        self.report_viewer_total_queuing_optionmenu_packaging.configure(values=pckglist_report)
        self.report_viewer_ticket_history_optionmenu_packaging.configure(values=pckglist_report)
        self.report_viewer_sla_optionmenu_packaging.configure(values=pckglist_report)

        file.commit()
        file.close()

    # untuk mengatur isi option menu warehouse sesuai master (tabble tapping point yang col function)
    def optionMenu_warehouseAssign(self):
        # set tupple for warehouse name to apply to optionmenu
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT remark FROM tb_tappingPointReg WHERE functions = ?",("Warehouse",))
        rec = c.fetchall()
        warehouselist = ["Otomatis"]
        for i in rec:
            warehouselist.append(i[0])
        
        # set package optionmenu
        self.optionmenu_warehouse.configure(values=warehouselist)
        self.optionmenu_warehouse_distribution.configure(values=warehouselist)
        self.optionmenu_warehouse_rfidPairing.configure(values=warehouselist[1:])
            
        file.commit()
        file.close()

    def checkBox_All_tappingPoint(self):
        global check_all, productTapPoint
        if check_all == False:
            check_all = True
            self.checkbox_packaging_tappingPoint_zak.select()
            self.checkbox_packaging_tappingPoint_zak.configure(state="disabled")
            self.checkbox_packaging_tappingPoint_curah.select()
            self.checkbox_packaging_tappingPoint_curah.configure(state="disabled")
            self.checkbox_packaging_tappingPoint_jumbo.select()
            self.checkbox_packaging_tappingPoint_jumbo.configure(state="disabled")
            self.checkbox_packaging_tappingPoint_pallet.select()
            self.checkbox_packaging_tappingPoint_pallet.configure(state="disabled")

            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # set product optionmenu
            c.execute("SELECT product_grp FROM tb_productReg")
            rec = c.fetchall()
            for i in rec:
                productTapPoint.append(i[0])
            productTapPoint = list(OrderedDict.fromkeys(productTapPoint))
            self.optionmenu_product_tappingPoint.configure(values=productTapPoint)

            file.commit()
            file.close()
        else:
            check_all = False
            self.checkbox_packaging_tappingPoint_zak.deselect()
            self.checkbox_packaging_tappingPoint_zak.configure(state="normal")
            self.checkbox_packaging_tappingPoint_curah.deselect()
            self.checkbox_packaging_tappingPoint_curah.configure(state="normal")
            self.checkbox_packaging_tappingPoint_jumbo.deselect()
            self.checkbox_packaging_tappingPoint_jumbo.configure(state="normal")
            self.checkbox_packaging_tappingPoint_pallet.deselect()
            self.checkbox_packaging_tappingPoint_pallet.configure(state="normal")

            productTapPoint = ["-"]
            self.optionmenu_product_tappingPoint.configure(values=productTapPoint)

    def checkBox_ZAK_tappingPoint(self):
        global check_zak, productTapPoint
        if check_zak == False:
            check_zak = True

            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # set product optionmenu
            c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",("ZAK",))
            rec = c.fetchall()
            for i in rec:
                productTapPoint.append(i[0])
            productTapPoint = list(OrderedDict.fromkeys(productTapPoint))
            self.optionmenu_product_tappingPoint.configure(values=productTapPoint)

            file.commit()
            file.close()
        else:
            check_zak = False

            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # set product optionmenu
            c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",("ZAK",))
            rec = c.fetchall()
            for i in rec:
                for j in productTapPoint:
                    if i[0] == j:
                        productTapPoint.remove(i[0])
            self.optionmenu_product_tappingPoint.configure(values=productTapPoint)

            file.commit()
            file.close()

    def checkBox_CURAH_tappingPoint(self):
        global check_curah, productTapPoint
        if check_curah == False:
            check_curah = True

            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # set product optionmenu
            c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",("CURAH",))
            rec = c.fetchall()
            for i in rec:
                productTapPoint.append(i[0])
            productTapPoint = list(OrderedDict.fromkeys(productTapPoint))
            self.optionmenu_product_tappingPoint.configure(values=productTapPoint)

            file.commit()
            file.close()
        else:
            check_curah = False

            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # set product optionmenu
            c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",("CURAH",))
            rec = c.fetchall()
            for i in rec:
                for j in productTapPoint:
                    if i[0] == j:
                        productTapPoint.remove(i[0])
            self.optionmenu_product_tappingPoint.configure(values=productTapPoint)

            file.commit()
            file.close()

    def checkBox_JUMBO_tappingPoint(self):
        global check_jumbo, productTapPoint
        if check_jumbo == False:
            check_jumbo = True

            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # set product optionmenu
            c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",("JUMBO",))
            rec = c.fetchall()
            for i in rec:
                productTapPoint.append(i[0])
            productTapPoint = list(OrderedDict.fromkeys(productTapPoint))
            self.optionmenu_product_tappingPoint.configure(values=productTapPoint)

            file.commit()
            file.close()
        else:
            check_jumbo = False

            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # set product optionmenu
            c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",("JUMBO",))
            rec = c.fetchall()
            for i in rec:
                for j in productTapPoint:
                    if i[0] == j:
                        productTapPoint.remove(i[0])
            self.optionmenu_product_tappingPoint.configure(values=productTapPoint)

            file.commit()
            file.close()

    def checkBox_PALLET_tappingPoint(self):
        global check_pallet, productTapPoint
        if check_pallet == False:
            check_pallet = True

            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # set product optionmenu
            c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",("PALLET",))
            rec = c.fetchall()
            for i in rec:
                productTapPoint.append(i[0])
            productTapPoint = list(OrderedDict.fromkeys(productTapPoint))
            self.optionmenu_product_tappingPoint.configure(values=productTapPoint)

            file.commit()
            file.close()
        else:
            check_pallet = False

            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # set product optionmenu
            c.execute("SELECT product_grp FROM tb_productReg WHERE packaging = ?",("PALLET",))
            rec = c.fetchall()
            for i in rec:
                for j in productTapPoint:
                    if i[0] == j:
                        productTapPoint.remove(i[0])
            self.optionmenu_product_tappingPoint.configure(values=productTapPoint)

            file.commit()
            file.close()

    def checkBox_get_tappingPoint(self):
        all = self.checkbox_packaging_tappingPoint_all.get()
        zak = self.checkbox_packaging_tappingPoint_zak.get()
        curah = self.checkbox_packaging_tappingPoint_curah.get()
        jumbo = self.checkbox_packaging_tappingPoint_jumbo.get()
        pallet = self.checkbox_packaging_tappingPoint_pallet.get()

        packaging = [all,zak,curah,jumbo,pallet]

        packaging_tappingPoint = ""

        if all == "ALL":
            packaging_tappingPoint = "ALL"
        else:
            for i in packaging:
                if i != 0:
                    if len(packaging_tappingPoint) == 0:
                        packaging_tappingPoint = packaging_tappingPoint + i
                    else:
                        packaging_tappingPoint = packaging_tappingPoint + "/" + i
        
        return packaging_tappingPoint

    def print_ticket(self):
        try:
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT ticket, activity, packaging FROM tb_waitinglist where no_polisi = ?",(self.entry_truckNo_manualCall.get(),))
            rec = c.fetchone()
            ticket = rec[0]
            activity = rec[1]
            pckging = rec[2]
            file.commit()
            file.close()
        except:
            pass
        if activity == "LOADING":
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT wh_id FROM tb_loadingList where no_polisi = ?",(self.entry_truckNo_manualCall.get(),))
            rec = c.fetchone()
            wrhouse = "Gudang " + rec[0]
            file.commit()
            file.close()
        else:
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT wh_id FROM tb_unloading where no_polisi = ?",(self.entry_truckNo_manualCall.get(),))
            rec = c.fetchone()
            wrhouse = "Gudang " + rec[0]
            file.commit()
            file.close()

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

    def delete_waiting_list(self):
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("DELETE from tb_waitinglist where no_polisi = ?",(self.entry_truckNo_manualCall.get(),))
        file.commit()
        file.close()

        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("DELETE from tb_calling where no_polisi = ?",(self.entry_truckNo_manualCall.get(),))
        file.commit()
        file.close()

        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("DELETE from tb_queuing_flow where no_polisi = ?",(self.entry_truckNo_manualCall.get(),))
        file.commit()
        file.close()

        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("DELETE from tb_overSLA where no_polisi = ?",(self.entry_truckNo_manualCall.get(),))
        file.commit()
        file.close()

        self.entry_truckNo_manualCall.delete(0,tk.END)

        # real all data in all tab and show in each tab
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()

        # data alert inbox
        for record in self.manualCall_table.get_children():
            self.manualCall_table.delete(record)
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_waitinglist")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.manualCall_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("evenrow",))
            else:
                self.manualCall_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("oddrow",))
            count+=1

        file.commit()
        file.close()       

    def refresh_waiting_list(self):
        # real all data in all tab and show in each tab
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()

        # data alert inbox
        for record in self.manualCall_table.get_children():
            self.manualCall_table.delete(record)
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_waitinglist")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.manualCall_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("evenrow",))
            else:
                self.manualCall_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("oddrow",))
            count+=1

        file.commit()
        file.close()

    def refresh_pairing_status(self):
        # real all data in all tab and show in each tab
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()

        # data alert inbox
        for record in self.rfidPairing_table.get_children():
            self.rfidPairing_table.delete(record)
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_queuing_status")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.rfidPairing_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18], i[19], i[20]), tags=("evenrow",))
            else:
                self.rfidPairing_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18], i[19], i[20]), tags=("oddrow",))
            count+=1

        file.commit()
        file.close()

    def call_waiting_list(self):
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
            # cek apakah area warehouse tersedia
            # get ticket dan no plat berdasarkan oid terkecil di tb waitinglist
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT activity FROM tb_waitinglist where no_polisi = ?",(self.entry_truckNo_manualCall.get(),))
            rec = c.fetchone()
            activity = rec[0]
            file.commit()
            file.close()
            if activity == "LOADING":
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                c.execute("SELECT wh_id FROM tb_loadingList where no_polisi = ?",(self.entry_truckNo_manualCall.get(),))
                rec = c.fetchone()
                wh_id = rec[0]
                file.commit()
                file.close()
            else:
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                c.execute("SELECT wh_id FROM tb_unloading where no_polisi = ?",(self.entry_truckNo_manualCall.get(),))
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
            c.execute("SELECT max FROM tb_tappingPointReg where remark = ?",("Gudang "+str(int(wh_id)),))
            rec = c.fetchone()
            max_warehouse = rec[0]
            file.commit()
            file.close()
            # kurangi max dengan jumlah  jumlah warehouse yang ditugaskan di tb_queuing status
            max_value_warehouse = int(max_warehouse) - int(count_warehouse)
            # jika wh tersdia lakukan algoritma mengubah oid
            if max_value_warehouse > 0:
                # get oid where no_polisi selected
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                c.execute("SELECT oid FROM tb_waitinglist where no_polisi = ?",(self.entry_truckNo_manualCall.get(),))
                rec = c.fetchone()
                oid_selected = rec[0]
                file.commit()
                file.close()
                # get max oid
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                c.execute("SELECT MAX(oid) FROM tb_waitinglist")
                rec = c.fetchone()
                oid_max = rec[0]
                file.commit()
                file.close()
                # get min oid
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                c.execute("SELECT MIN(oid) FROM tb_waitinglist")
                rec = c.fetchone()
                oid_min = rec[0]
                file.commit()
                file.close()
                if oid_selected == oid_max:
                    # urutkan dari oid yang paling besar
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT oid, no_polisi FROM tb_waitinglist ORDER BY oid DESC")
                    rec = c.fetchall()
                    file.commit()
                    file.close()
                    # tambah satu semua
                    for i in rec:
                        oid_new = str(int(i[0])+1)
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()
                        c.execute("UPDATE tb_waitinglist SET oid=? WHERE no_polisi = ?",
                            oid_new,
                            i[1]
                        )
                        file.commit()
                        file.close()
                    # jadikan oid yang selected (yang max) menjadi oid min (1)
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("UPDATE tb_waitinglist SET oid=? WHERE no_polisi = ?",
                        oid_min,
                        self.entry_truckNo_manualCall.get()
                    )
                    file.commit()
                    file.close()
                elif oid_selected == oid_min:
                    pass
                elif oid_selected != oid_max:
                    # urutkan dari oid yang paling besar
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT oid, no_polisi FROM tb_waitinglist ORDER BY oid DESC")
                    rec = c.fetchall()
                    file.commit()
                    file.close()
                    # tambah satu semua
                    for i in rec:
                        oid_new = str(int(i[0])+1)
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()
                        c.execute("UPDATE tb_waitinglist SET oid=? WHERE no_polisi = ?",
                            oid_new,
                            i[1]
                        )
                        file.commit()
                        file.close()
                    # jadikan oid yang selected (yang max) menjadi oid min (1)
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("UPDATE tb_waitinglist SET oid=? WHERE no_polisi = ?",
                        oid_min,
                        self.entry_truckNo_manualCall.get()
                    )
                    file.commit()
                    file.close()
                    # oid yang > oid_selected dikurang 1
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT oid, no_polisi FROM tb_waitinglist ORDER BY oid DESC")
                    rec = c.fetchall()
                    file.commit()
                    file.close()
                    for i in rec:
                        if int(i[0]) > int(oid_selected):
                            oid_new = str(int(i[0])-1)
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()
                            c.execute("UPDATE tb_waitinglist SET oid=? WHERE no_polisi = ?",
                                oid_new,
                                i[1]
                            )
                            file.commit()
                            file.close()
            
                # real all data in all tab and show in each tab
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                # data alert inbox
                for record in self.manualCall_table.get_children():
                    self.manualCall_table.delete(record)
                # get data from database
                count = 0
                c.execute("SELECT * FROM tb_waitinglist")
                rec = c.fetchall()
                # input to tabel GUI
                for i in rec:
                    if count % 2 == 0:
                        self.manualCall_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("evenrow",))
                    else:
                        self.manualCall_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("oddrow",))
                    count+=1
                file.commit()
                file.close()

            else:
                messagebox.showerror("showerror", "Area Gudang {} Penuh\nSilahkan tunggu sampai area tersedia untuk panggilan berikutya".format(wh_id))
        else:
            messagebox.showerror("showerror", "Area Gate IN Penuh\nSilahkan tunggu sampai area tersedia untuk panggilan berikutya")

    def skip_waiting_list(self):
        # get oid where no_polisi selected
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT oid FROM tb_waitinglist where no_polisi = ?",(self.entry_truckNo_manualCall.get(),))
        rec = c.fetchone()
        oid_selected = rec[0]
        file.commit()
        file.close()
        # get max oid
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT oid FROM tb_waitinglist where oid = (SELECT MAX(oid) FROM tb_waitinglist)")
        rec = c.fetchone()
        oid_max = rec[0]
        file.commit()
        file.close()
        # ubah oid_selected to max oid + 1
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("UPDATE tb_waitinglist SET oid=? WHERE no_polisi = ?",
            str(int(oid_max)+1),
            self.entry_truckNo_manualCall.get()
        )
        file.commit()
        file.close()
        # ubah oid_selected +1  to oid_selected
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("UPDATE tb_waitinglist SET oid=? WHERE oid = ?",
            oid_selected,
            str(int(oid_selected)+1)
        )
        file.commit()
        file.close()
        # ubah oid yang sudah jadi str(int(oid_max)+1) menjadi str(int(oid_selected)+1)
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("UPDATE tb_waitinglist SET oid=? WHERE oid = ?",
            str(int(oid_selected)+1),
            str(int(oid_max)+1)
        )
        file.commit()
        file.close()

        # real all data in all tab and show in each tab
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # data alert inbox
        for record in self.manualCall_table.get_children():
            self.manualCall_table.delete(record)
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_waitinglist")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.manualCall_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("evenrow",))
            else:
                self.manualCall_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("oddrow",))
            count+=1
        file.commit()
        file.close()

    def gatein_waiting_list(self):
        # get no police
        no_polisi = self.entry_truckNo_manualCall.get()

        # get activity, packaging, product, ticket, arrival
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT activity, packaging, product, ticket, arrival FROM tb_waitinglist WHERE no_polisi = ?", (no_polisi,))
        rec = c.fetchone()
        activity, packaging, product, ticket, arrival = rec
        file.commit()
        file.close()

        #4 get kuota transaksi based on activity -> tabel unloading/loading (activity=remark)
        if activity ==  "LOADING":
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT wh_id FROM tb_loadingList where no_polisi = ? and status = ? and product = ?", (no_polisi, "Activate", product))
            rec = c.fetchone()
            wh_id = rec[0]
            file.commit()
            file.close()
        elif activity == "UNLOADING":
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT wh_id FROM tb_unloading where no_polisi = ? and status = ? and product = ?", (no_polisi, "Activate", product))
            rec = c.fetchone()
            wh_id = rec[0]
            file.commit()
            file.close()

        # cek dlu apakah no_plat sudah ada di tb_queuing_flow
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT no_polisi FROM tb_queuing_flow where no_polisi = ?",(no_polisi,))
        rec = c.fetchone()
        if rec == None: # kalo ga ada
            # Insert into tb_queuing_flow()
            c.execute("INSERT INTO tb_queuing_flow (no_polisi, getin, weight_bridge1, warehouse, cargo_covering, weight_bridge2, getout) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (no_polisi, "False", "False", "False", "False", "False", "True"))
            file.commit()

            # Insert into tb_calling
            c.execute("INSERT INTO tb_calling (no_polisi, call_wb1, call_wh, call_wb2, call_cc, call_getout, call_end) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (no_polisi, "True", "True", "True", "True", "True", "True"))
            file.commit()

            # Insert into tb_overSLA
            c.execute("INSERT INTO tb_overSLA (no_polisi, over_wb1, over_wh, over_wb2, over_cc) VALUES (?, ?, ?, ?, ?)",
                    (no_polisi, "False", "False", "False", "False"))
            file.commit()
            file.close()
        
        #get bluetooth name
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT rfid1 FROM tb_vetting where no_polisi = ?",(no_polisi,))
        rec = c.fetchone()
        blth_name = rec[0]
        file.commit()
        file.close()

        # insert to tb_temporaryRFID
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT getout FROM tb_queuing_flow where no_polisi = ?",(no_polisi,))
        rec = c.fetchone()
        cek_getout = rec[0]
        file.commit()
        file.close()
        time_detected = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
        if cek_getout == "True":
            location = "Pairing/Registration"
            # Insert into tb_temporaryRFID
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("INSERT INTO tb_temporaryRFID (rfid, no_polisi, location, time_detected) VALUES (?, ?, ?, ?)",(blth_name, no_polisi, location, time_detected))
            file.commit()
            file.close()

        #
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT TOP 1 oid, rfid, no_polisi, location, time_detected FROM tb_temporaryRFID where no_polisi = ? ORDER BY oid DESC",(no_polisi,))
        rec = c.fetchone()
        rfid_last = rec[1]
        no_polisi_last = rec[2]
        location_last_temp = rec[3]
        time_detected_last = rec[4]
        file.commit()
        file.close()

        # jika plat di lokasi getin, cek apakah plat no sudah ada di tb queuing status, jika belum maka insert to tabel queuing status dan hapus di waiting list
        #> location gatein
        if location_last_temp == "Pairing/Registration":
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT TOP 1 oid, rfid, no_polisi, location, time_detected FROM tb_temporaryRFID where no_polisi = ? ORDER BY oid DESC",(no_polisi,))
            rec = c.fetchone()
            # check apakah packaging dan activity ini perlu kesini
            c.execute("SELECT registration FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
            rec = c.fetchone()
            flow_regis = rec[0]
            file.commit()
            if flow_regis == "YES":
                # cek apakah getout true
                c.execute("SELECT getout FROM tb_queuing_flow where no_polisi = ?",(no_polisi_last,))
                rec = c.fetchone()
                Qflow_getout = rec[0]
                file.commit()
                if Qflow_getout == "True":
                    # get nilai calling_time
                    c.execute("SELECT calling_time FROM tb_calling_time where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    if rec != None:
                        calling_time = rec[0]
                    else:
                        calling_time = ""
                    file.commit()
                    # cek apakah plat no sudah ada di tb queuing status
                    c.execute("SELECT no_polisi FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    file.commit()
                    # jika belum maka calling dan insert to tabel queuing status
                    if rec == None:
                        # insert to tabel queuing status
                        c.execute("INSERT INTO tb_queuing_status (rfid,ticket,no_polisi,activity,packaging,product,arrival,calling,pairing,wb1_in,wb1_out,wh_in,wh_out,wb2_in,wb2_out,cc_in,cc_out,unpairing,total_minute,wh_id) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                            (rfid_last,
                            ticket,
                            no_polisi_last,
                            activity,
                            packaging,
                            product,
                            arrival,
                            calling_time,
                            time_detected_last,
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            wh_id)
                        )
                        file.commit()
                    
                    # hapus data degn noplat di waiting list
                    c.execute("DELETE FROM tb_waitinglist where no_polisi = ?",(no_polisi_last,))
                    file.commit()
                    # update tb_queuing_flow getin true
                    c.execute("UPDATE tb_queuing_flow SET getin=?, getout=? WHERE no_polisi = ?",
                        ("True",
                        "False",
                        no_polisi_last)
                    )
                    file.commit()
                    # hitung durasi pairing | calling -> input to total_minute di tb_queuing_status (jam:menit:detik)
                    # get pairing & calling timetotal_minute
                    c.execute("SELECT pairing, arrival FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    pairing = rec[0]
                    arrival = rec[1]
                    file.commit()
                    # hitung durasi
                    format = '%H:%M:%S'
                    duration = datetime.strptime(pairing, format) - datetime.strptime(arrival, format)
                    # update and insert duration to tb_queuing_status
                    c.execute("UPDATE tb_queuing_status SET total_minute=? WHERE no_polisi = ?",
                        (str(duration),
                        no_polisi_last)
                    )
                    file.commit()
                    # hapus yang di tb_calling_time
                    c.execute("DELETE from tb_calling_time where no_polisi = ?",(no_polisi_last,))
                    file.commit()
                    # next calling for next place
                    c.execute("SELECT call_wb1 FROM tb_calling where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    call_wb1 = rec[0]
                    file.commit()
                    # insert table start (insert karena ini harus yang pertama)
                    try:
                        c.execute("MERGE INTO tb_temporaryRFID_start AS target USING (VALUES (?, ?, ?, ?)) AS source (rfid, no_polisi, location, time_detected) ON target.no_polisi = source.no_polisi WHEN MATCHED THEN UPDATE SET target.rfid = source.rfid, target.location = source.location, target.time_detected = source.time_detected WHEN NOT MATCHED THEN  INSERT (rfid, no_polisi, location, time_detected) VALUES (source.rfid, source.no_polisi, source.location, source.time_detected);",
                            (rfid_last,
                            no_polisi_last,
                            location_last_temp,
                            time_detected_last)
                        )
                        file.commit()
                    except:
                        pass
                    # calling
                    if call_wb1 == "True":
                        # cek apakah covering atau tidak
                        c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
                        rec = c.fetchone()
                        flow_cc = rec[0]
                        file.commit()
                        # input to tb_playSound
                        c.execute("INSERT INTO tb_playSound (location, activity, ticket, covering_status) VALUES (?,?,?,?)",
                            (location_last_temp,
                            activity,
                            ticket,
                            flow_cc)
                        )
                        file.commit()
                        # update tb_calling
                        c.execute("UPDATE tb_calling SET call_wb1=? WHERE no_polisi = ?",
                            ("False",
                            no_polisi_last)
                        )
                        file.commit()

                # get location start
                c.execute("SELECT location FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                rec = c.fetchone()
                location_start = rec[0]
                file.commit()
                # jika location_last_temp == location start insert and update table last
                if location_last_temp == location_start:
                    try:
                        c.execute("MERGE INTO tb_temporaryRFID_last AS target USING (VALUES (?, ?, ?, ?)) AS source (rfid, no_polisi, location, time_detected) ON target.no_polisi = source.no_polisi WHEN MATCHED THEN  UPDATE SET target.rfid = source.rfid, target.location = source.location, target.time_detected = source.time_detected WHEN NOT MATCHED THEN  INSERT (rfid, no_polisi, location, time_detected) VALUES (source.rfid, source.no_polisi, source.location, source.time_detected);",
                            (rfid_last,
                            no_polisi_last,
                            location_last_temp,
                            time_detected_last)
                        )
                        file.commit()
                    except:
                        pass

    def set_warehouse_pairing_status(self):
        wh_id = self.optionmenu_warehouse_rfidPairing.get()
        wh_id = wh_id[7:]
        if int(wh_id) < 10:
            wh_id = "0"+wh_id
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("UPDATE tb_queuing_status SET wh_id=? WHERE ticket = ?",
            wh_id,
            self.entry_noTiket_rfidPairing.get()
        )
        file.commit()
        file.close()

        # real all data in all tab and show in each tab
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # data alert inbox
        for record in self.rfidPairing_table.get_children():
            self.rfidPairing_table.delete(record)
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_queuing_status")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.rfidPairing_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18], i[19], i[20]), tags=("evenrow",))
            else:
                self.rfidPairing_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18], i[19], i[20]), tags=("oddrow",))
            count+=1
        file.commit()
        file.close()

    def tap_manual(self):
        # get no police
        ticket = self.entry_noTiket_rfidPairing.get()
        location = self.optionmenu_manualTaping.get()
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT no_polisi, activity, packaging, product FROM tb_queuing_status where ticket = ?",(ticket,))
        rec = c.fetchone()
        no_polisi = rec[0]
        activity = rec[1]
        packaging = rec[2]
        product = rec[3]
        file.commit()
        file.close()

        #4 get kuota transaksi based on activity -> tabel unloading/loading (activity=remark)
        if activity ==  "LOADING":
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT berat FROM tb_loadingList where no_polisi = ? and status = ? and product = ?", (no_polisi, "Activate", product))
            rec = c.fetchone()
            berat = rec
            file.commit()
            file.close()
        elif activity == "UNLOADING":
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT berat FROM tb_unloading where no_polisi = ? and status = ? and product = ?", (no_polisi, "Activate", product))
            rec = c.fetchone()
            berat = rec
            file.commit()
            file.close()

        #get bluetooth name
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT rfid1 FROM tb_vetting where no_polisi = ?",(no_polisi,))
        rec = c.fetchone()
        blth_name = rec[0]
        file.commit()
        file.close()

        # insert to tb_temporaryRFID
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT getin, weight_bridge1, warehouse, cargo_covering, weight_bridge2, getout FROM tb_queuing_flow where no_polisi = ?",(no_polisi,))
        rec = c.fetchone()
        cek_getin, cek_weight_bridge1, cek_warehouse, cek_cargo_covering, cek_weight_bridge2, cek_getout = rec
        time_detected = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
        if cek_getin == "True":
            if location == "Weight Bridge 1":
                # Insert into tb_temporaryRFID
                c.execute("INSERT INTO tb_temporaryRFID (rfid, no_polisi, location, time_detected) VALUES (?, ?, ?, ?)",(blth_name, no_polisi, location, time_detected))
                file.commit()
        if cek_weight_bridge1 == "True":
            if location == "Weight Bridge 1" or location == "Warehouse":
                # Insert into tb_temporaryRFID
                c.execute("INSERT INTO tb_temporaryRFID (rfid, no_polisi, location, time_detected) VALUES (?, ?, ?, ?)",(blth_name, no_polisi, location, time_detected))
                file.commit()
        if cek_warehouse == "True":
            if location == "Warehouse" or location == "Weight Bridge 2":
                # Insert into tb_temporaryRFID
                c.execute("INSERT INTO tb_temporaryRFID (rfid, no_polisi, location, time_detected) VALUES (?, ?, ?, ?)",(blth_name, no_polisi, location, time_detected))
                file.commit()
        if cek_weight_bridge2 == "True":
            c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
            rec = c.fetchone()
            cek_cc = rec[0]
            if location == "Weight Bridge 2" or (location == "Cargo Covering" and cek_cc == "YES") or (location == "Unpairing/Unregistration" and cek_cc == "NO"):
                # Insert into tb_temporaryRFID
                c.execute("INSERT INTO tb_temporaryRFID (rfid, no_polisi, location, time_detected) VALUES (?, ?, ?, ?)",(blth_name, no_polisi, location, time_detected))
                file.commit()
        if cek_cargo_covering == "True":
            if location == "Cargo Covering":
                # Insert into tb_temporaryRFID
                c.execute("INSERT INTO tb_temporaryRFID (rfid, no_polisi, location, time_detected) VALUES (?, ?, ?, ?)",(blth_name, no_polisi, location, time_detected))
                file.commit()

        #
        c.execute("SELECT TOP 1 oid, rfid, no_polisi, location, time_detected FROM tb_temporaryRFID where no_polisi = ? ORDER BY oid DESC",(no_polisi,))
        rec = c.fetchone()
        rfid_last = rec[1]
        no_polisi_last = rec[2]
        location_last_temp = rec[3]
        time_detected_last = rec[4]
        file.commit()

        # get time duration if lacation_start != location_last
        #> location wb1
        if location_last_temp == "Weight Bridge 1":
            c.execute("SELECT weightbridge1 FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
            rec = c.fetchone()
            flow_wb1 = rec[0]
            file.commit()
            if flow_wb1 == "YES":
                # jika location = wb 1 check apakah getin sudah true? jika sudah next proses
                c.execute("SELECT getin FROM tb_queuing_flow where no_polisi = ?",(no_polisi_last,))
                rec = c.fetchone()
                Qflow_regis = rec[0]
                file.commit()
                if Qflow_regis == "True":
                    if Qflow_regis == "True":
                        c.execute("SELECT location, time_detected FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        location_start = rec[0]
                        time_wb1_in = rec[1]
                        file.commit()
                        if location_last_temp != location_start:
                            # get time start -> wb1 in, input to tabel queuing status
                            c.execute("UPDATE tb_queuing_status SET wb1_in=? WHERE no_polisi = ?",
                                (time_detected_last,
                                no_polisi_last)
                            )
                            file.commit()
                            # get wb1_in and calling time to calculate timetotal_minute
                            c.execute("SELECT wb1_in, arrival FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                            rec = c.fetchone()
                            wb1_in_time = rec[0]
                            arrival = rec[1]
                            file.commit()
                            # hitung durasi
                            format = '%H:%M:%S'
                            duration = datetime.strptime(wb1_in_time, format) - datetime.strptime(arrival, format)
                            # perbarui total_minute tabel queuing status
                            c.execute("UPDATE tb_queuing_status SET total_minute=? WHERE no_polisi = ?",
                                (str(duration),
                                no_polisi_last)
                            )
                            file.commit()
                            # perbahrui location start
                            try:
                                c.execute("MERGE INTO tb_temporaryRFID_start AS target USING (VALUES (?, ?, ?, ?)) AS source (rfid, no_polisi, location, time_detected) ON target.no_polisi = source.no_polisi WHEN MATCHED THEN  UPDATE SET target.rfid = source.rfid, target.location = source.location, target.time_detected = source.time_detected WHEN NOT MATCHED THEN  INSERT (rfid, no_polisi, location, time_detected) VALUES (source.rfid, source.no_polisi, source.location, source.time_detected);",
                                    (rfid_last,
                                    no_polisi_last,
                                    location_last_temp,
                                    time_detected_last)
                                )
                                file.commit()
                            except:
                                pass

                        # get block in tb_tappingPointReg where function == Weight Bridge 1
                        c.execute("SELECT block FROM tb_tappingPointReg where functions = ?",("Weight Bridge 1",))
                        rec = c.fetchone()
                        block = rec[0]
                        file.commit()
                        # get time start dan last
                        c.execute("SELECT time_detected FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        time_start = rec[0]
                        file.commit()
                        # blocking suara calling selama block
                        format = '%H:%M:%S'
                        duration = datetime.strptime(time_detected_last, format) - datetime.strptime(time_start, format)
                        if int(duration.seconds) > int(block):
                            c.execute("SELECT call_wh FROM tb_calling where no_polisi = ?",(no_polisi_last,))
                            rec = c.fetchone()
                            call_wh = rec[0]
                            file.commit()
                            if call_wh == "True":
                                # cek apakah covering atau tidak
                                c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
                                rec = c.fetchone()
                                flow_cc = rec[0]
                                file.commit()
                                # input to tb_playSound
                                c.execute("INSERT INTO tb_playSound (location, activity, ticket, covering_status) VALUES (?,?,?,?)",
                                    (location_last_temp,
                                    activity,
                                    ticket,
                                    flow_cc)
                                )
                                file.commit()
                                # update tb_calling
                                c.execute("UPDATE tb_calling SET call_wh=? WHERE no_polisi = ?",
                                    ("False",
                                    no_polisi_last)
                                )
                                file.commit()
                            # wb1 false -> true | getin true -> false
                            c.execute("UPDATE tb_queuing_flow SET getin=?, weight_bridge1=? WHERE no_polisi = ?",
                                ("False",
                                "True",
                                no_polisi_last)
                            )
                            file.commit()   
                    # get location start
                    c.execute("SELECT location FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    location_start = rec[0]
                    file.commit()
                    # jika location_last_temp == location start insert and update table last
                    if location_last_temp == location_start:
                        try:
                            c.execute("MERGE INTO tb_temporaryRFID_last AS target USING (VALUES (?, ?, ?, ?)) AS source (rfid, no_polisi, location, time_detected) ON target.no_polisi = source.no_polisi WHEN MATCHED THEN  UPDATE SET target.rfid = source.rfid, target.location = source.location, target.time_detected = source.time_detected WHEN NOT MATCHED THEN  INSERT (rfid, no_polisi, location, time_detected) VALUES (source.rfid, source.no_polisi, source.location, source.time_detected);",
                                (rfid_last,
                                no_polisi_last,
                                location_last_temp,
                                time_detected_last)
                            )
                            file.commit()
                        except:
                            pass
                    # hitung duration wb1 (start last) -> cek sla
                    # get range berat min max di tb_warehouseSla
                    c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla where packaging = ? AND product_grp = ?",(packaging, product))
                    rec = c.fetchall()
                    file.commit()
                    for i in rec:
                        weight_min = i[0]
                        weight_max = i[1]
                        # get SLA wb1 where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                        if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                            c.execute("SELECT w_bridge1 FROM tb_warehouseSla where packaging = ? AND product_grp = ? AND weight_min = ? AND weight_max = ?",(packaging, product, weight_min, weight_max))
                            rec = c.fetchone()
                            sla_wb1 = rec[0]
                            file.commit()
                    # cek apakah lebih dari sla?
                    c.execute("SELECT wb1_in FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    wb1_in_time = rec[0]
                    file.commit()
                    sla_wb1 = int(sla_wb1)*60
                    waktu_now = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
                    format = '%H:%M:%S'
                    duration = datetime.strptime(waktu_now, format) - datetime.strptime(wb1_in_time, format)
                    c.execute("SELECT over_wb1 FROM tb_overSLA where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    over_wb1 = rec[0]
                    file.commit()
                    if over_wb1 == "False":
                        if int(duration.seconds) > sla_wb1:
                            # masuk alert
                            date = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
                            waktu = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
                            status = "WAITING"
                            category = "OVER SLA"
                            message = "Truck {} over SLA at Jembatan Timbang 1".format(no_polisi_last)
                            c.execute("INSERT INTO tb_alert (date,time,status,category,message,notes,no_polisi,arrival) VALUES (?,?,?,?,?,?,?,?)",
                                (date,
                                waktu,
                                status,
                                category,
                                message,
                                "",
                                no_polisi_last,
                                arrival)
                            )
                            file.commit()
                            c.execute("UPDATE tb_overSLA SET over_wb1=? WHERE no_polisi = ?",
                                ("True",
                                no_polisi_last)
                            )
                            file.commit()

        #> location warehouse
        if location_last_temp == "Warehouse":
            # check apakah packaging dan activity ini perlu kesini
            c.execute("SELECT warehouse FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
            rec = c.fetchone()
            flow_wh = rec[0]
            file.commit()
            if flow_wh == "YES":
                # get gudang berdasarkan penugasan
                c.execute("SELECT wh_id FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                rec = c.fetchone()
                wh = rec[0]
                file.commit()

                loc_gudang = "Gudang {}".format(int(wh))

                if loc_gudang == "Gudang {}".format(int(wh)):
                    # jika location = wh check apakah wb1 sudah true? jika sudah next proses
                    c.execute("SELECT weight_bridge1 FROM tb_queuing_flow where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    Qflow_wb1 = rec[0]
                    file.commit()
                    if Qflow_wb1 == "True":
                        # get location start
                        c.execute("SELECT location FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        location_start = rec[0]
                        file.commit()
                        if location_last_temp != location_start:
                            # get time last wb1, input wb1-out time tabel queuing status
                            c.execute("SELECT time_detected FROM tb_temporaryRFID_last where no_polisi = ?",(no_polisi_last,))
                            rec = c.fetchone()
                            time_wb1_last = rec[0]
                            file.commit()
                            c.execute("UPDATE tb_queuing_status SET wb1_out=? WHERE no_polisi = ?",
                                (time_wb1_last,
                                no_polisi_last)
                            )
                            file.commit()
                            # perbarui total_minute tabel queuing status
                            c.execute("SELECT wb1_out, arrival FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                            rec = c.fetchone()
                            wb1_out_time = rec[0]
                            arrival = rec[1]
                            file.commit()
                            # hitung durasi
                            format = '%H:%M:%S'
                            duration = datetime.strptime(wb1_out_time, format) - datetime.strptime(arrival, format)
                            c.execute("UPDATE tb_queuing_status SET total_minute=? WHERE no_polisi = ?",
                                (str(duration),
                                no_polisi_last)
                            )
                            file.commit()
                            # input wh-in
                            c.execute("UPDATE tb_queuing_status SET wh_in=? WHERE no_polisi = ?",
                                (time_detected_last,
                                no_polisi_last)
                            )
                            file.commit()
                            # perbarui total minute
                            c.execute("SELECT wh_in, arrival FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                            rec = c.fetchone()
                            wh_in_time = rec[0]
                            arrival = rec[1]
                            file.commit()
                            # hitung durasi
                            format = '%H:%M:%S'
                            duration = datetime.strptime(wh_in_time, format) - datetime.strptime(arrival, format)
                            c.execute("UPDATE tb_queuing_status SET total_minute=? WHERE no_polisi = ?",
                                (str(duration),
                                no_polisi_last)
                            )
                            file.commit()
                            # perbahrui location start
                            try:
                                c.execute("MERGE INTO tb_temporaryRFID_start AS target USING (VALUES (?, ?, ?, ?)) AS source (rfid, no_polisi, location, time_detected) ON target.no_polisi = source.no_polisi WHEN MATCHED THEN  UPDATE SET target.rfid = source.rfid, target.location = source.location, target.time_detected = source.time_detected WHEN NOT MATCHED THEN  INSERT (rfid, no_polisi, location, time_detected) VALUES (source.rfid, source.no_polisi, source.location, source.time_detected);",
                                    (rfid_last,
                                    no_polisi_last,
                                    location_last_temp,
                                    time_detected_last)
                                )
                                file.commit()
                            except:
                                pass
                            # suara tappin wh
                            # cek apakah covering atau tidak
                            c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
                            rec = c.fetchone()
                            flow_cc = rec[0]
                            file.commit()
                            # input to tb_playSound
                            c.execute("INSERT INTO tb_playSound (location, activity, ticket, covering_status) VALUES (?,?,?,?)",
                                (location_last_temp,
                                "tap_in",
                                ticket,
                                flow_cc)
                            )
                            file.commit()
                        
                        # get block in tb_tappingPointReg where function == Weight Bridge 1
                        c.execute("SELECT block FROM tb_tappingPointReg where functions = ?",("Warehouse",))
                        rec = c.fetchone()
                        block = rec[0]
                        file.commit()
                        # get time start dan last
                        c.execute("SELECT time_detected FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        time_start = rec[0]
                        file.commit()
                        # blocking suara calling selama block
                        format = '%H:%M:%S'
                        duration = datetime.strptime(time_detected_last, format) - datetime.strptime(time_start, format)
                        if int(duration.seconds) > int(block):
                            c.execute("SELECT call_wb2 FROM tb_calling where no_polisi = ?",(no_polisi_last,))
                            rec = c.fetchone()
                            call_wb2 = rec[0]
                            file.commit()
                            if call_wb2 == "True":
                                # cek apakah covering atau tidak
                                c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
                                rec = c.fetchone()
                                flow_cc = rec[0]
                                file.commit()
                                # input to tb_playSound
                                c.execute("INSERT INTO tb_playSound (location, activity, ticket, covering_status) VALUES (?,?,?,?)",
                                    (location_last_temp,
                                    "tap_out",
                                    ticket,
                                    flow_cc)
                                )
                                file.commit()
                                # update tb_calling
                                c.execute("UPDATE tb_calling SET call_wb2=? WHERE no_polisi = ?",
                                    ("False",
                                    no_polisi_last)
                                )
                                file.commit()
                            
                            # wb1 false -> true | getin true -> false
                            c.execute("UPDATE tb_queuing_flow SET weight_bridge1=?, warehouse=? WHERE no_polisi = ?",
                                ("False",
                                "True",
                                no_polisi_last)
                            )
                            file.commit()
                    # perbarui last
                    # get location start
                    c.execute("SELECT location FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    location_start = rec[0]
                    file.commit()
                    # jika location_last_temp == location start insert and update table last
                    if location_last_temp == location_start:
                        try:
                            c.execute("MERGE INTO tb_temporaryRFID_last AS target USING (VALUES (?, ?, ?, ?)) AS source (rfid, no_polisi, location, time_detected) ON target.no_polisi = source.no_polisi WHEN MATCHED THEN  UPDATE SET target.rfid = source.rfid, target.location = source.location, target.time_detected = source.time_detected WHEN NOT MATCHED THEN  INSERT (rfid, no_polisi, location, time_detected) VALUES (source.rfid, source.no_polisi, source.location, source.time_detected);",
                                (rfid_last,
                                no_polisi_last,
                                location_last_temp,
                                time_detected_last)
                            )
                            file.commit()
                        except:
                            pass

        #> location wb2
        if location_last_temp == "Weight Bridge 2":
            # check apakah packaging dan activity ini perlu kesini
            c.execute("SELECT weightbridge2 FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
            rec = c.fetchone()
            flow_wb2 = rec[0]
            file.commit()
            if flow_wb2 == "YES":
                # jika location = wb 2 check apakah wh sudah true? jika sudah next proses
                c.execute("SELECT warehouse FROM tb_queuing_flow where no_polisi = ?",(no_polisi_last,))
                rec = c.fetchone()
                Qflow_wh = rec[0]
                file.commit()
                if Qflow_wh == "True":
                    # get location start
                    # jika location last tidak sama dengan location start, get time last wh, input wh-out time tabel queuing status
                    c.execute("SELECT location FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    location_start = rec[0]
                    file.commit()
                    if location_last_temp != location_start:
                        # get time last wb1, input wb1-out time tabel queuing status
                        c.execute("SELECT time_detected FROM tb_temporaryRFID_last where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        time_wh_last = rec[0]
                        file.commit()
                        c.execute("UPDATE tb_queuing_status SET wh_out=? WHERE no_polisi = ?",
                            (time_wh_last,
                            no_polisi_last)
                        )
                        file.commit()
                        # perbarui total_minute tabel queuing status
                        c.execute("SELECT wh_out, arrival FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        wh_out_time = rec[0]
                        arrival = rec[1]
                        file.commit()
                        # hitung durasi
                        format = '%H:%M:%S'
                        duration = datetime.strptime(wh_out_time, format) - datetime.strptime(arrival, format)
                        c.execute("UPDATE tb_queuing_status SET total_minute=? WHERE no_polisi = ?",
                            (str(duration),
                            no_polisi_last)
                        )
                        file.commit()
                        # input wb-in
                        c.execute("UPDATE tb_queuing_status SET wb2_in=? WHERE no_polisi = ?",
                            (time_detected_last,
                            no_polisi_last)
                        )
                        file.commit()
                        # perbarui total minute
                        c.execute("SELECT wb2_in, arrival FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        wb2_in_time = rec[0]
                        arrival = rec[1]
                        file.commit()
                        # hitung durasi
                        format = '%H:%M:%S'
                        duration = datetime.strptime(wb2_in_time, format) - datetime.strptime(arrival, format)
                        c.execute("UPDATE tb_queuing_status SET total_minute=? WHERE no_polisi = ?",
                            (str(duration),
                            no_polisi_last)
                        )
                        file.commit()
                        # perbahrui location start
                        try:
                            c.execute("MERGE INTO tb_temporaryRFID_start AS target USING (VALUES (?, ?, ?, ?)) AS source (rfid, no_polisi, location, time_detected) ON target.no_polisi = source.no_polisi WHEN MATCHED THEN  UPDATE SET target.rfid = source.rfid, target.location = source.location, target.time_detected = source.time_detected WHEN NOT MATCHED THEN  INSERT (rfid, no_polisi, location, time_detected) VALUES (source.rfid, source.no_polisi, source.location, source.time_detected);",
                                (rfid_last,
                                no_polisi_last,
                                location_last_temp,
                                time_detected_last)
                            )
                            file.commit()
                        except:
                            pass

                    # blocking suara calling selama xx
                    c.execute("SELECT block FROM tb_tappingPointReg where functions = ?",("Weight Bridge 2",))
                    rec = c.fetchone()
                    block = rec[0]
                    file.commit()
                    # get time start dan last
                    c.execute("SELECT time_detected FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    time_start = rec[0]
                    file.commit()
                    # blocking suara calling selama block
                    format = '%H:%M:%S'
                    duration = datetime.strptime(time_detected_last, format) - datetime.strptime(time_start, format)
                    if int(duration.seconds) > int(block):
                        c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
                        rec = c.fetchone()
                        flow_cc = rec[0]
                        file.commit()
                        if flow_cc == "YES":
                            c.execute("SELECT call_cc FROM tb_calling where no_polisi = ?",(no_polisi_last,))
                            rec = c.fetchone()
                            call_cc = rec[0]
                            file.commit()
                            if call_cc == "True":
                                # cek apakah covering atau tidak
                                c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
                                rec = c.fetchone()
                                flow_cc = rec[0]
                                file.commit()
                                # input to tb_playSound
                                c.execute("INSERT INTO tb_playSound (location, activity, ticket, covering_status) VALUES (?,?,?,?)",
                                    (location_last_temp,
                                    activity,
                                    ticket,
                                    flow_cc)
                                )
                                file.commit()
                                # Update tb_calling
                                c.execute("UPDATE tb_calling SET call_cc=? WHERE no_polisi = ?",
                                    ("False",
                                    no_polisi_last)
                                )
                                file.commit()
                            # wb1 false -> true | getin true -> false
                            c.execute("UPDATE tb_queuing_flow SET warehouse=?, weight_bridge2=? WHERE no_polisi = ?",
                                ("False",
                                "True",
                                no_polisi_last)
                            )
                            file.commit()
                        else:
                            c.execute("SELECT call_getout FROM tb_calling where no_polisi = ?",(no_polisi_last,))
                            rec = c.fetchone()
                            call_getout = rec[0]
                            file.commit()
                            if call_getout == "True":
                                # cek apakah covering atau tidak
                                c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
                                rec = c.fetchone()
                                flow_cc = rec[0]
                                file.commit()
                                # input to tb_playSound
                                c.execute("INSERT INTO tb_playSound (location, activity, ticket, covering_status) VALUES (?,?,?,?)",
                                    (location_last_temp,
                                    activity,
                                    ticket,
                                    flow_cc)
                                )
                                file.commit()
                                # Update tb_calling
                                c.execute("UPDATE tb_calling SET call_getout=? WHERE no_polisi = ?",
                                    ("False",
                                    no_polisi_last)
                                )
                                file.commit()
                            # wb1 false -> true | getin true -> false
                            c.execute("UPDATE tb_queuing_flow SET warehouse=?, weight_bridge2=? WHERE no_polisi = ?",
                                ("False",
                                "True",
                                no_polisi_last)
                            )
                            file.commit()
                # perbahrui location last
                c.execute("SELECT location FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                rec = c.fetchone()
                location_start = rec[0]
                file.commit()
                # jika location_last_temp == location start insert and update table last
                if location_last_temp == location_start:
                    try:
                        c.execute("MERGE INTO tb_temporaryRFID_last AS target USING (VALUES (?, ?, ?, ?)) AS source (rfid, no_polisi, location, time_detected) ON target.no_polisi = source.no_polisi WHEN MATCHED THEN  UPDATE SET target.rfid = source.rfid, target.location = source.location, target.time_detected = source.time_detected WHEN NOT MATCHED THEN  INSERT (rfid, no_polisi, location, time_detected) VALUES (source.rfid, source.no_polisi, source.location, source.time_detected);",
                            (rfid_last,
                            no_polisi_last,
                            location_last_temp,
                            time_detected_last)
                        )
                        file.commit()
                    except:
                        pass
                # hitung duration wb1 (start last) -> cek sla
                # get range berat min max di tb_warehouseSla
                c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla where packaging = ? AND product_grp = ?",(packaging, product))
                rec = c.fetchall()
                file.commit()
                for i in rec:
                    weight_min = i[0]
                    weight_max = i[1]
                    # get SLA wb2 where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                    if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                        c.execute("SELECT w_bridge2 FROM tb_warehouseSla where packaging = ? AND product_grp = ? AND weight_min = ? AND weight_max = ?",(packaging, product, weight_min, weight_max))
                        rec = c.fetchone()
                        sla_wb2 = rec[0]
                        file.commit()
                # cek apakah lebih dari sla?
                c.execute("SELECT wb2_in FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                rec = c.fetchone()
                wb2_in_time = rec[0]
                file.commit()
                sla_wb2 = int(sla_wb2)*60
                waktu_now = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
                format = '%H:%M:%S'
                duration = datetime.strptime(waktu_now, format) - datetime.strptime(wb2_in_time, format)
                c.execute("SELECT over_wb2 FROM tb_overSLA where no_polisi = ?",(no_polisi_last,))
                rec = c.fetchone()
                over_wb2 = rec[0]
                file.commit()
                if over_wb2 == "False":
                    if int(duration.seconds) > sla_wb2:
                        # masuk alert
                        date = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
                        waktu = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
                        status = "WAITING"
                        category = "OVER SLA"
                        message = "Truck {} over SLA at Jembatan Timbang 2".format(no_polisi_last)
                        c.execute("INSERT INTO tb_alert (date,time,status,category,message,notes,no_polisi,arrival) VALUES (?,?,?,?,?,?,?,?)",
                            (date,
                            waktu,
                            status,
                            category,
                            message,
                            "",
                            no_polisi_last,
                            arrival)
                        )
                        file.commit()
                        c.execute("UPDATE tb_overSLA SET over_wb2=? WHERE no_polisi = ?",
                            ("True",
                            no_polisi_last)
                        )
                        file.commit()

        #> location cargo covering
        if location_last_temp == "Cargo Covering":
            # check apakah packaging dan activity ini perlu kesini
            c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
            rec = c.fetchone()
            flow_cc = rec[0]
            file.commit()
            if flow_cc == "YES":
                # jika location = cc check apakah wb2 sudah true? jika sudah next proses
                c.execute("SELECT weight_bridge2 FROM tb_queuing_flow where no_polisi = ?",(no_polisi_last,))
                rec = c.fetchone()
                Qflow_wb2 = rec[0]
                file.commit()
                if Qflow_wb2 == "True":
                    c.execute("SELECT location FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    location_start = rec[0]
                    file.commit()
                    if location_last_temp != location_start:
                        # get time last wb2, input wb2-out time tabel queuing status
                        c.execute("SELECT time_detected FROM tb_temporaryRFID_last where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        time_wb2_last = rec[0]
                        file.commit()
                        c.execute("UPDATE tb_queuing_status SET wb2_out=? WHERE no_polisi = ?",
                            (time_wb2_last,
                            no_polisi_last)
                        )
                        file.commit()
                        # perbarui total_minute tabel queuing status
                        c.execute("SELECT wb2_out, arrival FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        wb2_out_time = rec[0]
                        arrival = rec[1]
                        file.commit()
                        # hitung durasi
                        format = '%H:%M:%S'
                        duration = datetime.strptime(wb2_out_time, format) - datetime.strptime(arrival, format)
                        c.execute("UPDATE tb_queuing_status SET total_minute=? WHERE no_polisi = ?",
                            (str(duration),
                            no_polisi_last)
                        )
                        file.commit()
                        # input cc-in
                        c.execute("UPDATE tb_queuing_status SET cc_in=? WHERE no_polisi = ?",
                            (time_detected_last,
                            no_polisi_last)
                        )
                        file.commit()
                        # perbarui total minute
                        c.execute("SELECT cc_in, arrival FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        cc_in_time = rec[0]
                        arrival = rec[1]
                        file.commit()
                        # hitung durasi
                        format = '%H:%M:%S'
                        duration = datetime.strptime(cc_in_time, format) - datetime.strptime(arrival, format)
                        c.execute("UPDATE tb_queuing_status SET total_minute=? WHERE no_polisi = ?",
                            (str(duration),
                            no_polisi_last)
                        )
                        file.commit()
                        # perbahrui location start
                        try:
                            c.execute("MERGE INTO tb_temporaryRFID_start AS target USING (VALUES (?, ?, ?, ?)) AS source (rfid, no_polisi, location, time_detected) ON target.no_polisi = source.no_polisi WHEN MATCHED THEN  UPDATE SET target.rfid = source.rfid, target.location = source.location, target.time_detected = source.time_detected WHEN NOT MATCHED THEN  INSERT (rfid, no_polisi, location, time_detected) VALUES (source.rfid, source.no_polisi, source.location, source.time_detected);",
                                (rfid_last,
                                no_polisi_last,
                                location_last_temp,
                                time_detected_last)
                            )
                            file.commit()
                        except:
                            pass

                # blocking suara calling selama xx
                c.execute("SELECT block FROM tb_tappingPointReg where functions = ?",("Cargo Covering",))
                rec = c.fetchone()
                block = rec[0]
                file.commit()
                # get time start dan last
                c.execute("SELECT time_detected FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                rec = c.fetchone()
                time_start = rec[0]
                file.commit()
                # blocking suara calling selama block
                format = '%H:%M:%S'
                duration = datetime.strptime(time_detected_last, format) - datetime.strptime(time_start, format)
                if int(duration.seconds) > int(block):
                    c.execute("SELECT call_getout FROM tb_calling where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    call_getout = rec[0]
                    file.commit()
                    if call_getout == "True":
                        # cek apakah covering atau tidak
                        c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
                        rec = c.fetchone()
                        flow_cc = rec[0]
                        file.commit()
                        # input to tb_playSound
                        c.execute("INSERT INTO tb_playSound (location, activity, ticket, covering_status) VALUES (?,?,?,?)",
                            (location_last_temp,
                            activity,
                            ticket,
                            flow_cc)
                        )
                        file.commit()
                        # Update tb_calling
                        c.execute("UPDATE tb_calling SET call_getout=? WHERE no_polisi = ?",
                            ("False",
                            no_polisi_last)
                        )
                        file.commit()
                    # wb1 false -> true | getin true -> false
                    c.execute("UPDATE tb_queuing_flow SET weight_bridge2=?, cargo_covering=? WHERE no_polisi = ?",
                        ("False",
                        "True",
                        no_polisi_last)
                    )
                    file.commit()
            # perbahrui location last
            c.execute("SELECT location FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
            rec = c.fetchone()
            location_start = rec[0]
            file.commit()
            # jika location_last_temp == location start insert and update table last
            if location_last_temp == location_start:
                try:
                    c.execute("MERGE INTO tb_temporaryRFID_last AS target USING (VALUES (?, ?, ?, ?)) AS source (rfid, no_polisi, location, time_detected) ON target.no_polisi = source.no_polisi WHEN MATCHED THEN  UPDATE SET target.rfid = source.rfid, target.location = source.location, target.time_detected = source.time_detected WHEN NOT MATCHED THEN  INSERT (rfid, no_polisi, location, time_detected) VALUES (source.rfid, source.no_polisi, source.location, source.time_detected);",
                        (rfid_last,
                        no_polisi_last,
                        location_last_temp,
                        time_detected_last)
                    )
                    file.commit()
                except:
                    pass
            # hitung duration wb1 (start last) -> cek sla
            # get range berat min max di tb_warehouseSla
            c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla where packaging = ? AND product_grp = ?",(packaging, product))
            rec = c.fetchall()
            file.commit()
            for i in rec:
                weight_min = i[0]
                weight_max = i[1]
                # get SLA wb2 where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                    c.execute("SELECT covering_time FROM tb_warehouseSla where packaging = ? AND product_grp = ? AND weight_min = ? AND weight_max = ?",(packaging, product, weight_min, weight_max))
                    rec = c.fetchone()
                    sla_covering = rec[0]
                    file.commit()
            # cek apakah lebih dari sla?
            c.execute("SELECT cc_in FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
            rec = c.fetchone()
            cc_in_time = rec[0]
            file.commit()
            sla_covering = int(sla_covering)*60
            waktu_now = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
            format = '%H:%M:%S'
            duration = datetime.strptime(waktu_now, format) - datetime.strptime(cc_in_time, format)
            c.execute("SELECT over_cc FROM tb_overSLA where no_polisi = ?",(no_polisi_last,))
            rec = c.fetchone()
            over_cc = rec[0]
            file.commit()
            if over_cc == "False":
                if int(duration.seconds) > sla_covering:
                    # masuk alert
                    date = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
                    waktu = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
                    status = "WAITING"
                    category = "OVER SLA"
                    message = "Truck {} over SLA at Cargo Covering".format(no_polisi_last)
                    c.execute("INSERT INTO tb_alert (date,time,status,category,message,notes,no_polisi,arrival) VALUES (?,?,?,?,?,?,?,?)",
                        (date,
                        waktu,
                        status,
                        category,
                        message,
                        "",
                        no_polisi_last,
                        arrival)
                    )
                    file.commit()
                    c.execute("UPDATE tb_overSLA SET over_cc=%s WHERE no_polisi = %s",
                        ("True",
                        no_polisi_last)
                    )
                    file.commit()

    def gate_out(self):
        # get no police
        ticket = self.entry_noTiket_rfidPairing.get()
        location = self.optionmenu_manualTaping.get()
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT no_polisi, activity, packaging, product FROM tb_queuing_status where ticket = ?",(ticket,))
        rec = c.fetchone()
        no_polisi = rec[0]
        activity = rec[1]
        packaging = rec[2]
        product = rec[3]
        file.commit()
        file.close()

        #4 get kuota transaksi based on activity -> tabel unloading/loading (activity=remark)
        if activity ==  "LOADING":
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT berat, no_so FROM tb_loadingList where no_polisi = ? and status = ? and product = ?", (no_polisi, "Activate", product))
            rec = c.fetchone()
            berat, no_so = rec
            file.commit()
            file.close()
        elif activity == "UNLOADING":
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT berat, no_so FROM tb_unloading where no_polisi = ? and status = ? and product = ?", (no_polisi, "Activate", product))
            rec = c.fetchone()
            berat, no_so = rec
            file.commit()
            file.close()

        #get bluetooth name
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT rfid1 FROM tb_vetting where no_polisi = ?",(no_polisi,))
        rec = c.fetchone()
        blth_name = rec[0]
        file.commit()
        file.close()

        # insert to tb_temporaryRFID
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT getin, weight_bridge1, warehouse, cargo_covering, weight_bridge2, getout FROM tb_queuing_flow where no_polisi = ?",(no_polisi,))
        rec = c.fetchone()
        cek_getin, cek_weight_bridge1, cek_warehouse, cek_cargo_covering, cek_weight_bridge2, cek_getout = rec
        time_detected = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
        if cek_getin == "True":
            if location == "Weight Bridge 1":
                # Insert into tb_temporaryRFID
                c.execute("INSERT INTO tb_temporaryRFID (rfid, no_polisi, location, time_detected) VALUES (?, ?, ?, ?)",(blth_name, no_polisi, location, time_detected))
                file.commit()
        if cek_weight_bridge1 == "True":
            if location == "Weight Bridge 1" or location == "Warehouse":
                # Insert into tb_temporaryRFID
                c.execute("INSERT INTO tb_temporaryRFID (rfid, no_polisi, location, time_detected) VALUES (?, ?, ?, ?)",(blth_name, no_polisi, location, time_detected))
                file.commit()
        if cek_warehouse == "True":
            if location == "Warehouse" or location == "Weight Bridge 2":
                # Insert into tb_temporaryRFID
                c.execute("INSERT INTO tb_temporaryRFID (rfid, no_polisi, location, time_detected) VALUES (?, ?, ?, ?)",(blth_name, no_polisi, location, time_detected))
                file.commit()
        if cek_weight_bridge2 == "True":
            c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
            rec = c.fetchone()
            cek_cc = rec[0]
            if location == "Weight Bridge 2" or (location == "Cargo Covering" and cek_cc == "YES") or (location == "Unpairing/Unregistration" and cek_cc == "NO"):
                # Insert into tb_temporaryRFID
                c.execute("INSERT INTO tb_temporaryRFID (rfid, no_polisi, location, time_detected) VALUES (?, ?, ?, ?)",(blth_name, no_polisi, location, time_detected))
                file.commit()
        if cek_cargo_covering == "True":
            if location == "Cargo Covering":
                # Insert into tb_temporaryRFID
                c.execute("INSERT INTO tb_temporaryRFID (rfid, no_polisi, location, time_detected) VALUES (?, ?, ?, ?)",(blth_name, no_polisi, location, time_detected))
                file.commit()

        #
        c.execute("SELECT TOP 1 oid, rfid, no_polisi, location, time_detected FROM tb_temporaryRFID where no_polisi = ? ORDER BY oid DESC",(no_polisi,))
        rec = c.fetchone()
        rfid_last = rec[1]
        no_polisi_last = rec[2]
        location_last_temp = rec[3]
        time_detected_last = rec[4]
        file.commit()

        # Gateout
        location_last_temp = "Unpairing/Unregistration"
        c.execute("SELECT unregistration FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
        rec = c.fetchone()
        flow_unregis = rec[0]
        file.commit()
        if flow_unregis == "YES":
            c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
            rec = c.fetchone()
            flow_cc = rec[0]
            file.commit()
            if flow_cc == "YES":
                # jika location = unregis check apakah cc sudah true atau wb2 true? jika sudah next proses
                c.execute("SELECT cargo_covering FROM tb_queuing_flow where no_polisi = ?",(no_polisi_last,))
                rec = c.fetchone()
                Qflow_cc = rec[0]
                file.commit()
                if Qflow_cc == "True":
                    c.execute("SELECT location FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    location_start = rec[0]
                    file.commit()
                    if location_last_temp != location_start:
                        # get time last cc, input cc-out time tabel queuing status
                        c.execute("SELECT time_detected FROM tb_temporaryRFID_last where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        time_cc_last = rec[0]
                        file.commit()
                        c.execute("UPDATE tb_queuing_status SET cc_out=? WHERE no_polisi = ?",
                            (time_cc_last,
                            no_polisi_last)
                        )
                        file.commit()
                        # perbarui total_minute tabel queuing status
                        c.execute("SELECT cc_out, arrival FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        cc_out_time = rec[0]
                        arrival = rec[1]
                        file.commit()
                        # hitung durasi
                        format = '%H:%M:%S'
                        duration = datetime.strptime(cc_out_time, format) - datetime.strptime(arrival, format)
                        c.execute("UPDATE tb_queuing_status SET total_minute=? WHERE no_polisi = ?",
                            (str(duration),
                            no_polisi_last)
                        )
                        file.commit()
                        # input unpairing
                        c.execute("UPDATE tb_queuing_status SET unpairing=? WHERE no_polisi = ?",
                            (time_detected_last,
                            no_polisi_last)
                        )
                        file.commit()
                        # perbarui total minute
                        c.execute("SELECT unpairing, arrival FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        unpairing_time = rec[0]
                        arrival = rec[1]
                        file.commit()
                        # hitung durasi
                        format = '%H:%M:%S'
                        duration = datetime.strptime(unpairing_time, format) - datetime.strptime(arrival, format)
                        c.execute("UPDATE tb_queuing_status SET total_minute=? WHERE no_polisi = ?",
                            (str(duration),
                            no_polisi_last)
                        )
                        file.commit()
                        # get company and category
                        c.execute("SELECT company, category FROM tb_vetting where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        company = rec[0]
                        category = rec[1]
                        file.commit()
                        # masukan data status ke tb_queuing_history
                        c.execute("SELECT * FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchall()
                        file.commit()
                        for i in rec:
                            # hitung waiting_minute
                            format = '%H:%M:%S'
                            waiting_minute = datetime.strptime(i[7], format) - datetime.strptime(i[6], format)
                            if waiting_minute.seconds/60 >= 1:
                                waiting_minute = str(round(waiting_minute.seconds/60))
                            else:
                                waiting_minute = str(round(waiting_minute.seconds/60,2))
                            # hitung wb1_minute
                            format = '%H:%M:%S'
                            wb1_minute = datetime.strptime(i[10], format) - datetime.strptime(i[9], format)
                            if wb1_minute.seconds/60 >= 1:
                                wb1_minute = str(round(wb1_minute.seconds/60))
                            else:
                                wb1_minute = str(round(wb1_minute.seconds/60,2))
                            # hitung wh_minute
                            format = '%H:%M:%S'
                            wh_minute = datetime.strptime(i[12], format) - datetime.strptime(i[11], format)
                            if wh_minute.seconds/60 >= 1:
                                wh_minute = str(round(wh_minute.seconds/60))
                            else:
                                wh_minute = str(round(wh_minute.seconds/60,2))
                            # hitung wb2_minute
                            format = '%H:%M:%S'
                            wb2_minute = datetime.strptime(i[14], format) - datetime.strptime(i[13], format)
                            if wb2_minute.seconds/60 >= 1:
                                wb2_minute = str(round(wb2_minute.seconds/60))
                            else:
                                wb2_minute = str(round(wb2_minute.seconds/60,2))
                            # hitung cc_minute
                            format = '%H:%M:%S'
                            cc_minute = datetime.strptime(i[16], format) - datetime.strptime(i[15], format)
                            if cc_minute.seconds/60 >= 1:
                                cc_minute = str(round(cc_minute.seconds/60))
                            else:
                                cc_minute = str(round(cc_minute.seconds/60,2))
                            date_now = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
                            # insert to tb_queuing_history
                            c.execute("INSERT INTO tb_queuing_history (date,rfid,ticket,no_polisi,activity,packaging,product,arrival,calling,pairing,wb1_in,wb1_out,wh_in,wh_out,wb2_in,wb2_out,cc_in,cc_out,unpairing,total_minute,wh_id,warehouse,berat,waiting_minute,wb1_minute,wh_minute,wb2_minute,cc_minute,company,category,no_so) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                (date_now,
                                i[0],
                                i[1],
                                i[2],
                                i[3],
                                i[4],
                                i[5],
                                i[6],
                                i[7],
                                i[8],
                                i[9],
                                i[10],
                                i[11],
                                i[12],
                                i[13],
                                i[14],
                                i[15],
                                i[16],
                                i[17],
                                i[18],
                                i[19],
                                wh_name,
                                berat,
                                waiting_minute,
                                wb1_minute,
                                wh_minute,
                                wb2_minute,
                                cc_minute,
                                company,
                                category,
                                no_so)
                            )
                            file.commit()
                        # perbahrui location start
                        try:
                            c.execute("MERGE INTO tb_temporaryRFID_start AS target USING (VALUES (?, ?, ?, ?)) AS source (rfid, no_polisi, location, time_detected) ON target.no_polisi = source.no_polisi WHEN MATCHED THEN  UPDATE SET target.rfid = source.rfid, target.location = source.location, target.time_detected = source.time_detected WHEN NOT MATCHED THEN  INSERT (rfid, no_polisi, location, time_detected) VALUES (source.rfid, source.no_polisi, source.location, source.time_detected);",
                                (rfid_last,
                                no_polisi_last,
                                location_last_temp,
                                time_detected_last)
                            )
                            file.commit()
                        except:
                            pass
                    # getout false -> true | cc true -> false
                    c.execute("UPDATE tb_queuing_flow SET cargo_covering=?, getout=? WHERE no_polisi = ?",
                        ("False",
                        "True",
                        no_polisi_last)
                    )
                    file.commit()
                    # blocking suara calling selama xx
                    c.execute("SELECT block FROM tb_tappingPointReg where functions = ?",("Unpairing/Unregistration",))
                    rec = c.fetchone()
                    block = rec[0]
                    file.commit()
                    # get time start dan last
                    c.execute("SELECT time_detected FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    time_start = rec[0]
                    file.commit()
                    c.execute("SELECT time_detected FROM tb_temporaryRFID_last where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    time_last = rec[0]
                    file.commit()
                    # blocking suara calling selama block
                    format = '%H:%M:%S'
                    duration = datetime.strptime(time_last, format) - datetime.strptime(time_start, format)
                    if int(duration.seconds) > int(block):
                        c.execute("SELECT call_end FROM tb_calling where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        call_end = rec[0]
                        file.commit()
                        if call_end == "True":
                            if activity == "LOADING":
                                # cek apakah covering atau tidak
                                c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
                                rec = c.fetchone()
                                flow_cc = rec[0]
                                file.commit()
                                # input to tb_playSound
                                c.execute("INSERT INTO tb_playSound (location, activity, ticket, covering_status) VALUES (?,?,?,?)",
                                    (location_last_temp,
                                    activity,
                                    ticket,
                                    flow_cc)
                                )
                                file.commit()
                                # Update tb_calling
                                c.execute("UPDATE tb_calling SET call_end=? WHERE no_polisi = ?",
                                    ("False",
                                    no_polisi_last)
                                )
                                file.commit()
                            else:
                                # cek apakah covering atau tidak
                                c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
                                rec = c.fetchone()
                                flow_cc = rec[0]
                                file.commit()
                                # input to tb_playSound
                                c.execute("INSERT INTO tb_playSound (location, activity, ticket, covering_status) VALUES (?,?,?,?)",
                                    (location_last_temp,
                                    activity,
                                    ticket,
                                    flow_cc)
                                )
                                file.commit()
                                # Update tb_calling
                                c.execute("UPDATE tb_calling SET call_end=? WHERE no_polisi = ?",
                                    ("False",
                                    no_polisi_last)
                                )
                                file.commit()
                            time.sleep(5)
                            try:
                                # hapus semua data dengan plat no yg getout di tb_temporaryRFID
                                c.execute("DELETE FROM tb_temporaryRFID where no_polisi = ?",(no_polisi_last,))
                                file.commit()
                            except:
                                pass
                            try:
                                # hapus yg di tb_calling
                                c.execute("DELETE FROM tb_calling where no_polisi = ?",(no_polisi_last,))
                                file.commit()
                            except:
                                pass
                            try:
                                # hapus yg di tb_overSLA
                                c.execute("DELETE FROM tb_overSLA where no_polisi = ?",(no_polisi_last,))
                                file.commit()
                            except:
                                pass
                            try:
                                # hapus yg di tb_queuing_status
                                c.execute("DELETE FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                                file.commit()
                            except:
                                pass
                            try:
                                # hapus yg di tb_queuing_flow
                                c.execute("DELETE FROM tb_queuing_flow where no_polisi = ?",(no_polisi_last,))
                                file.commit()
                            except:
                                pass
                            try:
                                # hapus yg di tb_temporaryRFID_start
                                c.execute("DELETE FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                                file.commit()
                            except:
                                pass
                            try:
                                # hapus yg di tb_calling_time
                                c.execute("DELETE FROM tb_calling_time where no_polisi = ?",(no_polisi_last,))
                                file.commit()
                            except:
                                pass
            else:
                c.execute("SELECT weight_bridge2 FROM tb_queuing_flow where no_polisi = ?",(no_polisi_last,))
                rec = c.fetchone()
                Qflow_wb2 = rec[0]
                file.commit()
                if Qflow_wb2 == "True":
                    c.execute("SELECT location FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    location_start = rec[0]
                    file.commit()
                    if location_last_temp != location_start:
                        # get time last wb2, input wb2-out time tabel queuing status
                        c.execute("SELECT time_detected FROM tb_temporaryRFID_last where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        time_wb2_last = rec[0]
                        file.commit()
                        c.execute("UPDATE tb_queuing_status SET wb2_out=? WHERE no_polisi = ?",
                            (time_wb2_last,
                            no_polisi_last)
                        )
                        file.commit()
                        # perbarui total_minute tabel queuing status
                        c.execute("SELECT wb2_out, arrival FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        wb2_out_time = rec[0]
                        arrival = rec[1]
                        file.commit()
                        # hitung durasi
                        format = '%H:%M:%S'
                        duration = datetime.strptime(wb2_out_time, format) - datetime.strptime(arrival, format)
                        c.execute("UPDATE tb_queuing_status SET total_minute=? WHERE no_polisi = ?",
                            (str(duration),
                            no_polisi_last)
                        )
                        file.commit()
                        # input unpairing
                        c.execute("UPDATE tb_queuing_status SET unpairing=? WHERE no_polisi = ?",
                            (time_detected_last,
                            no_polisi_last)
                        )
                        file.commit()
                        # perbarui total minute
                        c.execute("SELECT unpairing, arrival FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        unpairing_time = rec[0]
                        arrival = rec[1]
                        file.commit()
                        # hitung durasi
                        format = '%H:%M:%S'
                        duration = datetime.strptime(unpairing_time, format) - datetime.strptime(arrival, format)
                        c.execute("UPDATE tb_queuing_status SET total_minute=? WHERE no_polisi = ?",
                            (str(duration),
                            no_polisi_last)
                        )
                        file.commit()
                        # get company and category
                        c.execute("SELECT company, category FROM tb_vetting where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        company = rec[0]
                        category = rec[1]
                        file.commit()
                        # masukan data status ke tb_queuing_history
                        c.execute("SELECT * FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchall()
                        file.commit()
                        for i in rec:
                            # hitung waiting_minute
                            format = '%H:%M:%S'
                            waiting_minute = datetime.strptime(i[7], format) - datetime.strptime(i[6], format)
                            if waiting_minute.seconds/60 >= 1:
                                waiting_minute = str(round(waiting_minute.seconds/60))
                            else:
                                waiting_minute = str(round(waiting_minute.seconds/60,2))
                            # hitung wb1_minute
                            format = '%H:%M:%S'
                            wb1_minute = datetime.strptime(i[10], format) - datetime.strptime(i[9], format)
                            if wb1_minute.seconds/60 >= 1:
                                wb1_minute = str(round(wb1_minute.seconds/60))
                            else:
                                wb1_minute = str(round(wb1_minute.seconds/60,2))
                            # hitung wh_minute
                            format = '%H:%M:%S'
                            wh_minute = datetime.strptime(i[12], format) - datetime.strptime(i[11], format)
                            if wh_minute.seconds/60 >= 1:
                                wh_minute = str(round(wh_minute.seconds/60))
                            else:
                                wh_minute = str(round(wh_minute.seconds/60,2))
                            # hitung wb2_minute
                            format = '%H:%M:%S'
                            wb2_minute = datetime.strptime(i[14], format) - datetime.strptime(i[13], format)
                            if wb2_minute.seconds/60 >= 1:
                                wb2_minute = str(round(wb2_minute.seconds/60))
                            else:
                                wb2_minute = str(round(wb2_minute.seconds/60,2))
                            # hitung cc_minute
                            cc_minute = ""
                            date_now = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
                            # insert to tb_queuing_history
                            c.execute("INSERT INTO tb_queuing_history (date,rfid,ticket,no_polisi,activity,packaging,product,arrival,calling,pairing,wb1_in,wb1_out,wh_in,wh_out,wb2_in,wb2_out,cc_in,cc_out,unpairing,total_minute,wh_id,warehouse,berat,waiting_minute,wb1_minute,wh_minute,wb2_minute,cc_minute,company,category,no_so) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                (date_now,
                                i[0],
                                i[1],
                                i[2],
                                i[3],
                                i[4],
                                i[5],
                                i[6],
                                i[7],
                                i[8],
                                i[9],
                                i[10],
                                i[11],
                                i[12],
                                i[13],
                                i[14],
                                i[15],
                                i[16],
                                i[17],
                                i[18],
                                i[19],
                                wh_name,
                                berat,
                                waiting_minute,
                                wb1_minute,
                                wh_minute,
                                wb2_minute,
                                cc_minute,
                                company,
                                category,
                                no_so)
                            )
                            file.commit()
                        # perbahrui location start
                        try:
                            c.execute("MERGE INTO tb_temporaryRFID_start AS target USING (VALUES (?, ?, ?, ?)) AS source (rfid, no_polisi, location, time_detected) ON target.no_polisi = source.no_polisi WHEN MATCHED THEN  UPDATE SET target.rfid = source.rfid, target.location = source.location, target.time_detected = source.time_detected WHEN NOT MATCHED THEN  INSERT (rfid, no_polisi, location, time_detected) VALUES (source.rfid, source.no_polisi, source.location, source.time_detected);",
                                (rfid_last,
                                no_polisi_last,
                                location_last_temp,
                                time_detected_last)
                            )
                            file.commit()
                        except:
                            pass
                    # getout false -> true | wb2 true -> false
                    c.execute("UPDATE tb_queuing_flow SET weight_bridge2=?, getout=? WHERE no_polisi = ?",
                        ("False",
                        "True",
                        no_polisi_last)
                    )
                    file.commit()
                    # blocking suara calling selama xx
                    c.execute("SELECT block FROM tb_tappingPointReg where functions = ?",("Unpairing/Unregistration",))
                    rec = c.fetchone()
                    block = rec[0]
                    file.commit()
                    # get time start dan last
                    c.execute("SELECT time_detected FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    time_start = rec[0]
                    file.commit()
                    c.execute("SELECT time_detected FROM tb_temporaryRFID_last where no_polisi = ?",(no_polisi_last,))
                    rec = c.fetchone()
                    time_last = rec[0]
                    file.commit()
                    # blocking suara calling selama block
                    format = '%H:%M:%S'
                    duration = datetime.strptime(time_last, format) - datetime.strptime(time_start, format)
                    if int(duration.seconds) > int(block):
                        c.execute("SELECT call_end FROM tb_calling where no_polisi = ?",(no_polisi_last,))
                        rec = c.fetchone()
                        call_end = rec[0]
                        file.commit()
                        if call_end == "True":
                            if activity == "LOADING":
                                # cek apakah covering atau tidak
                                c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
                                rec = c.fetchone()
                                flow_cc = rec[0]
                                file.commit()
                                # input to tb_playSound
                                c.execute("INSERT INTO tb_playSound (location, activity, ticket, covering_status) VALUES (?,?,?,?)",
                                    (location_last_temp,
                                    activity,
                                    ticket,
                                    flow_cc)
                                )
                                file.commit()
                                # Update tb_calling
                                c.execute("UPDATE tb_calling SET call_end=? WHERE no_polisi = ?",
                                    ("False",
                                    no_polisi_last)
                                )
                                file.commit()
                            else:
                                # cek apakah covering atau tidak
                                c.execute("SELECT covering FROM tb_warehouseFlow where packaging = ? AND activity = ?",(packaging, activity))
                                rec = c.fetchone()
                                flow_cc = rec[0]
                                file.commit()
                                # input to tb_playSound
                                c.execute("INSERT INTO tb_playSound (location, activity, ticket, covering_status) VALUES (?,?,?,?)",
                                    (location_last_temp,
                                    activity,
                                    ticket,
                                    flow_cc)
                                )
                                file.commit()
                                # Update tb_calling
                                c.execute("UPDATE tb_calling SET call_end=? WHERE no_polisi = ?",
                                    ("False",
                                    no_polisi_last)
                                )
                                file.commit()
                            time.sleep(5)
                            try:
                                # hapus semua data dengan plat no yg getout di tb_temporaryRFID
                                c.execute("DELETE FROM tb_temporaryRFID where no_polisi = ?",(no_polisi_last,))
                                file.commit()
                            except:
                                pass
                            try:
                                # hapus yg di tb_calling
                                c.execute("DELETE FROM tb_calling where no_polisi = ?",(no_polisi_last,))
                                file.commit()
                            except:
                                pass
                            try:
                                # hapus yg di tb_overSLA
                                c.execute("DELETE FROM tb_overSLA where no_polisi = ?",(no_polisi_last,))
                                file.commit()
                            except:
                                pass
                            try:
                                # hapus yg di tb_queuing_status
                                c.execute("DELETE FROM tb_queuing_status where no_polisi = ?",(no_polisi_last,))
                                file.commit()
                            except:
                                pass
                            try:
                                # hapus yg di tb_queuing_status
                                c.execute("DELETE FROM tb_queuing_flow where no_polisi = ?",(no_polisi_last,))
                                file.commit()
                            except:
                                pass
                            try:
                                # hapus yg di tb_temporaryRFID_start
                                c.execute("DELETE FROM tb_temporaryRFID_start where no_polisi = ?",(no_polisi_last,))
                                file.commit()
                            except:
                                pass
                            try:
                                # hapus yg di tb_calling_time
                                c.execute("DELETE FROM tb_calling_time where no_polisi = ?",(no_polisi_last,))
                                file.commit()
                            except:
                                pass

    def follow_up_inbox(self):
        selectedItem = self.alert_table.selection()[0]
        date_alert = self.alert_table.item(selectedItem)['values'][0]
        time_alert = self.alert_table.item(selectedItem)['values'][1]
        notes = self.notes_textbox.get("0.0", "end")
        # update status to FOLLOW UP di tb_alert
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("UPDATE tb_alert SET status=?, notes=? WHERE date = ? AND time = ?",
            "FOLLOW UP",
            notes,
            date_alert,
            time_alert
        )
        file.commit()
        file.close()
        #
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # data alert inbox
        for record in self.alert_table.get_children():
            self.alert_table.delete(record)
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_alert")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.alert_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=("evenrow",))
            else:
                self.alert_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=("oddrow",))
            count+=1
        file.commit()
        file.close()

    def done_inbox(self):
        selectedItem = self.alert_table.selection()[0]
        date_alert = self.alert_table.item(selectedItem)['values'][0]
        time_alert = self.alert_table.item(selectedItem)['values'][1]
        category = self.alert_table.item(selectedItem)['values'][3]
        message = self.alert_table.item(selectedItem)['values'][4]
        notes = self.alert_table.item(selectedItem)['values'][5]
        no_polisi = self.alert_table.item(selectedItem)['values'][6]
        arrival = self.alert_table.item(selectedItem)['values'][7]
        solving_time = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
        solving_date = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")

        format = '%H:%M:%S'
        solving_total_time = datetime.strptime(solving_time, format) - datetime.strptime(time_alert, format)

        # insert to tb_history
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("INSERT INTO tb_history (date,time,status,category,message,operator,warehouse,notes,no_polisi,arrival,solving_time,solving_date,solving_total_time) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
            date_alert,
            time_alert,
            "DONE",
            category,
            message,
            "",
            wh_name,
            notes,
            no_polisi,
            arrival,
            solving_time,
            solving_date,
            str(solving_total_time)
        )
        file.commit()
        file.close()
        # delete from tb_history
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("DELETE from tb_alert WHERE date = ? AND time = ?",
            date_alert,
            time_alert
        )
        file.commit()
        file.close()
        #
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # data alert inbox
        for record in self.alert_table.get_children():
            self.alert_table.delete(record)
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_alert")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.alert_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=("evenrow",))
            else:
                self.alert_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=("oddrow",))
            count+=1
        file.commit()
        file.close()

    def clear_alert(self):
        # get all data di tb_alert
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("SELECT * FROM tb_alert")
        rec = c.fetchall()
        file.commit()
        file.close()
        # insert all data to tb_history with status DONE
        for i in rec:
            date_alert = i[0]
            time_alert = i[1]
            status = "DONE"
            category = i[3]
            message = i[4]
            operator = ""
            notes = i[5]
            no_polisi = i[6]
            arrival = i[7]

            solving_time = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
            solving_date = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")

            format = '%H:%M:%S'
            solving_total_time = datetime.strptime(solving_time, format) - datetime.strptime(time_alert, format)

            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("INSERT INTO tb_history (date,time,status,category,message,operator,warehouse,notes,no_polisi,arrival,solving_time,solving_date,solving_total_time) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
            date_alert,
            time_alert,
            "DONE",
            category,
            message,
            operator,
            wh_name,
            notes,
            no_polisi,
            arrival,
            solving_time,
            solving_date,
            str(solving_total_time)
        )
            file.commit()
            file.close()
        # delete all data tb alert
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute("DELETE from tb_alert")
        file.commit()
        file.close()
        # refresh tree alert
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # data alert inbox
        for record in self.alert_table.get_children():
            self.alert_table.delete(record)
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_alert")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.alert_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=("evenrow",))
            else:
                self.alert_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=("oddrow",))
            count+=1
        file.commit()
        file.close()

    # untuk stop thread
    def stop(self):
        stop_event.set()
        try:
            thread.join()
        except:
            pass
        stop_event2.set()
        try:
            thread2.join()
        except:
            pass
        self.destroy()

    def thread1(self, stop_event): # (tb_alert & alert dashboard, alert notif, tb_waitinglist, pairing status tabel, tabel dashboard, cek waktu untuk monitor sla warehouse, tabel Displan data, tabel loading list, sound tapping)
        try:
            global count_alert, tupple_waitinglist, tupple_queuing_status, jumlah_inbox, tupple_dashboard, tupple_displan_data, tupple_loadinglist, tupple_soundPlay
            while not stop_event.is_set():
                # tb_alert & alert dashboard
                try:
                    #check if new data exist (tb_alert)
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute('SELECT COUNT(*) FROM tb_alert')
                    count_alert_new = c.fetchone()[0]
                    file.commit()
                    file.close()
                    if count_alert != count_alert_new:
                        try:
                            # real all data in all tab and show in each tab
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()

                            # data alert inbox
                            for record in self.alert_table.get_children():
                                self.alert_table.delete(record)
                            # get data from database
                            count = 0
                            c.execute("SELECT * FROM tb_alert")
                            rec = c.fetchall()
                            # delete semua di self.dashboard_alert
                            for widget in self.dashboard_alert.winfo_children():
                                widget.destroy()
                            # input to tabel GUI
                            for i in rec:
                                if count % 2 == 0:
                                    self.alert_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=("evenrow",))
                                else:
                                    self.alert_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7]), tags=("oddrow",))
                                count+=1
                                # add alert to dasboard
                                if i[3] == "OVER SLA":
                                    new_alert_label = customtkinter.CTkLabel(self.dashboard_alert, image=self.sla_alert_image, text=i[6]+"-"+"SLA", compound="top")
                                    new_alert_label.pack(side="left", padx=10)
                                    new_alert_label.bind("<Button-1>", self.goto_inbox)  # Bind left mouse button click event
                                    new_dict_alert_dashboard = {i[6]:new_alert_label}
                                    dict_alert_dashboard.update(new_dict_alert_dashboard)
                                elif i[3] == "REGISTRATION":
                                    new_alert_label = customtkinter.CTkLabel(self.dashboard_alert, image=self.registration_alert_image, text=i[6]+"-"+"REG", compound="top")
                                    new_alert_label.pack(side="left", padx=10)
                                    new_alert_label.bind("<Button-1>", self.goto_inbox)  # Bind left mouse button click event
                                    new_dict_alert_dashboard = {i[6]:new_alert_label}
                                    dict_alert_dashboard.update(new_dict_alert_dashboard)

                            file.commit()
                            file.close()
                            count_alert = count_alert_new
                        except:
                            pass
                except:
                    pass

                # alert notif
                if jumlah_inbox != count_alert_new:
                    if count_alert_new > jumlah_inbox:
                        playsound('./sound/inbox.mp3')
                    jumlah_inbox = count_alert_new
                    
                    if jumlah_inbox > 0:
                        text_color = "#ba0404"
                        bell_icon = ""
                        # bell_icon = "🔔"  # Add bell icon if jumlah_inbox is greater than 0
                    else:
                        text_color = "black"
                        bell_icon = ""  # No bell icon if jumlah_inbox is less than 0

                    self.report_button.configure(text="Report\t\t\t{} Inbox! {}".format(jumlah_inbox, bell_icon), text_color=text_color)
                
                # tb_waitinglist
                try:
                    #check if new data exist (tb_waitinglist)
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute('SELECT * FROM tb_waitinglist')
                    tupple_waitinglist_new = c.fetchall()
                    file.commit()
                    file.close()
                    if tupple_waitinglist != tupple_waitinglist_new:
                        try:
                            # real all data in all tab and show in each tab
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()

                            # delete all data first from table
                            for record in self.manualCall_table.get_children():
                                self.manualCall_table.delete(record)
                            # get data from database
                            count = 0
                            c.execute("SELECT * FROM tb_waitinglist")
                            rec = c.fetchall()
                            # input to tabel GUI
                            for i in rec:
                                if count % 2 == 0:
                                    self.manualCall_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("evenrow",))
                                else:
                                    self.manualCall_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9]), tags=("oddrow",))
                                count+=1

                            file.commit()
                            file.close()
                            tupple_waitinglist = tupple_waitinglist_new
                        except:
                            pass
                except:
                    pass

                # pairing status tabel
                try:
                    # pairing status tabel (UPDATE JIKA DATA UPDATE)
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT * FROM tb_queuing_status")
                    tupple_queuing_status_new = c.fetchall()
                    file.commit()
                    file.close()
                    if tupple_queuing_status != tupple_queuing_status_new:
                        try:
                            # read all data in all tab and show in each tab
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()

                            # data alert inbox
                            for record in self.rfidPairing_table.get_children():
                                self.rfidPairing_table.delete(record)
                            # get data from database
                            count = 0
                            c.execute("SELECT * FROM tb_queuing_status")
                            rec = c.fetchall()
                            # input to tabel GUI
                            for i in rec:
                                if count % 2 == 0:
                                    self.rfidPairing_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18], i[19], i[20]), tags=("evenrow",))
                                else:
                                    self.rfidPairing_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12], i[13], i[14], i[15], i[16], i[17], i[18], i[19], i[20]), tags=("oddrow",))
                                count+=1

                            file.commit()
                            file.close()
                            tupple_queuing_status = tupple_queuing_status_new
                        except:
                            pass
                except:
                    pass
                
                # tabel dashboard
                try:
                    # pairing status tabel (UPDATE JIKA DATA UPDATE)
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT no_polisi, location FROM tb_temporaryRFID_start")
                    tupple_dashboard_new = c.fetchall()
                    file.commit()
                    file.close()
                    if tupple_dashboard != tupple_dashboard_new:
                        # clean data in tree
                        for record in self.gatein_table.get_children():
                            self.gatein_table.delete(record)
                        for record in self.wb1_table.get_children():
                            self.wb1_table.delete(record)
                        for record in self.wh_table.get_children():
                            self.wh_table.delete(record)
                        for record in self.wb2_table.get_children():
                            self.wb2_table.delete(record)
                        for record in self.cc_table.get_children():
                            self.cc_table.delete(record)
                        for record in self.gateout_table.get_children():
                            self.gateout_table.delete(record)
                        # input in each tree
                        for i in tupple_dashboard_new:
                            if i[1] == "Pairing/Registration":
                                self.gatein_table.insert(parent="",index="end", text="", values=(i[0]), tags=("oddrow","evenrow"))
                            if i[1] == "Weight Bridge 1":
                                self.wb1_table.insert(parent="",index="end", text="", values=(i[0]), tags=("oddrow","evenrow"))
                            if i[1] == "Warehouse":
                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                c = file.cursor()
                                c.execute("SELECT wh_id FROM tb_queuing_status WHERE no_polisi = ?",(i[0],))
                                wh_id = c.fetchone()[0]
                                file.commit()
                                file.close()
                                self.wh_table.insert(parent="",index="end", text="", values=(i[0],wh_id), tags=("oddrow","evenrow"))
                            if i[1] == "Weight Bridge 2":
                                self.wb2_table.insert(parent="",index="end", text="", values=(i[0]), tags=("oddrow","evenrow"))
                            if i[1] == "Cargo Covering":
                                self.cc_table.insert(parent="",index="end", text="", values=(i[0]), tags=("oddrow","evenrow"))
                            if i[1] == "Unpairing/Unregistration":
                                self.gateout_table.insert(parent="",index="end", text="", values=(i[0]), tags=("oddrow","evenrow"))
                        tupple_dashboard = tupple_dashboard_new
                except:
                    pass

                # cek waktu untuk monitor sla warehouse
                try:
                    # get no plat yang ada di warehouse
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT no_polisi FROM tb_temporaryRFID_start where location = ?",("Warehouse",))
                    rec = c.fetchall()
                    file.commit()
                    file.close()
                    tanggal_now = time.strftime("%Y")+"-"+time.strftime("%m")+"-"+time.strftime("%d")
                    for i in rec:
                        platNo = i[0]
                        # get product & packaging
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()
                        c.execute("SELECT product, packaging, wh_in, arrival, activity FROM tb_queuing_status where no_polisi = ?",(platNo,))
                        rec1 = c.fetchone()
                        product_thread = rec1[0]
                        packaging_thread = rec1[1]
                        wh_in_time = rec1[2]
                        arrival = rec1[3]
                        activity_thread = rec1[4]
                        file.commit()
                        file.close()
                        # get berat
                        if activity_thread == "LOADING":
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()
                            c.execute("SELECT berat FROM tb_loadingList where no_polisi = ? and status = ? and product = ?",(platNo, "Activate", product_thread))
                            rec2 = c.fetchone()
                            berat = rec2[0]
                            file.commit()
                            file.close()
                        else:
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()
                            c.execute("SELECT berat FROM tb_unloading where no_polisi = ? and status = ? and product = ?",(platNo, "Activate", product_thread))
                            rec2 = c.fetchone()
                            berat = rec2[0]
                            file.commit()
                            file.close()
                        # sla berdasarkan berat
                        try:
                            # get weight min max
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()
                            c.execute("SELECT weight_min, weight_max FROM tb_warehouseSla where packaging = ? AND product_grp = ?",(packaging_thread, product_thread))
                            rec3 = c.fetchall()
                            file.commit()
                            file.close()
                            # get sla based on berat
                            for j in rec3:
                                weight_min = j[0]
                                weight_max = j[1]
                                # get SLA wh where packaging, product, berat in range weight_min and in weight_max di tb_warehouseSla
                                if int(berat) >= int(weight_min) and int(berat) <= int(weight_max):
                                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                    c = file.cursor()
                                    c.execute("SELECT wh_sla FROM tb_warehouseSla where packaging = ? AND product_grp = ? AND weight_min = ? AND weight_max = ?",(packaging_thread, product_thread, weight_min, weight_max))
                                    rec4 = c.fetchone()
                                    sla_wh = rec4[0]
                                    file.commit()
                                    file.close()
                        # sla default berdasarkan packaging
                        except:
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()
                            c.execute("SELECT sla FROM tb_packagingReg where packaging = ?",(packaging_thread,))
                            rec6 = c.fetchone()
                            sla_wh = rec6[0]
                            file.commit()
                            file.close()
                        sla_wh = int(sla_wh)*60
                        waktu_now = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
                        format = '%H:%M:%S'
                        duration = datetime.strptime(waktu_now, format) - datetime.strptime(wh_in_time, format)
                        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                        c = file.cursor()
                        c.execute("SELECT over_wh FROM tb_overSLA where no_polisi = ?",(platNo,))
                        rec5 = c.fetchone()
                        over_wh = rec5[0]
                        file.commit()
                        file.close()
                        if over_wh == "False":
                            if int(duration.seconds) > sla_wh:
                                # masuk alert
                                date = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
                                status = "WAITING"
                                category = "OVER SLA"
                                message = "Truck {} over SLA at Warehouse".format(platNo)
                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                c = file.cursor()
                                c.execute("INSERT INTO tb_alert (date,time,status,category,message,notes,no_polisi,arrival) VALUES (?,?,?,?,?,?,?,?)",
                                    date,
                                    waktu_now,
                                    status,
                                    category,
                                    message,
                                    "",
                                    platNo,
                                    arrival
                                )
                                file.commit()
                                file.close()
                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                c = file.cursor()
                                c.execute("UPDATE tb_overSLA SET over_wh=? WHERE no_polisi = ?",
                                    "True",
                                    platNo
                                )
                                file.commit()
                                file.close()
                except:
                    pass

                # tabel Displan data
                try:
                    # Displan Data tabel (UPDATE JIKA DATA UPDATE)
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT * FROM tb_displanData")
                    tupple_displan_data_new = c.fetchall()
                    file.commit()
                    file.close()
                    if tupple_displan_data != tupple_displan_data_new:
                        try:
                            # real all data in all tab and show in each tab
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()

                            # delete all data first from table
                            for record in self.displanData_table.get_children():
                                self.displanData_table.delete(record)
                            # get data from database
                            count = 0
                            c.execute("SELECT * FROM tb_displanData")
                            rec = c.fetchall()
                            # input to tabel GUI
                            for i in rec:
                                if count % 2 == 0:
                                    self.displanData_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("evenrow",))
                                else:
                                    self.displanData_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("oddrow",))
                                count+=1

                            file.commit()
                            file.close()
                            tupple_displan_data = tupple_displan_data_new
                        except:
                            pass
                except:
                    pass

                # tabel loading list
                try:
                    # loading list tabel (UPDATE JIKA DATA UPDATE)
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT * FROM tb_loadingList")
                    tupple_loadinglist_new = c.fetchall()
                    file.commit()
                    file.close()
                    if tupple_loadinglist != tupple_loadinglist_new:
                        try:
                            # real all data in all tab and show in each tab
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()

                            # delete all data first from table
                            for record in self.loadingList_table.get_children():
                                self.loadingList_table.delete(record)
                            # get data from database
                            count = 0
                            c.execute("SELECT * FROM tb_loadingList")
                            rec = c.fetchall()
                            # input to tabel GUI
                            urutan = 1
                            for i in rec:
                                if count % 2 == 0:
                                    self.loadingList_table.insert(parent="",index="end", iid=count, text="", values=(i[13], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12]), tags=("evenrow",))
                                else:
                                    self.loadingList_table.insert(parent="",index="end", iid=count, text="", values=(i[13], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12]), tags=("oddrow",))
                                count+=1
                                urutan+=1

                            file.commit()
                            file.close()
                            tupple_loadinglist = tupple_loadinglist_new
                        except:
                            pass
                except:
                    pass

                # sound tapping
                try:
                    # loading list tabel (UPDATE JIKA DATA UPDATE)
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT * FROM tb_playSound")
                    tupple_soundPlay_new = c.fetchall()
                    file.commit()
                    file.close()
                    if tupple_soundPlay != tupple_soundPlay_new:
                        try:
                            # get data where min(oid)
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()
                            c.execute("SELECT * FROM tb_playSound where oid = (SELECT MIN(oid) FROM tb_playSound)")
                            rec = c.fetchone()
                            location = rec[0]
                            activity = rec[1]
                            ticket = rec[2]
                            covering_status = rec[3]
                            file.commit()
                            file.close()

                            try:
                                # get gudang berapa
                                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                                c = file.cursor()
                                c.execute("SELECT wh_id FROM tb_queuing_status where ticket = ?",(ticket,))
                                rec = c.fetchone()
                                gudang = int(rec[0])
                                file.commit()
                                file.close()
                            except:
                                pass

                            # change false -> true in tb_inPlaySound
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()
                            c.execute("SELECT play FROM tb_inPlaySound")
                            if c.fetchone()[0] == 0:
                                c.execute("UPDATE tb_inPlaySound SET play = 1")
                            file.commit()
                            file.close()

                            # play sound
                            sound_directory = "./sound/taping/"
                            if location == "Pairing/Registration":
                                sound_directory = sound_directory+"gate_in/"+ticket[0]
                            if location == "Weight Bridge 1":
                                sound_directory = sound_directory+"wb1/"+ticket[0]
                            if location == "Warehouse":
                                sound_directory = sound_directory+"wh/"+activity+"/"+ticket[0]
                            if location == "Cargo Covering":
                                sound_directory = sound_directory+"cc/"+ticket[0]
                            if location == "Weight Bridge 2":
                                if covering_status == "YES":
                                    sound_directory = sound_directory+"wb2/covering/"+ticket[0]
                                else:
                                    sound_directory = sound_directory+"wb2/not_covering/"+ticket[0]
                            if location == "Unpairing/Unregistration":
                                sound_directory = sound_directory+"gate_out/"+activity+"/"+ticket[0]

                            playsound(sound_directory+"/"+"{}.mp3".format(ticket))

                            # playsound gudang berapa
                            if location == "Weight Bridge 1":
                                playsound("./sound/number/{}.mp3".format(gudang))

                            # delete data where min(oid)
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()
                            c.execute("DELETE FROM tb_playSound where oid = (SELECT MIN(oid) FROM tb_playSound)")
                            file.commit()
                            file.close()

                            # change true -> false in tb_inPlaySound
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()
                            c.execute("SELECT play FROM tb_inPlaySound")
                            if c.fetchone()[0] == 1:
                                c.execute("UPDATE tb_inPlaySound SET play = 0")
                            file.commit()
                            file.close()

                            tupple_soundPlay = tupple_soundPlay_new
                        except:
                            pass
                except:
                    pass

        except:
            pass

    # Function to save data to Excel
    def save_to_excel(self,treeview,columns):
        # Get data from Treview widget
        def get_treview_data(treeview):
            data = []
            for item in treeview.get_children():
                values = treeview.item(item)['values']
                data.append(values)
            return data

        # Create file dialog to select file location
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        
        if file_path:
            try:
                # Create Excel writer
                writer = pd.ExcelWriter(file_path)

                # Get data from Treview
                data = get_treview_data(treeview)
                
                # Convert data to DataFrame
                df = pd.DataFrame(data, columns=columns)
                
                # Write data to Excel
                df.to_excel(writer, index=False)
                
                # Close writer
                writer.save()
                writer.close()
                messagebox.showinfo("Success", "File Berhasil Tersimpan")
            except:
                messagebox.showerror("Error", "File Gagal Tersimpan")

    def export_excel_total_time(self):
        self.save_to_excel(self.report_viewer_total_time_table,self.report_viewer_total_time_table["columns"])

    def export_excel_alert_note_history(self):
        self.save_to_excel(self.alert_note_history_tabel,self.alert_note_history_tabel["columns"])

    def export_excel_transaction_time(self):
        self.save_to_excel(self.report_viewer_transaction_time_table,self.report_viewer_transaction_time_table["columns"])

    def export_excel_peak_hour(self):
        self.save_to_excel(self.report_viewer_peak_hour_table,self.report_viewer_peak_hour_table["columns"])

    def export_excel_total_queuing(self):
        self.save_to_excel(self.report_viewer_total_queuing_table,self.report_viewer_total_queuing_table["columns"])

    def export_excel_incident_report(self):
        self.save_to_excel(self.report_viewer_incident_report_table,self.report_viewer_incident_report_table["columns"])

    def export_excel_ticket_history(self):
        self.save_to_excel(self.report_viewer_ticket_history_table,self.report_viewer_ticket_history_table["columns"])

    def export_excel_sla(self):
        self.save_to_excel(self.report_viewer_sla_table,self.report_viewer_sla_table["columns"])

    def export_excel_auditlog(self):
        self.save_to_excel(self.report_viewer_auditlog_table,self.report_viewer_auditlog_table["columns"])

    def zoom_in_dashboard(self):
        global size_diagram, size_arrow, size_container
        size_diagram = size_diagram + 10
        size_arrow = size_arrow + 10
        size_container = size_container + 10
        self.gatein_image.configure(size=(size_diagram, size_diagram))
        self.wb1_image.configure(size=(size_diagram, size_diagram))
        self.wh_image.configure(size=(size_diagram, size_diagram))
        self.wb2_image.configure(size=(size_diagram, size_diagram))
        self.cc_image.configure(size=(size_diagram, size_diagram))
        self.gateout_image.configure(size=(size_diagram, size_diagram))
        self.arrow_image.configure(size=(size_arrow, size_arrow))
        self.container_image.configure(size=(size_container, size_container))

    def zoom_out_dashboard(self):
        global size_diagram, size_arrow, size_container
        size_diagram = size_diagram - 10
        size_arrow = size_arrow - 10
        size_container = size_container - 10
        self.gatein_image.configure(size=(size_diagram, size_diagram))
        self.wb1_image.configure(size=(size_diagram, size_diagram))
        self.wh_image.configure(size=(size_diagram, size_diagram))
        self.wb2_image.configure(size=(size_diagram, size_diagram))
        self.cc_image.configure(size=(size_diagram, size_diagram))
        self.gateout_image.configure(size=(size_diagram, size_diagram))
        self.arrow_image.configure(size=(size_arrow, size_arrow))
        self.container_image.configure(size=(size_container, size_container))

    def reset_queuing(self):
        Question = messagebox.askquestion("Reset Queuing", "Anda yakin ingin mengatur ulang sistem antrian?\nNB: Ini tidak menghapus data master dan report,\nini hanya akan menghapus data dalam sistem antrian")
        if Question == "yes":
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # tb_calling
            c.execute("DELETE from tb_calling")
            c.execute("DBCC CHECKIDENT ('tb_calling', RESEED, 0)")
            # tb_calling_time
            c.execute("DELETE from tb_calling_time")
            c.execute("DBCC CHECKIDENT ('tb_calling_time', RESEED, 0)")
            # tb_noantrian
            c.execute("DELETE from tb_noantrian")
            c.execute("DBCC CHECKIDENT ('tb_noantrian', RESEED, 0)")
            # reset tb_noantrian
            prefix = ["A","B","C","D","Z"]
            c.execute("select count(*) from tb_noantrian")
            for i in prefix:
                c.execute("INSERT INTO tb_noantrian (prefix,no_antrian) VALUES (?,?)",
                    i,
                    "0"
                )
            # tb_queuing_flow
            c.execute("DELETE from tb_queuing_flow")
            c.execute("DBCC CHECKIDENT ('tb_queuing_flow', RESEED, 0)")
            # tb_queuing_status
            c.execute("DELETE from tb_queuing_status")
            c.execute("DBCC CHECKIDENT ('tb_queuing_status', RESEED, 0)")
            # tb_temporaryRFID
            c.execute("DELETE from tb_temporaryRFID")
            c.execute("DBCC CHECKIDENT ('tb_temporaryRFID', RESEED, 0)")
            # tb_temporaryRFID_start
            c.execute("DELETE from tb_temporaryRFID_start")
            c.execute("DBCC CHECKIDENT ('tb_temporaryRFID_start', RESEED, 0)")
            # tb_temporaryRFID_last
            c.execute("DELETE from tb_temporaryRFID_last")
            c.execute("DBCC CHECKIDENT ('tb_temporaryRFID_last', RESEED, 0)")
            # tb_waitinglist
            c.execute("DELETE from tb_waitinglist")
            c.execute("DBCC CHECKIDENT ('tb_waitinglist', RESEED, 0)")
            # tb_overSLA
            c.execute("DELETE from tb_overSLA")
            c.execute("DBCC CHECKIDENT ('tb_overSLA', RESEED, 0)")
            # tb_ticket_skip
            c.execute("DELETE from tb_ticket_skip")
            c.execute("DBCC CHECKIDENT ('tb_ticket_skip', RESEED, 0)")
            file.commit()
            file.close()
            messagebox.showinfo("Reset Queuing", "Berhasil mengatur ulang sistem antrian")
        else:
            pass

    def reset_tappingSound(self):
        Question = messagebox.askquestion("Reset Tapping Sound", "Anda yakin ingin mengatur ulang tapping sound?")
        if Question == "yes":
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # tb_playSound
            c.execute("DELETE from tb_playSound")
            c.execute("DBCC CHECKIDENT ('tb_playSound', RESEED, 0)")
            file.commit()
            file.close()
            # change if true -> false in tb_inPlaySound
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            c.execute("SELECT play FROM tb_inPlaySound")
            if c.fetchone()[0] == 1:
                c.execute("UPDATE tb_inPlaySound SET play = 0")
            file.commit()
            file.close()
            messagebox.showinfo("Reset Tapping Sound", "Berhasil mengatur ulang tapping sound")
        else:
            pass

    def delete_report(self):
        Question = messagebox.askquestion("Delete Report", "Anda yakin ingin menghapus semua data report?\n(Termasuk data alert history)")
        if Question == "yes":
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # tb_calling
            c.execute("DELETE from tb_queuing_history")
            # tb_calling_time
            c.execute("DELETE from tb_history")
            # tb_noantrian
            c.execute("DELETE from tb_ticket_history")
            file.commit()
            file.close()
            messagebox.showinfo("Delete Report", "Berhasil menghapus seluruh data report")
        else:
            pass

    def upload_excel_vetting(self):
        Question = messagebox.askquestion("Import Excel", "Anda akan menghapus seluruh data.\nData yang terhapus akan diperbarui dengan data pada file excel.\nAnda Yakin?")
        if Question == "yes":
            filename = filedialog.askopenfilename(
                initialdir="C:/",
                title="Open A File",
                filetypes=(("xlsx files", "*.xlsx"), ("All Files", "*.*"))
            )

            if filename:
                try:
                    filename = r"{}".format(filename)
                    df = pd.read_excel(filename, dtype={'RFID 1': str, 'RFID 2': str})
                    expected_header = ['COMPANY', 'NO POLISI', 'REGISTRATION DATE', 'EXPIRED DATE', 'REMARK', 'STATUS', 'RFID 1', 'RFID 2', 'CATEGORY']
                    if not df.columns.tolist() == expected_header:
                        messagebox.showerror("Error", "Format file salah / File Excel Salah.")
                        return
                except ValueError:
                    messagebox.showerror("Error", "Invalid file format. Please select an Excel file.")
                    return
                except FileNotFoundError:
                    messagebox.showerror("Error", "File not found. Please select a valid file.")
                    return

            try:
                file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                c = file.cursor()
                # Delete all data in the table
                c.execute("DELETE FROM tb_vetting")
                file.commit()
                # Import data to the table
                for row in df.itertuples(index=False):
                    values = [str(val) if val else '' for val in row]
                    reg_date = str(values[2]).split(' ')[0] if values[2] else ''
                    expired_date = str(values[3]).split(' ')[0] if values[3] else ''

                    c.execute("""
                        INSERT INTO tb_vetting (company, no_polisi, reg_date, expired_date, remark, status, rfid1, rfid2, category)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, values[0], values[1], reg_date, expired_date, values[4], values[5], values[6], values[7], values[8])

                file.commit()
                file.close()
                messagebox.showinfo("Success", "Data tersimpan")
            except:
                messagebox.showerror("Error", "Gagal Menyimpan Data")

            # show data in tabel
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # data tab vetting
            for record in self.vetting_table.get_children():
                self.vetting_table.delete(record)
            count = 0
            c.execute("SELECT * FROM tb_vetting")
            rec = c.fetchall()
            urutan = 1
            for i in rec:
                if count % 2 == 0:
                    self.vetting_table.insert(parent="",index="end", iid=count, text="", values=(i[9], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
                else:
                    self.vetting_table.insert(parent="",index="end", iid=count, text="", values=(i[9], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
                count+=1
                urutan+=1
        else:
            pass

    def upload_excel_distribution(self):
        filename = filedialog.askopenfilename(
            initialdir="C:/",
            title="Open A File",
            filetypes=(("xlsx files", "*.xlsx"), ("All Files", "*.*"))
        )

        if filename:
            try:
                filename = r"{}".format(filename)
                df = pd.read_excel(filename)
                expected_header = ['COMPANY', 'NO POLISI', 'PACKAGING', 'PRODUCT', 'DATE', 'TIME (HH:MM)', 'STATUS', 'QUOTA', 'NO SO', 'BERAT']
                if not df.columns.tolist() == expected_header:
                    messagebox.showerror("Error", "Format file salah / File Excel Salah.")
                    return
            except ValueError:
                messagebox.showerror("Error", "Invalid file format. Please select an Excel file.")
                return
            except FileNotFoundError:
                messagebox.showerror("Error", "File not found. Please select a valid file.")
                return

        try:
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # Import data to the table
            for row in df.itertuples(index=False):
                values = [str(val) if val else '' for val in row]
                date_work = str(values[4]).split(' ')[0] if values[2] else ''

                c.execute("""
                    INSERT INTO tb_loadingList (company, no_polisi, packaging, product, date, time, remark, status, quota, no_so, berat, wh_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, values[0], values[1], values[2], values[3], date_work, (values[5])[0:5], "LOADING", values[6], values[7], values[8], values[9], "Auto")

            file.commit()
            file.close()
            messagebox.showinfo("Success", "Data tersimpan")
        except:
            messagebox.showerror("Error", "Gagal Menyimpan Data")

        # show data in tabel
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # data tab distribution -> loading list
        for record in self.loadingList_table.get_children():
            self.loadingList_table.delete(record)
        count = 0
        c.execute("SELECT * FROM tb_loadingList")
        rec = c.fetchall()
        # input to tabel GUI
        urutan = 1
        for i in rec:
            if count % 2 == 0:
                self.loadingList_table.insert(parent="",index="end", iid=count, text="", values=(i[13], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12]), tags=("evenrow",))
            else:
                self.loadingList_table.insert(parent="",index="end", iid=count, text="", values=(i[13], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11], i[12]), tags=("oddrow",))
            count+=1
            urutan+=1

    def upload_excel_unloading(self):
        filename = filedialog.askopenfilename(
            initialdir="C:/",
            title="Open A File",
            filetypes=(("xlsx files", "*.xlsx"), ("All Files", "*.*"))
        )

        if filename:
            try:
                filename = r"{}".format(filename)
                df = pd.read_excel(filename)
                expected_header = ['COMPANY', 'NO POLISI', 'PACKAGING', 'PRODUCT', 'DATE', 'TIME (HH:MM)', 'STATUS', 'QUOTA', 'BERAT', 'NO SO']
                if not df.columns.tolist() == expected_header:
                    messagebox.showerror("Error", "Format file salah / File Excel Salah.")
                    return
            except ValueError:
                messagebox.showerror("Error", "Invalid file format. Please select an Excel file.")
                return
            except FileNotFoundError:
                messagebox.showerror("Error", "File not found. Please select a valid file.")
                return

        try:
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            # Import data to the table
            for row in df.itertuples(index=False):
                values = [str(val) if val else '' for val in row]
                date_work = str(values[4]).split(' ')[0] if values[2] else ''

                c.execute("""
                    INSERT INTO tb_unloading (company, no_polisi, packaging, product, date, time, remark, status, quota, berat, wh_id, no_so)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, values[0], values[1], values[2], values[3], date_work, (values[5])[0:5], "UNLOADING", values[6], int(values[7]), values[8], "Auto", values[9])

            file.commit()
            file.close()
            messagebox.showinfo("Success", "Data tersimpan")
        except:
            messagebox.showerror("Error", "Gagal Menyimpan Data")

        # show data in tabel
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # data tab unloading
        for record in self.unloading_table.get_children():
            self.unloading_table.delete(record)
        count = 0
        c.execute("SELECT * FROM tb_unloading")
        rec = c.fetchall()
        # input to tabel GUI
        urutan = 1
        for i in rec:
            if count % 2 == 0:
                self.unloading_table.insert(parent="",index="end", iid=count, text="", values=(i[12], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("evenrow",))
            else:
                self.unloading_table.insert(parent="",index="end", iid=count, text="", values=(i[12], urutan, i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10], i[11]), tags=("oddrow",))
            count+=1
            urutan+=1

    def get_displan_api(self, stop_event):
        global stop_event2, thread2, save_api_entry, start_schedule, loadingDate, date_range

        file = sq.connect("./dataBase/dataBase.db")
        c = file.cursor()
        c.execute("SELECT * FROM tb_api_setting")
        rec = c.fetchall()
        
        if len(rec) != 0 :
            for i in rec:
                location = i[0]
                token = i[1]
                domain = i[2]
                url = i[3]+"api/DeliverySlip/GetLSForQueueing"
                schedule = i[4]
        else:
            token = ""
            location = ""
            domain = ""
            schedule = ""
            url = ""

        try:
            token = self.entry_token_api.get()
            location = self.entry_warehouse_id.get()
            domain = self.entry_domain_api.get()
            schedule = self.entry_download_displan_api.get()
            url = self.entry_url_api.get()+"api/DeliverySlip/GetLSForQueueing"
        except:
            pass

        try:
            self.entry_token_api.configure(state="disabled")
            self.entry_warehouse_id.configure(state="disabled")
            self.entry_domain_api.configure(state="disabled")
            self.entry_url_api.configure(state="disabled")
            self.entry_download_displan_api.configure(state="disabled")
            self.get_displan_api_button.configure(fg_color="dark red", text="Stop Get Data", hover_color="red", command=self.stop_get_api_data)
            if save_api_entry == True:
                # delete data in tb_api_setting
                try:
                    file = sq.connect("./dataBase/dataBase.db")
                    c = file.cursor()
                    c.execute("DELETE from tb_api_setting")
                    file.commit()
                    file.close()
                except:
                    pass
                # save data connection to sqlite table
                try:
                    file = sq.connect("./dataBase/dataBase.db")
                    c = file.cursor()
                    c.execute("INSERT INTO tb_api_setting VALUEs(:wh_id,:token,:domain,:url,:schedule_api)",{
                        "wh_id": location,
                        "token": token,
                        "domain": domain,
                        "url": self.entry_url_api.get(),
                        "schedule_api" : schedule
                    })
                    file.commit()
                    file.close()
                except:
                    pass
                save_api_entry = False
            schedule = int(schedule)*60
            # start thread get_api_data
            while not stop_event.is_set():
                # loadingDate = datetime.now().date()
                params = {
                    'Token': token,
                    'Location': location,
                    'Domain': domain,
                    'LoadingDate': loadingDate
                }
                # get data from API
                response = requests.get(url, params=params)
                json_string = response.content.decode('utf-8')
                data = json.loads(json_string)
                # tabel displan data
                time_now = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
                format = '%H:%M:%S'
                duration = datetime.strptime(time_now, format) - datetime.strptime(start_schedule, format)
                # print(duration)
                if int(duration.seconds) > schedule:
                    # cek dlu data tb_displan_data_container ada yang baru atau tidak (no_so)
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    c.execute("SELECT no_so FROM tb_displan_data_container")
                    rec_no_so = c.fetchall()
                    file.commit()
                    file.close()
                    # eliminasi data json yang sudah ada di tb_displan_data_container
                    for row in data['data']:
                        for no_so in rec_no_so:
                            if row['tr_sod_nbr'] == no_so[0]:
                                # hapus data tersebut dr json
                                data['data'] = [d for d in data["data"] if d.get("tr_sod_nbr") != row['tr_sod_nbr']]
                    # input data json yang tidak tereliminasi to tb_displan_data_container
                    if len(data['data']) > 0:
                        for row in data['data']:
                            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                            c = file.cursor()
                            c.execute("INSERT INTO tb_displan_data_container (company, truck, packaging, product, date, time, pkg_remark, remark, no_so, tr_part, satuan, berat) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                row['tr_shipto_name'],
                                row['tr_vehicle'],
                                row['tr_packaging'],
                                row['tr_part_name'],
                                row['tr_due_date'][0:10],
                                row['tr_due_date'][11:],
                                row['tr_packaging_rmks'],
                                row['tr_remarks'],
                                row['tr_sod_nbr'],
                                row['tr_part'],
                                row['tr_um'],
                                row['tr_qty_loc'],
                            )
                            file.commit()
                            file.close()

                    # Retrieve data from tb_displan_data_container based on date range
                    file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
                    c = file.cursor()
                    query = "SELECT company, truck, packaging, product, date, time, pkg_remark, remark, no_so, tr_part, satuan, berat " \
                            "FROM tb_displan_data_container " \
                            "WHERE date IN ({})".format(','.join(['?' for _ in date_range]))
                    
                    c.execute(query, [str(date) for date in date_range])
                    result_set = c.fetchall()

                    #cek apakah sama dengan data di tb_displanData sebelumnya (if beda)
                    c.execute("SELECT company, truck, packaging, product, date, time, pkg_remark, remark, no_so, tr_part, satuan, berat FROM tb_displanData")
                    rec = c.fetchall()
                    
                    if rec != result_set:
                        #> hapus data di tb_displanData
                        c.execute("DELETE FROM tb_displanData")
                        #> inpurt data result_set to tb_displanData
                        for row in result_set:
                            c.execute("INSERT INTO tb_displanData (company, truck, packaging, product, date, time, pkg_remark, remark, no_so, tr_part, satuan, berat) "
                                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                    row.company,
                                    row.truck,
                                    row.packaging,
                                    row.product,
                                    row.date,
                                    row.time[0:5],
                                    row.pkg_remark,
                                    row.remark,
                                    row.no_so,
                                    row.tr_part,
                                    row.satuan,
                                    row.berat)

                        # input to tabel loading list
                        c.execute("SELECT no_so FROM tb_displanData")
                        rec_displanData = c.fetchall()
                        c.execute("SELECT no_so FROM tb_loadingList")
                        rec_loadingList = c.fetchall()

                        for no_so_displanData in rec_displanData:
                            if no_so_displanData not in rec_loadingList:
                                # get data in tb_displanData where no_so_displanData
                                c.execute("SELECT company, truck, packaging, product, date, time, pkg_remark, remark, tr_part, satuan, berat FROM tb_displanData where no_so = ?",(no_so_displanData[0],))
                                rec = c.fetchone()

                                # Check if packaging needs special treatment
                                special_packaging = None
                                c.execute("SELECT remark, packaging FROM tb_packagingReg")
                                reg_records = c.fetchall()
                                for reg in reg_records:
                                    remark_values = reg.remark.split(',')
                                    if rec.packaging in remark_values:
                                        special_packaging = reg.packaging
                                        break

                                # Insert the no_so value into tb_loadingList
                                if special_packaging:
                                    c.execute("INSERT INTO tb_loadingList (company, no_polisi, packaging, product, date, time, remark, status, quota, no_so, satuan, berat, wh_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                        rec.company,      # company
                                        rec.truck,        # no_polisi
                                        special_packaging,
                                        rec.product,
                                        rec.date,
                                        str(rec.time)[0:5],
                                        "LOADING",
                                        'Activate',        # status (replace with appropriate value)
                                        '1',               # quota (replace with appropriate value)
                                        no_so_displanData[0],
                                        rec.satuan,
                                        int(rec.berat),
                                        'Auto'
                                        )
                                else:
                                    c.execute("INSERT INTO tb_loadingList (company, no_polisi, packaging, product, date, time, remark, status, quota, no_so, satuan, berat, wh_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                        rec.company,      # company
                                        rec.truck,        # no_polisi
                                        rec.packaging,
                                        rec.product,
                                        rec.date,
                                        str(rec.time)[0:5],
                                        "LOADING",
                                        'Activate',        # status (replace with appropriate value)
                                        '1',               # quota (replace with appropriate value)
                                        no_so_displanData[0],
                                        rec.satuan,
                                        int(rec.berat),
                                        'Auto'
                                        )

                    # Commit and close the database connection
                    file.commit()
                    file.close()

                    start_schedule = time_now
                    
        except:
            save_api_entry = True
            thread2 = threading.Thread(target=self.get_displan_api, args=(stop_event2,))
            self.get_displan_api_button.configure(fg_color="dark blue", text="GET API DATA", hover_color="blue", command=thread2.start)
            error_message = "Unable to connect to the API\nReconect API in menu System Setting > Conection Displan API"
            messagebox.showerror("Error", error_message)
            self.entry_token_api.configure(state="normal")
            self.entry_warehouse_id.configure(state="normal")
            self.entry_domain_api.configure(state="normal")
            self.entry_url_api.configure(state="normal")
            self.entry_download_displan_api.configure(state="normal")

    def stop_get_api_data(self):
        global stop_event2, thread2, save_api_entry
        stop_event2.set()
        thread2.join()
        save_api_entry = True
        stop_event2 = threading.Event()
        thread2 = threading.Thread(target=self.get_displan_api, args=(stop_event2,))
        self.get_displan_api_button.configure(fg_color="dark blue", text="GET API DATA", hover_color="blue", command=thread2.start)
        self.entry_token_api.configure(state="normal")
        self.entry_warehouse_id.configure(state="normal")
        self.entry_domain_api.configure(state="normal")
        self.entry_url_api.configure(state="normal")
        self.entry_download_displan_api.configure(state="normal")

    def change_password(self):
        username_akun = self.entry_username_user.get()
        password_akun = self.entry_password_user.get()
        new_password = self.entry_new_password.get()
        retype_new_password = self.entry_retype_new_password.get()

        # Check if username and password exist in the database
        db = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        cursor = db.cursor()
        cursor.execute("SELECT * FROM tb_userRegistration WHERE username = ? AND password = ?", (username_akun, password_akun))
        user = cursor.fetchone()

        if not user:
            messagebox.showerror("Error", "Username/password salah.")
            return

        if new_password == retype_new_password:
            # Update password in the database
            cursor.execute("UPDATE tb_userRegistration SET password = ? WHERE username = ?", (new_password, username_akun))
            db.commit()
            db.close()
            messagebox.showinfo("Success", "Password berhasil diubah.")
            # Clear the entry fields
            self.entry_username_user.delete(0, 'end')
            self.entry_password_user.delete(0, 'end')
            self.entry_new_password.delete(0, 'end')
            self.entry_retype_new_password.delete(0, 'end')
        else:
            messagebox.showerror("Error", "Retype new password salah.")

    def login(self, event=None):
        global username_log
        username_log = self.entry_username_log.get()
        password_log = self.entry_password_log.get()
        
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()

        # check apakah username sudah ada atau belum
        c.execute("SELECT * FROM tb_userRegistration where username = ? AND password = ?",(username_log,password_log))
        rec = c.fetchone()
        file.commit()
        file.close()

        if rec != None:

            self.dashboard_button.configure(state="Normal")
            self.report_button.configure(state="Normal")
            self.Qsettings_button.configure(state="Normal")
            self.truck_data_button.configure(state="Normal")
            self.settings_button.configure(state="Normal")
            self.user_button.configure(state="Normal")

            alertNoteMenu = rec[4]
            reportViewerMenu = rec[5]
            packagingMenu = rec[6]
            productMenu = rec[7]
            tappingPointMenu = rec[8]
            warehouseFlowMenu = rec[9]
            rfidManualCallMenu = rec[10]
            vettingMenu = rec[11]
            unloadingMenu = rec[12]
            distributionMenu = rec[13]
            databaseSettingMenu = rec[14]
            displanApiMenu = rec[15]
            otherSettingMenu = rec[16]
            userProfileMenu = rec[17]
            userRegistrationMenu = rec[18]

            if alertNoteMenu == False:
                self.tabview_report.delete("Alert Note")
            if reportViewerMenu == False:
                self.tabview_report.delete("Report Viewer")
            if alertNoteMenu == False and reportViewerMenu == False:
                self.report_button.configure(state="disabled")

            if packagingMenu == False:
                self.tabview_qsettings.delete("Packaging")
            if productMenu == False:
                self.tabview_qsettings.delete("Product")
            if tappingPointMenu == False:
                self.tabview_qsettings.delete("Tapping Point")
            if warehouseFlowMenu == False:
                self.tabview_qsettings.delete("Warehouse Flow")
            if rfidManualCallMenu == False:
                self.tabview_qsettings.delete("Queuing Manual Call")
            if packagingMenu == False and productMenu == False and tappingPointMenu == False and warehouseFlowMenu == False and rfidManualCallMenu == False:
                self.Qsettings_button.configure(state="disabled")

            if vettingMenu == False:
                self.tabview_truckData.delete("Vetting")
            if unloadingMenu == False:
                self.tabview_truckData.delete("Unloading")
            if distributionMenu == False:
                self.tabview_truckData.delete("Distribution")
            if vettingMenu == False and unloadingMenu == False and distributionMenu == False:
                self.truck_data_button.configure(state="disabled")

            if databaseSettingMenu == False:
                self.tabview_settings.delete("Database Setting")
            if displanApiMenu == False:
                self.tabview_settings.delete("Connection Displan API")
            if otherSettingMenu == False:
                self.tabview_settings.delete("Other Settings")
            if databaseSettingMenu == False and displanApiMenu == False and otherSettingMenu == False:
                self.settings_button.configure(state="disabled")

            if userProfileMenu == False:
                self.tabview_user_settings.delete("User Profile")
            if userRegistrationMenu == False:
                self.tabview_user_settings.delete("User Registration")
            if userProfileMenu == False and userRegistrationMenu == False:
                self.user_button.configure(state="disabled")

            self.login_button.configure(fg_color="gray", state="disabled")
            self.logout_button.configure(fg_color="dark red", state="normal")
            self.select_frame_by_name("dashboard")

            self.location.configure(text=f"🚩 Lokasi : {wh_name}")
            self.user.configure(text=f"👷🏼 User : {username_log}")

            # save log
            file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
            c = file.cursor()
            date_log = time.strftime("%d")+"/"+time.strftime("%m")+"/"+time.strftime("%Y")
            time_log = time.strftime("%H")+":"+time.strftime("%M")+":"+time.strftime("%S")
            transaction_log = "AUDIT LOG"
            status = "DONE"
            remark = f"Login application: {username_log}"
            # write data to database
            c.execute("INSERT INTO tb_log_report (warehouse,date_log,time_log,transaction_log,operator,status,remark) VALUES (?,?,?,?,?,?,?)",
                wh_name,
                date_log,
                time_log,
                transaction_log,
                username_log,
                status,
                remark
            )
            file.commit()
            file.close()

            messagebox.showinfo("Login Berhasil", f"Selamat Datang {username_log}")

            self.entry_username_log.configure(state="readonly")
            self.entry_password_log.configure(state="readonly")
        else:
            messagebox.showerror("Error", "Password atau Username salah")

    def logout(self):
        self.dashboard_button.configure(state="disabled")
        self.report_button.configure(state="disabled")
        self.Qsettings_button.configure(state="disabled")
        self.truck_data_button.configure(state="disabled")
        self.settings_button.configure(state="disabled")
        self.user_button.configure(state="disabled")

        self.user.configure(text="Silahkan Login")

        self.login_button.configure(fg_color="dark green", state="normal")
        self.logout_button.configure(fg_color="gray", state="disabled")

        self.entry_username_log.delete(0,tk.END)
        self.entry_password_log.delete(0,tk.END)

        messagebox.showinfo("Info", "Logout Berhasil")
        self.stop()

    def set_call_rules(self):
        delay_repeat = self.entry_call_rules_repeat.get()
        skip_call_time = self.entry_call_rules_skip.get()
        delete_call = self.entry_call_rules_delete.get()

        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()

        c.execute("SELECT COUNT(*) FROM tb_call_rules")
        count = c.fetchone()[0]

        if count > 0:
            c.execute("UPDATE tb_call_rules SET delay_repeat=?, skip_call_time=?, delete_call=?",
                    delay_repeat, skip_call_time, delete_call)
        else:
            c.execute("INSERT INTO tb_call_rules (delay_repeat, skip_call_time, delete_call) VALUES (?,?,?)",
                    delay_repeat, skip_call_time, delete_call)

        file.commit()
        file.close()

        messagebox.showinfo("Success", "Pengaturan Disimpan")
      
    def set_schedule_loadingListDB(self):
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()

        # Clear the table
        c.execute('DELETE FROM tb_loadingScheduleSetting')

        # Insert the new values
        c.execute('''
            INSERT INTO tb_loadingScheduleSetting (
                allow_h1,
                allow_h2,
                allow_h3,
                allow_h_minus_1,
                allow_h_minus_2,
                allow_h_minus_3
            ) VALUES (?, ?, ?, ?, ?, ?)
        ''', (
            int(self.checkbox_loading_schedule_hplus1.get()),
            int(self.checkbox_loading_schedule_hplus2.get()),
            int(self.checkbox_loading_schedule_hplus3.get()),
            int(self.checkbox_loading_schedule_hmin1.get()),
            int(self.checkbox_loading_schedule_hmin2.get()),
            int(self.checkbox_loading_schedule_hmin3.get())
        ))

        # Commit the changes and close the connection
        c.commit()
        c.close()
    
    def update_date_range(self):
        global date_range, loadingDate
        date_range = [loadingDate]
        if (self.checkbox_loading_schedule_hplus1.get())=="1":
            date_range.append(loadingDate + timedelta(days=1))
        if (self.checkbox_loading_schedule_hplus2.get())=="1":
            date_range.append(loadingDate + timedelta(days=2))
        if (self.checkbox_loading_schedule_hplus3.get())=="1":
            date_range.append(loadingDate + timedelta(days=3))
        if (self.checkbox_loading_schedule_hmin1.get())=="1":
            date_range.append(loadingDate + timedelta(days=-1))
        if (self.checkbox_loading_schedule_hmin2.get())=="1":
            date_range.append(loadingDate + timedelta(days=-2))
        if (self.checkbox_loading_schedule_hmin3.get())=="1":
            date_range.append(loadingDate + timedelta(days=-3))
        self.set_schedule_loadingListDB()

    def set_schedule_loadingList(self):
        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        c.execute('SELECT TOP 1 * FROM tb_loadingScheduleSetting')
        rec = c.fetchone()
        if rec != None:
            if rec[0] == True:
                self.checkbox_loading_schedule_hplus1.select()
                date_range.append(loadingDate + timedelta(days=1))
            if rec[1] == True:
                self.checkbox_loading_schedule_hplus2.select()
                date_range.append(loadingDate + timedelta(days=2))
            if rec[2] == True:
                self.checkbox_loading_schedule_hplus3.select()
                date_range.append(loadingDate + timedelta(days=3))
            if rec[3] == True:
                self.checkbox_loading_schedule_hmin1.select()
                date_range.append(loadingDate + timedelta(days=-1))
            if rec[4] == True:
                self.checkbox_loading_schedule_hmin2.select()
                date_range.append(loadingDate + timedelta(days=-2))
            if rec[5] == True:
                self.checkbox_loading_schedule_hmin3.select()
                date_range.append(loadingDate + timedelta(days=-3))
        c.commit()
        c.close()

    def toggle_navigation(self):
        if self.navigation_frame.winfo_children():
            # If there are no children in the navigation frame, add back all buttons and labels
            for widget in [self.navigation_frame_label, self.dashboard_button, self.report_button,
                        self.Qsettings_button, self.truck_data_button, self.settings_button,
                        self.user_button]:
                widget.grid()

        else:
            # If all buttons and labels are present, remove them
            for widget in [self.navigation_frame_label, self.dashboard_button, self.report_button,
                        self.Qsettings_button, self.truck_data_button, self.settings_button,
                        self.user_button]:
                widget.grid_remove()

        # Toggle the visibility of the navigation frame
        if self.navigation_frame.winfo_viewable():
            self.navigation_frame.grid_remove()  # Hide the navigation frame
        else:
            self.navigation_frame.grid()  # Show the navigation frame

    def cancel_modify_vetting(self):
        query = self.entry_policeNo.get()
        selections = []
        for child in self.vetting_table.get_children():
            if query in self.vetting_table.item(child)['values']:
                selections.append(child)
        self.vetting_table.selection_remove(selections)

        self.entry_company.delete(0,tk.END)
        self.entry_policeNo.delete(0,tk.END)
        self.entry_remark.delete(0,tk.END)
        self.entry_rfid1.delete(0,tk.END)
        self.entry_rfid2.delete(0,tk.END)

    def cancel_modify_unloading(self):
        query = self.entry_policeNo_unloading.get()
        selections = []
        for child in self.unloading_table.get_children():
            if query in self.unloading_table.item(child)['values']:
                selections.append(child)
        self.unloading_table.selection_remove(selections)

        self.entry_company_unloading.delete(0,tk.END)
        self.entry_policeNo_unloading.delete(0,tk.END)
        self.entry_quota_unloading.delete(0,tk.END)
        self.entry_berat_unloading.delete(0,tk.END)
        self.optionmenu_noSo.delete(0,tk.END)

    def cancel_modify_distribution(self):
        query = self.entry_policeNo_distribution.get()
        selections = []
        for child in self.loadingList_table.get_children():
            if query in self.loadingList_table.item(child)['values']:
                selections.append(child)
        self.loadingList_table.selection_remove(selections)
        for child in self.displanData_table.get_children():
            if query in self.displanData_table.item(child)['values']:
                selections.append(child)
        self.displanData_table.selection_remove(selections)

        self.entry_company_distribution.delete(0,tk.END)
        self.entry_policeNo_distribution.delete(0,tk.END)
        self.entry_quota_distribution.delete(0,tk.END)
        self.entry_berat_distribution.delete(0,tk.END)
        self.entry_remark_distribution.delete(0,tk.END)
        self.entry_noSO_distribution.delete(0,tk.END)
        self.entry_pkgRemark_distribution.delete(0,tk.END)
        self.entry_trPart_distribution.delete(0,tk.END)

    def cancel_packaging(self):
        query = self.entry_packaging.get()
        selections = []
        for child in self.packaging_table.get_children():
            if query in self.packaging_table.item(child)['values']:
                selections.append(child)
        self.packaging_table.selection_remove(selections)
        
        # data warehouse sla setting
        for record in self.warehouse_slaSetting_table.get_children():
            self.warehouse_slaSetting_table.delete(record)

        file = pyodbc.connect(f'DRIVER={driver_db};SERVER={server_db};DATABASE={dbname};UID={username};PWD={password}')
        c = file.cursor()
        # get data from database
        count = 0
        c.execute("SELECT * FROM tb_warehouseSla")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.warehouse_slaSetting_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10]), tags=("evenrow",))
            else:
                self.warehouse_slaSetting_table.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8], i[9], i[10]), tags=("oddrow",))
            count+=1

        self.entry_packaging.delete(0,tk.END)
        self.entry_prefix.delete(0,tk.END)
        self.entry_remark_packaging.delete(0,tk.END)
        self.entry_sla_packaging.delete(0,tk.END)

    def cancel_warehouse_slaSetting(self):
        query = self.optionmenu_packaging_sla.get()
        selections = []
        for child in self.warehouse_slaSetting_table.get_children():
            if query in self.warehouse_slaSetting_table.item(child)['values']:
                selections.append(child)
        self.warehouse_slaSetting_table.selection_remove(selections)

        self.optionmenu_packaging_sla.set("Pilih Packaging")
        self.optionmenu_product_sla.set("Pilih Product")
        self.optionmenu_vehicle_sla.set("ALL")
        self.entry_weightMin.delete(0,tk.END)
        self.entry_weightMax.delete(0,tk.END)
        self.entry_waitingTime.delete(0,tk.END)
        self.entry_wBridge1Min.delete(0,tk.END)
        self.entry_warehouseMin.delete(0,tk.END)
        self.entry_wBridge2Min.delete(0,tk.END)
        self.entry_coveringMin.delete(0,tk.END)

    def cancel_product_registration(self):
        query = self.entry_product_group.get()
        selections = []
        for child in self.product_table.get_children():
            if query in self.product_table.item(child)['values']:
                selections.append(child)
        self.product_table.selection_remove(selections)

        self.entry_product_group.delete(0,tk.END)
        self.entry_product_group_remark.delete(0,tk.END)

    def cancel_tapping_point(self):
        query = self.entry_remark_tappingpoint.get()
        selections = []
        for child in self.tappingPoint_table.get_children():
            if query in self.tappingPoint_table.item(child)['values']:
                selections.append(child)
        self.tappingPoint_table.selection_remove(selections)

        self.entry_machineId.delete(0,tk.END)
        self.entry_tapId.delete(0,tk.END)
        self.entry_aliasId.delete(0,tk.END)

        self.checkbox_packaging_tappingPoint_all.deselect()
        self.checkbox_packaging_tappingPoint_zak.deselect()
        self.checkbox_packaging_tappingPoint_curah.deselect()
        self.checkbox_packaging_tappingPoint_jumbo.deselect()
        self.checkbox_packaging_tappingPoint_pallet.deselect()

        self.entry_max_tappingPoint.delete(0,tk.END)
        self.entry_remark_tappingpoint.delete(0,tk.END)
        self.entry_tapBlocking_tappingpoint.delete(0,tk.END)

        self.optionmenu_activity_tappingPoint.set("Pilih Activity")
        self.optionmenu_function_tappingPoint.set("Pilih Function")
        self.optionmenu_status_tappingPoint.set("Activate")

    def cancel_warehouse_flow(self):
        selections = []
        for child in self.warehouse_flow_table.get_children():
            selections.append(child)
        self.warehouse_flow_table.selection_remove(selections)

        self.optionmenu_packaging_warehouseflow.set("Pilih Packaging")
        self.optionmenu_activity_warehouseflow.set("Pilih Activity")
        self.checkbox_tappingPoint_warehouseflow_registration.deselect()
        self.checkbox_tappingPoint_warehouseflow_weightbridge2.deselect()
        self.checkbox_tappingPoint_warehouseflow_weightbridge1.deselect()
        self.checkbox_tappingPoint_warehouseflow_covering.deselect()
        self.checkbox_tappingPoint_warehouseflow_warehouse.deselect()
        self.checkbox_tappingPoint_warehouseflow_unregistration.deselect()

    def cancel_pairing_status(self):
        selections = []
        for child in self.rfidPairing_table.get_children():
            selections.append(child)
        self.rfidPairing_table.selection_remove(selections)

        self.optionmenu_warehouse_rfidPairing.set("")
        self.entry_noTiket_rfidPairing.delete(0,tk.END)

    def cancel_waiting_list(self):
        selections = []
        for child in self.manualCall_table.get_children():
            selections.append(child)
        self.manualCall_table.selection_remove(selections)

        self.entry_truckNo_manualCall.delete(0,tk.END)

    def goto_inbox(self, event):
        self.select_frame_by_name("report")
        self.tabview_report.set("Alert Note")
        self.tabview_alert_note.set("Inbox")

    def show_version(self, event):
        messagebox.showinfo("Information", "Copyright © 2024 AKR Technology Development.\nSoftware version 1.0")

class MyDateEntry(DateEntry):
    def __init__(self, master=None, align='left', **kw):
        DateEntry.__init__(self, master, **kw)
        self.align = align

    def drop_down(self):
        """Display or withdraw the drop-down calendar depending on its current state."""
        if self._calendar.winfo_ismapped():
            self._top_cal.withdraw()
        else:
            self._validate_date()
            date = self.parse_date(self.get())
            h = self._top_cal.winfo_reqheight()
            w = self._top_cal.winfo_reqwidth()
            x_max = self.winfo_screenwidth()
            y_max = self.winfo_screenheight()
            # default: left-aligned drop-down below the entry
            x = self.winfo_rootx()
            y = self.winfo_rooty() + self.winfo_height()
            if x + w > x_max:  # the drop-down goes out of the screen
                # right-align the drop-down
                x += self.winfo_width() - w
            if y + h > y_max:  # the drop-down goes out of the screen
                # bottom-align the drop-down
                y -= self.winfo_height() + h
            if self.winfo_toplevel().attributes('-topmost'):
                self._top_cal.attributes('-topmost', True)
            else:
                self._top_cal.attributes('-topmost', False)
            self._top_cal.geometry('+%i+%i' % (x, y))
            self._top_cal.deiconify()
            self._calendar.focus_set()
            self._calendar.selection_set(date)

if __name__ == "__main__":
    app = App()
    app.mainloop()


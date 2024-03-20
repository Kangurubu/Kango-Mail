import getpass
from imap_tools import MailBox, A

import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from functools import partial
from PIL import Image, ImageTk
from ttkbootstrap.tableview import Tableview

import pytz
from datetime import datetime

from tkhtmlview import HTMLText, RenderHTML

import webbrowser 

import time

class KangoMail (ttk.Frame):
	def __init__(self, root):
		self.root = root
		self.root.place_window_center()
		self.mail = None 
		self.dt = None
		self.timezone = pytz.timezone('Etc/GMT-3')  # UTC+3 Türkiye 
		self.welcomeFrame = ttk.Frame(root)
		self.display_login_screen()

	def __del__(self):
		if self.mail is not None:
			print("Done.")
			self.mail.logout()
			self.mail = None

	def connection(self, usernameVal, passVal):
		IMAPURL = ""  # Write the IMAP server you want to connect.
		if IMAPURL == "":
			print("Please edit the IMAPURL variable with the IMAP server you want to connect to at the function connection.")
			exit()

		print("Welcome.")
		try:
			self.mail =  MailBox(IMAPURL).login(usernameVal.get(), passVal.get(), 'INBOX')
			print("Connected!")
			self.display_mails()
		except Exception as e:
			print("Error! ", e)
			errorMessage = ttk.dialogs.dialogs.Messagebox.ok(f"There was an error with the login process. {e}", title="ERROR", alert=True, parent=self.root)
			return 1

	def display_login_screen(self):
		s = ttk.Style()
		s.configure('TFrame', background='#69676E')

		self.root.columnconfigure(0, weight=3)
		self.root.columnconfigure(1, weight=3)

		self.welcomeFrame.place(x=0, y=0)

		usernameLabel = ttk.Label(self.welcomeFrame,text="Username: ", bootstyle="inverse-secondary", font=("Times New Roman", 16))
		usernameLabel.grid(row=1, column=0, padx=15, pady=15)
		
		usernameVal = ttk.StringVar()
		usernameEntry = ttk.Entry(self.welcomeFrame, textvariable=usernameVal)
		usernameEntry.grid(row=1, column=1, columnspan=2, padx=15, sticky="ew")

		passwordLabel = ttk.Label(self.welcomeFrame,text="Password: ", bootstyle="inverse-secondary", font=("Times New Roman", 16))
		passwordLabel.grid(row=2, column=0, padx=15, pady=15)
		
		passVal = ttk.StringVar()
		passwordEntry = ttk.Entry(self.welcomeFrame, show="*", textvariable=passVal)
		passwordEntry.grid(row=2, column=1, columnspan=2 , padx=15,sticky="ew")
		
		logo = Image.open("assets/logo.png")
		tmp = ImageTk.PhotoImage(logo)
		logo = tk.Label(self.welcomeFrame,image=tmp)

		logo.image = tmp
		logo.grid(row=0, column=0, columnspan=3)

		lbtnStyle = ttk.Style()
		lbtnStyle.configure("primary.Outline.TButton", font=("Times New Roman", 24))
		loginButton = ttk.Button(self.welcomeFrame, text="Log In", style="primary.Outline.TButton", command=partial(self.connection,usernameVal,passVal))
		loginButton.grid(row=3,column=0, columnspan=3, pady=20)
		
	def display_mail_body(self):
		selectedRow = self.dt.get_rows(selected=True)[0]
		selectedUid = selectedRow.values[0]

		# Write the mail content to an html file and then open it using the browser. Might not be the best solution.
		# -- mark_seen is set as True. If desired, one can change it to False.
		for msg in self.mail.fetch(A(uid=selectedUid), mark_seen=True):
			with open("openedMail.html", "w", encoding="utf-8") as f:
				f.write(msg.html) 
			webbrowser.open("openedMail.html")  

		# Since the mail is seen now, remove the stripe.
		self.dt.get_rows(selected=True)[0]
		selectedRow.configure(None, tags="")

	def display_mails(self):
		self.welcomeFrame.place_forget()
		colors = self.root.style.colors
		
		style = ttk.Style()
		style.configure("Treeview.Heading", font=("Times New Roman", 16))

		rowdata = []
		unseen = []
		

		# Only the last 100 emails are fetched. Can be changed if desired.
		messages = self.mail.fetch(limit=100, mark_seen=False, reverse=True, bulk=True) 
		for msg in messages:
			# Change time to Türkiye time, remove the unnecessary +0300 at the end. Convert date to a more suitable form.
			formattedDate = msg.date.astimezone(self.timezone).replace(tzinfo=None).strftime('%d.%m.%Y %H:%M:%S') 
			rowdata.append(( msg.uid,msg.from_,msg.subject,formattedDate ))
			if not "\\Seen" in msg.flags:
				unseen.append(msg.uid)
			
		coldata = [
			{"text": "mailid", "stretch":False},
		    {"text": "From", "stretch": True},
		    {"text": "Subject", "stretch": True},
		    {"text": "Date", "stretch": True},
		]

		colors = self.root.style.colors
		self.dt = Tableview( 
		    master=self.root,
		    coldata=coldata,
		    rowdata=rowdata,
		    paginated=True,
		    searchable=True,
		    bootstyle=PRIMARY,
		    autofit=True,
		    stripecolor=(colors.info, colors.light),
		)

		for row in self.dt.get_rows():
			if row.values[0] in unseen:
				row.configure(None, tags="striped")

		self.dt.tablecolumns[0].hide()
		self.dt.pack(fill=Y, expand=YES, side=LEFT)

		openMailButton = ttk.Button(self.root, text="Open", style="primary.Outline.TButton", command=self.display_mail_body)
		openMailButton.pack(side=RIGHT, padx=40)

		self.root.geometry("") # This will resize the window to the size of the new widgets in mail table.
		root.after(60000, self.check_for_new_mails) 

	def check_for_new_mails(self):
		if self.mail:
			self.mail.idle.start()
			print("In the function.")
			responses = self.mail.idle.poll(timeout=60)
			self.mail.idle.stop()

			if responses:
				for msg in self.mail.fetch(A(seen=False), mark_seen=False):
					print(msg.uid, msg.from_, msg.subject)
					formattedDate = msg.date.astimezone(self.timezone).replace(tzinfo=None).strftime('%d.%m.%Y %H:%M:%S') 
					row = self.dt.insert_row(index=0, values=[msg.uid, msg.from_, msg.subject ,formattedDate])
					row.configure(None, tags="striped")
					self.dt.load_table_data()
			else:
				print("no updates")

			self.root.after(60000, self.check_for_new_mails) 
		

if __name__ == '__main__':

	root = ttk.Window(title="KangoMail", themename="pulse",iconphoto="assets/icon.png", size=[360,410])
	KangoMail(root)
	root.mainloop()

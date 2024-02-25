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

def connection(usernameVal,passVal):
	global mail
	IMAPURL = ""  # Write the IMAP server you want to connect.
	if IMAPURL == "":
		print("Please edit the IMAPURL variable with the IMAP server you want to connect to at the function connection.")
		exit()

	print("Welcome.")
	try:
		mail =  MailBox(IMAPURL).login(usernameVal.get(), passVal.get(), 'INBOX')
		print("Connected!")
	except Exception as e:
		print("Error! ", e)
		# TODO Go to an error function.
		return 1

	mailTableScreen()
	
	
def mailBody(dt, mail):
	selectedUid = dt.get_rows(selected=True)[0].values[0]
	for msg in mail.fetch(A(uid=selectedUid)):
		# mailTextWindow = ttk.Toplevel(title=f"{msg.subject}", size=[1080,720])
		
		# html_label = HTMLText(mailTextWindow, html=f'{msg.html}')
		# html_label.pack(fill="both", expand=True)
		# html_label.fit_height()

		# mailTextWindow.place_window_center()
		# mailBodyText = ttk.Label(mailTextWindow, text=msg.text)
		# mailBodyText.pack(fill=BOTH, expand=YES)
		# mailTextWindow.mainloop()

		with open("openedMail.html", "w", encoding="utf-8") as f:
			f.write(msg.html) 
		webbrowser.open("openedMail.html")  
		

def mailTableScreen():
	welcomeFrame.place_forget()
	colors = root.style.colors
	
	style = ttk.Style()
	style.configure("Treeview.Heading", font=("Times New Roman", 16))

	rowdata=[]
	timezone = pytz.timezone('Etc/GMT-3')  # UTC+3 Türkiye 

	messages = mail.fetch(limit=100, mark_seen=False, reverse=True, bulk=True) 
	for msg in messages:
		# Change time to Türkiye time, remove the unnecessary +0300 at the end. Convert date to more suitable form.
		formattedDate = msg.date.astimezone(timezone).replace(tzinfo=None).strftime('%d.%m.%Y %H:%M:%S') 
		rowdata.append(( msg.uid,msg.from_,msg.subject,formattedDate ))
		
	coldata = [
		{"text": "mailid", "stretch":False},
	    {"text": "From", "stretch": True},
	    {"text": "Subject", "stretch": True},
	    {"text": "Date", "stretch": True},
	]

	dt = Tableview( 
	    master=root,
	    coldata=coldata,
	    rowdata=rowdata,
	    paginated=True,
	    searchable=True,
	    bootstyle=PRIMARY,
	    autofit=True
	)
	# Line below can be used to open emails, too.
	# dt.view.bind("<<TreeviewSelect>>", lambda event: mailBody(event, dt, mail))

	dt.tablecolumns[0].hide()
	dt.pack(fill=Y, expand=YES, side=LEFT,)

	openMailButton = ttk.Button(root, text="Open", style="primary.Outline.TButton", command=partial(mailBody, dt, mail))
	openMailButton.pack(side=RIGHT, padx=40)

	root.geometry("") # This will resize the window to the size of the new widgets in mail table.
	return dt

def loginScreen():
	s = ttk.Style()
	s.configure('TFrame', background='#69676E')

	root.columnconfigure(0, weight=3)
	root.columnconfigure(1, weight=3)

	welcomeFrame.place(x=0, y=0)

	usernameLabel = ttk.Label(welcomeFrame,text="Username: ", bootstyle="inverse-secondary", font=("Times New Roman", 16))
	usernameLabel.grid(row=1, column=0, padx=15, pady=15)
	
	usernameVal = ttk.StringVar()
	usernameEntry = ttk.Entry(welcomeFrame, textvariable=usernameVal)
	usernameEntry.grid(row=1, column=1, columnspan=2, padx=15, sticky="ew")

	passwordLabel = ttk.Label(welcomeFrame,text="Password: ", bootstyle="inverse-secondary", font=("Times New Roman", 16))
	passwordLabel.grid(row=2, column=0, padx=15, pady=15)
	
	passVal = ttk.StringVar()
	passwordEntry = ttk.Entry(welcomeFrame, show="*", textvariable=passVal)
	passwordEntry.grid(row=2, column=1, columnspan=2 , padx=15,sticky="ew")
	
	logo = Image.open("assets/logo.png")
	tmp = ImageTk.PhotoImage(logo)
	logo = tk.Label(welcomeFrame,image=tmp)

	logo.image = tmp
	logo.grid(row=0, column=0, columnspan=3)

	lbtnStyle = ttk.Style()
	lbtnStyle.configure("primary.Outline.TButton", font=("Times New Roman", 24))
	loginButton = ttk.Button(welcomeFrame, text="Log In", style="primary.Outline.TButton", command=partial(connection,usernameVal,passVal))
	loginButton.grid(row=3,column=0, columnspan=3, pady=20)
	

root = ttk.Window(title="KangoMail", themename="pulse",iconphoto="assets/icon.png", size=[360,410])
welcomeFrame = ttk.Frame(root)
mail = None 
def main():
	root.place_window_center()
	colors = root.style.colors

	loginScreen()
	root.mainloop()
	if mail is not None:
		print("Done.")
		mail.logout()

if __name__ == '__main__':
	main()


from tkinter import*
from tkinter import filedialog
from tkinter.ttk import *
from tkinter import messagebox
import comtypes.client
import webbrowser
import os
import time

filepath = ""
dirname = ""
filename = ""
target_out=""



window = Tk()


my_filetypes = [('docx files', '.docx'), ('doc files', '.doc'), ('text files', '.txt'),]

def openFile():
	filepath = filedialog.askopenfilename(parent=window , title="Please Select a File - " , filetypes = my_filetypes)
	if(filepath != ""):
		filepath = filepath.replace("/" , "\\")
		dirname = os.path.dirname(filepath)

		filename = filepath[len(dirname)+1:-5]+".pdf"

		target_out = dirname+"\\"+filename
		#----------------------------------------
		format_code=17
		file_input = os.path.abspath(filepath)
		file_output = os.path.abspath(target_out)

		global progress
		progress = Progressbar(window, orient = HORIZONTAL,length = 250, mode = 'indeterminate')

		def bar():
		    progress['value'] = 20
		    window.update_idletasks()
		    time.sleep(0.1)
		  
		    progress['value'] = 40
		    window.update_idletasks()
		    time.sleep(0.1)
		  
		    progress['value'] = 50
		    window.update_idletasks()
		    time.sleep(0.1)
		  
		    progress['value'] = 60
		    window.update_idletasks()
		    time.sleep(0.1)
		  
		    progress['value'] = 80
		    window.update_idletasks()
		    time.sleep(0.1)
		  
		    progress['value'] = 100
		    window.update_idletasks()
		    time.sleep(0.1)
		  
		    progress['value'] = 80
		    window.update_idletasks()
		    time.sleep(0.1)
		  
		    progress['value'] = 60
		    window.update_idletasks()
		    time.sleep(0.1)
		  
		    progress['value'] = 50
		    window.update_idletasks()
		    time.sleep(0.1)
		  
		    progress['value'] = 40
		    window.update_idletasks()
		    time.sleep(0.1)
		  
		    progress['value'] = 20
		    window.update_idletasks()
		    time.sleep(0.1)
		    progress['value'] = 0

		def terminate():
			window.destroy()

		def callback1(url):
			webbrowser.open_new_tab(url)

		def callback2(url):
			command = 'explorer.exe ' + url
			os.system(command)
			terminate()

		def prog_start():
			lh.destroy()
			btn.destroy()

			#--------------before info ok alert-------
			press_convert.destroy()
			global before_wait 
			before_wait = Label(window, text = "Press 'OK' To Continue")
			before_wait.config(font =("Comic sans ms", 24))			
			before_wait.pack(pady=20)
			#-----------------------------------------
			messagebox.showinfo("Your File is being Analyzed","Press 'OK' And Have a Sip of Coffee !!") 
			#------------after info ok alert------
			before_wait.destroy()
			progress.pack_forget()
			global wait_msg
			wait_msg = Label(window, text = "Your File is being 'INCARNATED' ...")
			wait_msg.config(font =("Comic sans ms", 24))			
			wait_msg.pack(pady=20)
			progress.pack(pady = 20)
			bar()
			#------------------

		def convert_btn():
			button.destroy()
			lh.destroy()
			global press_convert
			press_convert = Label(window, text = "Press 'Convert' to Proceed...")
			press_convert.config(font =("Comic sans ms", 24))			
			press_convert.pack(pady=20)

			global btn
			btn = Button(text="Convert" , command=lambda:[prog_start(),convert()])
			btn.pack()


		def convert():
			wait_msg.destroy()
			progress.destroy()
			btn.destroy()
			word_app = comtypes.client.CreateObject('Word.Application')
			word_file = word_app.Documents.Open(file_input)

			if (word_file.SaveAs(file_output,FileFormat=format_code)) == 0:

				ms = Label(window, text="Your PDF File Is Saved At",font=('Helvetica bold', 15))
				ms.pack(pady=5)

				link = Label(window, text=target_out+" (Open Folder)",font=('Helveticaunderline', 15))
				link.pack(pady=10)
				link.bind("<Button-1>", lambda e:callback2(dirname))


			word_file.Close()
			word_app.Quit()
			messagebox.showinfo("Yay!! Your File is Ready ","Your PDF File Is Ready ")
			btn.pack_forget()	
			lh.pack_forget()
			t = Label(window, text = "To Open Your File")
			t.config(font =("Comic sans ms", 24))
			t.pack(pady=10)
			bt = Button(text="Click Here" , command=lambda:[callback1("file:///"+target_out.replace("\\" , "/")),terminate()])
			bt.pack(pady=10)

		convert_btn()


  
# label for heading
global lh 
lh = Label(window, text = "Choose One From Here..")
lh.config(font =("Comic sans ms", 24))
 

# window.iconbitmap(r'C:\Users\USER\Desktop\Pro\my_icon.ico')

#minimum window size value
window.title("Welcome to PDF Converter - Convert Your Files into PDF Free!! ")
window.minsize(800, 200) 
#maximum window size value
window.maxsize(800, 200)
#----------------------------------------------------------------------
windowWidth = window.winfo_reqwidth()
windowHeight = window.winfo_reqheight()
 
# Gets both half the screen width/height and window width/height
positionRight = int(window.winfo_screenwidth()/3 - windowWidth/1.2)
positionDown = int(window.winfo_screenheight()/3 - windowHeight/3)
 
# Positions the window in the center of the page.
window.geometry("+{}+{}".format(positionRight, positionDown))
#----------------------------------------------------------------------
# Progress bar widget
button = Button(text="Select File",command=openFile)

lh.pack()
button.pack()

window.mainloop()
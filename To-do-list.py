from tkinter import *
import random
import tkinter.ttk as ttk
from  tkinter import messagebox
import pickle
import xlsxwriter
import time
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
import os
from fpdf import FPDF

def main():

	global tasks, completed_tasks
	### Create empty lists
	tasks = []
	completed_tasks = []


	try:
	#Unpickle saved To-Do-List IFF some shit is present in corresponding. PICKLE file
		pickle_in = open("tasks.pickle", "rb")
		a = open("completed_tasks.pickle", "rb")
		fuck = a.read()
		if fuck != b'':
			tasks = pickle.load(pickle_in)
	except FileNotFoundError:
		pass

	try:
		#Unpickle saved completed task list IFF some shit is present in corresponding PICKLE file
		pickle_in_completed_tasks = open("completed_tasks.pickle", "rb")
		a = open("completed_tasks.pickle", "rb")
		fuck = a.read()
		if fuck != b'':
			completed_tasks = pickle.load(pickle_in_completed_tasks)

	except FileNotFoundError:
		pass

	######################################### CREATE A WINDOW (root) ########################################
	root = Tk()
	root.configure(background="white", bd=20)
	root.title("To-Do")
	root.geometry("+200+100")
	root.resizable(height=FALSE, width=FALSE)
	frame = Frame(root,).grid()
	root.iconbitmap('planner_icon.ico')



	### ---------------------------------- DEFINE ALL THE FUNCTIONS ---------------------------------- ###
	# - DEPENDABLE FUNCTIONS - #
	def pickle_out_to_do_tasks():
		global tasks
		pickle_out = open("tasks.pickle", "wb")
		pickle.dump(tasks, pickle_out, pickle.HIGHEST_PROTOCOL)
		pickle_out.close()

	def pickle_out_completed_tasks():
		global completed_tasks
		pickle_out_completed_tasks = open("completed_tasks.pickle", "wb")
		pickle.dump(completed_tasks, pickle_out_completed_tasks, pickle.HIGHEST_PROTOCOL)
		pickle_out_completed_tasks.close()

	def update_listbox():
		clear_listbox()
		global tasks
		for task in tasks:
			lb_tasks.insert(END, task)
		pickle_out_to_do_tasks()

	def clear_listbox():
		lb_tasks.delete(0, END)

	def update_completed_listbox():
		global completed_tasks
		clear_completed_tasks()
		for task in completed_tasks:
			lb_completed_tasks.insert(END, task)
		pickle_out_completed_tasks()

	def clear_completed_tasks():
		lb_completed_tasks.delete(0,END)


	# - DEPENDABLE FUNCTIONS - #


	def add_task():
		task = txt_input.get()
		if task != "":
			tasks.append(task)
			update_listbox()
		else:
			lbl_display.configure(text="Please enter a task.")
		txt_input.delete(0,END)

	def task_completed():
		global completed_tasks
		task = lb_tasks.get("active")
		completed_tasks.append(task)
		if task in tasks:
			tasks.remove(task)
		update_listbox()
		update_completed_listbox()
		pickle_out_completed_tasks()
		pickle_out_to_do_tasks()

	def sort_asc():
		tasks.sort()
		update_listbox()

	def sort_desc():
		tasks.sort()
		tasks.reverse()
		update_listbox()

	def del_one():
		task = lb_tasks.get("active")
		if task in tasks:
			tasks.remove(task)
		update_listbox()

	def del_all():
		confirmed = messagebox.askyesno("Confirm: Delete All?", "Do you really want to delete all tasks?")
		if confirmed == True:
			global tasks
			tasks = []
			update_listbox()

	def choose_random():
		global tasks
		if len(tasks) != 0:
			task = random.choice(tasks)
			lbl_display.configure(text=task)
		else:
			lbl_display.configure(text="No tasks to choose from")


	def show_number_of_tasks():
		number_of_tasks = len(tasks)
		msg = "Number of tasks = %s" %number_of_tasks
		lbl_display.configure(text=msg)

	def export_to_do_to_file():
		global radio_var
		global export_format_variable

		### ------------------- Create EXPORT FORMAT SELECTION windows ------------------- ###
		export_todo = Toplevel()
		export_todo.geometry("+460+100")
		export_todo.iconbitmap('planner_icon.ico')
		export_todo.title("Export as...")
		# export_todo.resizable(height=FALSE, width=FALSE)
		export_todo.configure(bd=20)

		def proceed():
			global radio_var, export_format_variable
			global tasks
			global radio_var
			global export_format_variable
			export_format_variable = radio_var.get() ### VARIABLE IS 1 FOR EXCEL :: VARIABLE IS 2 FOR PDF ###
			export_todo.quit()
			export_todo.destroy()

			### ------------ Ask user, the FILE SAVE PATH ------------ ###
			root.savetopath = askdirectory()

			### ------------------ ACTUAL EXPORTING PART ------------------ ###
			if export_format_variable == "1":
				file_location = "%s/To-do-list_%s.xlsx" %(root.savetopath, time.strftime("%d-%m-%Y_%H%M%S"))
				to_do_list_workbook = xlsxwriter.Workbook(file_location)
				to_do_list_worksheet = to_do_list_workbook.add_worksheet()
				# Widen the first column to make the text clearer.
				to_do_list_worksheet.set_column('A:A', 40)

				# Write into
				to_do_list_worksheet.write("A1", "Tasks left to do:")

				for i in range(len(tasks)):
					to_do_list_worksheet.write("A%d"%(i+2), tasks[i])

				to_do_list_workbook.close()
				lbl_display.configure(text="List exported successfully.")

			elif export_format_variable == "2":
				file_location = "%s/To-do-list_%s.pdf"%(root.savetopath, time.strftime("%d-%m-%Y_%H%M%S"))
				class PDF(FPDF):
					def header(self):
						# Logo
						self.image('planner_icon.png', 82, 11, 7)
						# Arial bold 15
						self.set_font('Helvetica', 'B', 20)
						# Move to the right
						self.cell(80)
						# Title
						self.cell(50, 10, "To do list" , 0, 0, 'L')
						# Line break
						self.ln(20)

					# Page footer
					def footer(self):
						# Position at 1.5 cm from bottom
						self.set_y(-15)
						# Arial italic 8
						self.set_font('Helvetica', '', 8)
						# Page number
						self.cell(0, 10, 'Page ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')

				# Instantiation of inherited class
				pdf = PDF()
				pdf.alias_nb_pages()
				pdf.add_page()

				pdf.set_font('Helvetica', 'I', 12)
				pdf.cell(w=10)
				pdf.cell(0, 8, "Tasks left to do:", 0, 1)

				pdf.set_font('Helvetica', '', 12)
				for i in range(0,len(tasks)):
					pdf.cell(w=10)
					pdf.cell(0, 10, "%d.\t%s" %(i+1, tasks[i]), 0, 1)
				pdf.output(file_location, 'F')
				lbl_display.configure(text="List exported successfully.")

			else:
				lbl_display.configure(text="No export format selected.")

		### ---------- OPTION RADIO BUTTONS ---------- ###
		radio_var = StringVar()
		# radio_var.set(0)
		lbl_export_options_to_do = Label(export_todo, text="Select export format:")
		radio_btn_excel_to_do = ttk.Radiobutton(export_todo, text="Excel file (.xlsx)", value='1', variable=radio_var)
		radio_btn_pdf_to_do   = ttk.Radiobutton(export_todo, text="PDF file (.pdf)",    value='2', variable=radio_var)
		btn_proceed_with_selected_option = ttk.Button(export_todo, text="Continue", command=proceed)

		lbl_export_options_to_do.grid(row=0, column=0, stick=W+E)
		radio_btn_excel_to_do.grid(row=1, column=0)
		radio_btn_pdf_to_do.grid(row=2, column=0)
		btn_proceed_with_selected_option.grid(row=3, column=0, stick=W+E)

		### ------------------- Mainloop EXPORT FORMAT SELECTION windows ------------------- ###
		export_todo.mainloop()

	def move_to_to_do():
		global completed_tasks
		global tasks
		task = lb_completed_tasks.get("active")
		tasks.append(task)
		update_listbox()
		if task in completed_tasks:
			completed_tasks.remove(task)

		update_completed_listbox()
		pickle_out_completed_tasks()
		pickle_out_to_do_tasks()

	def export_completed_tasks_to_file():
		global radio_var
		global export_format_variable

		### ------------------- Create EXPORT FORMAT SELECTION windows ------------------- ###
		export_todo = Toplevel()
		export_todo.geometry("+460+100")
		export_todo.title("Export as...")
		export_todo.iconbitmap('planner_icon.ico')
		export_todo.configure(bd=20)

		def proceed():
			global radio_var, export_format_variable
			global completed_tasks
			global radio_var
			global export_format_variable
			export_format_variable = radio_var.get() ### VARIABLE IS 1 FOR EXCEL :: VARIABLE IS 2 FOR PDF ###
			export_todo.quit()
			export_todo.destroy()

			##### ------ ASK FILESAVE PATH ------ #####
			root.savetopath = askdirectory()

			### ------------------ ACTUAL EXPORTING PART ------------------ ###
			if export_format_variable == "1":
				file_location = "%s/Completed-tasks-list_%s.xlsx" %(root.savetopath, time.strftime("%d-%m-%Y_%H%M%S"))
				to_do_list_workbook = xlsxwriter.Workbook(file_location)
				to_do_list_worksheet = to_do_list_workbook.add_worksheet()
				# Widen the first column to make the text clearer.
				to_do_list_worksheet.set_column('A:A', 40)

				# Write into
				to_do_list_worksheet.write("A1", "Completed tasks:")

				for i in range(len(completed_tasks)):
					to_do_list_worksheet.write("A%d"%(i+2), completed_tasks[i])

				to_do_list_workbook.close()
				lbl_display.configure(text="List exported successfully.")

			elif export_format_variable == "2":
				file_location = "%s/Completed-tasks-list_%s.pdf" %(root.savetopath, time.strftime("%d-%m-%Y_%H%M%S"))
				class PDF(FPDF):
					def header(self):
						# Logo
						self.image('planner_icon.png', 82, 11, 7)
						# Arial bold 15
						self.set_font('Helvetica', 'B', 20)
						# Move to the right
						self.cell(80)
						# Title
						self.cell(50, 10, "To do list" , 0, 0, 'L')
						# Line break
						self.ln(20)

					# Page footer
					def footer(self):
						# Position at 1.5 cm from bottom
						self.set_y(-15)
						# Arial italic 8
						self.set_font('Helvetica', '', 8)
						# Page number
						self.cell(0, 10, 'Page ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')

				# Instantiation of inherited class
				pdf = PDF()
				pdf.alias_nb_pages()
				pdf.add_page()

				pdf.set_font('Helvetica', 'I', 12)
				pdf.cell(w=10)
				pdf.cell(0, 8, "Completed tasks:", 0, 1)

				pdf.set_font('Helvetica', '', 12)
				for i in range(0,len(completed_tasks)):
					pdf.cell(w=10)
					pdf.cell(0, 10, "%d.\t%s" %(i+1, completed_tasks[i]), 0, 1)
				pdf.output(file_location, 'F')
				lbl_display.configure(text="List exported successfully.")

			else:
				lbl_display.configure(text="No export format selected.")




		### ---------- OPTION RADIO BUTTONS ---------- ###
		radio_var = StringVar()
		# radio_var.set(0)
		lbl_export_options_to_do = Label(export_todo, text="Select export format:")
		radio_btn_excel_to_do = ttk.Radiobutton(export_todo, text="Excel file (.xlsx)", value='1', variable=radio_var)
		radio_btn_pdf_to_do   = ttk.Radiobutton(export_todo, text="PDF file (.pdf)",    value='2', variable=radio_var)
		btn_proceed_with_selected_option = ttk.Button(export_todo, text="Continue", command=proceed)

		lbl_export_options_to_do.grid(row=0, column=0, stick=W+E)
		radio_btn_excel_to_do.grid(row=1, column=0)
		radio_btn_pdf_to_do.grid(row=2, column=0)
		btn_proceed_with_selected_option.grid(row=3, column=0, stick=W+E)

		### ------------------- Mainloop EXPORT FORMAT SELECTION windows ------------------- ###
		export_todo.mainloop()


	def exit():
		quit()

	### ---------------------------------- DEFINE ALL THE FUNCTIONS ---------------------------------- ###


	### ------------------------------------- DEFINE ALL WIDGETS ------------------------------------- ###
	#Creating widgets (Column 0)
	lbl_title = Label(frame, text="To-Do-List", bg="white")
	lbl_title.grid(row=0, column=0, stick=W+E)

	btn_add_task = ttk.Button(frame, text="Add a task", command=add_task)
	btn_add_task.grid(row=1, column=0, stick=W+E)

	btn_task_completed = ttk.Button(frame, text="Task completed", command=task_completed)
	btn_task_completed.grid(row=2, column=0, stick=W+E)

	btn_sort_asc = ttk.Button(frame, text="Sort A-z", command=sort_asc)
	btn_sort_asc.grid(row=3, column=0, stick=W+E)

	btn_sort_desc = ttk.Button(frame, text="Sort z-A", command=sort_desc)
	btn_sort_desc.grid(row=4, column=0, stick=W+E)

	btn_del_one = ttk.Button(frame, text="Delete task", command=del_one)
	btn_del_one.grid(row=5, column=0, stick=W+E)

	btn_del_all = ttk.Button(frame, text="Delete All", command=del_all)
	btn_del_all.grid(row=6, column=0, stick=W+E)

	btn_choose_random = ttk.Button(frame, text="Choose random", command=choose_random)
	btn_choose_random.grid(row=7, column=0, stick=W+E)

	btn_number_of_tasks = ttk.Button(frame, text="Number of tasks", command=show_number_of_tasks)
	btn_number_of_tasks.grid(row=8, column=0, stick=W+E)

	lbl_about = Label(frame, text="©2018, Aashay Umesh Sathe, E-mail: aashaysathe.as@gmail.com", fg="#bababa", bg="white")
	lbl_about.grid(row=9, column=0, columnspan=2, stick=W)



	#Creating widgets (Column 1)

	lbl_display = Label(frame, text="", bg="white")
	lbl_display.grid(row=0, column=1, columnspan=1, stick=W+E)

	txt_input = ttk.Entry(frame, width=50)
	txt_input.grid(row=1, column=1, columnspan=1, stick=W+E)

	#Defining To-Do-List Listbox
	lb_tasks = Listbox(frame, selectbackground="#92bdff", relief=FLAT)
	lb_tasks.grid(row=2, column=1, rowspan=6, columnspan=1, stick=N+E+W+S)
	update_listbox()
	#Defining To-Do-List Scrollbar
	lb_scrollbar = ttk.Scrollbar(lb_tasks)
	lb_scrollbar.pack(side=RIGHT, fill=Y)
	#Linking Scrollbar to Listbox functionally
	lb_tasks.config(yscrollcommand=lb_scrollbar.set)
	lb_scrollbar.config(command=lb_tasks.yview)

	btn_export_to_do_to_file = ttk.Button(frame, text="Export to-do to file", command=export_to_do_to_file)
	btn_export_to_do_to_file.grid(row=8, column=1, columnspan=1, stick=W+E)



	#Creating widgets (Column 2)

	lbl_completed_tasks_title = Label(frame, text="Completed tasks", width=30, bg="white")
	lbl_completed_tasks_title.grid(row=0, column=2)

	#Defining Completed Tasks Listbox
	lb_completed_tasks = Listbox(frame, selectbackground="#92bdff", relief=FLAT, background="#bababa")
	lb_completed_tasks.grid(row=1, column=2, rowspan=6, stick=N+E+W+S)
	#Defining Completed Tasks Scrollbar
	lb_completed_tasks_scrollbar = ttk.Scrollbar(lb_completed_tasks, orient=VERTICAL)
	lb_completed_tasks_scrollbar.pack(side=RIGHT, fill=Y)
	#Linking Scrollbar to Listbox functionally
	lb_completed_tasks.config(yscrollcommand=lb_completed_tasks_scrollbar.set)
	lb_completed_tasks_scrollbar.config(command=lb_completed_tasks.yview)

	btn_move_to_to_do = ttk.Button(frame, text="Task not completed", command=move_to_to_do)
	btn_move_to_to_do.grid(row=7, column=2, stick=E+W)

	btn_export_completed_tasks_to_file = ttk.Button(frame, text="Export completed to file", command=export_completed_tasks_to_file)
	btn_export_completed_tasks_to_file.grid(row=8, column=2, columnspan=1, stick=W+E)

	btn_exit = ttk.Button(frame, text="Exit", command=exit)
	btn_exit.grid(row=9, column=2, columnspan=1, stick=W+E)

	update_completed_listbox()

	#Defining Buttons
	# btn_clear_completed_tasks = ttk.Button(frame, text="Clear all completed tasks", command=clear_completed_tasks)
	# btn_clear_completed_tasks.grid(row=20, stick=E+W)

	### ------------------------------------- DEFINE ALL WIDGETS ------------------------------------- ###


	root.mainloop()
	######################################### WINDOW LOOPED (root) ########################################

if __name__ == '__main__':
    main()
from tkinter import *
from tkinter import messagebox
import os
import Function as fn
import Excel_func as xlfn


########################################## GUI funkce #############################################

######################## search project button ########################

def btn_search():
	projekt = fn.projekt(entry0.get())
	global info
	info = projekt.hledani_info_mesta()	#return informace_o_projektu (dict)
		
	canvas.delete('info','mesto_info')	# delete previous text
	# print(info)
	
	if 'k.u.' not in info.keys(): #pokud nenajde k.u.
		messagebox.showerror(title='Error', message="k.u. nenalezeno")
		canvas.create_text(
		244.0, 149.0,
		text = f"{info['nazev_projektu']}",
		fill = "#000000",
		font = ('calibri',10,'bold'),
		tag = 'info')		


	canvas.create_text(
		244.0, 149.0,
		text = f"{info['nazev_projektu']}",
		fill = "#000000",
		font = ('calibri',10,'bold'),
		tag = 'info')

	canvas.create_text(
    	244.0, 179.5,
    	text = f"{info['jmeno_projektanta']}",
    	fill = "#000000",
    	font = ('calibri',10,'bold'),
    	tag = 'info')

	canvas.create_text(
	    125.0, 212.0,
	    text = f"{info['rok_projektu']}",
	    fill = "#000000",
	    font = ('calibri',10,'bold'),
	    tag = 'info')

	canvas.create_text(
    	313.0, 209.0,
	    text = f"{info['kod_projektu']}",
	    fill = "#000000",
	    font = ('calibri',10,'bold'),
	    tag = 'info')

	canvas.create_text(
	    141.0, 246.0,
	    text = f"{info['Obec']}",
	    fill = "#000000",
	    font = ('calibri',10,'bold'),
	    tag = 'mesto_info')

	canvas.create_text(
	    305.0, 246.0,
	    text = f"{info['k.u.']}",
	    fill = "#000000",
	    font = ('calibri',10,'bold'),
	    tag = 'mesto_info')

	canvas.create_text(
	    136.0, 280.0,
	    text = f"{info['Okres']}",
	    fill = "#000000",
	    font = ('calibri',10,'bold'),
	    tag = 'mesto_info')

	canvas.create_text(
    	305.0, 280.0,
    	text = f"{info['Kraj']}",
    	fill = "#000000",
    	font = ('calibri',10,'bold'),
    	tag = 'mesto_info')

	print("Search Button Clicked")


################## Thunder button ###############################
def btn_thunder():
	ask =messagebox.askquestion(title='Are you sure about that?', message="Opravdu chcete pokračovat?")

	if ask == 'yes':
		print (ask)

		if pruvodni.get() == 1:
			print ("Pruvodni: On")
			info_z_search = info
			ZAD = fn.zadani(entry0.get())
			info_ZAD = ZAD.info_zadani()
			if info_ZAD is None:
				print ('Chyba zadani')
				return  
			info_z_search.update(info_ZAD)
			
			doc = fn.doc_manipulation(info_z_search)
			doc.pruvodni_zprava()

		else:
			print ("Pruvodni: Off")

		if stitky.get() == 1:
			print("Stitky: On")
			excel = xlfn.excel_manipulation(info)
			excel.stitky_pdf()
		else:
			print("Stitky: Off")

		if doklad1_obsah.get() == 1:
			print("Doklad1_obsah: On")
			
		else:
			print("Doklad1_obsah: Off")

		
	else:
		print (ask)
		pass

def close_win(win):
	pass


#################### Pop up window #################################
def pop_up(n_error, string):
	top = Toplevel(window)
	top.title(f"{n_error}")
	top.resizable(False, False)

	label = Label(top, text= f"{string}")
	label.pack()
	return 


#############################################################

cwd = os.getcwd()		#curent script dicrectory
img_dicrectory = rf"{cwd}\img"


############################ GUI ###############################
window = Tk()

window.geometry("400x500")
window.configure(bg = "#ffffff")
window.title("R-built s.r.o.")
window.iconbitmap(img_dicrectory+r'\R-built logo.ico')
canvas = Canvas(
    window,
    bg = "#ffffff",
    height = 500,
    width = 400,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge")
canvas.place(x = 0, y = 0)

background_img = PhotoImage(file = img_dicrectory+r"\background.png")
background = canvas.create_image(
    199.5, 259.0,
    image=background_img)

entry0_img = PhotoImage(file = img_dicrectory+r"\img_textBox0.png")
entry0_bg = canvas.create_image(
    169.0, 319.0,
    image = entry0_img)

entry0 = Entry(
    bd = 0,
    bg = "#d9d9d9",
    highlightthickness = 0,
    font = ("Calibri", 12, "bold"),
    text = 'rr/nn')

entry0.place(
    x = 144.0, y = 308,
    width = 50.0,
    height = 20)

canvas.create_text(
    305.0, 280.0,
    text = "",
    fill = "#000000",
    font = ("Calibri", 10, "bold"),
    tag = 'info')

canvas.create_text(
    136.0, 280.0,
    text = "",
    fill = "#000000",
    font = ('calibri',10,'bold'),
    tag = 'info')

canvas.create_text(
    305.0, 246.0,
    text = "",
    fill = "#000000",
    font = ('calibri',10,'bold'),
    tag = 'info')

canvas.create_text(
    141.0, 246.0,
    text = "",
    fill = "#000000",
    font = ('calibri',10,'bold'),
    tag = 'info')

canvas.create_text(
    313.0, 209.0,
    text = "",
    fill = "#000000",
    font = ('calibri',10,'bold'),
    tag = 'info')

canvas.create_text(
    125.0, 212.0,
    text = "",
    fill = "#000000",
    font = ('calibri',10,'bold'),
    tag = 'info')

canvas.create_text(
    244.0, 179.5,
    text = "",
    fill = "#000000",
    font = ('calibri',10,'bold'),
    tag = 'info')

canvas.create_text(
    244.0, 149.0,
    text = "",
    fill = "#000000",
    font = ('calibri',10,'bold'),
    tag = 'info')

img0 = PhotoImage(file = img_dicrectory+r"\img0.png")
b_search = Button(
    image = img0,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_search,
    relief = "flat")

b_search.place(
    x = 283, y = 305,
    width = 83,
    height = 28)

img1 = PhotoImage(file = img_dicrectory+r"\img1.png")
b_thunder = Button(
    image = img1,
    borderwidth = 0,
    highlightthickness = 0,
    command = btn_thunder,
    relief = "flat")

b_thunder.place(
    x = 302, y = 446,
    width = 67,
    height = 22)


################### checkbox ##########################
stitky = IntVar()
c_stitky = Checkbutton(
	window,
	variable = stitky,
	onvalue = 1,
	offvalue = 0,
	)
c_stitky.place(
	x = 105, y = 365,
    width = 13,
    height = 13,
	)

pruvodni = IntVar()
c_pruvodni = Checkbutton(
	window,
	variable = pruvodni,
	onvalue = 1,
	offvalue = 0,
	)
c_pruvodni.place(
	x = 105, y = 390,
    width = 13,
    height = 13,
	)
doklad1_obsah = IntVar()
c_doklad1_obsah = Checkbutton(
	window,
	variable = doklad1_obsah,
	onvalue = 1,
	offvalue = 0,
	)
c_doklad1_obsah.place(
	x = 105, y = 412,
    width = 13,
    height = 13,
	)
#######################################################
window.resizable(False, False)
window.mainloop()


##########################################################################



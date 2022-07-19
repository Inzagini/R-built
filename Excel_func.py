import openpyxl, glob, os, json
from openpyxl import Workbook as wb
from win32com import client
import Function as fn
from docx import Document


def test():
	projekt_ID = '22/41'
	projekt = fn.projekt(projekt_ID)
	info = projekt.hledani_info_mesta()
	
	excel = excel_manipulation(info)
	stitky = excel.stitky_pdf()

	return


class excel_manipulation:
	def __init__(self,info_dict):
		self.project_dict = info_dict
		self.user_name = os.getlogin()
		cwd = os.getcwd()
		#load potrebne odkazy	
		self.Titulni_strany = rf'{cwd}\Dokumenty\Titulní strany a příprava R.xlsx'

		#open setting na odkzay
		with open("setting.json","r",encoding = 'utf-8') as f:
			file = f.read()
			self.setting = json.loads(file)
		#open vzorovy excel
		self.open_xl()
		#vyplnit vzorovy excel
		self.vypln_excel_file = self.vypln_zad_list()

	def open_xl(self): #open zadani list na vyplneni info -> return aktivni excel list
		wb = openpyxl.load_workbook(self.Titulni_strany) #load seznam projectu file
		ws = wb.sheetnames

		for index,sheet in enumerate(ws): 
			if self.setting["Titulni_strany"]["zadani_list"] == sheet:
				company = 'R-built'
				wb.active = index
				ws= wb.active
				self.ws = ws
				self.wb = wb

		return	

	def vypln_zad_list(self):
		# user_name = self.user_name
		project_file_directory = self.project_dict["project_file_directory"]
		pruvodni_zprava_directory = rf'{project_file_directory}\dokumentace na SÚ\textová část'
		cwd = os.getcwd()

		files = glob.glob(rf'{pruvodni_zprava_directory}\*.docx')
		
		for file_name in files:
			# print (file_name)
			if 'A_Průvodní' or 'Technicka' in file_name:
				pruvodni_zprava_directory = file_name
				# print ("Hotovo: ", pruvodni_zprava_directory)
			else :
				print ("Pruvodni zprava nenalezena")
				return None

		with open(pruvodni_zprava_directory,'rb') as f:
				document = Document(f)

		for index,paragraph in enumerate(document.paragraphs):
			if 'Datum vydání' in paragraph.text:
				datum_vydani = paragraph.text.replace("Datum vydání", "").replace(":", "").lstrip()
				self.project_dict.setdefault("datum_vydani", datum_vydani)

		

		self.ws[str(self.setting["Titulni_strany"]["Projektant"])] = self.project_dict["jmeno_projektanta"]
		self.ws[str(self.setting["Titulni_strany"]["Nazev_stavby"])] = self.project_dict["nazev_projektu"]
		self.ws[str(self.setting["Titulni_strany"]["interni_cislo"])] = f'{self.project_dict["cislo_projektu"]}/{self.project_dict["rok_projektu"]}'
		self.ws[str(self.setting["Titulni_strany"]["datum"])] = self.project_dict["datum_vydani"]
		self.ws[str(self.setting["Titulni_strany"]["Cislo_stavby"])] = self.project_dict["kod_projektu"]
		self.ws[str(self.setting["Titulni_strany"]["Tel_cislo_projektant"])] = self.project_dict["Tel_cislo"]

		out_excel_file = rf'C:\Users\{self.user_name}\Desktop\ŠTÍTKY.xlsx'	
		self.wb.save(filename=out_excel_file)
		self.wb.close()

		return out_excel_file

	def stitky_pdf (self):

		wb = openpyxl.load_workbook(self.Titulni_strany) #load seznam projectu file
		ws = wb.sheetnames		

		
		
		excel = client.Dispatch("Excel.Application")
		excel.Visible = False

		sheets = excel.Workbooks.Open(self.vypln_excel_file)
		for index,sheet in enumerate(ws): 
			if self.setting["Titulni_strany"]["stitky"] == sheet:
				company = 'R-built'
				work_sheets = sheets.Worksheets[index]
				work_sheets.ExportAsFixedFormat(0, rf'C:\Users\{self.user_name}\Desktop\ŠTÍTKY.pdf')
				sheets.Close(True)
		wb.close()

		try:
			os.remove(self.vypln_excel_file)
		except:
			pass

		return


if __name__ == '__main__':
	test()
import os, sys, glob, json, openpyxl
from PyPDF2 import PdfMerger, PdfWriter,PdfReader
from openpyxl import Workbook as wb
import Function as fn

def test():


	projekt_ID = '22/63'
	project = fn.projekt(projekt_ID)
	info = project.search_project()
	
	
	PDF = PDF_manipulation(info)
	CEZ = PDF.CEZ()
	print(CEZ)
	
	# file = PDF.find_file('ČEZ','Sdělení')

	# print(file)



	return

class PDF_manipulation:

	def __init__(self,info_dict):
		self.project_dict = info_dict
		self.project_file_directory = self.project_dict["project_file_directory"]
		self.site_directory = rf'{self.project_file_directory}\dokumentace na SÚ\dokladová část\vyjádření-inženýrské sítě'
		
	#hledani spravny PDF file ve slozce -> list file
	def find_file(self,sit,pozadavek):

		dict_dir = {}

		#hledani vsech slozek splneni podminky sit
		for dir_ in os.listdir(self.site_directory):
			if sit in dir_:
				dict_dir.setdefault(f"{dir_}", self.site_directory + rf'\{dir_}')

		# site_directory = self.site_directory + rf'\{sit}'
		dict_file = {}
		for sit in dict_dir.keys():

			site_directory = dict_dir[f"{sit}"]

			if 'us' in os.listdir(site_directory):
				site_directory = site_directory + r'\us'

			
			for filename in os.listdir(site_directory):
				
				if pozadavek in filename:
					# print (filename)
					dict_file.setdefault(f"{sit}", site_directory+rf'\{filename}')
		# print(list_file)
		return dict_file

	def PDF_reader(self,pdf_file_directory):
		reader = PdfReader(pdf_file_directory)
		page = reader.pages[0]
		text = page.extractText()
		text_list = text.split("\n")

		return text_list

	#sit
	def CETIN (self):

		
		CETIN = {
		"cislo_jednaci" : "Číslo jednací:",
		"duvod_vyjadreni" : "Důvod vyjádření",
		"platnost_vyjadreni" : "Platnost Vyjádření",
		"stret" : "ojde ke střetu"
		}

		info_CETIN = {}
		
		# extrakt informaci
		def sub_CETIN():

			for text in pdf_text:
				# print('SUB:', text)
				if CETIN["cislo_jednaci"] in text:
					cislo_jednaci = text.split(" ")[2].lstrip()
					info_CETIN["cislo_jednaci"] = cislo_jednaci
					# print('Číslo jednací:', cislo_jednaci)	

				if CETIN["duvod_vyjadreni"] in text:
					text_index = text.find(CETIN["duvod_vyjadreni"])
					duvod_vyjadreni = text[text_index+len(CETIN["duvod_vyjadreni"])+2::]
					info_CETIN["duvod_vyjadreni"] = duvod_vyjadreni
					# print('Důvod vyjádření:', duvod_vyjadreni)

				if CETIN["platnost_vyjadreni"] in text:
					text_index = text.find(CETIN["platnost_vyjadreni"])
					f_index = text.find('(')
					platnost_vyjadreni = text[text_index+len(CETIN["platnost_vyjadreni"]):f_index].replace(" ", "")
					info_CETIN["platnost_vyjadreni"] = platnost_vyjadreni
					# print('Platnost vyjádření:', platnost_vyjadreni)
			return info_CETIN
		
		#list pozadovany souboru
		dict_file= self.find_file('CETIN','Vy')

		#loop pdf souboru z listu
		for key,pdf_file in dict_file.items():
			pdf_text = self.PDF_reader(pdf_file)

			 #hleda jestli dojde ke stretu
			for text in pdf_text:
				# print(index, text)

				if CETIN["stret"] in text:
					stret_string = text.split(" ")
					print('Test Stret...')

					#pokud nedojde ke stretu			
					if stret_string[0] == 'Nedojde':	
						info_CETIN = sub_CETIN()
						info_CETIN.setdefault("stret", "Nejsou dotčeni")
						return info_CETIN


					#pokud dojde ke stretu
					if stret_string[0] == 'Dojde':
						info_CETIN = sub_CETIN()

						if info_CETIN["duvod_vyjadreni"] == 'Informace o poloze sítě':
							pass
						if info_CETIN["duvod_vyjadreni"] == 'Územní souhlas':
							info_CETIN.setdefault("stret", "Střet")
							return info_CETIN
		
		return


	def CRA (self):

		CRA = {
			"cislo_jednaci" : "UPTS",
			"duvod_vyjadreni" : "ÒþHO",
			"platnost_vyjadreni" : "Platnost Vyjádření",
			"neni_dotcen_CRA" : r".9DãtåiGRVWLRY\MiGĜHQtNH[LVWHQFLVtWt9iPVGČOXMHPHåHYH9iPLY\]QDþHQpPĜHãHQpP~]HPtQHGRMGHNHVW\NXVåiGQêPSRG]HPQtPYHGHQtP]DĜt]HQtPYQDãtVSUiYČ",		

		}

		info_CRA = {}

		dict_file = self.find_file('CRA','VY')
		

		for pdf_file in dict_file.values():
			pdf_text = self.PDF_reader(pdf_file)

			for index,text in enumerate(pdf_text):
				print(index, text)

				# if CRA["stret"] in text:
				# 	pass
		return

	def CEZ (self):
		CEZ = {
		"cislo_jednaci" : "NAŠE ZNAČKA",
		"duvod_vyjadreni" : "Důvod vyjádření",
		"platnost_vyjadreni" : "Toto sdělení je platné do ",
		"stret" : "střet"
		}

		info_CEZ = {}

		#podpurna funkce pro ziskani informaci
		def sub_CEZ(key,pdf_file):

			sub_dict = {}
			# print(key)
			for index, text in enumerate(pdf_text):
				# print(index, text)


				if CEZ["stret"] in text:
					# print("Nachazi")
					sub_dict["stret"] = "střet"
				
				if CEZ["cislo_jednaci"] in text:
					cislo_jednaci_index = index

		
				if "cislo_jednaci_index" in locals() and "střet" in sub_dict.values():
					sub_dict["cislo_jednaci"] = pdf_text[cislo_jednaci_index+1].split(" ")[0]
					# print("Nachazi")
								

				if "cislo_jednaci_index" in locals() and "střet" not in sub_dict.values():
					sub_dict["cislo_jednaci"] = pdf_text[cislo_jednaci_index+1][10::]


				if CEZ["platnost_vyjadreni"] in text:
					text_index = text.find(CEZ["platnost_vyjadreni"])
					string_len = len(CEZ["platnost_vyjadreni"])
					date_len = len("XX.YY.ZZZZ")
					platnost_vyjadreni = text[text_index+string_len:string_len+date_len]
					sub_dict["platnost_vyjadreni"] = platnost_vyjadreni

			if "stret" not in sub_dict:
					sub_dict["stret"] = 'nechází'

			info_CEZ[f"{key}"] = sub_dict

			return info_CEZ



		dict_file= self.find_file('ČEZ','Sdělení')

		
		#loop pdf souboru z listu
		for key,pdf_file in dict_file.items():
			pdf_text = self.PDF_reader(pdf_file)
			# print(key)
			 #ziskani info z pdf


			info_CEZ =sub_CEZ(key,pdf_text)
			

		return info_CEZ

if __name__ == '__main__':
	test()
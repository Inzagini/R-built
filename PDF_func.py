import os, sys, glob, json, openpyxl, datetime
from PyPDF2 import PdfMerger, PdfWriter,PdfReader
from openpyxl import Workbook as wb
import Function as fn

def test():


	projekt_ID = '22/95'
	project = fn.projekt(projekt_ID)
	info = project.search_project()
	
	
	PDF = PDF_manipulation(info)
	GasNet = PDF.T_mobile()
	print(GasNet)
	
	# file = PDF.find_file('GasNet','stanovisko')

	# print(file)



	return

class PDF_manipulation:

	def __init__(self, info_dict):
		self.project_dict = info_dict
		self.project_f_dir = self.project_dict["project_f_dir"]
		self.site_directory = rf'{self.project_f_dir}\dokumentace na SÚ\dokladová část\vyjádření-inženýrské sítě'
		
	#hledani spravny PDF file ve slozce -> list file
	def find_file(self, sit, pozadavek):

		dict_dir = {}

		#hledani vsech slozek splneni podminky sit
		for dir_ in os.listdir(self.site_directory):
			if sit in dir_:
				dict_dir.setdefault(f"{dir_}", self.site_directory + rf'\{dir_}')

		dict_file = {}
		for sit in dict_dir.keys():

			site_directory = dict_dir[f"{sit}"]

			if 'us' in os.listdir(site_directory):
				site_directory = site_directory + r'\us'

			os.chdir(site_directory)
			for filename in glob.glob("*.pdf"):	
				if pozadavek in filename or pozadavek == 'Neni':
					dict_file.setdefault(f"{sit}", site_directory+rf'\{filename}')

		
		return dict_file

	def PDF_reader(self,pdf_file_directory):
		reader = PdfReader(pdf_file_directory)
		page = reader.pages[0]
		
		
		text = page.extract_text()	
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
					sub_dict["cislo_jednaci"] = cislo_jednaci
					# print('Číslo jednací:', cislo_jednaci)	

				if CETIN["duvod_vyjadreni"] in text:
					text_index = text.find(CETIN["duvod_vyjadreni"])
					duvod_vyjadreni = text[text_index+len(CETIN["duvod_vyjadreni"])+2::]
					sub_dict["duvod_vyjadreni"] = duvod_vyjadreni
					# print('Důvod vyjádření:', duvod_vyjadreni)

				if CETIN["platnost_vyjadreni"] in text:
					text_index = text.find(CETIN["platnost_vyjadreni"])
					f_index = text.find('(')
					platnost_vyjadreni = text[text_index+len(CETIN["platnost_vyjadreni"]):f_index].replace(" ", "")
					sub_dict["platnost_vyjadreni"] = platnost_vyjadreni
					# print('Platnost vyjádření:', platnost_vyjadreni)
			return sub_dict
		
		#list pozadovany souboru
		dict_file= self.find_file('CETIN','Vy')

		#loop pdf souboru z listu
		for key,pdf_file in dict_file.items():
			pdf_text = self.PDF_reader(pdf_file)

			sub_dict = {}
			 #hleda jestli dojde ke stretu
			
			for text in pdf_text:
				# print(index, text)

				if CETIN["stret"] in text:
					stret_string = text.split(" ")
					print('Test Stret...')

					#pokud nedojde ke stretu			
					if stret_string[0] == 'Nedojde':
						print('Nedojde')	
						sub_CETIN()
						sub_dict.setdefault("stret", "Nejsou dotčeni")
						


					#pokud dojde ke stretu
					if stret_string[0] == 'Dojde':
						print('Dojde')
						sub_CETIN()

						if sub_dict["duvod_vyjadreni"] == 'Informace o poloze sítě':
							pass
						if sub_dict["duvod_vyjadreni"] == 'Územní souhlas':
							sub_dict.setdefault("stret", "Střet")
			
			info_CETIN[f"{key}"] = sub_dict

		return info_CETIN
		


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

	def GasNet (self):
		GasNet = {
		"cislo_jednaci" : "naše značka",
		"duvod_vyjadreni" : "Účel stanoviska:",
		"platnost_vyjadreni" : "datum",
		"stret" : "nejsou umístěna"
		}

		info_GasNet = {}

		dict_file= self.find_file('GasNet','stanovisko')

		
		#loop pdf souboru z listu
		for key,pdf_file in dict_file.items():
			pdf_text = self.PDF_reader(pdf_file)
			# print(key)
			 #ziskani info z pdf
			sub_dict = {}
			for index, text in enumerate(pdf_text):
				# print(index, text)

				if GasNet["stret"] in text:
					# print("Nachazi")
					sub_dict["stret"] = "nenachází"
				
				if GasNet["cislo_jednaci"] in text:
					sub_dict["cislo_jednaci"] = pdf_text[index+1]

				if GasNet["platnost_vyjadreni"] in text:
					datum = pdf_text[index+1][:10]
					datum_ = [n for n in map(int,datum.split("."))]
					datum_novy = f"{datum_[0]}.{datum_[1]}.{datum_[2]+2}"
					
					sub_dict["platnost_vyjadreni"] = datum_novy
					
		# info_GasNet = 
		info_GasNet[f"{key}"] = sub_dict

		return info_GasNet


	def scvk (self):
		scvk = {
		"cislo_jednaci" : "Naše značka",
		"duvod_vyjadreni" : "Existence zařízení ve správě SČVK",
		"platnost_vyjadreni" : "Datová schránka",
		"stret" : "nenachází"
		}

		info_scvk = {}

		dict_file= self.find_file('SčVK','Neni')
		
		#loop pdf souboru z listu
		for key,pdf_file in dict_file.items():

			#pokud je uzemni souhlas (protoze to je obrazek)
			if 'us' in pdf_file:
				return None

			pdf_text = self.PDF_reader(pdf_file)
			
			if len(pdf_text) <= 1:
				
				return 'prazdny'

			# print(key)
			#ziskani info z pdf
			sub_dict = {}
			for index, text in enumerate(pdf_text):
				# print(index, text)

				if scvk["duvod_vyjadreni"] in text:
					duvod_vyjadreni = "Existence zařízení"
					sub_dict["duvod_vyjadreni"] = duvod_vyjadreni

				if scvk["stret"] in text:
					sub_dict["stret"] = "nenachází"

				if scvk["cislo_jednaci"] in text:
					cislo_jednaci = text[:text.find(scvk["cislo_jednaci"])]
					sub_dict["cislo_jednaci"] = cislo_jednaci
					
				
				if scvk["platnost_vyjadreni"] in text:
					platnost_vyjadreni = text[:text.find(scvk["platnost_vyjadreni"])]
					sub_dict["platnost_vyjadreni"] = platnost_vyjadreni


			#pokud to neni existence a je stret
			if sub_dict["duvod_vyjadreni"] and sub_dict["stret"] not in sub_dict.values():
				return None

		info_scvk[f"{key}"] = sub_dict

		return info_scvk

	def T_mobile (self):
		Tmobile = {
		"cislo_jednaci" : [("Naše zna",3),("Číslo jednací",2)],
		"duvod_vyjadreni" : "Účel stanoviska:",
		"platnost_vyjadreni" : "V Praze dne:",
		"stret" : "souhlas s realizací stavby"
		}

		info_Tmobile = {}

		dict_file= self.find_file('T-mobile','Neni')

		
		#loop pdf souboru z listu
		for key,pdf_file in dict_file.items():
			pdf_text = self.PDF_reader(pdf_file)
			# print(key)
			 #ziskani info z pdf
			sub_dict = {}
			for index, text in enumerate(pdf_text):
				# print(index, text)

				if Tmobile["platnost_vyjadreni"] in text:
					datum = pdf_text[index+1]
					datum_ = [n for n in map(int,datum.split("."))]
					platnost_vyjadreni = f"{datum_[0]}.{datum_[1]}.{datum_[2]+1}"
					sub_dict["platnost_vyjadreni"] = platnost_vyjadreni
					
				if Tmobile["stret"] in text:
					sub_dict["stret"] = 'souhlasí'

				
				cislo_jednaci = [sub_dict.setdefault("cislo_jednaci",  pdf_text[index+off_set]) for n, off_set in Tmobile["cislo_jednaci"] if n in text if len(pdf_text[index+off_set])>0 ]
				

		if "stret" not in sub_dict:
			sub_dict["stret"] = 'Kolize'
		
		info_Tmobile[f"{key}"] = sub_dict

		return info_Tmobile


	#zatim nejde
	# def Vodafone (self):
	# 	Vodafone = {
	# 	"cislo_jednaci" : "Naše zn",
	# 	"duvod_vyjadreni" : "Účel stanoviska:",
	# 	"platnost_vyjadreni" : "V Praze dne:",
	# 	"stret" : "souhlasí s realizací projektu"
	# 	}

	# 	info_Vodafone = {}

	# 	dict_file= self.find_file('Vodafone','Vyjad')

	# 	for key,pdf_file in dict_file.items():
	# 		pdf_text = self.PDF_reader(pdf_file)
	# 	# 	# print(key)
	# 	# 	 #ziskani info z pdf
	# 	# 	sub_dict = {}
	# 	# 	for index, text in enumerate(pdf_text):
	# 	# 		print(index, text)

	# 	return

if __name__ == '__main__':
	test()
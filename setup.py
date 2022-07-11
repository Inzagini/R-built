from distutils.core import setup
import py2exe

setup(
	options={
	        'py2exe': {
	                'compressed': 2,
	                'optimize': 2,
	                'bundle_files': 1,  # 1 = .exe; 2 = .zip; 3 = separate
	                'dist_dir': 'dist',  # Put .exe in dist/
	                'xref': False,
	                'skip_archive': False,
	                'ascii': False,
	                'includes' : ["Function","os","tkinter","sys","openpyxl","glob","json", "requests","docx","b4s","datetime"],

	                'dist_dir' : r"C:\Users\David\Desktop"
	                
	                #'unbuffered': True,  # Immediately flush output.
	        }
	},
	zipfile=None,
	windows = [
				{
					'script':"R-built Tools.py",
					'icon_resource': [(0, r'C:\Users\David\Desktop\R-built Tool\img\R-built logo.ico')]
				},
				
			],
	service = ["win32com"]
)
	
	
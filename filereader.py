'''
accepts filename, return content of file if it exists
else returns an error.
'''
import string
from pathlib import Path

class FileReader:	
	@property
	def content(self):
		# path_object = Path(input("\aEnter a valid filename: "))
		path_object = Path('bearingcodes.txt')
		if path_object.exists():
			with path_object.open() as file_handler:
				return [ line.strip() for line in file_handler.readlines() if line not in string.whitespace ]
		raise Exception("\aYou might have to check the file name.")

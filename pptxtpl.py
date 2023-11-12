from fnmatch import fnmatch
from io import BytesIO
from re import compile as regex
from zipfile import ZipFile, ZIP_DEFLATED
import json

def main(infn: str, numerics: dict, vars: dict, outfn: str):

	def replacer(param, empty):
		result = empty
		if param in numerics:
			param = numerics[param]
		if param in vars:
			result = vars[param]
		result = str(result).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
		return result

	replacers = [
		{ 're': regex('I([A-Z][A-Za-z0-9]+)I'), 'sub': lambda m: replacer(m.group(1), ' ') },
		{ 're': regex('>(([A-Z][a-z0-9]+){2,})<'), 'sub': lambda m: f">{replacer(m.group(1), m.group(1))}<" },
		{ 're': regex('>9([0-9]{3,4})9<'), 'sub': lambda m: f">{replacer(m.group(1), '0')}<" },
	]

	with open(outfn, 'wb') as outfile:
		processzip(infn, outfile, replacers)

def processzip(infile, outfile, replacers: list):

	with ZipFile(infile, 'r') as inpptx, ZipFile(outfile, 'w', ZIP_DEFLATED) as outpptx:
		for name in inpptx.namelist():
			with inpptx.open(name) as memberfile:
				if fnmatch(name, 'ppt/embeddings/*'):
					with BytesIO() as embed:
						processzip(memberfile, embed, replacers)
						embed.seek(0)
						outpptx.writestr(name, embed.read())
					continue

				memberdata = memberfile.read()

				if fnmatch(name, '*.xml'):
					xml = memberdata.decode()
					for re in replacers:
						xml = re['re'].sub(re['sub'], xml)
					memberdata = xml.encode()

				outpptx.writestr(name, memberdata)

def loadJson(file: str) -> dict:

	with open(file) as f:
		return json.load(f)

if __name__ == '__main__':

	import sys
	main(sys.argv[1], loadJson(sys.argv[2]), loadJson(sys.argv[3]), sys.argv[4])

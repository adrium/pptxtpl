from fnmatch import fnmatch
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
import json

def main(infn: str, outfn: str, vars: dict, numerics: dict):
	replacements = getreplacements(vars, numerics)

	with open(outfn, 'wb') as outfile:
		processzip(infn, outfile, replacements)

def processzip(infile, outfile, replacements: dict):
	with ZipFile(infile, 'r') as inpptx, ZipFile(outfile, 'w', ZIP_DEFLATED) as outpptx:
		for name in inpptx.namelist():
			with inpptx.open(name) as memberfile:
				if fnmatch(name, 'ppt/embeddings/*'):
					with BytesIO() as embed:
						processzip(memberfile, embed, replacements)
						embed.seek(0)
						outpptx.writestr(name, embed.read())
					continue

				memberdata = memberfile.read()

				if fnmatch(name, '*.xml'):
					xml = memberdata.decode()
					for x, y in replacements.items():
						xml = xml.replace(x, y)
					memberdata = xml.encode()

				outpptx.writestr(name, memberdata)

def getreplacements(vars: dict, numerics: dict):
	result = {}

	for k, n in numerics.items():
		result[f'>9{n}9</'] = '>0</'

	for k, v in vars.items():
		if (isinstance(v, str)):
			v = v.replace('&', '&amp;')

		result[f'{{{k}}}'] = f'{v}'
		result[f'>{k}</'] = f'>{v}</'

		if k in numerics:
			n = numerics[k]
			result[f'>9{n}9</'] = f'>{v}</'

	result['{</'] = '</'
	result['>}'] = '>'

	return result

def loadJson(file: str) -> dict:
	with open(file) as f:
		return json.load(f)

if __name__ == '__main__':
	import sys
	data = {}
	try:
		data = loadJson(sys.argv[3])
	except:
		print('failed to load ' + sys.argv[3])
	if data != {}:
		main(sys.argv[1], sys.argv[2], data, loadJson(sys.argv[4]))

import comtypes.client
import os

def init_powerpoint():
	powerpoint = comtypes.client.CreateObject('Powerpoint.Application')
	powerpoint.Visible = 1
	return powerpoint

def ppt_to_pdf(powerpoint, inputfilename, outputfilename, formattype=32):
	if outputfilename[-3:] != 'pdf':
		outputfilename = outputfilename + '.pdf'
	deck = powerpoint.Presentations.Open(inputfilename)
	deck.SaveAs(outputfilename, formattype)
	deck.Close()

def convert_files_in_folder(powerpoint, folder):
	files = os.listdir(folder)
	pptfiles = [i for i in files if f.endswith(('.ppt', ',pptx'))]
	for ppt in pptfiles:
		fullpath = os.path.join(cwd, ppt)
		ppt_to_pdf(powerpoint, fullpath, fullpath)

if __name__ == "__main__":
	powerpoint = init_powerpoint()
	cwd = os.getcwd()
	convert_files_in_folder(powerpoint, cwd)
	powerpoint.Quit()
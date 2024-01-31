### Notes ###
# To use script, manually open autocad and all 100+ files you want multi burst. Run with pyautocad installed etc.
# Burst as an expresstools command doesn't work through lisp and python was easier for me to use. 
#If you don't care about speed re-enable the purge and audit lines in the 'process' function

from pyautocad import Autocad
acad = Autocad()

import time
ratelimit = 5 #send command's aysync bits break on long commands. 5 is a trial and error that runs well on burst. change if call is rejected by callee

### portion removed due to defaulting to wrong autocad versions. Would have swapped to user input but trashed the section instead ###
#from glob import glob
#process = R"C:\Users..."
#dwgs = glob(process + "\\*.dwg")
#acad.doc.SendCommand('FILEDIA 0 ') #suppress open dialogue so the command works
#for dwg in dwgs:
#	acad.doc.SendCommand(f'OPEN {dwg}\n ')
#acad.doc.SendCommand('FILEDIA 1 ') #reset default

#wrapper function to cleanup process step.
def send_wait(cmd, limit=ratelimit):
	doc.SendCommand(cmd)
	time.sleep(limit)

def process(doc):
	#send_wait('SELECT ALL *\n BURST ') -> Stopped working at some point, Swapped to _ai_selall
	send_wait('_ai_selall\nBURST \n')
	send_wait('_ai_selall\nBURST \n')
	send_wait('_ai_selall\nBURST \n') # 99% of elements are exploded by now.
	
	#send_wait('-PURGE ALL *\nN\n ')  -> Moved to dwg trueview because DWG Convert is many times faster. 
	#send_wait('AUDIT Y ')

# main program logic.
for doc in acad.app.Documents:
	print(doc.Name, end="")
	process(doc)
	print(" -> Done")
send_wait('saveall ')
send_wait('close_all ')

import googlemaps, json, io, jdcal, openpyxl, os, sys
from openpyxl.styles import Alignment
from openpyxl import Workbook
from HTMLParser import HTMLParser
from datetime import datetime
import Tkinter, Tkconstants, tkFileDialog, os
from Tkinter import *

class Route(object):
	routeName = ""
	routeStops = []
	
	def __init__(self, routeName):
		self.routeName = routeName

class RouteStop(object):
	number = ""
	clientName = ""
	address = ""
	phone = ""
	hotMeals = 0
	coldMeals = 0
	deliveryInstructions = ""
	specialInstructions = ""
	sunday = False
	monday = False
	tuesday = False
	wednesday = False
	thursday = False
	friday = False
	saturday = False
	
    # The class "constructor" - It's actually an initializer 
	def __init__(self, number, clientName, address, phone):
		self.number = number
		self.clientName = clientName
		self.phone = phone
		self.address = address

class MLStripper(HTMLParser):
    def __init__(self):
        self.reset()
        self.fed = []
    def handle_starttag(self, tag, attrs):
        if('div' in tag):
            self.fed.append("\n")
    def handle_data(self, d):
        self.fed.append(d)
    def get_data(self):
        return ''.join(self.fed)

def strip_tags(html):
    s = MLStripper()
    s.feed(html)
    return s.get_data()

def compileWorkBooks(workbookName):
	wbRead = openpyxl.load_workbook(workbookName)
	
	routeStopList = []
	routeList = {}
	
	iterSheet = iter(wbRead.active)
	next(iterSheet)
	for row in iterSheet:
		routeStopList.append(RouteStop(row[0].value, row[1].value, row[2].value, row[3].value))
	wbRead.close()
	
	for routestop in routeStopList:
		if routestop.number[0] not in routeList:
			routeList[str(routestop.number[0])]=Route(routestop.number[0])
		routeList[routestop.number[0]].routeStops.append(routestop)
	for route in routeList:
		print "DETERMINE ROUTING FOR EACH ROUTE AND CREATE SHEET"


class App(Tkinter.Frame):
	def __init__(self, master):
		self.fileCount=0
		Tkinter.Frame.__init__(self, root)
		
		button_opt = {'fill':Tkconstants.BOTH, 'padx':5, 'pady':5}

		self.displayBox = Tkinter.Text()
		self.displayBox.pack()

		self.dir_opt = options = {}
		options['initialdir'] = os.path.dirname(os.path.realpath(sys.argv[0]))
		options['parent'] = root
		options['title'] = 'Select a file to parse'
		
		Tkinter.Button(self, text='Select File to Parse', command=self.askdirectory).pack(**button_opt)
	def askdirectory(self):
		self.fileCount=0
		self.selectedDir = tkFileDialog.askopenfilename(**self.dir_opt)
		self.displayBox.insert(Tkinter.END, self.selectedDir + " selected")
		compileWorkBooks(self.selectedDir)
		
if __name__=='__main__':
	root = Tkinter.Tk()
	App(root).pack()
	root.mainloop()
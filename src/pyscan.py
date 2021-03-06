#!/usr/bin/env python
import sys
from time import time, gmtime, strftime
from os.path import isdir, isfile, join
from os import mkdir, remove, listdir, __file__
import ConfigParser

import twain

from Tkinter import *
from tkMessageBox import askyesno, showinfo

from PIL import ImageTk, Image

COUNT = 0
ANGLE = 0

class Twain:
	def __init__(self, cfg='', onSelect=False):
		self.cfg = cfg

		## Call to dsm_entry, this will load twain32.dll or twain_dsm.dll
		
		try:
			self.SourceManager = twain.SourceManager(0)	
		except Exception as e:
			if 'dll' in e:
				self.SourceManager.destroy()
				self.SourceManager = twain.SourceManager(0)
			else:
				print(e)

		sl = self.SourceManager.GetSourceList()
		if not sl:
			self.message = "Aucun scanner n'est reconnu sur le reseau"
		else:
			self.productName = self.getProductName(onSelect)
			print(self.productName)
		self.capPixelType = self.getPixelType()
		self.capResolution = self.getResolution()

	def setErrs(self, err, method=''):
		self.message = """
		Si vous rencontrez ce message, veuillez vous assurez de la bonne connection de votre peripherique.
		- Debrancher/Rebrancher votre scanner.
		- Selectionner a nouveau votre scanner depuis le menu 'Scanner -> Selectionner Source'

		[file=pyscan.py, class=Twain, method=%s]
		[%s]
		"""%(method,err)

	def getProductName(self, onSelect):
		##	Check if source has been opened once from inifile
		#	if not use selectSource to get twain.SourceManager.ProductName
		if not hasSectionOrOption(self.cfg, needle='ProductName',option=True) or onSelect:
			if self.selectSource():
				writeSetting(self.cfg, self.productName, section='SCANNER_INFO', option='ProductName')
				name = self.Source.GetSourceName()
				self.Source.destroy()
				return name
			else:
				return ''
		else:
			return self.cfg.get('SCANNER_INFO', 'ProductName')

	def selectSource(self):
		##	Will be used once to get identity, then onSourceSelect
		#	Get twain source object
		try:
			self.Source = self.SourceManager.OpenSource()
		except Exception as e:
			if "ConditionCode = 4" in e:
				self.setErrs(e, 'selectSource')
			else:
				print(">>\n%s\n<<"%e)
			return False
		finally:
			if hasattr(self, 'Source'):
				self.productName = self.Source.GetSourceName()
				self.Source.destroy()
				return True

	def getInfo(self):
		source = self.SourceManager.OpenSource(self.productName)
		
		cap_PIXELTYPE, value = source.get_capability_current(twain.ICAP_PIXELTYPE) 
		message = "%s: %s %s"% ('PIXELTYPE', cap_PIXELTYPE, value)
		
		source.destroy()
		self.SourceManager.destroy()
		return (message)

	def getPixelType(self):
		if hasSectionOrOption(self.cfg, needle='PIXELTYPE', option=True):
			return self.cfg.get('SCANNER_INFO', 'PIXELTYPE')
		else:
			return 'bw'

	def getResolution(self):
		if hasSectionOrOption(self.cfg, needle='RESOLUTION', option=True):
			return self.cfg.get('SCANNER_INFO', 'RESOLUTION')
		else:		
			default = 'good'	
			writeSetting(self.cfg, default, 'SCANNER_INFO', 'RESOLUTION')
			return default

	def Acquire(self,filename):
		try:
			self.Source = self.SourceManager.OpenSource(self.productName)
		except Exception as e:
			if "ConditionCode = 4" in e:
				# Cancel Acquire !
				self.setErrs(e, 'Acquire')
			else:
				print(">>\n%s\n<<"%e)
			return False

		# Set color default: bw
		pixelTypeMap = {
		'bw': twain.TWPT_BW,
		'grey': twain.TWPT_GRAY,
		'color': twain.TWPT_RGB
		}
		pixelType = pixelTypeMap[self.capPixelType]
		self.Source.SetCapability(twain.ICAP_PIXELTYPE, twain.TWTY_UINT16, pixelType)

		# Set resolution default: 100
		resolutionMap = {
		'low': 50,
		'good': 100,
		'excellent': 150,
		'fantastic': 200
		}
		print(self.capResolution)
		resolution = resolutionMap[self.capResolution]
		self.Source.SetCapability(twain.ICAP_XRESOLUTION, twain.TWTY_FIX32, resolution)
		self.Source.SetCapability(twain.ICAP_YRESOLUTION, twain.TWTY_FIX32, resolution)

		self.Source.RequestAcquire(0,0)
		self.Source.ModalLoop()
		self.rv = self.Source.XferImageNatively()
		if self.rv:
			(self.handle, count) = self.rv
		self.save(filename)
		return True


	def save(self, filename):
		twain.DIBToBMFile(self.handle, filename)
		self.img = filename
		
		#	Unfortunatly i must destroy source/sm everytime in order to reset twain states
		self.Source.destroy()
		self.SourceManager.destroy()
		

class App(Frame):
	## --- This is class is used to handle the Tk application --- ##
	## ---------------------------------------------------------- ##
	## -------------- Interface and App management -------------- ##	

	def __init__(self, cfg='', master=None):
		Frame.__init__(self, master)
		print(__file__)
		self.errs = list()
		self.master = master

		self.state = False
		self.master.bind("<F11>", self.toggleFullscreen)		
		self.master.bind("q", self._quit)
		
		self.imgList = list()
		self.editBinded = False

		self.cfg = cfg
		## App geometry

		# getScreenSize and update setAppSize
		if hasSectionOrOption(self.cfg, needle='maxAppWidth', option=True) and hasSectionOrOption(self.cfg, needle='maxAppHeight', option=True):
			screenWidth, screenHeight = self.cfg.getint('APP_INFO', 'maxAppWidth'), self.cfg.getint('APP_INFO', 'maxAppHeight')
			self.maxAppSize = (screenWidth,screenHeight)
		else:
			screenWidth, screenHeight = self.master.winfo_screenwidth(),(self.master.winfo_screenheight()-5)
			writeSetting(self.cfg, str(screenWidth), 'APP_INFO', 'maxAppWidth')
			writeSetting(self.cfg, str(screenHeight), 'APP_INFO', 'maxAppHeight')
			self.maxAppSize = (screenWidth,screenHeight)

		self.minAppSize = (960,860)
		self.master.geometry("%dx%d+0+0" % (self.minAppSize))

		# isFullScreen
		if hasSectionOrOption(self.cfg, needle='LastClosedFullScreen', option=True):
			self.forceFullscreen(
				value=self.cfg.getboolean('APP_INFO','LastClosedFullScreen')
				)
		else:
			self.errs.append({'hasSectionOrOption=No Section Found': 'continue'})

		# hasWDReady
		errs, ready = hasWDReady()
		if not ready:
			for err in errs:			
				self.errs.append({err: 'close'})

		# hasAtLeastOneSource
		self.hasASource = False
		if hasSectionOrOption(self.cfg, needle='ProductName', option=True):
			self.scannerName = self.cfg.get('SCANNER_INFO', 'ProductName')
			self.hasASource = True
		else:
			if self.selectSource():
				self.hasASource = True

		if self.hasASource:
			self.master.bind('n', self.TwainAcquire)

		self.createWidgets()

	def __destroy__(self):
		self._quit()

	def confirm(self):
		ans = askyesno(title='Quitter', message='Etes vous sure de vouloir quitter ?')
		if ans:
			self.quit()

	def _quit(self, event=None):
		writeSetting(self.cfg, str(self.state), 'APP_INFO', 'LastClosedFullScreen')
		self.confirm()

	def createWidgets(self):
		# -------------- Labels --------------- #

		self.labelSourceName = Button(self.master,width=100)
		self.labelSourceName["text"] = self.scannerName

		self._imgHolder = Canvas(self.master, width=600,height=800)
		self.selectedItem = Label(self.master)

		self.labelHelp = Label(self.master, text=
			"<N> Nouvelle Acquisition; <A> Appliquer; <R> Retourner a gauche; <Shift-R> Retourner a droite; <Q> Quitter"
			)

		# -------------- MenuBar -------------- #

		self.menubar = Menu(self.master)
		self.master.config(menu=self.menubar)

		self.fileMenu = Menu(self.menubar, tearoff=0)
		self.fileMenu.add_command(
			label="Quitter",
			command=self._quit,
			underline=0
			)
		
		self.menubar.add_cascade(label="Fichier", menu=self.fileMenu)

		self.scanMenu = Menu(self.menubar, tearoff=0)
		self.scanMenu.add_command(
			label="Selectionner source",
			command=self.selectSource
			)
		self.scanMenu.add_command(
			label = "Nouvelle numerisation",
			command = self.TwainAcquire,
			state=(DISABLED if self.scannerName == '' else NORMAL)
			)
		# self.scanMenu.add_command(
		# 	label = "About(Scanner)",
		# 	command = self.showAbout,
		# 	state=(DISABLED if self.scannerName == '' else NORMAL)
		# 	)

		self.menubar.add_cascade(label="Scanner", menu=self.scanMenu)

		self.pixelTypeMenu = Menu(self.scanMenu, tearoff=0)
		self.pixelTypeMenu.add_command(
			label = "Noir et blanc",
			command = lambda: self.setPixelType('bw'),
			state=(DISABLED if self.scannerName == '' else NORMAL)
			)
		self.pixelTypeMenu.add_command(
			label = "Gris",
			command = lambda: self.setPixelType('grey'),
			state=(DISABLED if self.scannerName == '' else NORMAL)
			)
		self.pixelTypeMenu.add_command(
			label = "Couleur",
			command = lambda: self.setPixelType('color'),
			state=(DISABLED if self.scannerName == '' else NORMAL)
			)

		self.scanMenu.add_cascade(label="Mode de couleur", menu=self.pixelTypeMenu)		

		self.resolutionMenu = Menu(self.scanMenu, tearoff=0)
		self.resolutionMenu.add_command(
			label = "Faible (50 dpi)",
			command = lambda: self.setResolution('low'),
			state=(DISABLED if self.scannerName == '' else NORMAL)
			)
		self.resolutionMenu.add_command(
			label = "Bonne (100 dpi)",
			command = lambda: self.setResolution('good'),
			state=(DISABLED if self.scannerName == '' else NORMAL)
			)
		self.resolutionMenu.add_command(
			label = "Excellente (150 dpi)",
			command = lambda: self.setResolution('excellent'),
			state=(DISABLED if self.scannerName == '' else NORMAL)
			)
		self.resolutionMenu.add_command(
			label = "Fantastique (200dpi)",
			command = lambda: self.setResolution('fantastic'),
			state=(DISABLED if self.scannerName == '' else NORMAL)
			)		

		self.scanMenu.add_cascade(label="Resolution", menu=self.resolutionMenu)

		self.editMenu = Menu(self.menubar, tearoff=0)
		self.editMenu.add_command(
			label = 'Appliquer',
			command = self.applyChangesImage,
			state=DISABLED
			)
		self.editMenu.add_command(
			label="Reinitialiser canvas",
			command=self.resetCan,
			state=DISABLED
			)		
		self.editMenu.add_command(
			label = 'Retourner a GAUCHE',
			command = lambda: self.rotateCan(90),
			state=DISABLED
			)
		self.editMenu.add_command(
			label = 'Retourner a DROITE',
			command = lambda: self.rotateCan(-90),
			state=DISABLED
			)
		self.editMenu.add_command(
			label = 'Supprimer',
			command = self.delImage,
			state=DISABLED
			)
		
		self.menubar.add_cascade(label="Editer", menu=self.editMenu)

		# -------------- listBox -------------- #
		self.listbox = Listbox(self.master, selectmode=SINGLE, height=COUNT+1, width=15)
		def callback(event):
			sel = event.widget.curselection()
			if sel:
				i = sel[0]
				data = self.listbox.get(self.listbox.curselection())
				self.selectedItem.configure(text=data)
				self.resetCan()
				
				self.DisplayImage(data)
			else:
				self.selectedItem.configure(text='')
		self.listbox.bind("<<ListboxSelect>>", callback)		

		self.displayWidgets()

	def displayWidgets(self, event=None, _from=''):
		## Let's be logic..
		# Say hello (Selected Source OR onClickSelectSource update ini and Label/infos here)
		# ALWAYS, state['helloed']
		self.master.update()
		padx = (int(self.master.winfo_width()) - 865)*.5
		if padx <= 0:
			w,h = self.minAppSize
			padx = int(w - 715)*.5
		self.labelSourceName.grid(row=0,column=1,pady=5,padx=padx)

		# Place canvas on the grid
		if self._imgHolder:
			self._imgHolder.grid(row=1,column=1)

		self.labelHelp.grid(row=2,column=1)
		
		self.listbox.grid(column=0,row=1, rowspan=1)

	def updateWidgets(self):
		# update list from app property
		self.listbox.insert(COUNT,self.img)
		self.listbox.configure(height=COUNT+1)
		self.listbox.update()

	def updateSourceLabel(self, event=None):
		if scannerName:
			self.labelSourceName.configure(text=self.scannerName)

	def toggleFullscreen(self, event=None):
		self.state = not self.state
		self.master.attributes("-fullscreen", self.state)
		return "break"

	def forceFullscreen(self, value):
		self.master.attributes("-fullscreen", value)


	## -------------- \Interface and App management -------------- ##
	## ----------------------------------------------------------- ##
	## ---------------------- Twain Related ---------------------- ##


	def selectSource(self, event=None):
		# ... :s
		tw = Twain(cfg=self.cfg, onSelect=True)
		name = tw.productName
		self.scannerName = name
		del tw
		return True if not self.scannerName == '' else False

	def setPixelType(self, value):
		writeSetting(self.cfg, value, 'SCANNER_INFO', 'PIXELTYPE')

	def setResolution(self, value):
		writeSetting(self.cfg, value, 'SCANNER_INFO', 'RESOLUTION')

	def TwainAcquire(self, event=None):
		# Empty canvas
		if hasattr(self, '_imgHolder'):
			self.resetCan()

		# ... :s
		tw = Twain(cfg=self.cfg)
		if tw.Acquire('tmp/out_%d.png' % COUNT):	
			self.DisplayImage('tmp/out_%d.png' % COUNT)
		else:
			showinfo(title='error',message=tw.message)
		del tw

	def showAbout(self):
		tw = Twain(cfg=self.cfg)
		message = tw.getInfo()
		del tw
		showinfo(title='About',message=message)


	## --------------------- \Twain Related ---------------------- ##
	## ----------------------------------------------------------- ##
	## --------------------- Image management -------------------- ##
	

	def DisplayImage(self, _img, angle=ANGLE):
		img = Image.open(_img)
		img = img.resize((600,800), Image.ANTIALIAS)
		rotated = img.rotate(ANGLE)
		img = ImageTk.PhotoImage(rotated)
		self.rotated = rotated
		self._imgHolder.create_image(0,0, image=img, anchor='nw')
		self._imgHolder.image = img
		self._imgHolder.grid(row=1,column=1)
		self.img = _img

		# Bind and activate edit menu
		for i in range(0,5):
			self.editMenu.entryconfig(i, state=NORMAL)		
		if not self.editBinded:
			self.applyB_ID = self.master.bind('a', self.applyChangesImage, "+")
			self.rotateLeftB_ID = self.master.bind('r', lambda event, r=90: self.rotateCan(r), '+')
			self.preciseRotateLeftB_ID = self.master.bind('<Control-r>', lambda event, r=1: self.rotateCan(r), '+')
			self.rotateRightB_ID = self.master.bind('<Shift-R>', lambda event, r=-90: self.rotateCan(r), '+')
			self.preciseRotateRightB_ID = self.master.bind('<Control-Shift-r>', lambda event, r=-1: self.rotateCan(r), '+')
			self.deleteB_ID = self.master.bind('d', self.delImage, "+")
			self.editBinded = True

	def resetCan(self):
		global ANGLE
		self._imgHolder.image = ''
		self.img = ''
		ANGLE = 0

		# Unbind and deactivate edit menu
		for i in range(0,5):
			self.editMenu.entryconfig(i, state=DISABLED)		
		if self.editBinded:
			self.master.unbind('a', self.applyB_ID)
			self.master.unbind('r', self.rotateLeftB_ID)
			self.master.unbind('<Control-r>', self.preciseRotateLeftB_ID)
			self.master.unbind('<Shift-R>', self.rotateRightB_ID)
			self.master.unbind('<Control-Shift-r>', self.preciseRotateRightB_ID)
			self.master.unbind('d', self.deleteB_ID)
			self.editBinded = False		

	def rotateCan(self, angle):
		global ANGLE

		ANGLE += angle
		self.DisplayImage(self.img, angle=ANGLE)

	def applyChangesImage(self, event=None):
		global COUNT

		dt = strftime("%Y-%m-%d_%H-%M-%S")
		fn = join('images', ("%s_%d.png" % (dt, COUNT)))
		self.rotated.save(fn)
		
		self.img = fn

		self.imgList.append(self.img)
		self.updateWidgets()
		COUNT += 1

	def delImage(self, event=None):
		if hasattr(self._imgHolder, 'image'):
			if not self._imgHolder.image == '':
				if askyesno(title='Supprimer Image', message="Etes vous sure de vouloir supprimer l'image ?"):				 
					remove(self.img)
					# Get id of our object in the listbox to delete it
					# BUT NOT from listbox.getindex() or such
					who = self.listbox.get(0,END).index(self.img)
					self.listbox.delete(who)
					self.resetCan()
					self.updateWidgets()


	## -------------------- \Image management -------------------- ##
	## ----------------------------------------------------------- ##
	## ------------------------- Program ------------------------- ##
	

def writeSetting(cfg,value, section='',option=''):
	try:
		cfg.set(section, option, value)
	except ConfigParser.NoSectionError:
		cfg.add_section(section)
		cfg.set(section, option, value)
	finally:
		with open('config.ini', 'w') as fd:
			cfg.write(fd)

def hasSectionOrOption(cfg, needle='', section=False,option=False):	
	if section:
		return (True if cfg.has_section(needle) else False)
	elif option:
		for sec in cfg.sections():
			if cfg.has_option(sec, needle):
				return True
		return False

def hasWDReady():
	errs = list()
	# images will be used as return path for the app
	dirs = ['../tmp', '../images']	
	for _dir in dirs:
		if not isdir(_dir):
			fn = '%s/init.png' % _dir
			try:
				mkdir(_dir, 0o666)
				# Make sure we can read/write with this user
				with open(fn,'w') as fd:
					fd.write('%s created'% _dir)

			except OSError as e:
				errs.append(e)
				print(e)

			finally:
				if isfile(fn):
					remove(fn)
		else:
			## ALWAYS clear directory on __init__ !!
			# This is the container for our result, so keep that as clean as possible
			
			for f in listdir(_dir):
				try:
					remove(join(_dir,f))
				except OSError as e:
					errs.append(e)
					print(e)
				finally:
					print('Clean Directory')

	return (errs, (True if len(errs) == 0 else False))


def main():
	root = Tk()

	cfg = ConfigParser.SafeConfigParser()
	fn = "config.ini"
	cfg.read(fn)
	
	if not cfg:
		writeSetting(cfg, 'True', 'APP_INFO', 'FirstLaunch')

	for c in ['APP_INFO', 'SCANNER_INFO']:
		if not hasSectionOrOption(cfg, needle=c,section=True):
			cfg.add_section(c)		

	app = App(cfg=cfg,master=root)

	## Let's run the app
	app.master.title('Scanner')
	app.mainloop()

	## Log last closed, for later uses
	writeSetting(cfg, str(time()), 'APP_INFO', 'LastClosed')
	root.destroy()

main()
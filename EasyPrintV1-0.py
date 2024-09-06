from tkinter import *
from tkinter import ttk
from pyautocad import *
from comtypes import automation
import win32com
import win32com.client
import pythoncom
import time
import array
from ctypes import *

cads = Autocad(create_if_not_exists=True, visible = True)

acad = win32com.client.Dispatch("AutoCAD.Application")
xl = win32com.client.gencache.EnsureDispatch ("AutoCAD.Application")
doc = acad.ActiveDocument
ms = doc.ModelSpace

try:
    print('File {} connected.' .format(doc.Name))
except:
    print('AutoCAD in use.!!!')
    print('Press ESC on AutoCAD window then try again.')
    quit(1)
cads2 = cads.ActiveDocument
cads3 = cads2.ModelSpace

def layerexist(lay):
    layers = doc.Layers
    layers_nums = layers.Count
    layers_names = [layers.Item(i).Name for i in range(layers_nums)]    # List of ACAD layers
    if lay in layers_names:
        return True
    else:
        return False

def get_selection(self, text="Select objects"):
        """ Asks user to select objects

        :param text: prompt for selection
        """
        self.prompt(text)
        try:
            self.doc.SelectionSets.Item("SS1").Delete()
        except Exception:
            logger.debug('Delete selection failed')

        selection = self.doc.SelectionSets.Add('SS1')
        selection.SelectOnScreen()
        return selection

def GetBoundingBox(entity):
    # Create 3-d 'Variant' array of 'd'- ouble
    A = automation.VARIANT(array.array('d', [0,0,0]))
    B = automation.VARIANT(array.array('d', [0,0,0]))
    
    # Get the refernece / address
    vA = byref(A)
    vB = byref(B)
    
    # Call the Method from COM object
    entity.GetBoundingBox(vA,vB)
    # Return two points as 3-d
    return (A.value, B.value)

def VtFloat(list):
    """converts list in python into required float"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list)

def blockPrint(pt1,pt2):
    doc.ActiveLayout.ConfigName = str(printer_option.get())
    lotstyle = ms.Layout.GetCanonicalMediaNames()
    plotstyle = []
    for i in ms.Layout.GetCanonicalMediaNames():
        plotstyle.append(ms.Layout.GetLocaleMediaName(i))
    ind = plotstyle.index(str(paper_option.get()))
    paper_selected = (lotstyle[ind])
    ms.Layout.CanonicalMediaName = str(paper_selected)
    ms.Layout.StyleSheet = str(plotstyle_option.get())
    ms.Layout.SetWindowToPlot(VtFloat(pt1[0:2]),VtFloat(pt2[0:2]))
    ms.Layout.PlotType=win32com.client.constants.acWindow
    ms.Layout.CenterPlot = True
    ms.Layout.UseStandardScale = True
    ms.Layout.PlotRotation = win32com.client.constants.ac90degrees
    ms.Layout.StandardScale = win32com.client.constants.acScaleToFit
    doc.Plot.PlotToDevice()

def addboundbox(points_2d,lay):
    if not layerexist(lay):
        cads2.Layers.Add(lay)
    box = cads3.AddPolyLine(array.array("d",points_2d))
    box.Layer = lay
    return box


def printbutton():
    ss_se = []
    for se in cads3:
        if se.Entityname == 'AcDbBlockReference':
            if se.Name == ssName:
                ss_se.append(se)
                
    for i in ss_se:
        point1,point2 = GetBoundingBox(i)
        pt1 = []
        pt2 = []
        for x in point1:
            pt1.append(x)
        for y in point2:
            pt2.append(y)
        print(pt1)
        print(pt2)
        points = [pt1[0],pt1[1],0,pt1[0],pt2[1],0,pt2[0],pt2[1],0,pt2[0],pt1[1],0,pt1[0],pt1[1],0]
        layer = "EasyPrint"
        addboundbox(points,layer)
        blockPrint(pt1,pt2)
        
def previewbutton():
    return doc.Plot.DisplayPlotPreview(win32com.client.constants.acFullPreview)

root = Tk()
root.title('Easy Print')
#root.iconbitmap('abs.ico')
root.geometry("350x350")

def pick_Block():
    root.iconify()
    ss1 = cads.get_selection()
    global ssName
    ssName = ss1.Item(0).Name
    block_label = Label(PrintMethod_frame,text= ssName)
    block_label.grid(row=0, column=2,sticky = W)
    return ssName

# Creating Frames
################################################################
# 1. Name Frame:
drgName_frame = LabelFrame(root, text = "Drawing Name",padx=5,pady=5)
drgName_frame.grid(row=0,column=0,ipadx=10)

drgName_label = Label(drgName_frame,text= doc.Name)
drgName_label.grid(row=0, column=0,sticky = W)
#################################################################
# 2. Setting Frame
setting_frame = LabelFrame(root, text = "Settings",padx=2,pady=2)
setting_frame.grid(row=1,column=0,ipadx=10,sticky =W)

printer_label = Label(setting_frame,text= "Printer").grid(row=0, column=0,sticky = W)
printers = ms.Layout.GetPlotDeviceNames()
printer_option = ttk.Combobox(setting_frame,width = 40,value = printers)#,textvariable = printers\)
printer_option.grid(row=0,column=1,sticky = E,pady=5)

def papernames(e):
    doc.ActiveLayout.ConfigName = str(printer_option.get())
    papers = []
    papers.clear()
    for i in ms.Layout.GetCanonicalMediaNames():
        papers.append(ms.Layout.GetLocaleMediaName(i))
        paper_option.config(value = papers)
    return papers
printer_option.bind("<<ComboboxSelected>>",papernames)

selectedprinter = printer_option.get()

#################################################################

paper_size_label = Label(setting_frame,text= "Paper Size")
paper_size_label.grid(row=2, column=0,sticky = W)

paper_option = ttk.Combobox(setting_frame,width = 40)#,postcommand = papernames())#,value = papers,textvariable = papers)
paper_option.grid(row=2,column=1,sticky = E,pady=5)

#################################################################

plot_style_label = Label(setting_frame,text= "Plot Style")
plot_style_label.grid(row=4, column=0,sticky = W)
plotstyles = ms.Layout.GetPlotStyleTableNames()
plotstyle_option = ttk.Combobox(setting_frame,value = plotstyles)#,textvariable=plotstyles
plotstyle_option.grid(row=4,column=1,sticky = W,ipadx = 8,pady=5)

#################################################################

# 2. Print Method Frame
PrintMethod_frame = LabelFrame(root, text = "Print Method")
PrintMethod_frame.grid(row=2,column=0,ipadx=25,sticky =W)
Pick_button = Button(PrintMethod_frame, text="Pick Block", command=pick_Block)
Pick_button.grid(row=0,column=1,padx=10,pady=10)

#################################################################

# 7. Common frame for Buttons
common_frame = LabelFrame(root, text = "")
common_frame.grid(padx=10,sticky =W,pady=10)
Print_button = Button(common_frame, text="Print",padx=15,command = printbutton)
Print_button.grid(row=1,column=1,padx=5)
Preview_button = Button(common_frame, text="Preview",padx=15,command = previewbutton)
Preview_button.grid(row=1, column=2,padx=5)
Cancel_button = Button(common_frame, text="Cancel",padx=15,command = root.destroy)
Cancel_button.grid(row=1,column =3,padx=5)
Help_button = Button(common_frame, text="Help",padx=15)
Help_button.grid(row=1,column=4,padx=5)

root.mainloop()
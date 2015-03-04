from win32com.client import Dispatch
excel = Dispatch("Excel.Application")
excel.Visible = True
XlDirectionDown = 4  #Not needed yet
excel.DisplayAlerts = False
excel.Application.CutCopyMode=False

#Stage 1: Open the Internal Machine and clear sheets
machine = excel.Workbooks.Add("C:\\Users\\A_DO\\Dropbox\\2. The Machine\\Internal Machine\\The Bible rev110.xlsx")
#This is where u tell python what u want to label your specific excel file and where it's at
excel.Worksheets("Dom VOD (rep)").Select() #this is where you specify the tab you want to go to
excel.Range("A2:AR40000").Select() #This is where u specify the range u want to select
excel.Selection.ClearContents() 

excel.Worksheets("UK VOD (rep)").Select() 
excel.Range("A2:AR40000").Select() 
excel.Selection.ClearContents() 

excel.Worksheets("Dom LINEAR PB, this Mo (rep)").Select() 
excel.Range("A2:P40000").Select() 
excel.Selection.ClearContents()

excel.Worksheets("Dom LINEAR Hots, this Mo (rep)").Select() 
excel.Range("A3:Q40000").Select() 
excel.Selection.ClearContents()

excel.Worksheets("RKTV, this Mo (rep)").Select()
excel.Range("A2:Q400").Select() 
excel.Selection.ClearContents()

excel.Worksheets("UK LINEAR (rep)").Select()
excel.Range("A2:Q40000").Select() 
excel.Selection.ClearContents()

excel.Worksheets("EDIT Locating Media (rep)").Select()
excel.Range("A2:AF40000").Select()
excel.Selection.ClearContents()
 
excel.Worksheets("Schedules (rep)").Select()
excel.Range("A2:AF40000").Select()
excel.Selection.ClearContents()

excel.Worksheets("CO (rep)").Select()
excel.Range("A2:AF40000").Select()
excel.Selection.ClearContents()

excel.Worksheets("EDIT Complete (rep)").Select()
excel.Range("A2:AF40000").Select()
excel.Selection.ClearContents()

excel.Worksheets("QC Media Services Tracker (rep)").Select()
excel.Range("A2:AM40000").Select()
excel.Selection.ClearContents()

excel.Worksheets("QC3 (rep)").Select()
excel.Range("A2:AM40000").Select()
excel.Selection.ClearContents()

print ("Stage 1 Complete: Open the Internal Machine and clear sheets")

#Stage 2 domesticVOD, Open, Copy, Paste, Close
domesticVOD = excel.Workbooks.Add("\\\\10.3.65.159\\MediaServices\\Fulfillment_Tracking\\Domestic_VOD_Fulfillment\\VOD_Tracker_ALL_rev5.xlsx")
excel.Worksheets("Tracker").Select()
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A2:AR40000").Select()
excel.Selection.Copy()
machine.Activate()
excel.Worksheets("Dom VOD (rep)").Select() #this is where you specify the tab you want to go to
excel.Range("A2").Select()
excel.Selection.PasteSpecial(Paste=-4163)
domesticVOD.Close()

print ("Stage 2 Complete: domesticVOD")

#Stage 3 ukVOD, Open, Copy, Paste, Close
ukVOD = excel.Workbooks.Add("\\\\10.3.65.159\\MediaServices\\Fulfillment_Tracking\\UK_Fulfillment\\VOD\\UK VOD Tracker 2014.02.27.xlsx")
excel.Worksheets(1).Select()
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A2:AR40000").Select()
excel.Selection.Copy()
machine.Activate()
excel.Worksheets("UK VOD (rep)").Select()
excel.Range("A2").Select()
excel.Selection.PasteSpecial(Paste=-4163)
ukVOD.Close()

print ("Stage 3 Complete: ukVOD")

#Stage 4 domLinearPB Open, Copy, Paste, Close
domLinearPB1 = excel.Workbooks.Add("\\\\10.3.65.159\\MediaServices\\Fulfillment_Tracking\\Domestic_Linear_Fulfillment\\2015_03_Mar PBTV.xlsx")
excel.Worksheets(1).Select()
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A2:R200").Select()
excel.Selection.Copy()

machine.Activate()
excel.Worksheets("Dom LINEAR PB, this Mo (rep)").Select() 
excel.Range("A2").Select()
excel.Selection.PasteSpecial(Paste=-4163)
domLinearPB1.Close()

print ("Stage 4 Complete: domLinearPB")

#Stage 5 domLinearHots and RKTV, Open, Copy, Paste, Close
excel.AskToUpdateLinks = False #supposed to disable the update links popup
domLinearHots1 = excel.Workbooks.Add("\\\\10.3.65.159\\MediaServices\\Fulfillment_Tracking\\Domestic_Linear_Fulfillment\\2015_03_Mar_Linear_Hots_Fulfillment-dev.xlsx")
excel.Worksheets(1).Select()
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A3:Q200").Select()
excel.Selection.Copy()
machine.Activate()
excel.Worksheets("Dom LINEAR Hots, this Mo (rep)").Select()
excel.Range("A3").Select()
excel.Selection.PasteSpecial(Paste=-4163)

domLinearHots1.Activate()
excel.Worksheets("RKTV").Select() #RKTV part
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A2:Q200").Select()
excel.Selection.Copy()
machine.Activate()
excel.Worksheets("RKTV, this Mo (rep)").Select()
excel.Range("A2").Select()
excel.Selection.PasteSpecial(Paste=-4163)
domLinearHots1.Close()

print ("Stage 5 Complete: domLinearHots and RKTV")

#Stage 6 UK Linear, Open, Copy, Paste, Close
ukLinear = excel.Workbooks.Add("\\\\10.3.65.159\\MediaServices\\Fulfillment_Tracking\\UK_Fulfillment\\Linear\\UK_Linear_2014.xlsx")
excel.Worksheets(1).Select()
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A2:Q40000").Select() 
excel.Selection.Copy()
machine.Activate()
excel.Worksheets("UK LINEAR (rep)").Select()
excel.Range("A2").Select()
excel.Selection.PasteSpecial(Paste=-4163)
ukLinear.Close()

print ("Stage 6 Complete: ukLinear")

#Stage 7 QC3, Open, Copy, Paste, Close
qc3 = excel.Workbooks.Add("\\\\10.3.65.159\\MediaServices\\QC\\QC3\\QC3_Log_2015_2_26.xlsx")
excel.Worksheets("Main Data").Select()
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A2:AM40000").Select()
excel.Selection.Copy()
machine.Activate()
excel.Worksheets("QC3 (rep)").Select()
excel.Range("A2").Select()
excel.Selection.PasteSpecial(Paste=-4163)
qc3.Close()

print ("Stage 7 Complete: QC3")

#Stage 8 Edit Google Doc, Open, Copy, Paste, Close
editGoogle = excel.Workbooks.Add("C:\\Users\\A_DO\\Dropbox\\2. The Machine\\10. Edit Tracker\\EDIT.xlsx")

#Locating Media Tab
excel.Worksheets("1. Locating_Media_2015").Select() 
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A2:AF500").Select() 
excel.Selection.Copy()
machine.Activate()
excel.Worksheets("EDIT Locating Media (rep)").Select() 
excel.Range("A2").Select()
excel.Selection.PasteSpecial(Paste=-4163)

#Scheduling Tab
editGoogle.Activate() 
excel.Worksheets("2.Schedules").Select()
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A2:AF500").Select() 
excel.Selection.Copy()
machine.Activate()
excel.Worksheets("Schedules (rep)").Select() 
excel.Range("A2").Select()
excel.Selection.PasteSpecial(Paste=-4163)

#Content Orders Tab
editGoogle.Activate() 
excel.Worksheets(3).Select() #1st CO
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A2:AE1500").Select() 
excel.Selection.Copy()
machine.Activate()
excel.Worksheets("CO (rep)").Select()
excel.Range("A2").Select()
excel.Selection.PasteSpecial(Paste=-4163)

editGoogle.Activate() 
excel.Worksheets(4).Select() #2nd CO
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A2:AE1500").Select() 
excel.Selection.Copy()
machine.Activate()
excel.Worksheets("CO (rep)").Select()
excel.Range("A1500").Select()
excel.Selection.PasteSpecial(Paste=-4163)

editGoogle.Activate() 
excel.Worksheets(5).Select() #3rd CO
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A2:AE1500").Select() 
excel.Selection.Copy()
machine.Activate()
excel.Worksheets("CO (rep)").Select()
excel.Range("A3000").Select()
excel.Selection.PasteSpecial(Paste=-4163)

editGoogle.Activate() 
excel.Worksheets(6).Select() #4th CO
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A2:AE1500").Select() 
excel.Selection.Copy()
machine.Activate()
excel.Worksheets("CO (rep)").Select()
excel.Range("A4500").Select()
excel.Selection.PasteSpecial(Paste=-4163)

editGoogle.Activate() 
excel.Worksheets(7).Select() #5th CO
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A2:AE1500").Select() 
excel.Selection.Copy()
machine.Activate()
excel.Worksheets("CO (rep)").Select()
excel.Range("A6000").Select()
excel.Selection.PasteSpecial(Paste=-4163)

#Completed Tab
editGoogle.Activate() 
excel.Worksheets("10.  2015 Complete").Select()
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A2:AF1000").Select() 
excel.Selection.Copy()
machine.Activate()
excel.Worksheets("EDIT Complete (rep)").Select()
excel.Range("A2").Select()
excel.Selection.PasteSpecial(Paste=-4163)

editGoogle.Close() #Last stage of Edit Google Doc

print ("Stage 8 Complete: Edit GoogleDoc")

#Stage 9 QC3 Google Doc, Open, Copy, Paste, Close
qcGoogle = excel.Workbooks.Add("C:\\Users\\A_DO\Dropbox\\2. The Machine\\11. QC1 (Media Services)\\QC Tracker.xlsx")
excel.ActiveWorkbook.ActiveSheet.Columns(1).AutoFilter() #Turn off or on filter. Off more important
excel.Range("A2:AM40000").Select()
excel.Selection.Copy()
machine.Activate()
excel.Worksheets("QC Media Services Tracker (rep)").Select()
excel.Range("A2").Select()
excel.Selection.PasteSpecial(Paste=-4163)
qcGoogle.Close()

print ("Stage 9 Complete: QC GoogleDoc")

#Save
excel.DisplayAlerts = True #for some reason this stops excel from crashing domLinearHots
from datetime import datetime, date
date = "{:%Y.%m.%d }".format(datetime.now())
name = date + "The Bible rev110 (Python).xlsx"
machine.SaveAs("C:\\Users\\A_DO\\Dropbox\\2. The Machine\\Internal Machine\\"+name)

print "Done son!"






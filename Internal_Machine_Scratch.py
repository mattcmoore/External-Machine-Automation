import os
import glob
from win32com.client import Dispatch
excel = Dispatch('Excel.Application')
excel.Visible = True
XlDirectionDown = 4  #Not needed yet
excel.DisplayAlerts = False
excel.Application.CutCopyMode=False
excel.AskToUpdateLinks = False #supposed to disable the update links popup

machine = excel.Workbooks.Add(max(glob.iglob(r'C:\Users\andyd_000\Desktop\Python\Internal Machine Practice\Internal Machine\The Bible rev111*'), key=os.path.getctime))

dom_vod = max(glob.iglob(r'C:\Users\andyd_000\Desktop\Python\Internal Machine Practice\1. Domestic Tracker\*.xlsx'), key=os.path.getctime)
uk_vod = max(glob.iglob(r'C:\Users\andyd_000\Desktop\Python\Internal Machine Practice\2. UK Tracker\*.xlsx'), key=os.path.getctime)
dom_linear_pb = max(glob.iglob(r'C:\Users\andyd_000\Desktop\Python\Internal Machine Practice\9. DOM Linear Tracker\2014_12_December PBTV*'), key=os.path.getctime)
dom_linear_hots = max(glob.iglob(r'C:\Users\andyd_000\Desktop\Python\Internal Machine Practice\9. DOM Linear Tracker\2014_12_Dec_Linear_Hots*'), key=os.path.getctime)
uk_linear = max(glob.iglob(r'C:\Users\andyd_000\Desktop\Python\Internal Machine Practice\9. UK Linear\*.xlsx'), key=os.path.getctime)
qc3_log = max(glob.iglob(r'C:\Users\andyd_000\Desktop\Python\Internal Machine Practice\7. QC3 (Encompass)\Log (From Marcus)\*.xlsx'), key=os.path.getctime)
edit_google = max(glob.iglob(r'C:\Users\andyd_000\Desktop\Python\Internal Machine Practice\10. Edit Tracker\*.xlsx'), key=os.path.getctime)
qc_google = max(glob.iglob(r'C:\Users\andyd_000\Desktop\Python\Internal Machine Practice\11. QC1 (Media Services)\*xlsx'), key=os.path.getctime)

#Stage 1: Open the Internal Machine and clear sheets
machine_sheets = ['Dom VOD (rep)','UK VOD (rep)','Dom LINEAR PB, this Mo (rep)','Dom LINEAR Hots, this Mo (rep)','RKTV, this Mo (rep)','UK LINEAR (rep)','EDIT Locating Media (rep)','Schedules (rep)','CO (rep)','EDIT Complete (rep)','QC Media Services Tracker (rep)','QC3 (rep)']
for sheet in machine_sheets:
        if sheet == ('Dom LINEAR Hots, this Mo (rep)'):
                excel.Worksheets(sheet).Select()
                excel.Range('A3:Q40000').Select()
                excel.Selection.ClearContents()
        else:
                excel.Worksheets(sheet).Select() #this is where you specify the tab you want to go to
                excel.Range('A2:AR40000').Select() #This is where u specify the range u want to select
                excel.Selection.ClearContents()


#Stage 2: Copy and Paste trackers to specific sheets
sheet_to_trackers = {dom_vod : 'Dom VOD (rep)',uk_vod:'UK VOD (rep)',dom_linear_pb : 'Dom LINEAR PB, this Mo (rep)',dom_linear_hots:'Dom LINEAR Hots, this Mo (rep)',uk_linear:'UK LINEAR (rep)',qc3_log: 'QC3 (rep)',qc_google:'QC Media Services Tracker (rep)'}

trackers =[dom_vod, uk_vod, dom_linear_pb,dom_linear_hots,uk_linear,qc3_log,qc_google]
for tracker in trackers:
    if tracker == dom_linear_hots:
        excel.Workbooks.Add(tracker)
        excel.Worksheets(1).Select()
        excel.ActiveWorkbook.ActiveSheet.AutoFilterMode=False
        excel.Range('A3:AR40000').Copy()
        machine.Activate()
        excel.Worksheets(sheet_to_trackers[tracker]).Select()
        excel.Range('A3').PasteSpecial(Paste=-4163)
    elif tracker == dom_vod:
        excel.Workbooks.Add(tracker)
        excel.Worksheets(3).Select()
        excel.ActiveWorkbook.ActiveSheet.AutoFilterMode=False
        excel.Range('A2:AR40000').Copy()
        machine.Activate()
        excel.Worksheets(sheet_to_trackers[tracker]).Select()
        excel.Range('A2').PasteSpecial(Paste=-4163)
    else:
        excel.Workbooks.Add(tracker)
        excel.Worksheets(1).Select()
        excel.ActiveWorkbook.ActiveSheet.AutoFilterMode=False
        excel.Range('A2:AR40000').Copy()
        machine.Activate()
        excel.Worksheets(sheet_to_trackers[tracker]).Select()
        excel.Range('A2').PasteSpecial(Paste=-4163)

wb = excel.Workbooks.Add()

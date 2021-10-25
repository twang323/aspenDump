import sys
sys.path = [p for p in sys.path if p.startswith(r'C:\Users')]
#sys.path.append('C:\\Users\\username\\AppData\\Local\\Continuum\\miniconda2\\Lib\\site-packages\\win32com')
print(sys.path)
import PySimpleGUI as sg
import os

##def WriteSubject(event):
##    if event == 'Submit':
##        return True
##    else:
##        return None
    
#form = sg.FlexForm('Simple data entry form')  # begin with a blank form
 
layout = [[sg.Text('Browse Short Circuit Template')],
          [sg.FileBrowse(),sg.InputText('', key='_ZcZnZ_')],
          [sg.Text('County of INTCON', size=(15, 1)), sg.InputText('', key='_county_')], 
          [sg.Text('INR', size=(15, 1)), sg.InputText('INR', key='_INR_')], 
          [sg.Text('Gen LLC', size=(15, 1)), sg.InputText('', key='_company_')],
          [sg.Text('Gen MW', size=(15, 1)), sg.InputText('', key='_genMW_')],
          [sg.Text('Gen Type', size=(15, 1)), sg.InputCombo((None,'Wind', 'Solar'),\
                                                   key='_gentype_', default_value=None,size=(20, 1))],
          [sg.Text('Gen Name', size=(15, 1)), sg.InputText('', key='_genname_')],
          [sg.Text('Point of INTCON', size=(15, 1)), sg.InputText('', key='_POI_')],
          [sg.Text('kV of INTCON', size=(15, 1)), sg.InputCombo((None,'345', '138'),\
                                                   key='_genKV_',default_value=None, size=(20, 1))], 
          [sg.Text('ISD', size=(15, 1)), sg.InputText('11/30/20', key='_ISD_')],
          [sg.Text('ASPEN Case Year', size=(15, 1)), sg.InputText('20', key='_caseyear_')],
          [sg.Submit(), sg.Cancel()]]
 
window = sg.Window('Customize Short Circuit Report').Layout(layout)

event, values = window.Read()   
window.Close()
print(event)

from datetime import datetime
current_month = datetime.now().strftime('%B') # February
values['_month_'] = current_month

default_list = ['',None,'None','11/30/20','20','INR']
myDict = {key:val for key, val in values.items() if val not in default_list}
for item in myDict:
    print(item,myDict[item])
#replacedict = {'<county>':values['_county_']}
head, tail = os.path.split(values['_ZcZnZ_'])
templatefile = values['_ZcZnZ_']
templatefile = templatefile.replace('/', '\\')
print('template: ',templatefile)

time1 = datetime.now().strftime('_%H_%M')
savefile = values['_genname_']+\
            '('+values['_INR_']+')_Short Circuit Report'+time1+'.docx'
#savefile = head+'/'+savefile
savefile = head.replace('/', '\\')+'\\'+savefile
print(savefile)

import win32com.client
constants=win32com.client.constants
wordapp=win32com.client.gencache.EnsureDispatch('Word.Application')
worddoc=wordapp.Documents.Open(templatefile)
wordapp.Visible=True
wordapp.Selection.Find.ClearFormatting
wordapp.Selection.Find.Replacement.ClearFormatting
worddoc.Range(0,0).Select()
selection = wordapp.Selection
selection.Find.ClearFormatting()
selection.Find.Replacement.ClearFormatting()
selection.Find.Forward=True
for dummy in myDict.keys():
    selection.Find.Text = dummy
    #selection.Find.Replacement.Text = replacedict[dummy]
    selection.Find.Replacement.Text = myDict[dummy]
    selection.Find.Execute(Replace=constants.wdReplaceAll)

worddoc.TablesOfContents(1).Update
worddoc.SaveAs(savefile)
worddoc.Close()
wordapp.Application.Quit()

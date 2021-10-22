# File:"ASPENDump_TW16.py" using excelpy, python 3.7

import os,sys
sys.path = [p for p in sys.path if p.startswith(r'C:\Users')]
sys.path.append("C:\Program Files (x86)\PTI\PSSE34\PSSPY37")
sys.path.append("C:\Program Files (x86)\PTI\PSSE34\PSSLIB") #psse_utilities
sys.path.append("C:\Program Files (x86)\PTI\PSSE34\PSSBIN") #psspy
os.environ['PATH'] += ';' + "C:\Program Files (x86)\PTI\PSSE34\PSSBIN"

import excelpy #redirect,

def find_first_lines(txtfile):
    'put the outage results into an ascending order'
    infile = open(txtfile)
    lines = [line.strip() for line in infile.readlines()]
    first_row_per_section = [x for x in range(len(lines)) if 'Bus Fault on' in lines[x]]
    index_per_section = [int(lines[x].split('.',1)[0].strip()) for x in range(len(lines)) \
                         if 'Bus Fault on' in lines[x]]

    index_dict = list(zip(index_per_section, first_row_per_section))
    #index_dict.insert(len(index_dict),(1,12))
    index_dict = dict(index_dict) # the index is put in order
    #index_ordered = OrderedDict(sorted(index_dict.items()))
    return lines,index_dict

def create_xlsx_columns(lines,first_lines):
    outage,bus,mva_3ph,mva_1ph,amp_3ph,amp_1ph,z1,z2,z0,xr_3ph,xr_1ph = ([] for i in range(11))
    
    for id in range(len(first_lines)):
        row = first_lines[id+1]        
        if '3LG' in lines[row]:
            is3ph = True
        else:
            is3ph = False
        if 'Branch outage' in lines[row+1]:
            i = 1
        else:
            i = 0        
        if is3ph:
            bus.append(lines[row].split(':',1)[1].split('kV',1)[0].strip() +' kV')
            amp_3ph.append(lines[row+3+i].split('@',4)[3].split()[1].strip())
            z1.append(lines[row+5+i].split(' '*3)[0].strip())
            z2.append(lines[row+5+i].split(' '*3)[1].strip())
            z0.append(lines[row+5+i].split(' '*3)[2].strip())
            mva_3ph.append(lines[row+7+i].split('=',2)[1].split()[0].strip())
            xr_3ph.append(lines[row+7+i].split('=',3)[2].split()[0].strip())
        else:
            amp_1ph.append(lines[row+3+i].split('@',4)[3].split()[1].strip())
            mva_1ph.append(lines[row+7+i].split('=',2)[1].split()[0].strip())
            xr_1ph.append(lines[row+7+i].split('=',3)[2].split()[0].strip())
        if i == 1 and is3ph:
            outage.append(lines[row+1].split(':')[1].strip())
        elif i==0 and is3ph:
            outage.append(None)
    return outage,bus,mva_3ph,mva_1ph,amp_3ph,amp_1ph,xr_3ph,xr_1ph,z1,z0

def write_cell(output,id=1):
    output.set_cell('a'+str(id),'Contingency',fontStyle='bold')
    output.set_cell('b'+str(id),'Bus',fontStyle='bold')
    output.set_cell('c'+str(id),'SC MVA 3ph',fontStyle='bold')
    output.set_cell('d'+str(id),'SC MVA 1ph',fontStyle='bold')
    output.set_cell('e'+str(id),'SC AMP 3ph',fontStyle='bold')
    output.set_cell('f'+str(id),'SC AMP 1ph',fontStyle='bold')
    output.set_cell('g'+str(id),'X/R 3ph',fontStyle='bold')
    output.set_cell('h'+str(id),'X/R 1ph',fontStyle='bold')
    output.set_cell('i'+str(id),'Z1',fontStyle='bold')
    output.set_cell('j'+str(id),'Z0',fontStyle='bold')

    output.set_range(id+1,'a',list(zip(outage))) #one layer list
    output.set_range(id+1,'b',list(zip(bus))) #two layer list
    output.set_range(id+1,'c',list(zip(mva_3ph))) #one layer list
    output.set_range(id+1,'d',list(zip(mva_1ph))) #one layer list
    output.set_range(id+1,'e',list(zip(amp_3ph))) #one layer list
    output.set_range(id+1,'f',list(zip(amp_1ph))) #one layer list
    output.set_range(id+1,'g',list(zip(xr_3ph))) #one layer list
    output.set_range(id+1,'h',list(zip(xr_1ph))) #one layer list
    output.set_range(id+1,'i',list(zip(z1))) #one layer list
    output.set_range(id+1,'j',list(zip(z0))) #one layer list
#--------------------MAIN-----------------------------if __name__ == "__main__":
##xlsx_file = excelpy.workbook(r"ASPEN SC Formatting GUI.xlsx", mode='r')
##txt1 = (xlsx_file.get_cell((3,3))) #'18SSWG_2020_SUM1_Final_06252018_R1.sav'
##txt2 = (xlsx_file.get_cell((4,3)))
##txt = [txt1,txt2]

cwd_txt = []
for fname in os.listdir(os.getcwd()):
    if '.txt' in fname:
        cwd_txt.append(fname)
if len(cwd_txt) == 1:
    user_index = [0]
else:
    flag = False
    while(not flag):
        for i,j in enumerate(cwd_txt):
            print('{}. {}'.format(i,j))
        s = input('ENTER FILE NUMBER(S) NEEDED TO PRODUCE THE REPORT\n(use commas if more than one): ')
        user_index = map(int, s.split(',')) # user_index is an iterator.
        print('THE FILE(S) YOU CHOSE:')
        txt = [cwd_txt[x] for x in user_index]
        print(txt)
        s1=input('PLEASE CONFIRM THE SELECTED FILES Y/N: ')
        if s1 == 'Y' or s1 == 'y':
            flag = True
#print(cwd_txt)
#user_index = [ for x in user_index]
#print(user_index)

#print(txt)

lines,first_lines = find_first_lines(txt[0])
outage,bus,mva_3ph,mva_1ph,amp_3ph,amp_1ph,xr_3ph,xr_1ph,z1,z0=create_xlsx_columns(lines,first_lines)
length_txt1 = len(outage)+1
from datetime import datetime
time1 = datetime.now().strftime('_%H_%M')
FileName = 'output' + time1 + ".xlsx"
print("RESULT OUTPUT >>> "+ FileName)
output = excelpy.workbook(FileName, mode = 'w', sheet = 'None')
output.worksheet_add_after(newSheet = 'Fault Current Report',overwritesheet=1)
output.worksheet_delete('None')
output.set_active_sheet('Fault Current Report')
write_cell(output)
if len(txt)>1:
    lines,first_lines = find_first_lines(txt[1])
    outage,bus,mva_3ph,mva_1ph,amp_3ph,amp_1ph,xr_3ph,xr_1ph,z1,z0=create_xlsx_columns(lines,first_lines)
    write_cell(output,length_txt1+1)

output.save(FileName)
output.close()
print('Done!')

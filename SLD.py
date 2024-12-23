#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      PC
#
# Created:     12/09/2024
# Copyright:   (c) PC 2024
# Licence:     <your licence>
#-------------------------------------------------------------------------------

##res=[1,1,2] ## từng nhánh kết nối với MPT
##print(max(res)) ##max của MPT là số MPT
##        for i in range(mpt):
##            res.append(int(input()))
## creat GSU
import psse35
import psspy
import os
from psspy import _i,_f,_s
ierr=psspy.psseinit(50)
ierr=psspy.newcase_2(basemva=100, basefreq=50)
psspy.newdiagfile()
a = input()
res = [int(digit) for digit in str(a)]
##res = [1, 1, 2, 2 ]
print(res)
##res = [1,2]  ## input thêm điều kiện để loại trường hợp res 1,1,3
mpt = len(res)
len1=[]
len2 = [0] * (max(res) + 1)
dem = [0] * (max(res) + 1)
#mpt=4
psspy.bus_data_4(999997,0,[1,1,1,1],[0.0,1.0,0.0,1.1,0.9,1.1,0.9],r"""DUMMY_POI""")
for i in range(1,mpt+1):
    len1.append(-i*2)
    name = f"GSU{i}_SEC"
    sld =f"BU {10000*i + 1}"
    frombus =10000*i+1
    psspy.bus_data_4(frombus,0,[2,1,1,1],[0.0,1.0,0.0,1.1,0.9,1.1,0.9],name)
##    psspy.growdiagram_2(1,1,[sld],1,-i*2,[0,0,0,0,0,0,0,0,0,0,0,0,0])
    ## Gen
    psspy.plant_data_4(frombus,0,[999997,0],[1,100.0])
    psspy.machine_data_4(frombus,r"""1""",[1,1,0,0,0,1,0],[0.0,0.0,9999.0,-9999.0,9999.0,-9999.0,100.0,0.0,1.0,0.0,0.0,1.0,1.0,1.0,1.0,1.0,1.0],"")
##    psspy.growdiagram_2(1,1,[sld],1,-i*2,[0,0,0,0,0,0,0,0,0,0,0,0,0])

    name1 = f"GSU{i}_PRI"
    sld1 =f"BU {10000+1000*i}"
    tobus =10000+1000*i
    psspy.bus_data_4(tobus,0,[1,1,1,1],[0.0,1.0,0.0,1.1,0.9,1.1,0.9],name1)
    psspy.growdiagram_2(1,1,[sld1],3,-i*2,[0,0,0,0,0,0,0,0,0,0,0,0,0])
    psspy.two_winding_data_6(tobus,frombus,r"""1""",[1,tobus,1,0,0,0,5,0,tobus,0,0,1,0,1,1,1],[0.0,0.0001,100.0,1.0,0.0,0.0,1.0,0.0,1.0,1.0,1.0,1.0,0.0,0.0,1.05,0.95,1.1,0.9,0.0,0.0,0.0],[0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,
    0.0,0.0,0.0],"","")
    psspy.growdiagram_2(1,1,[sld],1,-i*2,[0,0,0,0,0,0,0,0,0,0,0,0,0])
for i in range(len(res)):
    len2[res[i]]=len2[res[i]]+len1[i]
    dem[res[i]]=dem[res[i]]+1
for i in range(1,len(len2)):
    len2[i]=len2[i]/dem[i]
position_110002=-(len(res)+1)
psspy.growdiagram_2(1,1,[r"""BU 999997"""],13,position_110002,[0,0,0,0,0,0,0,0,0,0,0,0,0]) # vi position_110002 
#creat MPT bus
mpt_number=[]
for i in range(1,max(res)+1):
    name = f"MPT{i}_SEC"
    sld =f"BU {100000 + 10000*i}"
    sec =100000 + 10000*i
    psspy.bus_data_4(sec,0,[1,1,1,1],[0.0,1.0,0.0,1.1,0.9,1.1,0.9],name)
    #thay doi vi tri bus 110000
    psspy.growdiagram_2(1,1,[sld],5,len2[i],[0,0,0,0,0,0,0,0,0,0,0,0,0])
    mpt_number.append(sec)

    name2 = f"MPT{i}_TER"
    sld2 =f"BU {100000 + 10000*i+3}"
    ter =100000 + 10000*i+3
    psspy.bus_data_4(ter,0,[1,1,1,1],[0.0,1.0,0.0,1.1,0.9,1.1,0.9],name2)
    psspy.growdiagram_2(1,1,[sld2],6,len2[i]+1,[0,0,0,0,0,0,0,0,0,0,0,0,0])

    if i==1:
        psspy.bus_data_4(110002,0,[1,1,1,1],[0.0,1.0,0.0,1.1,0.9,1.1,0.9],r"""MPT1_PRI""")
        psspy.growdiagram_2(1,1,[r"""BU 110002"""],9,position_110002,[0,0,0,0,0,0,0,0,0,0,0,0,0])

    ## three wnd
    psspy.three_wnd_imped_data_4(110002,sec,ter,r"""1""",[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,1],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],"","")

    psspy.three_wnd_winding_data_5(110002,sec,ter,r"""1""",1,[_i,_i,_i,_i,_i,1],[_f,_f,_f,_f,_f,1.02,0.98,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    psspy.three_wnd_winding_data_5(110002,sec,ter,r"""1""",1,[_i,_i,_i,_i,_i,-1],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    psspy.three_wnd_winding_data_5(110002,sec,ter,r"""1""",2,[_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,1.02,0.98,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    psspy.three_wnd_winding_data_5(110002,sec,ter,r"""1""",3,[_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,1.02,0.98,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])

    psspy.growdiagram_2(1,1,[sld],5,-i*2,[0,0,0,0,0,0,0,0,0,0,0,0,0])

##creat branch
for i in range(mpt):

    frombus = 10000+1000*(i+1)
    sld =f"BU {frombus}"
    tobus = mpt_number[res[i]-1]
    psspy.branch_data_3(frombus,tobus,r"""1""",[1,frombus,1,0,0,0],[0.0,0.0001,0.0,0.0,0.0,0.0,0.0,0.0,1.0,1.0,1.0,1.0],[0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],"")
    psspy.growdiagram_2(1,1,[sld],1,-i*2,[0,0,0,0,0,0,0,0,0,0,0,0,0])


psspy.bus_data_4(960000,0,[1,1,1,1],[0.0,1.0,0.0,1.1,0.9,1.1,0.9],r"""DUMMY_SUB""")
psspy.growdiagram_2(1,1,[r"""BU 960000"""],10,position_110002,[0,0,0,0,0,0,0,0,0,0,0,0,0])


psspy.bus_data_4(970000,0,[3,1,1,1],[0.0,1.0,0.0,1.1,0.9,1.1,0.9],r"""POI""")
psspy.growdiagram_2(1,1,[r"""BU 970000"""],15,position_110002,[0,0,0,0,0,0,0,0,0,0,0,0,0])

##mpt1-dummy_sub
psspy.branch_data_3(110002,960000,r"""1""",[1,960000,1,0,0,0],[0.0,0.0001,0.0,0.0,0.0,0.0,0.0,0.0,1.0,1.0,1.0,1.0],[0.0,0.0001,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],"")
psspy.growdiagram_2(1,1,[r"""BU 110002"""],1,position_110002,[0,0,0,0,0,0,0,0,0,0,0,0,0])

psspy.branch_data_3(960000,999997,r"""1""",[1,999997,1,0,0,0],[0.0,0.0001,0.0,0.0,0.0,0.0,0.0,0.0,1.0,1.0,1.0,1.0],[0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],"")
psspy.growdiagram_2(1,1,[r"""BU 960000"""],1,position_110002,[0,0,0,0,0,0,0,0,0,0,0,0,0])

psspy.plant_data_4(970000,0,[0,0],[1.0,100.0])
psspy.machine_data_4(970000,r"""1""",[1,1,0,0,0,0,0],[0.0,0.0,9999.0,-9999.0,9999.0,-9999.0,100.0,0.0,1.0,0.0,0.0,1.0,1.0,1.0,1.0,1.0,1.0],"")

psspy.branch_data_3(999997,970000,r"""1""",[1,970000,1,0,0,0],[0.0,0.0001,0.0,0.0,0.0,0.0,0.0,0.0,1.0,1.0,1.0,1.0],[0.0,0.0001,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0],"")
psspy.growdiagram_2(1,1,[r"""BU 970000"""],1,position_110002,[0,0,0,0,0,0,0,0,0,0,0,0,0])

psspy.newseq()
psspy.seq_branch_data_3(110002,960000,r"""1""",_i,[_f,0.0001,_f,_f,_f,_f,_f,_f])
psspy.seq_branch_data_3(970000,999997,r"""1""",_i,[_f,0.0001,_f,_f,_f,_f,_f,_f])

psspy.machine_chng_4(970000,r"""1""",[_i,_i,_i,_i,_i,_i,_i],[_f,_f,_f,_f,_f,_f,10000.0,_f,0.1,_f,_f,_f,_f,_f,_f,_f,_f],_s)
psspy.seq_machine_data_4(970000,r"""1""",_i,[_f,0.01,_f,0.01,_f,0.01,0.01,0.01,_f,_f,_f])
print(position_110002)


### creat Transformer
ierr=psspy.save(r'D:\1_GITHUB\PyD\test.sav')
psspy.savediagfile(r'D:\1_GITHUB\PyD\test.sld')


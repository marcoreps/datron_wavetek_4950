#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pyvisa
import time

rm = pyvisa.ResourceManager()
#F5700EP = rm.open_resource('GPIB0::1::INSTR') # Ethernet GPIB Dongle
#dmm = rm.open_resource('GPIB0::9::INSTR') # Ethernet GPIB Dongle

#F5700EP = rm.open_resource('GPIB0::1::INSTR') # Local GPIB Dongle
dmm = rm.open_resource('GPIB0::9::INSTR') # Local GPIB Dongle
dmm.timeout = 1000*60*2 
dmm.query_delay = 1
dmm.read_termination = "\n"

########## DMM and MFC ##########   
#F5700EP.write("*RST")
#F5700EP.write("*CLS")
#F5700EP.write("STBY")
#F5700EP.write("EXTGUARD OFF")
#print("SRC configured")


dmm.write("*RST")
dmm.write("TRIG_SRCE EXT")    
dmm.write("ACCURACY HIGH")
dmm.write("CORRECTN CERTIFIED")
dmm.write("BAND OFF")
dmm.write("LCL_GUARD")
dmm.write("LEAD_NO '31255'")
dmm.write("OHMS 10,PCENT_0,TWR,LCL_GUARD")
dmm.write("ENBCAL CERTIFIED")
print("D4950 configured")


#########################
#Zero
while 1:
    str_input = input("Make FWR short and then input: go\n")
    if (str_input == 'go'):
        break
    else:
        print("what")
#########################

if (dmm.query("CHSE? LEAD")=="0"):
    print("Characterized the Nominated Lead for 2-Wire Ohms Measurements")
    print(dmm.query("LEAD_NO?"))
else:
    print("Fail while characterizing the Nominated Lead for 2-Wire Ohms Measurements")
    quit()

for ix in range (0,5):
    if ix == 0:
        dmm.write("DCV 0.1,PCENT_0,LCL_GUARD")
    if ix == 1:
        dmm.write("DCV 1,PCENT_0,LCL_GUARD")
    if ix == 2:
        dmm.write("DCV 10,PCENT_0,LCL_GUARD")
    if ix == 3:
        dmm.write("DCV 100,PCENT_0,LCL_GUARD")
    if ix == 4:
        dmm.write("DCV 1000,PCENT_0,LCL_GUARD")

    if (dmm.query("CAL?")=="0"):
        print("DCV Zero adjust step "+str(ix)+" completed")
    else:
        print("Fail during DCV Zero adjust step"+str(ix))
        quit()

        
for ix in range (0,8):
    if ix == 0:
        dmm.write("OHMS 10,PCENT_0,FWR,LCL_GUARD")
        print("DMM OHMS FWR Range: 10 OHM, PCENT_0")
    elif ix == 1:
        dmm.write("OHMS 100,PCENT_0,FWR,LCL_GUARD")
        print("DMM OHMS FWR Range: 100 OHM, PCENT_0")
    elif ix == 2:
        dmm.write("OHMS 1000,PCENT_0,FWR,LCL_GUARD")
        print("DMM OHMS FWR Range: 1 KOHM, PCENT_0")
    elif ix == 3:
        dmm.write("OHMS 10000,PCENT_0,FWR,LCL_GUARD")
        print("DMM OHMS FWR Range: 10 KOHM, PCENT_0")
    elif ix == 4:
        dmm.write("OHMS 100000,PCENT_0,FWR,LCL_GUARD")
        print("DMM OHMS FWR Range: 100 KOHM, PCENT_0")
    elif ix == 5:
        dmm.write("OHMS 1000000,PCENT_0,FWR,LCL_GUARD")
        print("DMM OHMS FWR Range: 1 MOHM, PCENT_0")
    elif ix == 6:
        dmm.write("OHMS 10000000,PCENT_0,FWR,LCL_GUARD")
        print("DMM OHMS FWR Range: 10 MOHM, PCENT_0")
    elif ix == 7:
        dmm.write("OHMS 100000000,PCENT_0,FWR,LCL_GUARD")
        print("OHMS 100000000,PCENT_0,FWR,LCL_GUARD")
        
    if (dmm.query("CAL?")=="0"):
        print("OHM Zero adjust step "+str(ix)+" completed")
    else:
        print("Fail during OHM Zero adjust step"+str(ix))
        quit()

for ix in range (0,5):
    if ix == 0:
        dmm.write("DCI 0.0001,PCENT_0,LCL_GUARD")
        print("Zero DMM DCI Range: 100uA")
    elif ix == 1:
        dmm.write("DCI 0.001,PCENT_0,LCL_GUARD")
        print("Zero DMM DCI Range: 1mA")
    elif ix == 2:
        dmm.write("DCI 0.01,PCENT_0,LCL_GUARD")
        print("Zero DMM DCI Range: 10mA")
    elif ix == 3:
        dmm.write("DCI 0.1,PCENT_0,LCL_GUARD")
        print("Zero DMM DCI Range: 100mA")
    elif ix == 4:
        dmm.write("DCI 1,PCENT_0,LCL_GUARD")
        print("Zero DMM DCI Range: 1A")        

    if (dmm.query("CAL?")=="0"):
        print("DCI Zero adjust step "+str(ix)+" completed")
    else:
        print("Fail during DCI Zero adjust step"+str(ix))
        quit()

dmm.write("EXIT DATE")
print("All zero adjustments done.")
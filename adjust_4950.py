#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pyvisa as visa
import time
import logging


def finish():
    logging.info("Shutting down...")
    #Reset the DMM and MFC####
    F5700EP.write("OUT 0 V, 0 Hz")
    F5700EP.write("STBY")
    F5700EP.write("*RST")
    F5700EP.write("*CLS")
    F5700EP.close()
    dmm.close()
    quit()


def freq_adjust():
    dmm.write("ACV 1,FREQ_300,PCENT_100")
    F5700EP.write("OUT 1 V, 300 Hz")
    F5700EP.write("OPER")
    logging.info("Cal FREQ 300 Hz")
    time.sleep(settling_time)
    if(dmm.query("CAL_FREQ?") != '0\n'):
        logging.error("Error")
        finish()
    F5700EP.write("STBY")


def lead_characterization():
    dmm.write("OHMS 10,PCENT_0,TWR")
    F5700EP.write("OUT 0 OHM")
    F5700EP.write("OPER")
    logging.info("Characterizing test lead")
    time.sleep(settling_time)
    if(dmm.write("CHSE? LEAD") != '0\n'):
        logging.error("Error")
        finish()
    F5700EP.write("STBY")

def dcv():
    for v in [0.1,1,10,100,1000]:
        for pcent in [[0,"PCENT_0"],[1,"PCENT_100"],[-1,"PCENT_100"],[1.9,"PCENT_190"],[-1.9,"PCENT_190"]]:
            if pcent[1] == "PCENT_190" and not v == 10:
                continue
            dmm.write("DCV "+str(v)+","+pcent[1]+"")
            F5700EP.write("OUT "+str(pcent[0]*v))
            F5700EP.write("OPER")
            logging.info("Cal DCV "+str(v*pcent[0]))
            time.sleep(settling_time)
            if v == 1000 and pcent[0] == 1:
                time.sleep(600-settling_time)
            if(dmm.query("CAL?") != '0\n'):
                logging.error("Error")
                finish()
    F5700EP.write("STBY")

def acv():
    while 1:
        str_input = input("connect the cable for ACV 4W adjustments, and then input: go\n")
        if (str_input == 'go'):
            break
        else:
            print("input again")

    dmm.write("ACV 1,FREQ_1k,PCENT_100,FWR")
    F5700EP.write("OUT 1 V, 1 KHz")
    F5700EP.write("OPER")
    logging.info("Thermal sensor warm up ...")
    time.sleep(180)

    for v in [0.001,0.01,0.1,1,10]:
        for freq in ["1 kHz","1 MHz","10 Hz","20 Hz","30 Hz","40 Hz","55 Hz","300 Hz","10 kHz","20 kHz","30 kHz","50 kHz","100 kHz","300 kHz","500 kHz"]:
            cutstr = freq.split(" ")
            if cutstr[1] == "Hz":
                dmm.write("ACV "+str(v)+",FREQ_"+cutstr[0]+",PCENT_100,FWR")
            elif cutstr[1] == "kHz":
                  dmm.write("ACV "+str(v)+",FREQ_"+cutstr[0]+"k,PCENT_100,FWR")
            elif cutstr[1] == "MHz":
                  dmm.write("ACV "+str(v)+",FREQ_"+cutstr[0]+"M,PCENT_100,FWR")
            F5700EP.write("OUT "+str(v)+" V, "+freq)
            F5700EP.write("OPER")
            logging.info("Cal ACV "+str(v)+" V, "+freq)
            time.sleep(settling_time)
            if(dmm.query("CAL?") != '0\n'):
                logging.error("Error")
                finish()
            if v == 10 and freq == "1 kHz":
                dmm.write("ACV 10,FREQ_1k,PCENT_190,FWR")
                F5700EP.write("OUT 19 V, 1 kHz")
                logging.info("Cal ACV  19 V, 1 kHz")
                time.sleep(settling_time)
                if(dmm.query("CAL?") != '0\n'):
                    logging.error("Error")
                    finish()
                
    v = 100
    for freq in ["1 kHz","200 kHz","10 Hz","20 Hz","30 Hz","40 Hz","55 Hz","300 Hz","10 kHz","20 kHz","30 kHz","50 kHz","100 kHz"]:
        cutstr = freq.split(" ")
        if cutstr[1] == "Hz":
            dmm.write("ACV "+str(v)+",FREQ_"+cutstr[0]+",PCENT_100,FWR")
        elif cutstr[1] == "kHz":
              dmm.write("ACV "+str(v)+",FREQ_"+cutstr[0]+"k,PCENT_100,FWR")
        elif cutstr[1] == "MHz":
              dmm.write("ACV "+str(v)+",FREQ_"+cutstr[0]+"M,PCENT_100,FWR")
        F5700EP.write("OUT "+str(v)+" V, "+freq)
        F5700EP.write("OPER")
        logging.info("Cal ACV "+str(v)+" V, "+freq)
        time.sleep(settling_time)
        if(dmm.query("CAL?") != '0\n'):
            logging.error("Error")
            finish()
            
    v = 1000
    for freq in ["1 kHz","30 kHz","10 Hz","20 Hz","30 Hz","40 Hz","55 Hz","300 Hz","10 kHz","20 kHz"]:
        cutstr = freq.split(" ")
        if cutstr[1] == "Hz":
            dmm.write("ACV "+str(v)+",FREQ_"+cutstr[0]+",PCENT_100,FWR")
        elif cutstr[1] == "kHz":
              dmm.write("ACV "+str(v)+",FREQ_"+cutstr[0]+"k,PCENT_100,FWR")
        elif cutstr[1] == "MHz":
              dmm.write("ACV "+str(v)+",FREQ_"+cutstr[0]+"M,PCENT_100,FWR")
        F5700EP.write("OUT "+str(v)+" V, "+freq)
        F5700EP.write("OPER")
        logging.info("Cal ACV "+str(v)+" V, "+freq)
        if freq == "1 kHz":
            time.sleep(600)
        else:
            time.sleep(settling_time)
        if(dmm.query("CAL?") != '0\n'):
            logging.error("Error")
            finish()
            
    v = 700
    for freq in ["50 kHz","100 kHz"]:
        cutstr = freq.split(" ")
        dmm.write("ACV "+str(v)+",FREQ_"+cutstr[0]+"k,PCENT_100,FWR")
        F5700EP.write("OUT "+str(v)+" V, "+freq)
        F5700EP.write("OPER")
        logging.info("Cal ACV "+str(v)+" V, "+freq)
        if freq == "50 kHz":
            time.sleep(600)
        else:
            time.sleep(settling_time)
        if(dmm.query("CAL? 700") != '0\n'):
            logging.error("Error")
            finish()

    F5700EP.write("OUT 0.0 V, 0 Hz")
    time.sleep(10)
    F5700EP.write("STBY")


def dci():
    while 1:
        str_input = input("connect the cable for DCI & ACI adjustments, and then input: go\n")
        if (str_input == 'go'):
            break
        else:
            print("input again")
            
    F5700EP.write("OUT 0 A, 0 Hz")
    F5700EP.write("OPER")
    for i in [0.0001,0.001,0.01,0.1,1]:
        for pcent in [[0,"PCENT_0"],[1,"PCENT_100"],[-1,"PCENT_100"]]:
            dmm.write("DCI "+str(i)+","+pcent[1])
            F5700EP.write("OUT "+str(i*pcent[0]))
            logging.info("Cal DCI "+str(i*pcent[0])+" A")
            time.sleep(settling_time)
            if i == 1 and pcent[0] == 1:
                time.sleep(600-settling_time)
            if(dmm.query("CAL?") != '0\n'):
                logging.error("Error")
                finish()

    F5700EP.write("STBY")


def aci():
    for i in [0.0001,0.001,0.01,0.1,1]:
        for freq in ["300 Hz","5 kHz","10 Hz","20 Hz","30 Hz","40 Hz","55 Hz","1 kHz","10 kHz","20 kHz","30 kHz"]:
            cutstr = freq.split(" ")
            if cutstr[1] == "Hz":
                dmm.write("ACI "+str(i)+",FREQ_"+cutstr[0])
            elif cutstr[1] == "kHz":
                  dmm.write("ACI "+str(i)+",FREQ_"+cutstr[0]+"k")
            F5700EP.write("OUT "+str(i)+" A, "+freq)
            F5700EP.write("OPER")
            logging.info("Cal ACI "+str(i)+" A, "+freq)
            time.sleep(settling_time)
            if i == 1 and freq == "300 Hz":
                time.sleep(600-settling_time)
            if(dmm.query("CAL?") != '0\n'):
                logging.error("Error")
                finish()
        
    F5700EP.write("STBY")


def ohms():
    while 1:
        str_input = input("connect the cable for 4W OHMS adjustments, and then input: go\n")
        if (str_input == 'go'):
            break
        else:
            print("input again")

    F5700EP.write("OUT 0 OHM")
    F5700EP.write("EXTSENSE ON")
    F5700EP.write("OPER")
    time.sleep(settling_time)
    for r in [100000,1000000,10000000,100000000,10000,1000,100,10]:
        for band in [[0,"PCENT_0"],[1,"PCENT_100"]]:
            F5700EP.write("OUT "+str(r*band[0]))
            F5700EP.write("OPER")
            if r == 100000000:
                F5700EP.write("EXTSENSE OFF")
                dmm.write("OHMS "+str(r)+","+band[1]+",TWR")
            else:
                dmm.write("OHMS "+str(r)+","+band[1]+",FWR")
                F5700EP.write("EXTSENSE ON")
            F5700EP.write("OUT?")
            res = F5700EP.read()
            cutstr = res.split(",")
            logging.info("Cal OHMS FWR "+cutstr[0]+" Ohm")
            time.sleep(settling_time)
            if(dmm.query("CAL? "+cutstr[0]) != '0\n'):
                logging.error("Error")
                finish()

    F5700EP.write("STBY")
    
def do_all():
    # order recommended by manual
    # if you find that it doesn't matter
    # some connection reconfigs could be skipped
    freq_adjust()
    lead_characterization()
    dcv()
    acv()
    dci()
    aci()
    ohms()
    






settling_time = 60

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)-8s %(message)s')
logging.info("Starting 4950 adjustment ...")

rm = visa.ResourceManager()
#F5700EP = rm.open_resource("TCPIP::192.168.0.88::GPIB0,1") # Ethernet GPIB Dongle
#dmm = rm.open_resource("TCPIP::192.168.0.88::GPIB0,16") # Ethernet GPIB Dongle

F5700EP = rm.open_resource('GPIB0::1::INSTR') # Local GPIB Dongle
dmm = rm.open_resource('GPIB0::10::INSTR') # Local GPIB Dongle

########## DMM and MFC ##########   
F5700EP.write("*RST")
F5700EP.write("*CLS")
F5700EP.write("RANGELCK OFF")
#F5700EP.write("OUT 10 V, 0 Hz")
#F5700EP.write("RANGELCK ON")
F5700EP.write("STBY") 
F5700EP.write("OUT 0.0 V, 0 Hz")
F5700EP.write("EXTGUARD OFF")
time.sleep(5)
F5700EP.write("OPER")
time.sleep(5)
logging.info("SRC configured")

dmm.write("*RST")
dmm.write("*CLS")
time.sleep(2)
dmm.write("ENBCAL CERTIFIED")
dmm.write("TRIG_SRCE EXT")    
dmm.write("ACCURACY HIGH")
dmm.write("BAND OFF")
dmm.write("LCL_GUARD")
dmm.write("LEAD_NO '31255'")
dmm.write("CERT_AMB DEG23C")
dmm.timeout = 150000
logging.info("4950 configured")

        
ohms()



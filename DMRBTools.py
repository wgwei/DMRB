# -*- coding: utf-8 -*-
"""
Created on Fri Oct 20 14:45:34 2017

@author: Weigang Wei
"""

import csv

class HGV():
    def __init__(self,aadtHGV=0, ampmHGV=0):
        self.LGVx = 1.12 #van and <3.5 ton
        self.OGV1 = 1.20 # >3.5 tones <= 3 axles
        self.OGV2 = 1.20 # >3.5 tones >3 axles
        self.PSV = 0.97  # bus and coaches

class HourlyFlowGroup():
    def __init__(self, Network):
        if Network==0:
            self.SI = 1.06
        elif Network==1:
            self.SI = 1.00
        elif Network==2:
            self.SI = 1.00
        elif Network==3:
            self.SI = 1.10
        elif Network==4:
            self.SI = 1.10
        else:
            assert(False) # wrong input range 
            
        self.FGM1 = 0.446 - 0.159*self.SI
        self.FGM2 = 1.581 - 0.089*self.SI
        self.FGM3 = 1.630 + 0.326*self.SI
        self.FGM4 = 1.371 + 0.981*self.SI
        
        # 18 hours distribution in different hourly flow group
        self.hoursInGroup1 = 6  #19:00 to 24:00 and 06:00 to 07:00
        self.hoursInGroup2 = 8  #09:00 to 17:00
        self.hoursInGroup3 = 2  #07:00 to 09:00 AM peak
        self.hoursInGroup4 = 2  #17:00 to 19:00 PM peak
        
class TrafficFlow(HourlyFlowGroup):
    def __init__(self, Network, AADT=0, AMPeak=0,PMPeak=0, AAHT=0, HGVPerc=0):
        """ Hourly Flow Group - HFG
            Network =   0 motorway (MWY); 
                        1 Built-up Trunk (TBU); 
                        2 Built-up Principle(PBU);
                        3 Non Built-up Trunk(TNB); 
                        4 No Built-up Principle (PNB)
            AADT = annaul average daily traffic
            AMPeak = traffic between 07:00 and 09:00
            PMPeak = traffic between 17:00 and 19:00
            AAHT = annual average hourly traffic
        """
        
        HourlyFlowGroup.__init__(self, Network)
        HGV.__init__(self)
        
        self.Network = Network
        self.HGVPerc = HGVPerc
        
        if AADT>0:
            self.AADT = AADT
            self.AAWT = self._AADT_to_AAWT()
            self.HGVPerc = self._HGV_perc_AMPM()
        if AMPeak>0:
            print("AM peak traffic recognised !")
        if PMPeak>0:
            print("AM peak traffic recognised !")
        if (AAHT>0):
            print("AAHT")
    
    def _AADT_to_AAWT(self):
        """ AADT = annual average daily traffic
            AAWT = annual average weekday 18hour traffic
        """
        AAHT = self.AADT / 24.
        trafficG1 = AAHT * self.FGM1 * self.hoursInGroup1
        trafficG2 = AAHT * self.FGM2 * self.hoursInGroup2
        trafficG3 = AAHT * self.FGM3 * self.hoursInGroup3
        trafficG4 = AAHT * self.FGM4 * self.hoursInGroup4
        AAWT = trafficG1 + trafficG2 + trafficG3 + trafficG4
        return AAWT
    
    def _HGV_perc_AMPM(self):
        return self.HGVPerc*self.OGV1
        
    def _AMPMPeak_to_AAWT(self):
        pass
    
class Roads():
    def __init__(self, rn, nw, aadt, ampeak, pmpeak, aaht, hgv,speed):
        self.RoadName = rn
        self.Network =nw
        self.AADT = aadt
        self.AMPeak = ampeak
        self.PMPeak = pmpeak
        self.AAHT = aaht
        self.HGVPerc = hgv
        self.Speed = speed
            
def read_traffic_input(filename="traffic_input.CSV"):
    ''' 
        First line in the traffic input file should be introduction
        Sectond line should be the keys
        keys should be 
        ['RoadName', 'Network', 'AADT', 'AMPeak', 'PMPeak', 'AAHT', 'HGVPercentage', 'Speed km/h']
    ''' 
    with open(filename, "r") as f:
        roads = []
        traffic = csv.reader(f)
        for n, tr in enumerate(traffic):
            if n>1:
                roads.append(Roads(tr[0], tr[1], tr[2], tr[3], tr[4], tr[5], tr[6], tr[7]))
    return roads

def process(roads,outputFile="AAWT_output.CSV"):
    with open(outputFile, "w") as w:
        w.write("RoadName, AADT_24hr, AAWT_18hr, HGV%, Speed km/h\n")
        for rd in roads:
            print("\n", rd.RoadName)
            tf = TrafficFlow(Network=int(rd.Network), AADT=float(rd.AADT), HGVPerc=float(rd.HGVPerc))
            print(int(float(rd.AADT)), int(tf.AAWT), tf.HGVPerc)
            w.write("{},{},{},{}, {}\n".format(rd.RoadName, int(float(rd.AADT)), int(tf.AAWT), tf.HGVPerc, rd.Speed) )
    
    print("\n\ndone !")
        
def main():
    roads = read_traffic_input("traffic_input_DS2030.CSV")
    process(roads, "AAWT_output_DS2030.CSV")

if __name__=="__main__":
    main()
    
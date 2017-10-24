Attribute VB_Name = "DMRBTool"
Function AADT2AAWT(Network As Integer, AADT As Double)
Attribute AADT2AAWT.VB_Description = "AADT2AAWT(Newwork, AADT)"
' AADT = 24 hour annual average daily traffic
' AAWT = 18 hour annual weekday traffic
' Network 0 motorway (MWY); 1 Built-up Trunk (TBU); 2 Built-up Principle(PBU);3 Non Built-up Trunk(TNB);4 No Built-up Principle (PNB)
' FGM = Flow Group Multiplier. it is the ratio between the flow in this hourly flow group over AAHT
' SI  = seaonality index related to road network
' HFG = Hourly flwo group
' HFG 1 is between 19:00 and 07:00
' HFG 2 is between 09:00 and 17:00
' HFG 3 may be the AM peak 07:00 to 09:00
' HFG 4 may be the PM peak 17:00 to 19:00

    Dim SI As Double
    Dim FGM1, FGM2, FGM3, FGM4, AAHT, TrafficG1, TrafficG2, TrafficG3, TrafficG4 As Double

    Select Case Network
        Case 0
            SI = 1.06
        Case 1, 2
            SI = 1#
        Case 3, 4
            SI = 1.1
    End Select
    FGM1 = 0.446 - 0.159 * SI
    FGM2 = 1.581 - 0.089 * SI
    FGM3 = 1.63 + 0.326 * SI
    FGM4 = 1.371 + 0.981 * SI
    
    AAHT = AADT / 24#
    TrafficG1 = AAHT * FGM1 * 6 ' 6 hours in group 1
    TrafficG2 = AAHT * FGM2 * 8 ' 8 hours in group 2
    TrafficG3 = AAHT * FGM3 * 2 ' 2 hours in group 3
    TrafficG4 = AAHT * FGM4 * 2 ' 2 hours in group 4
    
    AADT2AAWT = TrafficG1 + TrafficG2 + TrafficG3 + TrafficG4

End Function

Function HGVWeekday(HGV_AADT As Double)
    HGVWeekday = HGV_AADT * 1.2
End Function


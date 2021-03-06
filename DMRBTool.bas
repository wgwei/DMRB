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
Attribute HGVWeekday.VB_Description = "Convert AADT HGV% to weekday only HGV%"
' convert HGV% of AADT to HGV% of weekday only
    HGVWeekday = HGV_AADT * 1.2
End Function

Function AMPMPeak2AAWT(Network As Integer, AMPeak As Double, PMPeak As Double)
    Dim SI As Double
    Dim FGM1, FGM2, FGM3, FGM4, AAHTam, AAHTpm, AAHT, TrafficG1, TrafficG2, TrafficG3, TrafficG4 As Double

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
    
    AAHTam = AMPeak / FGM3
    AAHTpm = PMPeak / FGM4
    AAHT = (AAHTam + AAHTpm) / 2#
    TrafficG1 = AAHT * FGM1 * 6 ' 6 hours in group 1
    TrafficG2 = AAHT * FGM2 * 8 ' 8 hours in group 2
    TrafficG3 = AAHT * FGM3 * 2 ' 2 hours in group 3
    TrafficG4 = AAHT * FGM4 * 2 ' 2 hours in group 4
    
    AMPMPeak2AAWT = TrafficG1 + TrafficG2 + TrafficG3 + TrafficG4
End Function

Function countValueBtwX_Y(values As Range, X As Double, Y As Double)
    Dim v As Variant
    Dim i As Long
    i = 0
    For Each v In values
        If v.Value > X And v.Value < Y Then
            i = i + 1
        End If
    Next v
    countValueBtwX_Y = i
End Function
Function ShortTermImpact(levelDiff As Double)
    Dim magnitude As String
    If levelDiff < 0 Then
        magnitude = "-"
    ElseIf levelDiff = 0 Then
         magnitude = "No change"
    ElseIf levelDiff > 0.09 And levelDiff <= 0.9 Then
        magnitude = "Negligible"
    ElseIf levelDiff >= 1 And levelDiff <= 2.9 Then
        magnitude = "Minor"
    ElseIf levelDiff >= 3 And levelDiff <= 4.9 Then
        magnitude = "Moderate"
    ElseIf levelDiff >= 5 Then
        magnitude = "Major"
    
    End If
    ShortTermImpact = magnitude
    
End Function

Function LongTermImpact(levelDiff As Double)
    Dim magnitude As String
    If levelDiff < 0 Then
        magnitude = "-"
    ElseIf levelDiff = 0 Then
         magnitude = "No change"
    ElseIf levelDiff > 0.09 And levelDiff <= 2.9 Then
        magnitude = "Negligible"
    ElseIf levelDiff >= 3 And levelDiff <= 4.9 Then
        magnitude = "Minor"
    ElseIf levelDiff >= 5 And levelDiff <= 9.9 Then
        magnitude = "Moderate"
    ElseIf levelDiff >= 10 Then
        magnitude = "Major"
    
    End If
    LongTermImpact = magnitude
    
End Function
Function NuisanceByL10(LA1018hr As Double)
    NuisanceByL10 = (100# / (1 + Exp(-(0.12 * LA1018hr - 9.08)))) / 100#
End Function

Function NuisanceChangeByL10Change(LA10change As Double)
    If LA10change >= 0 Then
        NuisanceChangeByL10Change = 21# * (LA10change ^ 0.33) / 100#
    Else
        NuisanceChangeByL10Change = -1#
    End If
End Function


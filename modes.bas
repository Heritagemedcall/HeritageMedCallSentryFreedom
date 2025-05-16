Attribute VB_Name = "modES"
Option Explicit


Global TempDevice As cESDevice

' Global ESDeviceTypeXREF as collection


'0602 is ack from network controller.
'need to see this before sending next message, or timeout (recent tests show 30 seconds maximum turnaround time)

Global Const MSGTYPE_REPEATERNID = -4
Global Const MSGTYPE_DELETEALL = -3
Global Const MSGTYPE_TWOWAYNID = -2
' gets NID for ES
Global Const MSGTYPE_REQTXSTAT = -1

Global Const MSGTYPE_SET_TIME = 0
Global Const MSGTYPE_CANNEDACK = 1
Global Const MSGTYPE_SETNID = 2
Global Const MSGTYPE_GETNCSETTINGS = 3
' gets NID for DN
Global Const MSGTYPE_GETNID = 4

Global Const MSGTYPE_CREATEFIELDMSG = 6
Global Const MSGTYPE_GETNCSERIAL = &H90

Global Const MSGTYPE_CUSTOM_CODED = &H26
Global Const MSGTYPE_CUSTOM_CUSTOM = &H28

Global Const MSGTYPE_GENERIC_PAGE = &H100

Global Const OUTBOUND_TIMEOUT = 5  ' seconds to wait for response from outbound request

Global gDefaultDeviceTypeIndex As Long

Global newserial  As String

Global gLoBattDelay As Integer



'Global Const MSGTYPE_CANNED2 = 2
'Global Const MSGTYPE_CANNED3 = 3
'Global Const MSGTYPE_CANNED4 = 4
'Global Const MSGTYPE_CANNED5 = 5
'
'Global Const MSGTYPE_CANNED7 = 7
'Global Const MSGTYPE_CANNED8 = 8

Public Type ModelPTI
  CLSPTI                 As Long
  Model                  As String
End Type



Public Type ESDeviceTypeType

  MIDPTI            As Long  ' MID * 256 + PTI
  MID               As Byte
  PTI               As Byte
  CLS               As Byte
  CLSPTI            As Long
  BiDi              As Boolean  ' bidirectional
  Model             As String
  desc              As String
  Announce          As String
  Announce2         As String
  Checkin           As Long  ' check in period
  Latching          As Integer
  Portable          As Integer
  Fixed             As Integer
  AllowDisable      As Integer
  NumInputs         As Integer  ' 1, 2 maybe 3?
  ClearByReset      As Integer  ' 0 or 1
  EnableInput0      As Integer  ' we'll do up to two inputs for now, add more later
  EnableInput1      As Integer

  Custom            As String   ' placeholder 2/8/12

  NoCheckin         As Integer
  SerialDevice      As Integer
  TemperatureDev    As Integer
  Notes             As String
  NoTamper          As Integer

  IgnoreTamper      As Integer

  AutoClear         As Long
  ' new with 226
  'input 1
  Repeats           As Long
  Pause             As Long
  RepeatUntil       As Integer
  SendCancel        As Integer

  'input 2
  Repeats_A         As Long
  Pause_A           As Long
  RepeatUntil_A     As Integer
  SendCancel_A      As Integer

  'input 3
  Repeats_B         As Long
  Pause_B           As Long
  RepeatUntil_B     As Integer
  SendCancel_B      As Integer

  ' GROUPS BUTTON 1
  OG1               As Long
  OG2               As Long
  OG3               As Long
  OG4               As Long
  OG5               As Long
  OG6               As Long

  NG1               As Long
  NG2               As Long
  NG3               As Long
  NG4               As Long
  NG5               As Long
  NG6               As Long

  GG1               As Long
  GG2               As Long
  GG3               As Long
  GG4               As Long
  GG5               As Long
  GG6               As Long

' GROUPS BUTTON 2
  OG1_A             As Long
  OG2_A             As Long
  OG3_A             As Long
  OG4_A             As Long
  OG5_A             As Long
  OG6_A             As Long
  
  NG1_A             As Long
  NG2_A             As Long
  NG3_A             As Long
  NG4_A             As Long
  NG5_A             As Long
  NG6_A             As Long

  GG1_A             As Long
  GG2_A             As Long
  GG3_A             As Long
  GG4_A             As Long
  GG5_A             As Long
  GG6_A             As Long

' DELAY BUTTON 1
  OG1D              As Long
  OG2D              As Long
  OG3D              As Long
  OG4D              As Long
  OG5D              As Long
  OG6D              As Long

  NG1D              As Long
  NG2D              As Long
  NG3D              As Long
  NG4D              As Long
  NG5D              As Long
  NG6D              As Long
  
  GG1D              As Long
  GG2D              As Long
  GG3D              As Long
  GG4D              As Long
  GG5D              As Long
  GG6D              As Long
  
' DELAY BUTTON 2
  OG1_AD            As Long
  OG2_AD            As Long
  OG3_AD            As Long
  OG4_AD            As Long
  OG5_AD            As Long
  OG6_AD            As Long


  NG1_AD            As Long
  NG2_AD            As Long
  NG3_AD            As Long
  NG4_AD            As Long
  NG5_AD            As Long
  NG6_AD            As Long

  GG1_AD            As Long
  GG2_AD            As Long
  GG3_AD            As Long
  GG4_AD            As Long
  GG5_AD            As Long
  GG6_AD            As Long



End Type

Global MAX_ESDEVICETYPES  As Integer
Global ESDeviceType()     As ESDeviceTypeType

Global AutoEnrollEnabled  As Boolean



Global SurveyDevice       As cESSurveyDevice

Global SurveyEnabled      As Boolean
Global WayPointForm       As Form

Global BatchSurveyEnabled As Boolean
Global BatchForm          As Form

Global Waypoints          As cWaypoints

Global InBounds           As cAlarms
Global Outbounds          As cOutBounds

Global SerialIns          As New Collection

Global LocatorWaitTime    As Double  ' how long to wait for locator packets to arrive default is 5 seconds
Private NextPacketSeq As Long

Private mNIDMatch   As Boolean

Public Sub SetDeviceTypes()
        Dim j                  As Integer

10      j = 0
20      ReDim ESDeviceType(j)
30      ESDeviceType(j).Model = "EN6040"
40      ESDeviceType(j).desc = "Receiver"
50      ESDeviceType(j).MID = &H0
60      ESDeviceType(j).PTI = &H0
70      ESDeviceType(j).CLS = &H0
80      ESDeviceType(j).BiDi = True
90      ESDeviceType(j).NumInputs = 0
100     ESDeviceType(j).Notes = ""   'Network Coordinator uses following byte stream as checkin: x1C x05 NON-ACKS STAT1 STAT0 CKSUM"
110     ESDeviceType(j).Announce = ""
        'ESDeviceType(j).Checkin = 9


120     j = j + 1
130     ReDim Preserve ESDeviceType(j)  ' 1
140     ESDeviceType(j).Model = "EN6080"
150     ESDeviceType(j).desc = "AGC"
160     ESDeviceType(j).MID = &H0
170     ESDeviceType(j).PTI = &H0
180     ESDeviceType(j).CLS = &H1
190     ESDeviceType(j).BiDi = True
        '200     ESDeviceType(j).Fixed = 1
200     ESDeviceType(j).NumInputs = 0
210     ESDeviceType(j).Notes = ""   'Network Coordinator uses following byte stream as checkin: x1C x05 NON-ACKS STAT1 STAT0 CKSUM"
220     ESDeviceType(j).Announce = ""
        'ESDeviceType(j).Checkin = 9

230     j = j + 1
240     ReDim Preserve ESDeviceType(j)  '2
250     ESDeviceType(j).Model = "ES1210"
260     ESDeviceType(j).desc = "Universal End Device"
270     ESDeviceType(j).MID = &HB2
280     ESDeviceType(j).PTI = &H0
290     ESDeviceType(j).CLS = &H3E
300     ESDeviceType(j).NumInputs = 1
310     ESDeviceType(j).Notes = ""
320     ESDeviceType(j).Announce = "Alarm"
330     ESDeviceType(j).NoTamper = 1
340     ESDeviceType(j).Checkin = 9


350     j = j + 1
360     ReDim Preserve ESDeviceType(j)  '2
370     ESDeviceType(j).Model = "EN1210-60"
380     ESDeviceType(j).desc = "Universal End Device"
390     ESDeviceType(j).MID = &HB2
400     ESDeviceType(j).PTI = &H2
410     ESDeviceType(j).CLS = &H3E
420     ESDeviceType(j).NumInputs = 1
430     ESDeviceType(j).Notes = ""
440     ESDeviceType(j).Announce = "Alarm"
450     ESDeviceType(j).NoTamper = 1
460     ESDeviceType(j).Checkin = 60 * 3

470     j = j + 1
480     ReDim Preserve ESDeviceType(j)  '2
490     ESDeviceType(j).Model = "EN1210-240"
500     ESDeviceType(j).desc = "Universal End Device"
510     ESDeviceType(j).MID = &HB2
520     ESDeviceType(j).PTI = &H6
530     ESDeviceType(j).CLS = &H3E
540     ESDeviceType(j).NumInputs = 1
550     ESDeviceType(j).Notes = ""
560     ESDeviceType(j).Announce = "Alarm"
570     ESDeviceType(j).NoTamper = 1
580     ESDeviceType(j).Checkin = 60 * 9


        '350     ESDeviceType(j).Fixed = 1
        '  j = j + 1
        '  ReDim Preserve ESDeviceType(j)
        '  ESDeviceType(j).Model = "ES1210SK"
        '  ESDeviceType(j).Desc = "Survey Kit Universal End Device"
        '  ESDeviceType(j).MID = &HB2
        '  ESDeviceType(j).PTI = &H4
        '  ESDeviceType(j).NumInputs = 1
        '  ESDeviceType(j).Notes = ""
        '  ESDeviceType(j).Announce = "Alarm"

590     j = j + 1
600     ReDim Preserve ESDeviceType(j)
610     ESDeviceType(j).Model = "ES1210W"
620     ESDeviceType(j).desc = "Door/Window"
630     ESDeviceType(j).MID = &HB2
640     ESDeviceType(j).PTI = &H3
650     ESDeviceType(j).CLS = &H3E
660     ESDeviceType(j).NumInputs = 2
670     ESDeviceType(j).Notes = ""
680     ESDeviceType(j).Announce = "Alarm"
690     ESDeviceType(j).NoTamper = 1
        '470     ESDeviceType(j).Fixed = 1
700     ESDeviceType(j).Checkin = 9

710     j = j + 1
720     ReDim Preserve ESDeviceType(j)
730     ESDeviceType(j).Model = "EN1210W-60"
740     ESDeviceType(j).desc = "Door/Window"
750     ESDeviceType(j).MID = &HB2
760     ESDeviceType(j).PTI = &H20
770     ESDeviceType(j).CLS = &H3E
780     ESDeviceType(j).NumInputs = 2
790     ESDeviceType(j).Notes = ""
800     ESDeviceType(j).Announce = "Alarm"
810     ESDeviceType(j).NoTamper = 1
820     ESDeviceType(j).Checkin = 60 * 3


830     j = j + 1
840     ReDim Preserve ESDeviceType(j)
850     ESDeviceType(j).Model = "ES1212"
860     ESDeviceType(j).desc = "Two Input Universal"
870     ESDeviceType(j).MID = &HB2
880     ESDeviceType(j).PTI = &H1
890     ESDeviceType(j).CLS = &H3E
900     ESDeviceType(j).NumInputs = 2
910     ESDeviceType(j).Notes = ""
920     ESDeviceType(j).Announce = "Alarm"
930     ESDeviceType(j).Checkin = 9

940     j = j + 1
950     ReDim Preserve ESDeviceType(j)
960     ESDeviceType(j).Model = "EN1212-60"
970     ESDeviceType(j).desc = "Two Input Universal"
980     ESDeviceType(j).MID = &HB2
990     ESDeviceType(j).PTI = &HA
1000    ESDeviceType(j).CLS = &H3E
1010    ESDeviceType(j).NumInputs = 2
1020    ESDeviceType(j).Notes = ""
1030    ESDeviceType(j).Announce = "Alarm"
1040    ESDeviceType(j).Checkin = 60 * 3



        '580     ESDeviceType(j).Fixed = 1

1050    j = j + 1
1060    ReDim Preserve ESDeviceType(j)
1070    ESDeviceType(j).Model = "ES1215"
1080    ESDeviceType(j).desc = "Universal End Device w/ Wall Tamper"
1090    ESDeviceType(j).MID = &HB2
1100    ESDeviceType(j).PTI = &H8
1110    ESDeviceType(j).CLS = &H3E
1120    ESDeviceType(j).NumInputs = 1
1130    ESDeviceType(j).Notes = ""
1140    ESDeviceType(j).Announce = "Alarm"
1150    ESDeviceType(j).Checkin = 9
        '690     ESDeviceType(j).Fixed = 1

1160    j = j + 1
1170    ReDim Preserve ESDeviceType(j)
1180    ESDeviceType(j).Model = "ES1215W"
1190    ESDeviceType(j).desc = "Door/Window w/ Wall Tamper"
1200    ESDeviceType(j).MID = &HB2
1210    ESDeviceType(j).PTI = &HB
1220    ESDeviceType(j).CLS = &H3E
1230    ESDeviceType(j).NumInputs = 2
1240    ESDeviceType(j).Notes = ""
1250    ESDeviceType(j).Announce = "Alarm"
1260    ESDeviceType(j).Checkin = 9
        '800     ESDeviceType(j).Fixed = 1

1270    j = j + 1
1280    ReDim Preserve ESDeviceType(j)
1290    ESDeviceType(j).Model = "ES1216"
1300    ESDeviceType(j).desc = "Two Input Universal w/ Wall Tamper"
1310    ESDeviceType(j).MID = &HB2
1320    ESDeviceType(j).PTI = &H9
1330    ESDeviceType(j).CLS = &H3E
1340    ESDeviceType(j).NumInputs = 2
1350    ESDeviceType(j).Notes = ""
1360    ESDeviceType(j).Announce = "Alarm"
1370    ESDeviceType(j).Checkin = 9
        '910     ESDeviceType(j).Fixed = 1

1380    j = j + 1
1390    ReDim Preserve ESDeviceType(j)
1400    ESDeviceType(j).Model = "EN1223D"
1410    ESDeviceType(j).desc = "Two Button Pendant"
1420    ESDeviceType(j).MID = &HB2
1430    ESDeviceType(j).PTI = &H19
1440    ESDeviceType(j).CLS = &H3E
1450    ESDeviceType(j).Portable = 1
1460    ESDeviceType(j).NumInputs = 2
1470    ESDeviceType(j).Notes = ""
1480    ESDeviceType(j).Announce = "Alarm"
1490    ESDeviceType(j).NoTamper = 1
1500    ESDeviceType(j).Checkin = 9

1510    j = j + 1
1520    ReDim Preserve ESDeviceType(j)
1530    ESDeviceType(j).Model = "EN1223S"
1540    ESDeviceType(j).desc = "One Button Pendant"
1550    ESDeviceType(j).MID = &HB2   '3E class   18 pti
1560    ESDeviceType(j).PTI = &H18
1570    ESDeviceType(j).CLS = &H3E
1580    ESDeviceType(j).Portable = 1
1590    ESDeviceType(j).NumInputs = 1
1600    ESDeviceType(j).Notes = ""
1610    ESDeviceType(j).Announce = "Alarm"
1620    ESDeviceType(j).NoTamper = 1
1630    ESDeviceType(j).Checkin = 9  '  10800 = 60 * 60 * 3

1640    j = j + 1
1650    ReDim Preserve ESDeviceType(j)
1660    ESDeviceType(j).Model = "EN1221S-60"
1670    ESDeviceType(j).desc = "One Button Pendant"
1680    ESDeviceType(j).MID = &HB2   '3E class   1D pti
1690    ESDeviceType(j).PTI = &H1D
1700    ESDeviceType(j).CLS = &H3E
1710    ESDeviceType(j).Portable = 1
1720    ESDeviceType(j).NumInputs = 1
1730    ESDeviceType(j).Notes = ""
1740    ESDeviceType(j).Announce = "Alarm"
1750    ESDeviceType(j).NoTamper = 1
1760    ESDeviceType(j).Checkin = 60 * 3  '  10800 = 60 * 60 * 3


1770    gDefaultDeviceTypeIndex = j

        '        Debug.Assert 0

        '  j = j + 1
        '  ReDim Preserve ESDeviceType(j)
        '  ESDeviceType(j).Model = "EN1223SK"
        '  ESDeviceType(j).Desc = "Survey Kit One Button Pendant"
        '  ESDeviceType(j).MID = &HB2
        '  ESDeviceType(j).PTI = &H1E
        '  ESDeviceType(j).Portable = 1
        '  ESDeviceType(j).NumInputs = 1
        '  ESDeviceType(j).Notes = ""
        '  ESDeviceType(j).Announce = "Alarm"

1780    j = j + 1
1790    ReDim Preserve ESDeviceType(j)
1800    ESDeviceType(j).Model = "ES1223S-60"
1810    ESDeviceType(j).desc = "One Button Necklace Pendant"
1820    ESDeviceType(j).MID = &HB2
1830    ESDeviceType(j).PTI = &H1A
1840    ESDeviceType(j).CLS = &H3E
1850    ESDeviceType(j).Portable = 1
1860    ESDeviceType(j).NumInputs = 1
1870    ESDeviceType(j).Notes = ""
1880    ESDeviceType(j).Announce = "Alarm"
1890    ESDeviceType(j).NoTamper = 1
1900    ESDeviceType(j).Checkin = 60 * 3  '  10800 = 60 * 60 * 3






1910    j = j + 1
1920    ReDim Preserve ESDeviceType(j)
1930    ESDeviceType(j).Model = "ES1233D"
1940    ESDeviceType(j).desc = "Two Button Necklace Pendant"
1950    ESDeviceType(j).MID = &HB2
1960    ESDeviceType(j).PTI = &H15
1970    ESDeviceType(j).CLS = &H3E
1980    ESDeviceType(j).Portable = 1
1990    ESDeviceType(j).NumInputs = 2
2000    ESDeviceType(j).Notes = ""
2010    ESDeviceType(j).Announce = "Alarm"
2020    ESDeviceType(j).NoTamper = 1
2030    ESDeviceType(j).Checkin = 9

2040    j = j + 1
2050    ReDim Preserve ESDeviceType(j)
2060    ESDeviceType(j).Model = "ES1233S"
2070    ESDeviceType(j).desc = "One Button Necklace Pendant"
2080    ESDeviceType(j).MID = &HB2
2090    ESDeviceType(j).PTI = &H14
2100    ESDeviceType(j).CLS = &H3E
2110    ESDeviceType(j).Portable = 1
2120    ESDeviceType(j).Announce = "Alarm"
2130    ESDeviceType(j).NoTamper = 1
2140    ESDeviceType(j).Checkin = 9

2150    j = j + 1
2160    ReDim Preserve ESDeviceType(j)
2170    ESDeviceType(j).Model = "ES1234D"
2180    ESDeviceType(j).desc = "Three Channel Necklace Pendant"
2190    ESDeviceType(j).MID = &HB2
2200    ESDeviceType(j).PTI = &H38
2210    ESDeviceType(j).CLS = &H3E
2220    ESDeviceType(j).Portable = 1
2230    ESDeviceType(j).NumInputs = 2
2240    ESDeviceType(j).Notes = ""
2250    ESDeviceType(j).Announce = "Alarm"
2260    ESDeviceType(j).NoTamper = 1
2270    ESDeviceType(j).Checkin = 9

2280    j = j + 1
2290    ReDim Preserve ESDeviceType(j)
2300    ESDeviceType(j).Model = "ES1235D"
2310    ESDeviceType(j).desc = "Two Button Necklace Pendant"
2320    ESDeviceType(j).MID = &HB2
2330    ESDeviceType(j).PTI = &H11
2340    ESDeviceType(j).CLS = &H3E
2350    ESDeviceType(j).Portable = 1
2360    ESDeviceType(j).NumInputs = 2
2370    ESDeviceType(j).Notes = ""
2380    ESDeviceType(j).Announce = "Alarm"
2390    ESDeviceType(j).NoTamper = 1
2400    ESDeviceType(j).Checkin = 9

2410    j = j + 1
2420    ReDim Preserve ESDeviceType(j)
2430    ESDeviceType(j).Model = "ES1235S"
2440    ESDeviceType(j).desc = "One Button Necklace Pendant"
2450    ESDeviceType(j).MID = &HB2
2460    ESDeviceType(j).PTI = &H10
2470    ESDeviceType(j).CLS = &H3E
2480    ESDeviceType(j).Portable = 1
2490    ESDeviceType(j).NumInputs = 1
2500    ESDeviceType(j).Notes = ""
2510    ESDeviceType(j).Announce = "Alarm"
2520    ESDeviceType(j).NoTamper = 1
2530    ESDeviceType(j).Checkin = 9

2540    j = j + 1
2550    ReDim Preserve ESDeviceType(j)
2560    ESDeviceType(j).Model = "ES1235SF"
2570    ESDeviceType(j).desc = "One Button Fixed Position"
2580    ESDeviceType(j).MID = &HB2
2590    ESDeviceType(j).PTI = &H12
2600    ESDeviceType(j).CLS = &H3E
2610    ESDeviceType(j).Portable = 0
2620    ESDeviceType(j).NumInputs = 1
2630    ESDeviceType(j).Notes = ""
2640    ESDeviceType(j).Announce = "Alarm"
2650    ESDeviceType(j).NoTamper = 1
2660    ESDeviceType(j).Checkin = 9

2670    j = j + 1
2680    ReDim Preserve ESDeviceType(j)
2690    ESDeviceType(j).Model = "ES1236D"
2700    ESDeviceType(j).desc = "Three Channel Beltclip Pendant"
2710    ESDeviceType(j).MID = &HB2
2720    ESDeviceType(j).PTI = &H39
2730    ESDeviceType(j).CLS = &H3E
2740    ESDeviceType(j).NumInputs = 3
2750    ESDeviceType(j).Portable = 1
2760    ESDeviceType(j).Announce = "Alarm"
2770    ESDeviceType(j).NoTamper = 1
2780    ESDeviceType(j).Checkin = 9
2790    j = j + 1
2800    ReDim Preserve ESDeviceType(j)
2810    ESDeviceType(j).Model = "ES1238D"
2820    ESDeviceType(j).desc = "Two Channel Beltclip Pendant"
2830    ESDeviceType(j).MID = &HB2
2840    ESDeviceType(j).PTI = &H3A
2850    ESDeviceType(j).CLS = &H3E
2860    ESDeviceType(j).NumInputs = 2
2870    ESDeviceType(j).Portable = 1
2880    ESDeviceType(j).Announce = "Alarm"
2890    ESDeviceType(j).NoTamper = 1
2900    ESDeviceType(j).Checkin = 9

2910    j = j + 1
2920    ReDim Preserve ESDeviceType(j)
2930    ESDeviceType(j).Model = "EN1240"
2940    ESDeviceType(j).desc = "Activity Sensor"
2950    ESDeviceType(j).MID = &HB2
2960    ESDeviceType(j).PTI = &H31
2970    ESDeviceType(j).CLS = &H3E
2980    ESDeviceType(j).NumInputs = 1
2990    ESDeviceType(j).Notes = ""
3000    ESDeviceType(j).Announce = "Activity"
3010    ESDeviceType(j).Checkin = 9


3020    j = j + 1
3030    ReDim Preserve ESDeviceType(j)
3040    ESDeviceType(j).Model = "EN1241-60"
3050    ESDeviceType(j).desc = "Activity Sensor"
3060    ESDeviceType(j).MID = &HB2
3070    ESDeviceType(j).PTI = &H2D
3080    ESDeviceType(j).CLS = &H3E
3090    ESDeviceType(j).NumInputs = 1
3100    ESDeviceType(j).Notes = ""
3110    ESDeviceType(j).Announce = "Activity"
3120    ESDeviceType(j).Checkin = 60 * 3

3130    j = j + 1
3140    ReDim Preserve ESDeviceType(j)
3150    ESDeviceType(j).Model = "ES1242"  ' chenged back to   ES1242 2/4/16 changed to dash 1/25/2016
3160    ESDeviceType(j).desc = "Residential Smoke Detector"
3170    ESDeviceType(j).MID = &HB2
3180    ESDeviceType(j).PTI = &H2C
3190    ESDeviceType(j).CLS = &H3E
3200    ESDeviceType(j).NumInputs = 1
3210    ESDeviceType(j).Notes = ""
3220    ESDeviceType(j).Announce = "Smoke"
3230    ESDeviceType(j).Checkin = 9


3240    j = j + 1
3250    ReDim Preserve ESDeviceType(j)
3260    ESDeviceType(j).Model = "EN1244"
3270    ESDeviceType(j).desc = "Smoke Detector"
3280    ESDeviceType(j).MID = &HB2
3290    ESDeviceType(j).PTI = &H21
3300    ESDeviceType(j).CLS = &H3E
3310    ESDeviceType(j).NumInputs = 1
3320    ESDeviceType(j).Notes = ""
3330    ESDeviceType(j).Announce = "Smoke"
3340    ESDeviceType(j).Checkin = 9



3350    j = j + 1
3360    ReDim Preserve ESDeviceType(j)
3370    ESDeviceType(j).Model = "EN1245"
3380    ESDeviceType(j).desc = "CO Detector"
3390    ESDeviceType(j).MID = &HB2
3400    ESDeviceType(j).PTI = &H2E
3410    ESDeviceType(j).CLS = &H3E
3420    ESDeviceType(j).NumInputs = 1
3430    ESDeviceType(j).Notes = ""
3440    ESDeviceType(j).Announce = "CO Detected"
3450    ESDeviceType(j).Checkin = 9



3460    j = j + 1
3470    ReDim Preserve ESDeviceType(j)
3480    ESDeviceType(j).Model = "ES1247"
3490    ESDeviceType(j).desc = "Glass Break Detector"
3500    ESDeviceType(j).MID = &HB2
3510    ESDeviceType(j).PTI = &H32
3520    ESDeviceType(j).CLS = &H3E
3530    ESDeviceType(j).NumInputs = 1
3540    ESDeviceType(j).Notes = ""
3550    ESDeviceType(j).Announce = "Glass Break"
3560    ESDeviceType(j).Fixed = 0
3570    ESDeviceType(j).Checkin = 9

3580    j = j + 1
3590    ReDim Preserve ESDeviceType(j)
3600    ESDeviceType(j).Model = "ES1249"
3610    ESDeviceType(j).desc = "Bill Trap"
3620    ESDeviceType(j).MID = &HB2
3630    ESDeviceType(j).PTI = &H30
3640    ESDeviceType(j).CLS = &H3E
3650    ESDeviceType(j).NumInputs = 1
3660    ESDeviceType(j).Notes = ""
3670    ESDeviceType(j).Announce = "Bill Trap"
3680    ESDeviceType(j).Fixed = 0
3690    ESDeviceType(j).Checkin = 9


3700    j = j + 1
3710    ReDim Preserve ESDeviceType(j)
3720    ESDeviceType(j).Model = "EN1252"
3730    ESDeviceType(j).desc = "Two Input Extended Range"
3740    ESDeviceType(j).MID = &HB2
3750    ESDeviceType(j).PTI = &H7
3760    ESDeviceType(j).CLS = &H3E
3770    ESDeviceType(j).NumInputs = 2
3780    ESDeviceType(j).Notes = ""
3790    ESDeviceType(j).Announce = "Alarm"
3800    ESDeviceType(j).Fixed = 0
3810    ESDeviceType(j).Checkin = 9

3820    j = j + 1
3830    ReDim Preserve ESDeviceType(j)
3840    ESDeviceType(j).Model = "ES1255"
3850    ESDeviceType(j).desc = "End Device Arm/Disarm"
3860    ESDeviceType(j).MID = &HB2
3870    ESDeviceType(j).PTI = &H24
3880    ESDeviceType(j).CLS = &H3E
3890    ESDeviceType(j).NumInputs = 2
3900    ESDeviceType(j).Notes = ""
3910    ESDeviceType(j).Announce = "Arm"
3920    ESDeviceType(j).Checkin = 9

3930    j = j + 1
3940    ReDim Preserve ESDeviceType(j)
3950    ESDeviceType(j).Model = "ES1260"
3960    ESDeviceType(j).desc = "Inovonics PIR"
3970    ESDeviceType(j).MID = &HB2
3980    ESDeviceType(j).PTI = &H28
3990    ESDeviceType(j).CLS = &H3E
4000    ESDeviceType(j).NumInputs = 1
4010    ESDeviceType(j).Notes = ""
4020    ESDeviceType(j).Announce = "Intruder"
4030    ESDeviceType(j).Checkin = 9

4040    j = j + 1
4050    ReDim Preserve ESDeviceType(j)
4060    ESDeviceType(j).Model = "ES1262"
4070    ESDeviceType(j).desc = "Bosch PIR"
4080    ESDeviceType(j).MID = &HB2
4090    ESDeviceType(j).PTI = &H29
4100    ESDeviceType(j).CLS = &H3E
4110    ESDeviceType(j).NumInputs = 1
4120    ESDeviceType(j).Notes = ""
4130    ESDeviceType(j).Announce = "Alarm"
4140    ESDeviceType(j).Checkin = 9

4150    j = j + 1
4160    ReDim Preserve ESDeviceType(j)
4170    ESDeviceType(j).Model = "ES1265"
4180    ESDeviceType(j).desc = "Ceiling Mount PIR"
4190    ESDeviceType(j).MID = &HB2
4200    ESDeviceType(j).PTI = &H2A
4210    ESDeviceType(j).CLS = &H3E
4220    ESDeviceType(j).NumInputs = 1
4230    ESDeviceType(j).Notes = ""
4240    ESDeviceType(j).Announce = "Intruder"
4250    ESDeviceType(j).Checkin = 9

4260    j = j + 1
4270    ReDim Preserve ESDeviceType(j)
4280    ESDeviceType(j).Model = "ES1254"
4290    ESDeviceType(j).desc = "Four Channel Belt Pendant"
4300    ESDeviceType(j).MID = &HB2
4310    ESDeviceType(j).PTI = &H3B
4320    ESDeviceType(j).CLS = &H3E
4330    ESDeviceType(j).NumInputs = 4
4340    ESDeviceType(j).Notes = ""
4350    ESDeviceType(j).Announce = "Alarm"
4360    ESDeviceType(j).Portable = 1
4370    ESDeviceType(j).Checkin = 9


        '  j = j + 1
        '  ReDim preserve ESDeviceType(j)
        '  ESDeviceType(j).Model = "ES1720TH"
        '  ESDeviceType(j).Desc = "Thermostat w/ Slide and Override"
        '  ESDeviceType(j).CLS = &HC0
        '  ESDeviceType(j).PTI = &H14

        '2810    j = j + 1
        '2820    ReDim Preserve ESDeviceType(j)
        '2830    ESDeviceType(j).Model = "ES1723S"  ' A
        '2840    ESDeviceType(j).Desc = "Temperature Internal No Options"
        '2850    ESDeviceType(j).mid = &HC0
        '2860    ESDeviceType(j).CLS = &H3C
        '2870    ESDeviceType(j).PTI = &HA
        '2880    ESDeviceType(j).NumInputs = 1
        '2885    ESDeviceType(j).TemperatureDev = 1
        '2890    ESDeviceType(j).Notes = ""
        '2900    ESDeviceType(j).Announce = "Temperature"




4380    j = j + 1
4390    ReDim Preserve ESDeviceType(j)
4400    ESDeviceType(j).Model = "ES1723"  ' 2A
4410    ESDeviceType(j).desc = "Temperature Sensor"
4420    ESDeviceType(j).MID = &HC0
4430    ESDeviceType(j).PTI = &H99
4440    ESDeviceType(j).CLS = &H3C
4450    ESDeviceType(j).TemperatureDev = 1
4460    ESDeviceType(j).NumInputs = 2
4470    ESDeviceType(j).Notes = ""
4480    ESDeviceType(j).Announce = "Temperature"
4490    ESDeviceType(j).Checkin = 9

        '3010    j = j + 1
        '3020    ReDim Preserve ESDeviceType(j)
        '3030    ESDeviceType(j).Model = "ES1723D"  ' 17
        '3040    ESDeviceType(j).Desc = "Temperature Internal-External No Options"
        '3050    ESDeviceType(j).mid = &HC0
        '3060    ESDeviceType(j).CLS = &H3C
        '3070    ESDeviceType(j).PTI = &H17
        '3075    ESDeviceType(j).TemperatureDev = 1
        '3080    ESDeviceType(j).NumInputs = 2
        '3090    ESDeviceType(j).Notes = ""
        '3100    ESDeviceType(j).Announce = "Temperature"
        '
        '
        '
        '3110    j = j + 1
        '3120    ReDim Preserve ESDeviceType(j)
        '3130    ESDeviceType(28).Model = "ES1723DO"  '37
        '3140    ESDeviceType(28).Desc = "Temperature Internal-External Options"
        '3150    ESDeviceType(j).mid = &HC0
        '3160    ESDeviceType(j).CLS = &H3C
        '3170    ESDeviceType(28).PTI = &H37
        '3175    ESDeviceType(j).TemperatureDev = 1
        '3180    ESDeviceType(j).NumInputs = 2
        '3190    ESDeviceType(j).Notes = ""
        '3200    ESDeviceType(j).Announce = "Temperature"


4500    j = j + 1
4510    ReDim Preserve ESDeviceType(j)
4520    ESDeviceType(j).Model = "EN1751"
4530    ESDeviceType(j).desc = "Water Detector"
4540    ESDeviceType(j).MID = &HB2
4550    ESDeviceType(j).PTI = &H23
4560    ESDeviceType(j).CLS = &H3E
4570    ESDeviceType(j).NumInputs = 1
4580    ESDeviceType(j).Notes = ""
4590    ESDeviceType(j).Announce = ""
4600    ESDeviceType(j).NoTamper = 0  '??? custom unit
4610    ESDeviceType(j).Checkin = 9


        '' Version 1961 works 1974 broken.

4620    j = j + 1
4630    ReDim Preserve ESDeviceType(j)
4640    ESDeviceType(j).Model = "EN1941"
4650    ESDeviceType(j).desc = "End Device Binary Input"
4660    ESDeviceType(j).MID = &HB2
4670    ESDeviceType(j).PTI = &HC    ' PTI CLASH with XS!!!!
4680    ESDeviceType(j).CLS = &H3E
4690    ESDeviceType(j).NumInputs = 2
4700    ESDeviceType(j).Notes = ""
4710    ESDeviceType(j).Announce = ""
4720    ESDeviceType(j).NoTamper = 1  '??? custom unit
4730    ESDeviceType(j).Checkin = 9

4740    j = j + 1
4750    ReDim Preserve ESDeviceType(j)
4760    ESDeviceType(j).Model = "EN1941-60"
4770    ESDeviceType(j).desc = "End Device Binary Input"

4780    ESDeviceType(j).MID = &HB2
4790    ESDeviceType(j).PTI = &HF    ' we got a differenet one here.... why not XS ????
4800    ESDeviceType(j).CLS = &H3E
4810    ESDeviceType(j).NumInputs = 2
4820    ESDeviceType(j).Notes = ""
4830    ESDeviceType(j).Announce = ""
4840    ESDeviceType(j).NoTamper = 1  '??? custom unit
4850    ESDeviceType(j).Checkin = 60 * 3

        '3310    ESDeviceType(j).Fixed = 1

        '4640    j = j + 1
        '4650    ReDim Preserve ESDeviceType(j)
        '4660    ESDeviceType(j).Model = "EN1941XS"
        '4670    ESDeviceType(j).desc = "End Device Serial Input"
        '4680    ESDeviceType(j).MID = &HB2
        ''4690    ESDeviceType(j).PTI = &HC    ' same as standard 1941 "WTF!?" shoud be 12-decimal &h0C
        ''4700    ESDeviceType(j).CLS = &H18 ' &H3E not a security device makes it different
        '4690    ESDeviceType(j).PTI = &H34
        '4700    ESDeviceType(j).CLS = &H3E ' &H3E not a security device makes it different
        '
        '4710    ESDeviceType(j).NumInputs = 2
        '4720    ESDeviceType(j).Notes = ""
        '4730    ESDeviceType(j).Announce = ""
        '4740    ESDeviceType(j).Checkin = 9

4860    j = j + 1
4870    ReDim Preserve ESDeviceType(j)
4880    ESDeviceType(j).Model = "ES3941XS"
4890    ESDeviceType(j).desc = "Two Way End Device Serial Input"
4900    ESDeviceType(j).MID = &HB2
4910    ESDeviceType(j).PTI = &H74
4920    ESDeviceType(j).CLS = &H3E
4930    ESDeviceType(j).BiDi = True
4940    ESDeviceType(j).NumInputs = 1
4950    ESDeviceType(j).Notes = ""
4960    ESDeviceType(j).Announce = ""
4970    ESDeviceType(j).Checkin = 9

4980    j = j + 1
4990    ReDim Preserve ESDeviceType(j)
5000    ESDeviceType(j).Model = "ES3942XS"
5010    ESDeviceType(j).desc = "Two Way End Device Serial Input"
5020    ESDeviceType(j).MID = &HB2
5030    ESDeviceType(j).PTI = &H73
5040    ESDeviceType(j).CLS = &H3E
5050    ESDeviceType(j).BiDi = True
5060    ESDeviceType(j).NumInputs = 1
5070    ESDeviceType(j).Notes = ""
5080    ESDeviceType(j).Announce = ""
5090    ESDeviceType(j).Checkin = 9

5100    j = j + 1
5110    ReDim Preserve ESDeviceType(j)  ' PCA
5120    ESDeviceType(j).Model = "EN3954"
5130    ESDeviceType(j).desc = "Two Way Communicator"
5140    ESDeviceType(j).MID = &HB2
5150    ESDeviceType(j).PTI = &H5B
5160    ESDeviceType(j).CLS = &H39
5170    ESDeviceType(j).BiDi = True
5180    ESDeviceType(j).NumInputs = 0
5190    ESDeviceType(j).Notes = ""
5200    ESDeviceType(j).Announce = ""
5210    ESDeviceType(j).NoTamper = 1
5220    PCA_DEV_NAME = ESDeviceType(j).Model
5230    ESDeviceType(j).Checkin = 9

5240    If USE6080 = 0 Then
5250      j = j + 1
5260      ReDim Preserve ESDeviceType(j)  ' this and the EN5040 are the same, with the 5040 going in to DN mode
5270      ESDeviceType(j).Model = "EN5000"
5280      ESDeviceType(j).desc = "High Power Repeater (BC)"
5290      ESDeviceType(j).MID = &H1
5300      ESDeviceType(j).PTI = &H0
5310      ESDeviceType(j).CLS = &H41
5320      ESDeviceType(j).BiDi = True
5330      ESDeviceType(j).NumInputs = 0
5340      ESDeviceType(j).Notes = ""
5350      ESDeviceType(j).Announce = ""
5360      ESDeviceType(j).Checkin = 9
5370    End If

5380    j = j + 1
5390    ReDim Preserve ESDeviceType(j)
5400    ESDeviceType(j).Model = "EN5040"
5410    ESDeviceType(j).desc = "High Power Repeater"
5420    ESDeviceType(j).MID = &H1

5430    If USE6080 = 0 Then
5440      ESDeviceType(j).PTI = &H1
5450    Else
5460      ESDeviceType(j).PTI = &H1  ' always '10  ' decimal
5470    End If

5480    ESDeviceType(j).CLS = &H41
5490    ESDeviceType(j).BiDi = True
5500    ESDeviceType(j).NumInputs = 0
5510    ESDeviceType(j).Notes = ""
5520    ESDeviceType(j).Announce = ""
5530    ESDeviceType(j).Checkin = 9

5540    j = j + 1
5550    ReDim Preserve ESDeviceType(j)
5560    ESDeviceType(j).Model = COM_DEV_NAME
5570    ESDeviceType(j).desc = "Serial Input"
5580    ESDeviceType(j).MID = &H0
5590    ESDeviceType(j).PTI = &HFF
5600    ESDeviceType(j).CLS = &H0
5610    ESDeviceType(j).NumInputs = 1
5620    ESDeviceType(j).Notes = ""
5630    ESDeviceType(j).SerialDevice = 1
5640    ESDeviceType(j).Announce = "Serial Event"
5650    ESDeviceType(j).NoTamper = 1



5660    j = j + 1
5670    ReDim Preserve ESDeviceType(j)
5680    ESDeviceType(j).Model = "Dukane 3"
5690    ESDeviceType(j).desc = "Dukane 3"
5700    ESDeviceType(j).MID = &HD3
5710    ESDeviceType(j).PTI = &HDD
5720    ESDeviceType(j).CLS = &H3
5730    ESDeviceType(j).NumInputs = 1
5740    ESDeviceType(j).Notes = ""
5750    ESDeviceType(j).SerialDevice = 0
5760    ESDeviceType(j).Announce = "Alarm"
5770    ESDeviceType(j).NoTamper = 1
5780    ESDeviceType(j).Checkin = 1000
5790    ESDeviceType(j).NoCheckin = 1


5800    j = j + 1
5810    ReDim Preserve ESDeviceType(j)
5820    ESDeviceType(j).Model = "Dukane 5"
5830    ESDeviceType(j).desc = "Dukane 5"
5840    ESDeviceType(j).MID = &HD3
5850    ESDeviceType(j).PTI = &HDD
5860    ESDeviceType(j).CLS = &H5
5870    ESDeviceType(j).NumInputs = 1
5880    ESDeviceType(j).Notes = ""
5890    ESDeviceType(j).SerialDevice = 0
5900    ESDeviceType(j).Announce = "Alarm"
5910    ESDeviceType(j).NoTamper = 1
5920    ESDeviceType(j).Checkin = 1000
5930    ESDeviceType(j).NoCheckin = 1

        ' no longer used
5940    If 0 Then
5950      j = j + 1
5960      ReDim Preserve ESDeviceType(j)
5970      ESDeviceType(j).Model = "EN5080/81"
5980      ESDeviceType(j).desc = "Locator"
5990      ESDeviceType(j).MID = &H1
6000      ESDeviceType(j).PTI = &H1
6010      ESDeviceType(j).CLS = &H36
6020      ESDeviceType(j).BiDi = True
6030      ESDeviceType(j).NumInputs = 0
6040      ESDeviceType(j).Notes = ""
6050      ESDeviceType(j).Announce = ""
6060      ESDeviceType(j).Checkin = 60 * 3  ' 60 minutes
6070    End If


        ' add remote console monitoring

6080    j = j + 1

6090    ReDim Preserve ESDeviceType(j)
6100    ESDeviceType(j).Model = "REMOTE"
6110    ESDeviceType(j).desc = "Remote"
6120    ESDeviceType(j).MID = &H1
6130    ESDeviceType(j).PTI = &HEE
6140    ESDeviceType(j).CLS = &HEE
6150    ESDeviceType(j).NumInputs = 0
6160    ESDeviceType(j).Notes = ""
6170    ESDeviceType(j).Announce = "Remote Disconnected"
6180    ESDeviceType(j).Checkin = 0
6190    ESDeviceType(j).NoTamper = 1
6200    ESDeviceType(j).NoCheckin = 1


6210    j = j + 1
6220    ReDim Preserve ESDeviceType(j)
6230    ESDeviceType(j).Model = "Central Office"
6240    ESDeviceType(j).desc = "Central Office"
6250    ESDeviceType(j).MID = &HD6
6260    ESDeviceType(j).PTI = &HD6
6270    ESDeviceType(j).CLS = &H6
6280    ESDeviceType(j).NumInputs = 0
6290    ESDeviceType(j).Notes = ""
6300    ESDeviceType(j).SerialDevice = 2
6310    ESDeviceType(j).Announce = "CO Error"
6320    ESDeviceType(j).NoTamper = 1
6330    ESDeviceType(j).Checkin = 1000
6340    ESDeviceType(j).NoCheckin = 0


6350    j = j + 1
6360    ReDim Preserve ESDeviceType(j)
6370    ESDeviceType(j).Model = "SDACT2"
6380    ESDeviceType(j).desc = "SDACT2"
6390    ESDeviceType(j).MID = &HD6
6400    ESDeviceType(j).PTI = &HD6
6410    ESDeviceType(j).CLS = &H7
6420    ESDeviceType(j).NumInputs = 1
6430    ESDeviceType(j).Notes = ""
6440    ESDeviceType(j).SerialDevice = 2
6450    ESDeviceType(j).Announce = "Central Office Error"
6460    ESDeviceType(j).NoTamper = 1
6470    ESDeviceType(j).Checkin = 0
6480    ESDeviceType(j).NoCheckin = 1




6490    MAX_ESDEVICETYPES = j

6500    For j = 0 To MAX_ESDEVICETYPES
6510      ESDeviceType(j).MIDPTI = ESDeviceType(j).MID * 256& + ESDeviceType(j).PTI
6520      ESDeviceType(j).CLSPTI = ESDeviceType(j).CLS * 256& + ESDeviceType(j).PTI
6530    Next

End Sub


Public Function GetNextPacketSeq() As Long

  NextPacketSeq = NextPacketSeq + 1
  If NextPacketSeq > 2000000 Then
    NextPacketSeq = 1
  End If
  GetNextPacketSeq = NextPacketSeq

End Function



Function ProcessESPacket(packet As cESPacket) As Long

        ' for 6080, this is called from modMain.DOREAD
        ' 6080 packets are redirected to mod6080.Process6080Packet

        ' 6040 modMain -> packetizer.Process
        ' 6040 modMain.DOREAD -> ProcessESPacket packetizer.GetPacket

        Dim d                  As cESDevice
        Dim IsAlarm            As Boolean ' alarm 1
        Dim IsAlarm_A          As Boolean ' alarm 2
        Dim IsAlarm_B          As Boolean ' alarm 3
        Dim Ready              As Boolean
        Dim ArmedStatus        As String
        Dim armed              As Boolean
        Dim disarmed           As Boolean
        Dim FirstHopDevice     As cESDevice



10      If packet Is Nothing Then
20        Exit Function
30      End If


        
        
40      If packet.Is6080 Then
50        Process6080Packet packet ' 6080 packets are processed there, not here
60        Set packet = Nothing
70      Else



      

'
'          If packet.Serial = "B210326A" Then
'            Debug.Print "ProcessESPacket packet.Status "; packet.HexPacket
'          End If
'
'
'          If packet.Serial = "B2F99F19" Then
'
'          Debug.Print "ProcessESPacket packet.Status "; packet.Status
'
'          End If


          ' new tiny pendant is B2851BB5
80        If packet.IsLocatorPacket Then
            ' no longer using locator devices
            'dbg "Located Packet " & Packet.LocatedSerial
90        End If

100       frmMain.PacketToggle

110       Set d = Devices.Item(1) ' Network Coordinator - Always

120       d.LastSupervise = Now      ' set time for receiver com port activity
          'dbg "ProcessESPacket  " & Packet.serial
130       If d.Dead = 1 Then
140         PostEvent d, packet, Nothing, EVT_COMM_RESTORE, 0
150       End If

          'Trace packet.serial & "  " & GetESModel(packet.MIDPTI) & " " & packet.HexPacket
          'Debug.Print packet.HexPacket
          'If Len(Packet.HexPacket) < 10 Then
          '  Debug.Print Packet.HexPacket
          'End If
          
160       If gShowPacketData Then
170         Call dbgPackets(packet.HexPacket)
180       End If


190       If packet.HexPacket = "0602" Or packet.HexPacket = "060208" Then
200         Outbounds.Ready = True
210         Set packet = Nothing
220         Exit Function
230       End If

240       If packet.IsAggregatePacket Then
250         Set packet = Nothing
260         Exit Function
270       End If


280       IsAlarm = False
290       IsAlarm_A = False
300       IsAlarm_B = False

          'Debug.Assert Not (packet.Serial = "B2851BB5")
          


330       Select Case packet.CMD
            
            Case &H11
340           If packet.ClassByte = &H39 Then
350             Outbounds.ACK packet
360           End If
370         Case &H17                ' PCA Reset, coming online
380           Outbounds.AddMessage packet.Serial, MSGTYPE_SET_TIME, "", 0

390         Case Else
400           If packet.TransType = &H31 Then  ' standard ES
410             If packet.SubCommand = &H7 Then
420               GlobalNID = packet.NID
430             End If
440           ElseIf packet.TransType = &H35 Then  ' DN
450             Outbounds.Ready = True
460             If packet.SubCommand = &H82 Then
470               GlobalNID = packet.NID  ' Gets NID
480             ElseIf packet.SubCommand = &H90 Then
490               newserial = Right("00000000" & packet.Serial, 8)
500             End If

510           End If

520       End Select

530       If packet.ACKREQ Then      ' if incoming packet needs an ACK, give it one, PCA and Two-way devices
540         Outbounds.AddMessage packet.Serial, MSGTYPE_CANNEDACK, packet.CMD, 0
550       End If


560       Set d = Devices.Device(packet.Serial) ' get device from serial number

          If Not d Is Nothing Then
            If Len(d.Configurationstring) Then
                If packet.alarm <> 0 Then
                  Dim resetpacket As cESPacket
                  Set resetpacket = MakeResetPacket(d.Configurationstring)
                  If resetpacket Is Nothing Then
                    ' just eat it
                    Exit Function
                  Else
                    ProcessESPacket resetpacket
                    Exit Function
                  End If
                    
                End If
            End If
          End If

          ' *********************  Handle survey devices if active
          ''' Debug.Print "Device, First Hop serial " & packet.Serial & "  " & packet.FirstHopSerial

          'Debug.Assert Not (packet.FirstHopSerial = "01219E4A")

          '016CC62D
          '0020650E ' RX


570       If packet.FirstHopSerial = Configuration.RxSerial Then
580         If Configuration.NoNCs Then
              ' skipit
590         Else
600           Set FirstHopDevice = Devices.Device(packet.FirstHopSerial)
610           If FirstHopDevice Is Nothing Then
620             Set FirstHopDevice = Devices.Item(1)
630           End If
640         End If
650       Else

660         Set FirstHopDevice = Devices.Device(packet.FirstHopSerial)
670       End If

680       If SurveyEnabled Or BatchSurveyEnabled Then
690         If packet.Serial = Configuration.SurveyDevice Then
              'Debug.Assert 0
700           Debug.Print "If SurveyEnabled Or BatchSurveyEnabled Then "; packet.Serial
              ' if I recall, they discontinued the locator devices.
710           If packet.IsLocatorPacket Then  ' if it's a locator packet then...
720             If (Not d Is Nothing) Then  ' it (the locator) must also be an enrolled device
730               If SurveyEnabled Then

740                 AutoSurvey packet
750               End If
760               If BatchSurveyEnabled Then
770                 BatchSurvey packet
780               End If
790             End If
'             If packet.LocatedSerial = Configuration.SurveyDevice Then
'                  '              InBounds.ProcessPacket Packet
'                  '            Exit Function
'             End If

820           ElseIf (Not FirstHopDevice Is Nothing) Then  ' the first hop must have been enrolled
830             'Debug.Print "ElseIf (Not FirstHopDevice Is Nothing) Then "; packet.Serial

840             If SurveyEnabled Then
850               AutoSurvey packet
860             End If
870             If BatchSurveyEnabled Then
880               BatchSurvey packet
890             End If
900             If packet.Serial = Configuration.SurveyDevice Then
                  '              InBounds.ProcessPacket Packet
                  '              Exit Function
910             End If
920           End If
930         End If
940       End If
          ' *********************  Handle AutoEnroll and Strays



950       If d Is Nothing Then
            ' either not entered into system, or awaiting Enrollment
            ' first we need to decide if it's Es or Directed Network
960         If packet.Serial <> "00000000" Or packet.RegisterPCA <> 0 Then
970           If AutoEnrollEnabled Or RemoteAutoEnroller.RemoteEnrollEnabled Then
980             If (packet.ClassByte = &H3E) Or (packet.ClassByte = &H3C) Then  'security device or temperature 1723 device
990               If packet.Reset Then
1000                packet.Registering = 1
1010                AutoEnroll packet
1020                Set packet = Nothing
1030                Exit Function

1040              End If
1050            End If
                ' handle Repeaters
1060            If packet.CONFIGACK Then
1070              AutoEnroll packet
1080              Set packet = Nothing
1090              Exit Function
1100            End If
1110            If packet.HardReset Then
                  'dbg "Sent Configure NID as a response to a hard reset with mismatched NID"
1120              Outbounds.AddMessage packet.Serial, MSGTYPE_REPEATERNID, "", 0
1130              Set packet = Nothing
1140              Exit Function
1150            End If
1160            If packet.RegisterPCA = 1 Then  ' we're registering a PCA
1170              AutoEnroll packet
1180              Outbounds.AddMessage packet.Serial, MSGTYPE_TWOWAYNID, "", 0
1190              Outbounds.AddMessage packet.Serial, MSGTYPE_SET_TIME, "", 0
1200              Set packet = Nothing
1210              Exit Function
1220            End If
                ' not a PCA
1230            If packet.ClassByte = 54 Or packet.Reset = 0 Then  ' &h36 for locator
1240              If packet.DataLength = 16 Then  ' short packet &h10
1250                If gDirectedNetwork Then
1260                  Outbounds.AddMessage packet.Serial, MSGTYPE_TWOWAYNID, "", 0
1270                  Set packet = Nothing
1280                  Exit Function
1290                Else
1300                  AutoEnroll packet
1310                  Outbounds.AddMessage packet.Serial, MSGTYPE_TWOWAYNID, "", 0
1320                End If
1330              ElseIf packet.Reset = 1 Then  '
1340                If gDirectedNetwork Then
1350                  AutoEnroll packet
                      'Outbounds.AddMessage Packet.serial, MSGTYPE_TWOWAYNID, "", 0
1360                End If
1370              End If

1380            ElseIf packet.Reset = 1 Then  ' real packet
1390              If gDirectedNetwork Then
                    'If Packet.ClassByte = &H41 Then
1400                If packet.ClassByte = &H0 Then
1410                  If packet.MID = 1 Then
                        ' not our code
                        ' interim code as NON-directed NET
                        'Outbounds.AddMessage Packet.serial, MSGTYPE_TWOWAYNID, "", 0
1420                    AutoEnroll packet
1430                  Else
1440                    AutoEnroll packet
1450                  End If
1460                Else
1470                  AutoEnroll packet
1480                End If
1490              Else
1500                AutoEnroll packet
1510              End If
1520            End If
1530          Else                   ' *********************  Handle Strays
1540            PostEvent d, packet, Nothing, EVT_STRAY, packet.inputnum
1550          End If
1560        End If

1570      Else                       '******************* d is a valid device! ***************

1580        If (d.Ignored <> 0) Then
1590          If d.Alarm_A And packet.Alarm0 = 0 Then
                ' let it pass, need to clear alarm
1600          ElseIf d.Alarm_B And packet.Alarm1 = 0 Then
                ' let it pass, need to clear alarm
1610          Else
                ' eat it
1620            Set packet = Nothing

1630            Exit Function
1640          End If
1650        End If

1660        d.FirstHop = packet.FirstHopSerial
1670        d.LastLevel = packet.LEvel
1680        d.LastMargin = packet.Margin
1690        d.Layer = packet.LayerID
1700        d.IncrementJam packet.Jammed



1710        If (packet.CONFIGACK) Then
1720          d.LastConfigResponse = Now
1730        End If
1740        If d.Model <> "REMOTE" Then
1750          d.LastSeen = Now
1760        End If

            ' ********************** FILTER CERTAIN PACKETS **********************
1770        If packet.DataLength = 16 Then  ' short packet &h10
1780          If packet.ClassByte = 0 Then
1790            If packet.Reset <> 0 Then
                  'If Packet.MID = 0 Or Packet.MID = 1 Then
1800              d.NID = packet.NID
                  'End If
1810            End If
1820          End If
1830          Set packet = Nothing
1840          Exit Function
1850        End If

            ' by virtue of getting this far, the packet is from an enrolled device
1860        If packet.IsLocatorPacket Then
1870          InBounds.ProcessPacket packet
1880          Set packet = Nothing
1890          Exit Function
1900        End If


'1910        If packet.PTI <> d.PTI Then  ' bad packet (Regression Fuck up )
1910        If packet.CLSPTI <> d.CLSPTI Then  ' bad packet
              
1920          PostEvent d, packet, Nothing, EVT_PTI_MISMATCH, 0
1930          dbg "mismatch"

1940          Exit Function
1950        End If

1960        If Not (CBool(d.IgnoreTamper)) Then  ' 9/4/2014 allow low battery to pass thru if tamper and battery are simultaneous with ignore tamper set

1970          If packet.Battery = 1 And packet.Tamper = 1 Then  ' this must have been for something flaky

1980            PostEvent d, packet, Nothing, EVT_BATT_TAMPER, 0
1990            Set packet = Nothing
2000            Exit Function
2010          End If

2020        End If

            ' *********** process device alarms etc here ***********

2030        d.LastSupervise = Now    ' packet.DateTime  ' checkin supervise update
2040        If packet.RegisterPCA <> 0 Then
2050          PostEvent d, packet, Nothing, EVT_PCA_REG, 0
2060          Set packet = Nothing
2070          Exit Function
2080        End If

            ' ************* Fetch Temperature Alarms State New stuff

2090        If packet.ClassByte = &H3C Then
              ' could maybe streamline this
2100          d.Temperature = packet.Temperature0
2110          d.Temperature_A = packet.Temperature1
2120          Call packet.SetAlarm0(d.TemperatureAlarm)
2130          Call packet.SetAlarm1(d.TemperatureAlarm_a)
              ' limited to 2 inputs

2140        End If


            ' ************* handle assure events for input #1, or pass thru as alarm


2150        IsAlarm = True
2160        If d.IsAway = 0 Then     ' if it's not on vacation then
2170          If d.AssurInput <= 1 Then
2180            If d.Assur = 1 Then  ' an assurance device
2190              IsAlarm = False    ' it's an assurance checkin
2200            End If
2210          End If
2220        Else                     ' if it's on vacation
2230          If d.AssurInput <= 1 Then
2240            If d.Assur = 1 Then  ' it's an assurance device
2250              If d.AssurSecure = 0 Then  ' they took it home
2260                IsAlarm = False  ' ignore... not supposed to check in if away
2270              End If
2280            End If
2290          End If
2300        End If

            ' ************* handle assure events for input #2, or pass thru as alarm
2310        IsAlarm_A = True
2320        If d.IsAway = 0 Then     ' if it's not on vacation then
2330          If d.AssurInput = 2 Then
2340            If d.Assur = 1 Then  ' an assurance device
2350              IsAlarm_A = False  ' it's an assurance checkin
2360            End If
2370          End If
2380        Else                     ' if it's on vacation
2390          If d.AssurInput = 2 Then
2400            If d.Assur = 1 Then  ' it's an assurance device
2410              If d.AssurSecure_A = 0 Then  ' they took it home
2420                IsAlarm_A = False  ' ignore... not supposed to check in if away
2430              End If
2440            End If
2450          End If
2460        End If

            ' ************* handle assure events for input #3, or pass thru as alarm

2470        IsAlarm_B = True
2480        If d.IsAway = 0 Then     ' if it's not on vacation then
2490          If d.AssurInput = 3 Then
2500            If d.Assur = 1 Then  ' an assurance device
2510              IsAlarm_B = False  ' it's an assurance checkin
2520            End If
2530          End If
2540        Else                     ' if it's on vacation
2550          If d.AssurInput = 3 Then
2560            If d.Assur = 1 Then  ' it's an assurance device
2570              If d.AssurSecure_B = 0 Then  ' they took it home
2580                IsAlarm_B = False  ' ignore... not supposed to check in if away
2590              End If
2600            End If
2610          End If
2620        End If

2630        If d.isDisabled Then
2640          'IsAlarm = False
2650        End If
2660        If d.isDisabled_A Then
2670          'IsAlarm_A = False
2680        End If
2690        If d.isDisabled_B Then
2700          'IsAlarm_B = False
2710        End If

            'Trace "IsAlarma " & str(IsAlarm_a) & "  InputNum " & packet.inputnum
            ' need to do this for both isalarm and isalarm_A



2720        If (IsAlarm = False) Then  ' this is an assurance event, not an alarm
2730          If d.AssurInput = 1 Then
2740            If d.AssurBit = 1 Then
2750              If packet.Alarm0 = 1 Then
2760                d.AssurBit = 0
2770                PostEvent d, packet, Nothing, EVT_ASSUR_CHECKIN, 1
2780              End If
2790            End If
2800          End If

2810        ElseIf (IsAlarm) Then
              ' Handle Alarm Bit  ' added isaway 9/8/2011
              ' removed is away 7/24/14 per Jerry
2820          If CBool(packet.Alarm0) And (d.isDisabled = 0) Then   ' And (d.IsAway = 0) Then    ' check status and last alarm for this device


2830            If d.alarm = 0 Then  ' alarm0
2840              Ready = DateDiff("s", d.LastAlarm, packet.DateTime) > gWindow  ' default window is 15 seconds
2850              If Ready Or d.IsSerialDevice Then
                    ' New Alarm since we get multiple hits
                    'If d.isDisabled Then
                    
2860                Select Case d.AlarmMask
                      Case 2
2870                    PostEvent d, packet, Nothing, EVT_EXTERN, 1
2880                  Case 1
2890                    PostEvent d, packet, Nothing, EVT_ALERT, 1
2900                  Case Else
2910                    'Trace " EVT_EMERGENCY 1"
2920                    PostEvent d, packet, Nothing, EVT_EMERGENCY, 1
2930                End Select
2940              End If
2950            End If

2960          ElseIf packet.Alarm0 = 0 And d.ClearByReset Then
2970            If packet.Reset Then
                  'If D.Alarm = 0 Then
2980              d.alarm = 0
2990              If DateDiff("s", d.LastRestore, packet.DateTime) > 0 Then  ' > gWindow Then
3000                If d.AlarmMask = 1 Then
3010                  PostEvent d, packet, Nothing, EVT_ALERT_RESTORE, 1
3020                Else
3030                  'Trace " EVT_EMERGENCY_RESTORE 1"
3040                  PostEvent d, packet, Nothing, EVT_EMERGENCY_RESTORE, 1
3050                End If
3060              End If
                  'End If
3070            End If
3080          ElseIf packet.Alarm0 = 0 Then
3090            If d.alarm = 1 Then
3100              If DateDiff("s", d.LastRestore, packet.DateTime) > 0 Then  ' > gWindow Then
3110                d.alarm = 0
3120                If d.AlarmMask = 1 Then  ' either 0 for emergency, or 1 for alert
3130                  PostEvent d, packet, Nothing, EVT_ALERT_RESTORE, 1
3140                Else
3150                  'Trace " EVT_EMERGENCY_RESTORE 1"
3160                  PostEvent d, packet, Nothing, EVT_EMERGENCY_RESTORE, 1
3170                End If
3180              End If
3190            End If

3200          End If

3210        End If
            ' end of alarm (0)


            'isalarm_A
3220        If (Not IsAlarm_A) Then  ' this is an assurance event, not an alarm
3230          If d.AssurInput = 2 Then
3240            If d.AssurBit = 1 Then
3250              If packet.Alarm1 = 1 Then
3260                d.AssurBit = 0
3270                PostEvent d, packet, Nothing, EVT_ASSUR_CHECKIN, 2
3280              End If
3290            End If
3300          End If

3310        ElseIf IsAlarm_A Then
              ' Handle Alarm Bit  ' added isaway 9/8/2011
              ' removed is away 7/24/14 per Jerry
3320          If CBool(packet.Alarm1) And (d.isDisabled = 0) Then     'And (d.IsAway = 0) Then  ' check status and last alarm for this device
3330            If d.Alarm_A = 0 Then
3340              Ready = DateDiff("s", d.LastAlarm_A, packet.DateTime) > gWindow
3350              If Ready Then      ' default window is 15 seconds
                    ' New Alarm since we get multiple hits
3360                If d.AlarmMask_A = 1 Then
3370                  PostEvent d, packet, Nothing, EVT_ALERT, 2
3380                Else
3390                  'Trace " EVT_EMERGENCY 2"
3400                  PostEvent d, packet, Nothing, EVT_EMERGENCY, 2
3410                End If
3420              End If
3430            End If

3440          ElseIf packet.Alarm1 = 0 And d.ClearByReset Then
3450            If packet.Reset Then
3460              If d.Alarm_A <> 0 Then
3470                d.Alarm_A = 0
3480                If DateDiff("s", d.LastRestore_A, packet.DateTime) > 0 Then  ' > gWindow Then
3490                  If d.AlarmMask_A = 1 Then
3500                    PostEvent d, packet, Nothing, EVT_ALERT_RESTORE, 2
3510                  Else
3520                    'Trace " EVT_EMERGENCY_RESTORE 2"
3530                    PostEvent d, packet, Nothing, EVT_EMERGENCY_RESTORE, 2
3540                  End If
3550                End If
3560              End If
3570            End If
3580          ElseIf packet.Alarm1 = 0 Then
3590            If d.Alarm_A = 1 Then
3600              If DateDiff("s", d.LastRestore_A, packet.DateTime) > 0 Then  '  gWindow Then
3610                d.Alarm_A = 0
3620                If d.AlarmMask_A = 1 Then
3630                  PostEvent d, packet, Nothing, EVT_ALERT_RESTORE, 2
3640                Else
3650                  'Trace " EVT_EMERGENCY_RESTORE 2"
3660                  PostEvent d, packet, Nothing, EVT_EMERGENCY_RESTORE, 2
3670                End If
3680              End If
3690            End If
3700          End If
3710        End If

            ' ************************ new for using tamper as 3rd input ************************ 12/2013
3720        If CBool(d.UseTamperAsInput) Then  ' NOT using tamper as input 3
3730          packet.Alarm2 = packet.Tamper
3740          If (Not IsAlarm_B) Then    ' this is an assurance event, not an alarm
3750            If d.AssurInput = 3 Then
3760              If d.AssurBit = 1 Then
3770                If packet.Alarm2 And 1 Then
3780                  d.AssurBit = 0
3790                  PostEvent d, packet, Nothing, EVT_ASSUR_CHECKIN, 3
3800                End If
3810              End If
3820            End If

3830          ElseIf IsAlarm_B Then
                ' Handle Alarm Bit  ' added isaway 9/8/2011
                ' removed is away 7/24/14 per Jerry
3840            If CBool(packet.Alarm2) And (d.isDisabled = 0) Then   ' And (d.IsAway = 0) Then  ' check status and last alarm for this device
3850              If d.Alarm_B = 0 Then
3860                Ready = DateDiff("s", d.LastAlarm_B, packet.DateTime) > gWindow
3870                If Ready Then    ' default window is 15 seconds
                      ' New Alarm since we get multiple hits
3880                  If d.AlarmMask_B = 1 Then
3890                    PostEvent d, packet, Nothing, EVT_ALERT, 3
3900                  Else
3910                    'Trace " EVT_EMERGENCY 2"
3920                    PostEvent d, packet, Nothing, EVT_EMERGENCY, 3
3930                  End If
3940                End If
3950              End If

3960            ElseIf packet.Alarm2 = 0 And d.ClearByReset Then
3970              If packet.Reset Then
3980                If d.Alarm_B <> 0 Then
3990                  d.Alarm_B = 0
4000                  If DateDiff("s", d.LastRestore_B, packet.DateTime) > 0 Then  ' > gWindow Then
4010                    If d.AlarmMask_B = 1 Then
4020                      PostEvent d, packet, Nothing, EVT_ALERT_RESTORE, 3
4030                    Else
4040                      'Trace " EVT_EMERGENCY_RESTORE 3"
4050                      PostEvent d, packet, Nothing, EVT_EMERGENCY_RESTORE, 3
4060                    End If
4070                  End If
4080                End If
4090              End If
4100            ElseIf packet.Alarm2 = 0 Then
4110              If d.Alarm_B = 1 Then
4120                If DateDiff("s", d.LastRestore_B, packet.DateTime) > 0 Then  '  gWindow Then
4130                  d.Alarm_B = 0
4140                  If d.AlarmMask_B = 1 Then
4150                    PostEvent d, packet, Nothing, EVT_ALERT_RESTORE, 3
4160                  Else
4170                    'Trace " EVT_EMERGENCY_RESTORE 3"
4180                    PostEvent d, packet, Nothing, EVT_EMERGENCY_RESTORE, 3
4190                  End If
4200                End If
4210              End If
4220            End If
4230          End If



4240        Else                     ' NOT using tamper as input 3
4250          If packet.Tamper = 1 Then  ' check status and last tamper for this device
4260            If d.Tamper = 0 Then
4270              If DateDiff("s", d.LastTamper, packet.DateTime) > gWindow Then  ' New Alarm since we get multiple hits
4280                PostEvent d, packet, Nothing, EVT_TAMPER, 0
4290              End If
4300            End If
4310          ElseIf packet.Tamper = 0 Then
4320            If d.Tamper = 1 Then
4330              If DateDiff("s", d.LastTamperRestore, packet.DateTime) > gWindow Then
4340                PostEvent d, packet, Nothing, EVT_TAMPER_RESTORE, 0
4350              End If
4360            End If
4370          End If
4380        End If                   ' NOT using tamper as input 3

            '   End If


4390        If packet.LineLoss = 1 And d.LineLoss = 0 Then
4400          PostEvent d, packet, Nothing, EVT_LINELOSS, 0
4410        ElseIf packet.LineLoss = 0 And d.LineLoss = 1 Then
4420          PostEvent d, packet, Nothing, EVT_LINELOSS_RESTORE, 0
4430        End If



4440        If packet.Battery = 1 Then

              'dbg packet.Serial & " Low Battery Flag Set"
4450          If d.Battery = 0 Then
                ' just set the bit

4460            If d.IsPortable Then

4470              d.Battery = 1
4480            Else

4490              PostEvent d, packet, Nothing, EVT_BATTERY_FAIL, 0
4500            End If
4510          ElseIf d.Battery = 1 Then
4520            If d.BatteryDelayTimeout = True Then
4530              PostEvent d, packet, Nothing, EVT_BATTERY_FAIL, 0
4540            End If
4550          End If

4560        ElseIf packet.Battery = 0 Then
4570          If d.Battery = 1 Then
4580            PostEvent d, packet, Nothing, EVT_BATTERY_RESTORE, 0  ' if there's no alarm, then it does not log restore
4590          End If
4600        End If


            ' we were getting extranous repeaters
            ' locators are processed early on elsewhere in this function
            '10/28/08
4610        If d.IsPortable And packet.alarm <> 0 Then  ' only locate portable devices
4620          If packet.FirstHopSerial = "00000000" Or packet.FirstHopSerial = "0020640E" Then  ' skip it
4630          Else
4640            If (FirstHopDevice Is Nothing) Then
4650              Debug.Print " FirstHopDevice Is Nothing " & packet.FirstHopSerial
4660            End If
4670          End If
4680          If (Not FirstHopDevice Is Nothing) Then  ' Fisthop must be in our list of devices
4690            Debug.Print "CALL InBounds.ProcessPacket packet " & packet.Serial & " " & packet.FirstHopSerial
4700            InBounds.ProcessPacket packet
4710          End If
4720        End If
4730      End If                     ' is a valid device

4740      Set packet = Nothing
4750      Set d = Nothing
4760      Set FirstHopDevice = Nothing

4770    End If

End Function

Function digest(packet As cESPacket) As String
  Exit Function
  Dim hfile         As Integer
  Dim filename As String
  filename = App.Path & "\digest.csv"
  limitFileSize filename
  hfile = FreeFile
  Open filename For Append As hfile
  ' raw packets
  Print #hfile, _
        "x"; Right("00000000" & packet.Serial, 8); _
        ",x"; Right("00" & Hex(packet.Stat0), 2); _
        ",x"; Right("00" & Hex(packet.Stat1), 2); _
        ",x"; Right("00" & Hex(packet.FirstHopMID), 2); _
        ",x"; Right("000000" & Hex(packet.FirstHopUID), 6); _
        ","; packet.LEvel; _
        ","; packet.Margin; _
        ",t"; Format(packet.DateTime, " hh-nn-ss")
  Close hfile


End Function




Public Sub ReadESDeviceTypes()

'  UPDATE Devices SET Devices.ignoretamper = 1 WHERE model='ES1210W';


  Dim SQL           As String
  Dim rs            As ADODB.Recordset

  SQL = "SELECT * FROM DeviceTypes"
  Set rs = ConnExecute(SQL)
  Do Until rs.EOF
    SetESCustom rs
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing


End Sub
Public Function IsTemperatureDevice(ByVal CLSPTI As Long)
  IsTemperatureDevice = IIf(CLSPTI = &H3C99, 1, 0) ' hard coded for 1723
  
End Function

Public Sub SetESCustom(rs As ADODB.Recordset)
        Dim j             As Integer
10    For j = 0 To MAX_ESDEVICETYPES
20    If 0 = StrComp(ESDeviceType(j).Model, rs("model") & "", vbTextCompare) Then
30    ESDeviceType(j).desc = rs("Description") & ""
40    ESDeviceType(j).Checkin = IIf(IsNull(rs("Checkin")), 60, rs("Checkin"))
50    If ESDeviceType(j).Checkin = 0 Then  ' in minutes
60      ESDeviceType(j).Checkin = 240  ' 4 hours
70    End If
80    ESDeviceType(j).Portable = IIf(rs("IsPortable") = 0, 0, 1)
90    ESDeviceType(j).AllowDisable = IIf(rs("AllowDisable") = 0, 0, 1)
100   ESDeviceType(j).Announce = rs("Announce") & ""
110   ESDeviceType(j).Announce2 = rs("Announce2") & ""
120   ESDeviceType(j).ClearByReset = IIf(rs("ClearByReset") = 1, 1, 0)
      ' new 4/12/07
130   ESDeviceType(j).AutoClear = Val("" & rs("AutoClear"))

      'new with build 226
140   ESDeviceType(j).Repeats = Val("" & rs("Repeats"))
150   ESDeviceType(j).Pause = Val("" & rs("Pause"))
160   ESDeviceType(j).RepeatUntil = IIf(rs("repeatuntil") = 1, 1, 0)
170   ESDeviceType(j).SendCancel = IIf(rs("SendCancel") = 1, 1, 0)
180   ESDeviceType(j).IgnoreTamper = IIf(rs("IgnoreTamper") = 1, 1, 0)


190   ESDeviceType(j).Repeats_A = Val("" & rs("Repeats_A"))
200   ESDeviceType(j).Pause_A = Val("" & rs("Pause_A"))
210   ESDeviceType(j).RepeatUntil_A = IIf(rs("repeatuntil_A") = 1, 1, 0)
220   ESDeviceType(j).SendCancel_A = IIf(rs("SendCancel_A") = 1, 1, 0)


230   ESDeviceType(j).OG1 = Val(rs("OG1") & "")
240   ESDeviceType(j).OG2 = Val(rs("OG2") & "")
250   ESDeviceType(j).OG3 = Val(rs("OG3") & "")
260   ESDeviceType(j).OG4 = Val(rs("OG4") & "")
270   ESDeviceType(j).OG5 = Val(rs("OG5") & "")
280   ESDeviceType(j).OG6 = Val(rs("OG6") & "")

290   ESDeviceType(j).OG1D = Val(rs("OG1d") & "")
300   ESDeviceType(j).OG2D = Val(rs("OG2d") & "")
310   ESDeviceType(j).OG3D = Val(rs("OG3d") & "")
320   ESDeviceType(j).OG4D = Val(rs("OG4d") & "")
330   ESDeviceType(j).OG5D = Val(rs("OG5d") & "")
340   ESDeviceType(j).OG6D = Val(rs("OG6d") & "")


350   ESDeviceType(j).OG1_A = Val(rs("OG1_A") & "")
360   ESDeviceType(j).OG2_A = Val(rs("OG2_A") & "")
370   ESDeviceType(j).OG3_A = Val(rs("OG3_A") & "")
380   ESDeviceType(j).OG4_A = Val(rs("OG4_A") & "")
390   ESDeviceType(j).OG5_A = Val(rs("OG5_A") & "")
400   ESDeviceType(j).OG6_A = Val(rs("OG6_A") & "")

410   ESDeviceType(j).OG1_AD = Val(rs("OG1_Ad") & "")
420   ESDeviceType(j).OG2_AD = Val(rs("OG2_Ad") & "")
430   ESDeviceType(j).OG3_AD = Val(rs("OG3_Ad") & "")
440   ESDeviceType(j).OG4_AD = Val(rs("OG4_Ad") & "")
450   ESDeviceType(j).OG5_AD = Val(rs("OG5_Ad") & "")
460   ESDeviceType(j).OG6_AD = Val(rs("OG6_Ad") & "")



470   ESDeviceType(j).NG1 = Val(rs("NG1") & "")
480   ESDeviceType(j).NG2 = Val(rs("NG2") & "")
490   ESDeviceType(j).NG3 = Val(rs("NG3") & "")
500   ESDeviceType(j).NG4 = Val(rs("NG4") & "")
510   ESDeviceType(j).NG5 = Val(rs("NG5") & "")
520   ESDeviceType(j).NG6 = Val(rs("NG6") & "")

530   ESDeviceType(j).NG1D = Val(rs("NG1d") & "")
540   ESDeviceType(j).NG2D = Val(rs("NG2d") & "")
550   ESDeviceType(j).NG3D = Val(rs("NG3d") & "")
560   ESDeviceType(j).NG4D = Val(rs("NG4d") & "")
570   ESDeviceType(j).NG5D = Val(rs("NG5d") & "")
580   ESDeviceType(j).NG6D = Val(rs("NG6d") & "")

590   ESDeviceType(j).NG1_A = Val(rs("NG1_A") & "")
600   ESDeviceType(j).NG2_A = Val(rs("NG2_A") & "")
610   ESDeviceType(j).NG3_A = Val(rs("NG3_A") & "")
620   ESDeviceType(j).NG4_A = Val(rs("NG4_A") & "")
630   ESDeviceType(j).NG5_A = Val(rs("NG5_A") & "")
640   ESDeviceType(j).NG6_A = Val(rs("NG6_A") & "")


650   ESDeviceType(j).NG1_AD = Val(rs("NG1_Ad") & "")
660   ESDeviceType(j).NG2_AD = Val(rs("NG2_Ad") & "")
670   ESDeviceType(j).NG3_AD = Val(rs("NG3_Ad") & "")
680   ESDeviceType(j).NG4_AD = Val(rs("NG4_Ad") & "")
690   ESDeviceType(j).NG5_AD = Val(rs("NG5_Ad") & "")
700   ESDeviceType(j).NG6_AD = Val(rs("NG6_Ad") & "")


710   ESDeviceType(j).GG1 = Val(rs("gG1") & "")
720   ESDeviceType(j).GG2 = Val(rs("gG2") & "")
730   ESDeviceType(j).GG3 = Val(rs("gG3") & "")
740   ESDeviceType(j).GG4 = Val(rs("gG4") & "")
750   ESDeviceType(j).GG5 = Val(rs("gG5") & "")
760   ESDeviceType(j).GG6 = Val(rs("gG6") & "")

770   ESDeviceType(j).GG1D = Val(rs("gG1d") & "")
780   ESDeviceType(j).GG2D = Val(rs("gG2d") & "")
790   ESDeviceType(j).GG3D = Val(rs("gG3d") & "")
800   ESDeviceType(j).GG4D = Val(rs("gG4d") & "")
810   ESDeviceType(j).GG5D = Val(rs("gG5d") & "")
820   ESDeviceType(j).GG6D = Val(rs("gG6d") & "")

830   ESDeviceType(j).GG1_A = Val(rs("gG1_A") & "")
840   ESDeviceType(j).GG2_A = Val(rs("gG2_A") & "")
850   ESDeviceType(j).GG3_A = Val(rs("gG3_A") & "")
860   ESDeviceType(j).GG4_A = Val(rs("gG4_A") & "")
870   ESDeviceType(j).GG5_A = Val(rs("gG5_A") & "")
880   ESDeviceType(j).GG6_A = Val(rs("gG6_A") & "")


890   ESDeviceType(j).GG1_AD = Val(rs("gG1_Ad") & "")
900   ESDeviceType(j).GG2_AD = Val(rs("gG2_Ad") & "")
910   ESDeviceType(j).GG3_AD = Val(rs("gG3_Ad") & "")
920   ESDeviceType(j).GG4_AD = Val(rs("gG4_Ad") & "")
930   ESDeviceType(j).GG5_AD = Val(rs("gG5_Ad") & "")
940   ESDeviceType(j).GG6_AD = Val(rs("gG6_Ad") & "")



950   Exit For
960   End If

970   Next
980   If j > MAX_ESDEVICETYPES + 1 Then
990     LogGeneric EVT_MAXDEVICE
1000  End If
End Sub



Public Function GetESModel(ByVal CLSPTI As Long) As String
'Public Function GetESModel(ByVal MIDPTI As Long) As String
  Dim j             As Integer
  GetESModel = "Unknown"
  For j = 0 To MAX_ESDEVICETYPES
    If ESDeviceType(j).CLSPTI = CLSPTI Then
      '    If ESDeviceType(j).MIDPTI = MIDPTI Then
      GetESModel = ESDeviceType(j).Model
      Exit For
    End If
  Next
End Function

'Public Function GetESDesc(ByVal MIDPTI As Long) As String
Public Function GetESDesc(ByVal CLSPTI As Long) As String
  Dim j             As Integer
  GetESDesc = "Unknown"
  For j = 0 To MAX_ESDEVICETYPES
    If ESDeviceType(j).CLSPTI = CLSPTI Then
      'If ESDeviceType(j).MIDPTI = MIDPTI Then
      GetESDesc = ESDeviceType(j).desc
      Exit For
    End If
  Next
End Function

Public Function GetESDescFromModel(ByVal Model As String) As String
  Dim j             As Integer
  
  If InStr(1, Model, "ES1242", vbTextCompare) Then
    Model = "ES1242"
  End If
  
  
  GetESDescFromModel = "Unknown"
  For j = 0 To MAX_ESDEVICETYPES
    If ESDeviceType(j).Model = Model Then
      GetESDescFromModel = ESDeviceType(j).desc
      Exit For
    End If
  Next

End Function



Public Function AutoEnroll(packet As cESPacket) As Boolean
  If AutoEnrollEnabled Or RemoteAutoEnroller.RemoteEnrollEnabled Then
    If packet.Registering Or packet.RegisterPCA Or packet.CONFIGACK Then
      If MASTER Then
        If RemoteAutoEnroller.RemoteEnrollEnabled Then
          'dbg "Auto Enrolling Packet " & packet.Serial & vbCrLf
          RemoteAutoEnroller.AutoEnroll packet
        Else
          frmTransmitter.AutoEnroll packet
        End If
      End If
    End If
  End If
End Function


Public Function BatchSurvey(packet As cESPacket) As Boolean
' if doing a batch, it will come here

  If Configuration.NoNCs Then
    If packet.FirstHopSerial = Configuration.RxSerial Then
      Exit Function
    End If
  End If

  If BatchSurveyEnabled Then
    If BatchForm Is Nothing Then
      BatchSurveyEnabled = False
    Else
      BatchForm.ProcessPacket packet
    End If
  End If

End Function
Public Function AutoSurvey(packet As cESPacket) As Boolean
' if doing a single, it will come here
  If Configuration.NoNCs Then
    If packet.FirstHopSerial = Configuration.RxSerial Then
      Exit Function
    End If
  End If
  
  
  If SurveyEnabled Then
    If WayPointForm Is Nothing Then
      SurveyEnabled = False
    Else
      WayPointForm.ProcessPacket packet
      Debug.Print "WayPointForm.ProcessPacket packet ", packet.Serial
    End If
  End If

End Function
Public Function CreateDeleteAll(ByVal Serial As String) As String

'[0x41] - Delete all messages command.
'[MsgSeq (2)]  Application controller


  Dim Header        As String
  Dim Msglen        As String
  Dim RadioType     As String
  Dim UID           As String
  Dim MsgClass      As String
  Dim DeliveryCode  As String
  Dim commandbyte   As String
  Dim Sequence      As String
  Dim Content       As String

  Dim message       As String
  Dim numbytes      As Integer


  Header = "50"  ' Outbound Broadcast
  Msglen = "00"  ' to be calculated
  RadioType = "00"  ' PCA/2Way
  UID = Right("00000000" & Serial, 8)  ' Destination Serial 4 bytes
  '[Payload]
  MsgClass = "19"  ' PCA
  DeliveryCode = "00"  ' No ACK, No Group
  '[CMD] - Command sent to application controller.
  '[CMD Response] - 01: Command understood; 00: Invalid command].
  commandbyte = "41"  ' DELETE

  Sequence = GetNextMessageID()  ' Get next unique message ID "nnnn"

  message = Header & Msglen & RadioType & UID & MsgClass & DeliveryCode & commandbyte & Sequence & Content


  numbytes = Len(message) / 2
  Mid(message, 3, 2) = Right("00" & Hex(numbytes), 2)  ' insert message length

  CreateDeleteAll = message & HexChecksum(message)



  '[0x50]           Header for broadcast outbound message.
  '[LEN]            Message length, excluding checksum.
  '[0x00] -         Radio type: Enhanced two-way
  '[UID (4) dest]   The unique ID of the destination node.
  '[Payload]        The maximum length of the payload is 90 bytes, including all message bytes below.
  '     [0x19]      Message class byte for messages to the PCA.
  '     [Delivery Code] Identifies specific protocol for message delivery.
  '           Bit 7 Set if Node ACK is requested.
  '           Bit 6 Reserved.
  '           Bits 0-5 Group code (0-63), to address message to a group of PCAs.
  '     [CMD]       Command Byte.
  '     [MsgSeq (2)] Application controller generated two-byte,unique sequence number for this message.
  '     [Message content] Application-specific information. Maximum message content length is 85 bytes (90 - 5).
  '[CKSUM]  Checksum.




End Function

Public Function CreateCannedACK(ByVal Serial As String, ByVal ResponseTo As String) As String
  Dim Header        As String
  Dim Msglen        As String
  Dim RadioType     As String
  Dim UID           As String
  Dim MsgClass      As String
  Dim DeliveryCode  As String
  Dim commandbyte   As String
  Dim Content       As String
  Dim message       As String
  Dim numbytes      As Integer


  Header = "50"  ' Outbound Broadcast
  Msglen = "00"  ' to be calculated
  RadioType = "00"  ' PCA/2Way
  UID = Right("00000000" & Serial, 8)  ' Destination Serial 4 bytes
  '[Payload]
  MsgClass = "19"  ' PCA
  DeliveryCode = "00"  ' No ACK, No Group
  '[CMD] - Command sent to application controller.
  '[CMD Response] - 01: Command understood; 00: Invalid command].

  ResponseTo = Right("00" & Hex(Val(ResponseTo)), 2)

  commandbyte = "12"  ' ACK

  Content = "01"  ' Valid/recognized command message type

  message = Header & Msglen & RadioType & UID & MsgClass & DeliveryCode & commandbyte & ResponseTo & Content


  numbytes = Len(message) / 2
  Mid(message, 3, 2) = Right("00" & Hex(numbytes), 2)  ' insert message length

  CreateCannedACK = message & HexChecksum(message)



  '[0x50]           Header for broadcast outbound message.
  '[LEN]            Message length, excluding checksum.
  '[0x00] -         Radio type: Enhanced two-way
  '[UID (4) dest]   The unique ID of the destination node.
  '[Payload]        The maximum length of the payload is 90 bytes, including all message bytes below.
  '     [0x19]      Message class byte for messages to the PCA.
  '     [Delivery Code] Identifies specific protocol for message delivery.
  '           Bit 7 Set if Node ACK is requested.
  '           Bit 6 Reserved.
  '           Bits 0-5 Group code (0-63), to address message to a group of PCAs.
  '     [CMD]       Command Byte.
  '     [MsgSeq (2)] Application controller generated two-byte,unique sequence number for this message.
  '     [Message content] Application-specific information. Maximum message content length is 85 bytes (90 - 5).
  '[CKSUM]  Checksum.



End Function
Public Function CreateFieldMessage(ByVal Serial As String, ByVal MessageID As Integer, ByVal text As String) As String
' Max of 32 different field messages

  Dim Header        As String
  Dim Msglen        As String
  Dim RadioType     As String
  Dim UID           As String
  Dim MsgClass      As String
  Dim DeliveryCode  As String
  Dim commandbyte   As String
  Dim Sequence      As String
  Dim Content       As String

  Dim message       As String
  Dim numbytes      As Integer
  Dim j             As Integer

  Header = "50"  ' Outbound Broadcast
  Msglen = "00"  ' to be calculated
  RadioType = "00"  ' PCA/2Way
  UID = Right("00000000" & Serial, 8)  ' Destination Serial 4 bytes
  '[Payload]
  MsgClass = "19"  ' PCA
  DeliveryCode = "00"  ' No ACK, No Group
  commandbyte = "66"  ' Field Message
  Sequence = GetNextMessageID()  ' Get next unique message ID "nnnn"

  Content = "01" & Right("00" & Hex(MessageID), 2)
  For j = 1 To Len(text)
    If j > 14 Then Exit For
    Content = Content & Right("00" & Hex(Asc(MID(text, j, 1))), 2)
  Next
  Content = Content & "00"

  message = Header & Msglen & RadioType & UID & MsgClass & DeliveryCode & commandbyte & Sequence & Content

  numbytes = Len(message) / 2
  Mid(message, 3, 2) = Right("00" & Hex(numbytes), 2)

  CreateFieldMessage = message & HexChecksum(message)

End Function

Public Function CreateTextMessage(ByVal Serial As String, ByVal MessageID As Integer, ByVal text As String) As String
' Max of 32 different field messages

  Dim Header        As String
  Dim Msglen        As String
  Dim RadioType     As String
  Dim UID           As String
  Dim MsgClass      As String
  Dim DeliveryCode  As String
  Dim commandbyte   As String
  Dim Sequence      As String
  Dim Content       As String

  Dim message       As String
  Dim numbytes      As Integer
  Dim j             As Integer

  Header = "50"  ' Outbound Broadcast
  Msglen = "00"  ' to be calculated
  RadioType = "00"  ' PCA/2Way
  UID = Right("00000000" & Serial, 8)  ' Destination Serial 4 bytes
  '[Payload]
  MsgClass = "19"  ' PCA
  DeliveryCode = "00"  ' No ACK, No Group
  commandbyte = "66"  ' Field Message
  Sequence = GetNextMessageID()  ' Get next unique message ID "nnnn"

  Content = "01" & Right("00" & Hex(MessageID), 2)
  For j = 1 To Len(text)
    If j > 14 Then Exit For
    Content = Content & Right("00" & Hex(Asc(MID(text, j, 1))), 2)
  Next
  Content = Content & "00"

  message = Header & Msglen & RadioType & UID & MsgClass & DeliveryCode & commandbyte & Sequence & Content

  numbytes = Len(message) / 2
  Mid(message, 3, 2) = Right("00" & Hex(numbytes), 2)

  CreateTextMessage = message & HexChecksum(message)


End Function
Public Function CreateCustomCustom(ByVal Serial As String, ByVal prompt As String, ByVal ResponseCount As Integer, Responses As Collection) As String

  Dim Header        As String
  Dim Msglen        As String
  Dim RadioType     As String
  Dim UID           As String
  Dim MsgClass      As String
  Dim DeliveryCode  As String
  Dim commandbyte   As String
  Dim Sequence      As String
  Dim Content       As String

  Dim message       As String
  Dim numbytes      As Integer
  Dim j             As Integer
  Dim text          As String

  ' sample
  ' 50 22 00 B2 0F 69 50 19 00 28 00 05 4D 65 73 73 61 67 65 00 02 52 65 70 6C 79 00 43 61 6E 63 65 6C 00 4B

  Header = "50"  ' Outbound Broadcast
  Msglen = "00"  ' to be calculated
  RadioType = "00"  ' PCA/2Way
  UID = Right("00000000" & Serial, 8)  ' Destination Serial 4 bytes
  '[Payload]
  MsgClass = "19"  ' PCA
  DeliveryCode = "80"  ' No ACK, No Group
  commandbyte = "28"  ' Custom Message, Custom Response


  Sequence = GetNextMessageID()  ' Get next unique message ID "nnnn" (two bytes)

  ' 50 22 00 B2 0F 69 50 19 00 28 00 05 4D 65 73 73 61 67 65 00 02 52 65 70 6C 79 00 43 61 6E 63 65 6C 00 4B
  '                                     ^Message starts here    ^numresponses followed by responses and checksum

  For j = 1 To Len(prompt)
    'If j > 14 Then Exit For
    Content = Content & Right("00" & Hex(Asc(MID(prompt, j, 1))), 2)
  Next
  Content = Content & "00"  ' null char
  ' need to add responsecount and responses
  Content = Content & Right("00" & Hex(ResponseCount), 2)


  For ResponseCount = 1 To ResponseCount

    text = Responses(ResponseCount) & ""
    For j = 1 To Len(text)
      'If j > 14 Then Exit For
      Content = Content & Right("00" & Hex(Asc(MID(text, j, 1))), 2)
    Next
    Content = Content & "00"  ' null char
  Next
  message = Header & Msglen & RadioType & UID & MsgClass & DeliveryCode & commandbyte & Sequence & Content

  numbytes = Len(message) / 2
  ' set number of bytes in string
  Mid(message, 3, 2) = Right("00" & Hex(numbytes), 2)

  CreateCustomCustom = message & HexChecksum(message)



End Function

Public Function ConfigureRepeaterNID(ByVal Serial As String) As String

  Dim message       As String
  message = "200700" & Serial
  ConfigureRepeaterNID = message & HexChecksum(message)


  '[0x20] - Configure NID command.
  '[0x07] - Message length, excluding checksum.
  '[0x00] - Radio type, enhanced two-way.
  '[UID (4)] - Destination unique ID.
  '[CKSUM] - Checksum.


End Function
Public Function ReportNCSerial() As String
'[0x34]   Network coordinator configuration header
'[LEN]    Length of this message, 0x03
'[0x90]   subcommand to report network coordinator serial number
'[CKSUM]  (0xC7)
  ReportNCSerial = "340390C7"
End Function
Public Function SetNCNID(ByVal NID As String) As String
  Dim message       As String
  '300405E00
  If gDirectedNetwork Then
    message = "340402" & NID
    SetNCNID = message & HexChecksum(message)
  Else
    message = "300405" & NID
    SetNCNID = message & HexChecksum(message)
  End If


End Function

Public Function ConfigureTwoWayNID(ByVal Serial As String) As String
' same as ConfigureRepeaterNID
  Dim message       As String
  message = "200700" & Serial
  ConfigureTwoWayNID = message & HexChecksum(message)


  '[0x20] - Configure NID command.
  '[0x07] - Message length, excluding checksum.
  '[0x00] - Radio type, enhanced two-way.
  '[UID (4)] - Destination unique ID.
  '[CKSUM] - Checksum.


End Function
Public Function RequestTXNID() As String
  Dim message       As String
  message = "340382"
  RequestTXNID = message & HexChecksum(message)

End Function
Public Function RequestTXStatus() As String
  Dim message       As String
  message = "300307"
  RequestTXStatus = message & HexChecksum(message)
  '[0x30] - RF gateway configuration header.
  '[0x03] - Length of this message.
  '[0x07] - Subcommand to report configuration.
  '[CKSUM] - Checksum.

End Function
Public Function CreateSetTimeMsg(ByVal Serial As String) As String
  Dim Header        As String
  Dim Msglen        As String
  Dim RadioType     As String
  Dim UID           As String
  Dim MsgClass      As String
  Dim DeliveryCode  As String
  Dim commandbyte   As String
  Dim Sequence      As String
  Dim Content       As String

  Dim message       As String
  Dim numbytes      As Integer


  Header = "50"  ' Outbound Broadcast
  Msglen = "00"  ' to be calculated
  RadioType = "00"  ' PCA/2Way
  UID = Right("00000000" & Serial, 8)  ' Destination Serial 4 bytes
  '[Payload]
  MsgClass = "19"  ' PCA
  DeliveryCode = "00"  ' No ACK, No Group
  commandbyte = "40"  ' Set Date/Time
  Sequence = GetNextMessageID()  ' Get next unique message ID "nnnn"
  Content = GetHexDateTime()  ' Get current Date/Time in Hex Format "xxxxxxxx"

  message = Header & Msglen & RadioType & UID & MsgClass & DeliveryCode & commandbyte & Sequence & Content



  numbytes = Len(message) / 2
  Mid(message, 3, 2) = Right("00" & Hex(numbytes), 2)

  CreateSetTimeMsg = message & HexChecksum(message)

End Function


Public Function HexChecksum(ByVal s As String) As String
  Dim LenS          As Integer
  Dim j             As Integer
  Dim ck            As Integer

  LenS = Len(s)

  If LenS Mod 2 <> 0 Then
    ' error
  Else
    For j = 1 To LenS - 1 Step 2
      ck = ck + Val("&H" & MID(s, j, 2))
      ck = ck And &HFF
    Next
  End If
  HexChecksum = Right("00" & Hex(ck), 2)


End Function
Public Function GetNextMessageID() As String
  Dim NewID         As Long
  NewID = Configuration.ESLastMessage + 1
  If NewID >= &H7000& Then  ' decimal 28672
    ' 7ffe is the real max accepted (32766)
    NewID = 0
  End If
  Configuration.ESLastMessage = NewID
  WriteSetting "Configuration", "ESLastMessage", NewID
  GetNextMessageID = Right("0000" & Hex(NewID), 4)

End Function
Public Function GetHexDateTime() As String
'Seconds Since 2000 01 01 ' signed long good until year 2050+
  GetHexDateTime = Right("00000000" & Hex(DateDiff("s", "1/1/2000", Now)), 8)
End Function

Public Function GetRoomByID(ByVal RoomID As Long) As String


  Dim SQL           As String
  Dim rs            As Recordset
  SQL = "SELECT Room from rooms where RoomID = " & RoomID
  Set rs = ConnExecute(SQL)
  If Not rs.EOF Then
    GetRoomByID = rs(0) & ""
  End If
  rs.Close
  Set rs = Nothing

End Function

Public Function GetRoomByName(ByVal Room As String) As Long


  Dim SQL           As String
  Dim rs            As Recordset


  SQL = "SELECT RoomID from rooms where Room = " & q(Room)
  Set rs = ConnExecute(SQL)
  If Not rs.EOF Then
    GetRoomByName = IIf(IsNull(rs(0)), 0, rs(0))
  End If
  rs.Close
  Set rs = Nothing

End Function


Public Function GetRepeaterName(ByVal Serial As String) As String
  Dim d             As cESDevice

  'What condition would allow the software to list a portable device with unkown location? 10/4/12

  Set d = Devices.Device(Serial)
  If d Is Nothing Then
    GetRepeaterName = "Unknown"
  Else
    If d.RoomID <> 0 Then
      GetRepeaterName = GetRoomByID(d.RoomID)
    Else
      GetRepeaterName = d.Description
    End If
  End If
End Function

Public Function GetRepeaterRoom(ByVal Serial As String) As String
  Dim d             As cESDevice
  Set d = Devices.Device(Serial)
  If d Is Nothing Then
    GetRepeaterRoom = "Unknown"
  Else

    GetRepeaterRoom = d.Description
  End If
End Function

Function GetESDeviceTypeByModel(ByVal Model As String) As ESDeviceTypeType
  Dim j             As Long

  If InStr(1, Model, "ES1242", vbTextCompare) Then
    Model = "ES1242"
  End If

  ' default 1 button pendant
  GetESDeviceTypeByModel = ESDeviceType(gDefaultDeviceTypeIndex)

  For j = 0 To MAX_ESDEVICETYPES
    If StrComp(ESDeviceType(j).Model, Model, vbTextCompare) = 0 Then
      GetESDeviceTypeByModel = ESDeviceType(j)
      Exit For
    End If
  Next


End Function

Function GetDeviceTypeByMIDPTI(ByVal MIDPTI As Long) As ESDeviceTypeType
  Dim j             As Integer

  ' default 1 button pendant
  GetDeviceTypeByMIDPTI = ESDeviceType(gDefaultDeviceTypeIndex)

  For j = 0 To MAX_ESDEVICETYPES
    If ESDeviceType(j).MIDPTI = MIDPTI Then
      GetDeviceTypeByMIDPTI = ESDeviceType(j)
      Exit For
    End If
  Next
End Function
Function GetDeviceTypeByCLSPTI(ByVal CLSPTI As Long) As ESDeviceTypeType
  Dim j             As Integer

  ' default 1 button pendant
  GetDeviceTypeByCLSPTI = ESDeviceType(gDefaultDeviceTypeIndex)

  For j = 0 To MAX_ESDEVICETYPES
    If ESDeviceType(j).CLSPTI = CLSPTI Then
      GetDeviceTypeByCLSPTI = ESDeviceType(j)
      Exit For
    End If
  Next
End Function
Function GetCLSPTI(ByVal Model As String) As Long
  Dim j             As Integer

  ' default 1 button pendant
  
  
  GetCLSPTI = 0

  For j = 0 To MAX_ESDEVICETYPES
    If StrComp(ESDeviceType(j).Model, Model, vbTextCompare) = 0 Then
      GetCLSPTI = ESDeviceType(j).CLSPTI
      Exit For
    End If
  Next
End Function



Function GetMIDPTI(ByVal Model As String) As Long
  Dim j             As Integer

  ' default 1 button pendant
  
  
  GetMIDPTI = 0

  For j = 0 To MAX_ESDEVICETYPES
    If StrComp(ESDeviceType(j).Model, Model, vbTextCompare) = 0 Then
      GetMIDPTI = ESDeviceType(j).MIDPTI
      Exit For
    End If
  Next
End Function

Function GetDeviceTypeByModel(ByVal Model As String) As ESDeviceTypeType
  Dim j             As Integer


  If InStr(1, Model, "ES1242", vbTextCompare) Then
    Model = "ES1242"
  End If


  ' default 1 button pendant
  GetDeviceTypeByModel = ESDeviceType(gDefaultDeviceTypeIndex)

  For j = 0 To MAX_ESDEVICETYPES
    If 0 = StrComp(ESDeviceType(j).Model, Model, vbTextCompare) Then
      GetDeviceTypeByModel = ESDeviceType(j)
      Exit For
    End If
  Next
End Function



Function GetDeviceModelByMIDPTI(ByVal MIDPTI As Long) As String

  Dim DeviceType    As ESDeviceTypeType

  DeviceType = GetDeviceTypeByMIDPTI(MIDPTI)

  GetDeviceModelByMIDPTI = DeviceType.Model

End Function

Function GetDeviceModelByCLSPTI(ByVal CLSPTI As Long) As String

  Dim DeviceType    As ESDeviceTypeType

  DeviceType = GetDeviceTypeByCLSPTI(CLSPTI)

  GetDeviceModelByCLSPTI = DeviceType.Model

End Function




Public Sub SortRepeaters(Repeaters As Collection)  ' get in decending order
  Dim offset        As Integer
  Dim maxrow        As Integer
  Dim limit         As Integer
  Dim row           As Integer
  Dim switch        As Integer
  Dim MinRow        As Integer
  Dim Temp          As cRepeater


  Dim a()           As cRepeater

  maxrow = Repeaters.Count

  ReDim a(1 To maxrow)
  For row = 1 To maxrow
    Set a(row) = Repeaters(row)
  Next


  MinRow = 1

  offset = maxrow \ 2
  Do While offset > 0
    limit = maxrow - offset
    Do
      switch = 0
      For row = MinRow To limit
        'If a(row) > a(row + offset) Then
        If a(row).LEvel < a(row + offset).LEvel Then  ' may need to incorporate margin
          Set Temp = a(row)
          Set a(row) = a(row + offset)
          Set a(row + offset) = Temp
          Set Temp = Nothing
          switch = row
        End If
      Next row
      limit = switch - offset
    Loop While switch

    offset = offset \ 2
  Loop

  Set Repeaters = New Collection

  For row = 1 To maxrow
    Repeaters.Add a(row)
  Next
End Sub
Public Function CreatePageMessage(ByVal Serial As String, ByVal message As String, ByVal Alert As Boolean)

  Dim Msglen        As Integer  ' up to 90
  Dim NeedACK       As Boolean
  Dim DeliveryCode  As Byte
  Dim Buffer        As String

  NeedACK = True

  Buffer = "500000" & Serial  'Header, Length placeholder, Radiotype and serial
  Buffer = Buffer & "19"  ' PCA class byte
  DeliveryCode = DeliveryCode Or IIf(NeedACK, (BIT_7), 0)
  Buffer = Buffer & Right("00" & Hex(DeliveryCode), 2)
  Buffer = Buffer & "26"
  Buffer = Buffer & GetNextMessageID()
  '  [Message content]  Application-specific information. Maximum message content length is 85 bytes (90 - 5).
  If Len(message) > 62 Then
    message = left(message, 62)  ' chop off at 62 bytes
  End If
  Buffer = Buffer & StringToHex(message & vbNullChar)
  Buffer = Buffer & "00"  ' no reponse string

  ' get message length
  Msglen = Len(Buffer) / 2
  Mid(Buffer, 3, 2) = Right("00" & Hex(Msglen), 2)

  ' append checksum and return it all
  CreatePageMessage = Buffer & HexChecksum(Buffer)

End Function

Public Function StringToHex(ByVal text As String) As String
  Dim j             As Long
  Dim Buffer        As String

  For j = 1 To Len(text)
    Buffer = Buffer & Right("00" & Hex(Asc(MID(text, j, 1))), 2)
  Next
  StringToHex = Buffer
End Function

Public Function GetOutputMask(ByVal ScreenNumber As Long, ByVal Announce As String) As cOutputMask

  Dim OutputMask    As cOutputMask


  Dim rs            As Recordset
  Set rs = ConnExecute("SELECT * FROM ScreenMasks Where Screen = " & ScreenNumber)
  If Not rs.EOF Then
    Set OutputMask = New cOutputMask
    OutputMask.Announce = Announce

    OutputMask.OG1 = Val("" & rs("og1"))
    OutputMask.OG2 = Val("" & rs("og2"))
    OutputMask.OG3 = Val("" & rs("og3"))
    OutputMask.OG4 = Val("" & rs("og4"))
    OutputMask.OG5 = Val("" & rs("og5"))
    OutputMask.OG6 = Val("" & rs("og6"))


    OutputMask.NG1 = Val("" & rs("ng1"))
    OutputMask.NG2 = Val("" & rs("ng2"))
    OutputMask.NG3 = Val("" & rs("ng3"))
    OutputMask.NG4 = Val("" & rs("ng4"))
    OutputMask.NG5 = Val("" & rs("ng5"))
    OutputMask.NG6 = Val("" & rs("ng6"))


    OutputMask.OG1D = Val("" & rs("og1d"))
    OutputMask.OG2D = Val("" & rs("og2d"))
    OutputMask.OG3D = Val("" & rs("og3d"))
    OutputMask.OG4D = Val("" & rs("og4d"))
    OutputMask.OG5D = Val("" & rs("og5d"))
    OutputMask.OG6D = Val("" & rs("og6d"))


    OutputMask.NG1D = Val("" & rs("ng1d"))
    OutputMask.NG2D = Val("" & rs("ng2d"))
    OutputMask.NG3D = Val("" & rs("ng3d"))
    OutputMask.NG4D = Val("" & rs("ng4d"))
    OutputMask.NG5D = Val("" & rs("ng5d"))
    OutputMask.NG6D = Val("" & rs("ng6d"))


    If InIDE Then
      '      Stop
    End If




    OutputMask.Repeats = Val("" & rs("Repeats"))
    OutputMask.RepeatUntil = Val("" & rs("RepeatUntil"))
    OutputMask.Pause = Val("" & rs("Pause"))
    OutputMask.SendCancel = Val("" & rs("SendCancel"))
    OutputMask.ScreenName = "" & rs("ScreenName")
    OutputMask.RepeatTwice = False
  End If
  rs.Close
  Set rs = Nothing
  Set GetOutputMask = OutputMask

End Function
Public Function GetNCNID() As Integer
  Dim t             As Long
  GlobalNID = 0

  Dim Timeout As Date

  If gDirectedNetwork Then
    dbg "Getting DN NID"
    Outbounds.AddMessage "", MSGTYPE_GETNID, "", 0
  Else
    dbg "Getting BC NID"
    Outbounds.AddMessage "", MSGTYPE_REQTXSTAT, "", 0
  End If
  Timeout = DateAdd("s", OUTBOUND_TIMEOUT, Now)  ' 10 seconds
  't = Win32.timeGetTime + (1000 * OUTBOUND_TIMEOUT)
  Do While Timeout > Now
    If GlobalNID <> 0 Then
      dbg "GOT NID"
      Exit Do
    End If
    DoEvents
  Loop
  dbg "Exit GET NID"
End Function
Public Function GetNCSerial() As String
  Dim t             As Long
  Dim Timeout As Date

  newserial = ""
  If gDirectedNetwork Then
    Outbounds.AddMessage "", MSGTYPE_GETNCSERIAL, "", 0
    't = Win32.timeGetTime + (1000 * OUTBOUND_TIMEOUT)  ' 5 seconds
    
    Timeout = DateAdd("s", OUTBOUND_TIMEOUT, Now)
    
    Do While Timeout > Now ' t > Win32.timeGetTime
      If Len(newserial) > 0 Then
        
        Exit Do
      End If
      DoEvents
    Loop
    dbg "NCSerial " & newserial
    
    If Len(newserial) > 0 Then
      GetNCSerial = newserial
    Else
      GetNCSerial = Configuration.RxSerial
    End If
    'Outbounds.Ready = True
  Else
    GetNCSerial = Configuration.RxSerial
  End If



End Function

Public Function SetNID(ByVal NID As Integer) As Boolean
  Outbounds.AddMessage "", MSGTYPE_SETNID, Right("00" & Hex(NID), 2), 0
End Function

Public Function SyncNIDs() As Boolean
  Dim d             As cESDevice
  Dim j             As Integer
  Dim i             As Integer
  Dim delay         As Long
  Dim start         As Long
  Dim starttime     As Double

  delay = 2




  starttime = CDbl(Now)

  For j = 1 To Devices.Count
    DoEvents


    Set d = Devices.Devices(j)
    'If 1 = 2 Then
    'For i = 1 To UBound(ESDeviceType)
    '  If 0 = StrComp(d.model, ESDeviceType(i).model, vbTextCompare) Then
    '    If ESDeviceType(i).BiDi Then
    '      Outbounds.AddMessage d.serial, MSGTYPE_TWOWAYNID, "", 0
    '    End If
    '    Exit For
    '  End If
    'Next
    'End If

    If 1 = 1 Then
      dbg "checking device model, serial " & d.Model & ", " & d.Serial
      Select Case UCase(d.Model)
      Case "EN5000", "EN5040"
        Outbounds.AddMessage d.Serial, MSGTYPE_REPEATERNID, "", 0
        ' create outbound message to set NID
        dbg "Set Repeater NID " & d.Serial
        starttime = DateAdd("s", delay, Now)
      Case "EN3954", "EN5081"
        Outbounds.AddMessage d.Serial, MSGTYPE_TWOWAYNID, "", 0
        dbg "Set Two-Way NID " & d.Serial
        starttime = DateAdd("s", delay, Now)
        ' create outbound message to set NID
      End Select
    End If

  Next
End Function
Public Function SaveSerialDevice(Device As cESDevice, ByVal Username As String)
10      Dim rs            As ADODB.Recordset: Set rs = New ADODB.Recordset

20      On Error GoTo SaveSerialDevice_Error

30      rs.Open "SELECT * FROM devices WHERE serial = " & q(Device.Serial), conn, gCursorType, gLockType
40      If Not rs.EOF Then
50        rs("SerialTapProtocol") = Device.SerialTapProtocol
60        rs("SerialSkip") = Device.SerialSkip
70        rs("SerialMessageLen") = Device.SerialMessageLen
80        rs("SerialAutoClear") = Device.SerialAutoClear
90        rs("SerialInclude") = left(Device.SerialInclude, 255)
100       rs("SerialExclude") = left(Device.SerialExclude, 255)
110       rs("SerialPort") = Device.SerialPort
120       rs("SerialBaud") = Device.SerialBaud
130       rs("SerialParity") = Device.SerialParity
140       rs("SerialBits") = Device.Serialbits
150       rs("SerialFlow") = Device.SerialFlow
160       rs("SerialStopBits") = Device.SerialStopbits
170       rs("SerialSettings") = Device.SerialSettings
180       rs("SerialEOLChar") = Device.SerialEOLChar
190       rs("SerialPreamble") = Device.SerialPreamble


260       rs.Update
270       SaveSerialDevice = True
280     End If
290     rs.Close
300     Set rs = Nothing


SaveSerialDevice_Resume:
310     On Error GoTo 0
320     Exit Function

SaveSerialDevice_Error:

330     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modES.SaveSerialDevice." & Erl
340     Resume SaveSerialDevice_Resume


End Function



Public Function SaveTemperatureDevice(Device As cESDevice, ByVal Username As String)
10      Dim rs                 As ADODB.Recordset: Set rs = New ADODB.Recordset

20      On Error GoTo SaveTemperatureDevice_Error

30      rs.Open "SELECT * FROM devices WHERE serial = " & q(Device.Serial), conn, gCursorType, gLockType
40      If Not rs.EOF Then
50        rs("lowset") = Device.LowSet
60        rs("lowset_a") = Device.LowSet_A
70        rs("hiset") = Device.HiSet
80        rs("hiset_a") = Device.HiSet_A
90        rs("EnableTemp") = Device.EnableTemperature
100       rs("EnableTemp_a") = Device.EnableTemperature_A
110       rs.Update
120       SaveTemperatureDevice = True
130     End If


SaveTemperatureDevice_Resume:
140     On Error Resume Next
150     rs.Close
160     Set rs = Nothing

170     On Error GoTo 0
180     Exit Function

SaveTemperatureDevice_Error:

190     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modES.SaveTemperatureDevice." & Erl
200     Resume SaveTemperatureDevice_Resume


End Function
Public Function SaveRoom(Room As cRoom, ByVal User As String) As Boolean
10      Dim rs                 As ADODB.Recordset: Set rs = New ADODB.Recordset

20      On Error GoTo SaveRoom_Error

30      If Room.RoomID = 0 Then
40        rs.Open "SELECT * FROM rooms WHERE room = " & q(Room.Room), conn, gCursorType, gLockType
50        If Not rs.EOF Then
60          SaveRoom = False
70          rs.Close
80          Exit Function
90        End If
100     Else
110       rs.Open "SELECT * FROM rooms WHERE roomID = " & Room.RoomID, conn, gCursorType, gLockType
120     End If

130     dbgGeneral "SaveRoom rs.Open "

140     If rs.EOF Then
150       dbgGeneral "SaveRoom addnew "
160       rs.addnew
170     Else
180       dbgGeneral "SaveRoom ID = " & Room.RoomID
190     End If
200     rs("room") = Room.Room
210     rs("Building") = ""
220     rs("Assurdays") = Room.Assurdays
230     rs("Away") = Room.Away
240     rs("Vacation") = 0
        rs("lockw") = Room.locKW
250     rs("Deleted") = 0
255     rs("Flags") = Room.flags

260     dbgGeneral "SaveRoom before update "
270     rs.Update
280     dbgGeneral "SaveRoom updated "
290     rs.MoveLast
300     Room.RoomID = rs("roomid")
310     rs.Close
320     SaveRoom = True
330     dbgGeneral "SaveRoom OK "
SaveRoom_Resume:

340     Set rs = Nothing
350     RefreshJet
360     On Error GoTo 0
370     Exit Function

SaveRoom_Error:

380     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modES.SaveRoom." & Erl
390     dbgGeneral "SaveRoom Error " & Err.Number & " (" & Err.Description & ") at modES.SaveRoom." & Erl
400     Resume SaveRoom_Resume


End Function
Public Sub dbgTAPI(ByVal s As String)
  On Error Resume Next
  If gShowTAPIData Then
    'Debug.Print s
    Win32.OutputDebugString s & vbCrLf
  End If

End Sub

Public Sub dbgPackets(ByVal s As String)
  On Error Resume Next
  If gShowPacketData Then
    Win32.OutputDebugString s & vbCrLf
  End If

End Sub

Public Sub dbgloc(ByVal s As String)

  On Error Resume Next
  If gShowLocationData Then
    Win32.OutputDebugString s & vbCrLf
  End If
End Sub

Public Sub dbgHostRemote(ByVal s As String)

  On Error Resume Next
  If gShowHostRemoteData = 0 Then
    Win32.OutputDebugString s & vbCrLf
  End If
End Sub
Public Sub dbgGeneral(ByVal s As String)
  On Error Resume Next
  If gShowGeneralData Then
    Win32.OutputDebugString s & vbCrLf
  End If
End Sub


Public Sub dbg(ByVal s As String)
  On Error Resume Next
  Win32.OutputDebugString s & vbCrLf
End Sub
Public Function SetResidentAwayStatus(ByVal Away As Integer, ByVal ResID As Long, ByVal User As String) As Boolean
  Dim rs            As ADODB.Recordset
  Dim SQL           As String
10 dbgGeneral "modES.SetResidentAwayStatus Start RESID, away " & ResID & "," & Away

20 ConnExecute "UPDATE Residents SET Away = " & Away & " WHERE ResidentID = " & ResID

30 SetDevicesAwayByResident ResID, Away
40 'frmMain.ClearAwayAlarms ResID

50 Set rs = ConnExecute("SELECT Serial, RoomID FROM Devices WHERE ResidentID <> 0 and ResidentID = " & ResID)
60 Do Until rs.EOF
70  If FieldToNumber(rs("roomid")) <> 0 Then
80    SQL = "UPDATE Rooms SET Away = " & Away & " WHERE RoomID = " & rs("roomid")
90    ConnExecute SQL
100 End If
110 rs.MoveNext
120 Loop
130 rs.Close
140 Set rs = Nothing

150 LogVacation ResID, 0, Away, User

160 SetResidentAwayStatus = Away
170 dbgGeneral "modES.SetResidentAwayStatus Done RESID, away " & ResID & "," & Away
End Function
Public Function SetRoomAwayStatus(ByVal Away As Integer, ByVal RoomID As Long, ByVal User As String) As Boolean
  Dim rs            As ADODB.Recordset

  ConnExecute "UPDATE Rooms SET Away = " & Away & " WHERE RoomID = " & RoomID
  SetDevicesAwayByRoom RoomID, Away
  'frmMain.ClearAwayAlarmsByRoomID RoomID

  LogVacation 0, RoomID, Away, User
  SetRoomAwayStatus = Away

End Function
Public Sub UpdateDeviceCheckin(ByVal Model As String, ByVal Checkin As Long)
  Dim j             As Integer
  Dim d             As cESDevice
  
  

  
  For j = 1 To Devices.Devices.Count
    Set d = Devices.Devices(j)
    If 0 = StrComp(d.Model, Model, vbTextCompare) Then
      d.SupervisePeriod = Checkin
    End If
  Next

End Sub
Public Function GetNoCheckIn(ByVal Model As String) As Integer
  GetNoCheckIn = 0
  Dim j             As Integer

  ' If InStr(1, Model, "duk", vbTextCompare) Then Stop

  If InStr(1, Model, "ES1242", vbTextCompare) Then
    Model = "ES1242"
  End If


  For j = 1 To MAX_ESDEVICETYPES
    If 0 = StrComp(ESDeviceType(j).Model, Model, vbTextCompare) Then
      GetNoCheckIn = ESDeviceType(j).NoCheckin
      Exit For
    End If
  Next

End Function

Public Function GetSupervisePeriod(ByVal Model As String) As Long
  GetSupervisePeriod = gSupervisePeriod  ' default if not found
  Dim j             As Integer
  
  If InStr(1, Model, "ES1242", vbTextCompare) Then
    Model = "ES1242"
  End If


  For j = 1 To MAX_ESDEVICETYPES
    If 0 = StrComp(ESDeviceType(j).Model, Model, vbTextCompare) Then
      GetSupervisePeriod = ESDeviceType(j).Checkin
      Exit For
    End If
  Next


End Function
Public Function GetAutoClear(ByVal Model As String) As Long
  GetAutoClear = 0
  Dim j             As Integer

  If InStr(1, Model, "ES1242", vbTextCompare) Then
    Model = "ES1242"
  End If

  For j = 1 To MAX_ESDEVICETYPES
    If 0 = StrComp(ESDeviceType(j).Model, Model, vbTextCompare) Then
      GetAutoClear = ESDeviceType(j).AutoClear
      Exit For
    End If
  Next


End Function

Public Property Get GlobalNID() As Integer

  GlobalNID = gCurrentNID

End Property

Public Property Let GlobalNID(ByVal Value As Integer)
'dbg "Global NID set to " & Value
' write to INI File

  WriteSetting "Configuration", "SystemNID", Value
  gCurrentNID = Value

End Property

Public Property Get NIDMatch() As Boolean

  NIDMatch = mNIDMatch

End Property

Public Property Let NIDMatch(ByVal NIDMatch As Boolean)

  mNIDMatch = NIDMatch

End Property

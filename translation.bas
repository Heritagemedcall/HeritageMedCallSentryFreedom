Attribute VB_Name = "modTranslation"
Option Explicit

Global TranslationTable As Collection

Function NewTranslationItem(ByVal i6080code As Long, ByVal i6080Desc As String, ByVal HMCcode As Long, ByVal HMCDesc As String) As cTranlationItem
  Dim TI  As cTranlationItem
  Set TI = New cTranlationItem
  TI.i6080code = i6080code
  TI.i6080Desc = i6080Desc
  TI.HMCcode = HMCcode
  TI.HMCDesc = HMCDesc
  Set NewTranslationItem = TI
  Set NewTranslationItem = Nothing
End Function

Sub FillTraslationTable()
  Set TranslationTable = New Collection

  Dim TI  As cTranlationItem
  
  Set TI = NewTranslationItem(1, "Alarm1", 1, "Button1")
  
  Set TI = NewTranslationItem(2, "Alarm1 Clear", 2, "Button Clear")
  Set TI = NewTranslationItem(3, "Alarm2", 2, "Button Clear")



'1 Alarm1
'2 Alarm1 has cleared
'3 Alarm2
'4 Alarm2 has cleared
'5 Alarm3
'6 Alarm3 has cleared
'7 Alarm3
'8 Alarm3 has cleared
'9 Device is Inactive
'10 Device is Inactive has cleared (is now active)
'11 tamper activated
'12 tamper cleared
'13 EOL tamper activated
'14 EOL tamper cleared
'15 Low battery
'16 Low battery cleared
'17 Maintenance required
'18 Maintenance required cleared
'21 Device Reset
'25 Endpoint Configuration Fail
'26 Endpoint Configuration Success
'27 Repeater Power Loss
'28 Repeater Power Loss clear
'29 Repeater Reset
'31 Repeater Tamper
'32 Repeater tamper clear
'33 Repeater low battery
'34 Repeater low battery clear
'35 Repeater Jam
'36 Repeater jam clear
'43 Repeater is inactive
'44 Repeater is inactive clear (is now active)
'45 Repeater Configuration fail
'46 Repeater Configuration Success
'49 ACG Reset
'51 ACG Tamper
'52 ACG Tamper clear
'55 ACG Jammed
'56 ACG Jam clear
'57 ACG Inactive
'58 ACG Inactive clear
'59 ACG Configuration Fail
'60 ACG Configuration success
'61 ACG CRC Check Fail
'71 ACG F/W Update Success
'72 ACG F/W Update Failed
'91 ACG Battery Failed
'92 ACG Battery low
'93 ACG Battery OK
'96 ACG Shutdown Imminent
'97 ACG F/W Update Pending
'99 ACG IP Processor CRC Invalid
'100 ACG Reboot Requested
'125 ACG Checkin

End Sub




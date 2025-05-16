VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cESSurveyDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mLocations    As Collection  ' FA
Private mFirstHops    As Collection  ' ES
Private mSerial       As String

Private mPCASerial    As String

Public DateTime       As Date
Public ReadyToPost    As Boolean
Public MessageString  As String
Public Sequence       As Long
Public RequireACK     As Boolean

Public Location1      As String
Public Location2      As String
Public Location3      As String
Public Signal1        As Double
Public Signal2        As Double
Public Signal3        As Double

Public SendAckMSG     As Integer
Private mPagerID      As Long

Private Const Holdoff As Long = 10
Private LastResponse  As Date

Public Function ResponseCode(packet As cESPacket, Optional ByVal surveymode As Integer = 0) As Integer
  Dim Elapsed
' 0 is not one of the responses!
' 1 is OK
' 2 is Cancel
  If surveymode = PCA_MODE Then
    If packet.Sequence = Sequence Then
      ResponseCode = packet.ResponseCode
    End If
  ElseIf surveymode = EN1221_MODE Then
    Elapsed = DateDiff("s", LastResponse, Now)
    If Elapsed < LocatorWaitTime Then  'Holdoff Then
      ResponseCode = SURVEY_RC0
    ElseIf packet.Alarm0 <> 0 Then
      ResponseCode = SURVEY_RC0
      LastResponse = Now
    ElseIf packet.Alarm0 = 0 Then ' simulate button 2 on multi pendant
      If packet.Reset Then
        ResponseCode = SURVEY_RC1
      End If
    Else
      ResponseCode = SURVEY_RC0
    End If
  
  ElseIf surveymode = TWO_BUTTON_MODE Then
    Elapsed = DateDiff("s", LastResponse, Now)
    If Elapsed < LocatorWaitTime Then
      ResponseCode = SURVEY_RC0
    ElseIf packet.Alarm0 <> 0 Then
      ResponseCode = SURVEY_RC0
      LastResponse = Now
    ElseIf packet.Alarm1 <> 0 Then ' button 2 pushed
      ResponseCode = SURVEY_RC1
      LastResponse = Now
    Else
      ResponseCode = SURVEY_RC0
    End If
  End If

End Function
Public Function RequestSurvey(ByVal prompt As String, ByVal ResponseCount As Integer, Optional ByVal SendToPager As Boolean = False) As String
  
  
  If SendToPager Then
    RequestSurvey = PagerRequestSurvey(prompt, ResponseCount)
  Else
    RequestSurvey = PCARequestSurvey(prompt, ResponseCount)
  End If
  
  
End Function

Public Function PagerRequestSurvey(ByVal prompt As String, ByVal ResponseCount As Integer) As String
  Dim t As Long
  Dim Responses As Collection
  Set Responses = New Collection
  Responses.Add "OK"
  MessageString = prompt
  RequireACK = False
  PagerRequestSurvey = MessageString
  DateTime = Now
  LastResponse = Now

End Function

Public Function PCARequestSurvey(ByVal prompt As String, ByVal ResponseCount As Integer) As String
  'dbgloc "cESSurveyDevice.PCARequestSurvey prompt: " & prompt
  
  
  Dim t As Long
  
  Dim Responses As Collection
  
  
  Set Responses = New Collection
  Responses.Add "OK"
  If ResponseCount > 1 Then
    Responses.Add "SKIP"
  End If
  If ResponseCount > 2 Then
    Responses.Add "EXIT"
  End If

  MessageString = CreateCustomCustom(mPCASerial, prompt, ResponseCount, Responses)
  Sequence = Val("&h" & MID(MessageString, 21, 4) & "&")  ' bytes 11 And 12 of outbound
  RequireACK = True
  PCARequestSurvey = MessageString
  DateTime = Now
End Function

Public Function ProcessLocations() As Boolean
  ' sort on average FirstHopUID
  'Dim locations() As cESPacket
  Dim j                  As Integer
  Dim i                  As Integer
  Dim Count              As Integer

  Dim Serial             As String
  Dim LEvel              As String

  Dim Repeaters          As Collection
  Dim Repeater           As cRepeater
  Dim p                  As cESPacket

  Count = mFirstHops.Count
  If DateTime = 0 Then Exit Function

  If Count > 0 Then
    dbgloc "Processing Locations " & vbCrLf
    i = DateDiff("s", DateTime, Now)
    If i > LocatorWaitTime Then
      ' ******************* the new way

      Set Repeaters = New Collection

      For j = 1 To Count  ' accumulate distinct repeaters
        Set p = mFirstHops(j)
        If p.IsLocatorPacket Then
          Serial = p.Serial
          LEvel = p.LocatedLevel
        Else
          Serial = p.FirstHopSerial
          LEvel = p.LEvel
        End If

        For i = 1 To Repeaters.Count
          Set Repeater = Repeaters(i)
          dbgloc "Packet Ser Lvl " & Serial & " " & LEvel & vbCrLf
          If gDirectedNetwork Then  ' DNet
            dbgloc " DN = true"
            If Configuration.OnlyLocators Then  ' use only locators
              dbgloc " OnlyLocators = true"
              If p.IsLocatorPacket Then
                If Repeater.Serial = Serial Then
                  Exit For
                End If
              End If
            Else  ' use both locators and repeaters
              dbgloc " OnlyLocators = false"
              If p.IsLocatorPacket Then  ' handle locators
                dbgloc " IsLocatorPacket = true"
                If Repeater.Serial = Serial Then
                  Exit For
                End If
              Else  ' handle repeaters
                dbgloc " IsLocatorPacket = false"
                If Repeater.Serial = Serial Then
                  Exit For
                End If
              End If
            End If

            ' NOT DNet so just use the old way
          ElseIf Repeater.Serial = Serial Then
            Exit For
          End If
        Next

        dbgloc "Serial Level " & Serial & " " & LEvel & vbCrLf

        If i > Repeaters.Count Then
          dbgloc "Adding Repeater " & Serial & " " & LEvel & vbCrLf
          Set Repeater = New cRepeater
          Repeaters.Add Repeater
          Repeater.Serial = Serial
          Repeater.LEvel = LEvel
        Else
          Repeater.LEvel = Repeater.LEvel + LEvel
        End If
        Repeater.HitCount = Repeater.HitCount + 1

      Next

      For j = 1 To Repeaters.Count  ' average signals for each repeater
        Set Repeater = Repeaters(j)
        Repeater.LEvel = Repeater.LEvel / Repeater.HitCount
      Next


      SortRepeaters Repeaters  ' get in decending order
      dbgloc "Process Locations 'Repeaters '"
      For j = 1 To Repeaters.Count  ' test printout
        dbgloc Repeaters(j).Serial & Space(8) & Format(Repeaters(j).LEvel, "00.00")
      Next

      ' not sure if/why we need to clear this here
      Set mFirstHops = New Collection

      If Repeaters.Count >= 1 Then
        'If Not (loc1(1) Is Nothing) Then
        'Location1 = Right("00" & Hex(loc1(1).FirstHopMID), 2) & Right("00000000" & Hex(loc1(1).FirstHopUID), 8)
        Location1 = Repeaters(1).Serial  ' Right("00000000" & Hex(loc1(1).FirstHopMID) & Hex(loc1(1).FirstHopUID), 8)
        Signal1 = Format(Repeaters(1).LEvel, "0")
      End If
      'If Not (loc2(1) Is Nothing) Then
      If Repeaters.Count >= 2 Then
        Location2 = Repeaters(2).Serial  ' Right("00000000" & Hex(loc2(1).FirstHopMID) & Hex(loc2(1).FirstHopUID), 8)
        Signal2 = Format(Repeaters(2).LEvel, "0")
      End If
      If Repeaters.Count >= 3 Then
        Location3 = Repeaters(3).Serial  ' Right("00000000" & Hex(loc2(1).FirstHopMID) & Hex(loc2(1).FirstHopUID), 8)
        Signal3 = Format(Repeaters(3).LEvel, "0")
      End If

      dbgloc "RESULTS   Loc1 " & Location1 & " S1 " & Signal1 & " Loc2 " & Location2 & " S2 " & Signal2 & " Loc3 " & Location3 & " S3 " & Signal3 & vbCrLf
      Debug.Print "RESULTS   Loc1 " & Location1 & " S1 " & Signal1 & " Loc2 " & Location2 & " S2 " & Signal2 & " Loc3 " & Location3 & " S3 " & Signal3 & vbCrLf
      ProcessLocations = True
      ReadyToPost = True
    End If
  Else

  End If
End Function



Public Function AddLocater(p As cESPacket) As Integer
      ' returns 0 if locator not added.
      ' else returns 1 upon first locator

        Dim i As Long

10      On Error GoTo AddLocater_Error


20      If mFirstHops.Count = 0 Then  ' start window on first packet
30        DateTime = Now
40      End If

50      i = DateDiff("s", DateTime, p.DateTime)
        'dbgloc "Time out in " & i - (LocatorWaitTime) & " Sec"
60      If i <= (LocatorWaitTime + 3) Then  ' global setting
          ' need to modify
          ' dbgloc "CESurveydevice.AddLocater Device Located: " & IIf(p.IsLocatorPacket, p.LocatedSerial, p.Serial)
70        mFirstHops.Add p
80      Else
90        dbgloc "************ Locator Window Expired ***********"
100     End If ' i <= LocatorWaitTime



AddLocater_Resume:
110     On Error GoTo 0
120     Exit Function

AddLocater_Error:

130     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at cESSurveyDevice.AddLocater." & Erl
140     Resume AddLocater_Resume

End Function

Private Sub Class_Initialize()
  Set mLocations = New Collection
  Set mFirstHops = New Collection
End Sub

Public Property Get Serial() As String
  Serial = mSerial
End Property

Public Property Let Serial(ByVal Serial As String)
  mSerial = Serial
End Property

Public Property Get PCASerial() As String
  PCASerial = mPCASerial
End Property

Public Property Let PCASerial(ByVal PCASerial As String)
  mPCASerial = PCASerial
End Property

Public Property Get PagerID() As Long

  PagerID = mPagerID

End Property

Public Property Let PagerID(ByVal PagerID As Long)

  mPagerID = PagerID

End Property

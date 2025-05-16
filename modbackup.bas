Attribute VB_Name = "modBackup"
Option Explicit
'Private mLastBackup   As Date
Private mNextBackup           As Date
Private mNextBackupRemote     As Date

Sub ClearBackupDate()
  mNextBackup = 0
  mNextBackupRemote = 0
  CheckIfBackupDue
End Sub




Function CheckIfBackupDue() As Boolean
  If MASTER Then
    If Configuration.BackupEnabled Then
      If mNextBackup = 0 Then
        mNextBackup = GetNextBackupDate()
      Else
        If Now > mNextBackup Then
          DoBackup
          mNextBackup = GetNextBackupDate()
        Else
          'Debug.Print "Next Backup Due " & mNextBackup
        End If
      End If
    End If

    If Configuration.BackupEnabledRemote Then
      If mNextBackupRemote = 0 Then
        mNextBackupRemote = GetNextBackupDateRemote()
      Else
        If Now > mNextBackupRemote Then
          DoBackupRemote
          mNextBackupRemote = GetNextBackupDateRemote()
        Else
          'Debug.Print "Next Backup Remote Due " & mNextBackupRemote
        End If
      End If
    End If

  End If

End Function
Function GetNextBackupDate() As Date
  Dim j As Integer
  Dim MonthDayNumber        As Integer
  Dim WeekDayNumber         As Integer
  Dim Makedate              As Date
  Dim WeekDays(0 To 7)      As Integer
  Dim BackupDays(0 To 7)    As Integer
  Dim DailyBackups(0 To 7)  As Date
  
  Dim CurrentDate           As Date

  Dim MonthlyBackups() As String

  MonthDayNumber = Day(Now)
  WeekDayNumber = Weekday(Now, vbSunday) - 1

  Select Case Configuration.BackupType
    Case 1  ' monthly
         
     
      If Configuration.BackupDOM <= 0 Or Configuration.BackupDOM > 28 Then
        Configuration.BackupDOM = 1
      End If
      Makedate = DateSerial(Year(Now), Month(Now), Configuration.BackupDOM)
      Makedate = DateAdd("h", Configuration.BackupTime / 100, Makedate)
      If Makedate < Now Then
        Makedate = DateAdd("m", 1, Makedate)
      End If
      GetNextBackupDate = Makedate
      
      
    Case Else  ' days -  weekly
      Makedate = DateSerial(Year(Now), Month(Now), Day(Now)) ' create
      Makedate = DateAdd("h", Configuration.BackupTime / 100, Makedate)
      
      For j = 0 To 6
        If ((2 ^ j) And Configuration.BackupDOW) <> 0 Then
          BackupDays(j) = 1
          WeekDays(j) = j - WeekDayNumber
          If WeekDays(j) < 0 Then
            WeekDays(j) = WeekDays(j) + 7
          End If
          DailyBackups(j) = DateAdd("d", WeekDays(j), Makedate)
        End If
      Next
      
      If BackupDays(1) = 1 Then
        BackupDays(7) = 1
        DailyBackups(7) = DateAdd("d", 7, Makedate)
      End If
      
      
      SortDates DailyBackups()
      For j = 0 To 7

        'If BackupDays(j) = 1 Then
          If DailyBackups(j) > Now Then
            GetNextBackupDate = DailyBackups(j)

            Exit For
          End If
        'End If
      Next
'      For j = 0 To 7
'
'        If BackupDays(j) = 1 Then
'          If DailyBackups(j) > Now Then
'            GetNextBackupDate = DailyBackups(j)
'
'            Exit For
'          End If
'        End If
'      Next
        


  End Select

End Function

Function GetNextBackupDateRemote() As Date
  Dim j As Integer
  Dim MonthDayNumber        As Integer
  Dim WeekDayNumber         As Integer
  Dim Makedate              As Date
  Dim WeekDays(0 To 7)      As Integer
  Dim BackupDays(0 To 7)    As Integer
  Dim DailyBackups(0 To 7)  As Date

  Dim CurrentDate           As Date

  MonthDayNumber = Day(Now)
  WeekDayNumber = Weekday(Now, firstdayofweek:=vbSunday) - 1

  Select Case Configuration.BackupTypeRemote
    Case 1  ' monthly
      If Configuration.BackupDOM <= 0 Or Configuration.BackupDOMRemote > 28 Then
        Configuration.BackupDOMRemote = 1
      End If
      Makedate = DateSerial(Year(Now), Month(Now), Configuration.BackupDOMRemote)
      Makedate = DateAdd("h", Configuration.BackupTimeRemote / 100, Makedate)
      If Makedate < Now Then
        Makedate = DateAdd("m", 1, Makedate)
      End If
      GetNextBackupDateRemote = Makedate
    Case Else  ' days -  weekly
      Makedate = DateSerial(Year(Now), Month(Now), Day(Now)) ' create
      Makedate = DateAdd("h", Configuration.BackupTimeRemote / 100, Makedate)

      For j = 0 To 6
        If ((2 ^ j) And Configuration.BackupDOWRemote) <> 0 Then
          BackupDays(j) = 1
          WeekDays(j) = j - WeekDayNumber
          If WeekDays(j) < 0 Then
            WeekDays(j) = WeekDays(j) + 7
          End If
          DailyBackups(j) = DateAdd("d", WeekDays(j), Makedate)
        End If
      Next

      If BackupDays(1) = 1 Then
        BackupDays(7) = 1
        DailyBackups(7) = DateAdd("d", 7, Makedate)
      End If

' need to sort dates


      
      SortDates DailyBackups()
      For j = 0 To 7

        'If BackupDays(j) = 1 Then
          If DailyBackups(j) > Now Then
            GetNextBackupDateRemote = DailyBackups(j)

            Exit For
          End If
        'End If
      Next



  End Select

End Function
Function SortDates(DayArray() As Date) As Boolean
'BubbleSort IS GOOD ENOUGH

    Dim i As Long
    Dim Min As Long
    Dim Max As Long
    Dim Swap As Date
    Dim Swapped As Boolean
    
    Min = LBound(DayArray)
    Max = UBound(DayArray) - 1
    Do
        Swapped = False
        For i = Min To Max
            If DayArray(i) > DayArray(i + 1) Then
                Swap = DayArray(i)
                DayArray(i) = DayArray(i + 1)
                DayArray(i + 1) = Swap
                Swapped = True
            End If
        Next
        Max = Max - 1
    Loop Until Not Swapped




End Function

Function DoBackup() As Boolean
        Dim BackupRoot    As String
        Dim BackupSubDir  As String
        Dim AppPath       As String

        Dim IniFileName   As String
        Dim DBFileName    As String
        Dim UDLFilename   As String
        Dim DBShortName   As String
        Dim EXEFilename   As String


        Dim DBName        As String
        Dim SQl           As String
        Dim RA            As Long
        Dim rs            As ADODB.Recordset


        Dim counter       As Long
        Static Busy       As Boolean


10      On Error GoTo DoBackup_Error

20      If Busy Then Exit Function
30      Busy = True

40      AppPath = App.path
50      If Right(AppPath, 1) <> "\" Then
60        AppPath = AppPath & "\"
70      End If

80      EXEFilename = AppPath & App.exename & ".exe"
90      IniFileName = AppPath & "FREEDOM2.INI"
100     UDLFilename = AppPath & "FREEDOM2.UDL"

110     If gIsJET Then
120       DBFileName = conn.Properties("Data Source") & ""
130     Else
140       DBName = conn.Properties("Current Catalog")
150       DBFileName = DBName & ".bak"
160     End If

170     WriteSetting "Backup", "Date", Now
180     WriteSetting "Backup", "SourceDir", App.path
190     WriteSetting "Backup", "EXEName", EXEFilename
200     WriteSetting "Backup", "SourceMDB", IIf(gIsJET, DBFileName, "SQL " & DBName)


        Dim i             As Long
210     i = InStrRev(DBFileName, "\", -1, vbTextCompare)
220     If i > 0 And i < Len(DBFileName) Then
230       DBShortName = MID$(DBFileName, i + 1)
240     End If



250     BackupSubDir = "Backup_" & Format(Now, "yyyymmdd")
260     BackupRoot = Configuration.BackupFolder
270     If BackupRoot = "" Then
280       BackupRoot = AppPath & "Backups"
290     End If

300     If Right(BackupRoot, 1) <> "\" Then
310       BackupRoot = BackupRoot & "\"
320     End If

330     If DirExists(BackupRoot) Then
340       BackupSubDir = BackupRoot & BackupSubDir
350       If Not DirExists(BackupSubDir) Then
360         MkDir BackupSubDir
370       End If
380     End If
390     If DirExists(BackupSubDir) Then
400       WriteSetting "BACKUP", "Dest", BackupSubDir
          ' ok write files
410       If Right(BackupSubDir, 1) <> "\" Then
420         BackupSubDir = BackupSubDir & "\"
430       End If
440       If gIsJET Then
450         CopyFile DBFileName, BackupSubDir & DBShortName
460       Else
470         SQl = "BACKUP DATABASE " & DBName & " TO DISK = '" & BackupSubDir & DBFileName & "' WITH FORMAT, MEDIANAME = '" & DBName & "Backup', NAME = 'Full Backup of " & DBName & "';"
480         Set rs = conn.Execute(SQl, RA, adAsyncExecute)
            
490         Do While rs.State = adStateExecuting

500           'Debug.Print "Still Executing Local Backup "; Now
              counter = counter + 1
              If counter > 500 Then
                counter = 0
                DoEvents
              End If
520         Loop
530         Set rs = Nothing
540       End If
550       CopyFile IniFileName, BackupSubDir & "FREEDOM2.INI"
560       CopyFile UDLFilename, BackupSubDir & "FREEDOM2.UDL"
570       CopyFile EXEFilename, BackupSubDir & App.exename & ".exe"
580     Else
590       WriteSetting "Backup", "Dest Failed", BackupSubDir
600     End If



DoBackup_Resume:
        
610     Busy = False
620     On Error GoTo 0
630     Exit Function

DoBackup_Error:

640     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modBackup.DoBackup." & Erl
650     Resume DoBackup_Resume


End Function

Function CheckIfRemoteBackupRunning() As Boolean
  Dim RBU As Object
  Dim rc As Boolean
  On Error Resume Next
  rc = IsBackupRunning()
  If (False = rc) Then
    Shell App.path & "\RemoteBackup.exe", vbHide
  End If


End Function
Function IsBackupRunning() As Boolean
 
  Dim hwnd As Long
  ' If uploader is running this API call returns its handle.
  hwnd = Win32.FindWindow(vbNullString, "Remote Backup Configuration")
  IsBackupRunning = (hwnd <> 0)
End Function

Function DoBackupRemote() As Boolean
        Dim BackupRoot    As String
        Dim BackupSubDir  As String
        Dim AppPath       As String

        Dim IniFileName   As String
        Dim DBFileName    As String
        Dim UDLFilename   As String
        Dim DBShortName   As String
        Dim EXEFilename   As String


        Dim DBName        As String
        Dim SQl           As String
        Dim RA            As Long
        Dim rs            As ADODB.Recordset



        'USE Freedom2;
        'GO
        'BACKUP DATABASE Freedom2
        'TO DISK = 'C:\backupfolder\Freedom2.Bak'
        '   WITH FORMAT,
        '      MEDIANAME = 'Fredom2_Backup',
        '      NAME = 'Full Backup of Freedom2';
        'GO


        Static Busy       As Boolean

10      CheckIfRemoteBackupRunning

20      On Error GoTo DoBackupRemote_Error

30      If Busy Then Exit Function
40      Busy = True

50      AppPath = App.path
60      If Right(AppPath, 1) <> "\" Then
70        AppPath = AppPath & "\"
80      End If



        'DBFileName = conn.Properties("Data Source") & ""
90      If gIsJET Then
100       DBFileName = conn.Properties("Data Source") & ""
110     Else
120       DBName = conn.Properties("Current Catalog")
130       DBShortName = DBName & ".bak"
140     End If



150     IniFileName = AppPath & "FREEDOM2.INI"
        Dim i             As Long
160     i = InStrRev(DBFileName, "\", -1, vbTextCompare)
170     If i > 0 And i < Len(DBFileName) Then
180       DBShortName = MID$(DBFileName, i + 1)
190     End If

200     BackupRoot = Configuration.BackupFolderRemote
210     If BackupRoot = "" Then
220       BackupRoot = AppPath & "Backups"
230     End If

240     If Right(BackupRoot, 1) <> "\" Then
250       BackupRoot = BackupRoot & "\"
260     End If

270     If 0 <> StrComp(DBFileName, BackupRoot & DBShortName, vbTextCompare) Then

280       If DirExists(BackupRoot) Then
            ' ok write files
290         If (FileExists(BackupRoot & DBShortName)) Then
300           Kill (BackupRoot & DBShortName)
310         End If
320         If (FileExists(BackupRoot & "Freedom2.ini")) Then
330           Kill (BackupRoot & "Freedom2.ini")
340         End If

350         If gIsJET Then
360           CopyFile DBFileName, BackupRoot & DBShortName
370         Else
380           SQl = "BACKUP DATABASE " & DBName & " TO DISK = '" & BackupRoot & DBShortName & "' WITH FORMAT, MEDIANAME = '" & DBName & "Backup', NAME = 'Full Backup of " & DBName & "';"
390           Set rs = conn.Execute(SQl, RA, adAsyncExecute)
400           Do While rs.State = adStateExecuting
410             Debug.Print "Still Executing Local Backup"
420             DoEvents
430           Loop
440           Set rs = Nothing
450         End If


460         CopyFile IniFileName, BackupRoot & "freedom2.ini"
470         WriteSetting "RemoteBackup", "Folder", BackupRoot
480         WriteSetting "RemoteBackup", "Filename", DBShortName
490         WriteSetting "RemoteBackup", "FullPath", BackupRoot & DBShortName
500         WriteSetting "RemoteBackup", "Date Remote", Now
510         WriteSetting "RemoteBackup", "Ready", 1

520       Else
530         If gIsJET Then
540           WriteSetting "RemoteBackup", "Dest Failed Remote", BackupRoot & DBShortName
550         Else
560           WriteSetting "RemoteBackup", "Dest Failed Remote", BackupRoot & DBFileName
570         End If
580       End If
590     Else
600       LogProgramError "Error: Source and Destination Files are the Same: modBackup.DoBackupRemote."
610     End If
        


DoBackupRemote_Resume:
620     Busy = False
630     On Error GoTo 0
640     Exit Function

DoBackupRemote_Error:

650     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modBackup.DoBackupRemote." & Erl
660     Resume DoBackupRemote_Resume


End Function

Function CopyFile(ByVal Source As String, ByVal Dest As String) As Boolean
        Dim hSrc          As Integer
        Dim hDest         As Integer
        Dim Buffer        As String
        Dim t             As Long
        Dim FileLen       As Long
        Dim Remaining     As Long

        Dim ChunkSize     As Long

        Dim chunk()       As Byte


        Dim wdtimer       As Date

        Dim lastnow       As Date

        Dim MaxChunkSize  As Long
10      On Error GoTo CopyFile_Error

20      MaxChunkSize = 2 ^ 22  ' 2 ^ 28 = 268,435,456 bytes

30      t = Win32.timeGetTime
40      hSrc = FreeFile

45      gSuspendPackets = True

50      Open Source For Binary Access Read As #hSrc

60      hDest = FreeFile
70      Open Dest For Binary Access Write As #hDest

80      FileLen = LOF(hSrc)




90      Remaining = FileLen
100     If FileLen > MaxChunkSize Then
110       ChunkSize = MaxChunkSize
120     Else
130       ChunkSize = FileLen
140     End If


        '       Dim numchunks As Long
        '
        '      numchunks = Filelen \ ChunkSize

        '      ReDim chunks(numchunks)

150     'numchunks = 0

160     wdtimer = Now

        lastnow = DateAdd("s", 1, Now)
170     Do While Remaining > 0
180       'numchunks = numchunks + 1

          Buffer = Space(ChunkSize)
190       'ReDim chunk(1 To ChunkSize)

200       Get #hSrc, , Buffer
210       'Get #hSrc, , chunk


220       Put #hDest, , Buffer
230       'Put #hDest, , chunk
240       Remaining = Remaining - ChunkSize
250       ChunkSize = Min(ChunkSize, Remaining)
          If Now > lastnow Then
255       DoEvents
          lastnow = DateAdd("s", 1, Now)
          End If
260       If DateDiff("s", wdtimer, Now) > 10 Then
            
            Debug.Print "Checked Watchdog " & Now
270         CheckWatchdog
280         wdtimer = Now
290       End If

300     Loop

310     Close #hSrc
320     Close #hDest
330     t = Win32.timeGetTime - t
340     Debug.Print "File size: " & FileLen
350     Debug.Print "Time to read and write file: " & t & "ms"

CopyFile_Resume:

355     gSuspendPackets = False
360     On Error GoTo 0
370     Exit Function

CopyFile_Error:

380     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modBackup.CopyFile." & Erl
390     Resume CopyFile_Resume

End Function




Function DirExists(ByVal DirName As String) As Boolean
  Dim s As String
  On Error Resume Next
  s = Dir(DirName, vbDirectory)
  If Err.Number = 0 Then
    DirExists = Len(s) > 0
  End If

End Function

Function EnsurePathExists(ByVal path As String) As Boolean
  On Error Resume Next
  If Not (DirExists(path)) Then
    MkDir path
  End If
  EnsurePathExists = DirExists(path)
End Function

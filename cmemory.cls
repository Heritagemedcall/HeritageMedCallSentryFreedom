VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUSEX)

    


Private Type INT64
  LoPart                 As Long
  HiPart                 As Long
End Type

Private Type MEMORYSTATUSEX
  dwLength               As Long
  dwMemoryLoad           As Long
  ulTotalPhys            As INT64
  ulAvailPhys            As INT64
  ulTotalPageFile        As INT64
  ulAvailPageFile        As INT64
  ulTotalVirtual         As INT64
  ulAvailVirtual         As INT64
  ulAvailExtendedVirtual As INT64
End Type

Private Const MAX_PATH = 260
 
Private Type PROCESS_MEMORY_COUNTERS
    cb                          As Long
    PageFaultCount              As Long
    PeakWorkingSetSize          As Long
    WorkingSetSize              As Long
    QuotaPeakPagedPoolUsage     As Long
    QuotaPagedPoolUsage         As Long
    QuotaPeakNonPagedPoolUsage  As Long
    QuotaNonPagedPoolUsage      As Long
    PagefileUsage               As Long
    PeakPagefileUsage           As Long
End Type
Private Type PROCESSENTRY32
    dwSize                      As Long
    cntUsage                    As Long
    th32ProcessID               As Long
    th32DefaultHeapID           As Long
    th32ModuleID                As Long
    cntThreads                  As Long
    th32ParentProcessID         As Long
    pcPriClassBase              As Long
    dwFlags                     As Long
    szExeFile                   As String * MAX_PATH
End Type
Private Type OSVERSIONINFO
    dwOSVersionInfoSize         As Long
    dwMajorVersion              As Long
    dwMinorVersion              As Long
    dwBuildNumber               As Long
    dwPlatformId                As Long
    szCSDVersion                As String * 128
End Type
 
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal lHandle As Long, ByRef lpStructure As PROCESS_MEMORY_COUNTERS, ByVal lSize As Long) As Long
 
Private Const TH32CS_SNAPPROCESS                  As Long = 2
Private Const PROCESS_QUERY_LIMITED_INFORMATION   As Long = &H1000
Private Const PROCESS_QUERY_INFORMATION           As Long = &H400
Private Const INVALID_HANDLE_VALUE                As Long = -1
 
Dim IsVistaAndLater As Boolean

Private mTotalPhysicalMemory As String
Private mAvailablePhysicalMemory As String
Private mTotalVirtualMemory As String
Private mAvailableVirtualMemory As String
Private mMemoryLoad      As String



Public Property Get TotalPhysicalMemory() As String
  TotalPhysicalMemory = mTotalPhysicalMemory
End Property


Public Property Get AvailablePhysicalMemory() As String
  AvailablePhysicalMemory = mAvailablePhysicalMemory
End Property


Public Property Get TotalVirtualMemory() As String
  TotalVirtualMemory = mTotalVirtualMemory
End Property

Public Property Get AvailableVirtualMemory() As String
  AvailableVirtualMemory = mAvailableVirtualMemory
End Property

Public Property Get TotalPageFile() As String
  'TotalPageFile = mTotalPageFile
End Property

Public Property Get AvailablePageFile() As String

End Property


Public Property Get MemoryLoad() As String
  MemoryLoad = mMemoryLoad
End Property


Public Sub UpdateMemoryInformation()
  Dim udtMemStatEx       As MEMORYSTATUSEX

  udtMemStatEx.dwLength = Len(udtMemStatEx)
  Call GlobalMemoryStatusEx(udtMemStatEx)

  mTotalPhysicalMemory = NumberInKB(CLargeInt(udtMemStatEx.ulTotalPhys.LoPart, udtMemStatEx.ulTotalPhys.HiPart)) & " (" & Round(CLargeInt(udtMemStatEx.ulAvailPhys.LoPart, udtMemStatEx.ulAvailPhys.HiPart) / (CLargeInt(udtMemStatEx.ulTotalPhys.LoPart, udtMemStatEx.ulTotalPhys.HiPart)) * 100) & "% Free)"
  mAvailablePhysicalMemory = NumberInKB(CLargeInt(udtMemStatEx.ulAvailPhys.LoPart, udtMemStatEx.ulAvailPhys.HiPart))
  mTotalVirtualMemory = NumberInKB(CLargeInt(udtMemStatEx.ulTotalVirtual.LoPart, udtMemStatEx.ulTotalVirtual.HiPart))
  mAvailableVirtualMemory = NumberInKB(CLargeInt(udtMemStatEx.ulAvailVirtual.LoPart, udtMemStatEx.ulAvailVirtual.HiPart))
  'Me.Print "Total page file" + vbTab + NumberInKB(CLargeInt(udtMemStatEx.ulTotalPageFile.LoPart, udtMemStatEx.ulTotalPageFile.HiPart))
  'Me.Print "Available page file" + vbTab + NumberInKB(CLargeInt(udtMemStatEx.ulAvailPageFile.LoPart, udtMemStatEx.ulAvailPageFile.HiPart))
  'Me.Print "Available extended page file" + vbTab + NumberInKB(CLargeInt(udtMemStatEx.ulAvailExtendedVirtual.LoPart, udtMemStatEx.ulAvailExtendedVirtual.HiPart))
  mMemoryLoad = CStr(udtMemStatEx.dwMemoryLoad) + "%"

  Exit Sub
'
'    Dim hSnap   As Long
'    Dim pe      As PROCESSENTRY32
'    Dim hProc   As Long
'    Dim mi      As PROCESS_MEMORY_COUNTERS
'    Dim i       As Long
'    Dim li      As ListItem
'
'    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
'    If hSnap = INVALID_HANDLE_VALUE Then Exit Sub
'
'    pe.dwSize = Len(pe)
'
'    If Process32First(hSnap, pe) Then
'
'        Do
'            hProc = OpenProcess(IIf(IsVistaAndLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), False, pe.th32ProcessID)
'            If hProc Then
'                mi.cb = Len(mi)
'                GetProcessMemoryInfo hProc, mi, Len(mi)
'                If i >= lvwInfo.ListItems.Count Then
'                    Set li = lvwInfo.ListItems.Add(, , left$(pe.szExeFile, InStr(1, pe.szExeFile, vbNullChar)))
'                Else: Set li = lvwInfo.ListItems(i + 1)
'                End If
'                li.SubItems(1) = pe.th32ProcessID
'                li.SubItems(2) = LTrim(Format(mi.WorkingSetSize / 1024, "### ### ##0"))
'                li.SubItems(3) = LTrim(Format(mi.PagefileUsage / 1024, "### ### ##0"))
'                li.SubItems(4) = mi.PageFaultCount
'                li.SubItems(5) = LTrim(Format(mi.PeakPagefileUsage / 1024, "### ### ##0"))
'                li.SubItems(6) = LTrim(Format(mi.PeakWorkingSetSize / 1024, "### ### ##0"))
'                li.SubItems(7) = LTrim(Format(Int(mi.QuotaNonPagedPoolUsage / 1024), "### ### ##0"))
'                li.SubItems(8) = LTrim(Format(Int(mi.QuotaPagedPoolUsage / 1024), "### ### ##0"))
'                li.SubItems(9) = LTrim(Format(mi.QuotaPeakNonPagedPoolUsage / 1024, "### ### ##0"))
'                li.SubItems(10) = LTrim(Format(mi.QuotaPeakPagedPoolUsage / 1024, "### ### ##0"))
'                CloseHandle hProc
'                i = i + 1
'            End If
'
'        Loop While Process32Next(hSnap, pe)
'
'    End If
'
'    CloseHandle hSnap
'
'    If i < lvwInfo.ListItems.Count Then
'        Do Until lvwInfo.ListItems.Count = i
'            lvwInfo.ListItems.Remove (lvwInfo.ListItems.Count)
'        Loop
'    End If

End Sub

'This function converts the LARGE_INTEGER data type to a double
Private Function CLargeInt(Lo As Long, Hi As Long) As Double
  Dim dblLo              As Double
  Dim dblHi              As Double

  If Lo < 0 Then
    dblLo = 2 ^ 32 + Lo
  Else
    dblLo = Lo
  End If

  If Hi < 0 Then
    dblHi = 2 ^ 32 + Hi
  Else
    dblHi = Hi
  End If

  CLargeInt = dblLo + dblHi * 2 ^ 32

End Function

Public Function NumberInKB(ByVal vNumber As Currency) As String
  Dim strReturn          As String

  Select Case vNumber
    Case Is < 1024 ^ 1
      strReturn = CStr(vNumber) & " bytes"

    Case Is < 1024 ^ 2
      strReturn = CStr(Round(vNumber / 1024, 1)) & " KB"

    Case Is < 1024 ^ 3
      strReturn = CStr(Round(vNumber / 1024 ^ 2, 2)) & " MB"

    Case Is < 1024 ^ 4
      strReturn = CStr(Round(vNumber / 1024 ^ 3, 2)) & " GB"
  End Select

  NumberInKB = strReturn

End Function

Private Sub Class_Initialize()
    Dim ver As OSVERSIONINFO
    ver.dwOSVersionInfoSize = Len(ver)
    GetVersionEx ver
    IsVistaAndLater = ver.dwMajorVersion >= 6

End Sub

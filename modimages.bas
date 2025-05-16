Attribute VB_Name = "modImages"
Option Explicit
Const ChunkSize = 8192  '2048 increment multiples depending the image sizes

' power resizer spacific
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, iPic As StdPicture) As Long
Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type RGBtype
    b As Byte
    r As Byte
    g As Byte
End Type

Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Const DIB_RGB_COLORS = 0&
Public Const BI_RGB = 0&

Private Type BITMAPINFOHEADER
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Private Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type
Public Function PowerResize(Img As StdPicture, NewWidth As Long, NewHeight As Long) As StdPicture
    'Debug.Assert Img.Type = vbPicTypeBitmap 'Image must be a bitmap
        
    Dim SrcBmp As BITMAP
    GetObject Img.Handle, Len(SrcBmp), SrcBmp
        
    Dim srcBI As BITMAPINFO
    With srcBI.bmiHeader
        .biSize = Len(srcBI.bmiHeader)
        .biWidth = SrcBmp.bmWidth
        .biHeight = -SrcBmp.bmHeight
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
    End With

    'Create Source Bit Array
    Dim SrcBits() As RGBQUAD
    ReDim SrcBits(0 To SrcBmp.bmWidth - 1, 0 To SrcBmp.bmHeight - 1) As RGBQUAD

    'Grab Source Bits
    Dim lDc As Long
    lDc = CreateCompatibleDC(0)
    GetDIBits lDc, Img.Handle, 0, SrcBmp.bmHeight, SrcBits(0, 0), srcBI, DIB_RGB_COLORS
    DeleteDC lDc

    'Create Destination Bit Array
    Dim DblDstBits() As Double
    ReDim DblDstBits(0 To 3, 0 To NewWidth - 1, 0 To NewHeight - 1) As Double

    'Multipliers
    Dim xMult As Double, yMult As Double
    xMult = NewWidth / SrcBmp.bmWidth
    yMult = NewHeight / SrcBmp.bmHeight

    'Traversing variables
    Dim X As Long, XX As Long
    Dim y As Long, YY As Long
    
    'Low/High scan X/Y
    Dim lsX As Double, hsX As Double
    Dim lsY As Double, hsY As Double
    
    Dim OverlapWidth As Double
    Dim OverlapHeight As Double
    Dim Overlap As Double
    
    For X = 0 To SrcBmp.bmWidth - 1
        lsX = X * xMult
        hsX = X * xMult + xMult
        For y = 0 To SrcBmp.bmHeight - 1
            lsY = y * yMult
            hsY = y * yMult + yMult
            For XX = Fix(lsX) To IIf(Fix(hsX) = hsX, Fix(hsX), Fix(hsX + 1)) - 1
                For YY = Fix(lsY) To IIf(Fix(hsY) = hsY, Fix(hsY), Fix(hsY + 1)) - 1
                    OverlapWidth = 1
                    OverlapHeight = 1
                    
                    If XX < lsX Then OverlapWidth = 1# - (lsX - XX)
                    If XX + 1# > hsX Then OverlapWidth = OverlapWidth - (XX + 1# - hsX)
                    If YY < lsY Then OverlapHeight = 1# - (lsY - YY)
                    If YY + 1# > hsY Then OverlapHeight = OverlapHeight - (YY + 1# - hsY)
                    
                    Overlap = OverlapHeight * OverlapWidth
                    
                    DblDstBits(0, XX, YY) = DblDstBits(0, XX, YY) + SrcBits(X, y).rgbRed * Overlap
                    DblDstBits(1, XX, YY) = DblDstBits(1, XX, YY) + SrcBits(X, y).rgbGreen * Overlap
                    DblDstBits(2, XX, YY) = DblDstBits(2, XX, YY) + SrcBits(X, y).rgbBlue * Overlap
                    DblDstBits(3, XX, YY) = DblDstBits(3, XX, YY) + Overlap
                Next
            Next
        Next
    Next
    
    Dim DstBits() As RGBQUAD
    ReDim DstBits(0 To NewWidth - 1, 0 To NewHeight - 1) As RGBQUAD
    
    For X = 0 To NewWidth - 1
        For y = 0 To NewHeight - 1
            DstBits(X, y).rgbRed = Round(DblDstBits(0, X, y) / DblDstBits(3, X, y))
            DstBits(X, y).rgbGreen = Round(DblDstBits(1, X, y) / DblDstBits(3, X, y))
            DstBits(X, y).rgbBlue = Round(DblDstBits(2, X, y) / DblDstBits(3, X, y))
        Next
    Next
    
    Dim dstBI As BITMAPINFO
    With dstBI.bmiHeader
        .biSize = Len(dstBI.bmiHeader)
        .biWidth = NewWidth
        .biHeight = -NewHeight
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
    End With
    
    Dim hBmp As Long
    hBmp = CreateBitmap(NewWidth, NewHeight, 1, 32, ByVal 0)

    SetDIBits 0, hBmp, 0, NewHeight, DstBits(0, 0), dstBI, DIB_RGB_COLORS

    Dim IGuid As Guid
    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    Dim PicDst As PictDesc
    With PicDst
        .cbSizeofStruct = Len(PicDst)
        .hImage = hBmp
        .picType = vbPicTypeBitmap
    End With
    
    OleCreatePictureIndirect PicDst, IGuid, True, PowerResize
End Function

Function ReSizeIt(ByVal filename As String) As String

   Dim pic As StdPicture
   Dim NewPic As StdPicture
   Set pic = LoadPicture(filename)
   Set NewPic = PowerResize(pic, 64, 64)
   SavePicture NewPic, filename

End Function




Public Function SaveImageToDB(ByVal filename As String, datafield As Object) As Boolean
        Dim Size            As Long
        Dim Chunks          As Long
        Dim FragmentOffset  As Long
        Dim hfile           As Integer
        Dim chunk()         As Byte
        Dim offset          As Long
        Dim i               As Long
        Dim tempfilename    As String

10      On Error GoTo SaveImageToDB_Error

20      On Error GoTo BadImage


        tempfilename = ReduceImage(filename)


30      hfile = FreeFile
40      Open filename For Binary Access Read As hfile

50      Size = LOF(hfile)

60      Chunks = Size \ ChunkSize
70      FragmentOffset = Size Mod ChunkSize
80      ReDim chunk(FragmentOffset)
90      offset = FragmentOffset

100     Get hfile, , chunk()
110     datafield.AppendChunk chunk()
120     ReDim chunk(ChunkSize)
130     offset = FragmentOffset
140     For i = 1 To Chunks
150       Get hfile, , chunk()
160       datafield.AppendChunk chunk()
170       offset = offset + ChunkSize
180     Next
190     Close hfile

        Win32.DeleteFile tempfilename

200     SaveImageToDB = True
210     Exit Function
  
BadImage:
220     Close

SaveImageToDB_Resume:
230      On Error GoTo 0
240      Exit Function

SaveImageToDB_Error:

250     LogProgramError "Error " & Err.Number & " (" & Err.Description & ") at modImages.SaveImageToDB." & Erl
260     Resume SaveImageToDB_Resume

End Function

Public Function ReduceImage(ByVal filename As String) As String
  Dim imgSrc As cImage
  Dim imgDest As cImage
  Dim STDPIX As StdPicture
  Dim tempname As String
  Dim JPG As cJpeg
  Dim Buffer() As Byte

  

  tempname = App.path & "\~temp.jpg"

  Win32.DeleteFile tempname
  Const Qual = 70
  Set STDPIX = LoadPicture(filename)


  Set imgSrc = New cImage
  imgSrc.CopyStdPicture STDPIX
  Set imgDest = imgSrc.Resample(100, 100)


  Set JPG = New cJpeg
  JPG.Quality = Qual
  JPG.SampleHDC imgDest.hDC, 100, 100
  JPG.SaveFile tempname
  Set JPG = Nothing

  ReduceImage = tempname
End Function

Public Function GetImageFromDB(p As Object, datafield As Object) As Boolean

        Dim path            As String
        Dim hfile           As Integer
  
        Dim chunk()         As Byte
        Dim Size            As Long
        Dim Chunks          As Long
        Dim FragmentOffset  As Long
        Dim offset          As Long

        Dim i               As Long


20      On Error GoTo GetImageFromDB_Error

30      hfile = FreeFile
40      path = App.path
50      If Right(path, 1) <> "\" Then
60        path = path & "\"
70      End If

80      path = path & "~~temp.tmp"

90      Open path For Binary Access Write As hfile

100     Size = datafield.ActualSize
  
110     If Size > 0 Then

120     Chunks = Size \ ChunkSize
130     FragmentOffset = Size Mod ChunkSize

140     ReDim chunk(FragmentOffset) As Byte

150     chunk() = datafield.GetChunk(FragmentOffset)

160     Put hfile, , chunk()
170     offset = FragmentOffset
180     For i = 1 To Chunks
190       ReDim chunk(ChunkSize) As Byte
200       chunk() = datafield.GetChunk(ChunkSize)
210       Put hfile, , chunk()
220       offset = offset + ChunkSize
230     Next
240     Close hfile
250     p.Picture = LoadPicture(path)
260     Else
270       p.Picture = LoadPicture("")
280     End If

GetImageFromDB_Error:
290     Close


GetImageFromDB_Resume:
300      On Error GoTo 0
310      Exit Function



End Function


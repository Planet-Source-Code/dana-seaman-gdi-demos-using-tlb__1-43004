Attribute VB_Name = "GDIPlusAPI"
Option Explicit
                     
Private Const MODULE_NAME As String = "mdHelpers"

Public Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long
'=========================================================================
' Member vars
'=========================================================================
Private m_hToken        As Long
'=========================================================================
' GDI+ Constants
'=========================================================================
Public Const FlatnessDefault As Single = 1# / 4#


'-----------------------------------------------
' Helper Functions
'-----------------------------------------------

' Use this in lieu of the Color class constructor
' Thanks to Richard Mason for help with this
Public Function ColorARGB(ByVal alpha As Byte, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte) As Long
   Dim bytestruct       As COLORBYTES
   Dim result           As COLORLONG

   With bytestruct
      .AlphaByte = alpha
      .RedByte = Red
      .GreenByte = Green
      .BlueByte = Blue
   End With

   LSet result = bytestruct
   ColorARGB = result.longval
End Function

Public Function ColorSetAlpha(ByVal lColor As Long, ByVal alpha As Byte) As Long
   Dim bytestruct       As COLORBYTES
   Dim result           As COLORLONG

   result.longval = lColor
   LSet bytestruct = result

   bytestruct.AlphaByte = alpha

   LSet result = bytestruct
   ColorSetAlpha = result.longval
End Function

' Pass a GDI+ color to this function and get the VB compatible color
Public Function GetRGB_GDIP2VB(ByVal lColor As Long) As Long
   Dim argb             As COLORBYTES
   CopyMemory argb, lColor, 4
   GetRGB_GDIP2VB = RGB(argb.RedByte, argb.GreenByte, argb.BlueByte)
End Function

' Pass a VB/standard color to this function and get the GDI+ compatible color
Public Function GetRGB_VB2GDIP(ByVal lColor As Long, Optional ByVal alpha As Byte = 255) As Long
   Dim rgbq             As RGBQUAD
   CopyMemory rgbq, lColor, 4
   GetRGB_VB2GDIP = ColorARGB(alpha, rgbq.rgbBlue, rgbq.rgbGreen, rgbq.rgbRed)
End Function

' Built-in encoders for saving: (You can *try* to get other types also)
'   image/bmp
'   image/jpeg
'   image/gif
'   image/tiff
'   image/png
'
' Notes When Saving:
' The JPEG encoder supports the Transformation, Quality, LuminanceTable, and ChrominanceTable parameter categories.
' The TIFF encoder supports the Compression, ColorDepth, and SaveFlag parameter categories.
' The BMP, PNG, and GIF encoders no do not support additional parameters.
' Courtesy of: Vlad Vissoultchev
Public Function GetEncoderClsid(sMimeType As String) As CLSID
   Const FUNC_NAME     As String = "GetEncoderClsid"
   Dim lNumCoders       As Long
   Dim lSize            As Long
   Dim uInfo()          As ImageCodecInfo
   Dim lIdx             As Long
   
   On Error GoTo EH

   GdipError GdipGetImageEncodersSize(lNumCoders, lSize)
   If lSize > 0 Then
      ReDim uInfo(0 To lSize \ LenB(uInfo(0))) As ImageCodecInfo
      GdipError GdipGetImageEncoders(lNumCoders, lSize, uInfo(0))
      For lIdx = 0 To lNumCoders - 1
         If StrComp(PtrToStrW(uInfo(lIdx).MimeTypePtr), sMimeType, vbTextCompare) = 0 Then
            GetEncoderClsid = uInfo(lIdx).CLSID
            Exit For
         End If
      Next
   End If
   Exit Function
EH:
   RaiseError FUNC_NAME
End Function

' Courtesy of: Vlad Vissoultchev
Public Function GetDecoderClsid(sMimeType As String) As CLSID
   Const FUNC_NAME     As String = "GetDecoderClsid"
   Dim lNumCoders       As Long
   Dim lSize            As Long
   Dim uInfo()          As ImageCodecInfo
   Dim lIdx             As Long

   On Error GoTo EH

   GdipError GdipGetImageDecodersSize(lNumCoders, lSize)
   If lSize > 0 Then
      ReDim uInfo(0 To lSize \ LenB(uInfo(0))) As ImageCodecInfo
      GdipError GdipGetImageDecoders(lNumCoders, lSize, uInfo(0))
      For lIdx = 0 To lNumCoders - 1
         If StrComp(PtrToStrW(uInfo(lIdx).MimeTypePtr), sMimeType, vbTextCompare) = 0 Then
            GetDecoderClsid = uInfo(lIdx).CLSID
            Exit For
         End If
      Next
   End If
   Exit Function
EH:
   RaiseError FUNC_NAME
End Function
' Courtesy of: Vlad Vissoultchev
' Modified   : Dana Seaman - use CLSIDFromString Function in TLB
'                            to supply the magic IPictureGUID
Public Function LoadPictureEx( _
   sFileName As String, _
   Optional clrBack As OLE_COLOR = vbMagenta) As StdPicture
   Const FUNC_NAME     As String = "LoadPictureEx"
   Dim hImg             As Long
   Dim hBmp             As Long
   Dim uPictDesc        As PICTDESC

   On Error GoTo EH
   '--- state check
   InitGdip
   '--- load image
   GdipError GdipLoadImageFromFile(sFileName, hImg)
   '--- create HBITMAP
   GdipError GdipCreateHBITMAPFromBitmap(hImg, hBmp, clrBack)
   '--- fill struct
   With uPictDesc
      .size = Len(uPictDesc)
      .Type = vbPicTypeBitmap
      .hBmpOrIcon = hBmp
      .hPal = 0
   End With
   '--- Create picture from bitmap handle
   OleCreatePictureIndirect uPictDesc, CLSIDFromString(IPictureGUID), True, LoadPictureEx
   '--- deallocate
   GdipError GdipDisposeImage(hImg)
   Exit Function
EH:
   RaiseError FUNC_NAME
End Function

Public Function LoadDib( _
   sFileName As String, _
   uHead As BITMAPINFOHEADER, _
   pBits As Long) As Long
   Const FUNC_NAME     As String = "LoadDib"
   Dim hImg             As Long
   Dim lWidth           As Long
   Dim lHeight          As Long
   Dim hDC              As Long
   Dim hOldDIB          As Long
   Dim hGrfx            As Long

   On Error GoTo EH
   '--- state check
   InitGdip
   '--- load image
   GdipError GdipLoadImageFromFile(sFileName, hImg)
   '--- get dimensions
   GdipError GdipGetImageWidth(hImg, lWidth)
   GdipError GdipGetImageHeight(hImg, lHeight)
   '--- create DIB
   With uHead
      .biSize = Len(uHead)
      .biPlanes = 1
      .biBitCount = 32
      .biWidth = lWidth
      .biHeight = -lHeight
   End With
   '--- select DIB in mem hDC
   hDC = CreateCompatibleDC(0)
   LoadDib = CreateDIBSection(hDC, uHead, DIB_RGB_COLORS, pBits, 0, 0)
   hOldDIB = SelectObject(hDC, LoadDib)
   '--- paint GDI+ image to the hDC
   GdipError GdipCreateFromHDC(hDC, hGrfx)
   GdipError GdipDrawImageRectI(hGrfx, hImg, 0, 0, lWidth, lHeight)
   GdipError GdipDeleteGraphics(hGrfx)
   '--- deselect DIB
   SelectObject hDC, hOldDIB
   DeleteDC hDC
   '--- deallocate
   GdipError GdipDisposeImage(hImg)
   Exit Function
EH:
   RaiseError FUNC_NAME
End Function

Public Function SaveDib( _
   uHead As BITMAPINFOHEADER, _
   pBits As Long, _
   sFileName As String, _
   Optional sEncoder As String = "image/png") As Boolean
   Const FUNC_NAME      As String = "SaveDib"

   Dim hImg             As Long
   Dim uInfo            As BITMAPINFO
   Dim uEncParams       As EncoderParameters

   On Error GoTo EH
   '--- prepare struct
   uInfo.bmiHeader = uHead
   '--- create bitmap
   GdipError GdipCreateBitmapFromGdiDib(uInfo, ByVal pBits, hImg)
   '--- encode
   GdipError GdipSaveImageToFile(hImg, sFileName, GetEncoderClsid(sEncoder), uEncParams)
   '--- deallocate
   GdipError GdipDisposeImage(hImg)
   '--- success
   SaveDib = True
   Exit Function
EH:
   RaiseError FUNC_NAME
End Function

'=========================================================================
' Error handling (Vlad Vissoultchev)
'=========================================================================

Public Sub GdipError(ByVal lStatus As GpStatus)
   Const STR_ERRORS    As String = "Ok|Generic Error|Invalid Parameter|Out Of Memory|Object Busy|Insufficient Buffer|Not Implemented|Win32 Error|Wrong State|Aborted|File Not Found|Value Overflow|Access Denied|Unknown Image Format|Font Family Not Found|Font Style Not Found|Not TrueType Font|Unsupported GDI+ Version|GDI+ Not Initialized|Property Not Found|Property Not Supported"
   Dim vSplit           As Variant
   Dim sStatus          As String

   If lStatus <> Ok Then
      vSplit = Split(STR_ERRORS, "|")
      If lStatus >= LBound(vSplit) And lStatus <= UBound(vSplit) Then
         sStatus = vSplit(lStatus)
      Else
         sStatus = "Unknown"
      End If
      Err.Raise vbObjectError, "GDI+", sStatus & ": Status = " & lStatus
   End If
End Sub

Private Sub RaiseError(sFunc As String)
   Err.Raise Err.Number, MODULE_NAME & "." & sFunc & vbCrLf & Err.Source, Err.Description
End Sub

'=========================================================================
' Methods
'=========================================================================

Public Function InitGdip() As Boolean
   Const FUNC_NAME     As String = "InitGdip"
   Dim uInput           As GdiplusStartupInput

   If m_hToken = 0 Then
      uInput.GdiplusVersion = 1
      GdipError GdiplusStartup(m_hToken, uInput)
      '--- success
      InitGdip = True
   End If
   Exit Function
EH:
   RaiseError FUNC_NAME
End Function

Public Function ShutdownGdip() As Boolean
   If m_hToken <> 0 Then
      GdiplusShutdown m_hToken
      m_hToken = 0
   End If
End Function

Public Function AppPath() As String
   AppPath = App.path
   If Right$(AppPath, 1) <> "\" Then
      AppPath = AppPath & "\"
   End If
End Function

Public Function GetInstalledDecoders() As Collection
   Const FUNC_NAME     As String = "GetInstalledDecoders"
   Dim lNumCoders       As Long
   Dim lSize            As Long
   Dim uInfo()          As ImageCodecInfo
   Dim lIdx             As Long

   On Error GoTo EH
   '--- state check
   InitGdip
   GdipError GdipGetImageDecodersSize(lNumCoders, lSize)
   If lSize > 0 Then
      ReDim uInfo(0 To lSize \ LenB(uInfo(0))) As ImageCodecInfo
      GdipError GdipGetImageDecoders(lNumCoders, lSize, uInfo(0))
      Set GetInstalledDecoders = New Collection
      For lIdx = 0 To lNumCoders - 1
         GetInstalledDecoders.Add Array( _
            PtrToStrW(uInfo(lIdx).CodecNamePtr), _
            PtrToStrW(uInfo(lIdx).FilenameExtensionPtr), _
            PtrToStrW(uInfo(lIdx).FormatDescriptionPtr), _
            PtrToStrW(uInfo(lIdx).MimeTypePtr), _
            PtrToStrW(uInfo(lIdx).DllNamePtr))
      Next
   End If
   Exit Function
EH:
   RaiseError FUNC_NAME
End Function

Public Function GetInstalledEncoders() As Collection
   Const FUNC_NAME     As String = "GetInstalledEncoders"
   Dim lNumCoders       As Long
   Dim lSize            As Long
   Dim uInfo()          As ImageCodecInfo
   Dim lIdx             As Long

   On Error GoTo EH
   '--- state check
   InitGdip
   GdipError GdipGetImageEncodersSize(lNumCoders, lSize)
   If lSize > 0 Then
      ReDim uInfo(0 To lSize \ LenB(uInfo(0))) As ImageCodecInfo
      GdipError GdipGetImageEncoders(lNumCoders, lSize, uInfo(0))
      Set GetInstalledEncoders = New Collection
      For lIdx = 0 To lNumCoders - 1
         GetInstalledEncoders.Add Array( _
            PtrToStrW(uInfo(lIdx).CodecNamePtr), _
            PtrToStrW(uInfo(lIdx).FilenameExtensionPtr), _
            PtrToStrW(uInfo(lIdx).FormatDescriptionPtr), _
            PtrToStrW(uInfo(lIdx).MimeTypePtr), _
            PtrToStrW(uInfo(lIdx).DllNamePtr))
      Next
   End If
   Exit Function
EH:
   RaiseError FUNC_NAME
End Function

' From www.mvps.org/vbnet...i think
'   Dereferences an ANSI or Unicode string pointer
'   and returns a normal VB BSTR
Public Function PtrToStrW(ByVal lpsz As Long) As String
   Dim sOut             As String
   Dim lLen             As Long

   lLen = lstrlenW(lpsz)

   If (lLen > 0) Then
      'was sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
      sOut = String$(lLen * 2, vbNullChar)
      CopyMemory ByVal sOut, ByVal lpsz, lLen * 2
      PtrToStrW = StrConv(sOut, vbFromUnicode)
   End If
End Function

Public Function PtrToStrA(ByVal lpsz As Long) As String
   Dim sOut             As String
   Dim lLen             As Long

   lLen = lstrlenA(lpsz)

   If (lLen > 0) Then
      sOut = String$(lLen, vbNullChar)
      CopyMemory ByVal sOut, ByVal lpsz, lLen
      PtrToStrA = sOut
   End If
End Function

' This should hopefully simplify property item value retrieval
' NOTE: We are raising errors in this function; ensure the caller has error handing code.
'       The resulting arrays are using a base of one.
Public Function GetPropValue(item As PropertyItem) As Variant
   ' We need a valid pointer and length
   If item.ValuePtr = 0 Or item.Length = 0 Then Err.Raise 5, "GetPropValue"

   Select Case item.Type
      ' We'll make Undefined types a Btye array as it seems the safest choice...
      Case PropertyTagTypeByte, PropertyTagTypeUndefined:
         Dim bte()            As Byte: ReDim bte(1 To item.Length)
         CopyMemory bte(1), ByVal item.ValuePtr, item.Length
         GetPropValue = bte
         Erase bte

      Case PropertyTagTypeASCII:
         GetPropValue = PtrToStrA(item.ValuePtr)

      Case PropertyTagTypeShort:
         Dim short()          As Integer: ReDim short(1 To (item.Length / 2))
         CopyMemory short(1), ByVal item.ValuePtr, item.Length
         GetPropValue = short
         Erase short

      Case PropertyTagTypeLong, PropertyTagTypeSLONG:
         Dim lng()            As Long: ReDim lng(1 To (item.Length / 4))
         CopyMemory lng(1), ByVal item.ValuePtr, item.Length
         GetPropValue = lng
         Erase lng

      Case PropertyTagTypeRational, PropertyTagTypeSRational:
         Dim lngpair()        As Long: ReDim lngpair(1 To (item.Length / 8), 1 To 2)
         CopyMemory lngpair(1, 1), ByVal item.ValuePtr, item.Length
         GetPropValue = lngpair
         Erase lngpair

      Case Else: Err.Raise 461, "GetPropValue"
   End Select
End Function


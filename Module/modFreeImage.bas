Attribute VB_Name = "modFreeImage"
Option Explicit
'=========================
Private Const c_strModule As String = "modFreeImage"
'=========================
' Описание      : Visual Basic Wrapper for FreeImage 3
' Версия        : 2.24.4.453644468
' Дата          : 13.03.2024 10:43:24
' Автор         : Carsten Klein (cklein05@users.sourceforge.net)
' Примечание    : https://freeimage.sourceforge.io/download.html
'               : if lib is missed in sys folder, first call: FreeImage_LoadLibrary to load it from LibPath to memory
'               : it seems that Optional params in FreeImage declares, need to be replaced
'               : especially in: FreeImage_Allocate family - this consistently leads to silent crashes
' v.2.24.3      : 02.09.2021 - added FreeImage_LoadBitmapFromMemoryEx to allow load FIBITMAP suppported sources with option to choose bitmap for multibitmap sources by idx or size (ICO) {Кашкин Р.В. (KashRus@gmail.com)}
' v.2.24.1      : 31.07.2021 - x64 support. adapted very superficially {Кашкин Р.В. (KashRus@gmail.com)}
' v.2.24        : 16.03.2015 - original from http://downloads.sourceforge.net/freeimage/FreeImage3180Win32Win64.zip \FreeImage\Wrapper\VB6\src\
'=========================
Private Const c_strLibPath = "\INC\"                ' libraries path relative to CurrentProject.Path
#If Win64 Then          '<WIN64>
Private Const c_strLibName = "FreeImage_x64.dll"    ' x64 lib file name (v3.18)
#Else                   '<WIN32>
Private Const c_strLibName = "FreeImage.dll"        ' x32 lib file name (v3.17, because of error on call LoadLibrary for FI v3.18 on MSO2003 WinXP x86)
#End If                 '<WIN64>
Public FreeImage_IsLoaded As Boolean ' Flag to check if FI is loaded in memory
'// ==========================================================
'// Visual Basic Wrapper for FreeImage 3
'// Original FreeImage 3 functions and VB compatible derived functions
'// Design and implementation by
'// - Carsten Klein (cklein05@users.sourceforge.net)
'//
'// Main reference : Curland, Matthew., Advanced Visual Basic 6, Addison Wesley, ISBN 0201707128, (c) 2000
'//                  Steve McMahon, creator of the excellent site vbAccelerator at http://www.vbaccelerator.com/
'//                  MSDN Knowlede Base
'//
'// COVERED CODE IS PROVIDED UNDER THIS LICENSE ON AN "AS IS" BASIS, WITHOUT WARRANTY
'// OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, WITHOUT LIMITATION, WARRANTIES
'// THAT THE COVERED CODE IS FREE OF DEFECTS, MERCHANTABLE, FIT FOR A PARTICULAR PURPOSE
'// OR NON-INFRINGING. THE ENTIRE RISK AS TO THE QUALITY AND PERFORMANCE OF THE COVERED
'// CODE IS WITH YOU. SHOULD ANY COVERED CODE PROVE DEFECTIVE IN ANY RESPECT, YOU (NOT
'// THE INITIAL DEVELOPER OR ANY OTHER CONTRIBUTOR) ASSUME THE COST OF ANY NECESSARY
'// SERVICING, REPAIR OR CORRECTION. THIS DISCLAIMER OF WARRANTY CONSTITUTES AN ESSENTIAL
'// PART OF THIS LICENSE. NO USE OF ANY COVERED CODE IS AUTHORIZED HEREUNDER EXCEPT UNDER
'// THIS DISCLAIMER.
'//
'// Use at your own risk!
'// ==========================================================

'// ==========================================================
'// CVS
'// $Revision: 2.24 $
'// $Date: 2015/03/16 06:29:34 $
'// $Id: MFreeImage.bas,v 2.24 2015/03/16 06:29:34 cklein05 Exp $
'// ==========================================================
' Plain, single page image:
'
' type:     FIBITMAP
' creation: FreeImage_Load(), FreeImage_Allocate(), FreeImage_Clone(),[...]
' destruct: FreeImage_Unload()
'
' Multi Bitmap, all pages collection:
'
' type:     FIMULTIBITMAP
' creation: FreeImage_OpenMultiBitmap(), FreeImage_LoadMultiBitmapFromMemory()
' destruct: FreeImage_CloseMultiBitmap()
'
' Single page of a multi bitmap:
'
' type:     FIBITMAP
' creation: FreeImage_LockPage()
' destruct: FreeImage_UnlockPage()

'----------------------
' POINTER LENGTH
'----------------------
' !!! OLE_HANDLE Is 32-bit Long Even on 64-bit Windows !!!
' This means that you need to change any code that assumed OLE_HANDLE and other HANDLE types are interchangeable.
' ??? Something like this ???:
'    Dim n As LongPtr   'Dim n As IntPtr    ' intptr_t n;
'    Dim ptr As Long    'Dim n As Single    ' float *ptr;
'    ...
'    ptr = CLng(n)      'ptr = CSng(n)       ' ptr = (float *)(n);
'----------------------
' POINTER
'----------------------
'#If VBA7 = 0 Then       'LongPtr trick by @Greedo (https://github.com/Greedquest)
'Public Enum LongPtr
'    [_]
'End Enum
'#End If
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>  PtrSafe, LongPtr and LongLong
Private Const PTR_LENGTH As Long = 8
Private Const VARIANT_SIZE As Long = 24
#Else                   '<OFFICE97-2010>        Long
Private Const PTR_LENGTH As Long = 4
Private Const VARIANT_SIZE As Long = 16
#End If                 '<WIN32>
'----------------------
' Win32 API function, structure and constant declarations
'----------------------
Private Const NOERROR As Long = 0
'----------------------
' KERNEL32
'----------------------
#If VBA7 Then           '<OFFICE2010+>
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As LongPtr) As Long
#Else                   '<OFFICE97-2007>
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long
#End If                 '<VBA7>
'----------------------
' OLE
'----------------------
#If VBA7 Then           '<OFFICE2010+>
Private Declare PtrSafe Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Any, pclsid As Any) As Long
Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32.dll" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare PtrSafe Function OleLoadPicture Lib "oleaut32.dll" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvPic As Any) As LongPtr
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef lpPictDesc As PICTDESC, ByRef riid As GUID, ByVal fOwn As Long, ByRef lplpvObj As IPicture) As Long
Private Declare PtrSafe Function SafeArrayAllocDescriptor Lib "oleaut32.dll" (ByVal cDims As Long, ByRef ppsaOut As LongPtr) As Long
Private Declare PtrSafe Function SafeArrayDestroyDescriptor Lib "oleaut32.dll" (ByVal psa As LongPtr) As Long
Private Declare PtrSafe Sub SafeArrayDestroyData Lib "oleaut32.dll" (ByVal psa As LongPtr)
Private Declare PtrSafe Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As OLE_COLOR, ByVal hPal As Long, ByRef lpcolorref As Long) As Long
#Else                   '<OFFICE97-2007>
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Any, pclsid As Any) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32.dll" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvPic As Any) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (ByRef lpPictDesc As PICTDESC, ByRef riid As GUID, ByVal fOwn As Long, ByRef lplpvObj As IPicture) As Long
Private Declare Function SafeArrayAllocDescriptor Lib "oleaut32.dll" (ByVal cDims As Long, ByRef ppsaOut As Long) As Long
Private Declare Function SafeArrayDestroyDescriptor Lib "oleaut32.dll" (ByVal psa As Long) As Long
Private Declare Sub SafeArrayDestroyData Lib "oleaut32.dll" (ByVal psa As Long)
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As OLE_COLOR, ByVal hPal As Long, ByRef lpcolorref As Long) As Long
#End If                 '<VBA7>
Private Const CLR_INVALID As Long = &HFFFF&
'----------------------
' SAFEARRAY
'----------------------
Private Const FADF_AUTO As Long = (&H1)
Private Const FADF_FIXEDSIZE As Long = (&H10)
'Type SAFEARRAYBOUND         ' 8 bytes
'    cElements As Long               ' +0 Количество элементов в размерности
'    lLbound As Long                 ' +4 Нижняя граница размерности
'End Type
'Type SAFEARRAY
'    cDims           As Integer      ' +0  Число размерностей
'    fFeatures       As Integer      ' +2  Флаг, используется функциями SafeArray
'    cbElements      As Long         ' +4  Размер одного элемента в байтах
'    cLocks          As Long         ' +8 (x86) Cчетчик ссылок, указывающий количество блокировок, наложенных на массив.
'                    As LongLong     '    (x64)
'    pvData          As Long         ' +12(x86) Указатель на данные
'                    As LongPtr      ' +16(x64)
'    rgSAbound As SAFEARRAYBOUND     ' Повторяется для каждой размерности (размер = n*8 bytes, n- кол-во размерностей массива)
'                                    ' +16(x86) rgSAbound.cElements (Long) - Количество элементов в размерности
'                                    ' +24(x64)
'                                    ' +20(x86) rgSAbound.lLbound (Long)   - Нижняя граница размерности
'                                    ' +28(x64)
'End Type
Private Type SAFEARRAY1D
   cDims As Integer
   fFeatures As Integer
   cbElements As Long
#If VBA7 And Win64 Then     ' <OFFICE2010+ & WIN64> use: LongPtr and LongLong
   cLocks As LongLong       ' << !!???
#Else                       ' <OFFICE97-2007>       use: Long only
   cLocks As Long
#End If                     ' <VBA7 & WIN64>
   pvData As LongPtr
   cElements As Long
   lLbound As Long
End Type
Private Type SAFEARRAY2D
   cDims As Integer
   fFeatures As Integer
   cbElements As Long
#If VBA7 And Win64 Then     ' <OFFICE2010+ & WIN64> use: LongPtr and LongLong
   cLocks As LongLong       ' << !!???
#Else                       ' <OFFICE97-2007>       use: Long only
   cLocks As Long
#End If                     ' <VBA7 & WIN64>
   pvData As LongPtr
   cElements1 As Long
   lLbound1 As Long
   cElements2 As Long
   lLbound2 As Long
End Type
'----------------------
' MSVBA
'----------------------
#If VBA7 And Win64 Then     ' <WIN64 & OFFICE2010+>
Private Declare PtrSafe Function VarPtrArray Lib "VBE7.dll" Alias "VarPtr" (ByRef Ptr() As Any) As LongPtr
#ElseIf VBA7 Then           ' <WIN32 & OFFICE2010+>
Private Declare Function VarPtrArray Lib "VBE7.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
'#Else                       ' <OFFICE2003-2010>
'Private Declare Function VarPtrArray Lib "VBE6.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
#Else                       ' <OFFICE2000-2003>
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
'#Else                       ' <OFFICE97-2000>
'Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
#End If                     ' <VBA7 & WIN64>
'----------------------
' USER32
'----------------------
Private Const DCX_WINDOW As Long = &H1&
#If VBA7 Then           '<OFFICE2010+>
Private Declare PtrSafe Function ReleaseDC Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDC Lib "user32.dll" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowDC Lib "user32.dll" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDCEx Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal hRgnClip As LongPtr, ByVal fdwOptions As Long) As LongPtr
Private Declare PtrSafe Function GetDesktopWindow Lib "user32.dll" () As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "user32.dll" (ByVal hwnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32.dll" (ByVal hwnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function DestroyIcon Lib "user32.dll" (ByVal hIcon As LongPtr) As Long
Private Declare PtrSafe Function CreateIconIndirect Lib "user32.dll" (ByRef piconinfo As ICONINFO) As Long 'Ptr
#Else                   '<OFFICE97-2007>
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetDCEx Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgnClip As Long, ByVal fdwOptions As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32.dll" (ByRef piconinfo As ICONINFO) As Long
#End If                 '<VBA7>
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Enum PICTYPE
    PICTYPE_UNINITIALIZED = -1
    PICTYPE_NONE = 0
    PICTYPE_BITMAP = 1
    PICTYPE_METAFILE = 2
    PICTYPE_ICON = 3
    PICTYPE_ENHMETAFILE = 4
End Enum
'Public Type uPicDesc
'        Size As Long
'        Type As Long
'#If VBA7 Then
'        hPic As LongPtr
'        hPal As LongPtr
'#Else
'        hPic As Long
'        hPal As Long
'#End If
'End Type
Private Type PICTDESC
    Size As Long
    Type As Long         'PICTYPE
    hPic As LongPtr
    hPal As LongPtr
End Type
Private Type BITMAP_API
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As LongPtr
End Type
Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As LongPtr
    hbmColor As LongPtr
End Type
Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type
Private Const STRETCH_HALFTONE As Long = &H4&
'----------------------
' GDI32
'----------------------
Private Const HORZSIZE = 4              ' Horizontal size in millimeters
Private Const VERTSIZE = 6              ' Vertical size in millimeters
Private Const HORZRES = 8               ' Horizontal width in pixels
Private Const VERTRES = 10              ' Vertical width in pixels
Private Const LOGPIXELSX = 88           ' Logical pixels/inch in X
Private Const LOGPIXELSY = 90           ' Logical pixels/inch in Y
Private Const CBM_INIT As Long = &H4&
Private Const OBJ_BITMAP As Long = &H7&
Private Const CF_BITMAP = &H2&          ' Predefined Clipboard Formats
Private Const CF_ENHMETAFILE = &HE&
Private Const CF_METAFILEPICT = &H3&
Private Const CF_DIB = &H8&
Private Const MM_ANISOTROPIC = &H8&     ' Map mode anisotropic

Private Const SYSTEM_FONT = 13          ' System font
Private Const LF_FACESIZE As Long = 32
Private Const LF_FACESIZEW As Long = LF_FACESIZE * 2

Private Type POINT
    cX As Long
    cY As Long
End Type
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte  'lfFaceName As String * LF_FACESIZE
End Type
Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type
#If VBA7 Then           '<OFFICE2010+>
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetStretchBltMode Lib "gdi32.dll" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function SetStretchBltMode Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal nStretchMode As Long) As Long
Private Declare PtrSafe Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal dX As Long, ByVal dY As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal Scan As Long, ByVal NumScans As Long, ByVal Bits As LongPtr, ByVal BitsInfo As LongPtr, ByVal wUsage As Long) As Long
Private Declare PtrSafe Function StretchDIBits Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal dX As Long, ByVal dY As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByVal lpBits As LongPtr, ByVal lpBitsInfo As LongPtr, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function CreateDIBitmap Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal lpInfoHeader As LongPtr, ByVal dwUsage As Long, ByVal lpInitBits As LongPtr, ByVal lpInitInfo As LongPtr, ByVal wUsage As Long) As LongPtr
Private Declare PtrSafe Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal pbmi As LongPtr, ByVal iUsage As Long, ByRef ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal dwOffset As Long) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32.dll" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, ByVal lpBits As LongPtr, ByVal lpBI As LongPtr, ByVal wUsage As Long) As Long
Private Declare PtrSafe Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32.dll" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function GetCurrentObject Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal uObjectType As Long) As LongPtr
Private Declare PtrSafe Function SetMapMode Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal nMapMode As Long) As Long
Private Declare PtrSafe Function CreateIC Lib "gdi32.dll" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As LongPtr
Private Declare PtrSafe Function CreateEnhMetaFile Lib "gdi32.dll" Alias "CreateEnhMetaFileA" (ByVal hdcRef As LongPtr, ByVal lpFileName As String, lpRect As RECT, ByVal lpDescription As String) As LongPtr
Private Declare PtrSafe Function CloseEnhMetaFile Lib "gdi32.dll" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function GetEnhMetaFileBits Lib "gdi32.dll" (ByVal hEmf As LongPtr, ByVal cbBuffer As Long, lpbBuffer As Any) As Long
Private Declare PtrSafe Function SetWindowExtEx Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpSize As Any) As Long
Private Declare PtrSafe Function SetViewportExtEx Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpSize As Any) As Long
' used to create the checkerboard pattern on demand
Private Declare PtrSafe Function FillRect Lib "user32.dll" (ByVal hdc As LongPtr, ByRef lpRect As RECT, ByVal hBrush As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As LongPtr
' used to create font
Private Declare PtrSafe Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As LongPtr
Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As LongPtr
' used to text output
Private Declare PtrSafe Function SetTextColor Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
Private Declare PtrSafe Function SetBkColor Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
Private Declare PtrSafe Function SetBkMode Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal nBkMode As Long) As Long
Private Declare PtrSafe Function TextOut Lib "gdi32.dll" Alias "TextOutW" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal lpString As LongPtr, ByVal nCount As Long) As Long
Private Declare PtrSafe Function SetTextAlign Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32W" (ByVal hdc As LongPtr, ByVal lpsz As LongPtr, ByVal cbString As Long, ByRef lpSize As POINT) As Long
Private Declare PtrSafe Function GetTextMetrics Lib "gdi32.dll" Alias "GetTextMetricsW" (ByVal hdc As LongPtr, lpMetrics As TEXTMETRIC) As Long
#Else                       '<OFFICE97-2007>
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dX As Long, ByVal dY As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal Scan As Long, ByVal NumScans As Long, ByVal Bits As Long, ByVal BitsInfo As Long, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dX As Long, ByVal dY As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByVal lpBits As Long, ByVal lpBitsInfo As Long, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDIBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal lpInfoHeader As Long, ByVal dwUsage As Long, ByVal lpInitBits As Long, ByVal lpInitInfo As Long, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, ByVal pbmi As Long, ByVal iUsage As Long, ByRef ppvBits As Long, ByVal hSection As Long, ByVal dwOffset As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByVal lpBits As Long, ByVal lpBI As Long, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Private Declare Function CreateIC Lib "gdi32.dll" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function SetMapMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As Long, ByVal lpFileName As String, lpRect As RECT, ByVal lpDescription As String) As Long
Private Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetEnhMetaFileBits Lib "gdi32" (ByVal hEmf As Long, ByVal cbBuffer As Long, lpbBuffer As Any) As Long
Private Declare Function SetWindowExtEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Any) As Long
Private Declare Function SetViewportExtEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Any) As Long
' used to create the checkerboard pattern on demand
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
' used to create font
Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
' used to text output
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutW" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32.dll" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32W" (ByVal hdc As Long, ByVal lpsz As Long, ByVal cbString As Long, ByRef lpSize As POINT) As Long
Private Declare Function GetTextMetrics Lib "gdi32.dll" Alias "GetTextMetricsW" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
#End If                 '<VBA7>
'----------------------
'MSIMG32
'----------------------
#If VBA7 Then           '<OFFICE2010+>
Private Declare PtrSafe Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As LongPtr, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As LongPtr, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal lBlendFunction As Long) As Long
#Else                       ' <OFFICE97-2007>
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal lBlendFunction As Long) As Long
#End If                 '<VBA7>
Private Const AC_SRC_OVER = &H0
Private Const AC_SRC_ALPHA = &H1
'Private Const AC_SRC_NO_PREMULT_ALPHA = &H1
'Private Const AC_SRC_NO_ALPHA = &H2
'Private Const AC_DST_NO_PREMULT_ALPHA = &H10
'Private Const AC_DST_NO_ALPHA = &H20

Private Const BLACKONWHITE As Long = 1
Private Const WHITEONBLACK As Long = 2
Private Const COLORONCOLOR As Long = 3
Public Enum STRETCH_MODE
   SM_BLACKONWHITE = BLACKONWHITE
   SM_WHITEONBLACK = WHITEONBLACK
   SM_COLORONCOLOR = COLORONCOLOR
End Enum
Private Const SRCAND As Long = &H8800C6
Private Const SRCCOPY As Long = &HCC0020
Private Const SRCERASE As Long = &H440328
Private Const SRCINVERT As Long = &H660046
Private Const SRCPAINT As Long = &HEE0086
Private Const CAPTUREBLT As Long = &H40000000
Public Enum RASTER_OPERATOR
   ROP_SRCAND = SRCAND
   ROP_SRCCOPY = SRCCOPY
   ROP_SRCERASE = SRCERASE
   ROP_SRCINVERT = SRCINVERT
   ROP_SRCPAINT = SRCPAINT
End Enum
Private Const DIB_PAL_COLORS As Long = 1&
Private Const DIB_RGB_COLORS As Long = 0&
Public Enum DRAW_MODE
   DM_DRAW_DEFAULT = &H0
   DM_MIRROR_NONE = DM_DRAW_DEFAULT
   DM_MIRROR_VERTICAL = &H1
   DM_MIRROR_HORIZONTAL = &H2
   DM_MIRROR_BOTH = DM_MIRROR_VERTICAL Or DM_MIRROR_HORIZONTAL
End Enum
Public Enum HISTOGRAM_ORIENTATION
   HOR_TOP_DOWN = &H0
   HOR_BOTTOM_UP = &H1
End Enum
'----------------------
' FreeImage 3 types, constants and enumerations
'----------------------
'FREEIMAGE
' Version information
Public Const FREEIMAGE_MAJOR_VERSION As Long = 3
Public Const FREEIMAGE_MINOR_VERSION As Long = 17
Public Const FREEIMAGE_RELEASE_SERIAL As Long = 0
' Memory stream pointer operation flags
Public Const SEEK_SET As Long = 0
Public Const SEEK_CUR As Long = 1
Public Const SEEK_END As Long = 2
' Indexes for byte arrays, masks and shifts for treating pixels as words
' These coincide with the order of RGBQUAD and RGBTRIPLE
' Little Endian (x86 / MS Windows, Linux) : BGR(A) order
Public Const FI_RGBA_RED As Long = 2
Public Const FI_RGBA_GREEN As Long = 1
Public Const FI_RGBA_BLUE As Long = 0
Public Const FI_RGBA_ALPHA As Long = 3
Public Const FI_RGBA_RED_MASK As Long = &HFF0000
Public Const FI_RGBA_GREEN_MASK As Long = &HFF00
Public Const FI_RGBA_BLUE_MASK As Long = &HFF
Public Const FI_RGBA_ALPHA_MASK As Long = &HFF000000
Public Const FI_RGBA_RED_SHIFT As Long = 16
Public Const FI_RGBA_GREEN_SHIFT As Long = 8
Public Const FI_RGBA_BLUE_SHIFT As Long = 0
Public Const FI_RGBA_ALPHA_SHIFT As Long = 24
' The 16 bit macros only include masks and shifts, since each color element is not byte aligned
Public Const FI16_555_RED_MASK As Long = &H7C00
Public Const FI16_555_GREEN_MASK As Long = &H3E0
Public Const FI16_555_BLUE_MASK As Long = &H1F
Public Const FI16_555_RED_SHIFT As Long = 10
Public Const FI16_555_GREEN_SHIFT As Long = 5
Public Const FI16_555_BLUE_SHIFT As Long = 0
Public Const FI16_565_RED_MASK As Long = &HF800
Public Const FI16_565_GREEN_MASK As Long = &H7E0
Public Const FI16_565_BLUE_MASK As Long = &H1F
Public Const FI16_565_RED_SHIFT As Long = 11
Public Const FI16_565_GREEN_SHIFT As Long = 5
Public Const FI16_565_BLUE_SHIFT As Long = 0
' ICC profile support
Public Const FIICC_DEFAULT As Long = &H0
Public Const FIICC_COLOR_IS_CMYK As Long = &H1
Private Const FREE_IMAGE_ICC_COLOR_MODEL_MASK As Long = &H1
Public Enum FREE_IMAGE_ICC_COLOR_MODEL
   FIICC_COLOR_MODEL_RGB = &H0
   FIICC_COLOR_MODEL_CMYK = &H1
End Enum
' Load / Save flag constants
Public Const FIF_LOAD_NOPIXELS As Long = &H8000      ' load the image header only (not supported by all plugins)
Public Const BMP_DEFAULT As Long = 0
Public Const BMP_SAVE_RLE As Long = 1
Public Const CUT_DEFAULT As Long = 0
Public Const DDS_DEFAULT As Long = 0
Public Const EXR_DEFAULT As Long = 0                 ' save data as half with piz-based wavelet compression
Public Const EXR_FLOAT As Long = &H1                 ' save data as float instead of as half (not recommended)
Public Const EXR_NONE As Long = &H2                  ' save with no compression
Public Const EXR_ZIP As Long = &H4                   ' save with zlib compression, in blocks of 16 scan lines
Public Const EXR_PIZ As Long = &H8                   ' save with piz-based wavelet compression
Public Const EXR_PXR24 As Long = &H10                ' save with lossy 24-bit float compression
Public Const EXR_B44 As Long = &H20                  ' save with lossy 44% float compression - goes to 22% when combined with EXR_LC
Public Const EXR_LC As Long = &H40                   ' save images with one luminance and two chroma channels, rather than as RGB (lossy compression)
Public Const FAXG3_DEFAULT As Long = 0
Public Const GIF_DEFAULT As Long = 0
Public Const GIF_LOAD256 As Long = 1                 ' Load the image as a 256 color image with ununsed palette entries, if it's 16 or 2 color
Public Const GIF_PLAYBACK As Long = 2                ''Play' the GIF to generate each frame (as 32bpp) instead of returning raw frame data when loading
Public Const HDR_DEFAULT As Long = 0
Public Const ICO_DEFAULT As Long = 0
Public Const ICO_MAKEALPHA As Long = 1               ' convert to 32bpp and create an alpha channel from the AND-mask when loading
Public Const IFF_DEFAULT As Long = 0
Public Const J2K_DEFAULT  As Long = 0                ' save with a 16:1 rate
Public Const JP2_DEFAULT As Long = 0                 ' save with a 16:1 rate
Public Const JPEG_DEFAULT As Long = 0                ' loading (see JPEG_FAST); saving (see JPEG_QUALITYGOOD|JPEG_SUBSAMPLING_420)
Public Const JPEG_FAST As Long = &H1                 ' load the file as fast as possible, sacrificing some quality
Public Const JPEG_ACCURATE As Long = &H2             ' load the file with the best quality, sacrificing some speed
Public Const JPEG_CMYK As Long = &H4                 ' load separated CMYK "as is" (use 'OR' to combine with other flags)
Public Const JPEG_EXIFROTATE As Long = &H8           ' load and rotate according to Exif 'Orientation' tag if available
Public Const JPEG_GREYSCALE As Long = &H10           ' load and convert to a 8-bit greyscale image
Public Const JPEG_QUALITYSUPERB As Long = &H80       ' save with superb quality (100:1)
Public Const JPEG_QUALITYGOOD As Long = &H100        ' save with good quality (75:1)
Public Const JPEG_QUALITYNORMAL As Long = &H200      ' save with normal quality (50:1)
Public Const JPEG_QUALITYAVERAGE As Long = &H400     ' save with average quality (25:1)
Public Const JPEG_QUALITYBAD As Long = &H800         ' save with bad quality (10:1)
Public Const JPEG_PROGRESSIVE As Long = &H2000       ' save as a progressive-JPEG (use 'OR' to combine with other save flags)
Public Const JPEG_SUBSAMPLING_411 As Long = &H1000   ' save with high 4x1 chroma subsampling (4:1:1)
Public Const JPEG_SUBSAMPLING_420 As Long = &H4000   ' save with medium 2x2 medium chroma subsampling (4:2:0) - default value
Public Const JPEG_SUBSAMPLING_422 As Long = &H8000   ' save with low 2x1 chroma subsampling (4:2:2)
Public Const JPEG_SUBSAMPLING_444 As Long = &H10000  ' save with no chroma subsampling (4:4:4)
Public Const JPEG_OPTIMIZE As Long = &H20000         ' on saving, compute optimal Huffman coding tables (can reduce a few percent of file size)
Public Const JPEG_BASELINE As Long = &H40000         ' save basic JPEG, without metadata or any markers
Public Const KOALA_DEFAULT As Long = 0
Public Const LBM_DEFAULT As Long = 0
Public Const MNG_DEFAULT As Long = 0
Public Const PCD_DEFAULT As Long = 0
Public Const PCD_BASE As Long = 1                    ' load the bitmap sized 768 x 512
Public Const PCD_BASEDIV4 As Long = 2                ' load the bitmap sized 384 x 256
Public Const PCD_BASEDIV16 As Long = 3               ' load the bitmap sized 192 x 128
Public Const PCX_DEFAULT As Long = 0
Public Const PFM_DEFAULT As Long = 0
Public Const PICT_DEFAULT As Long = 0
Public Const PNG_DEFAULT As Long = 0
Public Const PNG_IGNOREGAMMA As Long = 1             ' avoid gamma correction
Public Const PNG_Z_BEST_SPEED As Long = &H1          ' save using ZLib level 1 compression flag (default value is 6)
Public Const PNG_Z_DEFAULT_COMPRESSION As Long = &H6 ' save using ZLib level 6 compression flag (default recommended value)
Public Const PNG_Z_BEST_COMPRESSION As Long = &H9    ' save using ZLib level 9 compression flag (default value is 6)
Public Const PNG_Z_NO_COMPRESSION As Long = &H100    ' save without ZLib compression
Public Const PNG_INTERLACED As Long = &H200          ' save using Adam7 interlacing (use | to combine with other save flags)
Public Const PNM_DEFAULT As Long = 0
Public Const PNM_SAVE_RAW As Long = 0                ' if set, the writer saves in RAW format (i.e. P4, P5 or P6)
Public Const PNM_SAVE_ASCII As Long = 1              ' if set, the writer saves in ASCII format (i.e. P1, P2 or P3)
Public Const PSD_DEFAULT As Long = 0
Public Const PSD_CMYK As Long = 1                    ' reads tags for separated CMYK (default is conversion to RGB)
Public Const PSD_LAB As Long = 2                     ' reads tags for CIELab (default is conversion to RGB)
Public Const RAS_DEFAULT As Long = 0
Public Const RAW_DEFAULT As Long = 0                 ' load the file as linear RGB 48-bit
Public Const RAW_PREVIEW As Long = 1                 ' try to load the embedded JPEG preview with included Exif Data or default to RGB 24-bit
Public Const RAW_DISPLAY As Long = 2                 ' load the file as RGB 24-bit
Public Const RAW_HALFSIZE As Long = 4                ' load the file as half-size color image
Public Const RAW_UNPROCESSED As Long = 8             ' load the file as FIT_UINT16 raw Bayer image
Public Const SGI_DEFAULT As Long = 0
Public Const TARGA_DEFAULT As Long = 0
Public Const TARGA_LOAD_RGB888 As Long = 1           ' if set, the loader converts RGB555 and ARGB8888 -> RGB888
Public Const TARGA_SAVE_RLE As Long = 2              ' if set, the writer saves with RLE compression
Public Const TIFF_DEFAULT As Long = 0
Public Const TIFF_CMYK As Long = &H1                 ' reads/stores tags for separated CMYK (use 'OR' to combine with compression flags)
Public Const TIFF_PACKBITS As Long = &H100           ' save using PACKBITS compression
Public Const TIFF_DEFLATE As Long = &H200            ' save using DEFLATE compression (a.k.a. ZLIB compression)
Public Const TIFF_ADOBE_DEFLATE As Long = &H400      ' save using ADOBE DEFLATE compression
Public Const TIFF_NONE As Long = &H800               ' save without any compression
Public Const TIFF_CCITTFAX3 As Long = &H1000         ' save using CCITT Group 3 fax encoding
Public Const TIFF_CCITTFAX4 As Long = &H2000         ' save using CCITT Group 4 fax encoding
Public Const TIFF_LZW As Long = &H4000               ' save using LZW compression
Public Const TIFF_JPEG As Long = &H8000              ' save using JPEG compression
Public Const TIFF_LOGLUV As Long = &H10000           ' save using LogLuv compression
Public Const WBMP_DEFAULT As Long = 0
Public Const XBM_DEFAULT As Long = 0
Public Const XPM_DEFAULT As Long = 0
Public Const WEBP_DEFAULT As Long = 0                ' save with good quality (75:1)
Public Const WEBP_LOSSLESS As Long = &H100           ' save in lossless mode
Public Const JXR_DEFAULT As Long = 0                 ' save with quality 80 and no chroma subsampling (4:4:4)
Public Const JXR_LOSSLESS As Long = &H64             ' save in lossless mode
Public Const JXR_PROGRESSIVE As Long = &H2000        ' save as a progressive-JXR (use Or to combine with other save flags)
' I/O image format identifiers
Public Enum FREE_IMAGE_FORMAT
   FIF_UNKNOWN = -1
   FIF_BMP = 0
   FIF_ICO = 1
   FIF_JPEG = 2
   FIF_JNG = 3
   FIF_KOALA = 4
   FIF_LBM = 5
   FIF_IFF = FIF_LBM
   FIF_MNG = 6
   FIF_PBM = 7
   FIF_PBMRAW = 8
   FIF_PCD = 9
   FIF_PCX = 10
   FIF_PGM = 11
   FIF_PGMRAW = 12
   FIF_PNG = 13
   FIF_PPM = 14
   FIF_PPMRAW = 15
   FIF_RAS = 16
   FIF_TARGA = 17
   FIF_TIFF = 18
   FIF_WBMP = 19
   FIF_PSD = 20
   FIF_CUT = 21
   FIF_XBM = 22
   FIF_XPM = 23
   FIF_DDS = 24
   FIF_GIF = 25
   FIF_HDR = 26
   FIF_FAXG3 = 27
   FIF_SGI = 28
   FIF_EXR = 29
   FIF_J2K = 30
   FIF_JP2 = 31
   FIF_PFM = 32
   FIF_PICT = 33
   FIF_RAW = 34
   FIF_WEBP = 35
   FIF_JXR = 36
End Enum
' Image load options
Public Enum FREE_IMAGE_LOAD_OPTIONS
   FILO_LOAD_NOPIXELS = FIF_LOAD_NOPIXELS         ' load the image header only (not supported by all plugins)
   FILO_LOAD_DEFAULT = 0
   FILO_GIF_DEFAULT = GIF_DEFAULT
   FILO_GIF_LOAD256 = GIF_LOAD256                 ' load the image as a 256 color image with ununsed palette entries, if it's 16 or 2 color
   FILO_GIF_PLAYBACK = GIF_PLAYBACK               ' 'play' the GIF to generate each frame (as 32bpp) instead of returning raw frame data when loading
   FILO_ICO_DEFAULT = ICO_DEFAULT
   FILO_ICO_MAKEALPHA = ICO_MAKEALPHA             ' convert to 32bpp and create an alpha channel from the AND-mask when loading
   FILO_JPEG_DEFAULT = JPEG_DEFAULT               ' for loading this is a synonym for FILO_JPEG_FAST
   FILO_JPEG_FAST = JPEG_FAST                     ' load the file as fast as possible, sacrificing some quality
   FILO_JPEG_ACCURATE = JPEG_ACCURATE             ' load the file with the best quality, sacrificing some speed
   FILO_JPEG_CMYK = JPEG_CMYK                     ' load separated CMYK "as is" (use 'OR' to combine with other load flags)
   FILO_JPEG_EXIFROTATE = JPEG_EXIFROTATE         ' load and rotate according to Exif 'Orientation' tag if available
   FILO_JPEG_GREYSCALE = JPEG_GREYSCALE           ' load and convert to a 8-bit greyscale image
   FILO_PCD_DEFAULT = PCD_DEFAULT
   FILO_PCD_BASE = PCD_BASE                       ' load the bitmap sized 768 x 512
   FILO_PCD_BASEDIV4 = PCD_BASEDIV4               ' load the bitmap sized 384 x 256
   FILO_PCD_BASEDIV16 = PCD_BASEDIV16             ' load the bitmap sized 192 x 128
   FILO_PNG_DEFAULT = PNG_DEFAULT
   FILO_PNG_IGNOREGAMMA = PNG_IGNOREGAMMA         ' avoid gamma correction
   FILO_PSD_CMYK = PSD_CMYK                       ' reads tags for separated CMYK (default is conversion to RGB)
   FILO_PSD_LAB = PSD_LAB                         ' reads tags for CIELab (default is conversion to RGB)
   FILO_RAW_DEFAULT = RAW_DEFAULT                 ' load the file as linear RGB 48-bit
   FILO_RAW_PREVIEW = RAW_PREVIEW                 ' try to load the embedded JPEG preview with included Exif Data or default to RGB 24-bit
   FILO_RAW_DISPLAY = RAW_DISPLAY                 ' load the file as RGB 24-bit
   FILO_RAW_HALFSIZE = RAW_HALFSIZE               ' load the file as half-size color image
   FILO_RAW_UNPROCESSED = RAW_UNPROCESSED         ' load the file as FIT_UINT16 raw Bayer image
   FILO_TARGA_DEFAULT = TARGA_LOAD_RGB888
   FILO_TARGA_LOAD_RGB888 = TARGA_LOAD_RGB888     ' if set, the loader converts RGB555 and ARGB8888 -> RGB888
   FISO_TIFF_DEFAULT = TIFF_DEFAULT
   FISO_TIFF_CMYK = TIFF_CMYK                     ' reads tags for separated CMYK
End Enum
' Image save options
Public Enum FREE_IMAGE_SAVE_OPTIONS
   FISO_SAVE_DEFAULT = 0
   FISO_BMP_DEFAULT = BMP_DEFAULT
   FISO_BMP_SAVE_RLE = BMP_SAVE_RLE
   FISO_EXR_DEFAULT = EXR_DEFAULT                 ' save data as half with piz-based wavelet compression
   FISO_EXR_FLOAT = EXR_FLOAT                     ' save data as float instead of as half (not recommended)
   FISO_EXR_NONE = EXR_NONE                       ' save with no compression
   FISO_EXR_ZIP = EXR_ZIP                         ' save with zlib compression, in blocks of 16 scan lines
   FISO_EXR_PIZ = EXR_PIZ                         ' save with piz-based wavelet compression
   FISO_EXR_PXR24 = EXR_PXR24                     ' save with lossy 24-bit float compression
   FISO_EXR_B44 = EXR_B44                         ' save with lossy 44% float compression - goes to 22% when combined with EXR_LC
   FISO_EXR_LC = EXR_LC                           ' save images with one luminance and two chroma channels, rather than as RGB (lossy compression)
   FISO_JPEG_DEFAULT = JPEG_DEFAULT               ' for saving this is a synonym for FISO_JPEG_QUALITYGOOD
   FISO_JPEG_QUALITYSUPERB = JPEG_QUALITYSUPERB   ' save with superb quality (100:1)
   FISO_JPEG_QUALITYGOOD = JPEG_QUALITYGOOD       ' save with good quality (75:1)
   FISO_JPEG_QUALITYNORMAL = JPEG_QUALITYNORMAL   ' save with normal quality (50:1)
   FISO_JPEG_QUALITYAVERAGE = JPEG_QUALITYAVERAGE ' save with average quality (25:1)
   FISO_JPEG_QUALITYBAD = JPEG_QUALITYBAD         ' save with bad quality (10:1)
   FISO_JPEG_PROGRESSIVE = JPEG_PROGRESSIVE       ' save as a progressive-JPEG (use 'OR' to combine with other save flags)
   FISO_JPEG_SUBSAMPLING_411 = JPEG_SUBSAMPLING_411      ' save with high 4x1 chroma subsampling (4:1:1)
   FISO_JPEG_SUBSAMPLING_420 = JPEG_SUBSAMPLING_420      ' save with medium 2x2 medium chroma subsampling (4:2:0) - default value
   FISO_JPEG_SUBSAMPLING_422 = JPEG_SUBSAMPLING_422      ' save with low 2x1 chroma subsampling (4:2:2)
   FISO_JPEG_SUBSAMPLING_444 = JPEG_SUBSAMPLING_444      ' save with no chroma subsampling (4:4:4)
   FISO_JPEG_OPTIMIZE = JPEG_OPTIMIZE                    ' compute optimal Huffman coding tables (can reduce a few percent of file size)
   FISO_JPEG_BASELINE = JPEG_BASELINE                    ' save basic JPEG, without metadata or any markers
   FISO_PNG_Z_BEST_SPEED = PNG_Z_BEST_SPEED              ' save using ZLib level 1 compression flag (default value is 6)
   FISO_PNG_Z_DEFAULT_COMPRESSION = PNG_Z_DEFAULT_COMPRESSION ' save using ZLib level 6 compression flag (default recommended value)
   FISO_PNG_Z_BEST_COMPRESSION = PNG_Z_BEST_COMPRESSION  ' save using ZLib level 9 compression flag (default value is 6)
   FISO_PNG_Z_NO_COMPRESSION = PNG_Z_NO_COMPRESSION      ' save without ZLib compression
   FISO_PNG_INTERLACED = PNG_INTERLACED           ' save using Adam7 interlacing (use | to combine with other save flags)
   FISO_PNM_DEFAULT = PNM_DEFAULT
   FISO_PNM_SAVE_RAW = PNM_SAVE_RAW               ' if set, the writer saves in RAW format (i.e. P4, P5 or P6)
   FISO_PNM_SAVE_ASCII = PNM_SAVE_ASCII           ' if set, the writer saves in ASCII format (i.e. P1, P2 or P3)
   FISO_TARGA_SAVE_RLE = TARGA_SAVE_RLE           ' if set, the writer saves with RLE compression
   FISO_TIFF_DEFAULT = TIFF_DEFAULT
   FISO_TIFF_CMYK = TIFF_CMYK                     ' stores tags for separated CMYK (use 'OR' to combine with compression flags)
   FISO_TIFF_PACKBITS = TIFF_PACKBITS             ' save using PACKBITS compression
   FISO_TIFF_DEFLATE = TIFF_DEFLATE               ' save using DEFLATE compression (a.k.a. ZLIB compression)
   FISO_TIFF_ADOBE_DEFLATE = TIFF_ADOBE_DEFLATE   ' save using ADOBE DEFLATE compression
   FISO_TIFF_NONE = TIFF_NONE                     ' save without any compression
   FISO_TIFF_CCITTFAX3 = TIFF_CCITTFAX3           ' save using CCITT Group 3 fax encoding
   FISO_TIFF_CCITTFAX4 = TIFF_CCITTFAX4           ' save using CCITT Group 4 fax encoding
   FISO_TIFF_LZW = TIFF_LZW                       ' save using LZW compression
   FISO_TIFF_JPEG = TIFF_JPEG                     ' save using JPEG compression
   FISO_TIFF_LOGLUV = TIFF_LOGLUV                 ' save using LogLuv compression
   FISO_WEBP_LOSSLESS = WEBP_LOSSLESS             ' save in lossless mode
   FISO_JXR_LOSSLESS = JXR_LOSSLESS               ' save in lossless mode
   FISO_JXR_PROGRESSIVE = JXR_PROGRESSIVE         ' save as a progressive-JXR (use Or to combine with other save flags)
End Enum
' Image types used in FreeImage
Public Enum FREE_IMAGE_TYPE
   FIT_UNKNOWN = 0           ' unknown type
   FIT_BITMAP = 1            ' standard image           : 1-, 4-, 8-, 16-, 24-, 32-bit
   FIT_UINT16 = 2            ' array of unsigned short  : unsigned 16-bit
   FIT_INT16 = 3             ' array of short           : signed 16-bit
   FIT_UINT32 = 4            ' array of unsigned long   : unsigned 32-bit
   FIT_INT32 = 5             ' array of long            : signed 32-bit
   FIT_FLOAT = 6             ' array of float           : 32-bit IEEE floating point
   FIT_DOUBLE = 7            ' array of double          : 64-bit IEEE floating point
   FIT_COMPLEX = 8           ' array of FICOMPLEX       : 2 x 64-bit IEEE floating point
   FIT_RGB16 = 9             ' 48-bit RGB image         : 3 x 16-bit
   FIT_RGBA16 = 10           ' 64-bit RGBA image        : 4 x 16-bit
   FIT_RGBF = 11             ' 96-bit RGB float image   : 3 x 32-bit IEEE floating point
   FIT_RGBAF = 12            ' 128-bit RGBA float image : 4 x 32-bit IEEE floating point
End Enum
' Image color types used in FreeImage
Public Enum FREE_IMAGE_COLOR_TYPE
   FIC_MINISWHITE = 0        ' min value is white
   FIC_MINISBLACK = 1        ' min value is black
   FIC_RGB = 2               ' RGB color model
   FIC_PALETTE = 3           ' color map indexed
   FIC_RGBALPHA = 4          ' RGB color model with alpha channel
   FIC_CMYK = 5              ' CMYK color model
End Enum
' Color quantization algorithm constants
Public Enum FREE_IMAGE_QUANTIZE
   FIQ_WUQUANT = 0           ' Xiaolin Wu color quantization algorithm
   FIQ_NNQUANT = 1           ' NeuQuant neural-net quantization algorithm by Anthony Dekker
   FIQ_LFPQUANT = 2          ' Lossless Fast Pseudo-Quantization Algorithm by Carsten Klein
End Enum
' Dithering algorithm constants
Public Enum FREE_IMAGE_DITHER
   FID_FS = 0                ' Floyd & Steinberg error diffusion
   FID_BAYER4x4 = 1          ' Bayer ordered dispersed dot dithering (order 2 dithering matrix)
   FID_BAYER8x8 = 2          ' Bayer ordered dispersed dot dithering (order 3 dithering matrix)
   FID_CLUSTER6x6 = 3        ' Ordered clustered dot dithering (order 3 - 6x6 matrix)
   FID_CLUSTER8x8 = 4        ' Ordered clustered dot dithering (order 4 - 8x8 matrix)
   FID_CLUSTER16x16 = 5      ' Ordered clustered dot dithering (order 8 - 16x16 matrix)
   FID_BAYER16x16 = 6        ' Bayer ordered dispersed dot dithering (order 4 dithering matrix)
End Enum
' Lossless JPEG transformation constants
Public Enum FREE_IMAGE_JPEG_OPERATION
   FIJPEG_OP_NONE = 0        ' no transformation
   FIJPEG_OP_FLIP_H = 1      ' horizontal flip
   FIJPEG_OP_FLIP_V = 2      ' vertical flip
   FIJPEG_OP_TRANSPOSE = 3   ' transpose across UL-to-LR axis
   FIJPEG_OP_TRANSVERSE = 4  ' transpose across UR-to-LL axis
   FIJPEG_OP_ROTATE_90 = 5   ' 90-degree clockwise rotation
   FIJPEG_OP_ROTATE_180 = 6  ' 180-degree rotation
   FIJPEG_OP_ROTATE_270 = 7  ' 270-degree clockwise (or 90 ccw)
End Enum
' Tone mapping operator constants
Public Enum FREE_IMAGE_TMO
   FITMO_DRAGO03 = 0         ' Adaptive logarithmic mapping (F. Drago, 2003)
   FITMO_REINHARD05 = 1      ' Dynamic range reduction inspired by photoreceptor physiology (E. Reinhard, 2005)
   FITMO_FATTAL02 = 2        ' Gradient domain high dynamic range compression (R. Fattal, 2002)
End Enum
' Up- / Downsampling filter constants
Public Enum FREE_IMAGE_FILTER
   FILTER_BOX = 0            ' Box, pulse, Fourier window, 1st order (constant) b-spline
   FILTER_BICUBIC = 1        ' Mitchell & Netravali's two-param cubic filter
   FILTER_BILINEAR = 2       ' Bilinear filter
   FILTER_BSPLINE = 3        ' 4th order (cubic) b-spline
   FILTER_CATMULLROM = 4     ' Catmull-Rom spline, Overhauser spline
   FILTER_LANCZOS3 = 5       ' Lanczos3 filter
End Enum
Public Enum FREE_IMAGE_RESCALE_OPTIONS
   FI_RESCALE_DEFAULT = &H0        ' default options; none of the following other options apply
   FI_RESCALE_TRUE_COLOR = &H1     ' for non-transparent greyscale images, convert to 24-bit if src bitdepth <= 8 (default is a 8-bit greyscale image)
   FI_RESCALE_OMIT_METADATA = &H2  ' do not copy metadata to the rescaled image
End Enum
' Color channel constants
Public Enum FREE_IMAGE_COLOR_CHANNEL
   FICC_RGB = 0              ' Use red, green and blue channels
   FICC_RED = 1              ' Use red channel
   FICC_GREEN = 2            ' Use green channel
   FICC_BLUE = 3             ' Use blue channel
   FICC_ALPHA = 4            ' Use alpha channel
   FICC_BLACK = 5            ' Use black channel
   FICC_REAL = 6             ' Complex images: use real part
   FICC_IMAG = 7             ' Complex images: use imaginary part
   FICC_MAG = 8              ' Complex images: use magnitude
   FICC_PHASE = 9            ' Complex images: use phase
End Enum
' Tag data type information constants (based on TIFF specifications)
Public Enum FREE_IMAGE_MDTYPE
   FIDT_NOTYPE = 0           ' placeholder
   FIDT_BYTE = 1             ' 8-bit unsigned integer
   FIDT_ASCII = 2            ' 8-bit bytes w/ last byte null
   FIDT_SHORT = 3            ' 16-bit unsigned integer
   FIDT_LONG = 4             ' 32-bit unsigned integer
   FIDT_RATIONAL = 5         ' 64-bit unsigned fraction
   FIDT_SBYTE = 6            ' 8-bit signed integer
   FIDT_UNDEFINED = 7        ' 8-bit untyped data
   FIDT_SSHORT = 8           ' 16-bit signed integer
   FIDT_SLONG = 9            ' 32-bit signed integer
   FIDT_SRATIONAL = 10       ' 64-bit signed fraction
   FIDT_FLOAT = 11           ' 32-bit IEEE floating point
   FIDT_DOUBLE = 12          ' 64-bit IEEE floating point
   FIDT_IFD = 13             ' 32-bit unsigned integer (offset)
   FIDT_PALETTE = 14         ' 32-bit RGBQUAD
End Enum
' Metadata models supported by FreeImage
Public Enum FREE_IMAGE_MDMODEL
   FIMD_NODATA = -1          '
   FIMD_COMMENTS = 0         ' single comment or keywords
   FIMD_EXIF_MAIN = 1        ' Exif-TIFF metadata
   FIMD_EXIF_EXIF = 2        ' Exif-specific metadata
   FIMD_EXIF_GPS = 3         ' Exif GPS metadata
   FIMD_EXIF_MAKERNOTE = 4   ' Exif maker note metadata
   FIMD_EXIF_INTEROP = 5     ' Exif interoperability metadata
   FIMD_IPTC = 6             ' IPTC/NAA metadata
   FIMD_XMP = 7              ' Abobe XMP metadata
   FIMD_GEOTIFF = 8          ' GeoTIFF metadata
   FIMD_ANIMATION = 9        ' Animation metadata
   FIMD_CUSTOM = 10          ' Used to attach other metadata types to a dib
   FIMD_EXIF_RAW = 11        ' Exif metadata as a raw buffer
End Enum
' These are the GIF_DISPOSAL metadata constants
Public Enum FREE_IMAGE_FRAME_DISPOSAL_METHODS
   FIFD_GIF_DISPOSAL_UNSPECIFIED = 0
   FIFD_GIF_DISPOSAL_LEAVE = 1
   FIFD_GIF_DISPOSAL_BACKGROUND = 2
   FIFD_GIF_DISPOSAL_PREVIOUS = 3
End Enum
' Constants used in FreeImage_FillBackground and FreeImage_EnlargeCanvas
Public Enum FREE_IMAGE_COLOR_OPTIONS
   FI_COLOR_IS_RGB_COLOR = &H0          ' RGBQUAD color is a RGB color (contains no valid alpha channel)
   FI_COLOR_IS_RGBA_COLOR = &H1         ' RGBQUAD color is a RGBA color (contains a valid alpha channel)
   FI_COLOR_FIND_EQUAL_COLOR = &H2      ' For palettized images: lookup equal RGB color from palette
   FI_COLOR_ALPHA_IS_INDEX = &H4        ' The color's rgbReserved member (alpha) contains the palette index to be used
End Enum
Public Const FI_COLOR_PALETTE_SEARCH_MASK = (FI_COLOR_FIND_EQUAL_COLOR Or FI_COLOR_ALPHA_IS_INDEX)     ' Flag to test, if any color lookup is performed
' The following enum constants are used by derived (wrapper) functions of the
' FreeImage 3 VB Wrapper
Public Enum FREE_IMAGE_CONVERSION_FLAGS
   FICF_MONOCHROME = &H1
   FICF_MONOCHROME_THRESHOLD = FICF_MONOCHROME
   FICF_MONOCHROME_DITHER = &H3
   FICF_GREYSCALE_4BPP = &H4
   FICF_PALLETISED_8BPP = &H8
   FICF_GREYSCALE_8BPP = FICF_PALLETISED_8BPP Or FICF_MONOCHROME
   FICF_GREYSCALE = FICF_GREYSCALE_8BPP
   FICF_RGB_15BPP = &HF
   FICF_RGB_16BPP = &H10
   FICF_RGB_24BPP = &H18
   FICF_RGB_32BPP = &H20
   FICF_RGB_ALPHA = FICF_RGB_32BPP
   FICF_KEEP_UNORDERED_GREYSCALE_PALETTE = &H0
   FICF_REORDER_GREYSCALE_PALETTE = &H1000
End Enum
Public Enum FREE_IMAGE_COLOR_DEPTH
   FICD_AUTO = &H0
   FICD_MONOCHROME = &H1
   FICD_MONOCHROME_THRESHOLD = FICF_MONOCHROME
   FICD_MONOCHROME_DITHER = &H3
   FICD_1BPP = FICD_MONOCHROME
   FICD_4BPP = &H4
   FICD_8BPP = &H8
   FICD_15BPP = &HF
   FICD_16BPP = &H10
   FICD_24BPP = &H18
   FICD_32BPP = &H20
End Enum
Public Enum FREE_IMAGE_ADJUST_MODE
   AM_STRECH = &H1
   AM_DEFAULT = AM_STRECH
   AM_ADJUST_BOTH = AM_STRECH
   AM_ADJUST_WIDTH = &H2
   AM_ADJUST_HEIGHT = &H4
   AM_ADJUST_OPTIMAL_SIZE = &H8
End Enum
Public Enum FREE_IMAGE_MASK_FLAGS
   FIMF_MASK_NONE = &H0
   FIMF_MASK_FULL_TRANSPARENCY = &H1
   FIMF_MASK_ALPHA_TRANSPARENCY = &H2
   FIMF_MASK_COLOR_TRANSPARENCY = &H4
   FIMF_MASK_FORCE_TRANSPARENCY = &H8
   FIMF_MASK_INVERSE_MASK = &H10
End Enum
Public Enum FREE_IMAGE_COLOR_FORMAT_FLAGS
   FICFF_COLOR_RGB = &H1
   FICFF_COLOR_BGR = &H2
   FICFF_COLOR_PALETTE_INDEX = &H4
   FICFF_COLOR_HAS_ALPHA = &H100
   FICFF_COLOR_ARGB = FICFF_COLOR_RGB Or FICFF_COLOR_HAS_ALPHA
   FICFF_COLOR_ABGR = FICFF_COLOR_BGR Or FICFF_COLOR_HAS_ALPHA
   FICFF_COLOR_FORMAT_ORDER_MASK = FICFF_COLOR_RGB Or FICFF_COLOR_BGR
End Enum
Public Enum FREE_IMAGE_MASK_CREATION_OPTION_FLAGS
   MCOF_CREATE_MASK_IMAGE = &H1
   MCOF_MODIFY_SOURCE_IMAGE = &H2
   MCOF_CREATE_AND_MODIFY = MCOF_CREATE_MASK_IMAGE Or MCOF_MODIFY_SOURCE_IMAGE
End Enum
Public Enum FREE_IMAGE_TRANSPARENCY_STATE_FLAGS
   FITSF_IGNORE_TRANSPARENCY = &H0
   FITSF_NONTRANSPARENT = &H1
   FITSF_TRANSPARENT = &H2
   FITSF_INCLUDE_ALPHA_TRANSPARENCY = &H4
End Enum
Public Enum FREE_IMAGE_ICON_TRANSPARENCY_OPTION_FLAGS
   ITOF_NO_TRANSPARENCY = &H0
   ITOF_USE_TRANSPARENCY_INFO = &H1
   ITOF_USE_TRANSPARENCY_INFO_ONLY = ITOF_USE_TRANSPARENCY_INFO
   ITOF_USE_COLOR_TRANSPARENCY = &H2
   ITOF_USE_COLOR_TRANSPARENCY_ONLY = ITOF_USE_COLOR_TRANSPARENCY
   ITOF_USE_TRANSPARENCY_INFO_OR_COLOR = ITOF_USE_TRANSPARENCY_INFO Or ITOF_USE_COLOR_TRANSPARENCY
   ITOF_USE_DEFAULT_TRANSPARENCY = ITOF_USE_TRANSPARENCY_INFO_OR_COLOR
   ITOF_USE_COLOR_TOP_LEFT_PIXEL = &H0
   ITOF_USE_COLOR_FIRST_PIXEL = ITOF_USE_COLOR_TOP_LEFT_PIXEL
   ITOF_USE_COLOR_TOP_RIGHT_PIXEL = &H20
   ITOF_USE_COLOR_BOTTOM_LEFT_PIXEL = &H40
   ITOF_USE_COLOR_BOTTOM_RIGHT_PIXEL = &H80
   ITOF_USE_COLOR_SPECIFIED = &H100
   ITOF_FORCE_TRANSPARENCY_INFO = &H400
End Enum
Private Const ITOF_USE_COLOR_BITMASK As Long = ITOF_USE_COLOR_TOP_RIGHT_PIXEL Or ITOF_USE_COLOR_BOTTOM_LEFT_PIXEL Or ITOF_USE_COLOR_BOTTOM_RIGHT_PIXEL Or ITOF_USE_COLOR_SPECIFIED
Public Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type
Public Type RGBTRIPLE
   rgbtBlue As Byte
   rgbtGreen As Byte
   rgbtRed As Byte
End Type
Public Const BITMAPFILEHEADERSIZE = &HE&
Public Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffset As Long
End Type
Public Const BITMAPINFOHEADERSIZE = &H28&
Public Type BITMAPINFOHEADER
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
Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD 'bmiColors(0) As RGBQUAD
End Type
Private Type DIBSECTION
    dsBm As BITMAP_API
    dsBmih As BITMAPINFOHEADER
    dsBitFields(0 To 2) As Long
    dshSection As Long ' < ???
    dsOffset As Long
End Type

Public Const BI_RGB As Long = 0
Public Const BI_RLE8 As Long = 1
Public Const BI_RLE4 As Long = 2
Public Const BI_BITFIELDS As Long = 3
Public Const BI_JPEG As Long = 4
Public Const BI_PNG As Long = 5
#If VBA7 Then           '<OFFICE2010+>
Public Type FIICCPROFILE
   Flags As Integer         ' info flag
   Size As Long             ' profile's size measured in bytes
   Data As LongPtr          ' points to a block of contiguous memory containing the profile
End Type
#Else                       ' <OFFICE97-2007>
Public Type FIICCPROFILE
   Flags As Integer         ' info flag
   Size As Long             ' profile's size measured in bytes
   Data As Long             ' points to a block of contiguous memory containing the profile
End Type
#End If                     ' <VBA7>
' 48-bit RGB
Public Type FIRGB16
   Red As Integer
   Green As Integer
   Blue As Integer
End Type
' 64-bit RGBA
Public Type FIRGBA16
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type
' 96-bit RGB Float
Public Type FIRGBF
   Red As Single
   Green As Single
   Blue As Single
End Type
' 128-bit RGBA Float
Public Type FIRGBAF
   Red As Single
   Green As Single
   Blue As Single
   Alpha As Single
End Type
' data structure for COMPLEX type (complex number)
Public Type FICOMPLEX
   r As Double           ' real part
   i As Double           ' imaginary part
End Type
#If VBA7 Then           '<OFFICE2010+>
Public Type FITAG
   Key As LongPtr
   Description As LongPtr
   ID As Integer
   Type As Integer
   Count As Long
   Length As Long
   Value As LongPtr
End Type
#Else                       ' <OFFICE97-2007>
Public Type FITAG
   Key As Long
   Description As Long
   ID As Integer
   Type As Integer
   Count As Long
   Length As Long
   Value As Long
End Type
#End If                     ' <VBA7>

Public Type FIRATIONAL
   Numerator As Variant
   Denominator As Variant
End Type
#If VBA7 Then           '<OFFICE2010+>
Public Type FREE_IMAGE_TAG
   Model As FREE_IMAGE_MDMODEL
   TagPtr As LongPtr
   Key As String
   Description As String
   ID As Long
   Type As FREE_IMAGE_MDTYPE
   Count As Long
   Length As Long
   StringValue As String
   palette() As RGBQUAD
   RationalValue() As FIRATIONAL
   Value As Variant
End Type
Public Type FreeImageIO
   read_proc As LongPtr
   write_proc As LongPtr
   seek_proc As LongPtr
   tell_proc As LongPtr
End Type
Public Type Plugin
   format_proc As LongPtr
   description_proc As LongPtr
   extension_proc As LongPtr
   regexpr_proc As LongPtr
   open_proc As LongPtr
   close_proc As LongPtr
   pagecount_proc As LongPtr
   pagecapability_proc As LongPtr
   load_proc As LongPtr
   save_proc As LongPtr
   validate_proc As LongPtr
   mime_proc As LongPtr
   supports_export_bpp_proc As LongPtr
   supports_export_type_proc As LongPtr
   supports_icc_profiles_proc As LongPtr
End Type
#Else                       ' <OFFICE97-2007>
Public Type FREE_IMAGE_TAG
   Model As FREE_IMAGE_MDMODEL
   TagPtr As Long
   Key As String
   Description As String
   ID As Long
   Type As FREE_IMAGE_MDTYPE
   Count As Long
   Length As Long
   StringValue As String
   palette() As RGBQUAD
   RationalValue() As FIRATIONAL
   Value As Variant
End Type
Public Type FreeImageIO
   read_proc As Long
   write_proc As Long
   seek_proc As Long
   tell_proc As Long
End Type
Public Type Plugin
   format_proc As Long
   description_proc As Long
   extension_proc As Long
   regexpr_proc As Long
   open_proc As Long
   close_proc As Long
   pagecount_proc As Long
   pagecapability_proc As Long
   load_proc As Long
   save_proc As Long
   validate_proc As Long
   mime_proc As Long
   supports_export_bpp_proc As Long
   supports_export_type_proc As Long
   supports_icc_profiles_proc As Long
End Type
#End If                     ' <VBA7>

' The following structures are used by derived (wrapper) functions of the
' FreeImage 3 VB Wrapper
Public Type ScanLineRGBTRIBLE
   Data() As RGBTRIPLE
End Type
Public Type ScanLinesRGBTRIBLE
   Scanline() As ScanLineRGBTRIBLE
End Type
'----------------------
' FreeImage 3 function declarations
'----------------------
' The FreeImage 3 functions are declared in the same order as they are described
' in the FreeImage 3 API documentation (mostly). The documentation's outline is
' included as comments.
#If VBA7 Then           '<OFFICE2010+>
' Initialization / Deinitialization functions
Public Declare PtrSafe Sub FreeImage_Initialise Lib "FreeImage_x64.dll" (Optional ByVal LoadLocalPluginsOnly As Long)
Public Declare PtrSafe Sub FreeImage_DeInitialise Lib "FreeImage_x64.dll" ()
' Version functions
Private Declare PtrSafe Function p_FreeImage_GetVersion Lib "FreeImage_x64.dll" Alias "FreeImage_GetVersion" () As LongPtr
Private Declare PtrSafe Function p_FreeImage_GetCopyrightMessage Lib "FreeImage_x64.dll" Alias "FreeImage_GetCopyrightMessage" () As LongPtr
' Message output functions
Public Declare PtrSafe Sub FreeImage_SetOutputMessage Lib "FreeImage_x64.dll" Alias "FreeImage_SetOutputMessageStdCall" (ByVal omf As LongPtr)
' Allocate / Clone / Unload functions
Public Declare PtrSafe Function FreeImage_Allocate Lib "FreeImage_x64.dll" (ByVal Width As Long, ByVal Height As Long, ByVal BitsPerPixel As Long, Optional ByVal RedMask As Long = 0&, Optional ByVal GreenMask As Long = 0&, Optional ByVal BlueMask As Long = 0&) As LongPtr
Public Declare PtrSafe Function FreeImage_AllocateT Lib "FreeImage_x64.dll" (ByVal ImageType As FREE_IMAGE_TYPE, ByVal Width As Long, ByVal Height As Long, Optional ByVal BitsPerPixel As Long = 8, Optional ByVal RedMask As Long = 0&, Optional ByVal GreenMask As Long = 0&, Optional ByVal BlueMask As Long = 0&) As LongPtr
Public Declare PtrSafe Function FreeImage_Clone Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Sub FreeImage_Unload Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr)
' Header loading functions
Public Declare PtrSafe Function FreeImage_HasPixelsLng Lib "FreeImage_x64.dll" Alias "FreeImage_HasPixels" (ByVal BITMAP As LongPtr) As Long
' Load / Save functions
Public Declare PtrSafe Function FreeImage_Load Lib "FreeImage_x64.dll" (ByVal Format As FREE_IMAGE_FORMAT, ByVal FileName As String, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As LongPtr
Public Declare PtrSafe Function FreeImage_LoadFromHandle Lib "FreeImage_x64.dll" (ByVal Format As FREE_IMAGE_FORMAT, ByVal IO As LongPtr, ByVal Handle As LongPtr, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As LongPtr
Private Declare PtrSafe Function p_FreeImage_LoadU Lib "FreeImage_x64.dll" Alias "FreeImage_LoadU" (ByVal Format As FREE_IMAGE_FORMAT, ByVal FileName As LongPtr, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As LongPtr
Private Declare PtrSafe Function p_FreeImage_Save Lib "FreeImage_x64.dll" Alias "FreeImage_Save" (ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As LongPtr, ByVal FileName As String, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long
Private Declare PtrSafe Function p_FreeImage_SaveU Lib "FreeImage_x64.dll" Alias "FreeImage_SaveU" (ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As LongPtr, ByVal FileName As LongPtr, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long
Private Declare PtrSafe Function p_FreeImage_SaveToHandle Lib "FreeImage_x64.dll" Alias "FreeImage_SaveToHandle" (ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As LongPtr, ByVal IO As LongPtr, ByVal Handle As LongPtr, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long
' Memory I/O stream functions
Public Declare PtrSafe Function FreeImage_OpenMemory Lib "FreeImage_x64.dll" (Optional ByRef Data As Byte, Optional ByVal SizeInBytes As Long) As LongPtr
Public Declare PtrSafe Function FreeImage_OpenMemoryByPtr Lib "FreeImage_x64.dll" Alias "FreeImage_OpenMemory" (Optional ByVal DataPtr As LongPtr, Optional ByVal SizeInBytes As Long) As LongPtr
Public Declare PtrSafe Sub FreeImage_CloseMemory Lib "FreeImage_x64.dll" (ByVal Stream As LongPtr)
Public Declare PtrSafe Function FreeImage_LoadFromMemory Lib "FreeImage_x64.dll" (ByVal Format As FREE_IMAGE_FORMAT, ByVal Stream As LongPtr, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As LongPtr
Public Declare PtrSafe Function FreeImage_TellMemory Lib "FreeImage_x64.dll" (ByVal Stream As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_ReadMemory Lib "FreeImage_x64.dll" (ByVal BufferPtr As LongPtr, ByVal Size As Long, ByVal Count As Long, ByVal Stream As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_WriteMemory Lib "FreeImage_x64.dll" (ByVal BufferPtr As LongPtr, ByVal Size As Long, ByVal Count As Long, ByVal Stream As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_LoadMultiBitmapFromMemory Lib "FreeImage_x64.dll" (ByVal Format As FREE_IMAGE_FORMAT, ByVal Stream As LongPtr, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As LongPtr
Public Declare PtrSafe Function FreeImage_SaveMultiBitmapToMemory Lib "FreeImage_x64.dll" (ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As LongPtr, ByVal Stream As LongPtr, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long
Private Declare PtrSafe Function p_FreeImage_SaveToMemory Lib "FreeImage_x64.dll" Alias "FreeImage_SaveToMemory" (ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As LongPtr, ByVal Stream As LongPtr, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long
Private Declare PtrSafe Function p_FreeImage_SeekMemory Lib "FreeImage_x64.dll" Alias "FreeImage_SeekMemory" (ByVal Stream As LongPtr, ByVal Offset As Long, ByVal Origin As Long) As Long
Private Declare PtrSafe Function p_FreeImage_AcquireMemory Lib "FreeImage_x64.dll" Alias "FreeImage_AcquireMemory" (ByVal Stream As LongPtr, ByRef DataPtr As LongPtr, ByRef SizeInBytes As Long) As Long
' Plugin / Format functions
Public Declare PtrSafe Function FreeImage_RegisterLocalPlugin Lib "FreeImage_x64.dll" (ByVal InitProcAddress As LongPtr, Optional ByVal Format As String, Optional ByVal Description As String, Optional ByVal Extension As String, Optional ByVal RegExpr As String) As FREE_IMAGE_FORMAT
Public Declare PtrSafe Function FreeImage_RegisterExternalPlugin Lib "FreeImage_x64.dll" (ByVal path As String, Optional ByVal Format As String, Optional ByVal Description As String, Optional ByVal Extension As String, Optional ByVal RegExpr As String) As FREE_IMAGE_FORMAT
Public Declare PtrSafe Function FreeImage_GetFIFCount Lib "FreeImage_x64.dll" () As Long
Public Declare PtrSafe Function FreeImage_SetPluginEnabled Lib "FreeImage_x64.dll" (ByVal Format As FREE_IMAGE_FORMAT, ByVal Value As Long) As Long
Public Declare PtrSafe Function FreeImage_IsPluginEnabled Lib "FreeImage_x64.dll" (ByVal Format As FREE_IMAGE_FORMAT) As Long
Public Declare PtrSafe Function FreeImage_GetFIFFromFormat Lib "FreeImage_x64.dll" (ByVal Format As String) As FREE_IMAGE_FORMAT
Public Declare PtrSafe Function FreeImage_GetFIFFromMime Lib "FreeImage_x64.dll" (ByVal MimeType As String) As FREE_IMAGE_FORMAT
Public Declare PtrSafe Function FreeImage_GetFIFFromFilename Lib "FreeImage_x64.dll" (ByVal FileName As String) As FREE_IMAGE_FORMAT
Private Declare PtrSafe Function p_FreeImage_GetFormatFromFIF Lib "FreeImage_x64.dll" Alias "FreeImage_GetFormatFromFIF" (ByVal Format As FREE_IMAGE_FORMAT) As LongPtr
Private Declare PtrSafe Function p_FreeImage_GetFIFExtensionList Lib "FreeImage_x64.dll" Alias "FreeImage_GetFIFExtensionList" (ByVal Format As FREE_IMAGE_FORMAT) As LongPtr
Private Declare PtrSafe Function p_FreeImage_GetFIFDescription Lib "FreeImage_x64.dll" Alias "FreeImage_GetFIFDescription" (ByVal Format As FREE_IMAGE_FORMAT) As LongPtr
Private Declare PtrSafe Function p_FreeImage_GetFIFRegExpr Lib "FreeImage_x64.dll" Alias "FreeImage_GetFIFRegExpr" (ByVal Format As FREE_IMAGE_FORMAT) As LongPtr
Private Declare PtrSafe Function p_FreeImage_GetFIFMimeType Lib "FreeImage_x64.dll" Alias "FreeImage_GetFIFMimeType" (ByVal Format As FREE_IMAGE_FORMAT) As LongPtr
Private Declare PtrSafe Function p_FreeImage_GetFIFFromFilenameU Lib "FreeImage_x64.dll" Alias "FreeImage_GetFIFFromFilenameU" (ByVal FileName As LongPtr) As FREE_IMAGE_FORMAT
Private Declare PtrSafe Function p_FreeImage_FIFSupportsReading Lib "FreeImage_x64.dll" Alias "FreeImage_FIFSupportsReading" (ByVal Format As FREE_IMAGE_FORMAT) As Long
Private Declare PtrSafe Function p_FreeImage_FIFSupportsWriting Lib "FreeImage_x64.dll" Alias "FreeImage_FIFSupportsWriting" (ByVal Format As FREE_IMAGE_FORMAT) As Long
Private Declare PtrSafe Function p_FreeImage_FIFSupportsExportBPP Lib "FreeImage_x64.dll" Alias "FreeImage_FIFSupportsExportBPP" (ByVal Format As FREE_IMAGE_FORMAT, ByVal BitsPerPixel As Long) As Long
Private Declare PtrSafe Function p_FreeImage_FIFSupportsExportType Lib "FreeImage_x64.dll" Alias "FreeImage_FIFSupportsExportType" (ByVal Format As FREE_IMAGE_FORMAT, ByVal ImageType As FREE_IMAGE_TYPE) As Long
Private Declare PtrSafe Function p_FreeImage_FIFSupportsICCProfiles Lib "FreeImage_x64.dll" Alias "FreeImage_FIFSupportsICCProfiles" (ByVal Format As FREE_IMAGE_FORMAT) As Long
Private Declare PtrSafe Function p_FreeImage_FIFSupportsNoPixels Lib "FreeImage_x64.dll" Alias "FreeImage_FIFSupportsNoPixels" (ByVal Format As FREE_IMAGE_FORMAT) As Long
' Multipaging functions
Public Declare PtrSafe Function FreeImage_GetPageCount Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Public Declare PtrSafe Sub FreeImage_AppendPage Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal PageBitmap As LongPtr)
Public Declare PtrSafe Sub FreeImage_InsertPage Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal Page As Long, ByVal PageBitmap As LongPtr)
Public Declare PtrSafe Sub FreeImage_DeletePage Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal Page As Long)
Public Declare PtrSafe Function FreeImage_LockPage Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal Page As Long) As LongPtr
Private Declare PtrSafe Function p_FreeImage_OpenMultiBitmap Lib "FreeImage_x64.dll" Alias "FreeImage_OpenMultiBitmap" (ByVal Format As FREE_IMAGE_FORMAT, ByVal FileName As String, ByVal CreateNew As Long, ByVal ReadOnly As Long, ByVal KeepCacheInMemory As Long, ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As LongPtr
Private Declare PtrSafe Function p_FreeImage_CloseMultiBitmap Lib "FreeImage_x64.dll" Alias "FreeImage_CloseMultiBitmap" (ByVal BITMAP As LongPtr, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long
Private Declare PtrSafe Function p_FreeImage_MovePage Lib "FreeImage_x64.dll" Alias "FreeImage_MovePage" (ByVal BITMAP As LongPtr, ByVal TargetPage As Long, ByVal SourcePage As Long) As Long
Private Declare PtrSafe Function p_FreeImage_GetLockedPageNumbers Lib "FreeImage_x64.dll" Alias "FreeImage_GetLockedPageNumbers" (ByVal BITMAP As LongPtr, ByRef PagesPtr As LongPtr, ByRef Count As Long) As Long
Private Declare PtrSafe Sub p_FreeImage_UnlockPage Lib "FreeImage_x64.dll" Alias "FreeImage_UnlockPage" (ByVal BITMAP As LongPtr, ByVal PageBitmap As LongPtr, ByVal ApplyChanges As Long)
' Filetype request functions
Public Declare PtrSafe Function FreeImage_GetFileType Lib "FreeImage_x64.dll" (ByVal FileName As String, Optional ByVal Size As Long) As FREE_IMAGE_FORMAT
Public Declare PtrSafe Function FreeImage_GetFileTypeFromHandle Lib "FreeImage_x64.dll" (ByVal IO As LongPtr, ByVal Handle As LongPtr, Optional ByVal Size As Long) As FREE_IMAGE_FORMAT
Public Declare PtrSafe Function FreeImage_GetFileTypeFromMemory Lib "FreeImage_x64.dll" (ByVal Stream As LongPtr, Optional ByVal Size As Long) As FREE_IMAGE_FORMAT
Private Declare PtrSafe Function p_FreeImage_GetFileTypeU Lib "FreeImage_x64.dll" Alias "FreeImage_GetFileTypeU" (ByVal FileName As LongPtr, Optional ByVal Size As Long) As FREE_IMAGE_FORMAT
' Image type request functions
Public Declare PtrSafe Function FreeImage_GetImageType Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As FREE_IMAGE_TYPE
' FreeImage helper functions
Private Declare PtrSafe Function p_FreeImage_IsLittleEndian Lib "FreeImage_x64.dll" Alias "FreeImage_IsLittleEndian" () As Long
Private Declare PtrSafe Function p_FreeImage_LookupX11Color Lib "FreeImage_x64.dll" Alias "FreeImage_LookupX11Color" (ByVal Color As String, ByRef Red As Long, ByRef Green As Long, ByRef Blue As Long) As Long
Private Declare PtrSafe Function p_FreeImage_LookupSVGColor Lib "FreeImage_x64.dll" Alias "FreeImage_LookupSVGColor" (ByVal Color As String, ByRef Red As Long, ByRef Green As Long, ByRef Blue As Long) As Long
' Pixel access functions
Public Declare PtrSafe Function FreeImage_GetBits Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_GetScanline Lib "FreeImage_x64.dll" Alias "FreeImage_GetScanLine" (ByVal BITMAP As LongPtr, ByVal Scanline As Long) As LongPtr
Public Declare PtrSafe Function FreeImage_GetColorsUsed Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_GetBPP Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Private Declare PtrSafe Function p_FreeImage_GetPixelIndex Lib "FreeImage_x64.dll" Alias "FreeImage_GetPixelIndex" (ByVal BITMAP As LongPtr, ByVal x As Long, ByVal y As Long, ByRef Value As Byte) As Long
Private Declare PtrSafe Function p_FreeImage_GetPixelColor Lib "FreeImage_x64.dll" Alias "FreeImage_GetPixelColor" (ByVal BITMAP As LongPtr, ByVal x As Long, ByVal y As Long, ByRef Value As RGBQUAD) As Long
Private Declare PtrSafe Function p_FreeImage_GetPixelColorByLong Lib "FreeImage_x64.dll" Alias "FreeImage_GetPixelColor" (ByVal BITMAP As LongPtr, ByVal x As Long, ByVal y As Long, ByRef Value As Long) As Long
Private Declare PtrSafe Function p_FreeImage_SetPixelIndex Lib "FreeImage_x64.dll" Alias "FreeImage_SetPixelIndex" (ByVal BITMAP As LongPtr, ByVal x As Long, ByVal y As Long, ByRef Value As Byte) As Long
Private Declare PtrSafe Function p_FreeImage_SetPixelColor Lib "FreeImage_x64.dll" Alias "FreeImage_SetPixelColor" (ByVal BITMAP As LongPtr, ByVal x As Long, ByVal y As Long, ByRef Value As RGBQUAD) As Long
Private Declare PtrSafe Function p_FreeImage_SetPixelColorByLong Lib "FreeImage_x64.dll" Alias "FreeImage_SetPixelColor" (ByVal BITMAP As LongPtr, ByVal x As Long, ByVal y As Long, ByRef Value As Long) As Long
' DIB info functions
Public Declare PtrSafe Function FreeImage_GetWidth Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_GetHeight Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_GetLine Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_GetPitch Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_GetDIBSize Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_GetMemorySize Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_GetPalette Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_GetDotsPerMeterX Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_GetDotsPerMeterY Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Public Declare PtrSafe Sub FreeImage_SetDotsPerMeterX Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal resolution As Long)
Public Declare PtrSafe Sub FreeImage_SetDotsPerMeterY Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal resolution As Long)
Public Declare PtrSafe Function FreeImage_GetInfoHeader Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_GetInfo Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_GetColorType Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As FREE_IMAGE_COLOR_TYPE
Public Declare PtrSafe Function FreeImage_GetRedMask Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_GetGreenMask Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_GetBlueMask Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_GetTransparencyCount Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_GetTransparencyTable Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Sub FreeImage_SetTransparencyTable Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal TransTablePtr As LongPtr, ByVal Count As Long)
Public Declare PtrSafe Function FreeImage_SetTransparentIndex Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal Index As Long) As Long
Public Declare PtrSafe Function FreeImage_GetTransparentIndex Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Public Declare PtrSafe Function FreeImage_GetThumbnail Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As Long
Private Declare PtrSafe Function p_FreeImage_HasBackgroundColor Lib "FreeImage_x64.dll" Alias "FreeImage_HasBackgroundColor" (ByVal BITMAP As LongPtr) As Long
Private Declare PtrSafe Function p_FreeImage_GetBackgroundColor Lib "FreeImage_x64.dll" Alias "FreeImage_GetBackgroundColor" (ByVal BITMAP As LongPtr, ByRef BackColor As RGBQUAD) As Long
Private Declare PtrSafe Function p_FreeImage_GetBackgroundColorAsLong Lib "FreeImage_x64.dll" Alias "FreeImage_GetBackgroundColor" (ByVal BITMAP As LongPtr, ByRef BackColor As Long) As Long
Private Declare PtrSafe Function p_FreeImage_SetBackgroundColor Lib "FreeImage_x64.dll" Alias "FreeImage_SetBackgroundColor" (ByVal BITMAP As LongPtr, ByRef BackColor As RGBQUAD) As Long
Private Declare PtrSafe Function p_FreeImage_SetBackgroundColorAsLong Lib "FreeImage_x64.dll" Alias "FreeImage_SetBackgroundColor" (ByVal BITMAP As LongPtr, ByRef BackColor As Long) As Long
Private Declare PtrSafe Function p_FreeImage_HasRGBMasks Lib "FreeImage_x64.dll" Alias "FreeImage_HasRGBMasks" (ByVal BITMAP As LongPtr) As Long
Private Declare PtrSafe Function p_FreeImage_IsTransparent Lib "FreeImage_x64.dll" Alias "FreeImage_IsTransparent" (ByVal BITMAP As LongPtr) As Long
Private Declare PtrSafe Function p_FreeImage_SetThumbnail Lib "FreeImage_x64.dll" Alias "FreeImage_SetThumbnail" (ByVal BITMAP As LongPtr, ByVal Thumbnail As LongPtr) As Long
Private Declare PtrSafe Sub p_FreeImage_SetTransparent Lib "FreeImage_x64.dll" Alias "FreeImage_SetTransparent" (ByVal BITMAP As LongPtr, ByVal Value As Long)
' ICC profile functions
Public Declare PtrSafe Function FreeImage_CreateICCProfile Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByRef Data As LongPtr, ByVal Size As Long) As LongPtr
Public Declare PtrSafe Sub FreeImage_DestroyICCProfile Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr)
Private Declare PtrSafe Function p_FreeImage_GetICCProfile Lib "FreeImage_x64.dll" Alias "FreeImage_GetICCProfile" (ByVal BITMAP As LongPtr) As LongPtr
' Line conversion functions
Public Declare PtrSafe Sub FreeImage_ConvertLine1To4 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine8To4 Lib "FreeImage_x64.dll" Alias "FreeImage_ConvertLine1To8" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long, ByVal PalettePtr As LongPtr)
Public Declare PtrSafe Sub FreeImage_ConvertLine16To4_555 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine16To4_565 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine24To4 Lib "FreeImage_x64.dll" Alias "FreeImage_ConvertLine1To24" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine32To4 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine1To8 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine4To8 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine16To8_555 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine16To8_565 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine24To8 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine32To8 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine1To16_555 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long, ByVal PalettePtr As LongPtr)
Public Declare PtrSafe Sub FreeImage_ConvertLine4To16_555 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long, ByVal PalettePtr As LongPtr)
Public Declare PtrSafe Sub FreeImage_ConvertLine8To16_555 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long, ByVal PalettePtr As LongPtr)
Public Declare PtrSafe Sub FreeImage_ConvertLine16_565_To16_555 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine24To16_555 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine32To16_555 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine1To16_565 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long, ByVal PalettePtr As LongPtr)
Public Declare PtrSafe Sub FreeImage_ConvertLine4To16_565 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long, ByVal PalettePtr As LongPtr)
Public Declare PtrSafe Sub FreeImage_ConvertLine8To16_565 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long, ByVal PalettePtr As LongPtr)
Public Declare PtrSafe Sub FreeImage_ConvertLine16_555_To16_565 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine24To16_565 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine32To16_565 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine1To24 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long, ByVal PalettePtr As LongPtr)
Public Declare PtrSafe Sub FreeImage_ConvertLine4To24 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long, ByVal PalettePtr As LongPtr)
Public Declare PtrSafe Sub FreeImage_ConvertLine8To24 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long, ByVal PalettePtr As LongPtr)
Public Declare PtrSafe Sub FreeImage_ConvertLine16To24_555 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine16To24_565 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine32To24 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine1To32 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long, ByVal PalettePtr As LongPtr)
Public Declare PtrSafe Sub FreeImage_ConvertLine4To32 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long, ByVal PalettePtr As LongPtr)
Public Declare PtrSafe Sub FreeImage_ConvertLine8To32 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long, ByVal PalettePtr As LongPtr)
Public Declare PtrSafe Sub FreeImage_ConvertLine16To32_555 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine16To32_565 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
Public Declare PtrSafe Sub FreeImage_ConvertLine24To32 Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal SourcePtr As LongPtr, ByVal WidthInPixels As Long)
' Smart conversion functions
Public Declare PtrSafe Function FreeImage_ConvertTo4Bits Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_ConvertTo8Bits Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_ConvertToGreyscale Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_ConvertTo16Bits555 Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_ConvertTo16Bits565 Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_ConvertTo24Bits Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_ConvertTo32Bits Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_ColorQuantize Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal QuantizeMethod As FREE_IMAGE_QUANTIZE) As LongPtr
Public Declare PtrSafe Function FreeImage_Threshold Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal threshold As Byte) As LongPtr
Public Declare PtrSafe Function FreeImage_Dither Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal DitherMethod As FREE_IMAGE_DITHER) As LongPtr
Public Declare PtrSafe Function FreeImage_ConvertToFloat Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_ConvertToRGBF Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_ConvertToRGBAF Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_ConvertToUINT16 Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_ConvertToRGB16 Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Public Declare PtrSafe Function FreeImage_ConvertToRGBA16 Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr) As LongPtr
Private Declare PtrSafe Function p_FreeImage_ColorQuantizeEx Lib "FreeImage_x64.dll" Alias "FreeImage_ColorQuantizeEx" (ByVal BITMAP As LongPtr, Optional ByVal QuantizeMethod As FREE_IMAGE_QUANTIZE = FIQ_WUQUANT, Optional ByVal PaletteSize As Long = 256, Optional ByVal ReserveSize As Long = 0, Optional ByVal ReservePalettePtr As LongPtr = 0) As LongPtr
Private Declare PtrSafe Function p_FreeImage_ConvertFromRawBits Lib "FreeImage_x64.dll" Alias "FreeImage_ConvertFromRawBits" (ByVal BitsPtr As LongPtr, ByVal Width As Long, ByVal Height As Long, ByVal Pitch As Long, ByVal BitsPerPixel As Long, ByVal RedMask As Long, ByVal GreenMask As Long, ByVal BlueMask As Long, ByVal TopDown As Long) As LongPtr
Private Declare PtrSafe Function p_FreeImage_ConvertFromRawBitsEx Lib "FreeImage_x64.dll" Alias "FreeImage_ConvertFromRawBitsEx" (ByVal CopySource As Long, ByVal BitsPtr As LongPtr, ByVal ImageType As FREE_IMAGE_TYPE, ByVal Width As Long, ByVal Height As Long, ByVal Pitch As Long, ByVal BitsPerPixel As Long, ByVal RedMask As Long, ByVal GreenMask As Long, ByVal BlueMask As Long, ByVal TopDown As Long) As LongPtr
Private Declare PtrSafe Function p_FreeImage_ConvertToStandardType Lib "FreeImage_x64.dll" Alias "FreeImage_ConvertToStandardType" (ByVal BITMAP As LongPtr, ByVal ScaleLinear As Long) As LongPtr
Private Declare PtrSafe Function p_FreeImage_ConvertToType Lib "FreeImage_x64.dll" Alias "FreeImage_ConvertToType" (ByVal BITMAP As LongPtr, ByVal DestinationType As FREE_IMAGE_TYPE, ByVal ScaleLinear As Long) As LongPtr
Private Declare PtrSafe Sub p_FreeImage_ConvertToRawBits Lib "FreeImage_x64.dll" Alias "FreeImage_ConvertToRawBits" (ByVal BitsPtr As LongPtr, ByVal BITMAP As LongPtr, ByVal Pitch As Long, ByVal BitsPerPixel As Long, ByVal RedMask As Long, ByVal GreenMask As Long, ByVal BlueMask As Long, ByVal TopDown As Long)
' Tone mapping operators
Public Declare PtrSafe Function FreeImage_ToneMapping Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal Operator As FREE_IMAGE_TMO, Optional ByVal FirstArgument As Double, Optional ByVal SecondArgument As Double) As Long
Public Declare PtrSafe Function FreeImage_TmoDrago03 Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, Optional ByVal gamma As Double = 2.2, Optional ByVal Exposure As Double) As Long
Public Declare PtrSafe Function FreeImage_TmoReinhard05 Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, Optional ByVal Intensity As Double, Optional ByVal contrast As Double) As Long
Public Declare PtrSafe Function FreeImage_TmoReinhard05Ex Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, Optional ByVal Intensity As Double, Optional ByVal contrast As Double, Optional ByVal Adaptation As Double = 1, Optional ByVal ColorCorrection As Double) As Long
Public Declare PtrSafe Function FreeImage_TmoFattal02 Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, Optional ByVal ColorSaturation As Double = 0.5, Optional ByVal Attenuation As Double = 0.85) As Long
' ZLib functions
Public Declare PtrSafe Function FreeImage_ZLibCompress Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal TargetSize As Long, ByVal SourcePtr As LongPtr, ByVal SourceSize As Long) As Long
Public Declare PtrSafe Function FreeImage_ZLibUncompress Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal TargetSize As Long, ByVal SourcePtr As LongPtr, ByVal SourceSize As Long) As Long
Public Declare PtrSafe Function FreeImage_ZLibGZip Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal TargetSize As Long, ByVal SourcePtr As LongPtr, ByVal SourceSize As Long) As Long
Public Declare PtrSafe Function FreeImage_ZLibGUnzip Lib "FreeImage_x64.dll" (ByVal TargetPtr As LongPtr, ByVal TargetSize As Long, ByVal SourcePtr As LongPtr, ByVal SourceSize As Long) As Long
Public Declare PtrSafe Function FreeImage_ZLibCRC32 Lib "FreeImage_x64.dll" (ByVal CRC As Long, ByVal SourcePtr As LongPtr, ByVal SourceSize As Long) As Long
'----------------------
' Metadata functions
'----------------------
' tag creation / destruction
Private Declare PtrSafe Function p_FreeImage_CreateTag Lib "FreeImage_x64.dll" () As LongPtr
Private Declare PtrSafe Function p_FreeImage_CloneTag Lib "FreeImage_x64.dll" (ByVal Tag As LongPtr) As LongPtr
Private Declare PtrSafe Sub p_FreeImage_DeleteTag Lib "FreeImage_x64.dll" (ByVal Tag As LongPtr)
' tag getters and setters (only those actually needed by wrapper functions)
Private Declare PtrSafe Function p_FreeImage_SetTagKey Lib "FreeImage_x64.dll" (ByVal Tag As LongPtr, ByVal Key As String) As Long
Private Declare PtrSafe Function p_FreeImage_SetTagValue Lib "FreeImage_x64.dll" (ByVal Tag As LongPtr, ByVal ValuePtr As LongPtr) As Long
' metadata iterators
Public Declare PtrSafe Function FreeImage_FindFirstMetadata Lib "FreeImage_x64.dll" (ByVal Model As FREE_IMAGE_MDMODEL, ByVal BITMAP As LongPtr, ByRef Tag As LongPtr) As LongPtr
Public Declare PtrSafe Sub FreeImage_FindCloseMetadata Lib "FreeImage_x64.dll" (ByVal hFind As LongPtr)
Private Declare PtrSafe Function p_FreeImage_FindNextMetadata Lib "FreeImage_x64.dll" Alias "FreeImage_FindNextMetadata" (ByVal hFind As LongPtr, ByRef Tag As LongPtr) As Long
' metadata setters and getters
Private Declare PtrSafe Function p_FreeImage_SetMetadata Lib "FreeImage_x64.dll" Alias "FreeImage_SetMetadata" (ByVal Model As Long, ByVal BITMAP As LongPtr, ByVal Key As String, ByVal Tag As LongPtr) As Long
Private Declare PtrSafe Function p_FreeImage_GetMetadata Lib "FreeImage_x64.dll" Alias "FreeImage_GetMetadata" (ByVal Model As Long, ByVal BITMAP As LongPtr, ByVal Key As String, ByRef Tag As LongPtr) As Long
Private Declare PtrSafe Function p_FreeImage_SetMetadataKeyValue Lib "FreeImage_x64.dll" Alias "FreeImage_SetMetadataKeyValue" (ByVal Model As Long, ByVal BITMAP As LongPtr, ByVal Key As String, ByVal Tag As String) As Long
' metadata helper functions
Public Declare PtrSafe Function FreeImage_GetMetadataCount Lib "FreeImage_x64.dll" (ByVal Model As Long, ByVal BITMAP As LongPtr) As Long
Private Declare PtrSafe Function p_FreeImage_CloneMetadata Lib "FreeImage_x64.dll" Alias "FreeImage_CloneMetadata" (ByVal BitmapDst As LongPtr, ByVal BitmapSrc As LongPtr) As Long
' tag to string conversion functions
Private Declare PtrSafe Function p_FreeImage_TagToString Lib "FreeImage_x64.dll" Alias "FreeImage_TagToString" (ByVal Model As Long, ByVal Tag As LongPtr, Optional ByVal Make As String = vbNullString) As LongPtr
'----------------------
' JPEG lossless transformation functions
'----------------------
Private Declare PtrSafe Function p_FreeImage_JPEGTransform Lib "FreeImage_x64.dll" Alias "FreeImage_JPEGTransform" (ByVal SourceFile As String, ByVal DestFile As String, ByVal Operation As FREE_IMAGE_JPEG_OPERATION, ByVal Perfect As Long) As Long
Private Declare PtrSafe Function p_FreeImage_JPEGTransformU Lib "FreeImage_x64.dll" Alias "FreeImage_JPEGTransformU" (ByVal SourceFile As LongPtr, ByVal DestFile As LongPtr, ByVal Operation As FREE_IMAGE_JPEG_OPERATION, ByVal Perfect As Long) As Long
Private Declare PtrSafe Function p_FreeImage_JPEGCrop Lib "FreeImage_x64.dll" Alias "FreeImage_JPEGCrop" (ByVal SourceFile As String, ByVal DestFile As String, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Long
Private Declare PtrSafe Function p_FreeImage_JPEGCropU Lib "FreeImage_x64.dll" Alias "FreeImage_JPEGCropU" (ByVal SourceFile As LongPtr, ByVal DestFile As LongPtr, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Long
Private Declare PtrSafe Function p_FreeImage_JPEGTransformCombined Lib "FreeImage_x64.dll" Alias "FreeImage_JPEGTransformCombined" (ByVal SourceFile As String, ByVal DestFile As String, ByVal Operation As FREE_IMAGE_JPEG_OPERATION, ByRef Left As Long, ByRef Top As Long, ByRef Right As Long, ByRef Bottom As Long, ByVal Perfect As Long) As Long
Private Declare PtrSafe Function p_FreeImage_JPEGTransformCombinedU Lib "FreeImage_x64.dll" Alias "FreeImage_JPEGTransformCombinedU" (ByVal SourceFile As LongPtr, ByVal DestFile As LongPtr, ByVal Operation As FREE_IMAGE_JPEG_OPERATION, ByRef Left As Long, ByRef Top As Long, ByRef Right As Long, ByRef Bottom As Long, ByVal Perfect As Long) As Long
Private Declare PtrSafe Function p_FreeImage_JPEGTransformCombinedFromMemory Lib "FreeImage_x64.dll" Alias "FreeImage_JPEGTransformCombinedFromMemory" (ByVal SourceStream As LongPtr, ByVal DestStream As LongPtr, ByVal Operation As FREE_IMAGE_JPEG_OPERATION, ByRef Left As Long, ByRef Top As Long, ByRef Right As Long, ByRef Bottom As Long, ByVal Perfect As Long) As Long
'----------------------
' Image manipulation toolkit functions
'----------------------
' rotation and flipping
Public Declare PtrSafe Function FreeImage_RotateClassic Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal Angle As Double) As LongPtr
Public Declare PtrSafe Function FreeImage_Rotate Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal Angle As Double, Optional ByRef Color As Any = 0) As LongPtr
Private Declare PtrSafe Function p_FreeImage_RotateEx Lib "FreeImage_x64.dll" Alias "FreeImage_RotateEx" (ByVal BITMAP As LongPtr, ByVal Angle As Double, ByVal ShiftX As Double, ByVal ShiftY As Double, ByVal OriginX As Double, ByVal OriginY As Double, ByVal UseMask As Long) As LongPtr
Private Declare PtrSafe Function p_FreeImage_FlipHorizontal Lib "FreeImage_x64.dll" Alias "FreeImage_FlipHorizontal" (ByVal BITMAP As LongPtr) As Long
Private Declare PtrSafe Function p_FreeImage_FlipVertical Lib "FreeImage_x64.dll" Alias "FreeImage_FlipVertical" (ByVal BITMAP As LongPtr) As Long
' upsampling / downsampling
Public Declare PtrSafe Function FreeImage_Rescale Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal Width As Long, ByVal Height As Long, Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_CATMULLROM) As LongPtr
Public Declare PtrSafe Function FreeImage_RescaleRect Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal Width As Long, ByVal Height As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_CATMULLROM, Optional ByVal Flags As FREE_IMAGE_RESCALE_OPTIONS) As LongPtr
Private Declare PtrSafe Function p_FreeImage_MakeThumbnail Lib "FreeImage_x64.dll" Alias "FreeImage_MakeThumbnail" (ByVal BITMAP As LongPtr, ByVal MaxPixelSize As Long, Optional ByVal Convert As Long) As LongPtr
' color manipulation functions (point operations)
Public Declare PtrSafe Function FreeImage_SwapPaletteIndices Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByRef IndexA As Byte, ByRef IndexB As Byte) As Long
Private Declare PtrSafe Function p_FreeImage_AdjustCurve Lib "FreeImage_x64.dll" Alias "FreeImage_AdjustCurve" (ByVal BITMAP As LongPtr, ByVal LookupTablePtr As LongPtr, ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As Long
Private Declare PtrSafe Function p_FreeImage_AdjustGamma Lib "FreeImage_x64.dll" Alias "FreeImage_AdjustGamma" (ByVal BITMAP As LongPtr, ByVal gamma As Double) As Long
Private Declare PtrSafe Function p_FreeImage_AdjustBrightness Lib "FreeImage_x64.dll" Alias "FreeImage_AdjustBrightness" (ByVal BITMAP As LongPtr, ByVal Percentage As Double) As Long
Private Declare PtrSafe Function p_FreeImage_AdjustContrast Lib "FreeImage_x64.dll" Alias "FreeImage_AdjustContrast" (ByVal BITMAP As LongPtr, ByVal Percentage As Double) As Long
Private Declare PtrSafe Function p_FreeImage_Invert Lib "FreeImage_x64.dll" Alias "FreeImage_Invert" (ByVal BITMAP As LongPtr) As Long
Private Declare PtrSafe Function p_FreeImage_GetHistogram Lib "FreeImage_x64.dll" Alias "FreeImage_GetHistogram" (ByVal BITMAP As LongPtr, ByRef HistogramPtr As LongPtr, Optional ByVal Channel As FREE_IMAGE_COLOR_CHANNEL = FICC_BLACK) As Long
Private Declare PtrSafe Function p_FreeImage_GetAdjustColorsLookupTable Lib "FreeImage_x64.dll" Alias "FreeImage_GetAdjustColorsLookupTable" (ByVal LookupTablePtr As LongPtr, ByVal Brightness As Double, ByVal contrast As Double, ByVal gamma As Double, ByVal Invert As Long) As Long
Private Declare PtrSafe Function p_FreeImage_AdjustColors Lib "FreeImage_x64.dll" Alias "FreeImage_AdjustColors" (ByVal BITMAP As LongPtr, ByVal Brightness As Double, ByVal contrast As Double, ByVal gamma As Double, ByVal Invert As Long) As Long
Private Declare PtrSafe Function p_FreeImage_ApplyColorMapping Lib "FreeImage_x64.dll" Alias "FreeImage_ApplyColorMapping" (ByVal BITMAP As LongPtr, ByVal SourceColorsPtr As LongPtr, ByVal DestinationColorsPtr As LongPtr, ByVal Count As Long, ByVal IgnoreAlpha As Long, ByVal Swap As Long) As Long
Private Declare PtrSafe Function p_FreeImage_SwapColors Lib "FreeImage_x64.dll" Alias "FreeImage_SwapColors" (ByVal BITMAP As LongPtr, ByRef ColorA As RGBQUAD, ByRef ColorB As RGBQUAD, ByVal IgnoreAlpha As Long) As Long
Private Declare PtrSafe Function p_FreeImage_SwapColorsByLong Lib "FreeImage_x64.dll" Alias "FreeImage_SwapColors" (ByVal BITMAP As LongPtr, ByRef ColorA As Long, ByRef ColorB As Long, ByVal IgnoreAlpha As Long) As Long
Private Declare PtrSafe Function p_FreeImage_ApplyPaletteIndexMapping Lib "FreeImage_x64.dll" Alias "FreeImage_ApplyPaletteIndexMapping" (ByVal BITMAP As LongPtr, ByVal SourceIndicesPtr As LongPtr, ByVal DestinationIndicesPtr As LongPtr, ByVal Count As Long, ByVal Swap As Long) As Long
' channel processing functions
Public Declare PtrSafe Function FreeImage_GetChannel Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As LongPtr
Public Declare PtrSafe Function FreeImage_GetComplexChannel Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As LongPtr
Private Declare PtrSafe Function p_FreeImage_SetChannel Lib "FreeImage_x64.dll" Alias "FreeImage_SetChannel" (ByVal BitmapDst As LongPtr, ByVal BitmapSrc As LongPtr, ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As Long
Private Declare PtrSafe Function p_FreeImage_SetComplexChannel Lib "FreeImage_x64.dll" Alias "FreeImage_SetComplexChannel" (ByVal BitmapDst As LongPtr, ByVal BitmapSrc As LongPtr, ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As Long
' copy / paste / composite functions
Public Declare PtrSafe Function FreeImage_Copy Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As LongPtr
Public Declare PtrSafe Function FreeImage_CreateView Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As LongPtr
Public Declare PtrSafe Function FreeImage_Composite Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, Optional ByVal UseFileBackColor As Long, Optional ByRef AppBackColor As Any, Optional ByVal BackgroundBitmap As LongPtr) As LongPtr
Private Declare PtrSafe Function p_FreeImage_Paste Lib "FreeImage_x64.dll" Alias "FreeImage_Paste" (ByVal BitmapDst As LongPtr, ByVal BitmapSrc As LongPtr, ByVal Left As Long, ByVal Top As Long, ByVal Alpha As Long) As Long
Private Declare PtrSafe Function p_FreeImage_PreMultiplyWithAlpha Lib "FreeImage_x64.dll" Alias "FreeImage_PreMultiplyWithAlpha" (ByVal BITMAP As LongPtr) As Long
' background filling functions
Public Declare PtrSafe Function FreeImage_FillBackground Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByRef Color As Any, Optional ByVal Options As FREE_IMAGE_COLOR_OPTIONS = FI_COLOR_IS_RGB_COLOR) As Long
Public Declare PtrSafe Function FreeImage_EnlargeCanvas Lib "FreeImage_x64.dll" (ByVal BITMAP As LongPtr, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByRef Color As Any, Optional ByVal Options As FREE_IMAGE_COLOR_OPTIONS = FI_COLOR_IS_RGB_COLOR) As LongPtr
Public Declare PtrSafe Function FreeImage_AllocateEx Lib "FreeImage_x64.dll" (ByVal Width As Long, ByVal Height As Long, Optional ByVal BitsPerPixel As Long = 8, Optional ByRef Color As Any, Optional ByVal Options As FREE_IMAGE_COLOR_OPTIONS, Optional ByVal PalettePtr As LongPtr = 0, Optional ByVal RedMask As Long = 0, Optional ByVal GreenMask As Long = 0&, Optional ByVal BlueMask As Long = 0&) As LongPtr
Public Declare PtrSafe Function FreeImage_AllocateExT Lib "FreeImage_x64.dll" (ByVal ImageType As FREE_IMAGE_TYPE, ByVal Width As Long, ByVal Height As Long, Optional ByVal BitsPerPixel As Long = 8, Optional ByRef Color As Any, Optional ByVal Options As FREE_IMAGE_COLOR_OPTIONS, Optional ByVal PalettePtr As LongPtr, Optional ByVal RedMask As Long = 0&, Optional ByVal GreenMask As Long = 0&, Optional ByVal BlueMask As Long = 0&) As LongPtr
' miscellaneous algorithms
Public Declare PtrSafe Function FreeImage_MultigridPoissonSolver Lib "FreeImage_x64.dll" (ByVal LaplacianBitmap As LongPtr, Optional ByVal Cyles As Long = 3) As LongPtr
#Else                       ' <OFFICE97-2007>
' Initialization / Deinitialization functions
Public Declare Sub FreeImage_Initialise Lib "FreeImage.dll" Alias "_FreeImage_Initialise@4" (Optional ByVal LoadLocalPluginsOnly As Long)
Public Declare Sub FreeImage_DeInitialise Lib "FreeImage.dll" Alias "_FreeImage_DeInitialise@0" ()
' Version functions
Private Declare Function p_FreeImage_GetVersion Lib "FreeImage.dll" Alias "_FreeImage_GetVersion@0" () As Long
Private Declare Function p_FreeImage_GetCopyrightMessage Lib "FreeImage.dll" Alias "_FreeImage_GetCopyrightMessage@0" () As Long
' Message output functions
Public Declare Sub FreeImage_SetOutputMessage Lib "FreeImage.dll" Alias "_FreeImage_SetOutputMessageStdCall@4" (ByVal omf As Long)
' Allocate / Clone / Unload functions
Public Declare Function FreeImage_Allocate Lib "FreeImage.dll" Alias "_FreeImage_Allocate@24" (ByVal Width As Long, ByVal Height As Long, ByVal BitsPerPixel As Long, Optional ByVal RedMask As Long = 0&, Optional ByVal GreenMask As Long = 0&, Optional ByVal BlueMask As Long = 0&) As Long
Public Declare Function FreeImage_AllocateT Lib "FreeImage.dll" Alias "_FreeImage_AllocateT@28" (ByVal ImageType As FREE_IMAGE_TYPE, ByVal Width As Long, ByVal Height As Long, Optional ByVal BitsPerPixel As Long = 8, Optional ByVal RedMask As Long = 0&, Optional ByVal GreenMask As Long = 0&, Optional ByVal BlueMask As Long = 0&) As Long
Public Declare Function FreeImage_Clone Lib "FreeImage.dll" Alias "_FreeImage_Clone@4" (ByVal BITMAP As Long) As Long
Public Declare Sub FreeImage_Unload Lib "FreeImage.dll" Alias "_FreeImage_Unload@4" (ByVal BITMAP As Long)
' Header loading functions
Public Declare Function FreeImage_HasPixelsLng Lib "FreeImage.dll" Alias "_FreeImage_HasPixels@4" (ByVal BITMAP As Long) As Long
' Load / Save functions
Public Declare Function FreeImage_Load Lib "FreeImage.dll" Alias "_FreeImage_Load@12" (ByVal Format As FREE_IMAGE_FORMAT, ByVal FileName As String, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long
Public Declare Function FreeImage_LoadFromHandle Lib "FreeImage.dll" Alias "_FreeImage_LoadFromHandle@16" (ByVal Format As FREE_IMAGE_FORMAT, ByVal IO As Long, ByVal Handle As Long, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long
Private Declare Function p_FreeImage_LoadU Lib "FreeImage.dll" Alias "_FreeImage_LoadU@12" (ByVal Format As FREE_IMAGE_FORMAT, ByVal FileName As Long, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long
Private Declare Function p_FreeImage_Save Lib "FreeImage.dll" Alias "_FreeImage_Save@16" (ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As Long, ByVal FileName As String, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long
Private Declare Function p_FreeImage_SaveU Lib "FreeImage.dll" Alias "_FreeImage_SaveU@16" (ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As Long, ByVal FileName As Long, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long
Private Declare Function p_FreeImage_SaveToHandle Lib "FreeImage.dll" Alias "_FreeImage_SaveToHandle@20" (ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As Long, ByVal IO As Long, ByVal Handle As Long, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long
' Memory I/O stream functions
Public Declare Function FreeImage_OpenMemory Lib "FreeImage.dll" Alias "_FreeImage_OpenMemory@8" (Optional ByRef Data As Byte, Optional ByVal SizeInBytes As Long) As Long
Public Declare Function FreeImage_OpenMemoryByPtr Lib "FreeImage.dll" Alias "_FreeImage_OpenMemory@8" (Optional ByVal DataPtr As Long, Optional ByVal SizeInBytes As Long) As Long
Public Declare Sub FreeImage_CloseMemory Lib "FreeImage.dll" Alias "_FreeImage_CloseMemory@4" (ByVal Stream As Long)
Public Declare Function FreeImage_LoadFromMemory Lib "FreeImage.dll" Alias "_FreeImage_LoadFromMemory@12" (ByVal Format As FREE_IMAGE_FORMAT, ByVal Stream As Long, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long
Public Declare Function FreeImage_TellMemory Lib "FreeImage.dll" Alias "_FreeImage_TellMemory@4" (ByVal Stream As Long) As Long
Public Declare Function FreeImage_ReadMemory Lib "FreeImage.dll" Alias "_FreeImage_ReadMemory@16" (ByVal BufferPtr As Long, ByVal Size As Long, ByVal Count As Long, ByVal Stream As Long) As Long
Public Declare Function FreeImage_WriteMemory Lib "FreeImage.dll" Alias "_FreeImage_WriteMemory@16" (ByVal BufferPtr As Long, ByVal Size As Long, ByVal Count As Long, ByVal Stream As Long) As Long
Public Declare Function FreeImage_LoadMultiBitmapFromMemory Lib "FreeImage.dll" Alias "_FreeImage_LoadMultiBitmapFromMemory@12" (ByVal Format As FREE_IMAGE_FORMAT, ByVal Stream As Long, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long
Public Declare Function FreeImage_SaveMultiBitmapToMemory Lib "FreeImage.dll" Alias "_FreeImage_SaveMultiBitmapToMemory@16" (ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As Long, ByVal Stream As Long, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long
Private Declare Function p_FreeImage_SaveToMemory Lib "FreeImage.dll" Alias "_FreeImage_SaveToMemory@16" (ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As Long, ByVal Stream As Long, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long
Private Declare Function p_FreeImage_SeekMemory Lib "FreeImage.dll" Alias "_FreeImage_SeekMemory@12" (ByVal Stream As Long, ByVal Offset As Long, ByVal Origin As Long) As Long
Private Declare Function p_FreeImage_AcquireMemory Lib "FreeImage.dll" Alias "_FreeImage_AcquireMemory@12" (ByVal Stream As Long, ByRef DataPtr As Long, ByRef SizeInBytes As Long) As Long
' Plugin / Format functions
Public Declare Function FreeImage_RegisterLocalPlugin Lib "FreeImage.dll" Alias "_FreeImage_RegisterLocalPlugin@20" (ByVal InitProcAddress As Long, Optional ByVal Format As String, Optional ByVal Description As String, Optional ByVal Extension As String, Optional ByVal RegExpr As String) As FREE_IMAGE_FORMAT
Public Declare Function FreeImage_RegisterExternalPlugin Lib "FreeImage.dll" Alias "_FreeImage_RegisterExternalPlugin@20" (ByVal path As String, Optional ByVal Format As String, Optional ByVal Description As String, Optional ByVal Extension As String, Optional ByVal RegExpr As String) As FREE_IMAGE_FORMAT
Public Declare Function FreeImage_GetFIFCount Lib "FreeImage.dll" Alias "_FreeImage_GetFIFCount@0" () As Long
Public Declare Function FreeImage_SetPluginEnabled Lib "FreeImage.dll" Alias "_FreeImage_SetPluginEnabled@8" (ByVal Format As FREE_IMAGE_FORMAT, ByVal Value As Long) As Long
Public Declare Function FreeImage_IsPluginEnabled Lib "FreeImage.dll" Alias "_FreeImage_IsPluginEnabled@4" (ByVal Format As FREE_IMAGE_FORMAT) As Long
Public Declare Function FreeImage_GetFIFFromFormat Lib "FreeImage.dll" Alias "_FreeImage_GetFIFFromFormat@4" (ByVal Format As String) As FREE_IMAGE_FORMAT
Public Declare Function FreeImage_GetFIFFromMime Lib "FreeImage.dll" Alias "_FreeImage_GetFIFFromMime@4" (ByVal MimeType As String) As FREE_IMAGE_FORMAT
Public Declare Function FreeImage_GetFIFFromFilename Lib "FreeImage.dll" Alias "_FreeImage_GetFIFFromFilename@4" (ByVal FileName As String) As FREE_IMAGE_FORMAT
Private Declare Function p_FreeImage_GetFormatFromFIF Lib "FreeImage.dll" Alias "_FreeImage_GetFormatFromFIF@4" (ByVal Format As FREE_IMAGE_FORMAT) As Long
Private Declare Function p_FreeImage_GetFIFExtensionList Lib "FreeImage.dll" Alias "_FreeImage_GetFIFExtensionList@4" (ByVal Format As FREE_IMAGE_FORMAT) As Long
Private Declare Function p_FreeImage_GetFIFDescription Lib "FreeImage.dll" Alias "_FreeImage_GetFIFDescription@4" (ByVal Format As FREE_IMAGE_FORMAT) As Long
Private Declare Function p_FreeImage_GetFIFRegExpr Lib "FreeImage.dll" Alias "_FreeImage_GetFIFRegExpr@4" (ByVal Format As FREE_IMAGE_FORMAT) As Long
Private Declare Function p_FreeImage_GetFIFMimeType Lib "FreeImage.dll" Alias "_FreeImage_GetFIFMimeType@4" (ByVal Format As FREE_IMAGE_FORMAT) As Long
Private Declare Function p_FreeImage_GetFIFFromFilenameU Lib "FreeImage.dll" Alias "_FreeImage_GetFIFFromFilenameU@4" (ByVal FileName As Long) As FREE_IMAGE_FORMAT
Private Declare Function p_FreeImage_FIFSupportsReading Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsReading@4" (ByVal Format As FREE_IMAGE_FORMAT) As Long
Private Declare Function p_FreeImage_FIFSupportsWriting Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsWriting@4" (ByVal Format As FREE_IMAGE_FORMAT) As Long
Private Declare Function p_FreeImage_FIFSupportsExportBPP Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsExportBPP@8" (ByVal Format As FREE_IMAGE_FORMAT, ByVal BitsPerPixel As Long) As Long
Private Declare Function p_FreeImage_FIFSupportsExportType Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsExportType@8" (ByVal Format As FREE_IMAGE_FORMAT, ByVal ImageType As FREE_IMAGE_TYPE) As Long
Private Declare Function p_FreeImage_FIFSupportsICCProfiles Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsICCProfiles@4" (ByVal Format As FREE_IMAGE_FORMAT) As Long
Private Declare Function p_FreeImage_FIFSupportsNoPixels Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsNoPixels@4" (ByVal Format As FREE_IMAGE_FORMAT) As Long
' Multipaging functions
Public Declare Function FreeImage_GetPageCount Lib "FreeImage.dll" Alias "_FreeImage_GetPageCount@4" (ByVal BITMAP As Long) As Long
Public Declare Sub FreeImage_AppendPage Lib "FreeImage.dll" Alias "_FreeImage_AppendPage@8" (ByVal BITMAP As Long, ByVal PageBitmap As Long)
Public Declare Sub FreeImage_InsertPage Lib "FreeImage.dll" Alias "_FreeImage_InsertPage@12" (ByVal BITMAP As Long, ByVal Page As Long, ByVal PageBitmap As Long)
Public Declare Sub FreeImage_DeletePage Lib "FreeImage.dll" Alias "_FreeImage_DeletePage@8" (ByVal BITMAP As Long, ByVal Page As Long)
Public Declare Function FreeImage_LockPage Lib "FreeImage.dll" Alias "_FreeImage_LockPage@8" (ByVal BITMAP As Long, ByVal Page As Long) As Long
Private Declare Function p_FreeImage_OpenMultiBitmap Lib "FreeImage.dll" Alias "_FreeImage_OpenMultiBitmap@24" (ByVal Format As FREE_IMAGE_FORMAT, ByVal FileName As String, ByVal CreateNew As Long, ByVal ReadOnly As Long, ByVal KeepCacheInMemory As Long, ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As Long
Private Declare Function p_FreeImage_CloseMultiBitmap Lib "FreeImage.dll" Alias "_FreeImage_CloseMultiBitmap@8" (ByVal BITMAP As Long, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Long
Private Declare Function p_FreeImage_MovePage Lib "FreeImage.dll" Alias "_FreeImage_MovePage@12" (ByVal BITMAP As Long, ByVal TargetPage As Long, ByVal SourcePage As Long) As Long
Private Declare Function p_FreeImage_GetLockedPageNumbers Lib "FreeImage.dll" Alias "_FreeImage_GetLockedPageNumbers@12" (ByVal BITMAP As Long, ByRef PagesPtr As Long, ByRef Count As Long) As Long
Private Declare Sub p_FreeImage_UnlockPage Lib "FreeImage.dll" Alias "_FreeImage_UnlockPage@12" (ByVal BITMAP As Long, ByVal PageBitmap As Long, ByVal ApplyChanges As Long)
' Filetype request functions
Public Declare Function FreeImage_GetFileType Lib "FreeImage.dll" Alias "_FreeImage_GetFileType@8" (ByVal FileName As String, Optional ByVal Size As Long) As FREE_IMAGE_FORMAT
Public Declare Function FreeImage_GetFileTypeFromHandle Lib "FreeImage.dll" Alias "_FreeImage_GetFileTypeFromHandle@12" (ByVal IO As Long, ByVal Handle As Long, Optional ByVal Size As Long) As FREE_IMAGE_FORMAT
Public Declare Function FreeImage_GetFileTypeFromMemory Lib "FreeImage.dll" Alias "_FreeImage_GetFileTypeFromMemory@8" (ByVal Stream As Long, Optional ByVal Size As Long) As FREE_IMAGE_FORMAT
Private Declare Function p_FreeImage_GetFileTypeU Lib "FreeImage.dll" Alias "_FreeImage_GetFileTypeU@8" (ByVal FileName As Long, Optional ByVal Size As Long) As FREE_IMAGE_FORMAT
' Image type request functions
Public Declare Function FreeImage_GetImageType Lib "FreeImage.dll" Alias "_FreeImage_GetImageType@4" (ByVal BITMAP As Long) As FREE_IMAGE_TYPE
' FreeImage helper functions
Private Declare Function p_FreeImage_IsLittleEndian Lib "FreeImage.dll" Alias "_FreeImage_IsLittleEndian@0" () As Long
Private Declare Function p_FreeImage_LookupX11Color Lib "FreeImage.dll" Alias "_FreeImage_LookupX11Color@16" (ByVal Color As String, ByRef Red As Long, ByRef Green As Long, ByRef Blue As Long) As Long
Private Declare Function p_FreeImage_LookupSVGColor Lib "FreeImage.dll" Alias "_FreeImage_LookupSVGColor@16" (ByVal Color As String, ByRef Red As Long, ByRef Green As Long, ByRef Blue As Long) As Long
' Pixel access functions
Public Declare Function FreeImage_GetBits Lib "FreeImage.dll" Alias "_FreeImage_GetBits@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetScanline Lib "FreeImage.dll" Alias "_FreeImage_GetScanLine@8" (ByVal BITMAP As Long, ByVal Scanline As Long) As Long
Private Declare Function p_FreeImage_GetPixelIndex Lib "FreeImage.dll" Alias "_FreeImage_GetPixelIndex@16" (ByVal BITMAP As Long, ByVal x As Long, ByVal y As Long, ByRef Value As Byte) As Long
Private Declare Function p_FreeImage_GetPixelColor Lib "FreeImage.dll" Alias "_FreeImage_GetPixelColor@16" (ByVal BITMAP As Long, ByVal x As Long, ByVal y As Long, ByRef Value As RGBQUAD) As Long
Private Declare Function p_FreeImage_GetPixelColorByLong Lib "FreeImage.dll" Alias "_FreeImage_GetPixelColor@16" (ByVal BITMAP As Long, ByVal x As Long, ByVal y As Long, ByRef Value As Long) As Long
Private Declare Function p_FreeImage_SetPixelIndex Lib "FreeImage.dll" Alias "_FreeImage_SetPixelIndex@16" (ByVal BITMAP As Long, ByVal x As Long, ByVal y As Long, ByRef Value As Byte) As Long
Private Declare Function p_FreeImage_SetPixelColor Lib "FreeImage.dll" Alias "_FreeImage_SetPixelColor@16" (ByVal BITMAP As Long, ByVal x As Long, ByVal y As Long, ByRef Value As RGBQUAD) As Long
Private Declare Function p_FreeImage_SetPixelColorByLong Lib "FreeImage.dll" Alias "_FreeImage_SetPixelColor@16" (ByVal BITMAP As Long, ByVal x As Long, ByVal y As Long, ByRef Value As Long) As Long
' DIB info functions
Public Declare Function FreeImage_GetColorsUsed Lib "FreeImage.dll" Alias "_FreeImage_GetColorsUsed@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetBPP Lib "FreeImage.dll" Alias "_FreeImage_GetBPP@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetWidth Lib "FreeImage.dll" Alias "_FreeImage_GetWidth@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetHeight Lib "FreeImage.dll" Alias "_FreeImage_GetHeight@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetLine Lib "FreeImage.dll" Alias "_FreeImage_GetLine@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetPitch Lib "FreeImage.dll" Alias "_FreeImage_GetPitch@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetDIBSize Lib "FreeImage.dll" Alias "_FreeImage_GetDIBSize@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetMemorySize Lib "FreeImage.dll" Alias "_FreeImage_GetMemorySize@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetPalette Lib "FreeImage.dll" Alias "_FreeImage_GetPalette@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetDotsPerMeterX Lib "FreeImage.dll" Alias "_FreeImage_GetDotsPerMeterX@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetDotsPerMeterY Lib "FreeImage.dll" Alias "_FreeImage_GetDotsPerMeterY@4" (ByVal BITMAP As Long) As Long
Public Declare Sub FreeImage_SetDotsPerMeterX Lib "FreeImage.dll" Alias "_FreeImage_SetDotsPerMeterX@8" (ByVal BITMAP As Long, ByVal resolution As Long)
Public Declare Sub FreeImage_SetDotsPerMeterY Lib "FreeImage.dll" Alias "_FreeImage_SetDotsPerMeterY@8" (ByVal BITMAP As Long, ByVal resolution As Long)
Public Declare Function FreeImage_GetInfoHeader Lib "FreeImage.dll" Alias "_FreeImage_GetInfoHeader@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetInfo Lib "FreeImage.dll" Alias "_FreeImage_GetInfo@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetColorType Lib "FreeImage.dll" Alias "_FreeImage_GetColorType@4" (ByVal BITMAP As Long) As FREE_IMAGE_COLOR_TYPE
Private Declare Function p_FreeImage_HasRGBMasks Lib "FreeImage.dll" Alias "_FreeImage_HasRGBMasks@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetRedMask Lib "FreeImage.dll" Alias "_FreeImage_GetRedMask@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetGreenMask Lib "FreeImage.dll" Alias "_FreeImage_GetGreenMask@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetBlueMask Lib "FreeImage.dll" Alias "_FreeImage_GetBlueMask@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetTransparencyCount Lib "FreeImage.dll" Alias "_FreeImage_GetTransparencyCount@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetTransparencyTable Lib "FreeImage.dll" Alias "_FreeImage_GetTransparencyTable@4" (ByVal BITMAP As Long) As Long
Public Declare Sub FreeImage_SetTransparencyTable Lib "FreeImage.dll" Alias "_FreeImage_SetTransparencyTable@12" (ByVal BITMAP As Long, ByVal TransTablePtr As Long, ByVal Count As Long)
Public Declare Function FreeImage_SetTransparentIndex Lib "FreeImage.dll" Alias "_FreeImage_SetTransparentIndex@8" (ByVal BITMAP As Long, ByVal Index As Long) As Long
Public Declare Function FreeImage_GetTransparentIndex Lib "FreeImage.dll" Alias "_FreeImage_GetTransparentIndex@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_GetThumbnail Lib "FreeImage.dll" Alias "_FreeImage_GetThumbnail@4" (ByVal BITMAP As Long) As Long
Private Declare Function p_FreeImage_IsTransparent Lib "FreeImage.dll" Alias "_FreeImage_IsTransparent@4" (ByVal BITMAP As Long) As Long
Private Declare Function p_FreeImage_HasBackgroundColor Lib "FreeImage.dll" Alias "_FreeImage_HasBackgroundColor@4" (ByVal BITMAP As Long) As Long
Private Declare Function p_FreeImage_GetBackgroundColor Lib "FreeImage.dll" Alias "_FreeImage_GetBackgroundColor@8" (ByVal BITMAP As Long, ByRef BackColor As RGBQUAD) As Long
Private Declare Function p_FreeImage_GetBackgroundColorAsLong Lib "FreeImage.dll" Alias "_FreeImage_GetBackgroundColor@8" (ByVal BITMAP As Long, ByRef BackColor As Long) As Long
Private Declare Function p_FreeImage_SetBackgroundColor Lib "FreeImage.dll" Alias "_FreeImage_SetBackgroundColor@8" (ByVal BITMAP As Long, ByRef BackColor As RGBQUAD) As Long
Private Declare Function p_FreeImage_SetBackgroundColorAsLong Lib "FreeImage.dll" Alias "_FreeImage_SetBackgroundColor@8" (ByVal BITMAP As Long, ByRef BackColor As Long) As Long
Private Declare Function p_FreeImage_SetThumbnail Lib "FreeImage.dll" Alias "_FreeImage_SetThumbnail@8" (ByVal BITMAP As Long, ByVal Thumbnail As Long) As Long
Private Declare Sub p_FreeImage_SetTransparent Lib "FreeImage.dll" Alias "_FreeImage_SetTransparent@8" (ByVal BITMAP As Long, ByVal Value As Long)
' ICC profile functions
Public Declare Function FreeImage_CreateICCProfile Lib "FreeImage.dll" Alias "_FreeImage_CreateICCProfile@12" (ByVal BITMAP As Long, ByRef Data As Long, ByVal Size As Long) As Long
Public Declare Sub FreeImage_DestroyICCProfile Lib "FreeImage.dll" Alias "_FreeImage_DestroyICCProfile@4" (ByVal BITMAP As Long)
Private Declare Function p_FreeImage_GetICCProfile Lib "FreeImage.dll" Alias "_FreeImage_GetICCProfile@4" (ByVal BITMAP As Long) As Long
' Line conversion functions
Public Declare Sub FreeImage_ConvertLine1To4 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To4@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine8To4 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To8@16" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long, ByVal PalettePtr As Long)
Public Declare Sub FreeImage_ConvertLine16To4_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To4_555@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine16To4_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To4_565@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine24To4 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To24@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine32To4 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine32To4@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine1To8 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To8@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine4To8 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine4To8@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine16To8_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To8_555@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine16To8_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To8_565@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine24To8 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine24To8@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine32To8 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine32To8@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine1To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To16_555@16" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long, ByVal PalettePtr As Long)
Public Declare Sub FreeImage_ConvertLine4To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine4To16_555@16" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long, ByVal PalettePtr As Long)
Public Declare Sub FreeImage_ConvertLine8To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine8To16_555@16" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long, ByVal PalettePtr As Long)
Public Declare Sub FreeImage_ConvertLine16_565_To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16_565_To16_555@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine24To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine24To16_555@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine32To16_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine32To16_555@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine1To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To16_565@16" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long, ByVal PalettePtr As Long)
Public Declare Sub FreeImage_ConvertLine4To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine4To16_565@16" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long, ByVal PalettePtr As Long)
Public Declare Sub FreeImage_ConvertLine8To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine8To16_565@16" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long, ByVal PalettePtr As Long)
Public Declare Sub FreeImage_ConvertLine16_555_To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16_555_To16_565@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine24To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine24To16_565@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine32To16_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine32To16_565@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine1To24 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To24@16" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long, ByVal PalettePtr As Long)
Public Declare Sub FreeImage_ConvertLine4To24 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine4To24@16" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long, ByVal PalettePtr As Long)
Public Declare Sub FreeImage_ConvertLine8To24 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine8To24@16" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long, ByVal PalettePtr As Long)
Public Declare Sub FreeImage_ConvertLine16To24_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To24_555@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine16To24_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To24_565@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine32To24 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine32To24@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine1To32 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine1To32@16" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long, ByVal PalettePtr As Long)
Public Declare Sub FreeImage_ConvertLine4To32 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine4To32@16" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long, ByVal PalettePtr As Long)
Public Declare Sub FreeImage_ConvertLine8To32 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine8To32@16" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long, ByVal PalettePtr As Long)
Public Declare Sub FreeImage_ConvertLine16To32_555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To32_555@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine16To32_565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine16To32_565@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
Public Declare Sub FreeImage_ConvertLine24To32 Lib "FreeImage.dll" Alias "_FreeImage_ConvertLine24To32@12" (ByVal TargetPtr As Long, ByVal sourcePtr As Long, ByVal WidthInPixels As Long)
' Smart conversion functions
Public Declare Function FreeImage_ConvertTo4Bits Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo4Bits@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_ConvertTo8Bits Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo8Bits@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_ConvertToGreyscale Lib "FreeImage.dll" Alias "_FreeImage_ConvertToGreyscale@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_ConvertTo16Bits555 Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo16Bits555@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_ConvertTo16Bits565 Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo16Bits565@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_ConvertTo24Bits Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo24Bits@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_ConvertTo32Bits Lib "FreeImage.dll" Alias "_FreeImage_ConvertTo32Bits@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_ColorQuantize Lib "FreeImage.dll" Alias "_FreeImage_ColorQuantize@8" (ByVal BITMAP As Long, ByVal QuantizeMethod As FREE_IMAGE_QUANTIZE) As Long
Public Declare Function FreeImage_Threshold Lib "FreeImage.dll" Alias "_FreeImage_Threshold@8" (ByVal BITMAP As Long, ByVal threshold As Byte) As Long
Public Declare Function FreeImage_Dither Lib "FreeImage.dll" Alias "_FreeImage_Dither@8" (ByVal BITMAP As Long, ByVal DitherMethod As FREE_IMAGE_DITHER) As Long
Public Declare Function FreeImage_ConvertToFloat Lib "FreeImage.dll" Alias "_FreeImage_ConvertToFloat@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_ConvertToRGBF Lib "FreeImage.dll" Alias "_FreeImage_ConvertToRGBF@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_ConvertToRGBAF Lib "FreeImage.dll" Alias "_FreeImage_ConvertToRGBAF@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_ConvertToUINT16 Lib "FreeImage.dll" Alias "_FreeImage_ConvertToUINT16@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_ConvertToRGB16 Lib "FreeImage.dll" Alias "_FreeImage_ConvertToRGB16@4" (ByVal BITMAP As Long) As Long
Public Declare Function FreeImage_ConvertToRGBA16 Lib "FreeImage.dll" Alias "_FreeImage_ConvertToRGBA16@4" (ByVal BITMAP As Long) As Long
Private Declare Function p_FreeImage_ColorQuantizeEx Lib "FreeImage.dll" Alias "_FreeImage_ColorQuantizeEx@20" (ByVal BITMAP As Long, Optional ByVal QuantizeMethod As FREE_IMAGE_QUANTIZE = FIQ_WUQUANT, Optional ByVal PaletteSize As Long = 256, Optional ByVal ReserveSize As Long = 0, Optional ByVal ReservePalettePtr As Long = 0) As Long
Private Declare Function p_FreeImage_ConvertFromRawBits Lib "FreeImage.dll" Alias "_FreeImage_ConvertFromRawBits@36" (ByVal BitsPtr As Long, ByVal Width As Long, ByVal Height As Long, ByVal Pitch As Long, ByVal BitsPerPixel As Long, ByVal RedMask As Long, ByVal GreenMask As Long, ByVal BlueMask As Long, ByVal TopDown As Long) As Long
Private Declare Function p_FreeImage_ConvertFromRawBitsEx Lib "FreeImage.dll" Alias "_FreeImage_ConvertFromRawBitsEx@44" (ByVal CopySource As Long, ByVal BitsPtr As Long, ByVal ImageType As FREE_IMAGE_TYPE, ByVal Width As Long, ByVal Height As Long, ByVal Pitch As Long, ByVal BitsPerPixel As Long, ByVal RedMask As Long, ByVal GreenMask As Long, ByVal BlueMask As Long, ByVal TopDown As Long) As Long
Private Declare Function p_FreeImage_ConvertToStandardType Lib "FreeImage.dll" Alias "_FreeImage_ConvertToStandardType@8" (ByVal BITMAP As Long, ByVal ScaleLinear As Long) As Long
Private Declare Function p_FreeImage_ConvertToType Lib "FreeImage.dll" Alias "_FreeImage_ConvertToType@12" (ByVal BITMAP As Long, ByVal DestinationType As FREE_IMAGE_TYPE, ByVal ScaleLinear As Long) As Long
Private Declare Sub p_FreeImage_ConvertToRawBits Lib "FreeImage.dll" Alias "_FreeImage_ConvertToRawBits@32" (ByVal BitsPtr As Long, ByVal BITMAP As Long, ByVal Pitch As Long, ByVal BitsPerPixel As Long, ByVal RedMask As Long, ByVal GreenMask As Long, ByVal BlueMask As Long, ByVal TopDown As Long)
' Tone mapping operators
Public Declare Function FreeImage_ToneMapping Lib "FreeImage.dll" Alias "_FreeImage_ToneMapping@24" (ByVal BITMAP As Long, ByVal Operator As FREE_IMAGE_TMO, Optional ByVal FirstArgument As Double, Optional ByVal SecondArgument As Double) As Long
Public Declare Function FreeImage_TmoDrago03 Lib "FreeImage.dll" Alias "_FreeImage_TmoDrago03@20" (ByVal BITMAP As Long, Optional ByVal gamma As Double = 2.2, Optional ByVal Exposure As Double) As Long
Public Declare Function FreeImage_TmoReinhard05 Lib "FreeImage.dll" Alias "_FreeImage_TmoReinhard05@20" (ByVal BITMAP As Long, Optional ByVal Intensity As Double, Optional ByVal contrast As Double) As Long
Public Declare Function FreeImage_TmoReinhard05Ex Lib "FreeImage.dll" Alias "_FreeImage_TmoReinhard05Ex@36" (ByVal BITMAP As Long, Optional ByVal Intensity As Double, Optional ByVal contrast As Double, Optional ByVal Adaptation As Double = 1, Optional ByVal ColorCorrection As Double) As Long
Public Declare Function FreeImage_TmoFattal02 Lib "FreeImage.dll" Alias "_FreeImage_TmoFattal02@20" (ByVal BITMAP As Long, Optional ByVal ColorSaturation As Double = 0.5, Optional ByVal Attenuation As Double = 0.85) As Long
' ZLib functions
Public Declare Function FreeImage_ZLibCompress Lib "FreeImage.dll" Alias "_FreeImage_ZLibCompress@16" (ByVal TargetPtr As Long, ByVal TargetSize As Long, ByVal sourcePtr As Long, ByVal SourceSize As Long) As Long
Public Declare Function FreeImage_ZLibUncompress Lib "FreeImage.dll" Alias "_FreeImage_ZLibUncompress@16" (ByVal TargetPtr As Long, ByVal TargetSize As Long, ByVal sourcePtr As Long, ByVal SourceSize As Long) As Long
Public Declare Function FreeImage_ZLibGZip Lib "FreeImage.dll" Alias "_FreeImage_ZLibGZip@16" (ByVal TargetPtr As Long, ByVal TargetSize As Long, ByVal sourcePtr As Long, ByVal SourceSize As Long) As Long
Public Declare Function FreeImage_ZLibGUnzip Lib "FreeImage.dll" Alias "_FreeImage_ZLibGUnzip@16" (ByVal TargetPtr As Long, ByVal TargetSize As Long, ByVal sourcePtr As Long, ByVal SourceSize As Long) As Long
Public Declare Function FreeImage_ZLibCRC32 Lib "FreeImage.dll" Alias "_FreeImage_ZLibCRC32@12" (ByVal CRC As Long, ByVal sourcePtr As Long, ByVal SourceSize As Long) As Long
'----------------------
' Metadata functions
'----------------------
' tag creation / destruction
Private Declare Function p_FreeImage_CreateTag Lib "FreeImage.dll" Alias "_p_FreeImage_CreateTag@0" () As Long
Private Declare Sub p_FreeImage_DeleteTag Lib "FreeImage.dll" Alias "_p_FreeImage_DeleteTag@4" (ByVal Tag As Long)
Private Declare Function p_FreeImage_CloneTag Lib "FreeImage.dll" Alias "_p_FreeImage_CloneTag@4" (ByVal Tag As Long) As Long
' tag getters and setters (only those actually needed by wrapper functions)
Private Declare Function p_FreeImage_SetTagKey Lib "FreeImage.dll" Alias "_p_FreeImage_SetTagKey@8" (ByVal Tag As Long, ByVal Key As String) As Long
Private Declare Function p_FreeImage_SetTagValue Lib "FreeImage.dll" Alias "_p_FreeImage_SetTagValue@8" (ByVal Tag As Long, ByVal ValuePtr As Long) As Long
' metadata iterators
Public Declare Function FreeImage_FindFirstMetadata Lib "FreeImage.dll" Alias "_FreeImage_FindFirstMetadata@12" (ByVal Model As FREE_IMAGE_MDMODEL, ByVal BITMAP As Long, ByRef Tag As Long) As Long
Public Declare Sub FreeImage_FindCloseMetadata Lib "FreeImage.dll" Alias "_FreeImage_FindCloseMetadata@4" (ByVal hFind As Long)
Private Declare Function p_FreeImage_FindNextMetadata Lib "FreeImage.dll" Alias "_FreeImage_FindNextMetadata@8" (ByVal hFind As Long, ByRef Tag As Long) As Long
' metadata setters and getters
Private Declare Function p_FreeImage_SetMetadata Lib "FreeImage.dll" Alias "_FreeImage_SetMetadata@16" (ByVal Model As Long, ByVal BITMAP As Long, ByVal Key As String, ByVal Tag As Long) As Long
Private Declare Function p_FreeImage_GetMetadata Lib "FreeImage.dll" Alias "_FreeImage_GetMetadata@16" (ByVal Model As Long, ByVal BITMAP As Long, ByVal Key As String, ByRef Tag As Long) As Long
Private Declare Function p_FreeImage_SetMetadataKeyValue Lib "FreeImage.dll" Alias "_FreeImage_SetMetadataKeyValue@16" (ByVal Model As Long, ByVal BITMAP As Long, ByVal Key As String, ByVal Tag As String) As Long
' metadata helper functions
Public Declare Function FreeImage_GetMetadataCount Lib "FreeImage.dll" Alias "_FreeImage_GetMetadataCount@8" (ByVal Model As Long, ByVal BITMAP As Long) As Long
Private Declare Function p_FreeImage_CloneMetadata Lib "FreeImage.dll" Alias "_FreeImage_CloneMetadata@8" (ByVal BitmapDst As Long, ByVal BitmapSrc As Long) As Long
' tag to string conversion functions
Private Declare Function p_FreeImage_TagToString Lib "FreeImage.dll" Alias "_FreeImage_TagToString@12" (ByVal Model As Long, ByVal Tag As Long, Optional ByVal Make As String = vbNullString) As Long
'----------------------
' JPEG lossless transformation functions
'----------------------
Private Declare Function p_FreeImage_JPEGTransform Lib "FreeImage.dll" Alias "_FreeImage_JPEGTransform@16" (ByVal SourceFile As String, ByVal DestFile As String, ByVal Operation As FREE_IMAGE_JPEG_OPERATION, ByVal Perfect As Long) As Long
Private Declare Function p_FreeImage_JPEGTransformU Lib "FreeImage.dll" Alias "_FreeImage_JPEGTransformU@16" (ByVal SourceFile As Long, ByVal DestFile As Long, ByVal Operation As FREE_IMAGE_JPEG_OPERATION, ByVal Perfect As Long) As Long
Private Declare Function p_FreeImage_JPEGCrop Lib "FreeImage.dll" Alias "_FreeImage_JPEGCrop@24" (ByVal SourceFile As String, ByVal DestFile As String, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Long
Private Declare Function p_FreeImage_JPEGCropU Lib "FreeImage.dll" Alias "_FreeImage_JPEGCropU@24" (ByVal SourceFile As Long, ByVal DestFile As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Long
Private Declare Function p_FreeImage_JPEGTransformCombined Lib "FreeImage.dll" Alias "_FreeImage_JPEGTransformCombined@32" (ByVal SourceFile As String, ByVal DestFile As String, ByVal Operation As FREE_IMAGE_JPEG_OPERATION, ByRef Left As Long, ByRef Top As Long, ByRef Right As Long, ByRef Bottom As Long, ByVal Perfect As Long) As Long
Private Declare Function p_FreeImage_JPEGTransformCombinedU Lib "FreeImage.dll" Alias "_FreeImage_JPEGTransformCombinedU@32" (ByVal SourceFile As Long, ByVal DestFile As Long, ByVal Operation As FREE_IMAGE_JPEG_OPERATION, ByRef Left As Long, ByRef Top As Long, ByRef Right As Long, ByRef Bottom As Long, ByVal Perfect As Long) As Long
Private Declare Function p_FreeImage_JPEGTransformCombinedFromMemory Lib "FreeImage.dll" Alias "_FreeImage_JPEGTransformCombinedFromMemory@32" (ByVal SourceStream As Long, ByVal DestStream As Long, ByVal Operation As FREE_IMAGE_JPEG_OPERATION, ByRef Left As Long, ByRef Top As Long, ByRef Right As Long, ByRef Bottom As Long, ByVal Perfect As Long) As Long
'----------------------
' Image manipulation toolkit functions
'----------------------
' rotation and flipping
Public Declare Function FreeImage_RotateClassic Lib "FreeImage.dll" Alias "_FreeImage_RotateClassic@12" (ByVal BITMAP As Long, ByVal Angle As Double) As Long
Public Declare Function FreeImage_Rotate Lib "FreeImage.dll" Alias "_FreeImage_Rotate@16" (ByVal BITMAP As Long, ByVal Angle As Double, Optional ByRef Color As Any = 0) As Long
Private Declare Function p_FreeImage_RotateEx Lib "FreeImage.dll" Alias "_FreeImage_RotateEx@48" (ByVal BITMAP As Long, ByVal Angle As Double, ByVal ShiftX As Double, ByVal ShiftY As Double, ByVal OriginX As Double, ByVal OriginY As Double, ByVal UseMask As Long) As Long
Private Declare Function p_FreeImage_FlipHorizontal Lib "FreeImage.dll" Alias "_FreeImage_FlipHorizontal@4" (ByVal BITMAP As Long) As Long
Private Declare Function p_FreeImage_FlipVertical Lib "FreeImage.dll" Alias "_FreeImage_FlipVertical@4" (ByVal BITMAP As Long) As Long
' upsampling / downsampling
Public Declare Function FreeImage_Rescale Lib "FreeImage.dll" Alias "_FreeImage_Rescale@16" (ByVal BITMAP As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_CATMULLROM) As Long
Public Declare Function FreeImage_RescaleRect Lib "FreeImage.dll" Alias "_FreeImage_RescaleRect@36" (ByVal BITMAP As Long, ByVal Width As Long, ByVal Height As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_CATMULLROM, Optional ByVal Flags As FREE_IMAGE_RESCALE_OPTIONS) As Long
Private Declare Function p_FreeImage_MakeThumbnail Lib "FreeImage.dll" Alias "_FreeImage_MakeThumbnail@12" (ByVal BITMAP As Long, ByVal MaxPixelSize As Long, Optional ByVal Convert As Long) As Long
' color manipulation functions (point operations)
Public Declare Function FreeImage_SwapPaletteIndices Lib "FreeImage.dll" Alias "_FreeImage_SwapPaletteIndices@12" (ByVal BITMAP As Long, ByRef IndexA As Byte, ByRef IndexB As Byte) As Long
Private Declare Function p_FreeImage_AdjustCurve Lib "FreeImage.dll" Alias "_FreeImage_AdjustCurve@12" (ByVal BITMAP As Long, ByVal LookupTablePtr As Long, ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As Long
Private Declare Function p_FreeImage_AdjustGamma Lib "FreeImage.dll" Alias "_FreeImage_AdjustGamma@12" (ByVal BITMAP As Long, ByVal gamma As Double) As Long
Private Declare Function p_FreeImage_AdjustBrightness Lib "FreeImage.dll" Alias "_FreeImage_AdjustBrightness@12" (ByVal BITMAP As Long, ByVal Percentage As Double) As Long
Private Declare Function p_FreeImage_AdjustContrast Lib "FreeImage.dll" Alias "_FreeImage_AdjustContrast@12" (ByVal BITMAP As Long, ByVal Percentage As Double) As Long
Private Declare Function p_FreeImage_Invert Lib "FreeImage.dll" Alias "_FreeImage_Invert@4" (ByVal BITMAP As Long) As Long
Private Declare Function p_FreeImage_GetHistogram Lib "FreeImage.dll" Alias "_FreeImage_GetHistogram@12" (ByVal BITMAP As Long, ByRef HistogramPtr As Long, Optional ByVal Channel As FREE_IMAGE_COLOR_CHANNEL = FICC_BLACK) As Long
Private Declare Function p_FreeImage_GetAdjustColorsLookupTable Lib "FreeImage.dll" Alias "_FreeImage_GetAdjustColorsLookupTable@32" (ByVal LookupTablePtr As Long, ByVal Brightness As Double, ByVal contrast As Double, ByVal gamma As Double, ByVal Invert As Long) As Long
Private Declare Function p_FreeImage_AdjustColors Lib "FreeImage.dll" Alias "_FreeImage_AdjustColors@32" (ByVal BITMAP As Long, ByVal Brightness As Double, ByVal contrast As Double, ByVal gamma As Double, ByVal Invert As Long) As Long
Private Declare Function p_FreeImage_ApplyColorMapping Lib "FreeImage.dll" Alias "_FreeImage_ApplyColorMapping@24" (ByVal BITMAP As Long, ByVal SourceColorsPtr As Long, ByVal DestinationColorsPtr As Long, ByVal Count As Long, ByVal IgnoreAlpha As Long, ByVal Swap As Long) As Long
Private Declare Function p_FreeImage_SwapColors Lib "FreeImage.dll" Alias "_FreeImage_SwapColors@16" (ByVal BITMAP As Long, ByRef ColorA As RGBQUAD, ByRef ColorB As RGBQUAD, ByVal IgnoreAlpha As Long) As Long
Private Declare Function p_FreeImage_SwapColorsByLong Lib "FreeImage.dll" Alias "_FreeImage_SwapColors@16" (ByVal BITMAP As Long, ByRef ColorA As Long, ByRef ColorB As Long, ByVal IgnoreAlpha As Long) As Long
Private Declare Function p_FreeImage_ApplyPaletteIndexMapping Lib "FreeImage.dll" Alias "_FreeImage_ApplyPaletteIndexMapping@20" (ByVal BITMAP As Long, ByVal SourceIndicesPtr As Long, ByVal DestinationIndicesPtr As Long, ByVal Count As Long, ByVal Swap As Long) As Long
' channel processing functions
Public Declare Function FreeImage_GetChannel Lib "FreeImage.dll" Alias "_FreeImage_GetChannel@8" (ByVal BITMAP As Long, ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As Long
Public Declare Function FreeImage_GetComplexChannel Lib "FreeImage.dll" Alias "_FreeImage_GetComplexChannel@8" (ByVal BITMAP As Long, ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As Long
Private Declare Function p_FreeImage_SetChannel Lib "FreeImage.dll" Alias "_FreeImage_SetChannel@12" (ByVal BitmapDst As Long, ByVal BitmapSrc As Long, ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As Long
Private Declare Function p_FreeImage_SetComplexChannel Lib "FreeImage.dll" Alias "_FreeImage_SetComplexChannel@12" (ByVal BitmapDst As Long, ByVal BitmapSrc As Long, ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As Long
' copy / paste / composite functions
Public Declare Function FreeImage_Copy Lib "FreeImage.dll" Alias "_FreeImage_Copy@20" (ByVal BITMAP As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Long
Public Declare Function FreeImage_CreateView Lib "FreeImage.dll" Alias "_FreeImage_CreateView@20" (ByVal BITMAP As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Long
Public Declare Function FreeImage_Composite Lib "FreeImage.dll" Alias "_FreeImage_Composite@16" (ByVal BITMAP As Long, Optional ByVal UseFileBackColor As Long, Optional ByRef AppBackColor As Any, Optional ByVal BackgroundBitmap As Long) As Long
Private Declare Function p_FreeImage_Paste Lib "FreeImage.dll" Alias "_FreeImage_Paste@20" (ByVal BitmapDst As Long, ByVal BitmapSrc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Alpha As Long) As Long
Private Declare Function p_FreeImage_PreMultiplyWithAlpha Lib "FreeImage.dll" Alias "_FreeImage_PreMultiplyWithAlpha@4" (ByVal BITMAP As Long) As Long
' background filling functions
Public Declare Function FreeImage_FillBackground Lib "FreeImage.dll" Alias "_FreeImage_FillBackground@12" (ByVal BITMAP As Long, ByRef Color As Any, Optional ByVal Options As FREE_IMAGE_COLOR_OPTIONS = FI_COLOR_IS_RGB_COLOR) As Long
Public Declare Function FreeImage_EnlargeCanvas Lib "FreeImage.dll" Alias "_FreeImage_EnlargeCanvas@28" (ByVal BITMAP As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByRef Color As Any, Optional ByVal Options As FREE_IMAGE_COLOR_OPTIONS = FI_COLOR_IS_RGB_COLOR) As Long
Public Declare Function FreeImage_AllocateEx Lib "FreeImage.dll" Alias "_FreeImage_AllocateEx@36" (ByVal Width As Long, ByVal Height As Long, Optional ByVal BitsPerPixel As Long = 8, Optional ByRef Color As Any, Optional ByVal Options As FREE_IMAGE_COLOR_OPTIONS, Optional ByVal PalettePtr As Long, Optional ByVal RedMask As Long = 0&, Optional ByVal GreenMask As Long = 0&, Optional ByVal BlueMask As Long = 0&) As Long
Public Declare Function FreeImage_AllocateExT Lib "FreeImage.dll" Alias "_FreeImage_AllocateExT@36" (ByVal ImageType As FREE_IMAGE_TYPE, ByVal Width As Long, ByVal Height As Long, Optional ByVal BitsPerPixel As Long = 8, Optional ByRef Color As Any, Optional ByVal Options As FREE_IMAGE_COLOR_OPTIONS, Optional ByVal PalettePtr As Long, Optional ByVal RedMask As Long = 0&, Optional ByVal GreenMask As Long = 0&, Optional ByVal BlueMask As Long = 0&) As Long
' miscellaneous algorithms
Public Declare Function FreeImage_MultigridPoissonSolver Lib "FreeImage.dll" Alias "_FreeImage_MultigridPoissonSolver@8" (ByVal LaplacianBitmap As Long, Optional ByVal Cyles As Long = 3) As Long
#End If                 ' <VBA7>
'----------------------
' Load FreeImage Library from relative path
'----------------------
#If VBA7 Then           ' <OFFICE2010+>
Private Declare PtrSafe Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
#Else                   ' <OFFICE97-2007>
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
#End If                 ' <VBA7>

Private Const Pi As Double = 3.14159265358979 '3.14159265358979

'----------------------
' Load/Unload library functions
'----------------------
Public Function FreeImage_LoadLibrary(Optional InitErrorHandler As Boolean)
Dim hMod As LongPtr: hMod = GetModuleHandle(c_strLibName): FreeImage_IsLoaded = (hMod <> 0): If FreeImage_IsLoaded Then Exit Function
Dim strPath As String: strPath = CurrentProject.path & c_strLibPath & c_strLibName: hMod = LoadLibrary(strPath): FreeImage_IsLoaded = (hMod <> 0)
If FreeImage_IsLoaded And InitErrorHandler Then FreeImage_InitErrorHandler
End Function
Public Function FreeImage_UnLoadLibrary()
Dim hMod As LongPtr: hMod = GetModuleHandle(c_strLibName): FreeImage_IsLoaded = (hMod <> 0): If FreeImage_IsLoaded Then FreeLibrary hMod
End Function
'----------------------
' Initialization functions
'----------------------
Public Function FreeImage_IsAvailable(Optional ByRef Version As String) As Boolean
   On Error Resume Next
   Version = FreeImage_GetVersion()
   FreeImage_IsAvailable = (Err.Number = NOERROR)
   On Error GoTo 0
End Function
'----------------------
' Error handling functions
'----------------------
Public Sub FreeImage_InitErrorHandler()
' Call this function once for using the FreeImage 3 error handling callback.
' The 'FreeImage_ErrorHandler' function is called on each FreeImage 3 error.
   Call FreeImage_SetOutputMessage(AddressOf FreeImage_ErrorHandler)
End Sub
Private Sub FreeImage_ErrorHandler(ByVal Format As FREE_IMAGE_FORMAT, ByVal Message As LongPtr)
Dim strErrorMessage As String
Dim strImageFormat As String
   strErrorMessage = p_GetStringFromPointerA(Message)
   strImageFormat = FreeImage_GetFormatFromFIF(Format)
   Debug.Print "[FreeImage] Error: " & strErrorMessage
   Debug.Print "            Image: " & strImageFormat
   Debug.Print "            Code:  " & Format
End Sub
'----------------------
' String returning functions wrappers
'----------------------
Public Function FreeImage_GetVersion() As String: FreeImage_GetVersion = p_GetStringFromPointerA(p_FreeImage_GetVersion): End Function
Public Function FreeImage_GetCopyrightMessage() As String: FreeImage_GetCopyrightMessage = p_GetStringFromPointerA(p_FreeImage_GetCopyrightMessage): End Function
Public Function FreeImage_GetFormatFromFIF(ByVal Format As FREE_IMAGE_FORMAT) As String: FreeImage_GetFormatFromFIF = p_GetStringFromPointerA(p_FreeImage_GetFormatFromFIF(Format)): End Function
Public Function FreeImage_GetFIFExtensionList(ByVal Format As FREE_IMAGE_FORMAT) As String: FreeImage_GetFIFExtensionList = p_GetStringFromPointerA(p_FreeImage_GetFIFExtensionList(Format)): End Function
Public Function FreeImage_GetFIFDescription(ByVal Format As FREE_IMAGE_FORMAT) As String: FreeImage_GetFIFDescription = p_GetStringFromPointerA(p_FreeImage_GetFIFDescription(Format)): End Function
Public Function FreeImage_GetFIFRegExpr(ByVal Format As FREE_IMAGE_FORMAT) As String: FreeImage_GetFIFRegExpr = p_GetStringFromPointerA(p_FreeImage_GetFIFRegExpr(Format)): End Function
Public Function FreeImage_GetFIFMimeType(ByVal Format As FREE_IMAGE_FORMAT) As String: FreeImage_GetFIFMimeType = p_GetStringFromPointerA(p_FreeImage_GetFIFMimeType(Format)): End Function
'----------------------
' UNICODE dealing functions wrappers
'----------------------
Public Function FreeImage_LoadU(ByVal Format As FREE_IMAGE_FORMAT, ByVal FileName As String, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As LongPtr: FreeImage_LoadU = p_FreeImage_LoadU(Format, StrPtr(FileName), Flags): End Function
Public Function FreeImage_SaveU(ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As LongPtr, ByVal FileName As String, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Boolean: FreeImage_SaveU = (p_FreeImage_SaveU(Format, BITMAP, StrPtr(FileName), Flags) = 1): End Function
Public Function FreeImage_GetFileTypeU(ByVal FileName As String, Optional ByVal Size As Long = 0) As FREE_IMAGE_FORMAT: FreeImage_GetFileTypeU = p_FreeImage_GetFileTypeU(StrPtr(FileName), Size): End Function
Public Function FreeImage_GetFIFFromFilenameU(ByVal FileName As String) As FREE_IMAGE_FORMAT: FreeImage_GetFIFFromFilenameU = p_FreeImage_GetFIFFromFilenameU(StrPtr(FileName)): End Function
Public Function FreeImage_GetFIFFromMemory(ByRef Data As Variant) As FREE_IMAGE_FORMAT
Dim hStream As LongPtr, lDataPtr As LongPtr
Dim Format As FREE_IMAGE_FORMAT: Format = FIF_UNKNOWN
Dim SizeInBytes As Long
   lDataPtr = p_GetMemoryBlockPtrFromVariant(Data, SizeInBytes)
   hStream = FreeImage_OpenMemoryByPtr(lDataPtr, SizeInBytes)
   If (hStream) Then
      Format = FreeImage_GetFileTypeFromMemory(hStream)
      Call FreeImage_CloseMemory(hStream)
   End If
   FreeImage_GetFIFFromMemory = Format
End Function
'----------------------
' Boolean returning functions wrappers
'----------------------
Public Function FreeImage_HasPixels(ByVal BITMAP As LongPtr) As Boolean: FreeImage_HasPixels = (FreeImage_HasPixelsLng(BITMAP) = 1): End Function
Public Function FreeImage_HasRGBMasks(ByVal BITMAP As LongPtr) As Boolean: FreeImage_HasRGBMasks = (p_FreeImage_HasRGBMasks(BITMAP) = 1): End Function
Public Function FreeImage_Save(ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As LongPtr, ByVal FileName As String, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Boolean: FreeImage_Save = (p_FreeImage_Save(Format, BITMAP, FileName, Flags) = 1): End Function
Public Function FreeImage_SaveToHandle(ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As LongPtr, ByVal IO As LongPtr, ByVal Handle As Long, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Boolean: FreeImage_SaveToHandle = (p_FreeImage_SaveToHandle(Format, BITMAP, IO, Handle, Flags) = 1): End Function
Public Function FreeImage_IsTransparent(ByVal BITMAP As LongPtr) As Boolean: FreeImage_IsTransparent = (p_FreeImage_IsTransparent(BITMAP) = 1): End Function
Public Sub FreeImage_SetTransparent(ByVal BITMAP As LongPtr, ByVal Value As Boolean): Call p_FreeImage_SetTransparent(BITMAP, IIf(Value, 1, 0)): End Sub
Public Function FreeImage_HasBackgroundColor(ByVal BITMAP As LongPtr) As Boolean: FreeImage_HasBackgroundColor = (p_FreeImage_HasBackgroundColor(BITMAP) = 1): End Function
Public Function FreeImage_GetBackgroundColor(ByVal BITMAP As LongPtr, ByRef BackColor As RGBQUAD) As Boolean: FreeImage_GetBackgroundColor = (p_FreeImage_GetBackgroundColor(BITMAP, BackColor) = 1): End Function
Public Function FreeImage_GetBackgroundColorAsLong(ByVal BITMAP As LongPtr, ByRef BackColor As Long) As Boolean: FreeImage_GetBackgroundColorAsLong = (p_FreeImage_GetBackgroundColorAsLong(BITMAP, BackColor) = 1): End Function
Public Function FreeImage_GetBackgroundColorEx(ByVal BITMAP As LongPtr, ByRef Alpha As Byte, ByRef Red As Byte, ByRef Green As Byte, ByRef Blue As Byte) As Boolean
' gets the background color of an image as FreeImage_GetBackgroundColor() does but provides it's result as four different byte values, one for each color component.
Dim bkColor As RGBQUAD
   FreeImage_GetBackgroundColorEx = (p_FreeImage_GetBackgroundColor(BITMAP, bkColor) = 1)
   With bkColor
      Alpha = .rgbReserved
      Red = .rgbRed
      Green = .rgbGreen
      Blue = .rgbBlue
   End With
End Function
Public Function FreeImage_SetBackgroundColor(ByVal BITMAP As LongPtr, ByRef BackColor As RGBQUAD) As Boolean: FreeImage_SetBackgroundColor = (p_FreeImage_SetBackgroundColor(BITMAP, BackColor) = 1): End Function
Public Function FreeImage_SetBackgroundColorAsLong(ByVal BITMAP As LongPtr, ByVal BackColor As Long) As Boolean: FreeImage_SetBackgroundColorAsLong = (p_FreeImage_SetBackgroundColorAsLong(BITMAP, BackColor) = 1): End Function
Public Function FreeImage_SetBackgroundColorEx(ByVal BITMAP As LongPtr, ByVal Alpha As Byte, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte) As Boolean
' sets the color at position (x|y) as FreeImage_SetPixelColor() does but the color value to set must be provided four different byte values, one for each color component.
Dim tColor As RGBQUAD: With tColor: .rgbReserved = Alpha: .rgbRed = Red: .rgbGreen = Green: .rgbBlue = Blue: End With
   FreeImage_SetBackgroundColorEx = (p_FreeImage_SetBackgroundColor(BITMAP, tColor) = 1)
End Function
Public Function FreeImage_GetPixelIndex(ByVal BITMAP As LongPtr, ByVal x As Long, ByVal y As Long, ByRef Value As Byte) As Boolean: FreeImage_GetPixelIndex = (p_FreeImage_GetPixelIndex(BITMAP, x, y, Value) = 1): End Function
Public Function FreeImage_GetPixelColor(ByVal BITMAP As LongPtr, ByVal x As Long, ByVal y As Long, ByRef Value As RGBQUAD) As Boolean: FreeImage_GetPixelColor = (p_FreeImage_GetPixelColor(BITMAP, x, y, Value) = 1): End Function
Public Function FreeImage_GetPixelColorByLong(ByVal BITMAP As LongPtr, ByVal x As Long, ByVal y As Long, ByRef Value As Long) As Boolean: FreeImage_GetPixelColorByLong = (p_FreeImage_GetPixelColorByLong(BITMAP, x, y, Value) = 1): End Function
Public Function FreeImage_GetPixelColorEx(ByVal BITMAP As LongPtr, ByVal x As Long, ByVal y As Long, ByRef Alpha As Byte, ByRef Red As Byte, ByRef Green As Byte, ByRef Blue As Byte) As Boolean
' gets the color at position (x|y) as FreeImage_GetPixelColor() does but provides it's result as four different byte values, one for each color component.
Dim Value As RGBQUAD
   FreeImage_GetPixelColorEx = (p_FreeImage_GetPixelColor(BITMAP, x, y, Value) = 1)
   With Value
      Alpha = .rgbReserved
      Red = .rgbRed
      Green = .rgbGreen
      Blue = .rgbBlue
   End With
End Function
Public Function FreeImage_SetPixelIndex(ByVal BITMAP As LongPtr, ByVal x As Long, ByVal y As Long, ByRef Value As Byte) As Boolean: FreeImage_SetPixelIndex = (p_FreeImage_SetPixelIndex(BITMAP, x, y, Value) = 1): End Function
Public Function FreeImage_SetPixelColor(ByVal BITMAP As LongPtr, ByVal x As Long, ByVal y As Long, ByRef Value As RGBQUAD) As Boolean: FreeImage_SetPixelColor = (p_FreeImage_SetPixelColor(BITMAP, x, y, Value) = 1): End Function
Public Function FreeImage_SetPixelColorByLong(ByVal BITMAP As LongPtr, ByVal x As Long, ByVal y As Long, ByRef Value As Long) As Boolean: FreeImage_SetPixelColorByLong = (p_FreeImage_SetPixelColorByLong(BITMAP, x, y, Value) = 1): End Function
Public Function FreeImage_SetPixelColorEx(ByVal BITMAP As LongPtr, ByVal x As Long, ByVal y As Long, ByVal Alpha As Byte, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte) As Boolean
' sets the color at position (x|y) as FreeImage_SetPixelColor() does but the color value to set must be provided four different byte values, one for each color component.
Dim Value As RGBQUAD
   With Value
      .rgbReserved = Alpha
      .rgbRed = Red
      .rgbGreen = Green
      .rgbBlue = Blue
   End With
   FreeImage_SetPixelColorEx = (p_FreeImage_SetPixelColor(BITMAP, x, y, Value) = 1)
End Function
Public Function FreeImage_SaveToMemory(ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As LongPtr, ByVal Stream As LongPtr, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS) As Boolean: FreeImage_SaveToMemory = (p_FreeImage_SaveToMemory(Format, BITMAP, Stream, Flags) = 1): End Function
Public Function FreeImage_AcquireMemory(ByVal Stream As LongPtr, ByRef DataPtr As LongPtr, ByRef SizeInBytes As Long) As Boolean: FreeImage_AcquireMemory = (p_FreeImage_AcquireMemory(Stream, DataPtr, SizeInBytes) = 1): End Function
Public Function FreeImage_SeekMemory(ByVal Stream As LongPtr, ByVal Offset As Long, ByVal Origin As Long) As Boolean:   FreeImage_SeekMemory = (p_FreeImage_SeekMemory(Stream, Offset, Origin) = 1): End Function
Public Function FreeImage_FlipHorizontal(ByVal BITMAP As LongPtr) As Boolean:   FreeImage_FlipHorizontal = (p_FreeImage_FlipHorizontal(BITMAP) = 1): End Function
Public Function FreeImage_FlipVertical(ByVal BITMAP As LongPtr) As Boolean:   FreeImage_FlipVertical = (p_FreeImage_FlipVertical(BITMAP) = 1): End Function
Public Function FreeImage_AdjustCurve(ByVal BITMAP As LongPtr, ByVal LookupTablePtr As LongPtr, ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As Boolean:   FreeImage_AdjustCurve = (p_FreeImage_AdjustCurve(BITMAP, LookupTablePtr, Channel) = 1): End Function
Public Function FreeImage_AdjustGamma(ByVal BITMAP As LongPtr, ByVal gamma As Double) As Boolean:   FreeImage_AdjustGamma = (p_FreeImage_AdjustGamma(BITMAP, gamma) = 1): End Function
Public Function FreeImage_AdjustBrightness(ByVal BITMAP As LongPtr, ByVal Percentage As Double) As Boolean:   FreeImage_AdjustBrightness = (p_FreeImage_AdjustBrightness(BITMAP, Percentage) = 1): End Function
Public Function FreeImage_AdjustContrast(ByVal BITMAP As LongPtr, ByVal Percentage As Double) As Boolean:   FreeImage_AdjustContrast = (p_FreeImage_AdjustContrast(BITMAP, Percentage) = 1): End Function
Public Function FreeImage_Invert(ByVal BITMAP As LongPtr) As Boolean:   FreeImage_Invert = (p_FreeImage_Invert(BITMAP) = 1): End Function
Public Function FreeImage_GetHistogram(ByVal BITMAP As LongPtr, ByRef HistogramPtr As LongPtr, Optional ByVal Channel As FREE_IMAGE_COLOR_CHANNEL = FICC_BLACK) As Boolean:   FreeImage_GetHistogram = (p_FreeImage_GetHistogram(BITMAP, HistogramPtr, Channel) = 1): End Function
Public Function FreeImage_AdjustColors(ByVal BITMAP As LongPtr, Optional ByVal Brightness As Double, Optional ByVal contrast As Double, Optional ByVal gamma As Double = 1, Optional ByVal Invert As Boolean) As Boolean: Dim lInvert As Long: FreeImage_AdjustColors = (p_FreeImage_AdjustColors(BITMAP, Brightness, contrast, gamma, IIf(Invert, 1, 0)) = 1): End Function
Public Function FreeImage_SetChannel(ByVal BitmapDst As LongPtr, ByVal BitmapSrc As LongPtr, ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As Boolean: FreeImage_SetChannel = (p_FreeImage_SetChannel(BitmapDst, BitmapSrc, Channel) = 1): End Function
Public Function FreeImage_SetComplexChannel(ByVal BitmapDst As LongPtr, ByVal BitmapSrc As LongPtr, ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As Boolean: FreeImage_SetComplexChannel = (p_FreeImage_SetComplexChannel(BitmapDst, BitmapSrc, Channel) = 1): End Function
Public Function FreeImage_PreMultiplyWithAlpha(ByVal BITMAP As LongPtr) As Boolean:   FreeImage_PreMultiplyWithAlpha = (p_FreeImage_PreMultiplyWithAlpha(BITMAP) = 1): End Function
Public Function FreeImage_FillBackgroundEx(ByVal BITMAP As LongPtr, ByRef Color As RGBQUAD, Optional ByVal Options As FREE_IMAGE_COLOR_OPTIONS) As Boolean: FreeImage_FillBackgroundEx = (FreeImage_FillBackground(BITMAP, Color, Options) = 1): End Function
Public Function FreeImage_FillBackgroundByLong(ByVal BITMAP As LongPtr, ByRef Color As Long, Optional ByVal Options As FREE_IMAGE_COLOR_OPTIONS) As Boolean: FreeImage_FillBackgroundByLong = (FreeImage_FillBackground(BITMAP, Color, Options) = 1): End Function
Public Function FreeImage_SetThumbnail(ByVal BITMAP As LongPtr, ByVal Thumbnail As Long) As Boolean:   FreeImage_SetThumbnail = (p_FreeImage_SetThumbnail(BITMAP, Thumbnail) = 1): End Function
Public Function FreeImage_OpenMultiBitmap(ByVal Format As FREE_IMAGE_FORMAT, ByVal FileName As String, Optional ByVal CreateNew As Boolean, Optional ByVal ReadOnly As Boolean, Optional ByVal KeepCacheInMemory As Boolean, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS) As LongPtr: FreeImage_OpenMultiBitmap = p_FreeImage_OpenMultiBitmap(Format, FileName, IIf(CreateNew, 1, 0), IIf(ReadOnly And Not CreateNew, 1, 0), IIf(KeepCacheInMemory, 1, 0), Flags): End Function
Public Sub FreeImage_UnlockPage(ByVal BITMAP As LongPtr, ByVal PageBitmap As LongPtr, ByVal ApplyChanges As Boolean): Call p_FreeImage_UnlockPage(BITMAP, PageBitmap, IIf(ApplyChanges, 1, 0)): End Sub
Public Function FreeImage_MakeThumbnail(ByVal BITMAP As LongPtr, ByVal MaxPixelSize As Long, Optional ByVal Convert As Boolean) As LongPtr: FreeImage_MakeThumbnail = p_FreeImage_MakeThumbnail(BITMAP, MaxPixelSize, IIf(Convert, 1, 0)): End Function
Public Function FreeImage_GetAdjustColorsLookupTable(ByVal LookupTablePtr As LongPtr, Optional ByVal Brightness As Double, Optional ByVal contrast As Double, Optional ByVal gamma As Double, Optional ByVal Invert As Boolean) As Long: FreeImage_GetAdjustColorsLookupTable = p_FreeImage_GetAdjustColorsLookupTable(LookupTablePtr, Brightness, contrast, gamma, IIf(Invert, 1, 0)): End Function
Public Function FreeImage_ApplyColorMapping(ByVal BITMAP As LongPtr, ByVal SourceColorsPtr As LongPtr, ByVal DestinationColorsPtr As LongPtr, ByVal Count As Long, Optional ByVal IgnoreAlpha As Boolean = True, Optional ByVal Swap As Boolean) As Long: FreeImage_ApplyColorMapping = p_FreeImage_ApplyColorMapping(BITMAP, SourceColorsPtr, DestinationColorsPtr, Count, IIf(IgnoreAlpha, 1, 0), IIf(Swap, 1, 0)): End Function
Public Function FreeImage_SwapColors(ByVal BITMAP As LongPtr, ByRef ColorA As RGBQUAD, ByRef ColorB As RGBQUAD, Optional ByVal IgnoreAlpha As Boolean = True) As Long: FreeImage_SwapColors = p_FreeImage_SwapColors(BITMAP, ColorA, ColorB, IIf(IgnoreAlpha, 1, 0)): End Function
Public Function FreeImage_SwapColorsByLong(ByVal BITMAP As LongPtr, ByRef ColorA As Long, ByRef ColorB As Long, Optional ByVal IgnoreAlpha As Boolean = True) As Long: FreeImage_SwapColorsByLong = p_FreeImage_SwapColorsByLong(BITMAP, ColorA, ColorB, IIf(IgnoreAlpha, 1, 0)): End Function
Public Function FreeImage_ConvertFromRawBits(ByVal BitsPtr As LongPtr, ByVal Width As Long, ByVal Height As Long, ByVal Pitch As Long, ByVal BitsPerPixel As Long, Optional ByVal RedMask As Long, Optional ByVal GreenMask As Long, Optional ByVal BlueMask As Long, Optional ByVal TopDown As Boolean) As LongPtr: FreeImage_ConvertFromRawBits = p_FreeImage_ConvertFromRawBits(BitsPtr, Width, Height, Pitch, BitsPerPixel, RedMask, GreenMask, BlueMask, IIf(TopDown, 1, 0)): End Function
Public Function FreeImage_ConvertFromRawBitsEx(ByVal CopySource As Boolean, ByVal BitsPtr As LongPtr, ByVal ImageType As FREE_IMAGE_TYPE, ByVal Width As Long, ByVal Height As Long, ByVal Pitch As Long, ByVal BitsPerPixel As Long, Optional ByVal RedMask As Long, Optional ByVal GreenMask As Long, Optional ByVal BlueMask As Long, Optional ByVal TopDown As Boolean) As LongPtr: FreeImage_ConvertFromRawBitsEx = p_FreeImage_ConvertFromRawBitsEx(IIf(CopySource, 1, 0), BitsPtr, ImageType, Width, Height, Pitch, BitsPerPixel, RedMask, GreenMask, BlueMask, IIf(TopDown, 1, 0)): End Function
Public Sub FreeImage_ConvertToRawBits(ByVal BitsPtr As LongPtr, ByVal BITMAP As LongPtr, ByVal Pitch As Long, ByVal BitsPerPixel As Long, Optional ByVal RedMask As Long, Optional ByVal GreenMask As Long, Optional ByVal BlueMask As Long, Optional ByVal TopDown As Boolean): Call p_FreeImage_ConvertToRawBits(BitsPtr, BITMAP, Pitch, BitsPerPixel, RedMask, GreenMask, BlueMask, IIf(TopDown, 1, 0)): End Sub
Public Function FreeImage_ConvertToStandardType(ByVal BITMAP As LongPtr, Optional ByVal ScaleLinear As Boolean = True) As LongPtr: FreeImage_ConvertToStandardType = p_FreeImage_ConvertToStandardType(BITMAP, IIf(ScaleLinear, 1, 0)): End Function
Public Function FreeImage_ConvertToType(ByVal BITMAP As LongPtr, ByVal DestinationType As FREE_IMAGE_TYPE, Optional ByVal ScaleLinear As Boolean = True) As LongPtr: FreeImage_ConvertToType = p_FreeImage_ConvertToType(BITMAP, DestinationType, IIf(ScaleLinear, 1, 0)): End Function
Public Function FreeImage_Paste(ByVal BitmapDst As LongPtr, ByVal BitmapSrc As LongPtr, ByVal Left As Long, ByVal Top As Long, ByVal Alpha As Long) As Boolean: FreeImage_Paste = (p_FreeImage_Paste(BitmapDst, BitmapSrc, Left, Top, Alpha) = 1): End Function
Public Function FreeImage_RotateEx(ByVal BITMAP As LongPtr, ByVal Angle As Double, Optional ByVal ShiftX As Double, Optional ByVal ShiftY As Double, Optional ByVal OriginX As Double, Optional ByVal OriginY As Double, Optional ByVal UseMask As Boolean) As LongPtr:   FreeImage_RotateEx = p_FreeImage_RotateEx(BITMAP, Angle, ShiftX, ShiftY, OriginX, OriginY, IIf(UseMask, 1, 0)): End Function

'----------------------
' Color conversion helper functions
'----------------------
Public Function ConvertColor(ByVal Color As Long) As Long
' This helper function converts a VB-style color value (like vbRed), which
' uses the ABGR format into a RGBQUAD compatible color value, using the ARGB
' format, needed by FreeImage and vice versa.
   ConvertColor = ((Color And &HFF000000) Or ((Color And &HFF&) * &H10000) Or ((Color And &HFF00&)) Or ((Color And &HFF0000) \ &H10000))
End Function
Public Function ConvertOleColor(ByVal Color As OLE_COLOR) As Long
' This helper function converts an OLE_COLOR value (like vbButtonFace), which uses the BGR format into a RGBQUAD compatible color value, using the ARGB format, needed by FreeImage.
' generally ingnores the specified color's alpha value but, in contrast to ConvertColor, also has support for system colors, which have the format &H80bbggrr.
' You should not use this function to convert any color provided by FreeImage
' in ARGB format into a VB-style ABGR color value. Use function ConvertColor instead.
Dim lColorRef As Long: If (OleTranslateColor(Color, 0, lColorRef) = 0) Then ConvertOleColor = ConvertColor(lColorRef)
End Function
'----------------------
' Extended functions derived from FreeImage 3 functions usually dealing with arrays
'----------------------
Public Sub FreeImage_UnloadEx(ByRef BITMAP As LongPtr)
' Extended version of FreeImage_Unload, which additionally sets the passed Bitmap handle to zero after unloading.
   If (BITMAP <> 0) Then Call FreeImage_Unload(BITMAP): BITMAP = 0
End Sub
Public Function FreeImage_GetPaletteEx(ByVal BITMAP As LongPtr) As RGBQUAD()
' returns a VB style array of type RGBQUAD, containing
' the palette data of the Bitmap. This array provides read and write access
' to the actual palette data provided by FreeImage. This is done by
' creating a VB array with an own SAFEARRAY descriptor making the
' array point to the palette pointer returned by FreeImage_GetPalette().
' This makes you use code like you would in C/C++:
' // this code assumes there is a bitmap loaded and
' // present in a variable called 'dib'
' if(FreeImage_GetBPP(Bitmap) == 8) {
'   // Build a greyscale palette
'   RGBQUAD *pal = FreeImage_GetPalette(Bitmap);
'   for (int i = 0; i < 256; i++) {
'     pal[i].rgbRed = i;
'     pal[i].rgbGreen = i;
'     pal[i].rgbBlue = i;
'   }
' As in C/C++ the array is only valid while the DIB is loaded and the
' palette data remains where the pointer returned by FreeImage_GetPalette
' has pointed to when this function was called. So, a good thing would
' be, not to keep the returned array in scope over the lifetime of the
' Bitmap. Best practise is, to use this function within another routine and
' assign the return value (the array) to a local variable only. As soon
' as this local variable goes out of scope (when the calling function
' returns to it's caller), the array and the descriptor is automatically
' cleaned up by VB.
' does not make a deep copy of the palette data, but only
' wraps a VB array around the FreeImage palette data. So, it can be called
' frequently "on demand" or somewhat "in place" without a significant
' performance loss.
' To learn more about this technique I recommend reading chapter 2 (Leveraging
' Arrays) of Matthew Curland's book "Advanced Visual Basic 6"

' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the FreeImage_DestroyLockedArrayRGBQUAD() function.
   
    If (BITMAP = 0) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
       .cbElements = 4                              ' size in bytes of RGBQUAD structure
       .cDims = 1                                   ' the array has only 1 dimension
       .cElements = FreeImage_GetColorsUsed(BITMAP) ' the number of elements in the array is the number of used colors in the Bitmap
       .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE     ' need AUTO and FIXEDSIZE for safety issues,so the array can not be modified in sizeor erased; according to Matthew Curland never use FIXEDSIZE alone
       .pvData = FreeImage_GetPalette(BITMAP)       ' let the array point to the memory block, the FreeImage palette pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetPaletteEx), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetPaletteExClone(ByVal BITMAP As LongPtr) As RGBQUAD()
' returns a redundant clone of a Bitmap's palette as a VB style array of type RGBQUAD.
Dim lColors As Long: lColors = FreeImage_GetColorsUsed(BITMAP):   If (lColors <= 0) Then Exit Function
Dim atPal() As RGBQUAD: ReDim atPal(lColors - 1)
    Call CopyMemory(atPal(0), ByVal FreeImage_GetPalette(BITMAP), lColors * 4)
    Call p_Swap(ByVal VarPtrArray(atPal), ByVal VarPtrArray(FreeImage_GetPaletteExClone))
End Function
Public Function FreeImage_GetPaletteExLong(ByVal BITMAP As LongPtr) As Long()
' returns a VB style array of type Long, containing the palette data of the Bitmap.
' This array provides read and write access to the actual palette data provided by FreeImage.
' This is done by creating a VB array with an own SAFEARRAY descriptor
' making the array point to the palette pointer returned by FreeImage_GetPalette().
' The function actually returns an array of type RGBQUAD with each element packed into a Long.
' This is possible, since the RGBQUAD structure is also four bytes in size.
' Palette data, stored in an array of type Long may be passed ByRef to a function through an optional paremeter.
' For an example have a look at function FreeImage_ConvertColorDepth()
' This makes you use code like you would in C/C++:
' // this code assumes there is a bitmap loaded and
' // present in a variable called 'dib'
' if(FreeImage_GetBPP(Bitmap) == 8) {
'   // Build a greyscale palette
'   RGBQUAD *pal = FreeImage_GetPalette(Bitmap);
'   for (int i = 0; i < 256; i++) {
'     pal[i].rgbRed = i;
'     pal[i].rgbGreen = i;
'     pal[i].rgbBlue = i;
'   }
' As in C/C++ the array is only valid while the DIB is loaded and the palette data remains where the pointer returned by FreeImage_GetPalette()
' has pointed to when this function was called. So, a good thing would be, not to keep the returned array in scope over the lifetime of the Bitmap.
' Best practise is, to use this function within another routine and assign the return value (the array) to a local variable only.
' As soon as this local variable goes out of scope (when the calling function returns to it's caller), the array and the descriptor is automatically cleaned up by VB.
' does not make a deep copy of the palette data, but only wraps a VB array around the FreeImage palette data.
' So, it can be called frequently "on demand" or somewhat "in place" without a significant performance loss.
' To learn more about this technique I recommend reading chapter 2 (Leveraging Arrays) of Matthew Curland's book "Advanced Visual Basic 6"

' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArray' function.
    If (BITMAP = 0) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
       .cbElements = 4                              ' size in bytes of RGBQUAD structure
       .cDims = 1                                   ' the array has only 1 dimension
       .cElements = FreeImage_GetColorsUsed(BITMAP) ' the number of elements in the array is the number of used colors in the Bitmap
       .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE     ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
       .pvData = FreeImage_GetPalette(BITMAP)       ' let the array point to the memory block, the FreeImage palette pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetPaletteExLong), lpSA, PTR_LENGTH)  '4)
End Function
Public Function FreeImage_GetPaletteExLongClone(ByVal BITMAP As LongPtr) As Long()
' returns a redundant clone of a Bitmap's palette as a' VB style array of type Long.
Dim lColors As Long: lColors = FreeImage_GetColorsUsed(BITMAP): If (lColors <= 0) Then Exit Function
Dim alPal() As Long: ReDim alPal(lColors - 1)
    Call CopyMemory(alPal(0), ByVal FreeImage_GetPalette(BITMAP), lColors * 4)
    Call p_Swap(ByVal VarPtrArray(alPal), ByVal VarPtrArray(FreeImage_GetPaletteExLongClone))
End Function
Public Function FreeImage_SetPalette(ByVal BITMAP As LongPtr, ByRef palette() As RGBQUAD) As Long
' sets the palette of a palletised bitmap using a RGBQUAD array. Does nothing on high color bitmaps.
' This operation makes a deep copy of the provided palette data so, after this function
' has returned, changes to the RGBQUAD array are no longer reflected by the bitmap's palette.
   FreeImage_SetPalette = FreeImage_GetColorsUsed(BITMAP)
   If (FreeImage_SetPalette > 0) Then Call CopyMemory(ByVal FreeImage_GetPalette(BITMAP), palette(0), FreeImage_SetPalette * 4)
End Function
Public Function FreeImage_SetPaletteLong(ByVal BITMAP As LongPtr, ByRef palette() As Long) As Long
' sets the palette of a palletised bitmap using a RGBQUAD array. Does nothing on high color bitmaps.
' This operation makes a deep copy of the provided palette data so, after this function
' has returned, changes to the Long array are no longer reflected by the bitmap's palette.
   FreeImage_SetPaletteLong = FreeImage_GetColorsUsed(BITMAP)
   If (FreeImage_SetPaletteLong > 0) Then Call CopyMemory(ByVal FreeImage_GetPalette(BITMAP), palette(0), FreeImage_SetPaletteLong * 4)
End Function
Public Function FreeImage_GetTransparencyTableEx(ByVal BITMAP As LongPtr) As Byte()
' returns a VB style Byte array, containing the transparency
' table of the Bitmap. This array provides read and write access to the actual
' transparency table provided by FreeImage. This is done by creating a VB array
' with an own SAFEARRAY descriptor making the array point to the transparency
' table pointer returned by FreeImage_GetTransparencyTable().
' This makes you use code like you would in C/C++:
' // this code assumes there is a bitmap loaded and
' // present in a variable called 'dib'
' if(FreeImage_GetBPP(Bitmap) == 8) {
'   // Remove transparency information
'   byte *transt = FreeImage_GetTransparencyTable(Bitmap);
'   for (int i = 0; i < 256; i++) {
'     transt[i].rgbRed = 255;
'   }
' As in C/C++ the array is only valid while the DIB is loaded and the transparency
' table remains where the pointer returned by FreeImage_GetTransparencyTable() has
' pointed to when this function was called. So, a good thing would be, not to keep
' the returned array in scope over the lifetime of the DIB. Best practise is, to use
' within another routine and assign the return value (the array) to a
' local variable only. As soon as this local variable goes out of scope (when the
' calling function returns to it's caller), the array and the descriptor is
' automatically cleaned up by VB.
' does not make a deep copy of the transparency table, but only
' wraps a VB array around the FreeImage transparency table. So, it can be called
' frequently "on demand" or somewhat "in place" without a significant
' performance loss.
' To learn more about this technique I recommend reading chapter 2 (Leveraging
' Arrays) of Matthew Curland's book "Advanced Visual Basic 6"

' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the FreeImage_DestroyLockedArray() function.
   
    If (BITMAP = 0) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
       .cDims = 1                                          ' the array has only 1 dimension
       .cbElements = 1                                     ' size in bytes of a byte element
       .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE            ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
       .pvData = FreeImage_GetTransparencyTable(BITMAP)    ' let the array point to the memory block, the FreeImage transparency table pointer points to
       .cElements = FreeImage_GetTransparencyCount(BITMAP) ' the number of elements in the array is equal to the number transparency table entries
If .cElements > 0 Then Stop
    ' When the bitmap is not palletised, FreeImage_GetTransparencyCount always returns 0
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, LenB(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetTransparencyTableEx), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetTransparencyTableExClone(ByVal BITMAP As LongPtr) As Byte()
' returns a copy of a DIB's transparency table as VB style array of type Byte.
' So, the array provides read access only from the DIB's point of view.
Dim lpTransparencyTable As LongPtr: lpTransparencyTable = FreeImage_GetTransparencyTable(BITMAP): If (lpTransparencyTable = 0) Then Exit Function
Dim lEntries As Long: lEntries = FreeImage_GetTransparencyCount(BITMAP): If (lEntries <= 0) Then Exit Function
Dim abBuffer() As Byte: ReDim abBuffer(lEntries - 1)
    Call CopyMemory(abBuffer(0), ByVal lpTransparencyTable, lEntries)
    Call p_Swap(ByVal VarPtrArray(abBuffer), ByVal VarPtrArray(FreeImage_GetTransparencyTableExClone))
End Function
Public Sub FreeImage_SetTransparencyTableEx(ByVal BITMAP As LongPtr, ByRef Table() As Byte, Optional ByRef Count As Long = -1)
' sets a DIB's transparency table to the contents of the parameter table().
' When the optional parameter Count is omitted, the number of entries used is taken
' from the number of elements stored in the array, but will never be never greater than 256.
   If ((Count > UBound(Table) + 1) Or (Count < 0)) Then Count = UBound(Table) + 1
   If (Count > 256) Then Count = 256
   Call FreeImage_SetTransparencyTable(BITMAP, VarPtr(Table(0)), Count)
End Sub
Public Function FreeImage_IsTransparencyTableTransparent(ByVal BITMAP As LongPtr) As Boolean
' checks whether a Bitmap's transparency table contains any transparent colors or not.
' When an image has a transparency table and is transparent, what can be tested
' with 'FreeImage_IsTransparent()', the image still may display opaque when there
' are no transparent colors defined in the image's transparency table. This
' function reads the Bitmap's transparency table directly to determine whether
' there are transparent colors defined or not.
' The return value of this function does not relay on the image's transparency
' setting but only on the image's transparency table
    If (BITMAP = 0) Then Exit Function
Dim abTransTable() As Byte: abTransTable = FreeImage_GetTransparencyTableEx(BITMAP)
Dim i As Long
    For i = 0 To UBound(abTransTable)
       FreeImage_IsTransparencyTableTransparent = (abTransTable(i) < 255)
       If (FreeImage_IsTransparencyTableTransparent) Then Exit For
    Next i
End Function
Public Function FreeImage_GetAdjustColorsLookupTableEx(ByRef LookupTable() As Byte, Optional ByVal Brightness As Double, Optional ByVal contrast As Double, Optional ByVal gamma As Double = 1, Optional ByVal Invert As Boolean) As Long
' is an extended wrapper for FreeImage_GetAdjustColorsLookupTable(), which
' takes a real VB style Byte array LUT() to receive the created lookup table. The LUT()
' parameter must not be fixed sized or locked, since it is (re-)dimensioned in this
' function to contain 256 entries.
   ReDim LookupTable(255)
   FreeImage_GetAdjustColorsLookupTableEx = FreeImage_GetAdjustColorsLookupTable(VarPtr(LookupTable(0)), Brightness, contrast, gamma, Invert)
End Function
Public Function FreeImage_ApplyColorMappingEx(ByVal BITMAP As LongPtr, ByRef SourceColors() As RGBQUAD, ByRef DestinationColors() As RGBQUAD, Optional ByRef Count As Long = -1, Optional ByVal IgnoreAlpha As Boolean = True, Optional ByVal Swap As Boolean) As Long
' is an extended wrapper for FreeImage_ApplyColorMapping(), which takes real VB style RGBQUAD arrays for source and destination colors along with an optional ByRef Count parameter.
' If 'Count' is omitted upon entry, the number of entries of the smaller of both arrays
' is used for 'Count' and also passed back to the caller, due to this parameter's ByRef nature.
    If (BITMAP = 0) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to map colors on a 'header-only' bitmap.")
Dim nsrc As Long: nsrc = UBound(SourceColors) + 1
Dim ndst As Long: ndst = UBound(DestinationColors) + 1
      If (Count = -1) Then
         If (nsrc < ndst) Then Count = nsrc Else Count = ndst
      Else
         If (Count < nsrc) Then Count = nsrc
         If (Count < ndst) Then Count = ndst
      End If
      FreeImage_ApplyColorMappingEx = FreeImage_ApplyColorMapping(BITMAP, VarPtr(SourceColors(0)), VarPtr(DestinationColors(0)), Count, IgnoreAlpha, Swap)
End Function
Public Function FreeImage_ApplyPaletteIndexMappingEx(ByVal BITMAP As LongPtr, ByRef SourceIndices() As Byte, ByRef DestinationIndices() As Byte, Optional ByRef Count As Long = -1, Optional ByVal Swap As Boolean) As Long
' is an extended wrapper for FreeImage_ApplyIndexMapping(), which takes real VB style Byte arrays for source and destination indices along with an optional ByRef count parameter.
' If 'Count' is omitted upon entry, the number of entries of the smaller of both arrays
' is used for 'Count' and also passed back to the caller, due to this parameter's ByRef nature.

Dim nsrc As Long: nsrc = UBound(SourceIndices) + 1
Dim ndst As Long: ndst = UBound(DestinationIndices) + 1
   If (Count = -1) Then
      If (nsrc < ndst) Then Count = nsrc Else Count = ndst
   Else
      If (Count < nsrc) Then Count = nsrc
      If (Count < ndst) Then Count = ndst
   End If
Dim lSwap As Long: If (Swap) Then lSwap = 1
   FreeImage_ApplyPaletteIndexMappingEx = p_FreeImage_ApplyPaletteIndexMapping(BITMAP, VarPtr(SourceIndices(0)), VarPtr(DestinationIndices(0)), Count, lSwap)
End Function
Public Function FreeImage_ConvertFromRawBitsArray(ByRef Bits() As Byte, ByVal Width As Long, ByVal Height As Long, ByVal Pitch As Long, ByVal BitsPerPixel As Long, Optional ByVal RedMask As Long, Optional ByVal GreenMask As Long, Optional ByVal BlueMask As Long, Optional ByVal TopDown As Boolean) As LongPtr
   FreeImage_ConvertFromRawBitsArray = FreeImage_ConvertFromRawBits(VarPtr(Bits(0)), Width, Height, Pitch, BitsPerPixel, RedMask, GreenMask, BlueMask, TopDown)
End Function
Public Sub FreeImage_ConvertToRawBitsArray(ByRef Bits() As Byte, ByVal BITMAP As LongPtr, ByVal Pitch As Long, ByVal BitsPerPixel As Long, Optional ByVal RedMask As Long, Optional ByVal GreenMask As Long, Optional ByVal BlueMask As Long, Optional ByVal TopDown As Boolean)
    If (BITMAP = 0) Then Exit Sub
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to convert a 'header-only' bitmap.")
    If (Pitch > 0) Then
Dim lHeight As Long: lHeight = FreeImage_GetHeight(BITMAP)
         ReDim Bits((Pitch * lHeight) - 1)
         Call FreeImage_ConvertToRawBits(VarPtr(Bits(0)), BITMAP, Pitch, BitsPerPixel, RedMask, GreenMask, BlueMask, TopDown)
    End If
End Sub
Public Function FreeImage_GetHistogramEx(ByVal BITMAP As LongPtr, Optional ByVal Channel As FREE_IMAGE_COLOR_CHANNEL = FICC_BLACK, Optional ByRef Success As Boolean) As Long()
' returns a DIB's histogram data as VB style array of
' type Long. Since histogram data is never modified directly, it seems
' enough to return a clone of the data and no read/write accessible
' array wrapped around the actual pointer.

    If (BITMAP = 0) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to get histogram of a 'header-only' bitmap.")
Dim alResult() As LongPtr: ReDim alResult(255)
    Success = (p_FreeImage_GetHistogram(BITMAP, alResult(0), Channel) = 1)
    If (Success) Then Call p_Swap(VarPtrArray(FreeImage_GetHistogramEx), VarPtrArray(alResult))
End Function
Public Function FreeImage_AdjustCurveEx(ByVal BITMAP As LongPtr, ByRef LookupTable As Variant, Optional ByVal Channel As FREE_IMAGE_COLOR_CHANNEL = FICC_BLACK) As Boolean
' extends the FreeImage function 'FreeImage_AdjustCurve'
' to a more VB suitable function. The parameter 'LookupTable' may
' either be an array of type Byte or may contain the pointer to a memory
' block, what in VB is always the address of the memory block, since VB
' actually doesn's support native pointers.
' In case of providing the memory block as an array, make sure, that the
' array contains exactly 256 items. In case of providing an address of a
' memory block, the size of the memory block is assumed to be 256 bytes
' and it is up to the caller to ensure that it is large enough.
    If (BITMAP = 0) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to adjust a 'header-only' bitmap.")
Dim lpData As LongPtr, lSizeInBytes As Long
    If (IsArray(LookupTable)) Then
       lpData = p_GetMemoryBlockPtrFromVariant(LookupTable, lSizeInBytes)
    ElseIf (IsNumeric(LookupTable)) Then
       lSizeInBytes = 256
       lpData = CLng(LookupTable)
    End If
    If ((lpData <> 0) And (lSizeInBytes = 256)) Then FreeImage_AdjustCurveEx = (p_FreeImage_AdjustCurve(BITMAP, lpData, Channel) = 1)
End Function
Public Function FreeImage_GetLockedPageNumbersEx(ByVal BITMAP As LongPtr, Optional ByRef Count As Long) As Long()
' extends the FreeImage function FreeImage_GetLockedPageNumbers()
' to a more VB suitable function. The original FreeImage parameter 'pages', which
' is a pointer to an array of Long, containing all locked page numbers, was turned
' into a return value, which is a real VB-style array of type Long. The original
' Boolean return value, indicating if there are any locked pages, was dropped from
' this function. The caller has to check the 'Count' parameter, which works according
' to the FreeImage API documentation.
' returns an array of Longs, dimensioned from 0 to (Count - 1), that
' contains the page numbers of all currently locked pages of 'BITMAP', if 'Count' is
' greater than 0 after the function returns. If 'Count' is 0, there are no pages
' locked and the function returns an uninitialized array.
Dim lpPages As LongPtr
Dim lRet As Long: lRet = p_FreeImage_GetLockedPageNumbers(BITMAP, lpPages, Count)
    If (lRet <> 1) Or (Count = 0) Then Exit Function
Dim alResult() As Long: ReDim alResult(0 To Count - 1)
    Call CopyMemory(alResult(0), ByVal lpPages, Count * 4)
End Function
' Memory and Stream functions
Public Function FreeImage_GetFileTypeFromMemoryEx(ByRef Data As Variant, Optional ByRef SizeInBytes As Long) As FREE_IMAGE_FORMAT
' extends the FreeImage function FreeImage_GetFileTypeFromMemory()
' to a more VB suitable function. The parameter data of type Variant my
' me either an array of type Byte, Integer or Long or may contain the pointer
' to a memory block, what in VB is always the address of the memory block,
' since VB actually doesn's support native pointers.
' In case of providing the memory block as an array, the SizeInBytes may
' be omitted, zero or less than zero. Then, the size of the memory block
' is calculated correctly. When SizeInBytes is given, it is up to the caller
' to ensure, it is correct.
' In case of providing an address of a memory block, SizeInBytes must not
' be omitted.
' get both pointer and size in bytes of the memory block provided
' through the Variant parameter 'data'.
Dim lDataPtr As LongPtr: lDataPtr = p_GetMemoryBlockPtrFromVariant(Data, SizeInBytes)
Dim hStream As LongPtr:  hStream = FreeImage_OpenMemoryByPtr(lDataPtr, SizeInBytes)
   If (hStream) Then
      ' on success, detect image type
      FreeImage_GetFileTypeFromMemoryEx = FreeImage_GetFileTypeFromMemory(hStream)
      Call FreeImage_CloseMemory(hStream)
   Else
      FreeImage_GetFileTypeFromMemoryEx = FIF_UNKNOWN
   End If
End Function
Public Function FreeImage_LoadFromMemoryEx(ByRef Data As Variant, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS, Optional ByRef SizeInBytes As Long, Optional ByRef Format As FREE_IMAGE_FORMAT) As LongPtr
' loads a FreeImage bitmap from memory that has been passed
' through parameter 'Data'. This parameter is of type Variant and may actually
' be an array of type Byte, Integer or Long or may contain the address of an
' arbitrary block of memory.
' The parameter 'SizeInBytes' specifies the size of the passed block of memory
' in bytes. It may be omitted, if parameter 'Data' contains an array of type Byte,
' Integer or Long upon entry. In that case, or if 'SizeInBytes' is zero or less
' than zero, the size is determined directly from the array and also passed back
' to the caller through parameter 'SizeInBytes'.
' The parameter 'Format' is an OUT only parameter that contains the image type
' of the loaded image after the function returns.
' The parameter 'Flags' works according to the FreeImage API documentation.
' get both pointer and size in bytes of the memory block provided
' through the Variant parameter 'data'.
Dim lDataPtr As LongPtr: lDataPtr = p_GetMemoryBlockPtrFromVariant(Data, SizeInBytes)
Dim hStream As LongPtr:  hStream = FreeImage_OpenMemoryByPtr(lDataPtr, SizeInBytes)
   If (hStream) = 0 Then Exit Function
    ' on success, detect image type
    Format = FreeImage_GetFileTypeFromMemory(hStream)
    If (Format <> FIF_UNKNOWN) Then FreeImage_LoadFromMemoryEx = FreeImage_LoadFromMemory(Format, hStream, Flags)
    ' close the memory stream
    Call FreeImage_CloseMemory(hStream)
End Function
Public Function FreeImage_SaveToMemoryEx(ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As LongPtr, ByRef Data() As Byte, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS, Optional ByVal UnloadSource As Boolean) As Boolean
' saves a FreeImage bitmap into memory and returns it through the byte array passed in parameter 'Data()'.
' It makes a deep copy of the memory stream's byte buffer, into which the image has been saved.
' The memory stream is closed properly before the function returns.
' The provided byte array 'Data()' must not be a fixed sized array.
' It will be dimensioned to the size required to hold all the memory stream's data.
' The parameters 'Format', 'Bitmap' and 'Flags' work according to the FreeImage API documentation.
' The optional 'UnloadSource' parameter is for unloading the original image after it has been saved into memory.
' There is no need to clean up the DIB at the caller's site.
' The function returns True on success and False otherwise.
    If (BITMAP = 0) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to save a 'header-only' bitmap.")
Dim hStream As LongPtr: hStream = FreeImage_OpenMemory()
    If (hStream) Then
       FreeImage_SaveToMemoryEx = FreeImage_SaveToMemory(Format, BITMAP, hStream, Flags)
       If (FreeImage_SaveToMemoryEx) Then
Dim lpData As LongPtr, lSizeInBytes As Long
          If (p_FreeImage_AcquireMemory(hStream, lpData, lSizeInBytes)) Then
             On Error Resume Next
             ReDim Data(lSizeInBytes - 1)
             If (Err.Number = NOERROR) Then
                On Error GoTo 0
                Call CopyMemory(Data(0), ByVal lpData, lSizeInBytes)
             Else
                On Error GoTo 0
                FreeImage_SaveToMemoryEx = False
             End If
          Else
             FreeImage_SaveToMemoryEx = False
          End If
       End If
       Call FreeImage_CloseMemory(hStream)
    Else
       FreeImage_SaveToMemoryEx = False
    End If
    If (UnloadSource) Then Call FreeImage_Unload(BITMAP)
End Function
Public Function FreeImage_SaveToMemoryEx2(ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As LongPtr, ByRef Data() As Byte, ByRef Stream As LongPtr, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS, Optional ByVal UnloadSource As Boolean) As Boolean
' saves a FreeImage bitmap into memory and returns it through
' the byte array passed in parameter 'Data()'. In contrast to function
' 'FreeImage_SaveToMemoryEx', it does not make a deep copy of the memory
' stream's byte buffer, but directly wraps the array 'Data()' around the stream's
' byte buffer by calling function 'FreeImage_AcquireMemoryEx'.
' As a result, the memory stream must remain valid while the array 'Data()' is in use.
' In other words, the stream must be maintained by the caller of this function.
' The provided byte array 'Data()' must not be a fixed sized array.
' It will be dimensioned to the size required to hold all the memory stream's data.
' To reuse the caller's array variable that was passed through parameter 'Data()'
' before it goes out of the caller's scope, it must first be destroyed by passing
' it to the 'FreeImage_DestroyLockedArray' function.
' The parameter 'Stream' is an IN/OUT parameter, that keeps track of the memory
' stream, the VB array 'Data()' is based on. This parameter may contain an
' already opened FreeImage memory stream upon entry and will contain a valid
' memory stream when the function returns. It is left up to the caller to close
' this memory stream correctly.
' The array 'Data()' will no longer be valid and accessible after the stream has
' been closed, so the stream should only be closed after the passed byte array
' variable goes out of the caller's scope or is reused.
' The parameters 'Format', 'Bitmap' and 'Flags' work according to the FreeImage API documentation.
' The optional 'UnloadSource' parameter is for unloading the original image after it has been saved to memory.
' There is no need to clean up the DIB at the caller's site.
' The function returns True on success and False otherwise.
    If (BITMAP = 0) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to save a 'header-only' bitmap.")
    If (Stream = 0) Then Stream = FreeImage_OpenMemory()
    If (Stream) Then
       FreeImage_SaveToMemoryEx2 = FreeImage_SaveToMemory(Format, BITMAP, Stream, Flags)
       If (FreeImage_SaveToMemoryEx2) Then FreeImage_SaveToMemoryEx2 = FreeImage_AcquireMemoryEx(Stream, Data)
       ' Do not close the memory stream, since the returned array Data()
       ' directly points to the stream's data. The stream handle is passed back
       ' to the caller through parameter 'Stream'. The caller must close
       ' this stream after being done with the array.
    Else
       FreeImage_SaveToMemoryEx2 = False
    End If
    If (UnloadSource) Then Call FreeImage_Unload(BITMAP)
End Function
Public Function FreeImage_AcquireMemoryEx(ByVal Stream As LongPtr, ByRef Data() As Byte, Optional ByRef SizeInBytes As Long) As Boolean
' wraps the byte array passed through parameter 'Data()' around the
' memory acquired from the specified memory stream. After the function returns,
' the array passed in 'Data()' points directly to the stream's data pointer and so,
' provides full read and write access to the streams byte buffer.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArray' function.
    If (Stream = 0) Then Exit Function
Dim lpData As LongPtr: If (p_FreeImage_AcquireMemory(Stream, lpData, SizeInBytes)) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
       .cbElements = 1                           ' one element is one byte
       .cDims = 1                                ' the array has only 1 dimension
       .cElements = SizeInBytes                  ' the number of elements in the array is the size in bytes of the memory block
       .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
       .pvData = lpData                          ' let the array point to the memory block received by FreeImage_AcquireMemory
    End With
Dim lpSA As LongPtr: lpSA = p_DeRefPtr(VarPtrArray(Data))
    If (lpSA = 0) Then
       ' allocate memory for an array descriptor
       Call SafeArrayAllocDescriptor(1, lpSA)
       Call CopyMemory(ByVal VarPtrArray(Data), lpSA, PTR_LENGTH)
    Else
       Call SafeArrayDestroyData(lpSA)
    End If
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
End Function
Public Function FreeImage_ReadMemoryEx(ByRef Buffer As Variant, ByVal Stream As LongPtr, Optional ByRef Count As Long, Optional ByRef Size As Long) As Long
' is a wrapper for 'FreeImage_ReadMemory()' using VB style
' arrays instead of a void pointer.
' The variant parameter 'Buffer' may be a Byte, Integer or Long array or
' may contain a pointer to a memory block (the memory block's address).
' In the latter case, this function behaves exactly like
' function 'FreeImage_ReadMemory()'. Then, 'Count' and 'Size' must be valid
' upon entry.
' If 'Buffer' is an initialized (dimensioned) array, 'Count' and 'Size' may
' be omitted. Then, the array's layout is used to determine 'Count'
' and 'Size'. In that case, any provided value in 'Count' and 'Size' upon
' entry will override these calculated values as long as they are not
' exceeding the size of the array in 'Buffer'.
' If 'Buffer' is an uninitialized (not yet dimensioned) array of any valid
' type (Byte, Integer or Long) and, at least 'Count' is specified, the
' array in 'Buffer' is redimensioned by this function. If 'Buffer' is a
' fixed-size or otherwise locked array, a runtime error (10) occurs.
' If 'Size' is omitted, the array's element size is assumed to be the
' desired value.
' As FreeImage's function 'FreeImage_ReadMemory()', this function returns
' the number of items actually read.
' Example: (very freaky...)
'
' Dim alLongBuffer() As Long
' Dim lRet as Long
'
'    ' now reading 303 integers (2 byte) into an array of Longs
'    lRet = FreeImage_ReadMemoryEx(alLongBuffer, lMyStream, 303, 2)
'
'    ' now, lRet contains 303 and UBound(alLongBuffer) is 151 since
'    ' we need at least 152 Longs (0..151) to store (303 * 2) = 606 bytes
'    ' so, the higest two bytes of alLongBuffer(151) contain only unset
'    ' bits. Got it?
' Remark: This function's parameter order differs from FreeImage's
'         original funtion 'FreeImage_ReadMemory()'!
Dim lBufferPtr As LongPtr
Dim lSizeInBytes As Long
Dim lSize As Long
Dim lCount As Long
   If (VarType(Buffer) And vbArray) Then
      ' get both pointer and size in bytes of the memory block provided
      ' through the Variant parameter 'Buffer'.
      lBufferPtr = p_GetMemoryBlockPtrFromVariant(Buffer, lSizeInBytes, lSize)
      If (lBufferPtr = 0) Then
         ' array is not initialized
         If (Count > 0) Then
            ' only if we have a 'Count' value, redim the array
            If (Size <= 0) Then
               ' if 'Size' is omitted, use array's element size
               Size = lSize
            End If
            Select Case lSize
            Case 2
               ' Remark: -Int(-a) == ceil(a); a > 0
               ReDim Buffer(-Int(-Count * Size / 2) - 1) As Integer
            Case 4
               ' Remark: -Int(-a) == ceil(a); a > 0
               ReDim Buffer(-Int(-Count * Size / 4) - 1) As Long
            Case Else
               ReDim Buffer((Count * Size) - 1) As Byte
            End Select
            lBufferPtr = p_GetMemoryBlockPtrFromVariant(Buffer, lSizeInBytes, lSize)
         End If
      End If
      If (lBufferPtr) Then
         lCount = lSizeInBytes / lSize
         If (Size <= 0) Then
            ' use array's natural value for 'Size' when
            ' omitted
            Size = lSize
         End If
         If (Count <= 0) Then
            ' use array's natural value for 'Count' when
            ' omitted
            Count = lCount
         End If
         If ((Size * Count) > (lSize * lCount)) Then
            If (Size = lSize) Then
               Count = lCount
            Else
               ' Remark: -Fix(-a) == floor(a); a > 0
               Count = -Fix(-lSizeInBytes / Size)
               If (Count = 0) Then
                  Size = lSize
                  Count = lCount
               End If
            End If
         End If
         FreeImage_ReadMemoryEx = FreeImage_ReadMemory(lBufferPtr, Size, Count, Stream)
      End If
   ElseIf (VarType(Buffer) = vbLong) Then
      ' if Buffer is a Long, it specifies the address of a memory block
      ' then, we do not know anything about its size, so assume that 'Size'
      ' and 'Count' are correct and forward these directly to the FreeImage
      ' call.
      FreeImage_ReadMemoryEx = FreeImage_ReadMemory(CLng(Buffer), Size, Count, Stream)
   End If
End Function
Public Function FreeImage_WriteMemoryEx(ByRef Buffer As Variant, ByVal Stream As LongPtr, Optional ByRef Count As Long, Optional ByRef Size As Long) As Long
' is a wrapper for 'FreeImage_WriteMemory()' using VB style
' arrays instead of a void pointer.
' The variant parameter 'Buffer' may be a Byte, Integer or Long array or
' may contain a pointer to a memory block (the memory block's address).
' In the latter case, this function behaves exactly
' like 'FreeImage_WriteMemory()'. Then, 'Count' and 'Size' must be valid
' upon entry.
' If 'Buffer' is an initialized (dimensioned) array, 'Count' and 'Size' may
' be omitted. Then, the array's layout is used to determine 'Count'
' and 'Size'. In that case, any provided value in 'Count' and 'Size' upon
' entry will override these calculated values as long as they are not
' exceeding the size of the array in 'Buffer'.
' If 'Buffer' is an uninitialized (not yet dimensioned) array of any
' type, the function will do nothing an returns 0.
' Remark: This function's parameter order differs from FreeImage's
'         original funtion 'FreeImage_ReadMemory()'!
Dim lBufferPtr As LongPtr
Dim lSizeInBytes As Long
Dim lSize As Long
Dim lCount As Long
   If (VarType(Buffer) And vbArray) Then
      ' get both pointer and size in bytes of the memory block provided
      ' through the Variant parameter 'Buffer'.
      lBufferPtr = p_GetMemoryBlockPtrFromVariant(Buffer, lSizeInBytes, lSize)
      If (lBufferPtr) Then
         lCount = lSizeInBytes / lSize
         If (Size <= 0) Then
            ' use array's natural value for 'Size' when
            ' omitted
            Size = lSize
         End If
         If (Count <= 0) Then
            ' use array's natural value for 'Count' when
            ' omitted
            Count = lCount
         End If
         If ((Size * Count) > (lSize * lCount)) Then
            If (Size = lSize) Then
               Count = lCount
            Else
               ' Remark: -Fix(-a) == floor(a); a > 0
               Count = -Fix(-lSizeInBytes / Size)
               If (Count = 0) Then
                  Size = lSize
                  Count = lCount
               End If
            End If
         End If
         FreeImage_WriteMemoryEx = FreeImage_WriteMemory(lBufferPtr, Size, Count, Stream)
      End If
   ElseIf (VarType(Buffer) = vbLong) Then
      ' if Buffer is a Long, it specifies the address of a memory block
      ' then, we do not know anything about its size, so assume that 'Size'
      ' and 'Count' are correct and forward these directly to the FreeImage
      ' call.
      FreeImage_WriteMemoryEx = FreeImage_WriteMemory(CLng(Buffer), Size, Count, Stream)
   End If
End Function
Public Function FreeImage_LoadBitmapFromMemoryEx(ByRef Data As Variant, Optional Page As Long = -1, Optional ByVal Width As Long, Optional ByVal Height As Long, Optional BitDepth As Long, Optional ByRef Filter As FREE_IMAGE_FILTER, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS, Optional ByRef SizeInBytes As Long, Optional ByRef Format As FREE_IMAGE_FORMAT) As LongPtr
' loads a FreeImage bitmap from memory that has been passed through parameter 'Data'.
' If Bitmap is multipaged then select appropriate page and return it
' If Width/Height is set size picture to satisfy conditions
'-------------------------
' Data - array of Bytes, Integer or Long or address of an arbitrary block of memory containing image
' Page - page num of multipaged image to return. if <0 then  selected by size (for ICO only), on exit return real pagenum
' Width/Height/BitDepths - desired size and bit depths of image
' flags -
' SizeInBytes -
' Format -
'-------------------------
' v.1.0.0       : 02.09.2021 - исходная версия
'-------------------------
' ToDo: hold multipage handle for future work with it (for example for GIF frame switch)
'-------------------------
    On Error GoTo HandleError
' open the memory stream
Dim pdata As LongPtr:   pdata = p_GetMemoryBlockPtrFromVariant(Data, SizeInBytes): If (pdata) = 0 Then Exit Function
Dim hStream As LongPtr: hStream = FreeImage_OpenMemoryByPtr(pdata, SizeInBytes): If (hStream) = 0 Then Exit Function
' detect image type
    Format = FreeImage_GetFileTypeFromMemory(hStream): If (Format = FIF_UNKNOWN) Then Err.Raise 5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "The file specified has an unknown image format."
' check conditions
'    ' if set only one dimension make suppose target region square
'    If ((Width > 0) And (Height > 0)) Then
'    ElseIf (Height > 0) Then Width = Height
'    ElseIf (Width > 0) Then Height = Width
'    Else: Width = 0: Height = 0
'    End If
Dim pMult As LongPtr
Dim pResult As LongPtr
' load the image from memory stream only, if known image type
    Select Case Format
    Case FIF_ICO
    ' open ICO multibitmap
        Flags = FILO_ICO_MAKEALPHA
        pMult = FreeImage_LoadMultiBitmapFromMemory(Format, hStream, Flags)
        If (pMult = 0) Then Err.Raise 5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Can't load multipage bitmap."
        ' select page by size (for icons and if page num not set directly)
        If (Page < 0) Then
            pResult = FreeImage_GetIconBestMatch(pMult, Abs(Width), Abs(Height), BitDepth, Page)
        Else
            pResult = FreeImage_LockPage(pMult, Page)
        End If
        If FreeImage_GetBPP(pResult) = 32 Then FreeImage_PreMultiplyWithAlpha (pResult)
    Case FIF_GIF
    ' open GIF multibitmap
        Flags = FILO_GIF_PLAYBACK
        pMult = FreeImage_LoadMultiBitmapFromMemory(Format, hStream, Flags)
        If (pMult = 0) Then Err.Raise 5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Can't load multipage bitmap."
        If (Page < 0) Then Page = 0 ' если не задано выбираем первую страницу
        pResult = FreeImage_LockPage(pMult, Page)
        If FreeImage_GetBPP(pResult) = 32 Then FreeImage_PreMultiplyWithAlpha (pResult)
    Case FIF_TIFF
    ' open TIFF multibitmap
        pMult = FreeImage_LoadMultiBitmapFromMemory(Format, hStream, Flags)
        If (pMult = 0) Then Err.Raise 5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Can't load multipage bitmap."
        If (Page < 0) Then Page = 0 ' если не задано выбираем первую страницу
        pResult = FreeImage_LockPage(pMult, Page)
        'If FreeImage_GetBPP(pResult) = 32 Then FreeImage_PreMultiplyWithAlpha (pResult)
    Case FIF_PNG
    ' open PNG
        Flags = FILO_PNG_DEFAULT 'FILO_PNG_IGNOREGAMMA
        pResult = FreeImage_LoadFromMemory(Format, hStream, Flags)
        If FreeImage_GetBPP(pResult) = 32 Then FreeImage_PreMultiplyWithAlpha (pResult)
    Case Else
    ' open bitmap
        pResult = FreeImage_LoadFromMemory(Format, hStream, Flags)
        If FreeImage_GetBPP(pResult) = 32 Then FreeImage_PreMultiplyWithAlpha (pResult)
    End Select
    If (pResult = 0) Then Err.Raise 5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Can't load bitmap."
' size picture
    If Width = 0 Then GoTo HandleExit
    'pResult = FreeImage_RescaleByPixel(pResult, Width, Height, True, Filter)
HandleExit:  If (hStream <> 0) Then Call FreeImage_CloseMemory(hStream)
             'If (pMult <> 0) Then Call FreeImage_CloseMultiBitmap(pMult)
             FreeImage_LoadBitmapFromMemoryEx = pResult: Exit Function
HandleError: pResult = False: Err.Clear: Resume HandleExit
End Function
Public Function FreeImage_LoadMultiBitmapFromMemoryEx(ByRef Data As Variant, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS, Optional ByRef SizeInBytes As Long, Optional ByRef Format As FREE_IMAGE_FORMAT) As LongPtr
' loads a FreeImage multipage bitmap from memory that has been
' passed through parameter 'Data'. This parameter is of type Variant and may
' actually be an array of type Byte, Integer or Long or may contain the
' address of an arbitrary block of memory.
' The parameter 'SizeInBytes' specifies the size of the passed block of memory
' in bytes. It may be omitted, if parameter 'Data' contains an array of type Byte,
' Integer or Long upon entry. In that case, or if 'SizeInBytes' is zero or less
' than zero, the size is determined directly from the array and also passed back
' to the caller through parameter 'SizeInBytes'.
' The parameter 'Format' is an OUT only parameter that contains the image type
' of the loaded image after the function returns.
' The parameter 'Flags' works according to the FreeImage API documentation.
' get both pointer and size in bytes of the memory block provided
' through the Variant parameter 'Data'.
Dim lDataPtr As LongPtr:    lDataPtr = p_GetMemoryBlockPtrFromVariant(Data, SizeInBytes)
Dim hStream As LongPtr:     hStream = FreeImage_OpenMemoryByPtr(lDataPtr, SizeInBytes)
    If (hStream = 0) Then Exit Function
    ' on success, detect image type
    Format = FreeImage_GetFileTypeFromMemory(hStream)
    If (Format <> FIF_UNKNOWN) Then
       ' load the image from memory stream only, if known image type
        Select Case Format
        Case FIF_TIFF, FIF_GIF, FIF_ICO
            FreeImage_LoadMultiBitmapFromMemoryEx = FreeImage_LoadMultiBitmapFromMemory(Format, hStream, Flags)
        Case Else
            Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "FreeImage Library plugin '" & FreeImage_GetFormatFromFIF(Format) & "' " & "does not have any support for multi-page bitmaps.")
        End Select
    Else
        Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "The file specified has an unknown image format.")
    End If
'    ' close the memory stream
'    ' ??? if stream is closed then LockPage will crash app ???
'    Call FreeImage_CloseMemory(hStream)
End Function
Public Function FreeImage_SaveMultiBitmapToMemoryEx(ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As LongPtr, ByRef Data() As Byte, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS, Optional ByVal UnloadSource As Boolean) As Boolean
' saves a FreeImage multipage bitmap into memory and returns it
' through the byte array passed in parameter 'Data()'. It makes a deep copy of
' the memory stream's byte buffer, into which the image has been saved. The
' memory stream is closed properly before the function returns.
' The provided byte array 'Data()' must not be a fixed sized array. It will be
' dimensioned to the size required to hold all the memory stream's data.
' The parameters 'Format', 'Bitmap' and 'Flags' work according to the FreeImage
' API documentation.
' The optional 'UnloadSource' parameter is for unloading the original image
' after it has been saved into memory. There is no need to clean up the DIB at the caller's site.
' The function returns True on success and False otherwise.
    If (BITMAP = 0) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to save a 'header-only' bitmap.")
Dim hStream As LongPtr: hStream = FreeImage_OpenMemory()
Dim lpData As LongPtr
Dim lSizeInBytes As Long
    If (hStream) Then
        FreeImage_SaveMultiBitmapToMemoryEx = FreeImage_SaveMultiBitmapToMemory(Format, BITMAP, hStream, Flags)
        If (FreeImage_SaveMultiBitmapToMemoryEx) Then
            If (p_FreeImage_AcquireMemory(hStream, lpData, lSizeInBytes)) Then
                On Error Resume Next
                ReDim Data(lSizeInBytes - 1)
                If (Err.Number = NOERROR) Then
                    On Error GoTo 0
                    Call CopyMemory(Data(0), ByVal lpData, lSizeInBytes)
                Else
                    On Error GoTo 0
                    FreeImage_SaveMultiBitmapToMemoryEx = False
                End If
            Else
                FreeImage_SaveMultiBitmapToMemoryEx = False
            End If
        End If
        Call FreeImage_CloseMemory(hStream)
    Else
        FreeImage_SaveMultiBitmapToMemoryEx = False
    End If
    If (UnloadSource) Then Call p_FreeImage_CloseMultiBitmap(BITMAP)
End Function
Public Function FreeImage_SaveMultiBitmapToMemoryEx2(ByVal Format As FREE_IMAGE_FORMAT, ByVal BITMAP As LongPtr, ByRef Data() As Byte, ByRef Stream As LongPtr, Optional ByVal Flags As FREE_IMAGE_SAVE_OPTIONS, Optional ByVal UnloadSource As Boolean) As Boolean
' saves a FreeImage multipage bitmap into memory and returns it
' through the byte array passed in parameter 'Data()'. In contrast to function
' 'FreeImage_SaveToMemoryEx', it does not make a deep copy of the memory
' stream's byte buffer, but directly wraps the array 'Data()' around the stream's
' byte buffer by calling function 'FreeImage_AcquireMemoryEx'. As a result, the
' memory stream must remain valid while the array 'Data()' is in use. In other
' words, the stream must be maintained by the caller of this function.
' The provided byte array 'Data()' must not be a fixed sized array. It will be
' dimensioned to the size required to hold all the memory stream's data.
' To reuse the caller's array variable that was passed through parameter 'Data()'
' before it goes out of the caller's scope, it must first be destroyed by passing
' it to the 'FreeImage_DestroyLockedArray' function.
' The parameter 'Stream' is an IN/OUT parameter, that keeps track of the memory
' stream, the VB array 'Data()' is based on. This parameter may contain an
' already opened FreeImage memory stream upon entry and will contain a valid
' memory stream when the function returns. It is left up to the caller to close
' this memory stream correctly.
' The array 'Data()' will no longer be valid and accessible after the stream has
' been closed, so the stream should only be closed after the passed byte array
' variable goes out of the caller's scope or is reused.
' The parameters 'Format', 'Bitmap' and 'Flags' work according to the FreeImage
' API documentation.
' The optional 'UnloadSource' parameter is for unloading the original image
' after it has been saved to memory. There is no need to clean up the DIB
' at the caller's site.
' The function returns True on success and False otherwise.
    If Not (BITMAP) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to save a 'header-only' bitmap.")
    If (Stream = 0) Then Stream = FreeImage_OpenMemory()
    If (Stream) Then
       FreeImage_SaveMultiBitmapToMemoryEx2 = FreeImage_SaveMultiBitmapToMemory(Format, BITMAP, Stream, Flags)
       If (FreeImage_SaveMultiBitmapToMemoryEx2) Then FreeImage_SaveMultiBitmapToMemoryEx2 = FreeImage_AcquireMemoryEx(Stream, Data)
       ' Do not close the memory stream, since the returned array Data()
       ' directly points to the stream's data. The stream handle is passed back
       ' to the caller through parameter 'Stream'. The caller must close
       ' this stream after being done with the array.
    Else
       FreeImage_SaveMultiBitmapToMemoryEx2 = False
    End If
    If (UnloadSource) Then Call p_FreeImage_CloseMultiBitmap(BITMAP)
End Function
'----------------------
' Derived and hopefully useful functions
'----------------------
' Plugin and filename functions
Public Function FreeImage_IsExtensionValidForFIF(ByVal Format As FREE_IMAGE_FORMAT, ByVal Extension As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Boolean
' tests, whether a given filename extension is valid for a certain image format (fif).
   FreeImage_IsExtensionValidForFIF = (InStr(1, FreeImage_GetFIFExtensionList(Format) & ",", Extension & ",", Compare) > 0)
End Function
Public Function FreeImage_IsFilenameValidForFIF(ByVal Format As FREE_IMAGE_FORMAT, ByVal FileName As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Boolean
' tests, whether a given complete filename is valid for a certain image format (fif).
Dim i As Long:   i = InStrRev(FileName, ".")
Dim strExtension As String: If (i > 0) Then strExtension = Mid$(FileName, i + 1): FreeImage_IsFilenameValidForFIF = (InStr(1, FreeImage_GetFIFExtensionList(Format) & ",", strExtension & ",", Compare) > 0)
End Function
Public Function FreeImage_GetPrimaryExtensionFromFIF(ByVal Format As FREE_IMAGE_FORMAT) As String
' returns the primary (main or most commonly used?) extension
' of a certain image format (fif). This is done by returning the first of
' all possible extensions returned by FreeImage_GetFIFExtensionList(). That
' assumes, that the plugin returns the extensions in ordered form. If not,
' in most cases it is even enough, to receive any extension.
' is primarily used by the function 'SavePictureEx'.
Dim strExtensionList As String: strExtensionList = FreeImage_GetFIFExtensionList(Format)
Dim i As Long: i = InStr(strExtensionList, ",")
   If (i > 0) Then
      FreeImage_GetPrimaryExtensionFromFIF = Left$(strExtensionList, i - 1)
   Else
      FreeImage_GetPrimaryExtensionFromFIF = strExtensionList
   End If
End Function
Public Function FreeImage_IsGreyscaleImage(ByVal BITMAP As LongPtr) As Boolean
' returns a boolean value that is true, if the DIB is actually
' a greyscale image. Here, the only test condition is, that each palette
' entry must be a grey value, what means that each color component has the
' same value (red = green = blue).
' The FreeImage libraray doesn't offer a function to determine if a DIB is
' greyscale. The only thing you can do is to use the 'FreeImage_GetColorType'
' function, that returns either FIC_MINISWHITE or FIC_MINISBLACK for
' greyscale images. However, a DIB needs to have a ordered greyscale palette
' (linear ramp or inverse linear ramp) to be judged as FIC_MINISWHITE or
' FIC_MINISBLACK. DIB's with an unordered palette that are actually (visually)
' greyscale, are said to be (color-)palletized. That's also true for any 4 bpp
' image, since it will never have a palette that satifies the tests done
' in the 'FreeImage_GetColorType' function.
' So, there is a chance to omit some color depth conversions, when displaying
' an image in greyscale fashion. Maybe the problem will be solved in the
' FreeImage library one day.
Dim atRGB() As RGBQUAD
Dim i As Long
   Select Case FreeImage_GetBPP(BITMAP)
   Case 1, 4, 8
      atRGB = FreeImage_GetPaletteEx(BITMAP)
      FreeImage_IsGreyscaleImage = True
      For i = 0 To UBound(atRGB)
         With atRGB(i)
            If ((.rgbRed <> .rgbGreen) Or (.rgbRed <> .rgbBlue)) Then
               FreeImage_IsGreyscaleImage = False
               Exit For
            End If
         End With
      Next i
   End Select
End Function
' Bitmap resolution functions
Public Function FreeImage_GetResolutionX(ByVal BITMAP As LongPtr) As Long: FreeImage_GetResolutionX = Int(0.5 + 0.0254 * FreeImage_GetDotsPerMeterX(BITMAP)): End Function
Public Sub FreeImage_SetResolutionX(ByVal BITMAP As LongPtr, ByVal resolution As Long): Call FreeImage_SetDotsPerMeterX(BITMAP, Int(resolution / 0.0254 + 0.5)): End Sub
Public Function FreeImage_GetResolutionY(ByVal BITMAP As LongPtr) As Long: FreeImage_GetResolutionY = Int(0.5 + 0.0254 * FreeImage_GetDotsPerMeterY(BITMAP)): End Function
Public Sub FreeImage_SetResolutionY(ByVal BITMAP As LongPtr, ByVal resolution As Long): Call FreeImage_SetDotsPerMeterY(BITMAP, Int(resolution / 0.0254 + 0.5)): End Sub
' Bitmap Info functions
Public Function FreeImage_GetInfoHeaderEx(ByVal BITMAP As LongPtr) As BITMAPINFOHEADER
' is a wrapper around FreeImage_GetInfoHeader() and returns a fully
' populated BITMAPINFOHEADER structure for a given bitmap.
Dim lpInfoHeader As LongPtr: lpInfoHeader = FreeImage_GetInfoHeader(BITMAP)
   If (lpInfoHeader) Then Call CopyMemory(FreeImage_GetInfoHeaderEx, ByVal lpInfoHeader, LenB(FreeImage_GetInfoHeaderEx))
End Function
' Image color depth conversion wrapper
Public Function FreeImage_ConvertColorDepth(ByVal BITMAP As LongPtr, ByVal Conversion As FREE_IMAGE_CONVERSION_FLAGS, Optional ByVal UnloadSource As Boolean, Optional ByVal threshold As Byte = 128, Optional ByVal DitherMethod As FREE_IMAGE_DITHER = FID_FS, Optional ByVal QuantizeMethod As FREE_IMAGE_QUANTIZE = FIQ_WUQUANT) As LongPtr
' is an easy-to-use wrapper for color depth conversion, intended
' to work around some tweaks in the FreeImage library.
' The parameters 'Threshold' and 'eDitherMode' control how thresholding or
' dithering are performed. The 'QuantizeMethod' parameter determines, what
' quantization algorithm will be used when converting to 8 bit color images.
' The 'Conversion' parameter, which can contain a single value or an OR'ed
' combination of some of the FREE_IMAGE_CONVERSION_FLAGS enumeration values,
' determines the desired output image format.
' The optional 'UnloadSource' parameter is for unloading the original image, so
' you can "change" an image with this function rather than getting a new DIB
' pointer. There is no more need for a second DIB variable at the caller's site.
    If (BITMAP = 0) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to convert a 'header-only' bitmap.")
Dim lBPP As Long: lBPP = FreeImage_GetBPP(BITMAP)
Dim bForceLinearRamp As Boolean: bForceLinearRamp = ((Conversion And FICF_REORDER_GREYSCALE_PALETTE) = 0)
Dim hDIBsrc As LongPtr, hDIBdst As LongPtr
Dim lpReservePalette As LongPtr
Dim bAdjustReservePaletteSize As Boolean
    Select Case (Conversion And (Not FICF_REORDER_GREYSCALE_PALETTE))
    Case FICF_MONOCHROME_THRESHOLD:   If (lBPP > 1) Then hDIBdst = FreeImage_Threshold(BITMAP, threshold)
    Case FICF_MONOCHROME_DITHER:      If (lBPP > 1) Then hDIBdst = FreeImage_Dither(BITMAP, DitherMethod)
    Case FICF_GREYSCALE_4BPP
       If (lBPP <> 4) Then
          ' If the color depth is 1 bpp and the we don't have a linear ramp palette
          ' the bitmap is first converted to an 8 bpp greyscale bitmap with a linear
          ' ramp palette and then to 4 bpp.
          If ((lBPP = 1) And (FreeImage_GetColorType(BITMAP) = FIC_PALETTE)) Then
             hDIBsrc = BITMAP
             BITMAP = FreeImage_ConvertToGreyscale(BITMAP)
             Call FreeImage_Unload(hDIBsrc)
          End If
          hDIBdst = FreeImage_ConvertTo4Bits(BITMAP)
       Else
          ' The bitmap is already 4 bpp but may not have a linear ramp.
          ' If we force a linear ramp the bitmap is converted to 8 bpp with a linear ramp
          ' and then back to 4 bpp.
          If (((Not bForceLinearRamp) And (Not FreeImage_IsGreyscaleImage(BITMAP))) Or (bForceLinearRamp And (FreeImage_GetColorType(BITMAP) = FIC_PALETTE))) Then
             hDIBsrc = FreeImage_ConvertToGreyscale(BITMAP)
             hDIBdst = FreeImage_ConvertTo4Bits(hDIBsrc)
             Call FreeImage_Unload(hDIBsrc)
          End If
       End If
    Case FICF_GREYSCALE_8BPP
       ' Convert, if the bitmap is not at 8 bpp or does not have a linear ramp palette.
       If ((lBPP <> 8) Or (((Not bForceLinearRamp) And (Not FreeImage_IsGreyscaleImage(BITMAP))) Or (bForceLinearRamp And (FreeImage_GetColorType(BITMAP) = FIC_PALETTE)))) Then
          hDIBdst = FreeImage_ConvertToGreyscale(BITMAP)
       End If
    Case FICF_PALLETISED_8BPP
       ' note, that the FreeImage library only quantizes 24 or 32 bit images (expect FIQ_NNQUANT)
       ' do not convert any 8 bit images
       If (lBPP <> 8) Then
          ' images with a color depth of 24 bits can directly be
          ' converted with the FreeImage_ColorQuantize function;
          ' other images may need to be converted to 24 bits first
          If (lBPP = 24 Or (lBPP = 32 And QuantizeMethod <> FIQ_NNQUANT)) Then
             hDIBdst = FreeImage_ColorQuantize(BITMAP, QuantizeMethod)
          Else
             hDIBsrc = FreeImage_ConvertTo24Bits(BITMAP)
             hDIBdst = FreeImage_ColorQuantize(hDIBsrc, QuantizeMethod)
             Call FreeImage_Unload(hDIBsrc)
          End If
       End If
    Case FICF_RGB_15BPP: If (lBPP <> 15) Then hDIBdst = FreeImage_ConvertTo16Bits555(BITMAP)
    Case FICF_RGB_16BPP: If (lBPP <> 16) Then hDIBdst = FreeImage_ConvertTo16Bits565(BITMAP)
    Case FICF_RGB_24BPP: If (lBPP <> 24) Then hDIBdst = FreeImage_ConvertTo24Bits(BITMAP)
    Case FICF_RGB_32BPP: If (lBPP <> 32) Then hDIBdst = FreeImage_ConvertTo32Bits(BITMAP)
    End Select
    If (hDIBdst) Then
       FreeImage_ConvertColorDepth = hDIBdst
       If (UnloadSource) Then Call FreeImage_Unload(BITMAP)
    Else
       FreeImage_ConvertColorDepth = BITMAP
    End If
End Function
Public Function FreeImage_ColorQuantizeEx(ByVal BITMAP As LongPtr, Optional ByVal QuantizeMethod As FREE_IMAGE_QUANTIZE = FIQ_WUQUANT, Optional ByVal UnloadSource As Boolean, Optional ByVal PaletteSize As Long = 256, Optional ByVal ReserveSize As Long, Optional ByRef ReservePalette As Variant = Null) As LongPtr
' is a more VB-friendly wrapper around FreeImage_ColorQuantizeEx,
' which lets you specify the ReservePalette to be used not only as a pointer, but
' also as a real VB-style array of type Long, where each Long item takes a color
' in ARGB format (&HAARRGGBB). The native FreeImage function FreeImage_ColorQuantizeEx
' is declared private and named FreeImage_ColorQuantizeExInt and so hidden from the
' world outside the wrapper.
' In contrast to the FreeImage API documentation, ReservePalette is of type Variant
' and may either be a pointer to palette data (pointer to an array of type RGBQUAD
' == VarPtr(atMyPalette(0)) in VB) or an array of type Long, which then must contain
' the palette data in ARGB format. You can receive palette data as an array Longs
' from function FreeImage_GetPaletteExLong.
' Although ReservePalette is of type Variant, arrays of type RGBQUAD can not be
' passed, as long as RGBQUAD is not declared as a public type in a public object
' module. So, when dealing with RGBQUAD arrays, you are stuck on VarPtr or may
' use function FreeImage_GetPalettePtr, which is a more meaningfully named
' convenience wrapper around VarPtr.
' The optional 'UnloadSource' parameter is for unloading the original image, so
' you can "change" an image with this function rather than getting a new DIB
' pointer. There is no more need for a second DIB variable at the caller's site.
' All other parameters work according to the FreeImage API documentation.
' Note: Currently, any provided ReservePalette is only used, if quantize is
'       FIQ_NNQUANT. This seems to be either a bug or an undocumented
'       limitation of the FreeImage library (up to version 3.11.0).
    If BITMAP = 0 Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to quantize a 'header-only' bitmap.")
Dim hTmp As LongPtr, lpPalette As LongPtr
Dim lBlockSize As Long
Dim lElementSize As Long
    If (FreeImage_GetBPP(BITMAP) <> 24) Then
       hTmp = BITMAP
       BITMAP = FreeImage_ConvertTo24Bits(BITMAP)
       If (UnloadSource) Then Call FreeImage_Unload(hTmp)
       UnloadSource = True
    End If
    ' adjust PaletteSize
    If (PaletteSize < 2) Then
       PaletteSize = 2
    ElseIf (PaletteSize > 256) Then
       PaletteSize = 256
    End If
    lpPalette = p_GetMemoryBlockPtrFromVariant(ReservePalette, lBlockSize, lElementSize)
    FreeImage_ColorQuantizeEx = p_FreeImage_ColorQuantizeEx(BITMAP, QuantizeMethod, PaletteSize, ReserveSize, lpPalette)
    If (UnloadSource) Then Call FreeImage_Unload(BITMAP)
End Function
Public Function FreeImage_GetPalettePtr(ByRef palette() As RGBQUAD) As LongPtr: FreeImage_GetPalettePtr = VarPtr(palette(0)): End Function
' Image Rescale wrapper functions
Public Function FreeImage_RescaleEx(ByVal BITMAP As LongPtr, Optional ByVal Width As Variant, Optional ByVal Height As Variant, Optional ByVal IsPercentValue As Boolean, Optional ByVal UnloadSource As Boolean, Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC, Optional ByVal ForceCloneCreation As Boolean) As LongPtr
' is a easy-to-use wrapper for rescaling an image with the
' FreeImage library. It returns a pointer to a new rescaled DIB provided by FreeImage.
' The parameters 'Width', 'Height' and 'IsPercentValue' control
' the size of the new image. Here, the function tries to fake something like
' overloading known from Java. It depends on the parameter's data type passed
' through the Variant, how the provided values for width and height are
' actually interpreted. The following rules apply:
' In general, non integer values are either interpreted as percent values or
' factors, the original image size will be multiplied with. The 'IsPercentValue'
' parameter controls whether the values are percent values or factors. Integer
' values are always considered to be the direct new image size, not depending on
' the original image size. In that case, the 'IsPercentValue' parameter has no
' effect. If one of the parameters is omitted, the image will not be resized in
' that direction (either in width or height) and keeps it's original size. It is
' possible to omit both, but that makes actually no sense.
' The following table shows some of possible data type and value combinations
' that might by used with that function: (assume an original image sized 100x100 px)
' Parameter         |  Values |  Values |  Values |  Values |     Values |
' ----------------------------------------------------------------------
' Width             |    75.0 |    0.85 |     200 |     120 |      400.0 |
' Height            |   120.0 |     1.3 |     230 |       - |      400.0 |
' IsPercentValue    |    True |   False |    d.c. |    d.c. |      False | <- wrong option?
' ----------------------------------------------------------------------
' Result Size       |  75x120 |  85x130 | 200x230 | 120x100 |40000x40000 |
' Remarks           | percent |  factor |  direct |         |maybe not   |
'                                                           |what you    |
'                                                           |wanted,     |
'                                                           |right?      |
' The optional 'UnloadSource' parameter is for unloading the original image, so
' you can "change" an image with this function rather than getting a new DIB
' pointer. There is no more need for a second DIB variable at the caller's site.
' As of version 2.0 of the FreeImage VB wrapper, this function and all it's derived
' functions like FreeImage_RescaleByPixel() or FreeImage_RescaleByPercent(), do NOT
' return a clone of the image, if the new size desired is the same as the source
' image's size. That behaviour can be forced by setting the new parameter
' 'ForceCloneCreation' to True. Then, an image is also rescaled (and so
' effectively cloned), if the new width and height is exactly the same as the source
' image's width and height.
' Since this diversity may be confusing to VB developers, this function is also
' callable through three different functions called 'FreeImage_RescaleByPixel',
' 'FreeImage_RescaleByPercent' and 'FreeImage_RescaleByFactor'.
    If BITMAP = 0 Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to rescale a 'header-only' bitmap.")
Dim hDIBNew As LongPtr
Dim lNewWidth As Long, lNewHeight As Long
    If (Not IsMissing(Width)) Then
       Select Case VarType(Width)
       Case vbDouble, vbSingle, vbDecimal, vbCurrency
          lNewWidth = FreeImage_GetWidth(BITMAP) * Width
          If (IsPercentValue) Then lNewWidth = lNewWidth / 100
       Case Else
          lNewWidth = Width
       End Select
    End If
    If (Not IsMissing(Height)) Then
       Select Case VarType(Height)
       Case vbDouble, vbSingle, vbDecimal
          lNewHeight = FreeImage_GetHeight(BITMAP) * Height
          If (IsPercentValue) Then lNewHeight = lNewHeight / 100
       Case Else
          lNewHeight = Height
       End Select
    End If
    If ((lNewWidth > 0) And (lNewHeight > 0)) Then
       If (ForceCloneCreation) Then
          hDIBNew = FreeImage_Rescale(BITMAP, lNewWidth, lNewHeight, Filter)
       ElseIf ((lNewWidth <> FreeImage_GetWidth(BITMAP)) Or (lNewHeight <> FreeImage_GetHeight(BITMAP))) Then
          hDIBNew = FreeImage_Rescale(BITMAP, lNewWidth, lNewHeight, Filter)
       End If
    ElseIf (lNewWidth > 0) Then
       If ((lNewWidth <> FreeImage_GetWidth(BITMAP)) Or (ForceCloneCreation)) Then
          lNewHeight = lNewWidth / (FreeImage_GetWidth(BITMAP) / FreeImage_GetHeight(BITMAP))
          hDIBNew = FreeImage_Rescale(BITMAP, lNewWidth, lNewHeight, Filter)
       End If
    ElseIf (lNewHeight > 0) Then
       If ((lNewHeight <> FreeImage_GetHeight(BITMAP)) Or (ForceCloneCreation)) Then
          lNewWidth = lNewHeight * (FreeImage_GetWidth(BITMAP) / FreeImage_GetHeight(BITMAP))
          hDIBNew = FreeImage_Rescale(BITMAP, lNewWidth, lNewHeight, Filter)
       End If
    End If
    If (hDIBNew) Then
       FreeImage_RescaleEx = hDIBNew
       If (UnloadSource) Then Call FreeImage_Unload(BITMAP)
    Else
       FreeImage_RescaleEx = BITMAP
    End If
End Function
Public Function FreeImage_RescaleByPixel(ByVal BITMAP As LongPtr, Optional ByVal WidthInPixels As Long, Optional ByVal HeightInPixels As Long, Optional ByVal UnloadSource As Boolean, Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC, Optional ByVal ForceCloneCreation As Boolean) As LongPtr: FreeImage_RescaleByPixel = FreeImage_RescaleEx(BITMAP, WidthInPixels, HeightInPixels, False, UnloadSource, Filter, ForceCloneCreation): End Function
Public Function FreeImage_RescaleByPercent(ByVal BITMAP As LongPtr, Optional ByVal WidthPercentage As Double, Optional ByVal HeightPercentage As Double, Optional ByVal UnloadSource As Boolean, Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC, Optional ByVal ForceCloneCreation As Boolean) As LongPtr: FreeImage_RescaleByPercent = FreeImage_RescaleEx(BITMAP, WidthPercentage, HeightPercentage, True, UnloadSource, Filter, ForceCloneCreation): End Function
Public Function FreeImage_RescaleByFactor(ByVal BITMAP As LongPtr, Optional ByVal WidthFactor As Double, Optional ByVal HeightFactor As Double, Optional ByVal UnloadSource As Boolean, Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC, Optional ByVal ForceCloneCreation As Boolean) As LongPtr: FreeImage_RescaleByFactor = FreeImage_RescaleEx(BITMAP, WidthFactor, HeightFactor, False, UnloadSource, Filter, ForceCloneCreation): End Function
' Painting functions
Public Function FreeImage_PaintDC(ByVal hdc As LongPtr, ByVal BITMAP As LongPtr, Optional ByVal XDst As Long, Optional ByVal YDst As Long, Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, Optional ByVal Width As Long, Optional ByVal Height As Long) As Long
' draws a FreeImage DIB directly onto a device context (DC). There
' are many (selfexplaining?) parameters that control the visual result.
' Parameters 'XDst' and 'YDst' specify the point where the output should
' be painted and 'XSrc', 'YSrc', 'Width' and 'Height' span a rectangle
' in the source image 'Bitmap' that defines the area to be painted.
' If any of parameters 'Width' and 'Height' is zero, it is transparently substituted
' by the width or height of teh bitmap to be drawn, resprectively.
    If ((hdc = 0) Or (BITMAP = 0)) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to paint a 'header-only' bitmap.")
    If (Width = 0) Then Width = FreeImage_GetWidth(BITMAP)
    If (Height = 0) Then Height = FreeImage_GetHeight(BITMAP)
    FreeImage_PaintDC = SetDIBitsToDevice(hdc, XDst, YDst - ySrc, Width, Height, xSrc, ySrc, 0, Height, FreeImage_GetBits(BITMAP), FreeImage_GetInfo(BITMAP), DIB_RGB_COLORS)
End Function
Public Function FreeImage_PaintDCEx(ByVal hdc As LongPtr, ByVal BITMAP As LongPtr, Optional ByVal XDst As Long, Optional ByVal YDst As Long, Optional ByVal WidthDst As Long, Optional ByVal HeightDst As Long, Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, Optional ByVal WidthSrc As Long, Optional ByVal HeightSrc As Long, Optional ByVal DrawMode As DRAW_MODE = DM_DRAW_DEFAULT, Optional ByVal RasterOperator As RASTER_OPERATOR = ROP_SRCCOPY, Optional ByVal StretchMode As STRETCH_MODE = SM_COLORONCOLOR) As Long
' draws a FreeImage DIB directly onto a device context (DC).
' There are many (selfexplaining?) parameters that control the visual result.
' The main difference of this function compared to the 'FreeImage_PaintDC' is,
' that this function supports both mirroring and stretching of the image to be
' painted and so, is somewhat slower than 'FreeImage_PaintDC'.
    If ((hdc = 0) Or (BITMAP = 0)) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to paint a 'header-only' bitmap.")
Dim eLastStretchMode As STRETCH_MODE: eLastStretchMode = GetStretchBltMode(hdc)
    Call SetStretchBltMode(hdc, StretchMode)
    If (WidthSrc = 0) Then WidthSrc = FreeImage_GetWidth(BITMAP)
    If (WidthDst = 0) Then WidthDst = WidthSrc
    If (HeightSrc = 0) Then HeightSrc = FreeImage_GetHeight(BITMAP)
    If (HeightDst = 0) Then HeightDst = HeightSrc
    If (DrawMode And DM_MIRROR_VERTICAL) Then YDst = YDst + HeightDst: HeightDst = -HeightDst
    If (DrawMode And DM_MIRROR_HORIZONTAL) Then XDst = XDst + WidthDst:   WidthDst = -WidthDst
    FreeImage_PaintDCEx = StretchDIBits(hdc, XDst, YDst, WidthDst, HeightDst, xSrc, ySrc, WidthSrc, HeightSrc, FreeImage_GetBits(BITMAP), FreeImage_GetInfo(BITMAP), DIB_RGB_COLORS, RasterOperator)
    ' restore last mode
    Call SetStretchBltMode(hdc, eLastStretchMode)
End Function
Public Function FreeImage_PaintTransparent(ByVal hdc As LongPtr, ByVal BITMAP As LongPtr, Optional ByVal XDst As Long = 0, Optional ByVal YDst As Long = 0, Optional ByVal WidthDst As Long, Optional ByVal HeightDst As Long, Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, Optional ByVal WidthSrc As Long, Optional ByVal HeightSrc As Long, Optional ByVal Alpha As Byte = 255) As Long
' paints a device independent bitmap to any device context and
' thereby honors any transparency information associated with the bitmap.
' Furthermore, through the 'Alpha' parameter, an overall transparency level may be specified.
' For palletised images, any color set to be transparent in the transparency
' table, will be transparent. For high color images, only 32-bit images may
' have any transparency information associated in their alpha channel. Only
' these may be painted with transparency by this function.
' Since this is a wrapper for the Windows GDI function AlphaBlend(), 31-bit
' images, containing alpha (or per-pixel) transparency, must be 'premultiplied'
' for alpha transparent regions to actually show transparent.
' See MSDN help on the AlphaBlend() function.
' FreeImage also offers a function to premultiply 32-bit bitmaps with their alpha
' channel, according to the needs of AlphaBlend(). Have a look at function
' FreeImage_PreMultiplyWithAlpha().
' Overall transparency level may be specified for all bitmaps in all color
' depths supported by FreeImage. If needed, bitmaps are transparently converted
' to 32-bit and unloaded after the paint operation. This is also true for palletised bitmaps.
' Parameters 'hDC' and 'Bitmap' seem to be very self-explanatory. All other parameters
' are optional. The group of '*Dest*' parameters span a rectangle on the destination
' device context, used as drawing area for the bitmap. If these are omitted, the
' bitmap will be drawn starting at position 0,0 in the bitmap's actual size.
' The group of '*Src*' parameters span a rectangle on the source bitmap, used as
' cropping area for the paint operation. If both rectangles differ in size in any
' direction, the part of the image actually painted is stretched for to fit into
' the drawing area. If any of the parameters '*Width' or '*Height' are omitted,
' the bitmap's actual size (width or height) will be used.
' The 'Alpha' parameter specifies the overall transparency. It takes values in the
' range from 0 to 255. Using 0 will paint the bitmap fully transparent, 255 will
' paint the image fully opaque. The 'Alpha' value controls, how the non per-pixel
' portions of the image will be drawn.
    If ((hdc = 0) Or (BITMAP = 0)) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to paint a 'header-only' bitmap.")
Dim bIsTransparent As Boolean
    ' get image width if not specified
    If (WidthSrc = 0) Then WidthSrc = FreeImage_GetWidth(BITMAP)
    If (WidthDst = 0) Then WidthDst = WidthSrc
    ' get image height if not specified
    If (HeightSrc = 0) Then HeightSrc = FreeImage_GetHeight(BITMAP)
    If (HeightDst = 0) Then HeightDst = HeightSrc
Dim lpPalette As LongPtr: lpPalette = FreeImage_GetPalette(BITMAP)
    If (lpPalette) Then
' for palletized images
Dim lPaletteSize As Long: lPaletteSize = FreeImage_GetColorsUsed(BITMAP) * 4
Dim alPalOrg(255) As Long, alPalMod(255) As Long, alPalMask(255) As Long
       Call CopyMemory(alPalOrg(0), ByVal lpPalette, lPaletteSize)
       Call CopyMemory(alPalMod(0), ByVal lpPalette, lPaletteSize)
Dim abTT() As Byte: abTT = FreeImage_GetTransparencyTableEx(BITMAP)
Dim i As Long
       If ((Alpha = 255) And (HeightDst >= HeightSrc) And (WidthDst >= WidthSrc)) Then
    ' create a mask palette and a modified version of the original palette
          For i = 0 To UBound(abTT)
             If (abTT(i) = 0) Then
                alPalMask(i) = &HFFFFFFFF   ' white
                alPalMod(i) = &H0           ' black
                bIsTransparent = True
             End If
          Next i
          If (Not bIsTransparent) Then
        ' if there is no transparency in the image, paint it with a single SRCCOPY
             Call StretchDIBits(hdc, XDst, YDst, WidthDst, HeightDst, xSrc, ySrc, WidthSrc, HeightSrc, FreeImage_GetBits(BITMAP), FreeImage_GetInfo(BITMAP), DIB_RGB_COLORS, SRCCOPY)
          Else
        ' set mask palette and paint with SRCAND
             Call CopyMemory(ByVal lpPalette, alPalMask(0), lPaletteSize)
             Call StretchDIBits(hdc, XDst, YDst, WidthDst, HeightDst, xSrc, ySrc, WidthSrc, HeightSrc, FreeImage_GetBits(BITMAP), FreeImage_GetInfo(BITMAP), DIB_RGB_COLORS, SRCAND)
        ' set mask modified and paint with SRCPAINT
             Call CopyMemory(ByVal lpPalette, alPalMod(0), lPaletteSize)
             Call StretchDIBits(hdc, XDst, YDst, WidthDst, HeightDst, xSrc, ySrc, WidthSrc, HeightSrc, FreeImage_GetBits(BITMAP), FreeImage_GetInfo(BITMAP), DIB_RGB_COLORS, SRCPAINT)
        ' restore original palette
             Call CopyMemory(ByVal lpPalette, alPalOrg(0), lPaletteSize)
          End If
          ' we are done, do not paint with AlphaBlend() any more
          BITMAP = 0
       Else
    ' create a premultiplied palette
        ' since we have no real per pixel transparency in a palletized
        ' image, we only need to set all transparent colors to zero.
          For i = 0 To UBound(abTT)
             If (abTT(i) = 0) Then alPalMod(i) = 0
          Next i
          ' set premultiplied palette and convert to 32 bits
          Call CopyMemory(ByVal lpPalette, alPalMod(0), lPaletteSize)
          BITMAP = FreeImage_ConvertTo32Bits(BITMAP)
          ' restore original palette
          Call CopyMemory(ByVal lpPalette, alPalOrg(0), lPaletteSize)
       End If
    End If
    
    If (BITMAP = 0) Then Exit Function
Dim hMemDC As LongPtr
Dim hBitmap As LongPtr, hBitmapOld As LongPtr
Dim lBF As Long, tBF As BLENDFUNCTION
    hMemDC = CreateCompatibleDC(0): If (hMemDC = 0) Then Exit Function
    hBitmap = FreeImage_GetBitmap(BITMAP, hMemDC)
    hBitmapOld = SelectObject(hMemDC, hBitmap)

'If IsDebug Then
''Stop
'Dim fiTemp As LongPtr:
'fiTemp = FreeImage_CreateFromDC(hMemDC): FreeImage_Save FIF_BMP, fiTemp, CurrentProject.path & "\fiPictFromDC.bmp"
'fiTemp = FreeImage_CreateFromDC(hDC): FreeImage_Save FIF_BMP, fiTemp, CurrentProject.path & "\fiBackFromDC.bmp"
'End If
'Const SetHalfTone = True ' Stretch_Halftone not compatible with win9x
'Dim lStretchMode As Long: If SetHalfTone Then lStretchMode = SetStretchBltMode(hDC, STRETCH_HALFTONE)
    With tBF
       .BlendOp = AC_SRC_OVER
       .SourceConstantAlpha = Alpha
       If (FreeImage_GetBPP(BITMAP) = 32) Then .AlphaFormat = AC_SRC_ALPHA
    End With
    Call CopyMemory(lBF, tBF, PTR_LENGTH)  '4) ' convert typed BF to long value BF (not a ptr to type)
    Call AlphaBlend(hdc, XDst, YDst, WidthDst, HeightDst, hMemDC, xSrc, ySrc, WidthSrc, HeightSrc, lBF)
'If SetHalfTone Then SetStretchBltMode hDC, lStretchMode ' Stretch_Halftone not compatible with win9x
    Call SelectObject(hMemDC, hBitmapOld)
    Call DeleteObject(hBitmap)
    Call DeleteDC(hMemDC)
    If (lpPalette) Then Call FreeImage_Unload(BITMAP)
End Function
'----------------------
' Pixel access functions
'----------------------
Public Function FreeImage_GetBitsEx(ByVal BITMAP As LongPtr) As Byte()
' returns a two dimensional Byte array containing a DIB's
' data-bits. This is done by wrapping a true VB array around the memory
' block the returned pointer of FreeImage_GetBits() is pointing to. So, the
' array returned provides full read and write acces to the image's data.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the FreeImage_DestroyLockedArray() function.
    If (BITMAP = 0) Then Exit Function
Dim tSA As SAFEARRAY2D
    With tSA
       .cbElements = 1                           ' size in bytes per array element
       .cDims = 2                                ' the array has 2 dimensions
       .cElements1 = FreeImage_GetHeight(BITMAP) ' the number of elements in y direction (height of Bitmap)
       .cElements2 = FreeImage_GetPitch(BITMAP)  ' the number of elements in x direction (byte width of Bitmap)
       .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
       .pvData = FreeImage_GetBits(BITMAP)       ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(2, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetBitsEx), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetBitsExRGBTRIPLE(ByVal BITMAP As LongPtr) As RGBTRIPLE()
' returns a two dimensional RGBTRIPLE array containing a DIB's
' data-bits. This is done by wrapping a true VB array around the memory
' block the returned pointer of 'FreeImage_GetBits' is pointing to. So, the
' array returned provides full read and write acces to the image's data.
' only works with 24 bpp images and, since each FreeImage scanline
' is aligned to a 32-bit boundary, only if the image's width in pixels multiplied
' by three modulo four is zero. That means, that the image layout in memory must
' "naturally" be aligned to a 32-bit boundary, since arrays do not support padding.
' So, the function only returns an initialized array, if this equotion is true:
' (((ImageWidthPixels * 3) Mod 4) = 0)
' In other words, this is true for all images with no padding.
' For instance, only images with these widths will be suitable for this function:
' 100, 104, 108, 112, 116, 120, 124, ...
' Have a look at the wrapper function 'FreeImage_GetScanlinesRGBTRIPLE()' to have
' a way to work around that limitation.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArray' function.
    If (BITMAP = 0) Then Exit Function
    'If (FreeImage_GetImageType(BITMAP) <> FIT_BITMAP) Then Exit Function
    If (FreeImage_GetBPP(BITMAP) <> 24) Then Exit Function
    If (((FreeImage_GetWidth(BITMAP) * 3) Mod 4) <> 0) Then Exit Function
Dim tSA As SAFEARRAY2D
    With tSA
        .cbElements = 3                           ' size in bytes per array element
        .cDims = 2                                ' the array has 2 dimensions
        .cElements1 = FreeImage_GetHeight(BITMAP) ' the number of elements in y direction (height of Bitmap)
        .cElements2 = FreeImage_GetWidth(BITMAP)  ' the number of elements in x direction (byte width of Bitmap)
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in sizer erased; according to Matthew Curland never use FIXEDSIZE alone
        .pvData = FreeImage_GetBits(BITMAP)       ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(2, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetBitsExRGBTRIPLE), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetBitsExRGBQUAD(ByVal BITMAP As LongPtr) As RGBQUAD()
' returns a two dimensional RGBQUAD array containing a DIB's
' data-bits. This is done by wrapping a true VB array around the memory
' block the returned pointer of 'FreeImage_GetBits' is pointing to. So, the
' array returned provides full read and write acces to the image's data.
' only works with 32 bpp images. Since each scanline must
' "naturally" start at a 32-bit boundary if each pixel uses 32 bits, there
' are no padding problems like these known with 'FreeImage_GetBitsExRGBTRIPLE',
' so, this function is suitable for all 32 bpp images of any size.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArray' function.
    If (BITMAP = 0) Then Exit Function
    'If (FreeImage_GetImageType(BITMAP) <> FIT_BITMAP) Then Exit Function
    If (FreeImage_GetBPP(BITMAP) <> 32) Then Exit Function
Dim tSA As SAFEARRAY2D
    With tSA
        .cbElements = 4                           ' size in bytes per array element
        .cDims = 2                                ' the array has 2 dimensions
        .cElements1 = FreeImage_GetHeight(BITMAP) ' the number of elements in y direction (height of Bitmap)
        .cElements2 = FreeImage_GetWidth(BITMAP)  ' the number of elements in x direction (byte width of Bitmap)
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in sizer erased; according to Matthew Curland never use FIXEDSIZE alone
        .pvData = FreeImage_GetBits(BITMAP)       ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(2, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetBitsExRGBQUAD), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetBitsExLong(ByVal BITMAP As LongPtr) As Long()
' returns a two dimensional long array containing a DIB's
' data-bits. This is done by wrapping a true VB array around the memory
' block the returned pointer of 'FreeImage_GetBits' is pointing to. So, the
' array returned provides full read and write acces to the image's data.
' only works with 32 bpp images. Since each scanline must
' "naturally" start at a 32-bit boundary if each pixel uses 32 bits, there
' are no padding problems like these known with 'FreeImage_GetBitsExRGBTRIPLE',
' so, this function is suitable for all 32 bpp images of any size.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArray' function.
    If (BITMAP = 0) Then Exit Function
    'If (FreeImage_GetImageType(BITMAP) <> FIT_BITMAP) Then Exit Function
    If (FreeImage_GetBPP(BITMAP) <> 32) Then Exit Function
Dim tSA As SAFEARRAY2D
    With tSA
        .cbElements = 4                           ' size in bytes per array element
        .cDims = 2                                ' the array has 2 dimensions
        .cElements1 = FreeImage_GetHeight(BITMAP) ' the number of elements in y direction (height of Bitmap)
        .cElements2 = FreeImage_GetWidth(BITMAP)  ' the number of elements in x direction (byte width of Bitmap)
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
        .pvData = FreeImage_GetBits(BITMAP)       ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(2, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetBitsExLong), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetScanLinesRGBTRIPLE(ByVal BITMAP As LongPtr, ByRef Scanlines As ScanLinesRGBTRIBLE, Optional ByVal Reverse As Boolean) As Long
    If (BITMAP = 0) Then Exit Function
    If (FreeImage_GetImageType(BITMAP) <> FIT_BITMAP) Then Exit Function
    If (FreeImage_GetBPP(BITMAP) <> 24) Then Exit Function
Dim lHeight As Long:  lHeight = FreeImage_GetHeight(BITMAP)
    ReDim Scanlines.Scanline(lHeight - 1)
Dim i As Long
    For i = 0 To lHeight - 1
        If (Not Reverse) Then
            Scanlines.Scanline(i).Data = FreeImage_GetScanLineBITMAP24(BITMAP, i)
        Else
            Scanlines.Scanline(i).Data = FreeImage_GetScanLineBITMAP24(BITMAP, lHeight - i - 1)
        End If
    Next i
    FreeImage_GetScanLinesRGBTRIPLE = lHeight
End Function
Public Function FreeImage_GetScanLineEx(ByVal BITMAP As LongPtr, ByVal Scanline As Long) As Byte()
' returns a one dimensional Byte array containing a whole
' scanline's data-bits. This is done by wrapping a true VB array around
' the memory block the returned pointer of 'FreeImage_GetScanline' is
' pointing to. So, the array returned provides full read and write acces
' to the image's data.
' This is the most generic function of a complete function set dealing with
' scanline data, since this function returns an array of type Byte. It is
' up to the caller of the function to interpret these bytes correctly,
' according to the results of FreeImage_GetBPP and FreeImage_GetImageType.
' You may consider using any of the non-generic functions named
' 'FreeImage_GetScanLineXXX', that return an array of proper type, according
' to the images bit depth and type.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArray' function.
    If (BITMAP = 0) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
        .cbElements = 1                           ' size in bytes per array element
        .cDims = 1                                ' the array has only 1 dimension
        .cElements = FreeImage_GetLine(BITMAP)    ' the number of elements in the array
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
        .pvData = FreeImage_GetScanline(BITMAP, Scanline) ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetScanLineEx), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetScanLineBITMAP8(ByVal BITMAP As LongPtr, ByVal Scanline As Long) As Byte()
' returns a one dimensional Byte array containing a whole
' scanline's data-bits of a 8 bit bitmap image. This is done by wrapping
' a true VB array around the memory block the returned pointer of
' 'FreeImage_GetScanline' is pointing to. So, the array returned provides
' full read and write acces to the image's data.
' is just a thin wrapper for 'FreeImage_GetScanLineEx' but
' includes checking of the image's bit depth and type, as all of the
' non-generic scanline functions do.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArray' function.
    If (FreeImage_GetImageType(BITMAP) <> FIT_BITMAP) Then Exit Function
    Select Case FreeImage_GetBPP(BITMAP)
    Case 1, 4, 8: FreeImage_GetScanLineBITMAP8 = FreeImage_GetScanLineEx(BITMAP, Scanline)
    End Select
End Function
Public Function FreeImage_GetScanLineBITMAP16(ByVal BITMAP As LongPtr, ByVal Scanline As Long) As Integer()
' returns a one dimensional Integer array containing a whole
' scanline's data-bits of a 16 bit bitmap image. This is done by wrapping
' a true VB array around the memory block the returned pointer of
' 'FreeImage_GetScanline' is pointing to. So, the array returned provides
' full read and write acces to the image's data.
' The function includes checking of the image's bit depth and type and
' returns a non-initialized array if 'Bitmap' is an image of improper type.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArray' function.
    If (FreeImage_GetImageType(BITMAP) <> FIT_BITMAP) Then Exit Function
    If (FreeImage_GetBPP(BITMAP) <> 16) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
        .cbElements = 2                           ' size in bytes per array element
        .cDims = 1                                ' the array has only 1 dimension
        .cElements = FreeImage_GetWidth(BITMAP)   ' the number of elements in the array
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
        .pvData = FreeImage_GetScanline(BITMAP, Scanline) ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetScanLineBITMAP16), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetScanLineBITMAP24(ByVal BITMAP As LongPtr, ByVal Scanline As Long) As RGBTRIPLE()
' returns a one dimensional RGBTRIPLE array containing a whole
' scanline's data-bits of a 24 bit bitmap image. This is done by wrapping
' a true VB array around the memory block the returned pointer of
' 'FreeImage_GetScanline' is pointing to. So, the array returned provides
' full read and write acces to the image's data.
' The function includes checking of the image's bit depth and type and
' returns a non-initialized array if 'Bitmap' is an image of improper type.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArrayRGBTRIPLE' function.
    If (FreeImage_GetImageType(BITMAP) <> FIT_BITMAP) Then Exit Function
    If (FreeImage_GetBPP(BITMAP) <> 24) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
        .cbElements = 3                           ' size in bytes per array element
        .cDims = 1                                ' the array has only 1 dimension
        .cElements = FreeImage_GetWidth(BITMAP)   ' the number of elements in the array
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
        .pvData = FreeImage_GetScanline(BITMAP, Scanline) ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetScanLineBITMAP24), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetScanLineBITMAP32(ByVal BITMAP As LongPtr, ByVal Scanline As Long) As RGBQUAD()
' returns a one dimensional RGBQUAD array containing a whole
' scanline's data-bits of a 32 bit bitmap image. This is done by wrapping
' a true VB array around the memory block the returned pointer of
' 'FreeImage_GetScanline' is pointing to. So, the array returned provides
' full read and write acces to the image's data.
' The function includes checking of the image's bit depth and type and
' returns a non-initialized array if 'Bitmap' is an image of improper type.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArrayRGBQUAD' function.
    If (FreeImage_GetImageType(BITMAP) <> FIT_BITMAP) Then Exit Function
    If (FreeImage_GetBPP(BITMAP) <> 32) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
        .cbElements = 4                           ' size in bytes per array element
        .cDims = 1                                ' the array has only 1 dimension
        .cElements = FreeImage_GetWidth(BITMAP)   ' the number of elements in the array
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
        .pvData = FreeImage_GetScanline(BITMAP, Scanline) ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetScanLineBITMAP32), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetScanLineINT16(ByVal BITMAP As LongPtr, ByVal Scanline As Long) As Integer()
' returns a one dimensional Integer array containing a whole
' scanline's data-bits of a FIT_INT16 or FIT_UINT16 image. This is done
' by wrapping a true VB array around the memory block the returned pointer
' of 'FreeImage_GetScanline' is pointing to. So, the array returned
' provides full read and write acces to the image's data.
' The function includes checking of the image's bit depth and type and
' returns a non-initialized array if 'Bitmap' is an image of improper type.
' Since VB does not distinguish between signed and unsigned data types, both
' image types FIT_INT16 and FIT_UINT16 are handled with this function. If 'Bitmap'
' specifies an image of type FIT_UINT16, it is up to the caller to treat the
' array's Integers as unsigned, although VB knows signed Integers only.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArray' function.
Dim eImageType As FREE_IMAGE_TYPE: eImageType = FreeImage_GetImageType(BITMAP): If (eImageType <> FIT_UINT16) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
       .cbElements = 2                           ' size in bytes per array element
       .cDims = 1                                ' the array has only 1 dimension
       .cElements = FreeImage_GetWidth(BITMAP)   ' the number of elements in the array
       .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
       .pvData = FreeImage_GetScanline(BITMAP, Scanline) ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetScanLineINT16), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetScanLineINT32(ByVal BITMAP As LongPtr, ByVal Scanline As Long) As Long()
' returns a one dimensional Long array containing a whole
' scanline's data-bits of a FIT_INT32 or FIT_UINT32 image. This is done
' by wrapping a true VB array around the memory block the returned pointer
' of 'FreeImage_GetScanline' is pointing to. So, the array returned
' provides full read and write acces to the image's data.
' The function includes checking of the image's bit depth and type and
' returns a non-initialized array if 'Bitmap' is an image of improper type.
' Since VB does not distinguish between signed and unsigned data types, both
' image types FIT_INT32 and FIT_UINT32 are handled with this function. If 'Bitmap'
' specifies an image of type FIT_UINT32, it is up to the caller to treat the
' array's Longs as unsigned, although VB knows signed Longs only.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArray' function.
Dim eImageType As FREE_IMAGE_TYPE: eImageType = FreeImage_GetImageType(BITMAP): If (eImageType <> FIT_UINT32) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
        .cbElements = 4                           ' size in bytes per array element
        .cDims = 1                                ' the array has only 1 dimension
        .cElements = FreeImage_GetWidth(BITMAP)   ' the number of elements in the array
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
        .pvData = FreeImage_GetScanline(BITMAP, Scanline) ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetScanLineINT32), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetScanLineFLOAT(ByVal BITMAP As LongPtr, ByVal Scanline As Long) As Single()
' returns a one dimensional Single array containing a whole
' scanline's data-bits of a FIT_FLOAT image. This is done by wrapping
' a true VB array around the memory block the returned pointer of
' 'FreeImage_GetScanline' is pointing to. So, the array returned  provides
' full read and write acces to the image's data.
' The function includes checking of the image's bit depth and type and
' returns a non-initialized array if 'Bitmap' is an image of improper type.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArray' function.
Dim eImageType As FREE_IMAGE_TYPE: eImageType = FreeImage_GetImageType(BITMAP): If (eImageType <> FIT_FLOAT) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
        .cbElements = 4                           ' size in bytes per array element
        .cDims = 1                                ' the array has only 1 dimension
        .cElements = FreeImage_GetWidth(BITMAP)   ' the number of elements in the array
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues,
                                                  ' so the array can not be modified in size
                                                  ' or erased; according to Matthew Curland never
                                                  ' use FIXEDSIZE alone
        .pvData = FreeImage_GetScanline(BITMAP, Scanline) ' let the array point to the memory block, the
                                                  ' FreeImage scanline data pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetScanLineFLOAT), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetScanLineDOUBLE(ByVal BITMAP As LongPtr, ByVal Scanline As Long) As Double()
' returns a one dimensional Double array containing a whole
' scanline's data-bits of a FIT_DOUBLE image. This is done by wrapping
' a true VB array around the memory block the returned pointer of
' 'FreeImage_GetScanline' is pointing to. So, the array returned  provides
' full read and write acces to the image's data.
' The function includes checking of the image's bit depth and type and
' returns a non-initialized array if 'Bitmap' is an image of improper type.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArray' function.
    Dim eImageType As FREE_IMAGE_TYPE: eImageType = FreeImage_GetImageType(BITMAP): If (eImageType <> FIT_DOUBLE) Then Exit Function
    Dim tSA As SAFEARRAY1D
    With tSA
        .cbElements = 8                           ' size in bytes per array element
        .cDims = 1                                ' the array has only 1 dimension
        .cElements = FreeImage_GetWidth(BITMAP)   ' the number of elements in the array
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland neverse FIXEDSIZE alone
        .pvData = FreeImage_GetScanline(BITMAP, Scanline) ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
    Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetScanLineDOUBLE), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetScanLineCOMPLEX(ByVal BITMAP As LongPtr, ByVal Scanline As Long) As FICOMPLEX()
' returns a one dimensional FICOMPLEX array containing a whole
' scanline's data-bits of a FIT_COMPLEX image. This is done by wrapping
' a true VB array around the memory block the returned pointer of
' 'FreeImage_GetScanline' is pointing to. So, the array returned  provides
' full read and write acces to the image's data.
' The function includes checking of the image's bit depth and type and
' returns a non-initialized array if 'Bitmap' is an image of improper type.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArrayFICOMPLEX' function.
Dim eImageType As FREE_IMAGE_TYPE: eImageType = FreeImage_GetImageType(BITMAP): If (eImageType <> FIT_COMPLEX) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
        .cbElements = 16                          ' size in bytes per array element
        .cDims = 1                                ' the array has only 1 dimension
        .cElements = FreeImage_GetWidth(BITMAP)   ' the number of elements in the array
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
        .pvData = FreeImage_GetScanline(BITMAP, Scanline) ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
    ' For a complete source code documentation have a look at the function 'FreeImage_GetScanLineEx'
    Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetScanLineCOMPLEX), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetScanLineRGB16(ByVal BITMAP As LongPtr, ByVal Scanline As Long) As FIRGB16()
' returns a one dimensional FIRGB16 array containing a whole
' scanline's data-bits of a FIT_RGB16 image. This is done by wrapping
' a true VB array around the memory block the returned pointer of
' 'FreeImage_GetScanline' is pointing to. So, the array returned  provides
' full read and write acces to the image's data.
' The function includes checking of the image's bit depth and type and
' returns a non-initialized array if 'Bitmap' is an image of improper type.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArrayFIRGB16' function.
Dim eImageType As FREE_IMAGE_TYPE: eImageType = FreeImage_GetImageType(BITMAP): If (eImageType <> FIT_RGB16) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
        .cbElements = 6                           ' size in bytes per array element
        .cDims = 1                                ' the array has only 1 dimension
        .cElements = FreeImage_GetWidth(BITMAP)   ' the number of elements in the array
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
        .pvData = FreeImage_GetScanline(BITMAP, Scanline) ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetScanLineRGB16), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetScanLineRGBA16(ByVal BITMAP As LongPtr, ByVal Scanline As Long) As FIRGBA16()
' returns a one dimensional FIRGBA16 array containing a whole
' scanline's data-bits of a FIT_RGBA16 image. This is done by wrapping
' a true VB array around the memory block the returned pointer of
' 'FreeImage_GetScanline' is pointing to. So, the array returned  provides
' full read and write acces to the image's data.
' The function includes checking of the image's bit depth and type and
' returns a non-initialized array if 'Bitmap' is an image of improper type.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArrayFIRGBA16' function.
Dim eImageType As FREE_IMAGE_TYPE: eImageType = FreeImage_GetImageType(BITMAP): If (eImageType <> FIT_RGBA16) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
        .cbElements = 8                           ' size in bytes per array element
        .cDims = 1                                ' the array has only 1 dimension
        .cElements = FreeImage_GetWidth(BITMAP)   ' the number of elements in the array
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
        .pvData = FreeImage_GetScanline(BITMAP, Scanline) ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetScanLineRGBA16), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetScanLineRGBF(ByVal BITMAP As LongPtr, ByVal Scanline As Long) As FIRGBF()
' returns a one dimensional FIRGBF array containing a whole
' scanline's data-bits of a FIT_RGBF image. This is done by wrapping
' a true VB array around the memory block the returned pointer of
' 'FreeImage_GetScanline' is pointing to. So, the array returned  provides
' full read and write acces to the image's data.
' The function includes checking of the image's bit depth and type and
' returns a non-initialized array if 'Bitmap' is an image of improper type.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArrayFIRGBF' function.
Dim eImageType As FREE_IMAGE_TYPE: eImageType = FreeImage_GetImageType(BITMAP): If (eImageType <> FIT_RGBF) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
        .cbElements = 12                          ' size in bytes per array element
        .cDims = 1                                ' the array has only 1 dimension
        .cElements = FreeImage_GetWidth(BITMAP)   ' the number of elements in the array
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
        .pvData = FreeImage_GetScanline(BITMAP, Scanline) ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetScanLineRGBF), lpSA, PTR_LENGTH)
End Function
Public Function FreeImage_GetScanLineRGBAF(ByVal BITMAP As LongPtr, ByVal Scanline As Long) As FIRGBAF()
' returns a one dimensional FIRGBAF array containing a whole
' scanline's data-bits of a FIT_RGBAF image. This is done by wrapping
' a true VB array around the memory block the returned pointer of
' 'FreeImage_GetScanline' is pointing to. So, the array returned  provides
' full read and write acces to the image's data.
' The function includes checking of the image's bit depth and type and
' returns a non-initialized array if 'Bitmap' is an image of improper type.
' To reuse the caller's array variable, this function's result was assigned to,
' before it goes out of scope, the caller's array variable must be destroyed with
' the 'FreeImage_DestroyLockedArrayFIRGBAF' function.
Dim eImageType As FREE_IMAGE_TYPE: eImageType = FreeImage_GetImageType(BITMAP): If (eImageType <> FIT_RGBAF) Then Exit Function
Dim tSA As SAFEARRAY1D
    With tSA
       .cDims = 1                                ' the array has only 1 dimension
       .cbElements = 12                          ' size in bytes per array element
       .cElements = FreeImage_GetWidth(BITMAP)   ' the number of elements in the array
       .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE  ' need AUTO and FIXEDSIZE for safety issues, so the array can not be modified in size or erased; according to Matthew Curland never use FIXEDSIZE alone
       .pvData = FreeImage_GetScanline(BITMAP, Scanline) ' let the array point to the memory block, the FreeImage scanline data pointer points to
    End With
Dim lpSA As LongPtr
    Call SafeArrayAllocDescriptor(1, lpSA)
    Call CopyMemory(ByVal lpSA, tSA, Len(tSA))
    Call CopyMemory(ByVal VarPtrArray(FreeImage_GetScanLineRGBAF), lpSA, PTR_LENGTH)
End Function
'----------------------
' HBITMAP conversion functions
'----------------------
Public Function FreeImage_GetBitmap(ByVal BITMAP As LongPtr, Optional ByVal hdc As LongPtr, Optional ByVal UnloadSource As Boolean) As LongPtr
' returns an HBITMAP created by the CreateDIBSection() function which
' in turn has the same color depth as the original DIB. A reference DC may be provided
' through the 'hDC' parameter. The desktop DC will be used, if no reference DC is specified.
    If (BITMAP = 0) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to create a bitmap from a 'header-only' bitmap.")
Dim bReleaseDC As Boolean: If (hdc = 0) Then hdc = GetDC(0): bReleaseDC = True
    If (hdc = 0) Then Exit Function
Dim ppvBits As LongPtr: FreeImage_GetBitmap = CreateDIBSection(hdc, FreeImage_GetInfo(BITMAP), DIB_RGB_COLORS, ppvBits, 0, 0)
    If ((FreeImage_GetBitmap <> 0) And (ppvBits <> 0)) Then Call CopyMemory(ByVal ppvBits, ByVal FreeImage_GetBits(BITMAP), FreeImage_GetHeight(BITMAP) * FreeImage_GetPitch(BITMAP))
    If (UnloadSource) Then Call FreeImage_Unload(BITMAP)
    If (bReleaseDC) Then Call ReleaseDC(0, hdc)
End Function
Public Function FreeImage_GetBitmapForDevice(ByVal BITMAP As LongPtr, Optional ByVal hdc As LongPtr, Optional ByVal UnloadSource As Boolean) As LongPtr
' returns an HBITMAP created by the CreateDIBitmap() function which
' in turn has always the same color depth as the reference DC, which may be provided
' through the 'hDC' parameter. The desktop DC will be used, if no reference DC is specified.
    If (BITMAP = 0) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to create a bitmap from a 'header-only' bitmap.")
Dim bReleaseDC As Boolean: If (hdc = 0) Then hdc = GetDC(0): bReleaseDC = True
    If (hdc = 0) Then Exit Function
    FreeImage_GetBitmapForDevice = CreateDIBitmap(hdc, FreeImage_GetInfoHeader(BITMAP), CBM_INIT, FreeImage_GetBits(BITMAP), FreeImage_GetInfo(BITMAP), DIB_RGB_COLORS) 'DIB_PAL_COLORS)
If FreeImage_GetBitmapForDevice = 0 Then Stop
    If (UnloadSource) Then Call FreeImage_Unload(BITMAP)
    If (bReleaseDC) Then Call ReleaseDC(0, hdc)
End Function
'----------------------
' OlePicture conversion functions
'----------------------
Public Function FreeImage_GetOlePicture(ByVal BITMAP As LongPtr, Optional ByVal hdc As LongPtr, Optional ByVal UnloadSource As Boolean) As IPictureDisp
' This function creates a VB Picture object (OlePicture) from a FreeImage DIB.
' The original image need not remain valid nor loaded after the VB Picture
' object has been created.

' The optional parameter 'hDC' determines the device context (DC) used for
' transforming the device independent bitmap (DIB) to a device dependent
' bitmap (DDB). This device context's color depth is responsible for this
' transformation. This parameter may be null or omitted. In that case, the
' windows desktop's device context will be used, what will be the desired
' way in almost any cases.

' The optional 'UnloadSource' parameter is for unloading the original image
' after the OlePicture has been created, so you can easily "switch" from a
' FreeImage DIB to a VB Picture object. There is no need to unload the DIB
' at the caller's site if this argument is True.
    If (BITMAP = 0) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to create a bitmap from a 'header-only' bitmap.")
Dim bReleaseDC As Boolean
    If (hdc = 0) Then hdc = GetDC(0): bReleaseDC = True
    If (hdc = 0) Then Exit Function
Dim hBitmap As LongPtr: hBitmap = FreeImage_GetBitmap(BITMAP, hdc, UnloadSource): If (hBitmap = 0) Then Exit Function
'Dim hBitmap As LongPtr: hBitmap = FreeImage_GetBitmapForDevice(BITMAP, hDC, UnloadSource): If (hBitmap = 0) Then Exit Function
Dim tPicDesc As PICTDESC: With tPicDesc: .Type = PICTYPE_BITMAP: .hPic = hBitmap: .Size = LenB(tPicDesc): End With 'vbPicTypeBitmap: .hPal = 0
Dim tGuid As GUID: With tGuid: .Data1 = &H20400: .Data4(0) = &HC0: .Data4(7) = &H46: End With ' IDispatch Interface ID
Dim Ret As Long, cPictureDisp As IPictureDisp: Call OleCreatePictureIndirect(tPicDesc, tGuid, True, cPictureDisp)
    Set FreeImage_GetOlePicture = cPictureDisp
    If (bReleaseDC) Then Call ReleaseDC(0, hdc)
End Function
Public Function FreeImage_GetOlePictureIcon(ByVal BITMAP As LongPtr, Optional ByVal hdc As LongPtr, Optional ByVal UnloadSource As Boolean) As IPictureDisp
' creates a VB Picture object (OlePicture) of type picTypeIcon
' from FreeImage BITMAP handle. The BITMAP handle need not remain valid nor loaded
' after the VB Picture object has been created.
' The optional 'UnloadSource' parameter is for destroying the hIcon image
' after the OlePicture has been created, so you can easiely "switch" from a
' hIcon handle to a VB Picture object. There is no need to unload the hIcon
' at the caller's site if this argument is True.
    If (BITMAP = 0) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to create a bitmap from a 'header-only' bitmap.")
Dim bReleaseDC As Boolean
    If (hdc = 0) Then hdc = GetDC(0): bReleaseDC = True
    If (hdc = 0) Then Exit Function
Dim hIcon As LongPtr: hIcon = FreeImage_GetIcon(BITMAP, hdc:=hdc, UnloadSource:=UnloadSource): If (hIcon = 0) Then Exit Function
Dim tPicDesc As PICTDESC: With tPicDesc: .Type = PICTYPE_ICON: .hPic = hIcon: .Size = LenB(tPicDesc): End With 'vbPicTypeBitmap: .hPal = 0
Dim tGuid As GUID: With tGuid: .Data1 = &H20400: .Data4(0) = &HC0: .Data4(7) = &H46: End With ' IDispatch Interface ID
Dim Ret As Long, cPictureDisp As IPictureDisp: Call OleCreatePictureIndirect(tPicDesc, tGuid, True, cPictureDisp)
    Set FreeImage_GetOlePictureIcon = cPictureDisp
    If (bReleaseDC) Then Call ReleaseDC(0, hdc)
End Function
Public Function FreeImage_GetOlePictureEMF(ByVal BITMAP As LongPtr, Optional ByVal hdc As LongPtr, Optional ByVal UnloadSource As Boolean) As IPicture
' creates a VB Picture object EMF (OlePicture) from a FreeImage DIB.
' The original image need not remain valid nor loaded after the VB Picture object has been created.
    If BITMAP = 0 Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to create a picture from a 'header-only' bitmap.")
'    hBitmap = FreeImage_GetBitmapForDevice(BITMAP, hDC, UnloadSource)
'    If hBitmap = 0 Then Exit Function
    'MakeTransparent (-1&) ' делаем прозрачным
' получаем параметры изображения
Dim hEmf As LongPtr, hIC As LongPtr
Dim rc As RECT
Dim iDPIX As Long, iDPIY As Long
Dim iWEX As Long, iWEY As Long, iVEX As Long, iVEY As Long, iGCD As Long
Dim iBit As Long: iBit = FreeImage_GetBPP(BITMAP)
Dim iWpx As Long: iWpx = FreeImage_GetWidth(BITMAP)
Dim iHpx As Long: iHpx = FreeImage_GetHeight(BITMAP)
    ' EMF sizes must be in Himetric (1/100 mm) HimetricPerInch = 2540
    ' получаем параметры устройства вывода
    hIC = CreateIC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
    iDPIX = GetDeviceCaps(hIC, LOGPIXELSX):      iDPIY = GetDeviceCaps(hIC, LOGPIXELSY)
    ' берем размеры создаваемого изображения - переводим пиксели в Himetric
    rc.Right = Int(iWpx * 2540 / iDPIX + 0.5)
    rc.Bottom = Int(iHpx * 2540 / iDPIY + 0.5)
'Создаём "усовершенствованный" метафайл на котором будем рисовать загруженное изображение
    ' hIC - информационный контекст устройства которое
    iWEX = iWpx * GetDeviceCaps(hIC, HORZSIZE) * iDPIX * 10
    iWEY = iHpx * GetDeviceCaps(hIC, VERTSIZE) * iDPIY * 10
    iVEX = iWpx * GetDeviceCaps(hIC, HORZRES) * 254
    iVEY = iHpx * GetDeviceCaps(hIC, VERTRES) * 254
    iGCD = GCD(GCD(GCD(iWEX, iWEY), iVEX), iVEY) ' определяет НОД
' rc - д.б. в Himetric
    hdc = CreateEnhMetaFile(hIC, vbNullString, rc, vbNullString)
    SetMapMode hdc, MM_ANISOTROPIC
    SetWindowExtEx hdc, iWEX \ iGCD, iWEY \ iGCD, ByVal 0&
    SetViewportExtEx hdc, iVEX \ iGCD, iVEY \ iGCD, ByVal 0&
' рисуем DIB на метафайле
    Call FreeImage_PaintTransparent(hdc, BITMAP)  'Render DestDC:=hDC
    hEmf = CloseEnhMetaFile(hdc): hdc = 0
' function creates a stdPicture object from an image handle (bitmap or icon)
    If hEmf = 0 Then Exit Function
' fill tPictDesc structure with necessary parts for EMF
Dim tPicDesc As PICTDESC: With tPicDesc: .Type = PICTYPE_ENHMETAFILE: .hPic = hEmf: .hPal = 0: .Size = LenB(tPicDesc): End With
'Dim tGuid As GUID: With tGuid: .Data1 = &H7BF80981: .Data2 = &HBF32: .Data3 = &H101A: .Data4(0) = &H8B: .Data4(1) = &HBB: .Data4(3) = &HAA: .Data4(5) = &H30: .Data4(6) = &HC: .Data4(7) = &HAB: End With
Dim tGuid As GUID: With tGuid: .Data1 = &H20400: .Data4(0) = &HC0: .Data4(7) = &H46: End With ' IDispatch Interface ID
Dim cPictureDisp As IPictureDisp: Call OleCreatePictureIndirect(tPicDesc, tGuid, True, cPictureDisp)
    Set FreeImage_GetOlePictureEMF = cPictureDisp ': Exit Function
    If (UnloadSource) Then Call FreeImage_Unload(BITMAP)
End Function
Public Function FreeImage_GetOlePictureThumbnail(ByVal BITMAP As LongPtr, ByVal MaxPixelSize As Long, Optional ByVal hdc As LongPtr, Optional ByVal UnloadSource As Boolean) As IPicture
' is a IOlePicture aware wrapper for FreeImage_MakeThumbnail(). It
' returns a VB Picture object instead of a FreeImage DIB.
' The optional 'UnloadSource' parameter is for unloading the original image
' after the OlePicture has been created, so you can easiely "switch" from a
' FreeImage DIB to a VB Picture object. There is no need to clean up the DIB at the caller's site.
Dim hDIBThumbnail As LongPtr
    If Not (BITMAP) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to create a thumbnail picture from a 'header-only' bitmap.")
    hDIBThumbnail = FreeImage_MakeThumbnail(BITMAP, MaxPixelSize)
    Set FreeImage_GetOlePictureThumbnail = FreeImage_GetOlePicture(hDIBThumbnail, hdc, True)
    If (UnloadSource) Then Call FreeImage_Unload(BITMAP)
End Function
Public Function FreeImage_CreateFromOlePicture(ByRef Picture As IPicture) As LongPtr
' Creates a FreeImage DIB from a VB Picture object (OlePicture). This function
' returns a pointer to the DIB as, for instance, the FreeImage function
' 'FreeImage_Load' does. So, this could be a real replacement for 'FreeImage_Load'
' when working with VB Picture objects.
    If (Picture Is Nothing) Then Exit Function
Dim hBitmap As LongPtr: hBitmap = Picture.Handle: If (hBitmap = 0) Then Exit Function
Dim tBM As BITMAP_API
Dim lResult As Long: lResult = GetObject(hBitmap, LenB(tBM), tBM): If (lResult = 0) Then Exit Function
Dim hDib As LongPtr: hDib = FreeImage_Allocate(tBM.bmWidth, tBM.bmHeight, tBM.bmBitsPixel): If (hDib = 0) Then Exit Function
' The GetDIBits function clears the biClrUsed and biClrImportant BITMAPINFO
' members (dont't know why). So we save these infos below.
' This is needed for palletized images only.
Dim nColors As Long: nColors = FreeImage_GetColorsUsed(hDib)
Dim hdc As LongPtr: hdc = GetDC(0)
       lResult = GetDIBits(hdc, hBitmap, 0, FreeImage_GetHeight(hDib), FreeImage_GetBits(hDib), FreeImage_GetInfo(hDib), DIB_RGB_COLORS)
       If (lResult) Then
          FreeImage_CreateFromOlePicture = hDib
          If (nColors) Then
             ' restore BITMAPINFO members
Dim lpInfo As LongPtr: lpInfo = FreeImage_GetInfo(hDib)
             Call CopyMemory(ByVal lpInfo + 32, nColors, 4) ' FreeImage_GetInfo(Bitmap)->biClrUsed = nColors;
             Call CopyMemory(ByVal lpInfo + 36, nColors, 4) ' FreeImage_GetInfo(Bitmap)->biClrImportant = nColors;
          End If
       Else
          Call FreeImage_Unload(hDib)
       End If
       Call ReleaseDC(0, hdc)
End Function
Public Function FreeImage_CreateFromDC(ByVal hdc As LongPtr, Optional ByRef hBitmap As LongPtr) As LongPtr
' Creates a FreeImage DIB from a Device Context/Compatible Bitmap.
' returns a pointer to the DIB as, for instance, 'FreeImage_Load()' does.
' So, this could be a real replacement for FreeImage_Load() or
' 'FreeImage_CreateFromOlePicture()' when working with DCs and BITMAPs directly
' The 'hDC' parameter specifies a window device context (DC), the optional
' parameter 'hBitmap' may specify a handle to a memory bitmap. When 'hBitmap' is
' omitted, the bitmap currently selected into the given DC is used to create the DIB.
' When 'hBitmap' is not missing but NULL (0), the function uses the DC's currently
' selected bitmap. This bitmap's handle is stored in the ('ByRef'!) 'hBitmap' parameter
' and so, is avaliable at the caller's site when the function returns.
' The DIB returned by this function is a copy of the image specified by 'hBitmap' or
' the DC's current bitmap when 'hBitmap' is missing. The 'hDC' and also the 'hBitmap'
' remain untouched in this function, there will be no objects destroyed or freed.
' The caller is responsible to destroy or free the DC and BITMAP if necessary.
' first, check whether we got a hBitmap or not
   
Dim lResult As Long
    ' if not, the parameter may be missing or is NULL so get the DC's current bitmap
    If (hBitmap = 0) Then hBitmap = GetCurrentObject(hdc, OBJ_BITMAP)
Dim tBM As BITMAP_API
    lResult = GetObject(hBitmap, LenB(tBM), tBM): If lResult = 0 Then Exit Function
Dim hDib As LongPtr: hDib = FreeImage_Allocate(tBM.bmWidth, tBM.bmHeight, tBM.bmBitsPixel): If (hDib = 0) Then Exit Function
' The GetDIBits function clears the biClrUsed and biClrImportant BITMAPINFO
' members (dont't know why). So we save these infos below.
' This is needed for palletized images only.
Dim nColors As Long: nColors = FreeImage_GetColorsUsed(hDib)
    lResult = GetDIBits(hdc, hBitmap, 0, FreeImage_GetHeight(hDib), FreeImage_GetBits(hDib), FreeImage_GetInfo(hDib), DIB_RGB_COLORS)
    If (lResult) Then
       FreeImage_CreateFromDC = hDib
       If (nColors) Then
          ' restore BITMAPINFO members
Dim lpInfo As LongPtr: lpInfo = FreeImage_GetInfo(hDib)
          Call CopyMemory(ByVal lpInfo + 32, nColors, 4) ' FreeImage_GetInfo(Bitmap)->biClrUsed = nColors;
          Call CopyMemory(ByVal lpInfo + 36, nColors, 4) ' FreeImage_GetInfo(Bitmap)->biClrImportant = nColors;
       End If
    Else
       Call FreeImage_Unload(hDib)
    End If
End Function
Public Function FreeImage_CreateFromImageContainer(ByRef Container As Object, Optional ByVal IncludeDrawings As Boolean) As LongPtr
' Creates a FreeImage DIB from a VB container control that has at least a 'Picture' property.
' This function returns a pointer to the DIB as, for instance, 'FreeImage_Load()' does.
' So, this could be a real replacement for FreeImage_Load() or 'FreeImage_CreateFromOlePicture()'
' when working with image hosting controls like Forms or PictureBoxes.
' The 'IncludeDrawings' parameter controls whether drawings, drawn with VB
' methods like 'Container.Print()', 'Container.Line(x1, y1)-(x2, y2)' or
' 'Container.Circle(x, y), radius' as the controls 'BackColor' should be included
' into the newly created DIB. However, this only works, with control's that
' have their 'AutoRedraw' property set to 'True'.
' To get the control's picture as well as it's BackColor and custom drawings,
' uses the control's 'Image' property instead of the 'Picture' property.
' treats Forms and PictureBox controls explicitly, since the
' property sets and behaviours of these controls are publicly known.
' For any other control, the function checks for the existence of an 'Image' and
' 'AutoRedraw' property. If these are present and 'IncludeDrawings' is 'True',
' the function uses the control's 'Image' property instead of the 'Picture'
' property. This my be the case for UserControls. In any other case, the function
' uses the control's 'Picture' property if present. If none of these properties
' is present, a runtime error (5) is generated.
' Most of this function is actually implemented in the wrapper's private helper
' function 'p_GetIOlePictureFromContainer'.
    If (Not Container Is Nothing) Then FreeImage_CreateFromImageContainer = FreeImage_CreateFromOlePicture(p_GetIOlePictureFromContainer(Container, IncludeDrawings))
End Function
Public Function FreeImage_CreateFromScreen(Optional ByVal hwnd As LongPtr, Optional ByVal Left As Long, Optional ByVal Top As Long, Optional ByVal Width As Long, Optional ByVal Height As Long, Optional ByVal ClientAreaOnly As Boolean) As LongPtr
' Creates a FreeImage DIB from the screen which may either be the whole
' desktop/screen or a certain window. A certain window may be specified
' by it's window handle through the 'hWnd' parameter. By omitting this
' parameter, the whole screen/desktop window will be captured.
' hWnd - Window handler
' Left/Top, Width/Height - position in window and size of copied region
' ClientAreaOnly - to use only client area of certain window
Dim hdc As LongPtr, hMemDC As LongPtr, hMemBMP As LongPtr, hMemOldBMP As LongPtr
Dim tR As RECT
    If (hwnd = 0) Then
' get Desktop client area DC
        hwnd = GetDesktopWindow()
        hdc = GetDC(hwnd)         ' hDC = GetDCEx(hWnd, 0, 0)
        ' get width and height for desktop window
        If Width <= 0 Then Width = GetDeviceCaps(hdc, HORZRES) - Left
        If Height <= 0 Then Height = GetDeviceCaps(hdc, VERTRES) - Top
    ElseIf (ClientAreaOnly) Then
' get window's client area DC
        hdc = GetDC(hwnd)         ' hDC = GetDCEx(hWnd, 0, 0)
        Call GetClientRect(hwnd, tR)
        If Width <= 0 Then Width = tR.Right - Left
        If Height <= 0 Then Height = tR.Bottom - Top
    Else
' get window DC
        hdc = GetWindowDC(hwnd)   ' hDC = GetDCEx(hWnd, 0, DCX_WINDOW)
        Call GetWindowRect(hwnd, tR)
        If Width <= 0 Then Width = tR.Right - tR.Left - Left
        If Height <= 0 Then Height = tR.Bottom - tR.Top - Top
    End If
' create compatible memory DC and bitmap
    hMemDC = CreateCompatibleDC(hdc)
    hMemBMP = CreateCompatibleBitmap(hdc, Width, Height)
' select compatible bitmap
    hMemOldBMP = SelectObject(hMemDC, hMemBMP)
' blit bits
    'hDestDC - контекст устpойства, пpинимающего каpту бит.
    'x, y - веpхний левый угол пpямоугольника назначения.
    'nWidth - шиpина пpямоугольника назначения и каpты бит источника.
    'nHeight - высота пpямоугольника назначения и каpты бит источника.
    'hSrcDC - контекст устpойства, их котоpого копиpуется каpта бит, или ноль для pастpовой опеpации только на DestDC.
    'xSrc, ySrc - веpхний левый угол SrcDC.
    'dwRop - растровая операция: BLACKNESS, DSTINVERT, MERGECOPY, MERGEPAINT, NOTSRCCOPY, NOTSRCERASE, PATCOPY, PATINVERT, PATPAINT, SRCAND, SRCCOPY, SRCERASE, SRCINVERT, SRCPAINT, WHITNESS.
    Call BitBlt(hMemDC, 0, 0, Width, Height, hdc, Left, Top, SRCCOPY Or CAPTUREBLT)
' create FreeImage Bitmap from memory DC
    FreeImage_CreateFromScreen = FreeImage_CreateFromDC(hMemDC, hMemBMP)
    ' clean up
    Call SelectObject(hMemDC, hMemOldBMP)
    Call DeleteObject(hMemBMP)
    Call DeleteDC(hMemDC)
    Call ReleaseDC(hwnd, hdc)
End Function
'----------------------
' Microsoft Office / VBA PictureData supporting functions
'----------------------
Public Function FreeImage_GetPictureData(ByVal BITMAP As LongPtr, Optional ByVal UnloadSource As Boolean) As Byte()
' Creates an Office PictureData(DIB) Byte array from a FreeImage DIB.
' This format is suitable for Access.CommandButton.PictureData (not supports transparent color)
' The original image must not remain valid nor loaded after the PictureData array has been created.
' The optional 'UnloadSource' parameter is for unloading the original image
' after the PictureData Byte array has been created, so you can easily "switch"
' from a FreeImage DIB to an Office PictureData Byte array. There is no need to
' unload the DIB at the caller's site if this argument is True.
Dim lpInfo As LongPtr
Const SIZE_OF_LONG = 4
Const SIZE_OF_BITMAPINFOHEADER = &H28
Dim abResult() As Byte
    If BITMAP = 0 Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to create a PictureData array from a 'header-only' bitmap.")
Dim lHeaderSize As Long:    If (FreeImage_HasRGBMasks(BITMAP)) Then lHeaderSize = 3 * SIZE_OF_LONG
                            lHeaderSize = lHeaderSize + SIZE_OF_BITMAPINFOHEADER
Dim lImageSize As Long:     lImageSize = FreeImage_GetHeight(BITMAP) * FreeImage_GetPitch(BITMAP)
Dim lPaletteSize As Long:   lPaletteSize = FreeImage_GetColorsUsed(BITMAP) * 4 ' lPaletteSize = 0 - HighColor
Dim lOffset As Long:        lOffset = lOffset + lHeaderSize
    lpInfo = FreeImage_GetInfo(BITMAP)
' PictureData (DIB) is a BMP w/o BITMAPFILEHEADER &hE(14) DIB begins w BITMAPINFOHEADER
    ReDim abResult(0 To lHeaderSize + lPaletteSize + lImageSize - 1)
' BITMAPINFOHEADER ' 0-3   - format identifier, CF_DIBFILE = &h00000028(40)
    ' Copy the BITMAPINFOHEADER into the result array.
    Call CopyMemory(abResult(0), ByVal lpInfo, lHeaderSize)
    ' Copy the image's palette (if any) into the result array.
    If (lPaletteSize > 0) Then Call CopyMemory(abResult(lOffset), ByVal FreeImage_GetPalette(BITMAP), lPaletteSize): lOffset = lOffset + lPaletteSize
    ' Copy the image's bits into the result array.
    Call CopyMemory(abResult(lOffset), ByVal FreeImage_GetBits(BITMAP), lImageSize)
    Call p_Swap(ByVal VarPtrArray(abResult), ByVal VarPtrArray(FreeImage_GetPictureData))
    If (UnloadSource) Then FreeImage_Unload (BITMAP)
End Function
Public Function FreeImage_GetPictureDataEMF(ByVal BITMAP As LongPtr, Optional ByVal UnloadSource As Boolean) As Byte()
' Creates an Office PictureData(EMF) Byte array from a FreeImage DIB.
' This format is suitable for Access.Image.PictureData (supports transparent color)
' Convert BITMAP to EMF OLEPicture
Dim cPictureDisp As IPictureDisp: Set cPictureDisp = FreeImage_GetOlePictureEMF(BITMAP)
                    If cPictureDisp Is Nothing Then Exit Function
                    If cPictureDisp.Handle = 0 Then Exit Function
                    If cPictureDisp.Type <> PICTYPE_ENHMETAFILE Then Exit Function
Dim hEmf As Long:   hEmf = cPictureDisp.Handle
Dim cbSize As Long: cbSize = GetEnhMetaFileBits(hEmf, 0, ByVal 0&)
Dim abResult() As Byte, cbCopied As Long
' Convert to EMF PictureData
    ReDim abResult(0 To cbSize + 7) As Byte
    CopyMemory abResult(&H0), CF_ENHMETAFILE, 4&                    ' 0-3 - Format identifier, CF_ENHMETAFILE = &h0000000e(14)
                                                                    '       must be Long (!!!) check constant
    CopyMemory abResult(&H4), hEmf, 4&                              ' 4-7 - Metafile handle, OLE_HANDLE (32bit even in x64)
    cbCopied = GetEnhMetaFileBits(hEmf, cbSize, abResult(&H8))      ' 8-..  - Metafile body
    Call p_Swap(ByVal VarPtrArray(abResult), ByVal VarPtrArray(FreeImage_GetPictureDataEMF))
    If (UnloadSource) Then FreeImage_Unload (BITMAP)
End Function
Public Function FreeImage_GetIconBestMatch(ByVal MULTIBITMAP As LongPtr, Optional ByVal Width As Long, Optional ByVal Height As Long, Optional ByVal BitDepth As Long = 32, Optional BestMatch As Long, Optional ByVal UnloadSource As Boolean) As LongPtr
' Find the nearest match to the passed Size.
' based on cICOparser.GetBestMatch function by LaVolpe (from psc cd)
' the weighting is customized to favor larger icons over smaller ones
' when stretching would be needed. The thought is that stretching down almost always
' produces better quality graphics than stretching up.
Const cScrDepth = 32&       ' screen color depth
' Note that this routine is weighted for monitors set at 32bit.
' If this is not acceptable, then modify algorithm slightly
'   from adding weight of:  Abs(32 - bitDepth(lPage))
'   to adding weight of: Abs([ScreenColorDepth] - bitDepth(lPage))
Const cNumb = &H80000000    ' least desirable weight: some large number
Const cBase = cNumb \ 4     ' base weight to shift icons exceeding desirable size to the bottom of list

    On Error GoTo HandleError
    If MULTIBITMAP = 0 Then Exit Function
' vars for bitmap handle
Dim hDib As LongPtr, hMax As LongPtr, hTmp As LongPtr
' vars for weight values, start max value is lowest possible number
Dim lVal As Long, lMax As Long, lTmp As Long
Dim lPage As Long, lPages As Long
    lVal = cNumb: lMax = cNumb
    lPage = 0: BestMatch = lPage
    lPages = FreeImage_GetPageCount(MULTIBITMAP) - 1 ' -1 is only for not to do this in cycle
' if more then one page - check conditions
    If lPages > 0 Then
        If (BitDepth <= 0) Then BitDepth = cScrDepth
    ' if set only one dimension make suppose target region square
        If ((Width > 0) And (Height > 0)) Then
        ElseIf (Height > 0) Then Width = Height
        ElseIf (Width > 0) Then Height = Width
        Else:   Width = 0: Height = 0
        End If
    End If
' Run through icons and select best
    Do
' Lock page and get its handle
    ' crash on LockPage if Memory Stream was closed after opening MULTIBITMAP. Exception code: 0xc0000005 - (memory access violation)
        hDib = FreeImage_LockPage(MULTIBITMAP, lPage): If hDib = 0 Then GoTo HandleNext
' Calculate weight - get icon params and check compliance with the specified params
    ' skip if nothing to compare (single page or last but only correct in icon)
        If (lPage <> lPages) Or (hMax > 0) Then
        ' bitdepth
            lTmp = FreeImage_GetBPP(hDib): If lTmp = 0 Then GoTo HandleNext   ' if a image within icon file is faulty, we ignore it
        ' if prefer greatest icon under desirable size
        ' then change next to: lTmp = BitDepth - lTmp
            lTmp = lTmp - BitDepth ' << differ from below to make delta negative
        ' if prefer greatest icon under desirable size
        ' then change next and below to: If lTmp < 0 Then lTmp = lTmp + cBase  Else lTmp = -lTmp
            If lTmp > 0 Then lTmp = -lTmp + cBase ' Else lTmp = lTmp
            lVal = lTmp
        ' width
            lTmp = FreeImage_GetWidth(hDib): lTmp = Width - lTmp
            If lTmp > 0 Then lTmp = -lTmp + cBase ' Else lTmp = lTmp
            lVal = lVal + lTmp
        ' height
            lTmp = FreeImage_GetHeight(hDib): lTmp = Height - lTmp
            If lTmp > 0 Then lTmp = -lTmp + cBase ' Else lTmp = lTmp
            lVal = lVal + lTmp
        End If
' compare; biggest is best
        ' if only one page in icon then really need to do: hMax = hDib
        If lVal >= lMax Then                        ' best  match
            BestMatch = lPage: lMax = lVal          ' remember best page and weight
            hTmp = hMax: hMax = hDib: hDib = hTmp   ' swap bitmap handles
        End If
HandleNext:
        ' close page if has better
        If hDib <> 0 Then Call FreeImage_UnlockPage(MULTIBITMAP, hDib, 0)         ' close if not best
        If lVal = 0 Then Exit Do                    ' exact match exit
        If lPage = lPages Then Exit Do              ' last page exit
        lPage = lPage + 1
    Loop
HandleExit:  FreeImage_GetIconBestMatch = hMax: Exit Function
HandleError: hDib = False: Err.Clear: Resume HandleExit
End Function
Public Function FreeImage_CreateFromPictureData(ByRef PictureData() As Byte) As LongPtr
' Creates a FreeImage DIB from an Office PictureData Byte array.
' This function returns a pointer to the DIB as, for instance, the FreeImage function 'FreeImage_Load' does.
' So, this could be a real replacement for 'FreeImage_Load' when working with PictureData arrays.
Dim hDib As LongPtr
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
Dim tBMIH As BITMAPINFOHEADER
Dim lLength As Long: lLength = UBound(PictureData) + 1
   'If (lLength <= LenB(tBMIH)) Then Err.Raise vbObjectError + 512
   'If (lLength <= BITMAPINFOHEADERSIZE) Then Err.Raise vbObjectError + 512
   If (lLength > BITMAPINFOHEADERSIZE) Then
Dim lPaletteSize As Long
Dim lOffset As Long
Dim alMasks() As Long
    Call CopyMemory(tBMIH, PictureData(0), BITMAPINFOHEADERSIZE) 'LenB(tBMIH))
    With tBMIH
        'If (.biSize <> BITMAPINFOHEADERSIZE) Then Err.Raise vbObjectError + 512
        If (.biSize = BITMAPINFOHEADERSIZE) Then
            lOffset = BITMAPINFOHEADERSIZE
            Select Case .biBitCount
            Case 0
            Case 1, 4, 8
               If (.biClrUsed = 0) Then
                  lPaletteSize = 2 ^ .biBitCount * 4
               Else
                  lPaletteSize = .biClrUsed * 4
               End If
               hDib = FreeImage_Allocate(.biWidth, .biHeight, .biBitCount, 0, 0, 0)
               Call CopyMemory(ByVal FreeImage_GetPalette(hDib), PictureData(lOffset), lPaletteSize)
               lOffset = lOffset + lPaletteSize
            Case 16
               If (.biCompression = BI_BITFIELDS) Then
                  ReDim alMasks(2)
                  Call CopyMemory(alMasks(0), PictureData(lOffset), 12)
                  lOffset = lOffset + 12
                  hDib = FreeImage_Allocate(.biWidth, .biHeight, .biBitCount, alMasks(0), alMasks(1), alMasks(2))
               Else
                  hDib = FreeImage_Allocate(.biWidth, .biHeight, .biBitCount, FI16_555_RED_MASK, FI16_555_GREEN_MASK, FI16_555_BLUE_MASK)
               End If
            Case 24, 32
               hDib = FreeImage_Allocate(.biWidth, .biHeight, .biBitCount, 0, 0, 0)
            End Select
            If (hDib) Then
               Call CopyMemory(ByVal FreeImage_GetBits(hDib), PictureData(lOffset), lLength - lOffset)
               FreeImage_CreateFromPictureData = hDib
            End If
        End If
    End With
    End If
'' size picture
'    If width = 0 Then GoTo HandleExit
'    'pResult = FreeImage_RescaleByPixel(pResult, Width, Height, True, Filter)
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function FreeImage_CreateMask(ByVal hDib As LongPtr, Optional ByVal eMaskCreationOptions As FREE_IMAGE_MASK_CREATION_OPTION_FLAGS = MCOF_CREATE_MASK_IMAGE, Optional ByVal lBitDepth As Long = 1, Optional ByVal eMaskOptions As FREE_IMAGE_MASK_FLAGS = FIMF_MASK_FULL_TRANSPARENCY, Optional ByVal vntMaskColors As Variant, Optional ByVal eMaskColorsFormat As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB, Optional ByVal lColorTolerance As Long, Optional ByVal lciMaskColorDst As Long = vbWhite, Optional ByVal eMaskColorDstFormat As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB, Optional ByVal lciUnmaskColorDst As Long = vbBlack, Optional ByVal eUnmaskColorDstFormat As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB, Optional ByVal vlciMaskColorSrc As Variant, Optional ByVal eMaskColorSrcFormat As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB, Optional ByVal vlciUnmaskColorSrc As Variant, Optional ByVal eUnmaskColorSrcFormat As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB) As LongPtr
'
Dim hDIBResult As LongPtr
Dim lBitDepthSrc As Long
Dim lWidth As Long
Dim lHeight As Long
Dim bMaskColors As Boolean
Dim bMaskTransparency As Boolean
Dim bMaskFullTransparency As Boolean
Dim bMaskAlphaTransparency As Boolean
Dim bInvertMask As Boolean
Dim bHaveMaskColorSrc As Boolean
Dim bHaveUnmaskColorSrc As Boolean
Dim bCreateMaskImage As Boolean
Dim bModifySourceImage As Boolean
Dim alcMaskColors() As Long
Dim lMaskColorsMaxIndex As Long
Dim lciMaskColorSrc As Long
Dim lciUnmaskColorSrc As Long
Dim alPaletteSrc() As Long
Dim abTransparencyTableSrc() As Byte
Dim abBitsBSrc() As Byte
Dim atBitsTSrc As ScanLinesRGBTRIBLE
Dim atBitsQSrc() As RGBQUAD
Dim abBitValues(7) As Byte
Dim abBitMasks(7) As Byte
Dim abBitShifts(7) As Byte
Dim atPaletteDst() As RGBQUAD
Dim abBitsBDst() As Byte
Dim atBitsTDst As ScanLinesRGBTRIBLE
Dim atBitsQDst() As RGBQUAD
Dim bMaskPixel As Boolean
Dim x As Long
Dim X2 As Long
Dim lPixelIndex As Long
Dim y As Long
Dim i As Long
   ' check for a proper bit depth of the destination (mask) image
   If ((hDib) And ((lBitDepth = 1) Or (lBitDepth = 4) Or (lBitDepth = 8) Or (lBitDepth = 24) Or (lBitDepth = 32))) Then
      If (Not FreeImage_HasPixels(hDib)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to create a mask image from a 'header-only' bitmap.")
      ' check for a proper bit depth of the source image
      lBitDepthSrc = FreeImage_GetBPP(hDib)
      If ((lBitDepthSrc = 4) Or (lBitDepthSrc = 8) Or (lBitDepthSrc = 24) Or (lBitDepthSrc = 32)) Then
         
         ' get some information from eMaskCreationOptions
         bCreateMaskImage = (eMaskCreationOptions And MCOF_CREATE_MASK_IMAGE)
         bModifySourceImage = (eMaskCreationOptions And MCOF_MODIFY_SOURCE_IMAGE)
         
         If (bCreateMaskImage) Then
            ' check mask color format
            If (eMaskColorDstFormat And FICFF_COLOR_BGR) Then
               ' if mask color is in BGR format, convert to RGB format
               lciMaskColorDst = FreeImage_SwapColorLong(lciMaskColorDst)
            ElseIf (eMaskColorDstFormat And FICFF_COLOR_PALETTE_INDEX) Then
               ' if mask color is specified as palette index, check, whether the
               ' source image is a palletized image
               Select Case lBitDepthSrc
               Case 1: lciMaskColorDst = FreeImage_GetPaletteExLong(hDib)(lciMaskColorDst And &H1)
               Case 4: lciMaskColorDst = FreeImage_GetPaletteExLong(hDib)(lciMaskColorDst And &HF)
               Case 8: lciMaskColorDst = FreeImage_GetPaletteExLong(hDib)(lciMaskColorDst And &HFF)
               End Select
            End If
            ' check unmask color format
            If (eUnmaskColorDstFormat And FICFF_COLOR_BGR) Then
               ' if unmask color is in BGR format, convert to RGB format
               lciUnmaskColorDst = FreeImage_SwapColorLong(lciUnmaskColorDst)
            ElseIf (eUnmaskColorDstFormat And FICFF_COLOR_PALETTE_INDEX) Then
               ' if unmask color is specified as palette index, check, whether the
               ' source image is a palletized image
               Select Case lBitDepthSrc
               Case 1: lciUnmaskColorDst = FreeImage_GetPaletteExLong(hDib)(lciUnmaskColorDst And &H1)
               Case 4: lciUnmaskColorDst = FreeImage_GetPaletteExLong(hDib)(lciUnmaskColorDst And &HF)
               Case 8: lciUnmaskColorDst = FreeImage_GetPaletteExLong(hDib)(lciUnmaskColorDst And &HFF)
               End Select
            End If
         End If
         
         If (bModifySourceImage) Then
            ' check, whether source image can be modified
            bHaveMaskColorSrc = (Not IsMissing(vlciMaskColorSrc))
            bHaveUnmaskColorSrc = (Not IsMissing(vlciUnmaskColorSrc))
            Select Case lBitDepthSrc
            Case 4, 8
               If (bHaveMaskColorSrc) Then
                  ' get mask color as Long
                  lciMaskColorSrc = vlciMaskColorSrc
                  If (eMaskColorSrcFormat And FICFF_COLOR_PALETTE_INDEX) Then
                     If (lBitDepthSrc = 4) Then
                        lciMaskColorSrc = (lciMaskColorSrc And &HF)
                     Else
                        lciMaskColorSrc = (lciMaskColorSrc And &HFF)
                     End If
                  Else
                     If (eMaskColorSrcFormat And FICFF_COLOR_BGR) Then lciMaskColorSrc = FreeImage_SwapColorLong(lciMaskColorSrc, True)
                     lciMaskColorSrc = FreeImage_SearchPalette(hDib, lciMaskColorSrc)
                     bHaveMaskColorSrc = (lciMaskColorSrc <> -1)
                  End If
               End If
               If (bHaveUnmaskColorSrc) Then
                  ' get unmask color as Long
                  lciUnmaskColorSrc = vlciUnmaskColorSrc
                  If (eUnmaskColorSrcFormat And FICFF_COLOR_PALETTE_INDEX) Then
                     If (lBitDepthSrc = 4) Then
                        lciUnmaskColorSrc = (lciUnmaskColorSrc And &HF)
                     Else
                        lciUnmaskColorSrc = (lciUnmaskColorSrc And &HFF)
                     End If
                  Else
                     If (eUnmaskColorSrcFormat And FICFF_COLOR_BGR) Then
                        lciUnmaskColorSrc = FreeImage_SwapColorLong(lciUnmaskColorSrc, True)
                     End If
                     lciUnmaskColorSrc = FreeImage_SearchPalette(hDib, lciUnmaskColorSrc)
                     bHaveUnmaskColorSrc = (lciUnmaskColorSrc <> -1)
                  End If
               End If
               ' check, if source image still can be modified in any way
               bModifySourceImage = (bHaveMaskColorSrc Or bHaveUnmaskColorSrc)
            Case 24, 32
               If (bHaveMaskColorSrc) Then
                  ' get mask color as Long
                  lciMaskColorSrc = vlciMaskColorSrc
                  If (eMaskColorSrcFormat And FICFF_COLOR_BGR) Then lciMaskColorSrc = FreeImage_SwapColorLong(lciMaskColorSrc, (lBitDepthSrc = 24))
               End If
               If (bHaveUnmaskColorSrc) Then
                  ' get unmask color as Long
                  lciUnmaskColorSrc = vlciUnmaskColorSrc
                  If (eUnmaskColorSrcFormat And FICFF_COLOR_BGR) Then lciUnmaskColorSrc = FreeImage_SwapColorLong(lciUnmaskColorSrc, (lBitDepthSrc = 24))
               End If
            End Select
         End If
          
         If ((bModifySourceImage) Or (bCreateMaskImage)) Then
            ' get some information from eMaskOptions
            ' check for inverse mask
            bInvertMask = (eMaskOptions And FIMF_MASK_INVERSE_MASK)
            ' check for mask colors
            bMaskColors = (eMaskOptions And FIMF_MASK_COLOR_TRANSPARENCY)
            bMaskColors = bMaskColors And (Not IsMissing(vntMaskColors))
            If (bMaskColors) Then
               ' validate specified mask colors; all mask colors are transferred to
               ' an internal array of type Long
               If (Not IsArray(vntMaskColors)) Then
                  ' color masking is only done when the single mask color is
                  ' a numeric (color) value
                  bMaskColors = IsNumeric(vntMaskColors)
                  If (bMaskColors) Then
                     ' this is not an array of mask colors but only a single
                     ' color; this is also transferred into an internal array
                     lMaskColorsMaxIndex = 0
                     ReDim alcMaskColors(lMaskColorsMaxIndex)
                     alcMaskColors(lMaskColorsMaxIndex) = vntMaskColors
                  End If
               Else
                  ' transfer all valid color values (numeric) into an internal
                  ' array
                  ReDim alcMaskColors(UBound(vntMaskColors))
                  For i = LBound(vntMaskColors) To UBound(vntMaskColors)
                     bMaskColors = (IsNumeric(vntMaskColors(i)))
                     If (Not bMaskColors) Then
                        Exit For
                     Else
                        alcMaskColors(lMaskColorsMaxIndex) = vntMaskColors(i)
                        lMaskColorsMaxIndex = lMaskColorsMaxIndex + 1
                     End If
                  Next i
                  If (bMaskColors) Then lMaskColorsMaxIndex = lMaskColorsMaxIndex - 1
               End If
            End If
            ' check for transparency options
            If ((FreeImage_IsTransparent(hDib)) Or ((eMaskOptions And FIMF_MASK_FORCE_TRANSPARENCY) > 0)) Then
               bMaskFullTransparency = (eMaskOptions And FIMF_MASK_FULL_TRANSPARENCY)
               bMaskAlphaTransparency = (eMaskOptions And FIMF_MASK_ALPHA_TRANSPARENCY)
               bMaskTransparency = (bMaskFullTransparency Or bMaskAlphaTransparency)
            End If
            ' get image dimension
            lWidth = FreeImage_GetWidth(hDib)
            lHeight = FreeImage_GetHeight(hDib)
            ' create proper accessors for the source image
            Select Case lBitDepthSrc
            Case 4, 8 ' images with a bit depth of 4 or 8 bits will both be read through a byte array
               abBitsBSrc = FreeImage_GetBitsEx(hDib)
               ' depending on where to get the transparency information from,
               ' a palette or a transpareny table will be needed
               If (bMaskColors) Then alPaletteSrc = FreeImage_GetPaletteExLong(hDib)
               If (bMaskTransparency) Then abTransparencyTableSrc = FreeImage_GetTransparencyTableExClone(hDib)
               ' for 4 bit source images
               If (lBitDepthSrc = 4) Then
                  ' two additional arrays need to be filled with values
                  ' to mask and shift nibbles to bytes
                  ' index 0 stands for the high nibble of the byte
                  abBitMasks(0) = &HF0
                  abBitShifts(0) = &H10 ' a shift to right is implemented
                                        ' as division in VB
                  ' index 1 stands for the low nibble of the byte
                  abBitMasks(1) = &HF
                  abBitShifts(1) = &H1 ' no shift needed for low nibble
               End If
            Case 24: Call FreeImage_GetScanLinesRGBTRIPLE(hDib, atBitsTSrc)
               ' images with a depth of 24 bits could not be used
               ' through a two dimensional array in most cases, so get
               ' an array of individual scanlines (see remarks concerning
               ' pitch at function 'FreeImage_GetBitsExRGBTriple()')
            Case 32: atBitsQSrc = FreeImage_GetBitsExRGBQUAD(hDib)
            End Select
      
            ' create mask image if needed
            If (bCreateMaskImage) Then
               ' create mask image
               hDIBResult = FreeImage_Allocate(lWidth, lHeight, lBitDepth)
               ' if destination bit depth is 8 or below, a proper palette will
               ' be needed, so create a palette where the unmask color is at
               ' index 0 and the mask color is at index 1
               If (lBitDepth <= 8) Then
                  atPaletteDst = FreeImage_GetPaletteEx(hDIBResult)
                  Call CopyMemory(atPaletteDst(0), lciUnmaskColorDst, 4)
                  Call CopyMemory(atPaletteDst(1), lciMaskColorDst, 4)
               End If
               ' create proper accessors for the new mask image
               Select Case lBitDepth
               Case 1
                  abBitsBDst = FreeImage_GetBitsEx(hDIBResult)
                  x = 1
                  For i = 7 To 0 Step -1
                     abBitValues(i) = x
                     x = x * 2
                  Next i
               Case 4: abBitsBDst = FreeImage_GetBitsEx(hDIBResult): abBitValues(0) = &H10: abBitValues(1) = &H1
               Case 8: abBitsBDst = FreeImage_GetBitsEx(hDIBResult)
               Case 24: Call FreeImage_GetScanLinesRGBTRIPLE(hDIBResult, atBitsTDst)
                  ' images with a depth of 24 bits could not be used
                  ' through a two dimensional array in most cases, so get
                  ' an array of individual scanlines (see remarks concerning
                  ' pitch at function 'FreeImage_GetBitsExRGBTriple()')
               Case 32: atBitsQDst = FreeImage_GetBitsExRGBQUAD(hDIBResult)
               End Select
            End If
            ' walk the hole image
            For y = 0 To lHeight - 1
               For x = 0 To lWidth - 1
                  ' should transparency information be considered to create
                  ' the mask?
                  If (bMaskTransparency) Then
                     Select Case lBitDepthSrc
                     Case 4
                        X2 = x \ 2
                        lPixelIndex = (abBitsBSrc(X2, y) And abBitMasks(x Mod 2)) \ abBitShifts(x Mod 2)
                        bMaskPixel = (abTransparencyTableSrc(lPixelIndex) = 0)
                        If (Not bMaskPixel) Then bMaskPixel = ((abTransparencyTableSrc(lPixelIndex) < 255) And (bMaskAlphaTransparency))
                     Case 8
                        bMaskPixel = (abTransparencyTableSrc(abBitsBSrc(x, y)) = 0)
                        If (Not bMaskPixel) Then bMaskPixel = ((abTransparencyTableSrc(abBitsBSrc(x, y)) < 255) And (bMaskAlphaTransparency))
                     Case 24 ' no transparency information in 24 bit images reset bMaskPixel
                        bMaskPixel = False
                     Case 32
                        bMaskPixel = (atBitsQSrc(x, y).rgbReserved = 0)
                        If (Not bMaskPixel) Then bMaskPixel = ((atBitsQSrc(x, y).rgbReserved < 255) And (bMaskAlphaTransparency))
                     End Select
                  Else
                     ' clear 'bMaskPixel' if no transparency information was checked
                     ' since the flag might be still True from the last loop
                     bMaskPixel = False
                  End If
                  ' should color information be considered to create the mask?
                  ' do this only if the current pixel is not yet part of the mask
                  If ((bMaskColors) And (Not bMaskPixel)) Then
                     Select Case lBitDepthSrc
                     Case 4
                        X2 = x \ 2
                        lPixelIndex = (abBitsBSrc(X2, y) And abBitMasks(x Mod 2)) \ abBitShifts(x Mod 2)
                        If (eMaskColorsFormat And FICFF_COLOR_PALETTE_INDEX) Then
                           For i = 0 To lMaskColorsMaxIndex
                              If (lColorTolerance = 0) Then
                                 bMaskPixel = (lPixelIndex = alcMaskColors(i))
                              Else
                                 bMaskPixel = (FreeImage_CompareColorsLongLong(alPaletteSrc(lPixelIndex), alPaletteSrc(alcMaskColors(i)), lColorTolerance, FICFF_COLOR_RGB, FICFF_COLOR_RGB) = 0)
                              End If
                              If (bMaskPixel) Then Exit For
                           Next i
                        Else
                           For i = 0 To lMaskColorsMaxIndex
                              bMaskPixel = (FreeImage_CompareColorsLongLong(alPaletteSrc(lPixelIndex), alcMaskColors(i), lColorTolerance, FICFF_COLOR_RGB, (eMaskColorsFormat And FICFF_COLOR_FORMAT_ORDER_MASK)) = 0)
                              If (bMaskPixel) Then
                                 Exit For
                              End If
                           Next i
                        End If
                     Case 8
                        If (eMaskColorsFormat And FICFF_COLOR_PALETTE_INDEX) Then
                           For i = 0 To lMaskColorsMaxIndex
                              If (lColorTolerance = 0) Then
                                 bMaskPixel = (abBitsBSrc(x, y) = alcMaskColors(i))
                              Else
                                 bMaskPixel = (FreeImage_CompareColorsLongLong(alPaletteSrc(abBitsBSrc(x, y)), alPaletteSrc(alcMaskColors(i)), lColorTolerance, FICFF_COLOR_RGB, FICFF_COLOR_RGB) = 0)
                              End If
                              If (bMaskPixel) Then Exit For
                           Next i
                        Else
                           For i = 0 To lMaskColorsMaxIndex
                              bMaskPixel = (FreeImage_CompareColorsLongLong(alPaletteSrc(abBitsBSrc(x, y)), alcMaskColors(i), lColorTolerance, FICFF_COLOR_RGB, (eMaskColorsFormat And FICFF_COLOR_FORMAT_ORDER_MASK)) = 0)
                              If (bMaskPixel) Then Exit For
                           Next i
                        End If
                     Case 24
                        For i = 0 To lMaskColorsMaxIndex
                           bMaskPixel = (FreeImage_CompareColorsRGBTRIPLELong(atBitsTSrc.Scanline(y).Data(x), alcMaskColors(i), lColorTolerance, (eMaskColorsFormat And FICFF_COLOR_FORMAT_ORDER_MASK)) = 0)
                           If (bMaskPixel) Then Exit For
                        Next i
                     Case 32
                        For i = 0 To lMaskColorsMaxIndex
                           bMaskPixel = (FreeImage_CompareColorsRGBQUADLong(atBitsQSrc(x, y), alcMaskColors(i), lColorTolerance, (eMaskColorsFormat And FICFF_COLOR_FORMAT_ORDER_MASK)) = 0)
                           If (bMaskPixel) Then Exit For
                        Next i
                     End Select
                  End If
                  ' check whether a mask image needs to be created
                  If (bCreateMaskImage) Then
                     ' write current pixel to destination (mask) image
                     Select Case lBitDepth
                     Case 1: X2 = x \ 8: If ((bMaskPixel) Xor (bInvertMask)) Then abBitsBDst(X2, y) = abBitsBDst(X2, y) Or abBitValues(x Mod 8)
                     Case 4: X2 = x \ 2: If ((bMaskPixel) Xor (bInvertMask)) Then abBitsBDst(X2, y) = abBitsBDst(X2, y) Or abBitValues(x Mod 2)
                     Case 8: If ((bMaskPixel) Xor (bInvertMask)) Then abBitsBDst(x, y) = 1
                     Case 24: If ((bMaskPixel) Xor (bInvertMask)) Then Call CopyMemory(atBitsTDst.Scanline(y).Data(x), lciMaskColorDst, 3) Else Call CopyMemory(atBitsTDst.Scanline(y).Data(x), lciUnmaskColorDst, 3)
                     Case 32: If ((bMaskPixel) Xor (bInvertMask)) Then Call CopyMemory(atBitsQDst(x, y), lciMaskColorDst, 4) Else Call CopyMemory(atBitsQDst(x, y), lciUnmaskColorDst, 4)
                     End Select
                  End If
                  ' check whether a source image needs to be modified
                  If (bModifySourceImage) Then
                     Select Case lBitDepthSrc
                     Case 4
                        X2 = x \ 2
                        If ((bMaskPixel) Xor (bInvertMask)) Then
                           If (bHaveMaskColorSrc) Then abBitsBSrc(X2, y) = (abBitsBSrc(X2, y) And (Not abBitMasks(x Mod 2))) Or (lciMaskColorSrc * abBitShifts(x Mod 2))
                        ElseIf (bHaveUnmaskColorSrc) Then
                           abBitsBSrc(X2, y) = (abBitsBSrc(X2, y) And (Not abBitMasks(x Mod 2))) Or (lciUnmaskColorSrc * abBitShifts(x Mod 2))
                        End If
                     Case 8
                        If ((bMaskPixel) Xor (bInvertMask)) Then
                           If (bHaveMaskColorSrc) Then abBitsBSrc(x, y) = lciMaskColorSrc
                        ElseIf (bHaveUnmaskColorSrc) Then
                           abBitsBSrc(x, y) = lciUnmaskColorSrc
                        End If
                     Case 24
                        If ((bMaskPixel) Xor (bInvertMask)) Then
                           If (bHaveMaskColorSrc) Then Call CopyMemory(atBitsTSrc.Scanline(y).Data(x), lciMaskColorSrc, 3)
                        ElseIf (bHaveUnmaskColorSrc) Then
                           Call CopyMemory(atBitsTSrc.Scanline(y).Data(x), lciUnmaskColorSrc, 3)
                        End If
                     Case 32
                        If ((bMaskPixel) Xor (bInvertMask)) Then
                           If (bHaveMaskColorSrc) Then Call CopyMemory(atBitsQSrc(x, y), lciMaskColorSrc, 4)
                        ElseIf (bHaveUnmaskColorSrc) Then
                           Call CopyMemory(atBitsQSrc(x, y), lciUnmaskColorSrc, 4)
                        End If
                     End Select
                  End If
               Next x
            Next y
         End If
      End If
   End If
   FreeImage_CreateMask = hDIBResult
End Function
Public Function FreeImage_CreateMaskImage(ByVal hDib As LongPtr, Optional ByVal lBitDepth As Long = 1, Optional ByVal eMaskOptions As FREE_IMAGE_MASK_FLAGS = FIMF_MASK_FULL_TRANSPARENCY, Optional ByVal vntMaskColors As Variant, Optional ByVal eMaskColorsFormat As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB, Optional ByVal lColorTolerance As Long, Optional ByVal lciMaskColor As Long = vbWhite, Optional ByVal eMaskColorFormat As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB, Optional ByVal lciUnmaskColor As Long = vbBlack, Optional ByVal eUnmaskColorFormat As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB) As LongPtr: FreeImage_CreateMaskImage = FreeImage_CreateMask(hDib, MCOF_CREATE_MASK_IMAGE, lBitDepth, eMaskOptions, vntMaskColors, eMaskColorsFormat, lColorTolerance, lciMaskColor, eMaskColorFormat, lciUnmaskColor, eUnmaskColorFormat): End Function
Public Function FreeImage_CreateSimpleBWMaskImage(ByVal hDib As LongPtr, Optional ByVal lBitDepth As Long = 1, Optional ByVal eMaskOptions As FREE_IMAGE_MASK_FLAGS = FIMF_MASK_FULL_TRANSPARENCY, Optional ByVal vntMaskColors As Variant, Optional ByVal eMaskColorsFormat As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB, Optional ByVal lColorTolerance As Long) As LongPtr: FreeImage_CreateSimpleBWMaskImage = FreeImage_CreateMask(hDib, MCOF_CREATE_MASK_IMAGE, lBitDepth, eMaskOptions, vntMaskColors, eMaskColorsFormat, lColorTolerance, vbWhite, FICFF_COLOR_RGB, vbBlack, FICFF_COLOR_RGB): End Function
Public Function FreeImage_CreateMaskInPlace(ByVal hDib As LongPtr, Optional ByVal lBitDepth As Long = 1, Optional ByVal eMaskOptions As FREE_IMAGE_MASK_FLAGS = FIMF_MASK_FULL_TRANSPARENCY, Optional ByVal vntMaskColors As Variant, Optional ByVal eMaskColorsFormat As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB, Optional ByVal lColorTolerance As Long, Optional ByVal vlciMaskColor As Variant, Optional ByVal eMaskColorFormat As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB, Optional ByVal vlciUnmaskColor As Variant, Optional ByVal eUnmaskColorFormat As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB) As LongPtr: FreeImage_CreateMaskInPlace = FreeImage_CreateMask(hDib, MCOF_MODIFY_SOURCE_IMAGE, lBitDepth, eMaskOptions, vntMaskColors, eMaskColorsFormat, lColorTolerance, , , , , vlciMaskColor, eMaskColorFormat, vlciUnmaskColor, eUnmaskColorFormat): End Function
Public Function FreeImage_CreateSimpleBWMaskInPlace(ByVal hDib As LongPtr, Optional ByVal lBitDepth As Long = 1, Optional ByVal eMaskOptions As FREE_IMAGE_MASK_FLAGS = FIMF_MASK_FULL_TRANSPARENCY, Optional ByVal vntMaskColors As Variant, Optional ByVal eMaskColorsFormat As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB, Optional ByVal lColorTolerance As Long) As LongPtr: FreeImage_CreateSimpleBWMaskInPlace = FreeImage_CreateMask(hDib, MCOF_MODIFY_SOURCE_IMAGE, lBitDepth, eMaskOptions, vntMaskColors, eMaskColorsFormat, lColorTolerance, , , , , vbWhite, FICFF_COLOR_RGB, vbBlack, FICFF_COLOR_RGB): End Function
Public Function FreeImage_CreateMaskColors(ParamArray MaskColors() As Variant) As Variant
' this is just a FreeImage signed function that emulates VB's
' builtin Array() function, that makes a variant array from
' a ParamArray; so, a caller of the FreeImage_CreateMask() function
' can specify all mask colors inline in the call statement
' hDibMask = FreeImage_CreateMask(hDib, 1, FIMF_MASK_COLOR_TRANSPARENCY, FreeImage_CreateMaskColors(vbRed, vbGreen, vbBlack),  FICFF_COLOR_BGR, .... )
' keep in mind, that VB colors (vbRed, vbBlue, etc.) are OLE colors that have BRG format
   FreeImage_CreateMaskColors = MaskColors
End Function
Public Function FreeImage_SwapColorLong(ByVal Color As Long, Optional ByVal IgnoreAlpha As Boolean) As Long
' swaps both color components Red (R) and Blue (B) in either
' and RGB or BGR format color value stored in a Long value. This function is
' used to convert from a RGB to a BGR color value and vice versa.
   If (Not IgnoreAlpha) Then
      FreeImage_SwapColorLong = ((Color And &HFF000000) Or ((Color And &HFF&) * &H10000) Or (Color And &HFF00&) Or ((Color And &HFF0000) \ &H10000))
   Else
      FreeImage_SwapColorLong = (((Color And &HFF&) * &H10000) Or (Color And &HFF00&) Or ((Color And &HFF0000) \ &H10000))
   End If
End Function
Public Function FreeImage_CompareColorsLongLong(ByVal ColorA As Long, ByVal ColorB As Long, Optional ByVal Tolerance As Long, Optional ByVal ColorTypeA As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_ARGB, Optional ByVal ColorTypeB As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_ARGB) As Long
' compares two colors that both are specified as a 32 bit Long value.
' Use both parameters 'ColorTypeA' and 'ColorTypeB' to specify each color's
' format and 'Tolerance' to specify the matching tolerance.
' The function returns the result of the mathematical substraction
' ColorA - ColorB, so if both colors are equal, the function returns NULL (0)
' and any other value if both colors are different. Alpha transparency is taken into
' account only if both colors are said to have an alpha transparency component by
' both parameters 'ColorTypeA' and 'ColorTypeB'. If at least one of both colors
' has no alpha transparency component, the comparison only includes the bits for
' the red, green and blue component.
' The matching tolerance is applied to each color component (red, green, blue and
' alpha) separately. So, when 'Tolerance' contains a value greater than zero, the
' function returns NULL (0) when either both colors are exactly the same or the
' differences of each corresponding color components are smaller or equal than
' the given tolerance value.
Dim bFormatEqual As Boolean
Dim bAlphaEqual As Boolean
   
   If (((ColorTypeA And FICFF_COLOR_PALETTE_INDEX) Or (ColorTypeB And FICFF_COLOR_PALETTE_INDEX)) = 0) Then
      bFormatEqual = ((ColorTypeA And FICFF_COLOR_FORMAT_ORDER_MASK) = (ColorTypeB And FICFF_COLOR_FORMAT_ORDER_MASK))
      bAlphaEqual = ((ColorTypeA And FICFF_COLOR_HAS_ALPHA) And (ColorTypeB And FICFF_COLOR_HAS_ALPHA))
      If (bFormatEqual) Then
         If (bAlphaEqual) Then
            FreeImage_CompareColorsLongLong = ColorA - ColorB
         Else
            FreeImage_CompareColorsLongLong = (ColorA And &HFFFFFF) - (ColorB And &HFFFFFF)
         End If
      Else
         If (bAlphaEqual) Then
            FreeImage_CompareColorsLongLong = ColorA - ((ColorB And &HFF000000) Or ((ColorB And &HFF&) * &H10000) Or (ColorB And &HFF00&) Or ((ColorB And &HFF0000) \ &H10000))
         Else
            FreeImage_CompareColorsLongLong = (ColorA And &HFFFFFF) - (((ColorB And &HFF&) * &H10000) Or (ColorB And &HFF00&) Or ((ColorB And &HFF0000) \ &H10000))
         End If
      End If
      If ((Tolerance > 0) And (FreeImage_CompareColorsLongLong <> 0)) Then
         If (bFormatEqual) Then
            If (Abs(((ColorA \ &H10000) And &HFF) - ((ColorB \ &H10000) And &HFF)) <= Tolerance) Then
               If (Abs(((ColorA \ &H100) And &HFF) - ((ColorB \ &H100) And &HFF)) <= Tolerance) Then
                  If (Abs((ColorA And &HFF) - (ColorB And &HFF)) <= Tolerance) Then
                     If (bAlphaEqual) Then
                        If (Abs(((ColorA \ &H1000000) And &HFF) - ((ColorB \ &H1000000) And &HFF)) <= Tolerance) Then
                           FreeImage_CompareColorsLongLong = 0
                        End If
                     Else
                        FreeImage_CompareColorsLongLong = 0
                     End If
                  End If
               End If
            End If
         Else
            If (Abs(((ColorA \ &H10000) And &HFF) - (ColorB And &HFF)) <= Tolerance) Then
               If (Abs(((ColorA \ &H100) And &HFF) - ((ColorB \ &H100) And &HFF)) <= Tolerance) Then
                  If (Abs((ColorA And &HFF) - ((ColorB \ &H10000) And &HFF)) <= Tolerance) Then
                     If (bAlphaEqual) Then
                        If (Abs(((ColorA \ &H1000000) And &HFF) - ((ColorB \ &H1000000) And &HFF)) <= Tolerance) Then
                           FreeImage_CompareColorsLongLong = 0
                        End If
                     Else
                        FreeImage_CompareColorsLongLong = 0
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End Function
Public Function FreeImage_CompareColorsRGBTRIPLELong(ByRef ColorA As RGBTRIPLE, ByVal ColorB As Long, Optional ByVal Tolerance As Long, Optional ByVal ColorTypeB As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB) As Long
' This is a function derived from 'FreeImage_CompareColorsLongLong()' to make color
' comparisons between two colors whereby one color is provided as RGBTRIPLE and the
' other color is provided as Long value.
' Have a look at the documentation of 'FreeImage_CompareColorsLongLong()' to learn
' more about color comparisons.
Dim lcColorA As Long
   Call CopyMemory(lcColorA, ColorA, 3)
   FreeImage_CompareColorsRGBTRIPLELong = FreeImage_CompareColorsLongLong(lcColorA, ColorB, Tolerance, FICFF_COLOR_RGB, ColorTypeB)
End Function
Public Function FreeImage_CompareColorsRGBQUADLong(ByRef ColorA As RGBQUAD, ByVal ColorB As Long, Optional ByVal Tolerance As Long, Optional ByVal ColorTypeB As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_ARGB) As Long
' This is a function derived from 'FreeImage_CompareColorsLongLong()' to make color
' comparisons between two colors whereby one color is provided as RGBQUAD and the
' other color is provided as Long value.
' Have a look at the documentation of 'FreeImage_CompareColorsLongLong()' to learn
' more about color comparisons.
Dim lcColorA As Long
   Call CopyMemory(lcColorA, ColorA, 4)
   FreeImage_CompareColorsRGBQUADLong = FreeImage_CompareColorsLongLong(lcColorA, ColorB, Tolerance, FICFF_COLOR_ARGB, ColorTypeB)
End Function
Public Function FreeImage_SearchPalette(ByVal BITMAP As LongPtr, ByVal Color As Long, Optional ByVal Tolerance As Long, Optional ByVal ColorType As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB, Optional ByVal TransparencyState As FREE_IMAGE_TRANSPARENCY_STATE_FLAGS = FITSF_IGNORE_TRANSPARENCY) As Long
' searches an image's color palette for a certain color specified as a 32 bit Long value in either RGB or BGR format.
' A search tolerance may be specified in the 'Tolerance' parameter.
' If no transparency table was found for the specified image, transparency information will be ignored during the search.
' Then, the function behaves as if FITSF_IGNORE_TRANSPARENCY was specified for parameter TransparencyState.
' Use the 'TransparencyState' parameter to control, how the transparency state of the found palette entry affects the result.
' These values may be used:
' FITSF_IGNORE_TRANSPARENCY:        Returns the index of the first palette entry which
'                                   matches the red, green and blue components.
' FITSF_NONTRANSPARENT:             Returns the index of the first palette entry which
'                                   matches the red, green and blue components and is
'                                   nontransparent (fully opaque).
' FITSF_TRANSPARENT:                Returns the index of the first palette entry which
'                                   matches the red, green and blue components and is
'                                   fully transparent.
' FITSF_INCLUDE_ALPHA_TRANSPARENCY: Returns the index of the first palette entry which
'                                   matches the red, green and blue components as well
'                                   as the alpha transparency.
' When alpha transparency should be included in the palette search ('FITSF_INCLUDE_ALPHA_TRANSPARENCY'),
' the alpha transparency of the color searched is taken from the left most byte of 'Color'
' (Color is either in format ARGB or ABGR). The the alpha transparency of the palette entry
' actually comes from the image's transparency table rather than from the palette, since palettes
' do not contain transparency information.
Dim abTransparencyTable() As Byte
Dim alPalette() As Long
Dim i As Long
   If (FreeImage_GetImageType(BITMAP) = FIT_BITMAP) Then
      Select Case FreeImage_GetColorType(BITMAP)
      Case FIC_PALETTE, FIC_MINISBLACK, FIC_MINISWHITE
         FreeImage_SearchPalette = -1
         alPalette = FreeImage_GetPaletteExLong(BITMAP)
         If (FreeImage_GetTransparencyCount(BITMAP) > UBound(alPalette)) Then
            abTransparencyTable = FreeImage_GetTransparencyTableExClone(BITMAP)
         Else
            TransparencyState = FITSF_IGNORE_TRANSPARENCY
         End If
         For i = 0 To UBound(alPalette)
            If (FreeImage_CompareColorsLongLong(Color, alPalette(i), Tolerance, ColorType, FICFF_COLOR_RGB) = 0) Then
               Select Case TransparencyState
               Case FITSF_IGNORE_TRANSPARENCY
                  FreeImage_SearchPalette = i
                  Exit For
               Case FITSF_NONTRANSPARENT
                  If (abTransparencyTable(i) = 255) Then
                     FreeImage_SearchPalette = i
                     Exit For
                  End If
               Case FITSF_TRANSPARENT
                  If (abTransparencyTable(i) = 0) Then
                     FreeImage_SearchPalette = i
                     Exit For
                  End If
               Case FITSF_INCLUDE_ALPHA_TRANSPARENCY
                  If (abTransparencyTable(i) = ((Color And &HFF000000) \ 1000000)) Then
                     FreeImage_SearchPalette = i
                     Exit For
                  End If
               End Select
            End If
         Next i
      Case Else
         FreeImage_SearchPalette = -1
      End Select
   Else
      FreeImage_SearchPalette = -1
   End If
End Function
Public Function FreeImage_GetIcon(ByVal hDib As LongPtr, Optional ByVal eTransparencyOptions As FREE_IMAGE_ICON_TRANSPARENCY_OPTION_FLAGS = ITOF_USE_DEFAULT_TRANSPARENCY, Optional ByVal lciTransparentColor As Long, Optional ByVal eTransparentColorType As FREE_IMAGE_COLOR_FORMAT_FLAGS = FICFF_COLOR_RGB, Optional ByVal hdc As LongPtr, Optional ByVal UnloadSource As Boolean) As Long
' The optional 'UnloadSource' parameter is for unloading the original image after the OlePicture has been created, so you can easiely "switch" from a
' FreeImage DIB to a VB Picture object. There is no need to clean up the DIB at the caller's site.
    If (hDib) = 0 Then Exit Function
Dim hDIBsrc As LongPtr
Dim hDIBMask As LongPtr
Dim hBmpMask As LongPtr
Dim hBmp As LongPtr
Dim tIconInfo As ICONINFO
Dim bReleaseDC As Boolean
Dim bModifySourceImage As Boolean
Dim eMaskFlags As FREE_IMAGE_MASK_FLAGS
Dim lBitDepth As Long
Dim bPixelIndex As Byte
    If (Not FreeImage_HasPixels(hDib)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to create an icon from a 'header-only' bitmap.")
    lBitDepth = FreeImage_GetBPP(hDib)
    ' check whether the image supports transparency
    Select Case lBitDepth
    Case 4, 8
       If (eTransparencyOptions And ITOF_USE_TRANSPARENCY_INFO) Then
          If (FreeImage_IsTransparent(hDib)) Then
             eMaskFlags = FIMF_MASK_FULL_TRANSPARENCY
          ElseIf (eTransparencyOptions And ITOF_FORCE_TRANSPARENCY_INFO) Then
             If (FreeImage_IsTransparencyTableTransparent(hDib)) Then
                eMaskFlags = (FIMF_MASK_FULL_TRANSPARENCY And FIMF_MASK_FORCE_TRANSPARENCY)
             End If
          End If
       End If
       If ((eMaskFlags = FIMF_MASK_NONE) And (eTransparencyOptions And ITOF_USE_COLOR_TRANSPARENCY)) Then
          eMaskFlags = FIMF_MASK_COLOR_TRANSPARENCY
          Select Case (eTransparencyOptions And ITOF_USE_COLOR_BITMASK)
          Case ITOF_USE_COLOR_TOP_LEFT_PIXEL
             Call FreeImage_GetPixelIndex(hDib, 0, FreeImage_GetHeight(hDib) - 1, bPixelIndex)
             lciTransparentColor = bPixelIndex
             eTransparentColorType = FICFF_COLOR_PALETTE_INDEX
          Case ITOF_USE_COLOR_TOP_RIGHT_PIXEL
             Call FreeImage_GetPixelIndex(hDib, FreeImage_GetWidth(hDib) - 1, FreeImage_GetHeight(hDib) - 1, bPixelIndex)
             lciTransparentColor = bPixelIndex
             eTransparentColorType = FICFF_COLOR_PALETTE_INDEX
          Case ITOF_USE_COLOR_BOTTOM_LEFT_PIXEL
             Call FreeImage_GetPixelIndex(hDib, 0, 0, bPixelIndex)
             lciTransparentColor = bPixelIndex
             eTransparentColorType = FICFF_COLOR_PALETTE_INDEX
          Case ITOF_USE_COLOR_BOTTOM_RIGHT_PIXEL
             Call FreeImage_GetPixelIndex(hDib, FreeImage_GetWidth(hDib) - 1, 0, bPixelIndex)
             lciTransparentColor = bPixelIndex
             eTransparentColorType = FICFF_COLOR_PALETTE_INDEX
          End Select
       End If
       bModifySourceImage = True
    Case 24, 32
       If ((lBitDepth = 32) And (eTransparencyOptions And ITOF_USE_TRANSPARENCY_INFO)) Then
          If (FreeImage_IsTransparent(hDib)) Then eMaskFlags = FIMF_MASK_FULL_TRANSPARENCY
       End If
       If ((eMaskFlags = FIMF_MASK_NONE) And (eTransparencyOptions And ITOF_USE_COLOR_TRANSPARENCY)) Then
          eMaskFlags = FIMF_MASK_COLOR_TRANSPARENCY
          Select Case (eTransparencyOptions And ITOF_USE_COLOR_BITMASK)
          Case ITOF_USE_COLOR_TOP_LEFT_PIXEL
             Call FreeImage_GetPixelColorByLong(hDib, FreeImage_GetHeight(hDib) - 1, 0, lciTransparentColor)
             eTransparentColorType = FICFF_COLOR_RGB
          Case ITOF_USE_COLOR_TOP_RIGHT_PIXEL
             Call FreeImage_GetPixelColorByLong(hDib, FreeImage_GetHeight(hDib) - 1, FreeImage_GetWidth(hDib) - 1, lciTransparentColor)
             eTransparentColorType = FICFF_COLOR_RGB
          Case ITOF_USE_COLOR_BOTTOM_LEFT_PIXEL
             Call FreeImage_GetPixelColorByLong(hDib, 0, 0, lciTransparentColor)
             eTransparentColorType = FICFF_COLOR_RGB
          Case ITOF_USE_COLOR_BOTTOM_RIGHT_PIXEL
             Call FreeImage_GetPixelColorByLong(hDib, 0, FreeImage_GetWidth(hDib) - 1, lciTransparentColor)
             eTransparentColorType = FICFF_COLOR_RGB
          End Select
       End If
       bModifySourceImage = (lBitDepth = 24)
    End Select

    If (bModifySourceImage) Then
       hDIBsrc = FreeImage_Clone(hDib)
       hDIBMask = FreeImage_CreateMask(hDIBsrc, MCOF_CREATE_AND_MODIFY, 1, eMaskFlags, lciTransparentColor, eTransparentColorType, _
           vlciMaskColorSrc:=FreeImage_SearchPalette(hDIBsrc, 0, TransparencyState:=FITSF_NONTRANSPARENT), eUnmaskColorSrcFormat:=FICFF_COLOR_PALETTE_INDEX)
    Else
       hDIBsrc = hDib
       hDIBMask = FreeImage_CreateMaskImage(hDib, 1, FIMF_MASK_FULL_TRANSPARENCY)
    End If

    If (hdc = 0) Then hdc = GetDC(0): bReleaseDC = True
    
    hBmp = CreateDIBitmap(hdc, FreeImage_GetInfoHeader(hDIBsrc), CBM_INIT, FreeImage_GetBits(hDIBsrc), FreeImage_GetInfo(hDIBsrc), DIB_RGB_COLORS)
    hBmpMask = CreateDIBitmap(hdc, FreeImage_GetInfoHeader(hDIBMask), CBM_INIT, FreeImage_GetBits(hDIBMask), FreeImage_GetInfo(hDIBMask), DIB_RGB_COLORS)

    If (bModifySourceImage) Then Call FreeImage_Unload(hDIBsrc)
    If (UnloadSource) Then Call FreeImage_Unload(hDib)
    
    If ((hBmp <> 0) And (hBmpMask <> 0)) Then
       With tIconInfo
          .fIcon = 1 'True
          .hbmMask = hBmpMask
          .hbmColor = hBmp
       End With
       FreeImage_GetIcon = CreateIconIndirect(tIconInfo)
    End If
    If (bReleaseDC) Then Call ReleaseDC(0, hdc)
End Function
Public Function FreeImage_AdjustPictureBox(ByRef Control As Object, Optional ByVal Mode As FREE_IMAGE_ADJUST_MODE = AM_DEFAULT, Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC) As IPicture
' adjusts an already loaded picture in a VB PictureBox
' control in size. This is done by converting the picture to a Bitmap
' by FreeImage_CreateFromOlePicture. After resizing the Bitmap it is
' converted back to a Ole Picture object and re-assigned to the PictureBox control.
' The Control paramater is actually of type Object so any object or control
' providing Picture, hWnd, Width and Height properties can be used instead of a PictureBox control
' This may be useful when using compile time provided images in VB like
' logos or backgrounds that need to be resized during runtime. Using
' FreeImage's sophisticated rescaling methods is a much better aproach
' than using VB's stretchable Image control.
' One reason for resizing a usually fixed size logo or background image
' may be the following scenario:
' When running on a Windows machine using smaller or bigger fonts (what can
' be configured in the control panel by using different dpi fonts), the
' operation system automatically adjusts the sizes of Forms, Labels,
' TextBoxes, Frames and even PictureBoxes. So, the hole VB application is
' perfectly adapted to these font metrics with the exception of compile time
' provided images. Although the PictureBox control is resized, the containing
' image remains untouched. This problem could be solved with this function.
' is also wrapped by the function 'AdjustPicture', giving you
' a more VB common function name.
Const vbObjectOrWithBlockVariableNotSet As Long = 91
    If (Control Is Nothing) Then Err.Raise (vbObjectOrWithBlockVariableNotSet): Exit Function
Dim tR As RECT: Call GetClientRect(Control.hwnd, tR): If ((tR.Right = Control.Picture.Width) And (tR.Bottom = Control.Picture.Height)) Then Exit Function
Dim hDIBdst As LongPtr: hDIBdst = FreeImage_CreateFromOlePicture(Control.Picture)
    If (hDIBdst = 0) Then Exit Function
    If (Mode = AM_ADJUST_OPTIMAL_SIZE) Then
       If (Control.Picture.Width >= Control.Picture.Height) Then
          Mode = AM_ADJUST_WIDTH
       Else
          Mode = AM_ADJUST_HEIGHT
       End If
    End If
Dim lNewHeight As Long, lNewWidth As Long
    Select Case Mode
    Case AM_STRECH:         lNewWidth = tR.Right: lNewHeight = tR.Bottom
    Case AM_ADJUST_WIDTH:   lNewWidth = tR.Right: lNewHeight = lNewWidth / (Control.Picture.Width / Control.Picture.Height)
    Case AM_ADJUST_HEIGHT:  lNewHeight = tR.Bottom: lNewWidth = lNewHeight * (Control.Picture.Width / Control.Picture.Height)
    End Select
Dim hDIBsrc As LongPtr: hDIBsrc = hDIBdst
    hDIBdst = FreeImage_Rescale(hDIBdst, lNewWidth, lNewHeight, Filter)
    Call FreeImage_Unload(hDIBsrc)
    Set Control.Picture = FreeImage_GetOlePicture(hDIBdst, , True)
    Set FreeImage_AdjustPictureBox = Control.Picture
End Function
Public Function AdjustPicture(ByRef Control As Object, Optional ByRef Mode As FREE_IMAGE_ADJUST_MODE = AM_DEFAULT, Optional ByRef Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC) As IPicture
   Set AdjustPicture = FreeImage_AdjustPictureBox(Control, Mode, Filter)
End Function
Public Function FreeImage_LoadEx(ByVal FileName As String, Optional ByVal Options As FREE_IMAGE_LOAD_OPTIONS, Optional ByVal Width As Variant, Optional ByVal Height As Variant, Optional ByVal InPercent As Boolean, Optional ByVal Filter As FREE_IMAGE_FILTER, Optional ByRef Format As FREE_IMAGE_FORMAT) As LongPtr
' The function provides all image formats, the FreeImage library can read.
' The image format is determined from the image file to load, the optional parameter
' 'Format' is an OUT parameter that will contain the image format that has been loaded.
' The parameters 'Width', 'Height', 'InPercent' and 'Filter' make it possible
' to "load" the image in a resized version. 'Width', 'Height' specify the desired
' width and height, 'Filter' determines, what image filter should be used on the resizing process.
' The parameters 'Width', 'Height', 'InPercent' and 'Filter' map directly to the
' according parameters of the 'FreeImage_RescaleEx' function. So, read the
' documentation of the 'FreeImage_RescaleEx' for a complete understanding of the
' usage of these parameters.
Const vbInvalidPictureError As Long = 481
   Format = FreeImage_GetFileType(FileName)
   If (Format <> FIF_UNKNOWN) Then
      If (p_FreeImage_FIFSupportsReading(Format) = 1) Then
         FreeImage_LoadEx = FreeImage_Load(Format, FileName, Options)
         If (FreeImage_LoadEx) Then
            If ((Not IsMissing(Width)) Or (Not IsMissing(Height))) Then
               FreeImage_LoadEx = FreeImage_RescaleEx(FreeImage_LoadEx, Width, Height, InPercent, True, Filter)
            End If
         Else
            Call Err.Raise(vbInvalidPictureError)
         End If
      Else
         Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "FreeImage Library plugin '" & FreeImage_GetFormatFromFIF(Format) & "' " & "does not support reading.")
      End If
   Else
      Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "The file specified has an unknown image format.")
   End If
End Function
Public Function LoadPictureEx(Optional ByRef FileName As Variant, Optional ByRef Options As FREE_IMAGE_LOAD_OPTIONS, Optional ByRef Width As Variant, Optional ByRef Height As Variant, Optional ByRef InPercent As Boolean, Optional ByRef Filter As FREE_IMAGE_FILTER, Optional ByRef Format As FREE_IMAGE_FORMAT) As IPicture
' is an extended version of the VB method 'LoadPicture'.
' As the VB version it takes a filename parameter to load the image and throws the same errors in most cases.
' now is only a thin wrapper for the FreeImage_LoadEx() wrapper function (as compared to releases of this wrapper prior to version 1.8).
' So, have a look at this function's discussion of the parameters.
' However, we do mask out the FILO_LOAD_NOPIXELS load option, since this function shall create a VB Picture object, which does not support FreeImage's header-only loading option.
    If (IsMissing(FileName)) Then Exit Function
Dim hDib As LongPtr: hDib = FreeImage_LoadEx(FileName, (Options And (Not FILO_LOAD_NOPIXELS)), Width, Height, InPercent, Filter, Format)
    Set LoadPictureEx = FreeImage_GetOlePicture(hDib, , True)
End Function
Public Function FreeImage_SaveEx(ByVal BITMAP As LongPtr, ByVal FileName As String, Optional ByVal Format As FREE_IMAGE_FORMAT = FIF_UNKNOWN, Optional ByVal Options As FREE_IMAGE_SAVE_OPTIONS, Optional ByVal ColorDepth As FREE_IMAGE_COLOR_DEPTH, Optional ByVal Width As Variant, Optional ByVal Height As Variant, Optional ByVal InPercent As Boolean, Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC, Optional ByVal UnloadSource As Boolean) As Boolean
' is an easy to use replacement for FreeImage's FreeImage_Save() function which supports inline size- and color conversions as well as an
' auto image format detection algorithm that determines the desired image format by the given filename. An even more sophisticated algorithm may auto-detect the proper color depth for a explicitly given or auto-detected image format.
' The function provides all image formats, and save options, the FreeImagelibrary can write.
' The optional parameter 'Format' may contain the desired image format. When omitted, the function tries to get the image format from the filename extension.
' The optional parameter 'ColorDepth' may contain the desired color depth for the saved image.
' This can be either any value of the FREE_IMAGE_COLOR_DEPTH enumeration or the value FICD_AUTO what is the default value of the parameter.
' When 'ColorDepth' is FICD_AUTO, the function tries to get the most suitable color depth for the specified image format if the image's current color depth is not supported by the specified image format.
' Therefore, the function firstly reduces the color depth step by step until a proper color depth is found since an incremention would only increase the file's size with no quality benefit.
' Only when there is no lower color depth is found for the image format, the function starts to increase the color depth.
' Keep in mind that an explicitly specified color depth that is not supported by the image format results in a runtime error.
' For example, when saving a 24 bit image as GIF image, a runtime error occurs.
' The function checks, whether the given filename has a valid extension or not.
' If not, the "primary" extension for the used image format will be appended to the filename.
' The parameter 'Filename' remains untouched in this case.
' To learn more about the "primary" extension, read the documentation for the 'FreeImage_GetPrimaryExtensionFromFIF' function.
' The parameters 'Width', 'Height', 'InPercent' and 'Filter' make it possible to save the image in a resized version.
' 'Width', 'Height' specify the desired width and height, 'Filter' determines, what image filter should be used
' on the resizing process. Since FreeImage_SaveEx relies on FreeImage_RescaleEx,
' please refer to the documentation of FreeImage_RescaleEx to learn more about these four parameters.
' The optional 'UnloadSource' parameter is for unloading the saved image, so you can save and unload an image with this function in one operation.
' !!! CAUTION !!!: at current, the image is unloaded, even if the image was not saved correctly!
    If (BITMAP = 0) Then Exit Function
Dim hDIBRescale As LongPtr
Dim bConvertedOnRescale As Boolean
Dim bIsNewDIB As Boolean
Dim lBPP As Long
Dim lBPPOrg As Long
Dim strExtension As String
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to save 'header-only' bitmaps.")
    If ((Not IsMissing(Width)) Or (Not IsMissing(Height))) Then
       lBPP = FreeImage_GetBPP(BITMAP)
       hDIBRescale = FreeImage_RescaleEx(BITMAP, Width, Height, InPercent, UnloadSource, Filter)
       bIsNewDIB = (hDIBRescale <> BITMAP)
       BITMAP = hDIBRescale
       bConvertedOnRescale = (lBPP <> FreeImage_GetBPP(BITMAP))
    End If
    If (Format = FIF_UNKNOWN) Then Format = FreeImage_GetFIFFromFilename(FileName)
    If (Format = FIF_UNKNOWN) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unknown image format. Neither an explicit image format " & "was specified nor any known image format was determined " & "from the filename specified.")
    If ((p_FreeImage_FIFSupportsWriting(Format) <> 1) Or (p_FreeImage_FIFSupportsExportType(Format, FIT_BITMAP) <> 1)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "FreeImage Library plugin '" & FreeImage_GetFormatFromFIF(Format) & "' " & "is unable to write images of the image format requested.")

    If (Not FreeImage_IsFilenameValidForFIF(Format, FileName)) Then strExtension = "." & FreeImage_GetPrimaryExtensionFromFIF(Format)
    ' check color depth
    If (ColorDepth <> FICD_AUTO) Then
       ' mask out bit 1 (0x02) for the case ColorDepth is FICD_MONOCHROME_DITHER (0x03)
       ' FREE_IMAGE_COLOR_DEPTH values are true bit depths in general expect FICD_MONOCHROME_DITHER
       ' by masking out bit 1, 'FreeImage_FIFSupportsExportBPP()' tests for bitdepth 1
       ' what is correct again for dithered images.
       ColorDepth = (ColorDepth And (Not &H2))
       If (p_FreeImage_FIFSupportsExportBPP(Format, ColorDepth) <> 1) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "FreeImage Library plugin '" & FreeImage_GetFormatFromFIF(Format) & "' " & "is unable to write images with a color depth " & "of " & ColorDepth & " bpp.")
       If (FreeImage_GetBPP(BITMAP) <> ColorDepth) Then
          BITMAP = FreeImage_ConvertColorDepth(BITMAP, ColorDepth, (UnloadSource Or bIsNewDIB))
          bIsNewDIB = True
       End If
    Else
       If (lBPP = 0) Then lBPP = FreeImage_GetBPP(BITMAP)
       If (p_FreeImage_FIFSupportsExportBPP(Format, lBPP) <> 1) Then
          lBPPOrg = lBPP
          Do
             lBPP = p_GetPreviousColorDepth(lBPP)
          Loop While ((p_FreeImage_FIFSupportsExportBPP(Format, lBPP) <> 1) Or (lBPP = 0))
          If (lBPP = 0) Then
             lBPP = lBPPOrg
             Do
                lBPP = p_GetNextColorDepth(lBPP)
             Loop While ((p_FreeImage_FIFSupportsExportBPP(Format, lBPP) <> 1) Or (lBPP = 0))
          End If
          If (lBPP <> 0) Then
             BITMAP = FreeImage_ConvertColorDepth(BITMAP, lBPP, (UnloadSource Or bIsNewDIB))
             bIsNewDIB = True
          End If
       ElseIf (bConvertedOnRescale) Then
          ' restore original color depth
          ' always unload current DIB here, since 'bIsNewDIB' is True
          BITMAP = FreeImage_ConvertColorDepth(BITMAP, lBPP, True)
       End If
    End If
    FreeImage_SaveEx = FreeImage_Save(Format, BITMAP, FileName & strExtension, Options)
    If ((bIsNewDIB) Or (UnloadSource)) Then Call FreeImage_Unload(BITMAP)
End Function
Public Function SavePictureEx(ByRef Picture As IPicture, ByRef FileName As String, Optional ByRef Format As FREE_IMAGE_FORMAT, Optional ByRef Options As FREE_IMAGE_SAVE_OPTIONS, Optional ByRef ColorDepth As FREE_IMAGE_COLOR_DEPTH, Optional ByRef Width As Variant, Optional ByRef Height As Variant, Optional ByRef InPercent As Boolean, Optional ByRef Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC) As Boolean
' is an extended version of the VB method 'SavePicture'.
' As the VB version it takes a Picture object and a filename parameter to save the image and throws the same errors in most cases.
' now is only a thin wrapper for the FreeImage_SaveEx() wrapperfunction (as compared to releases of this wrapper prior to version 1.8).
Const vbObjectOrWithBlockVariableNotSet As Long = 91
Const vbInvalidPictureError As Long = 481
    If (Picture Is Nothing) Then Call Err.Raise(vbObjectOrWithBlockVariableNotSet): Exit Function
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Call Err.Raise(vbInvalidPictureError): Exit Function
    SavePictureEx = FreeImage_SaveEx(hDIBsrc, FileName, Format, Options, ColorDepth, Width, Height, InPercent, FILTER_BICUBIC, True)
End Function
Public Function SaveImageContainerEx(ByRef Container As Object, ByRef FileName As String, Optional ByVal IncludeDrawings As Boolean, Optional ByRef Format As FREE_IMAGE_FORMAT, Optional ByRef Options As FREE_IMAGE_SAVE_OPTIONS, Optional ByRef ColorDepth As FREE_IMAGE_COLOR_DEPTH, Optional ByRef Width As Variant, Optional ByRef Height As Variant, Optional ByRef InPercent As Boolean, Optional ByRef Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC) As Long
' is an extended version of the VB method 'SavePicture'.
' As the VB version it takes an image hosting control and a filename parameter to save the image and throws the same errors in most cases.
' merges the functionality of both wrapper functions 'SavePictureEx()' and 'FreeImage_CreateFromImageContainer()'.
' Basically this function is identical to 'SavePictureEx' expect that is does not take a IOlePicture (IPicture) object but a VB image hosting container control.
    Call SavePictureEx(p_GetIOlePictureFromContainer(Container, IncludeDrawings), FileName, Format, Options, ColorDepth, Width, Height, InPercent, Filter)
End Function
Public Function FreeImage_OpenMultiBitmapEx(ByVal FileName As String, Optional ByVal ReadOnly As Boolean, Optional ByVal KeepCacheInMemory As Boolean, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS, Optional ByRef Format As FREE_IMAGE_FORMAT) As LongPtr
    Format = FreeImage_GetFileType(FileName)
    If (Format = FIF_UNKNOWN) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "The file specified has an unknown image format."): Exit Function
    Select Case Format
    Case FIF_TIFF, FIF_GIF, FIF_ICO: FreeImage_OpenMultiBitmapEx = FreeImage_OpenMultiBitmap(Format, FileName, False, ReadOnly, KeepCacheInMemory, Flags)
    Case Else:                       Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "FreeImage Library plugin '" & FreeImage_GetFormatFromFIF(Format) & "' " & "does not have any support for multi-page bitmaps.")
    End Select
End Function
Public Function FreeImage_CreateMultiBitmapEx(ByVal FileName As String, Optional ByVal KeepCacheInMemory As Boolean, Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS, Optional ByRef Format As FREE_IMAGE_FORMAT) As LongPtr
    If (Format = FIF_UNKNOWN) Then Format = FreeImage_GetFIFFromFilename(FileName)
    If (Format = FIF_UNKNOWN) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "The file specified has an unknown image format."): Exit Function
    Select Case Format
    Case FIF_TIFF, FIF_GIF, FIF_ICO: FreeImage_CreateMultiBitmapEx = FreeImage_OpenMultiBitmap(Format, FileName, True, False, KeepCacheInMemory, Flags)
    Case Else:                       Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "FreeImage Library plugin '" & FreeImage_GetFormatFromFIF(Format) & "' " & "does not have any support for multi-page bitmaps.")
    End Select
End Function
'----------------------
' OlePicture aware toolkit, rescale and conversion functions
'----------------------
Public Function FreeImage_RescaleIOP(ByRef Picture As IPicture, Optional ByVal Width As Variant, Optional ByVal Height As Variant, Optional ByVal IsPercentValue As Boolean, Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC, Optional ByVal ForceCloneCreation As Boolean) As IPicture
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
    hDIBsrc = FreeImage_RescaleEx(hDIBsrc, Width, Height, IsPercentValue, True, Filter, ForceCloneCreation)
    Set FreeImage_RescaleIOP = FreeImage_GetOlePicture(hDIBsrc, , True)
End Function
Public Function FreeImage_RescaleByPixelIOP(ByRef Picture As IPicture, Optional ByVal WidthInPixels As Long, Optional ByVal HeightInPixels As Long, Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC, Optional ByVal ForceCloneCreation As Boolean) As IPicture
    Set FreeImage_RescaleByPixelIOP = FreeImage_RescaleIOP(Picture, WidthInPixels, HeightInPixels, False, Filter, ForceCloneCreation)
End Function
Public Function FreeImage_RescaleByPercentIOP(ByRef Picture As IPicture, Optional ByVal WidthPercentage As Double, Optional ByVal HeightPercentage As Double, Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC, Optional ByVal ForceCloneCreation As Boolean) As IPicture
    Set FreeImage_RescaleByPercentIOP = FreeImage_RescaleIOP(Picture, WidthPercentage, HeightPercentage, True, Filter, ForceCloneCreation)
End Function
Public Function FreeImage_RescaleByFactorIOP(ByRef Picture As IPicture, Optional ByVal WidthFactor As Double, Optional ByVal HeightFactor As Double, Optional ByVal Filter As FREE_IMAGE_FILTER = FILTER_BICUBIC, Optional ByVal ForceCloneCreation As Boolean) As IPicture
   Set FreeImage_RescaleByFactorIOP = FreeImage_RescaleIOP(Picture, WidthFactor, HeightFactor, False, Filter, ForceCloneCreation)
End Function
Public Function FreeImage_MakeThumbnailIOP(ByRef Picture As IPicture, ByVal MaxPixelSize As Long, Optional ByVal Convert As Boolean) As IPicture
' IOlePicture based wrapper for wrapper function FreeImage_MakeThumbnail()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
Dim hDIBdst As LongPtr:         hDIBdst = FreeImage_MakeThumbnail(hDIBsrc, MaxPixelSize, Convert)
    If (hDIBdst) Then Set FreeImage_MakeThumbnailIOP = FreeImage_GetOlePicture(hDIBdst, , True)
    Call FreeImage_Unload(hDIBsrc)
End Function
Public Function FreeImage_ConvertColorDepthIOP(ByRef Picture As IPicture, ByVal Conversion As FREE_IMAGE_CONVERSION_FLAGS, Optional ByVal threshold As Byte = 128, Optional ByVal DitherMethod As FREE_IMAGE_DITHER = FID_FS, Optional ByVal QuantizeMethod As FREE_IMAGE_QUANTIZE = FIQ_WUQUANT) As IPicture
' IOlePicture based wrapper for wrapper function FreeImage_ConvertColorDepth()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
    hDIBsrc = FreeImage_ConvertColorDepth(hDIBsrc, Conversion, True, threshold, DitherMethod, QuantizeMethod)
    Set FreeImage_ConvertColorDepthIOP = FreeImage_GetOlePicture(hDIBsrc, , True)
End Function
Public Function FreeImage_ColorQuantizeExIOP(ByRef Picture As IPicture, Optional ByVal QuantizeMethod As FREE_IMAGE_QUANTIZE = FIQ_WUQUANT, Optional ByVal PaletteSize As Long = 256, Optional ByVal ReserveSize As Long, Optional ByRef ReservePalette As Variant = Null) As IPicture
' IOlePicture based wrapper for wrapper function FreeImage_ColorQuantizeEx()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
    hDIBsrc = FreeImage_ColorQuantizeEx(hDIBsrc, QuantizeMethod, True, PaletteSize, ReserveSize, ReservePalette)
    Set FreeImage_ColorQuantizeExIOP = FreeImage_GetOlePicture(hDIBsrc, , True)
End Function
Public Function FreeImage_RotateClassicIOP(ByRef Picture As IPicture, ByVal Angle As Double) As IPicture
' IOlePicture based wrapper for FreeImage function FreeImage_RotateClassic()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
Dim hDIBdst As LongPtr
    Select Case FreeImage_GetBPP(hDIBsrc)
    Case 1, 8, 24, 32:          hDIBdst = FreeImage_RotateClassic(hDIBsrc, Angle)
        Set FreeImage_RotateClassicIOP = FreeImage_GetOlePicture(hDIBdst, , True)
    End Select
    Call FreeImage_Unload(hDIBsrc)
End Function
Public Function FreeImage_RotateIOP(ByRef Picture As IPicture, ByVal Angle As Double, Optional ByVal ColorPtr As Long) As IPicture
' IOlePicture based wrapper for FreeImage function FreeImage_Rotate()
' The optional ColorPtr parameter takes a pointer to (e.g. the address of) an RGB color value.
' So, all these assignments are valid for ColorPtr:
'
' Dim tColor As RGBQUAD
'
' VarPtr(tColor)
' VarPtr(&H33FF80)
' VarPtr(vbWhite) ' However, the VB color constants are in BGR format!
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
Dim hDIBdst As LongPtr
    Select Case FreeImage_GetBPP(hDIBsrc)
    Case 1, 8, 24, 32:          hDIBdst = FreeImage_Rotate(hDIBsrc, Angle, ByVal ColorPtr)
        Set FreeImage_RotateIOP = FreeImage_GetOlePicture(hDIBdst, , True)
    End Select
    Call FreeImage_Unload(hDIBsrc)
End Function
Public Function FreeImage_RotateExIOP(ByRef Picture As IPicture, ByVal Angle As Double, Optional ByVal ShiftX As Double, Optional ByVal ShiftY As Double, Optional ByVal OriginX As Double, Optional ByVal OriginY As Double, Optional ByVal UseMask As Boolean) As IPicture
' IOlePicture based wrapper for FreeImage function FreeImage_RotateEx()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
Dim hDIBdst As LongPtr
    Select Case FreeImage_GetBPP(hDIBsrc)
    Case 8, 24, 32:             hDIBdst = FreeImage_RotateEx(hDIBsrc, Angle, ShiftX, ShiftY, OriginX, OriginY, UseMask)
        Set FreeImage_RotateExIOP = FreeImage_GetOlePicture(hDIBdst, , True)
    End Select
    Call FreeImage_Unload(hDIBsrc)
End Function
Public Function FreeImage_FlipHorizontalIOP(ByRef Picture As IPicture) As IPicture
' IOlePicture based wrapper for FreeImage function FreeImage_FlipHorizontal()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
    Call p_FreeImage_FlipHorizontal(hDIBsrc)
    Set FreeImage_FlipHorizontalIOP = FreeImage_GetOlePicture(hDIBsrc, , True)
End Function
Public Function FreeImage_FlipVerticalIOP(ByRef Picture As IPicture) As IPicture
' IOlePicture based wrapper for FreeImage function FreeImage_FlipVertical()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
    Call p_FreeImage_FlipVertical(hDIBsrc)
    Set FreeImage_FlipVerticalIOP = FreeImage_GetOlePicture(hDIBsrc, , True)
End Function
Public Function FreeImage_AdjustCurveIOP(ByRef Picture As IPicture, ByRef LookupTable As Variant, Optional ByVal Channel As FREE_IMAGE_COLOR_CHANNEL = FICC_BLACK) As IPicture
' IOlePicture based wrapper for FreeImage function FreeImage_AdjustCurve()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
    Select Case FreeImage_GetBPP(hDIBsrc)
    Case 8, 24, 32: Call FreeImage_AdjustCurveEx(hDIBsrc, LookupTable, Channel)
    Set FreeImage_AdjustCurveIOP = FreeImage_GetOlePicture(hDIBsrc, , True)
    End Select
End Function
Public Function FreeImage_AdjustGammaIOP(ByRef Picture As IPicture, ByVal gamma As Double) As IPicture
' IOlePicture based wrapper for FreeImage function FreeImage_AdjustGamma()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
    Select Case FreeImage_GetBPP(hDIBsrc)
    Case 8, 24, 32: Call p_FreeImage_AdjustGamma(hDIBsrc, gamma)
        Set FreeImage_AdjustGammaIOP = FreeImage_GetOlePicture(hDIBsrc, , True)
    End Select
End Function
Public Function FreeImage_AdjustBrightnessIOP(ByRef Picture As IPicture, ByVal Percentage As Double) As IPicture
' IOlePicture based wrapper for FreeImage function FreeImage_AdjustBrightness()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
    Select Case FreeImage_GetBPP(hDIBsrc)
    Case 8, 24, 32: Call p_FreeImage_AdjustBrightness(hDIBsrc, Percentage)
        Set FreeImage_AdjustBrightnessIOP = FreeImage_GetOlePicture(hDIBsrc, , True)
    End Select
End Function
Public Function FreeImage_AdjustContrastIOP(ByRef Picture As IPicture, ByVal Percentage As Double) As IPicture
' IOlePicture based wrapper for FreeImage function FreeImage_AdjustContrast()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
    Select Case FreeImage_GetBPP(hDIBsrc)
    Case 8, 24, 32: Call p_FreeImage_AdjustContrast(hDIBsrc, Percentage)
        Set FreeImage_AdjustContrastIOP = FreeImage_GetOlePicture(hDIBsrc, , True)
    End Select
End Function
Public Function FreeImage_InvertIOP(ByRef Picture As IPicture) As IPicture
' IOlePicture based wrapper for FreeImage function FreeImage_Invert()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
    Call p_FreeImage_Invert(hDIBsrc)
    Set FreeImage_InvertIOP = FreeImage_GetOlePicture(hDIBsrc, , True)
End Function
Public Function FreeImage_GetChannelIOP(ByRef Picture As IPicture, ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As IPicture
' IOlePicture based wrapper for FreeImage function FreeImage_GetChannel()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
Dim hDIBdst As LongPtr
    Select Case FreeImage_GetBPP(hDIBsrc)
    Case 24, 32: hDIBdst = FreeImage_GetChannel(hDIBsrc, Channel)
        Set FreeImage_GetChannelIOP = FreeImage_GetOlePicture(hDIBdst, , True)
    End Select
    Call FreeImage_Unload(hDIBsrc)
End Function
Public Function FreeImage_SetChannelIOP(ByRef Picture As IPicture, ByVal BitmapSrc As Long, ByVal Channel As FREE_IMAGE_COLOR_CHANNEL) As IPicture
' IOlePicture based wrapper for FreeImage function FreeImage_SetChannel()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
    Select Case FreeImage_GetBPP(hDIBsrc)
    Case 24, 32: If (FreeImage_SetChannel(hDIBsrc, BitmapSrc, Channel)) Then Set FreeImage_SetChannelIOP = FreeImage_GetOlePicture(hDIBsrc, , True)
    End Select
    Call FreeImage_Unload(hDIBsrc)
End Function
Public Function FreeImage_CopyIOP(ByRef Picture As IPicture, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As IPicture
' IOlePicture based wrapper for FreeImage function FreeImage_Copy()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
Dim hDIBdst As LongPtr:         hDIBdst = FreeImage_Copy(hDIBsrc, Left, Top, Right, Bottom)
    If (hDIBdst) Then Set FreeImage_CopyIOP = FreeImage_GetOlePicture(hDIBdst, , True)
    Call FreeImage_Unload(hDIBsrc)
End Function
Public Function FreeImage_PasteIOP(ByRef PictureDst As IPicture, ByRef PictureSrc As IPicture, ByVal Left As Long, ByVal Top As Long, ByVal Alpha As Long, Optional ByVal KeepOriginalDestImage As Boolean) As IPicture
' IOlePicture based wrapper for FreeImage function FreeImage_Paste()
Dim hDIBdst As LongPtr:         hDIBdst = FreeImage_CreateFromOlePicture(PictureDst): If (hDIBdst = 0) Then Exit Function
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(PictureSrc): If (hDIBsrc = 0) Then Exit Function
    If FreeImage_Paste(hDIBdst, hDIBsrc, Left, Top, Alpha) Then
        Set FreeImage_PasteIOP = FreeImage_GetOlePicture(hDIBdst, , True)
        If (Not KeepOriginalDestImage) Then Set PictureDst = FreeImage_PasteIOP
    End If
    Call FreeImage_Unload(hDIBsrc)
End Function
Public Function FreeImage_CompositeIOP(ByRef Picture As IPicture, Optional ByVal UseFileBackColor As Boolean, Optional ByVal AppBackColor As OLE_COLOR, Optional ByRef BackgroundPicture As IPicture) As IPicture
' IOlePicture based wrapper for FreeImage function FreeImage_Composite()
Dim hDIBsrc As LongPtr:         hDIBsrc = FreeImage_CreateFromOlePicture(Picture): If (hDIBsrc = 0) Then Exit Function
Dim lUseFileBackColor As Long:  If (UseFileBackColor) Then lUseFileBackColor = 1
Dim hDIBbgd As LongPtr:         hDIBbgd = FreeImage_CreateFromOlePicture(BackgroundPicture)
Dim hDIBdst As LongPtr:         hDIBdst = FreeImage_Composite(hDIBsrc, lUseFileBackColor, ConvertColor(AppBackColor), hDIBbgd)
    If (hDIBdst) Then Set FreeImage_CompositeIOP = FreeImage_GetOlePicture(hDIBdst, , True)
    Call FreeImage_Unload(hDIBsrc)
    If (hDIBbgd) Then Call FreeImage_Unload(hDIBbgd)
End Function
'----------------------
' VB-coded Toolkit functions
'----------------------
Public Function FreeImage_GetColorizedPalette(ByVal Color As OLE_COLOR, Optional ByVal SplitValue As Variant = 0.5) As RGBQUAD()

Dim lSplitIndex As Long
Dim lSplitIndexInv As Long
   ' compute the split index
   Select Case VarType(SplitValue)
   Case vbByte, vbInteger, vbLong:      lSplitIndex = SplitValue
   Case vbDouble, vbSingle, vbDecimal:  lSplitIndex = 256 * SplitValue
   Case Else:                           lSplitIndex = 128
   End Select
   ' check ranges of split index
   If (lSplitIndex < 0) Then
      lSplitIndex = 0
   ElseIf (lSplitIndex > 255) Then
      lSplitIndex = 255
   End If
   lSplitIndexInv = 256 - lSplitIndex
   ' extract color components red, green and blue
Dim lRed As Long:   lRed = (Color And &HFF)
Dim lGreen As Long: lGreen = ((Color \ &H100&) And &HFF)
Dim lBlue As Long:  lBlue = ((Color \ &H10000) And &HFF)
Dim atPalette(255) As RGBQUAD, i As Long
   For i = 0 To lSplitIndex - 1
      With atPalette(i)
         .rgbRed = (lRed / lSplitIndex) * i
         .rgbGreen = (lGreen / lSplitIndex) * i
         .rgbBlue = (lBlue / lSplitIndex) * i
      End With
   Next i
   For i = lSplitIndex To 255
      With atPalette(i)
         .rgbRed = lRed + ((255 - lRed) / lSplitIndexInv) * (i - lSplitIndex)
         .rgbGreen = lGreen + ((255 - lGreen) / lSplitIndexInv) * (i - lSplitIndex)
         .rgbBlue = lBlue + ((255 - lBlue) / lSplitIndexInv) * (i - lSplitIndex)
      End With
   Next i
   FreeImage_GetColorizedPalette = atPalette
End Function
Public Function FreeImage_Colorize(ByVal BITMAP As LongPtr, ByVal Color As OLE_COLOR, Optional ByVal SplitValue As Variant = 0.5) As LongPtr
    If (BITMAP = 0) Then Exit Function
    If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to colorize a 'header-only' bitmap.")
    FreeImage_Colorize = FreeImage_ConvertToGreyscale(BITMAP)
    Call FreeImage_SetPalette(FreeImage_Colorize, FreeImage_GetColorizedPalette(Color, SplitValue))
End Function
Public Function FreeImage_Sepia(ByVal BITMAP As LongPtr, Optional ByVal SplitValue As Variant = 0.5) As LongPtr: FreeImage_Sepia = FreeImage_Colorize(BITMAP, &H658AA2, SplitValue): End Function ' RGB(162, 138, 101)
'----------------------
' Compression functions wrappers
'----------------------
Public Function FreeImage_ZLibCompressEx(ByRef Target As Variant, Optional ByRef TargetSize As Long, Optional ByRef Source As Variant, Optional ByVal SourceSize As Long, Optional ByVal Offset As Long) As Long
' is a more VB friendly wrapper for compressing data with
' the 'FreeImage_ZLibCompress' function.
' The parameter 'Target' may either be a VB style array of Byte, Integer
' or Long or a pointer to a memory block. If 'Target' is a pointer to a
' memory block (when it contains an address), 'TargetSize' must be
' specified and greater than zero. If 'Target' is an initialized array,
' the whole array will be used to store compressed data when 'TargetSize'
' is missing or below or equal to zero. If 'TargetSize' is specified, only
' the first TargetSize bytes of the array will be used.
' In each case, all rules according to the FreeImage API documentation
' apply, what means that the target buffer must be at least 0.1% greater
' than the source buffer plus 12 bytes.
' If 'Target' is an uninitialized array, the contents of 'TargetSize'
' will be ignored and the size of the array 'Target' will be handled
' internally. When the function returns, 'Target' will be initialized
' as an array of Byte and sized correctly to hold all the compressed
' data.
' Nearly all, that is true for the parameters 'Target' and 'TargetSize',
' is also true for 'Source' and 'SourceSize', expect that 'Source' should
' never be an uninitialized array. In that case, the function returns
' immediately.
' The optional parameter 'Offset' may contain a number of bytes to remain
' untouched at the beginning of 'Target', when an uninitialized array is
' provided through 'Target'. When 'Target' is either a pointer or an
' initialized array, 'Offset' will be ignored. This parameter is currently
' used by 'FreeImage_ZLibCompressVB' to store the length of the uncompressed
' data at the first four bytes of 'Target'.
' get the pointer and the size in bytes of the source memory block
Dim lSourceDataPtr As LongPtr
Dim lTargetDataPtr As LongPtr
Dim bTargetCreated As Boolean
   lSourceDataPtr = p_GetMemoryBlockPtrFromVariant(Source, SourceSize)
   If (lSourceDataPtr) Then
      ' when we got a valid pointer, get the pointer and the size in bytes
      ' of the target memory block
      lTargetDataPtr = p_GetMemoryBlockPtrFromVariant(Target, TargetSize)
      If (lTargetDataPtr = 0) Then
         ' if 'Target' is a null pointer, we will initialized it as an array
         ' of bytes; here we will take 'Offset' into account
         ReDim Target(SourceSize + Int(SourceSize * 0.1) + 12 + Offset) As Byte
         ' get pointer and size in bytes (will never be a null pointer)
         lTargetDataPtr = p_GetMemoryBlockPtrFromVariant(Target, TargetSize)
         ' adjust according to 'Offset'
         lTargetDataPtr = lTargetDataPtr + Offset
         TargetSize = TargetSize - Offset
         bTargetCreated = True
      End If
      ' compress source data
      FreeImage_ZLibCompressEx = FreeImage_ZLibCompress(lTargetDataPtr, TargetSize, lSourceDataPtr, SourceSize)
      ' the function returns the number of bytes needed to store the
      ' compressed data or zero on failure
      If (FreeImage_ZLibCompressEx) Then
         If (bTargetCreated) Then
            ' when we created the array, we need to adjust it's size
            ' according to the length of the compressed data
            ReDim Preserve Target(FreeImage_ZLibCompressEx - 1 + Offset)
         End If
      End If
   End If
End Function
Public Function FreeImage_ZLibUncompressEx(ByRef Target As Variant, Optional ByRef TargetSize As Long, Optional ByRef Source As Variant, Optional ByVal SourceSize As Long) As Long
' is a more VB friendly wrapper for compressing data with the 'FreeImage_ZLibUncompress' function.
' The parameter 'Target' may either be a VB style array of Byte, Integer
' or Long or a pointer to a memory block. If 'Target' is a pointer to a
' memory block (when it contains an address), 'TargetSize' must be
' specified and greater than zero. If 'Target' is an initialized array,
' the whole array will be used to store uncompressed data when 'TargetSize'
' is missing or below or equal to zero. If 'TargetSize' is specified, only
' the first TargetSize bytes of the array will be used.
' In each case, all rules according to the FreeImage API documentation
' apply, what means that the target buffer must be at least as large, to
' hold all the uncompressed data.
' Unlike the function 'FreeImage_ZLibCompressEx', 'Target' can not be
' an uninitialized array, since the size of the uncompressed data can
' not be determined by the ZLib functions, but must be specified by a
' mechanism outside the FreeImage compression functions' scope.
' Nearly all, that is true for the parameters 'Target' and 'TargetSize',
' is also true for 'Source' and 'SourceSize'.

' get the pointer and the size in bytes of the source memory block
Dim lSourceDataPtr As LongPtr
Dim lTargetDataPtr As LongPtr
   lSourceDataPtr = p_GetMemoryBlockPtrFromVariant(Source, SourceSize)
   If (lSourceDataPtr) Then
      ' when we got a valid pointer, get the pointer and the size in bytes
      ' of the target memory block
      lTargetDataPtr = p_GetMemoryBlockPtrFromVariant(Target, TargetSize)
      If (lTargetDataPtr) Then
         ' if we do not have a null pointer, uncompress the data
         FreeImage_ZLibUncompressEx = FreeImage_ZLibUncompress(lTargetDataPtr, TargetSize, lSourceDataPtr, SourceSize)
      End If
   End If
End Function
Public Function FreeImage_ZLibGZipEx(ByRef Target As Variant, Optional ByRef TargetSize As Long, Optional ByRef Source As Variant, Optional ByVal SourceSize As Long, Optional ByVal Offset As Long) As Long
' is a more VB friendly wrapper for compressing data with the 'FreeImage_ZLibGZip' function.
' The parameter 'Target' may either be a VB style array of Byte, Integer
' or Long or a pointer to a memory block. If 'Target' is a pointer to a
' memory block (when it contains an address), 'TargetSize' must be
' specified and greater than zero. If 'Target' is an initialized array,
' the whole array will be used to store compressed data when 'TargetSize'
' is missing or below or equal to zero. If 'TargetSize' is specified, only
' the first TargetSize bytes of the array will be used.
' In each case, all rules according to the FreeImage API documentation
' apply, what means that the target buffer must be at least 0.1% greater
' than the source buffer plus 24 bytes.
' If 'Target' is an uninitialized array, the contents of 'TargetSize'
' will be ignored and the size of the array 'Target' will be handled
' internally. When the function returns, 'Target' will be initialized
' as an array of Byte and sized correctly to hold all the compressed data.
' Nearly all, that is true for the parameters 'Target' and 'TargetSize',
' is also true for 'Source' and 'SourceSize', expect that 'Source' should
' never be an uninitialized array. In that case, the function returns immediately.
' The optional parameter 'Offset' may contain a number of bytes to remain
' untouched at the beginning of 'Target', when an uninitialized array is
' provided through 'Target'. When 'Target' is either a pointer or an
' initialized array, 'Offset' will be ignored. This parameter is currently
' used by 'FreeImage_ZLibGZipVB' to store the length of the uncompressed
' data at the first four bytes of 'Target'.
' get the pointer and the size in bytes of the source memory block
Dim lSourceDataPtr As LongPtr
Dim lTargetDataPtr As LongPtr
Dim bTargetCreated As Boolean
   lSourceDataPtr = p_GetMemoryBlockPtrFromVariant(Source, SourceSize)
   If (lSourceDataPtr) Then
      ' when we got a valid pointer, get the pointer and the size in bytes
      ' of the target memory block
      lTargetDataPtr = p_GetMemoryBlockPtrFromVariant(Target, TargetSize)
      If (lTargetDataPtr = 0) Then
         ' if 'Target' is a null pointer, we will initialized it as an array
         ' of bytes; here we will take 'Offset' into account
         ReDim Target(SourceSize + Int(SourceSize * 0.1) + 24 + Offset) As Byte
         ' get pointer and size in bytes (will never be a null pointer)
         lTargetDataPtr = p_GetMemoryBlockPtrFromVariant(Target, TargetSize)
         ' adjust according to 'Offset'
         lTargetDataPtr = lTargetDataPtr + Offset
         TargetSize = TargetSize - Offset
         bTargetCreated = True
      End If
      ' compress source data
      FreeImage_ZLibGZipEx = FreeImage_ZLibGZip(lTargetDataPtr, TargetSize, lSourceDataPtr, SourceSize)
      ' the function returns the number of bytes needed to store the
      ' compressed data or zero on failure
      If (FreeImage_ZLibGZipEx) Then
         If (bTargetCreated) Then
            ' when we created the array, we need to adjust it's size
            ' according to the length of the compressed data
            ReDim Preserve Target(FreeImage_ZLibGZipEx - 1 + Offset)
         End If
      End If
   End If
End Function
Public Function FreeImage_ZLibCRC32Ex(ByVal CRC As Long, Optional ByRef Source As Variant, Optional ByVal SourceSize As Long) As Long
' is a more VB friendly wrapper for compressing data with
' the 'FreeImage_ZLibCRC32' function.
' The parameter 'Source' may either be a VB style array of Byte, Integer
' or Long or a pointer to a memory block. If 'Source' is a pointer to a
' memory block (when it contains an address), 'SourceSize' must be
' specified and greater than zero. If 'Source' is an initialized array,
' the whole array will be used to calculate the new CRC when 'SourceSize'
' is missing or below or equal to zero. If 'SourceSize' is specified, only
' the first SourceSize bytes of the array will be used.
' get the pointer and the size in bytes of the source memory block
Dim lSourceDataPtr As LongPtr
   lSourceDataPtr = p_GetMemoryBlockPtrFromVariant(Source, SourceSize)
   If (lSourceDataPtr) Then
      ' if we do not have a null pointer, calculate the CRC including 'crc'
      FreeImage_ZLibCRC32Ex = FreeImage_ZLibCRC32(CRC, lSourceDataPtr, SourceSize)
   End If
End Function
Public Function FreeImage_ZLibGUnzipEx(ByRef Target As Variant, Optional ByRef TargetSize As Long, Optional ByRef Source As Variant, Optional ByVal SourceSize As Long) As Long
' is a more VB friendly wrapper for compressing data with
' the 'FreeImage_ZLibGUnzip' function.
' The parameter 'Target' may either be a VB style array of Byte, Integer
' or Long or a pointer to a memory block. If 'Target' is a pointer to a
' memory block (when it contains an address), 'TargetSize' must be
' specified and greater than zero. If 'Target' is an initialized array,
' the whole array will be used to store uncompressed data when 'TargetSize'
' is missing or below or equal to zero. If 'TargetSize' is specified, only
' the first TargetSize bytes of the array will be used.
' In each case, all rules according to the FreeImage API documentation
' apply, what means that the target buffer must be at least as large, to
' hold all the uncompressed data.
' Unlike the function 'FreeImage_ZLibGZipEx', 'Target' can not be
' an uninitialized array, since the size of the uncompressed data can
' not be determined by the ZLib functions, but must be specified by a
' mechanism outside the FreeImage compression functions' scope.
' Nearly all, that is true for the parameters 'Target' and 'TargetSize',
' is also true for 'Source' and 'SourceSize'.

Dim lSourceDataPtr As LongPtr
Dim lTargetDataPtr As LongPtr
' get the pointer and the size in bytes of the source memory block
   lSourceDataPtr = p_GetMemoryBlockPtrFromVariant(Source, SourceSize)
   If (lSourceDataPtr) Then
      ' when we got a valid pointer, get the pointer and the size in bytes
      ' of the target memory block
      lTargetDataPtr = p_GetMemoryBlockPtrFromVariant(Target, TargetSize)
      If (lTargetDataPtr) Then
         ' if we do not have a null pointer, uncompress the data
         FreeImage_ZLibGUnzipEx = FreeImage_ZLibGUnzip(lTargetDataPtr, TargetSize, lSourceDataPtr, SourceSize)
      End If
   End If
End Function
Public Function FreeImage_ZLibCompressVB(ByRef Data() As Byte, Optional ByVal IncludeSize As Boolean = True) As Byte()
' is another, even more VB friendly wrapper for the FreeImage
' 'FreeImage_ZLibCompress' function, that uses the 'FreeImage_ZLibCompressEx'
' function. This function is very easy to use, since it deals only with VB
' style Byte arrays.
' The parameter 'Data()' is a Byte array, providing the uncompressed source
' data, that will be compressed.
' The optional parameter 'IncludeSize' determines whether the size of the
' uncompressed data should be stored in the first four bytes of the returned
' byte buffer containing the compressed data or not. When 'IncludeSize' is
' True, the size of the uncompressed source data will be stored. This works
' in conjunction with the corresponding 'FreeImage_ZLibUncompressVB' function.
' The function returns a VB style Byte array containing the compressed data.
' start population the memory block with compressed data
' at offset 4 bytes, when the unclompressed size should
' be included
Dim lOffset As Long
Dim lpArrayDataPtr As LongPtr
   If (IncludeSize) Then
      lOffset = 4
   End If
   Call FreeImage_ZLibCompressEx(FreeImage_ZLibCompressVB, , Data, , lOffset)
   If (IncludeSize) Then
      ' get the pointer actual pointing to the array data of
      ' the Byte array 'FreeImage_ZLibCompressVB'
      lpArrayDataPtr = p_DeRefPtr(p_DeRefPtr(VarPtrArray(FreeImage_ZLibCompressVB)) + 12)
      ' copy uncompressed size into the first 4 bytes
      Call CopyMemory(ByVal lpArrayDataPtr, UBound(Data) + 1, 4)
   End If
End Function
Public Function FreeImage_ZLibUncompressVB(ByRef Data() As Byte, Optional ByVal SizeIncluded As Boolean = True, Optional ByVal SizeNeeded As Long) As Byte()
' is another, even more VB friendly wrapper for the FreeImage
' 'FreeImage_ZLibUncompress' function, that uses the 'FreeImage_ZLibUncompressEx'
' function. This function is very easy to use, since it deals only with VB
' style Byte arrays.
' The parameter 'Data()' is a Byte array, providing the compressed source
' data that will be uncompressed either withthe size of the uncompressed
' data included or not.
' When the optional parameter 'SizeIncluded' is True, the function assumes,
' that the first four bytes contain the size of the uncompressed data as a
' Long value. In that case, 'SizeNeeded' will be ignored.
' When the size of the uncompressed data is not included in the buffer 'Data()'
' containing the compressed data, the optional parameter 'SizeNeeded' must
' specify the size in bytes needed to hold all the uncompressed data.
' The function returns a VB style Byte array containing the uncompressed data.
Dim abBuffer() As Byte
   If (SizeIncluded) Then
      ' get uncompressed size from the first 4 bytes and allocate
      ' buffer accordingly
      Call CopyMemory(SizeNeeded, Data(0), 4)
      ReDim abBuffer(SizeNeeded - 1)
      Call FreeImage_ZLibUncompressEx(abBuffer, , VarPtr(Data(4)), UBound(Data) - 3)
      Call p_Swap(VarPtrArray(FreeImage_ZLibUncompressVB), VarPtrArray(abBuffer))
   ElseIf (SizeNeeded) Then
      ' no size included in compressed data, so just forward the
      ' call to 'FreeImage_ZLibUncompressEx' and trust on SizeNeeded
      ReDim abBuffer(SizeNeeded - 1)
      Call FreeImage_ZLibUncompressEx(abBuffer, , Data)
      Call p_Swap(VarPtrArray(FreeImage_ZLibUncompressVB), VarPtrArray(abBuffer))
   End If
End Function
Public Function FreeImage_ZLibGZipVB(ByRef Data() As Byte, Optional ByVal IncludeSize As Boolean = True) As Byte()
' is another, even more VB friendly wrapper for the FreeImage
' 'FreeImage_ZLibGZip' function, that uses the 'FreeImage_ZLibGZipEx'
' function. This function is very easy to use, since it deals only with VB
' style Byte arrays.
' The parameter 'Data()' is a Byte array, providing the uncompressed source
' data that will be compressed.
' The optional parameter 'IncludeSize' determines whether the size of the
' uncompressed data should be stored in the first four bytes of the returned
' byte buffer containing the compressed data or not. When 'IncludeSize' is
' True, the size of the uncompressed source data will be stored. This works
' in conjunction with the corresponding 'FreeImage_ZLibGUnzipVB' function.
' The function returns a VB style Byte array containing the compressed data.
' start population the memory block with compressed data
' at offset 4 bytes, when the unclompressed size should be included
Dim lOffset As Long
Dim lpArrayDataPtr As LongPtr
   If (IncludeSize) Then
      lOffset = 4
   End If
   Call FreeImage_ZLibGZipEx(FreeImage_ZLibGZipVB, , Data, , lOffset)
   If (IncludeSize) Then
      ' get the pointer actual pointing to the array data of
      ' the Byte array 'FreeImage_ZLibCompressVB'
      lpArrayDataPtr = p_DeRefPtr(p_DeRefPtr(VarPtrArray(FreeImage_ZLibGZipVB)) + 12)
      ' copy uncompressed size into the first 4 bytes
      Call CopyMemory(ByVal lpArrayDataPtr, UBound(Data) + 1, 4)
   End If
End Function
Public Function FreeImage_ZLibGUnzipVB(ByRef Data() As Byte, Optional ByVal SizeIncluded As Boolean = True, Optional ByVal SizeNeeded As Long) As Byte()
' is another, even more VB friendly wrapper for the FreeImage
' 'FreeImage_ZLibGUnzip' function, that uses the 'FreeImage_ZLibGUnzipEx'
' function. This function is very easy to use, since it deals only with VB style Byte arrays.
' The parameter 'Data()' is a Byte array, providing the compressed source
' data that will be uncompressed either withthe size of the uncompressed data included or not.
' When the optional parameter 'SizeIncluded' is True, the function assumes,
' that the first four bytes contain the size of the uncompressed data as a
' Long value. In that case, 'SizeNeeded' will be ignored.
' When the size of the uncompressed data is not included in the buffer 'Data()'
' containing the compressed data, the optional parameter 'SizeNeeded' must
' specify the size in bytes needed to hold all the uncompressed data.
' The function returns a VB style Byte array containing the uncompressed data.
Dim abBuffer() As Byte
   If (SizeIncluded) Then
      ' get uncompressed size from the first 4 bytes and allocate
      ' buffer accordingly
      Call CopyMemory(SizeNeeded, Data(0), 4)
      ReDim abBuffer(SizeNeeded - 1)
      Call FreeImage_ZLibGUnzipEx(abBuffer, , VarPtr(Data(4)), UBound(Data) - 3)
      Call p_Swap(VarPtrArray(FreeImage_ZLibGUnzipVB), VarPtrArray(abBuffer))
   ElseIf (SizeNeeded) Then
      ' no size included in compressed data, so just forward the
      ' call to 'FreeImage_ZLibUncompressEx' and trust on SizeNeeded
      ReDim abBuffer(SizeNeeded - 1)
      Call FreeImage_ZLibGUnzipEx(abBuffer, , Data)
      Call p_Swap(VarPtrArray(FreeImage_ZLibGUnzipVB), VarPtrArray(abBuffer))
   End If
End Function
'----------------------
' Public functions to destroy custom safearrays
'----------------------
Public Function FreeImage_DestroyLockedArray(ByRef Data As Variant) As Long
' Destroys an array, that was self created with a custom  array descriptor of type ('fFeatures' member) 'FADF_AUTO Or FADF_FIXEDSIZE'.
Dim lpArrayPtr As LongPtr
' Such arrays are returned by mostly all of the array-dealing wrapper
' functions. Since these should not destroy the actual array data, when
' going out of scope, they are craeted as 'FADF_FIXEDSIZE'.'
' So, VB sees them as fixed or temporarily locked, when you try to manipulate
' the array's dimensions. There will occur some strange effects, you should
' know about:
' 1. When trying to 'ReDim' the array, this run-time error will occur:
'    Error #10, 'This array is fixed or temporarily locked'
' 2. When trying to assign another array to the array variable, this
'    run-time error will occur:
'    Error #13, 'Type mismatch'
' 3. The 'Erase' statement has no effect on the array
' Although VB clears up these arrays correctly, when the array variable
' goes out of scope, you have to destroy the array manually, when you want
' to reuse the array variable in current scope.
' For an example assume, that you want do walk all scanlines in an image:
' For i = 0 To FreeImage_GetHeight(Bitmap)
'
'    ' assign scanline-array to array variable
'    abByte = FreeImage_GetScanLineEx(Bitmap, i)
'
'    ' do some work on it...
'
'    ' destroy the array (only the array, not the actual data)
'    Call FreeImage_DestroyLockedArray(dbByte)
' Next i
' The function returns zero on success and any other value on failure
' !! Attention !!
' uses a Variant parameter for passing the array to be
' destroyed. Since VB does not allow to pass an array of non public
' structures through a Variant parameter, this function can not be used
' with arrays of cutom type.
' You will get this compiler error: "Only public user defined types defined
' in public object modules can be used as parameters or return types for
' public procedures of class modules or as fields of public user defined types"
' So, there is a function in the wrapper called 'FreeImage_DestroyLockedArrayByPtr'
' that takes a pointer to the array variable which can be used to work around
' that VB limitation and furthermore can be used for any of these self-created
' arrays. To get the array variable's pointer, a declared version of the
' VB 'VarPtr' function can be used which works for all types of arrays expect
' String arrays. Declare this function like this in your code:
' Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" ( _
      ByRef Ptr() As Any) As Long
' Then an array could be destroyed by calling the 'FreeImage_DestroyLockedArrayByPtr'
' function like this:
' lResult = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(MyLockedArray))
' Additionally there are some handy wrapper functions available, one for each
' commonly used structure in FreeImage like RGBTRIPLE, RGBQUAD, FICOMPLEX etc.

' Currently, these functions do return 'FADF_AUTO Or FADF_FIXEDSIZE' arrays
' that must be destroyed using this or any of it's derived functions:
' FreeImage_GetPaletteEx()           with FreeImage_DestroyLockedArrayRGBQUAD()
' FreeImage_GetPaletteLong()         with FreeImage_DestroyLockedArray()
' FreeImage_SaveToMemoryEx2()        with FreeImage_DestroyLockedArray()
' FreeImage_AcquireMemoryEx()        with FreeImage_DestroyLockedArray()
' FreeImage_GetScanLineEx()          with FreeImage_DestroyLockedArray()
' FreeImage_GetScanLineBITMAP8()     with FreeImage_DestroyLockedArray()
' FreeImage_GetScanLineBITMAP16()    with FreeImage_DestroyLockedArray()
' FreeImage_GetScanLineBITMAP24()    with FreeImage_DestroyLockedArrayRGBTRIPLE()
' FreeImage_GetScanLineBITMAP32()    with FreeImage_DestroyLockedArrayRGBQUAD()
' FreeImage_GetScanLineINT16()       with FreeImage_DestroyLockedArray()
' FreeImage_GetScanLineINT32()       with FreeImage_DestroyLockedArray()
' FreeImage_GetScanLineFLOAT()       with FreeImage_DestroyLockedArray()
' FreeImage_GetScanLineDOUBLE()      with FreeImage_DestroyLockedArray()
' FreeImage_GetScanLineCOMPLEX()     with FreeImage_DestroyLockedArrayFICOMPLEX()
' FreeImage_GetScanLineRGB16()       with FreeImage_DestroyLockedArrayFIRGB16()
' FreeImage_GetScanLineRGBA16()      with FreeImage_DestroyLockedArrayFIRGBA16()
' FreeImage_GetScanLineRGBF()        with FreeImage_DestroyLockedArrayFIRGBF()
' FreeImage_GetScanLineRGBAF()       with FreeImage_DestroyLockedArrayFIRGBAF()
' ensure, this is an array
   If (VarType(Data) And vbArray) Then
      ' data is a VB array, what means a SAFEARRAY in C/C++, that is
      ' passed through a ByRef Variant variable, that is a pointer to
      ' a VARIANTARG structure
      ' the VARIANTARG structure looks like this:
      ' typedef struct tagVARIANT VARIANTARG;
      ' struct tagVARIANT
      '     {
      '     Union
      '         {
      '         struct __tagVARIANT
      '             {
      '             VARTYPE vt;
      '             WORD wReserved1;
      '             WORD wReserved2;
      '             WORD wReserved3;
      '             Union
      '                 {
      '                 [...]
      '             SAFEARRAY *parray;    // used when not VT_BYREF
      '                 [...]
      '             SAFEARRAY **pparray;  // used when VT_BYREF
      '                 [...]
      ' the data element (SAFEARRAY) has an offset of 8, since VARTYPE
      ' and WORD both have a length of 2 bytes; the pointer to the
      ' VARIANTARG structure is the VarPtr of the Variant variable in VB
      ' getting the contents of the data element (in C/C++: *(data + 8))
      lpArrayPtr = p_DeRefPtr(VarPtr(Data) + 8)
      ' call the 'FreeImage_DestroyLockedArrayByPtr' function to destroy
      ' the array properly
      Call FreeImage_DestroyLockedArrayByPtr(lpArrayPtr)
   Else
      FreeImage_DestroyLockedArray = -1
   End If
End Function
Public Function FreeImage_DestroyLockedArrayByPtr(ByVal arrayPtr As LongPtr) As Long
' Destroys a self-created array with a custom array descriptor by a pointer to the array variable.
Dim lpSA As LongPtr: lpSA = p_DeRefPtr(arrayPtr) ' dereference the pointer once (in C/C++: *ArrayPtr)
   ' now 'lpSA' is a pointer to the actual SAFEARRAY structure
   ' and could be a null pointer when the array is not initialized
   ' then, we have nothing to do here but return (-1) to indicate an "error"
   If (lpSA) Then
      ' destroy the array descriptor
      Call SafeArrayDestroyDescriptor(lpSA)
      ' make 'lpSA' a null pointer, that is an uninitialized array;
      ' keep in mind, that we here use 'ArrayPtr' as a ByVal argument,
      ' since 'ArrayPtr' is a pointer to lpSA (the address of lpSA);
      ' we need to zero these four bytes, 'ArrayPtr' points to
      Call CopyMemory(ByVal arrayPtr, 0&, PTR_LENGTH)
   Else
      ' the array is already uninitialized, so return an "error" value
      FreeImage_DestroyLockedArrayByPtr = -1
   End If
End Function
Public Function FreeImage_DestroyLockedArrayRGBTRIPLE(ByRef Data() As RGBTRIPLE) As Long: FreeImage_DestroyLockedArrayRGBTRIPLE = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(Data)): End Function
Public Function FreeImage_DestroyLockedArrayRGBQUAD(ByRef Data() As RGBQUAD) As Long: FreeImage_DestroyLockedArrayRGBQUAD = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(Data)): End Function
Public Function FreeImage_DestroyLockedArrayFICOMPLEX(ByRef Data() As FICOMPLEX) As Long: FreeImage_DestroyLockedArrayFICOMPLEX = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(Data)): End Function
Public Function FreeImage_DestroyLockedArrayFIRGB16(ByRef Data() As FIRGB16) As Long: FreeImage_DestroyLockedArrayFIRGB16 = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(Data)): End Function
Public Function FreeImage_DestroyLockedArrayFIRGBA16(ByRef Data() As FIRGBA16) As Long: FreeImage_DestroyLockedArrayFIRGBA16 = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(Data)): End Function
Public Function FreeImage_DestroyLockedArrayFIRGBF(ByRef Data() As FIRGBF) As Long: FreeImage_DestroyLockedArrayFIRGBF = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(Data)): End Function
Public Function FreeImage_DestroyLockedArrayFIRGBAF(ByRef Data() As FIRGBAF) As Long: FreeImage_DestroyLockedArrayFIRGBAF = FreeImage_DestroyLockedArrayByPtr(VarPtrArray(Data)): End Function
'----------------------
' Private IOlePicture related helper functions
'----------------------
Private Function p_GetIOlePictureFromContainer(ByRef Container As Object, Optional ByVal IncludeDrawings As Boolean) As IPicture
' Returns a VB IOlePicture object (IPicture) from a VB image hosting control.
   If (Not Container Is Nothing) Then
      Select Case TypeName(Container)
      Case "PictureBox", "Form"
         If (IncludeDrawings) Then
            If (Not Container.AutoRedraw) Then
               Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Custom drawings can only be included into the DIB when " & "the container's 'AutoRedraw' property is set to True.")
               Exit Function
            End If
            Set p_GetIOlePictureFromContainer = Container.Image
         Else
            Set p_GetIOlePictureFromContainer = Container.Picture
         End If
      Case Else
Dim bHasPicture As Boolean
Dim bHasImage As Boolean
Dim bIsAutoRedraw As Boolean
         On Error Resume Next
         bHasPicture = (Container.Picture <> 0)
         bHasImage = (Container.Image <> 0)
         bIsAutoRedraw = Container.AutoRedraw
         On Error GoTo 0
         If ((IncludeDrawings) And (bHasImage) And (bIsAutoRedraw)) Then
            Set p_GetIOlePictureFromContainer = Container.Image
         ElseIf (bHasPicture) Then
            Set p_GetIOlePictureFromContainer = Container.Picture
         Else
            Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Cannot create DIB from container control. Container " & "control has no 'Picture' property.")
         End If
      End Select
   End If
End Function
'----------------------
' Private image and color helper functions
'----------------------
Private Function p_GetPreviousColorDepth(ByVal BPP As Long) As Long
' Returns the 'previous' color depth of a given color depth. Here, 'previous' means the next smaller color depth.
   Select Case BPP
   Case 32: p_GetPreviousColorDepth = 24
   Case 24: p_GetPreviousColorDepth = 16
   Case 16: p_GetPreviousColorDepth = 15
   Case 15: p_GetPreviousColorDepth = 8
   Case 8:  p_GetPreviousColorDepth = 4
   Case 4:  p_GetPreviousColorDepth = 1
   End Select
End Function
Private Function p_GetNextColorDepth(ByVal BPP As Long) As Long
' Returns the 'next' color depth of a given color depth. Here, 'next' means the next greater color depth.
   Select Case BPP
   Case 1:  p_GetNextColorDepth = 4
   Case 4:  p_GetNextColorDepth = 8
   Case 8:  p_GetNextColorDepth = 15
   Case 15: p_GetNextColorDepth = 16
   Case 16: p_GetNextColorDepth = 24
   Case 24: p_GetNextColorDepth = 32
   End Select
End Function
Private Function GCD(ByVal a As Variant, ByVal b As Variant) As Variant
' calculate greatest common divisor
Dim vntTemp As Variant
   Do While (b)
      vntTemp = b
      ' calculate b = a % b (modulo)
      ' this could be just:
      ' b = a Mod b
      ' but VB's Mod operator fails for unsigned
      ' long values stored in currency variables
      ' so, we use the mathematical definition of
      ' the modulo operator taken from Wikipedia.
      b = a - p_Floor(a / b) * b
      a = vntTemp
   Loop
   GCD = a
End Function
Private Function p_Floor(ByRef a As Variant) As Variant
' This is a VB version of the floor() function.
   If (a < 0) Then
      p_Floor = VBA.Int(a)
   Else
      p_Floor = -VBA.Fix(-a)
   End If
End Function
Private Function p_GetValueBuffer(ByRef Value As Variant, ByVal MetaDataVarType As FREE_IMAGE_MDTYPE, ByRef ElementSize As Long, ByRef Buffer() As Byte) As Long
Dim lElementCount As Long
Dim bIsArray As Boolean
Dim abValueBuffer(7) As Byte
Dim cBytes As Long
Dim i As Long
' copies any value provided in the Variant 'Value'
' parameter into the byte array Buffer. 'Value' may contain an
' array as well. The values in the byte buffer are aligned to fit
' the FreeImage data type for tag values specified in
' 'MetaDataVarType'. For integer values, it does not matter, in
' which VB data type the values are provided. For example, it is
' possible to transform a provided byte array into a unsigned long array.
' The parameter 'ElementSize' is an OUT value, providing the actual
' size per element in the byte buffer in bytes to the caller.
' works for the types FIDT_BYTE, FIDT_SHORT, FIDT_LONG,
' FIDT_SBYTE , FIDT_SSHORT, FIDT_SLONG, FIDT_FLOAT, FIDT_DOUBLE
' and FIDT_IFD
   ElementSize = p_GetElementSize(MetaDataVarType)
   If (Not IsArray(Value)) Then
      lElementCount = 1
   Else
      On Error Resume Next
      lElementCount = UBound(Value) - LBound(Value) + 1
      On Error GoTo 0
      bIsArray = True
   End If
   If (lElementCount > 0) Then
      ReDim Buffer((lElementCount * ElementSize) - 1)
      If (Not bIsArray) Then
         cBytes = p_GetVariantAsByteBuffer(Value, abValueBuffer)
         If (cBytes > ElementSize) Then
            cBytes = ElementSize
         End If
         Call CopyMemory(Buffer(0), abValueBuffer(0), cBytes)
      Else
         For i = LBound(Value) To UBound(Value)
            cBytes = p_GetVariantAsByteBuffer(Value(i), abValueBuffer)
            If (cBytes > ElementSize) Then
               cBytes = ElementSize
            End If
            Call CopyMemory(Buffer(0 + (i * ElementSize)), abValueBuffer(0), cBytes)
         Next i
      End If
      p_GetValueBuffer = lElementCount
   End If
End Function
Private Function p_GetRationalValueBuffer(ByRef RationalValues() As FIRATIONAL, ByRef Buffer() As Byte) As Long
Dim lElementCount As Long
Dim abValueBuffer(7) As Byte
Dim cBytes As Long
Dim i As Long
   ' copies a number of elements from the FIRATIONAL array
   ' 'RationalValues' into the byte buffer 'Buffer'.
   ' From the caller's point of view, this function is the same as
   ' 'p_GetValueBuffer', except, it only works for arrays of FIRATIONAL.
   ' works for the types FIDT_RATIONAL and FIDT_SRATIONAL.
   lElementCount = UBound(RationalValues) - LBound(RationalValues) + 1
   ReDim Buffer(lElementCount * 8 + 1)
   For i = LBound(RationalValues) To UBound(RationalValues)
      cBytes = p_GetVariantAsByteBuffer(RationalValues(i).Numerator, abValueBuffer)
      If (cBytes > 4) Then
         cBytes = 4
      End If
      Call CopyMemory(Buffer(0 + (i * 8)), abValueBuffer(0), cBytes)
      cBytes = p_GetVariantAsByteBuffer(RationalValues(i).Denominator, abValueBuffer)
      If (cBytes > 4) Then
         cBytes = 4
      End If
      Call CopyMemory(Buffer(4 + (i * 8)), abValueBuffer(0), cBytes)
   Next i
   p_GetRationalValueBuffer = lElementCount
End Function
Private Function p_GetVariantAsByteBuffer(ByRef Value As Variant, ByRef Buffer() As Byte) As Long
Dim lLength As Long
   ' fills a byte buffer 'Buffer' with data taken
   ' from a Variant parameter. Depending on the Variant's type and,
   ' width, it copies N (lLength) bytes into the buffer starting
   ' at the buffer's first byte at Buffer(0). The function returns
   ' the number of bytes copied.
   ' It is much easier to assign the Variant to a variable of
   ' the proper native type first, since gathering a Variant's
   ' actual value is a hard job to implement for each subtype.
   Select Case VarType(Value)
   Case vbByte: Buffer(0) = Value: lLength = 1
   Case vbInteger: Dim iInteger As Integer: iInteger = Value: lLength = 2: Call CopyMemory(Buffer(0), iInteger, lLength)
   Case vbLong: Dim lLong As Long: lLong = Value: lLength = 4: Call CopyMemory(Buffer(0), lLong, lLength)
#If VBA7 Then           '<OFFICE2010+>
   Case vbLongLong: Dim lLongLong As LongLong: lLongLong = Value: lLength = 8: Call CopyMemory(Buffer(0), lLongLong, lLength)
#End If                 '<WIN32>
   Case vbCurrency: Dim cCurrency As Currency: cCurrency = Value / 10000: lLength = 8: Call CopyMemory(Buffer(0), cCurrency, lLength) ' since the Currency data type is a so called scaled integer, we have to divide by 10.000 first to get the proper bit layout.
   Case vbSingle: Dim sSingle As Single: sSingle = Value: lLength = 4: Call CopyMemory(Buffer(0), sSingle, lLength)
   Case vbDouble: Dim dblDouble As Double: dblDouble = Value: lLength = 8: Call CopyMemory(Buffer(0), dblDouble, lLength)
   End Select
   p_GetVariantAsByteBuffer = lLength
End Function
Private Function p_GetElementSize(ByVal vt As FREE_IMAGE_MDTYPE) As Long
' returns the width in bytes for any of the FreeImage metadata tag data types.
   Select Case vt
   Case FIDT_BYTE, FIDT_SBYTE, FIDT_UNDEFINED, FIDT_ASCII:          p_GetElementSize = 1
   Case FIDT_SHORT, FIDT_SSHORT:                                    p_GetElementSize = 2
   Case FIDT_LONG, FIDT_SLONG, FIDT_FLOAT, FIDT_PALETTE, FIDT_IFD:  p_GetElementSize = 4
   Case Else:                                                       p_GetElementSize = 8
   End Select
End Function
'----------------------
' Private pointer manipulation helper functions
'----------------------
Private Function p_GetStringFromPointerA(ByRef Ptr As LongPtr) As String
' creates and returns a VB BSTR variable from a C/C++ style string pointer by making a redundant deep copy of the string's characters.
Dim abBuffer() As Byte
Dim lLength As Long
   If (Ptr) Then
      ' get the length of the ANSI string pointed to by ptr
      lLength = lstrlen(Ptr)
      If (lLength) Then
         ' copy characters to a byte array
         ReDim abBuffer(lLength - 1)
         Call CopyMemory(abBuffer(0), ByVal Ptr, lLength)
         ' convert from byte array to unicode BSTR
         p_GetStringFromPointerA = StrConv(abBuffer, vbUnicode)
      End If
   End If
End Function
Private Function p_DeRefLong(ByVal Ptr As LongPtr) As Long: Call CopyMemory(p_DeRefLong, ByVal Ptr, 4): End Function
#If VBA7 Then           '<OFFICE2010+>
Private Function p_DeRefLongLong(ByVal Ptr As LongPtr) As LongLong: Call CopyMemory(p_DeRefLongLong, ByVal Ptr, 8): End Function
#End If                 '<OFFICE2010+>
Private Function p_DeRefPtr(ByVal Ptr As LongPtr) As LongPtr: Call CopyMemory(p_DeRefPtr, ByVal Ptr, PTR_LENGTH): End Function
Private Sub p_Swap(ByVal lpSrc As LongPtr, ByVal lpDst As LongPtr)
' swaps two DWORD memory blocks pointed to by lpSrc and lpDst, whereby lpSrc and lpDst are actually no pointer types but contain the pointer's address.
' in C/C++ this would be:
' void p_Swap(int lpSrc, int lpDst) {
'   int tmp = *(int*)lpSrc;
'   *(int*)lpSrc = *(int*)lpDst;
'   *(int*)lpDst = tmp;
' }
Dim lpTmp As LongPtr
   Call CopyMemory(lpTmp, ByVal lpSrc, PTR_LENGTH)
   Call CopyMemory(ByVal lpSrc, ByVal lpDst, PTR_LENGTH)
   Call CopyMemory(ByVal lpDst, lpTmp, PTR_LENGTH)
End Sub
Private Function p_GetMemoryBlockPtrFromVariant(ByRef Data As Variant, Optional ByRef SizeInBytes As Long, Optional ByRef ElementSize As Long) As LongPtr
' returns the pointer to the memory block provided through
' the Variant parameter 'data', which could be either a Byte, Integer or
' Long array or the address of the memory block itself. In the last case,
' the parameter 'SizeInBytes' must not be omitted or zero, since it's
' correct value (the size of the memory block) can not be determined by
' the address only. So, the function fails, if 'SizeInBytes' is omitted
' or zero and 'data' is not an array but contains a Long value (the address
' of a memory block) by returning Null.
' If 'data' contains either a Byte, Integer or Long array, the pointer to
' the actual array data is returned. The parameter 'SizeInBytes' will
' be adjusted correctly, if it was less or equal zero upon entry.
' The function returns Null (zero) if there was no supported memory block
' provided.
   ' do we have an array?
   If (VarType(Data) And vbArray) Then
      Select Case (VarType(Data) And (Not vbArray))
      Case vbByte
         ElementSize = 1
         p_GetMemoryBlockPtrFromVariant = p_GetArrayPtrFromVariantArray(Data)
         If (p_GetMemoryBlockPtrFromVariant) Then
            If (SizeInBytes <= 0) Then
               SizeInBytes = (UBound(Data) + 1)
            ElseIf (SizeInBytes > (UBound(Data) + 1)) Then
               SizeInBytes = (UBound(Data) + 1)
            End If
         End If
      Case vbInteger
         ElementSize = 2
         p_GetMemoryBlockPtrFromVariant = p_GetArrayPtrFromVariantArray(Data)
         If (p_GetMemoryBlockPtrFromVariant) Then
            If (SizeInBytes <= 0) Then
               SizeInBytes = (UBound(Data) + 1) * 2
            ElseIf (SizeInBytes > ((UBound(Data) + 1) * 2)) Then
               SizeInBytes = (UBound(Data) + 1) * 2
            End If
         End If
      Case vbLong
         ElementSize = 4
         p_GetMemoryBlockPtrFromVariant = p_GetArrayPtrFromVariantArray(Data)
         If (p_GetMemoryBlockPtrFromVariant) Then
            If (SizeInBytes <= 0) Then
               SizeInBytes = (UBound(Data) + 1) * 4
            ElseIf (SizeInBytes > ((UBound(Data) + 1) * 4)) Then
               SizeInBytes = (UBound(Data) + 1) * 4
            End If
         End If
#If VBA7 Then           '<OFFICE2010+>
      Case vbLongLong
         ElementSize = 8
         p_GetMemoryBlockPtrFromVariant = p_GetArrayPtrFromVariantArray(Data)
         If (p_GetMemoryBlockPtrFromVariant) Then
            If (SizeInBytes <= 0) Then
               SizeInBytes = (UBound(Data) + 1) * 8
            ElseIf (SizeInBytes > ((UBound(Data) + 1) * 8)) Then
               SizeInBytes = (UBound(Data) + 1) * 8
            End If
         End If
#End If                 '<WIN32>
      End Select
   Else
      ElementSize = 1
      If ((VarType(Data) = vbLong) And (SizeInBytes >= 0)) Then
         p_GetMemoryBlockPtrFromVariant = Data
      End If
   End If
End Function
Private Function p_GetArrayPtrFromVariantArray(ByRef Data As Variant) As LongPtr
' returns a pointer to the first array element of
' a VB array (SAFEARRAY) that is passed through a Variant type
' parameter. (Don't try this at home...)
Dim lDataPtr As LongPtr
   
   ' cache VarType in variable
Dim eVarType As VbVarType: eVarType = VarType(Data)
   
   ' ensure, this is an array
   If (eVarType And vbArray) Then
      
      ' data is a VB array, what means a SAFEARRAY in C/C++, that is
      ' passed through a ByRef Variant variable, that is a pointer to
      ' a VARIANTARG structure
      
      ' the VARIANTARG structure looks like this:
      
      ' typedef struct tagVARIANT VARIANTARG;
      ' struct tagVARIANT
      '     {
      '     Union
      '         {
      '         struct __tagVARIANT
      '             {
      '             VARTYPE vt;
      '             WORD wReserved1;
      '             WORD wReserved2;
      '             WORD wReserved3;
      '             Union
      '                 {
      '                 [...]
      '             SAFEARRAY *parray;    // used when not VT_BYREF
      '                 [...]
      '             SAFEARRAY **pparray;  // used when VT_BYREF
      '                 [...]
      
      ' the data element (SAFEARRAY) has an offset of 8, since VARTYPE
      ' and WORD both have a length of 2 bytes; the pointer to the
      ' VARIANTARG structure is the VarPtr of the Variant variable in VB
      
      ' getting the contents of the data element (in C/C++: *(data + 8))
      lDataPtr = p_DeRefPtr(VarPtr(Data) + 8)
      
      ' dereference the pointer again (in C/C++: *(lDataPtr))
      lDataPtr = p_DeRefPtr(lDataPtr)
      
      ' test, whether 'lDataPtr' now is a Null pointer
      ' in that case, the array is not yet initialized and so we can't dereference
      ' it another time since we have no permisson to acces address 0
      
      ' the contents of 'lDataPtr' may be Null now in case of an uninitialized
      ' array; then we can't access any of the SAFEARRAY members since the array
      ' variable doesn't event point to a SAFEARRAY structure, so we will return
      ' the null pointer
      
      If (lDataPtr) Then
         ' the contents of lDataPtr now is a pointer to the SAFEARRAY structure
            
         ' the SAFEARRAY structure looks like this:
         
         ' typedef struct FARSTRUCT tagSAFEARRAY {
         '    unsigned short cDims;       // Count of dimensions in this array.
         '    unsigned short fFeatures;   // Flags used by the SafeArray
         '                                // routines documented below.
         ' #if defined(WIN32)
         '    unsigned long cbElements;   // Size of an element of the array.
         '                                // Does not include size of
         '                                // pointed-to data.
         '    unsigned long cLocks;       // Number of times the array has been
         '                                // locked without corresponding unlock.
         ' #Else
         '    unsigned short cbElements;
         '    unsigned short cLocks;
         '    unsigned long handle;       // Used on Macintosh only.
         ' #End If
         '    void HUGEP* pvData;               // Pointer to the data.
         '    SAFEARRAYBOUND rgsabound[1];      // One bound for each dimension.
         ' } SAFEARRAY;
         ' in WIN32, the pvData element has an offset of 12 bytes from the base address of the structure,
         ' in WIN64, the pvData element has an offset of 16 bytes from the base address of the structure,
         ' so dereference the pvData pointer, what indeed is a pointer
         ' to the actual array (in C/C++: *(lDataPtr + 12)) or + 16)) (x64)
         lDataPtr = p_DeRefPtr(lDataPtr + 8 + PTR_LENGTH)
      End If
      
      ' return this value
      p_GetArrayPtrFromVariantArray = lDataPtr
      
      ' a more shorter form of this function would be:
      ' (doesn't work for uninitialized arrays, but will likely crash!)
      'p_GetArrayPtrFromVariantArray = pDeref(pDeref(pDeref(VarPtr(data) + 8)) + 8 + PTR_LENGTH)
   End If
End Function
'----------------------
' FreeImage based User functions
'----------------------
Public Function FreeImage_DrawText(Text As Variant, Optional ByVal hFont As LongPtr, _
    Optional ByVal DestWidth As Long, Optional ByVal DestHeight As Long, Optional ByVal Alignment As Long, _
    Optional FontColor As Long, Optional BackColor As Long = &HFF000000) As LongPtr
' create FIBITMAP with text on it
' Text      - text data to output (maybe string or array of strings)
' hFont     - handle of the font that'll be used to draw text
' DestWidth/DestHeight - size of text output region (in pixels)
' Alignment - text alignment
' FontColor - text color (VB color Style - BGR), also understand ole colors
' BackColor - background (VB color Style - BGR), < 0 - transparent, 0..&hFFFFFF - solid colors, also understand ole colors
'-------------------------
' v.1.0.0       : 24.12.2021 - исходная версия
'-------------------------
' ToDo: ???
'-------------------------
' http://www.vb-helper.com/howto_memory_bitmap_text.html
' https://www.codeproject.com/Articles/2102/Drawing-Lines-Shapes-or-Text-on-Bitmaps
' hack hinted in following link & remarked below in the code
' http://www.tech-archive.net/Archive/Development/microsoft.public.win32.programmer.gdi/2006-02/msg00111.html
'-------------------------
Const dX = 1, dY = 0 ' костыли пытаемся скорректировать обрезку пикселей
Dim Result As Long: Result = NOERROR
    On Error GoTo HandleError
' get Text lines
Dim sText As String, sLines() As String
Dim i As Long, iMin As Long, iMax As Long
    ' for the first time we take array from the end to get real text size,
    ' then from the begining to output it
    If IsArray(Text) Then
        If Len(Trim(Join(Text, vbNullString))) = 0 Then Err.Raise vbObjectError + 512   ' no text error
        sLines = Text: iMin = LBound(sLines): iMax = UBound(sLines): i = iMax
        sText = sLines(i)
    Else
        If Len(Trim(Text)) = 0 Then Err.Raise vbObjectError + 512                       ' no text error
        sText = Text
    End If
' get Font
'    If hFont = 0& Then Err.Raise vbObjectError + 512
    If hFont = 0 Then hFont = GetStockObject(SYSTEM_FONT) ' If hFont is null then get default font
Dim tFont As LongPtr
Dim LF As LOGFONT: GetObject hFont, LenB(LF), LF
' don't use rotated fonts; user rotates via Angle parameter
    If (LF.lfEscapement <> 0 Or LF.lfOrientation <> 0) Then LF.lfEscapement = 0: LF.lfOrientation = 0:    tFont = CreateFontIndirect(LF): If tFont <> 0 Then hFont = tFont
' prepare DC
Const NEWTRANSPARENT As Long = 3
Dim BITMAP As LongPtr
Dim hdc As LongPtr, hMemDC As LongPtr
Dim hMemBMP As LongPtr, hMemOldBMP As LongPtr
Dim hOldFont As LongPtr, hBrush As LongPtr
' create compatible memory DC
Dim bReleaseDC As Boolean
'    If (hDC = 0) Then hDC = GetDC(0): bReleaseDC = (hDC <> 0)
'    If (hDC = 0) Then Err.Raise vbObjectError + 512
    hdc = GetDC(0): bReleaseDC = (hdc <> 0)
    If (hdc = 0) Then Err.Raise vbObjectError + 512
    hMemDC = CreateCompatibleDC(hdc)
' select hFont into DC
    hOldFont = SelectObject(hMemDC, hFont)
Dim cY As Long
Dim Wt As Long, Ht As Long      '
Dim yStep As Long, tmOverhang As Long
Dim sz As POINT, tm As TEXTMETRIC
Dim cRect As RECT

' get text size with GetTextExtentPoint32 function
    If iMax > iMin Then
    ' if there is more then one text line
        Do While i > iMin
            GetTextExtentPoint32 hMemDC, StrPtr(sText), Len(sText), sz
            Ht = Ht + sz.cY: If sz.cX > Wt Then Wt = sz.cX
            i = i - 1: sText = sLines(i)
        Loop
    End If
    GetTextExtentPoint32 hMemDC, StrPtr(sText), Len(sText), sz: yStep = sz.cY
    Ht = Ht + sz.cY: If sz.cX > Wt Then Wt = sz.cX
    GetTextMetrics hMemDC, tm: tmOverhang = tm.tmOverhang   ' size correction for italic and bold fonts
    Wt = Wt + tmOverhang                                    ' add extra space used when bold and/or italic
Dim Wt1 As Long: Wt1 = DestWidth:  If Wt1 = 0 Then Wt1 = Wt 'Wt1 = DestWidth + dX: If Wt1 = 0 Then Wt1 = Wt
Dim Ht1 As Long: Ht1 = DestHeight: If Ht1 = 0 Then Ht1 = Ht 'Ht1 = estHeight + dY: If Ht1 = 0 Then Ht1 = Ht
' create compatible bitmap
    hMemBMP = CreateCompatibleBitmap(hdc, Wt1, Ht1)
' select compatible bitmap
    hMemOldBMP = SelectObject(hMemDC, hMemBMP)
' translate ole colors
Dim vbColorBack As Long: If (OleTranslateColor(BackColor, 0, vbColorBack) <> 0) Then vbColorBack = BackColor
Dim vbColorFont As Long: If (OleTranslateColor(FontColor, 0, vbColorFont) <> 0) Then vbColorFont = FontColor
' fill background
    If vbColorBack >= 0 Then hBrush = CreateSolidBrush(vbColorBack)
    If hBrush = 0 Then
    ' transparent background
        ' per the hack, by forcing White text on black background, GDI renders
        ' anti-aliasing in grayscale. We can get that grayscale and apply it as
        ' an Alpha ratio a little later and also fill in the true forecolor
        SetTextColor hMemDC, vbWhite
        SetBkColor hMemDC, vbBlack
    Else
    ' non transparent - fill background with color
        SetTextColor hMemDC, vbColorFont
        cRect.Right = Wt1 + 1: cRect.Bottom = Ht1 + 1
        FillRect hMemDC, cRect, hBrush
    End If
' Set transparent background then draw string
    hOldFont = SelectObject(hMemDC, hFont)
    Call SetBkMode(hMemDC, NEWTRANSPARENT)  ' make font background transparent anyway
    If hBrush <> 0& Then DeleteObject hBrush
    'If Not hMemDC = m_hDC Then ReleaseDC 0&, hMemDC
' Set text alignment
Dim xAnchor As Long
    If (Alignment And TA_CENTER) = TA_CENTER Then
        xAnchor = (Wt1 - tmOverhang) \ 2
    ElseIf (Alignment And TA_RIGHT) = TA_RIGHT Then
        xAnchor = (Wt1 - tmOverhang)
    Else
        xAnchor = 0
    End If
    If (Alignment And TA_BASELINE) = TA_BASELINE Then
        cY = (Ht1 - Ht - tmOverhang) \ 2
        Alignment = Alignment Xor TA_BASELINE
    ElseIf (Alignment And TA_BOTTOM) = TA_BOTTOM Then
        cY = Ht1 - Ht - tmOverhang
        Alignment = Alignment Xor TA_BOTTOM
    Else
        cY = 0
    End If
    xAnchor = xAnchor + dX: cY = cY + dY
Dim hOldAlign As Long: hOldAlign = SetTextAlign(hMemDC, Alignment)

' text output
    ' Note: Not using DrawText API. TextOutW supported on Win9x & unicode.
    ' DrawTextW not supported on Win9x
    Do
        If cY + yStep > 0 Then TextOut hMemDC, xAnchor, cY, StrPtr(sText), Len(sText)
        i = i + 1: If (i > iMax) Or (cY > Ht1) Then Exit Do
        cY = cY + yStep: sText = sLines(i)
    '' we assume that the height of the string for this font is constant? so no need change step each time
        'GetTextExtentPoint32 hMemDC, StrPtr(sText), Len(sText), sz: yStep = sz.cy
    Loop
' end output
' Copy the device context into the BitmapDst for future manipulations
    BITMAP = FreeImage_CreateFromDC(hMemDC, hMemBMP)
' Destroy the new font.
    SelectObject hMemDC, hOldFont
    SetTextAlign hMemDC, hOldAlign
' restore the alpha channel for text colors
    Call p_RestoreAlphaChannel(BITMAP, vbColorFont, vbColorBack, IIf(hBrush = 0, 0, &HFF))
    Call SelectObject(hMemDC, hMemOldBMP)
'Call FreeImage_PreMultiplyWithAlpha(BITMAP)
    ' clean up
    If hBrush <> 0 Then DeleteObject (hBrush)
    Call DeleteObject(hMemBMP)
    Call DeleteDC(hMemDC)
HandleExit:  If (bReleaseDC) Then Call ReleaseDC(0, hdc)
             FreeImage_DrawText = BITMAP: Exit Function
HandleError: BITMAP = 0: Err.Clear: Resume HandleExit
End Function
Public Function FreeImage_SetToControl(fiBITMAP As LongPtr, ByRef ObjectControl As Object) As Long
' set FIBITMAP to control
    If fiBITMAP = 0 Then Exit Function
    If ObjectControl Is Nothing Then Exit Function
Dim Result As Long: Result = NOERROR
    On Error GoTo HandleError
Stop
'Dim Wp As Long, Hp As Long: Wp = FreeImage_GetWidth(FIBITMAP): Hp = FreeImage_GetHeight(FIBITMAP)
'    With ObjectControl
'Dim Wc As Long: Wc = ConvTwipsToPixels(.Width, 0)
'Dim Hc As Long: Hc = ConvTwipsToPixels(.Height, 1)
'' rescale picture
'Dim k As Double: If Wp > 0 And Hp > 0 Then k = fMin(Wc / Wp, Hc / Hp) Else k = 4.94065645841247E-324
'Dim hTemp As LongPtr: hTemp = FreeImage_RescaleEx(FIBITMAP, k, k, False, False, FILTER_BICUBIC, True)
'        Wp = FreeImage_GetWidth(hTemp): Hp = FreeImage_GetHeight(hTemp)
'' rotate picture
'Dim dAngle As Double:    dAngle = fMod(angle, 360) ' нормализуем угол
'        If dAngle <> 0 Then hTemp = FreeImage_RotateEx(hTemp, dAngle, 0, 0, Wp / 2, Hp / 2, True)
'' set picture to image
'        .PictureData = FreeImage_GetPictureDataEMF(hTemp, True)
'    End With
HandleExit:  FreeImage_SetToControl = Result: Exit Function
HandleError: Result = Err: Err.Clear: Resume HandleExit
End Function
Public Function FreeImage_CompositeWithAlpha(fiDst As LongPtr, fiSrc As LongPtr, _
    Optional ByVal Left As Long, Optional ByVal Top As Long, Optional ByVal Width As Long, Optional ByVal Height As Long, _
    Optional Alpha As Long = &HFF, Optional PreMultiply As Boolean = False, Optional IgnoreBackAlpha As Boolean = True) As Long
' composite two FIBITMAPS with AlphaBlend (set fiSrc on fiDst and return result in fiDst)
'-------------------------
' fiDst - background FIBITMAP
' fiSrc - foreground FIBITMAP
' Left/Top/Width/Height - pos and size of foreground bitmap rect on back
' Alpha - Alpha for foreground bitmap
' PreMultiply - (not used) if True - do PreMultiply with Alpha
' IgnoreBackAlpha - (not used) if True - make background solid
'-------------------------
' create and select hBitmap in hDC
Dim Result As Long: Result = NOERROR
Dim hdc As LongPtr, hMemDC As LongPtr
Dim hMemBMP As LongPtr, hMemOldBMP As LongPtr
Dim bReleaseDC As Boolean
'    If (hDC = 0) Then hDC = GetDC(0): bReleaseDC = (hDC <> 0)
'    If (hDC = 0) Then Err.Raise vbObjectError + 512
    hdc = GetDC(0): bReleaseDC = (hdc <> 0)
    If (hdc = 0) Then Err.Raise vbObjectError + 512
    hMemDC = CreateCompatibleDC(hdc)
' create compatible bitmap
    hMemBMP = CreateCompatibleBitmap(hdc, FreeImage_GetWidth(fiDst), FreeImage_GetHeight(fiDst))
' select compatible bitmap
    hMemOldBMP = SelectObject(hMemDC, hMemBMP)
' do some deeds
'' load fore picture into DC with alpha (over background)
'    If PreMultiply Then Call FreeImage_PreMultiplyWithAlpha(fiDst)
' patch alpha channel on background
    'If IgnoreBackAlpha Then Call p_SetAlphaChannel(fiDst, Alpha:=&H0)
' load background into DC
    Call FreeImage_PaintDC(hMemDC, fiDst) 'Call FreeImage_PaintTransparent(hMemDC, fiDst)
' create FIBITMAP from DCS
'If IsDebug Then
'Stop
'Dim fiTemp As LongPtr: fiTemp = FreeImage_CreateFromDC(hMemDC)
'FreeImage_Save FIF_BMP, fiTemp, CurrentProject.path & "\fiBackFromDC.bmp"
'End If

' load fore picture into DC with alpha (over background)
'    If PreMultiply Then Call FreeImage_PreMultiplyWithAlpha(fiSrc)
' calculate size coords of the Src bitmap part that we will paint on Dst
    ' the source rectangle must lie completely within the source surface,
    ' otherwise an error occurs and the function returns FALSE.
    ' AlphaBlend fails if the width or height of the source
    ' or destination is negative.
    If Width <= 0 Then Width = FreeImage_GetWidth(fiSrc)
    If Height <= 0 Then Height = FreeImage_GetHeight(fiSrc)
    Call FreeImage_PaintTransparent(hMemDC, fiSrc, _
        XDst:=Left, YDst:=Top, WidthSrc:=Width, HeightSrc:=Height, _
        Alpha:=Alpha) 'Alpha
'If IsDebug Then
'Stop
'    fiTemp = FreeImage_CreateFromDC(hMemDC)
'FreeImage_Save FIF_BMP, fiTemp, CurrentProject.path & "\fiBackFromDC.bmp"
'End If
' create FIBITMAP from DCS
    fiDst = FreeImage_CreateFromDC(hMemDC)
' clean up
    Call SelectObject(hMemDC, hMemOldBMP)
    Call DeleteObject(hMemBMP)
    Call DeleteDC(hMemDC)
    If (bReleaseDC) Then Call ReleaseDC(0, hdc)
HandleExit:  FreeImage_CompositeWithAlpha = Result: Exit Function
HandleError: Result = Err: Err.Clear: Resume HandleExit
End Function

Public Function FreeImage_RotateExEx(ByVal BITMAP As LongPtr, ByVal Angle As Double, _
    Optional ByRef Color As Long = 0&, Optional UnloadSource As Boolean) As LongPtr ', _
' Function performs a rotation using a 3rd order (cubic) B-Spline,
' but rotated image retains size and aspect ratio of source image (destination image size is usually bigger)
' Rotation occurs around the center of the image area.
    On Error GoTo HandleExit
    If BITMAP = 0 Then Exit Function
' !!! держит центр, НО режет один пиксель если информация выходит за границы
' сначала расширить изображение, чтобы после поворота изображение поместилось целиком
' потом повернуть относительно центра с учетом смещения угла изображения
' и наконец обрезать по размерам после поворота
Dim radA As Single: radA = Angle * Pi / 180
Dim SinA As Single: SinA = Abs(Sin(radA))
Dim CosA As Single: CosA = Abs(Cos(radA))
' получаем размеры изображения до поворота
Dim Wp As Long:     Wp = FreeImage_GetWidth(BITMAP)
Dim Hp As Long:     Hp = FreeImage_GetHeight(BITMAP)
' размеры области изображения после поворота
Dim Wp1 As Single:  Wp1 = Wp * CosA + Hp * SinA
Dim Hp1 As Single:  Hp1 = Wp * SinA + Hp * CosA
' величина необходимого расширения области, чтобы после поворота изображение поместилось целиком
Dim dl As Long, dr As Long:    If Wp1 > Wp Then dl = (Wp1 - Wp + 1) \ 2: dr = dl ' +1 для округления вверх
Dim dt As Long, db As Long:    If Hp1 > Hp Then dt = (Hp1 - Hp + 1) \ 2: db = dt
' расширяем область изображения
Dim fiSrc As LongPtr, fiDst As LongPtr: fiSrc = BITMAP
    fiSrc = FreeImage_EnlargeCanvas(BITMAP, dl, dt, dr, db, Color, FI_COLOR_IS_RGBA_COLOR)
' поворачиваем изображение
    fiDst = p_FreeImage_RotateEx(fiSrc, Angle, 0, 0, FreeImage_GetWidth(fiSrc) / 2, FreeImage_GetHeight(fiSrc) / 2, 1): Call FreeImage_Unload(fiSrc)
' величина необходимого сжатия области после поворота (обрезка пустот)
    If (Wp1 < Wp) Then dl = (Wp1 - Wp) \ 2 Else dl = 0
    If (Hp1 < Hp) Then dt = (Hp1 - Hp) \ 2 Else dt = 0
    dr = dl: db = dt
' обрезаем область изображения
    FreeImage_RotateExEx = FreeImage_EnlargeCanvas(fiDst, dl, dt, dr, db, Color, FI_COLOR_IS_RGBA_COLOR): FreeImage_Unload (fiDst)
    If (UnloadSource) Then Call FreeImage_Unload(BITMAP)
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function FreeImage_ConvertToAlphaGreyScale(ByVal BITMAP As LongPtr, Optional ByVal UnloadSource As Boolean) As LongPtr '
' преобразует 32bit в greyscale сохраняя альфа канал оригинала
    If BITMAP = 0 Then Exit Function
    If FreeImage_GetBPP(BITMAP) <> 32 Then Exit Function
Dim fiSrc As LongPtr, fiDst As LongPtr: fiSrc = BITMAP
    fiSrc = FreeImage_Clone(BITMAP)
    fiDst = FreeImage_ConvertToGreyscale(BITMAP)
    fiDst = FreeImage_ConvertTo32Bits(fiDst)
Dim x As Long, y As Long
    ' overlay a RGBQUAD vs long array over the dib
Dim argbqSrc() As RGBQUAD:     argbqSrc() = FreeImage_GetBitsExRGBQUAD(fiSrc)
Dim argbqDst() As RGBQUAD:     argbqDst() = FreeImage_GetBitsExRGBQUAD(fiDst)
    For y = 0 To FreeImage_GetHeight(fiSrc) - 1
        For x = 0 To FreeImage_GetWidth(fiSrc) - 1
            argbqDst(x, y).rgbReserved = argbqSrc(x, y).rgbReserved
        Next x
    Next y
    Call FreeImage_DestroyLockedArrayRGBQUAD(argbqSrc()): FreeImage_Unload (fiSrc)
    Call FreeImage_DestroyLockedArrayRGBQUAD(argbqDst())
    FreeImage_ConvertToAlphaGreyScale = fiDst
    If UnloadSource Then FreeImage_Unload (BITMAP)
End Function

Public Function FreeImage_CreateCheckerBoard(ByVal Width As Long, ByVal Height As Long, Optional ByVal BITMAP As LongPtr, Optional Alpha As Long = -1, _
    Optional ByVal CheckerSize As Long = 16&, Optional ByVal ColorA As Long = &HFEFEFE, Optional ByVal ColorB As Long = &HC0C0C0) As LongPtr
' Create a checkerboard pattern
Dim Result As Long: Result = NOERROR
' The checker size is used for both the width and height of each square. Default value is 12.
' ColorA is the colored checker at the top left corner of the pattern. Default is white
' ColorB is the alternating checker color. Default is gray RGB: 192,192,192
    If (BITMAP = 0) Then
        If ((Width = 0) And (Height = 0)) Then Exit Function
    Else
        If (Not FreeImage_HasPixels(BITMAP)) Then Call Err.Raise(5, c_strModule, Error$(5) & vbCrLf & vbCrLf & "Unable to paint a 'header-only' bitmap.")
        If (Width = 0) Then Width = FreeImage_GetWidth(BITMAP)
        If (Height = 0) Then Height = FreeImage_GetHeight(BITMAP)
    End If
Dim fiBITMAP As LongPtr
Dim hdc As LongPtr, hMemDC As LongPtr
Dim hMemBMP As LongPtr, hMemOldBMP As LongPtr
Dim hBrush As LongPtr, hBr1 As LongPtr, hBr2 As LongPtr
Dim bReleaseDC As Boolean
'    If (hDC = 0) Then hDC = GetDC(0): bReleaseDC = (hDC <> 0)
'    If (hDC = 0) Then Err.Raise vbObjectError + 512
    hdc = GetDC(0): bReleaseDC = (hdc <> 0)
    If (hdc = 0) Then Err.Raise vbObjectError + 512
    hMemDC = CreateCompatibleDC(hdc)
' create compatible bitmap
    hMemBMP = CreateCompatibleBitmap(hdc, Width, Height)
' select compatible bitmap
    hMemOldBMP = SelectObject(hMemDC, hMemBMP)
' translate ole colors
Dim vbColorA As Long: If (OleTranslateColor(ColorA, 0, vbColorA) <> 0) Then vbColorA = ColorA
Dim vbColorB As Long: If (OleTranslateColor(ColorB, 0, vbColorB) <> 0) Then vbColorB = ColorB
' create color brushes
    hBr1 = CreateSolidBrush(vbColorA)
    hBr2 = CreateSolidBrush(vbColorB)
Dim x As Long, y As Long, bEven As Boolean
' create checker board
Dim cRect As RECT: cRect.Right = CheckerSize: cRect.Bottom = CheckerSize
    For y = 0& To Height - 1& Step CheckerSize
        If bEven Then hBrush = hBr2 Else hBrush = hBr1
        For x = 0& To Width - 1& Step CheckerSize
            FillRect hMemDC, cRect, hBrush
            If hBrush = hBr1 Then hBrush = hBr2 Else hBrush = hBr1
            OffsetRect cRect, CheckerSize, 0&
        Next x
        bEven = Not bEven
        OffsetRect cRect, -cRect.Left, CheckerSize
    Next y
    If BITMAP <> 0 Then
' paint transparent BITMAP with alpha over created CheckBoard
        Result = FreeImage_PreMultiplyWithAlpha(BITMAP)
        Result = FreeImage_PaintTransparent(hMemDC, BITMAP, Alpha:=Alpha)
    End If
' get checkboard to FIBITMAP from DC
    fiBITMAP = FreeImage_CreateFromDC(hMemDC, hMemBMP)
' patch CheckBoard alpha channel
    Call p_SetAlphaChannel(fiBITMAP, Alpha:=&HFF)
' clean up
    If (hBr1 <> 0) Then DeleteObject hBr1
    If (hBr2 <> 0) Then DeleteObject hBr2
    Call SelectObject(hMemDC, hMemOldBMP)
    If hBrush <> 0 Then DeleteObject (hBrush)
    Call DeleteObject(hMemBMP)
    Call DeleteDC(hMemDC)
HandleExit:  If (bReleaseDC) Then Call ReleaseDC(0, hdc)
             FreeImage_CreateCheckerBoard = fiBITMAP: Exit Function
HandleError: fiBITMAP = 0: Err.Clear: Resume HandleExit
End Function

Private Function p_SetAlphaChannel(fiBITMAP As LongPtr, Optional Alpha As Byte = &HFF) As Long
' set the alpha to whole pic
Dim Result As Long: Result = NOERROR
    On Error GoTo HandleError
    If FreeImage_GetBPP(fiBITMAP) <> 32 Then Err.Raise vbObjectError + 512
Dim x As Long, y As Long
    ' overlay a RGBQUAD vs long array over the dib
Dim argbqDIB() As RGBQUAD:     argbqDIB() = FreeImage_GetBitsExRGBQUAD(fiBITMAP)
    For y = 0 To FreeImage_GetHeight(fiBITMAP) - 1
        For x = 0 To FreeImage_GetWidth(fiBITMAP) - 1
            argbqDIB(x, y).rgbReserved = Alpha
        Next x
    Next y
HandleExit:  Call FreeImage_DestroyLockedArrayRGBQUAD(argbqDIB())
             p_SetAlphaChannel = Result: Exit Function
HandleError: Result = Err: Err.Clear: Resume HandleExit
End Function

Private Function p_RestoreAlphaChannel(fiBITMAP As LongPtr, Optional ForeColor As Long, Optional BackColor As Long, Optional BackAlpha As Byte) As Long
' walk thru the pixels and set the forecolor as needed
' FIBITMAP  - Handle to FreeImage picture containing processed text
' BackColor - Background color (VB style) '-1 - color is not defined
' ForeColor - Text color (VB style)       '-1 - color is not defined
' BackAlpha - Background transparency [0..255]
'---------------------
' used by FreeImage_DrawText to recover correct colors with transparency
' this code is based on c32bppDIB.spt_DrawText by LaVolpe (from psc cd)
' hack hinted in following link & remarked below in the code
' http://www.tech-archive.net/Archive/Development/microsoft.public.win32.programmer.gdi/2006-02/msg00111.html
'---------------------
' For place Fore over Back with Alpha formula is (we suppose that Back is solid and Fore is Alpha blend):
'   'ResColor = ForeColor * ForeAlpha + BackColor * (1 - ForeAlpha)
' so if we need extract ForeAlpha:
'   'ForeAlpha = (ResColor - BackColor)/(ForeColor - BackColor)
' if both Fore and Back are partitionaly transparent:
'   'ResAlpha = ForeAlpha + BackAlpha * (1 - ForeAlpha)
'   'ResColor = (ForeColor * ForeAlpha + BackColor * BackAlpha * (1 - ForeAlpha)) / ResAlpha
'---------------------
Dim Result As Long: Result = NOERROR
    On Error GoTo HandleError
    If FreeImage_GetBPP(fiBITMAP) <> 32 Then Err.Raise vbObjectError + 512
Dim x As Long, y As Long
' translate ole colors
Dim vbBackColor As Long: If (OleTranslateColor(BackColor, 0, vbBackColor) <> 0) Then vbBackColor = BackColor
Dim vbForeColor As Long: If (OleTranslateColor(ForeColor, 0, vbForeColor) <> 0) Then vbForeColor = ForeColor
Dim byForeAlpha As Long ', byBackAlpha As Byte
'Dim dbForeAlpha As Double, dbBackAlpha As Double
'    Select Case BackAlpha
'    Case &H0:   dblBackAlpha = 0
'    Case &HFF:  dblBackAlpha = 1
'    Case Else:  dblBackAlpha = BackAlpha / 255
'    End Select
'Stop
Dim alngDIB() As Long:         alngDIB() = FreeImage_GetBitsExLong(fiBITMAP)
    ' argbqDIB(x, y) is ABGR, so
    ' A = (argbqDIB(x, y))                 / 16^6
    ' B = (argbqDIB(x, y) and &h00FF0000&) / 16^4
    ' G = (argbqDIB(x, y) and &h0000FF00&) / 16^2
    ' R = (argbqDIB(x, y) and &h000000FF&)
    If BackAlpha = &H0 Then
' fully transparent background
    
'Dim argbbDIB() As Byte:     argbbDIB() = FreeImage_GetBitsEx(FIBITMAP)              ' overlay a Byte vs long array over the dib
Dim argbqDIB() As RGBQUAD:     argbqDIB() = FreeImage_GetBitsExRGBQUAD(fiBITMAP)    ' overlay a RGBQUAD vs long array over the dib
        Select Case vbForeColor
        Case vbBlack
            For y = 0 To FreeImage_GetHeight(fiBITMAP) - 1
                For x = 0 To FreeImage_GetWidth(fiBITMAP) - 1
                    With argbqDIB(x, y)
                        Select Case .rgbBlue
                        Case &H0:       alngDIB(x, y) = &H0&        ' fully transparent
                        Case &HFF:      alngDIB(x, y) = &HFF000000  ' fully opaque (nontransparent)
                        Case Is > &H7F: byForeAlpha = .rgbBlue: alngDIB(x, y) = ((byForeAlpha And &H7F&) * &H1000000) Or &H80000000
                        Case Else:      byForeAlpha = .rgbBlue: alngDIB(x, y) = &H1000000 * byForeAlpha
                        End Select
                    End With
                Next x
            Next y
        Case vbWhite
    ' white font - uses fewer calcs than other colors except black
            For y = 0 To FreeImage_GetHeight(fiBITMAP) - 1
                For x = 0 To FreeImage_GetWidth(fiBITMAP) - 1
                    With argbqDIB(x, y)
                        Select Case .rgbBlue
                        Case &H0:       alngDIB(x, y) = &H0&        ' fully transparent
                        Case &HFF:      alngDIB(x, y) = &HFFFFFFFF  ' fully opaque (nontransparent)
                        Case Is > &H7F: byForeAlpha = .rgbBlue: alngDIB(x, y) = ((byForeAlpha And &H7F&) * &H1000000 Or &H80000000) Or byForeAlpha Or (byForeAlpha * &H100) Or (byForeAlpha * &H10000)
                        Case Else:      .rgbReserved = .rgbBlue: byForeAlpha = .rgbBlue: alngDIB(x, y) = (byForeAlpha * &H1000000) Or byForeAlpha Or (byForeAlpha * &H100) Or (byForeAlpha * &H10000)
                        End Select
                    End With
                Next x
            Next y
       Case Else
    ' other color font - more intensive calcs
    ' VB-style color ABGR -> RGBQUAD
Dim rgbqForeColor As RGBQUAD: CopyMemory rgbqForeColor, ConvertColor(vbForeColor), 4& ' copy for faster reference
            For y = 0 To FreeImage_GetHeight(fiBITMAP) - 1
                For x = 0 To FreeImage_GetWidth(fiBITMAP) - 1
                    With argbqDIB(x, y)
                        Select Case .rgbBlue
                        Case &H0:       alngDIB(x, y) = &H0&        ' fully transparent
                        Case &HFF:      alngDIB(x, y) = &HFF000000 Or rgbqForeColor.rgbBlue Or (rgbqForeColor.rgbGreen * &H100&) Or (rgbqForeColor.rgbRed * &H10000)
                        Case Is > &H7F: byForeAlpha = .rgbBlue: alngDIB(x, y) = ((byForeAlpha And &H7F&) * &H1000000 Or &H80000000) Or _
                            ((rgbqForeColor.rgbBlue * byForeAlpha) \ 255) Or _
                            ((rgbqForeColor.rgbGreen * byForeAlpha) \ 255) * &H100 Or _
                            ((rgbqForeColor.rgbRed * byForeAlpha) \ 255) * &H10000
                        Case Else:      byForeAlpha = .rgbBlue: alngDIB(x, y) = (byForeAlpha * &H1000000) Or _
                            ((rgbqForeColor.rgbBlue * byForeAlpha) \ 255) Or _
                            ((rgbqForeColor.rgbGreen * byForeAlpha) \ 255) * &H100 Or _
                            ((rgbqForeColor.rgbRed * byForeAlpha) \ 255) * &H10000
                        End Select
                    End With
                Next x
            Next y
        End Select
'Stop
'Dim argbbDIB() As Byte:     argbbDIB() = FreeImage_GetBitsEx(FIBITMAP)
'ByteArray_WriteToFile argbbDIB, CurrentProject.path & "\TempTextFI.bin", False
'Call FreeImage_DestroyLockedArray(argbbDIB())
    ' done & remove the RGBQUAD ovelay
        'Call FreeImage_DestroyLockedArrayRGBQUAD(argbqDIB())
    ElseIf BackAlpha = &HFF Then
' used solid filled background
        For y = 0 To FreeImage_GetHeight(fiBITMAP) - 1
            For x = 0 To FreeImage_GetWidth(fiBITMAP) - 1
                Select Case alngDIB(x, y)
                Case 0
                Case Is < &H1000000: alngDIB(x, y) = (&HFF000000 Or alngDIB(x, y))
                End Select
            Next x
        Next y
    Else
' partitionaly transparent background
Stop
    End If
HandleExit:  Call FreeImage_DestroyLockedArray(alngDIB())
             p_RestoreAlphaChannel = Result: Exit Function
HandleError: Result = Err: Err.Clear: Resume HandleExit
End Function


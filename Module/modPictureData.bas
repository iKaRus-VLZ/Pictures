Attribute VB_Name = "modPictureData"
Option Explicit
Option Compare Database
Public IsDebug As Boolean
'=========================
Private Const c_strModule As String = "modPictureData"
'=========================
' ��������      : ������ ��� ������ � PictureData � BLOB ������
' ������        : 1.4.2.453644466
' ����          : 13.03.2024 10:43:06
' �����         : ������ �.�. (KashRus@gmail.com)
' ����������    : ��� ������ ����������� ���������� �������������� ���������� � ����������� �� ObjectDataType
'               : ������ ������ ��������� clsTransform - ����� ��� ������������� ��������� _
'               : ������������ ���� � ��������� ����: _
'               : LaVolpe  http://www.planet-source-code.com/vb/scripts/ShowCode.asp? txtCodeId=67466&lngWId=1 _
'               : �������� http://www.sql.ru/forum/348182-a/primer-jpeg-gif-iz-long-binary-polya-v-image-bez-kopirovaniya-vo-vremennyy-fayl?hl=createenhmetafile _
'               : ����������� �������� � ������� SysObjs
' v.1.4.2       : 20.12.2023 - ������ ������������� ��� ������ � �������� LaVolpe (��� �������������). �������� ������� ������ - ������� ������ �������� �� ������ ��� ��� ����
' v.1.4.1       : 12.12.2023 - ����������� ������ � PictureData_SetToControl
' v.1.4.0       : 08.12.2023 - ���������� PictureData_SetToControl
' v.1.3.2       : 11.02.2022 - ��������� ���� ������� ������ � OleObject (�� ��� ������ ����������������)
' v.1.3.0       : 02.09.2021 - ����������� ��� ���������� FreeImage
' v.1.2.4       : 06.06.2019 - ��������� � PictureData_SetToControl - ���������� ������ ���������������� ������ ��� ���������� �����������
' v.1.2.3       : 26.12.2018 - ��������� � PictureData_SetToControl - ��������� ����������. ���������� ������������, ����������� ����������� �������/����������
'=========================
' ToDo: ��������� >> PictureData_SetIcon �� �������� � x64 !!!
' - Forms.Image � FreeImage_GetOlePictureEMF ������ ������� ��������
' + PictureData_LoadFromEx - �������� ����������/�������� �� ����� � ��������� ��� ����������� ����������� ����� �� ��������� ���� � �� �� ��� �� �����
'=========================
#Const ObjectDataType = 0
' ObjectDataType = 0 - (x86/x64) ���������� ���������� FreeImage             https://freeimage.sourceforge.io/download.html
'                ��� ������ ������ ������ ��������� ������ modFreeImage (�������������� ��� x64 ������� Visual Basic Wrapper for FreeImage 3 by Carsten Klein (cklein05@users.sourceforge.net))
'                ��������!!! MSO 2003 (x86) �������� �� Load Library FreeImage.dll (v.3.18 x86), ������� ��� x86 ����������  FreeImage.dll (v.3.17 x86)
' ObjectDataType = 1 - (x86)     ���������� ���������������� ������ LaVolpe  http://www.planet-source-code.com/vb/scripts/ShowCode.asp? txtCodeId=67466&lngWId=1
'                ��� ������ ������ ������ ��������� clsPictureData[BMP|ICO|PNG|GIF|JPG] - ���������������� ������ LaVolpe ��� ������� � ������ �����������
Private Const NOERROR As Long = &H0
'=========================

' === Declare Const ===
Private Const c_strPathRes = "RES\" ' ������������� ���� � ��������� ������ (���������, ������� � ��.)

Private Const c_strObjectTable = "SysObjData"   ' ������� � ������� �������� ������ BLOB ��������
' �������� ����� �������
Private Const c_strKey = "ID"
Private Const c_strObjectKey = "NAME"
Private Const c_strObjectData = "BINDATA"
Private Const c_strObjectDesc = "DESC"
Private Const c_strObjectComm = "COMMENT"
'--------------------------------------------------------------------------------
' POINTER LENGTH
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>  PtrSafe, LongPtr and LongLong
Private Const PTR_LENGTH As Long = 8
#Else                   '<OFFICE97-2010>        Long
Private Const PTR_LENGTH As Long = 4
#End If                 '<WIN32>
'
Private Const Pi As Double = 3.14159265358979
' size convertion constants
Private Const PointsPerInch = 72
Private Const TwipsPerInch = 1440
Private Const CentimitersPerInch = 2.54                 '1 ���� = 127 / 50 ��
Private Const HimetricPerInch = 2540                    '1 ���� = 1000 * 127/50 himetrix
'HIMETRIC = (PIXEL * 2540) / 96
'PIXEL = (HIMETRIC * 96) / 2540
Private Const inch = TwipsPerInch                       '1 ���� = 1440 twips
Private Const pt = TwipsPerInch / PointsPerInch         '1 ����� = 20 twips
Private Const cm = TwipsPerInch / CentimitersPerInch    '1 �� = 566.929133858 twips
'
' ��������� OLE Object Data
Private Const LENGTH_FOR_SIZE = 4
Private Const CON_CHUNK_SIZE As Long = &H8000

Private Const OBJECT_SIGNATURE = &H1C15
Private Const OBJECT_HEADER_SIZE = &H14

Private Const OLE_HEADER_SIZE = &HC
Private Const CHECKSUM_SIGNATURE = &HFE05AD00
Private Const CHECKSUM_STRING_SIZE = 4

Private Const ALLOCSIZEINCR As Long = 65536
Private Const SYSCOLORMASK As Long = &H80000000

Private Const BlockSize = 32768

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
' Constants.
Private Const SRCAND As Long = &H8800C6
Private Const SRCCOPY As Long = &HCC0020
Private Const SRCERASE As Long = &H440328
Private Const SRCINVERT As Long = &H660046
Private Const SRCPAINT As Long = &HEE0086
Private Const CAPTUREBLT As Long = &H40000000

Private Const SHGFI_ICON As Long = &H100&
Private Const SHGFI_SMALLICON As Long = &H1&

'Font const
'used with fnWeight
Private Const FW_DONTCARE = 0
Private Const FW_THIN = 100
Private Const FW_EXTRALIGHT = 200
Private Const FW_LIGHT = 300
Private Const FW_NORMAL = 400
Private Const FW_MEDIUM = 500
Private Const FW_SEMIBOLD = 600
Private Const FW_BOLD = 700
Private Const FW_EXTRABOLD = 800
Private Const FW_HEAVY = 900
Private Const FW_BLACK = FW_HEAVY
Private Const FW_DEMIBOLD = FW_SEMIBOLD
Private Const FW_REGULAR = FW_NORMAL
Private Const FW_ULTRABOLD = FW_EXTRABOLD
Private Const FW_ULTRALIGHT = FW_EXTRALIGHT
'used with fdwCharSet
Private Const ANSI_CHARSET As Long = &H0
Private Const DEFAULT_CHARSET As Long = &H1
Private Const SYMBOL_CHARSET As Long = &H2
Private Const RUSSIAN_CHARSET As Long = &HCC
Private Const OEM_CHARSET As Long = &HFF
Private Const SHIFTJIS_CHARSET = 128
Private Const HANGEUL_CHARSET = 129
Private Const CHINESEBIG5_CHARSET = 136
'used with fdwOutputPrecision
Private Const OUT_CHARACTER_PRECIS = 2
Private Const OUT_DEFAULT_PRECIS = 0
Private Const OUT_DEVICE_PRECIS = 5
'used with fdwClipPrecision
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const CLIP_CHARACTER_PRECIS = 1
Private Const CLIP_STROKE_PRECIS = 2
'used with fdwQuality
Private Const DEFAULT_QUALITY As Long = 0
Private Const DRAFT_QUALITY  As Long = 1
Private Const PROOF_QUALITY  As Long = 2
Private Const NONANTIALIASED_QUALITY  As Long = 3
Private Const ANTIALIASED_QUALITY As Long = 4
Private Const CLEARTYPE_QUALITY As Long = 5
'used with fdwPitchAndFamily
Private Const DEFAULT_PITCH = 0
Private Const FIXED_PITCH = 1
Private Const VARIABLE_PITCH = 2
'used with SetBkMode
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2

' GDI and GDI+ constants
Private Const PLANES = 14       '  Number of planes
Private Const BITSPIXEL = 12    '  Number of bits per pixel
Private Const PATCOPY = &HF00021 ' (DWORD) dest = pattern
Private Const GDIP_WMF_PLACEABLEKEY = &H9AC6CDD7
Private Const UnitPixel = 2

Private Const FILE_ATTRIBUTE_NORMAL = &H80&

' === Declare Enums ===
#If ObjectDataType = 0 Then     'FI
Public Enum eAlignText
    TA_LEFT = 0                 '������� ����� ��������� �� ����� ������ �������� ��������������.
    TA_RIGHT = 2                '������� ����� ��������� �� ������ ������ �������� ��������������.
    TA_CENTER = 6               '������� ����� ������������� ������������� �� ������ �������� ��������������.
    TA_TOP = 0                  '������� ����� �� ������� ������ �������� ��������������.
    TA_BOTTOM = 8               '������� ����� �� ������ ������ �������� ��������������.
    TA_BASELINE = 24            '������� ����� ��������� �� ������� ����� ������.
    TA_RTLREADING = 256         '�������� Windows �� ������ �������� �������: ����� ����������� ��� ������� ������ ������ ������ , � ����������������� ������� ������ �� ��������� ����� �������. ��� ����������� ������ �����, ����� �����, ��������� � �������� ���������� ������������ ��� ��� ���������� ��� ��� ��������� �����.
'    TA_NOUPDATECP  ' ������� ������� �� �������������� ����� ������� ������ ������ ������.
'    TA_UPDATECP    ' ������� ������� �������������� ����� ������� ������ ������ ������.
'    TA_MASK  = (TA_BASELINE + TA_CENTER + TA_UPDATECP)
End Enum
#ElseIf ObjectDataType = 1 Then 'LV
Public Enum eConstants          'See SourceIconSizes
    HIGH_COLOR = &HFFFF00
    TRUE_COLOR = &HFF000000
    TRUE_COLOR_ALPHA = &HFFFFFFFF
End Enum
#End If                         'ObjectDataType
Public Enum StdPictureObjectType
    vbPicTypeNone = 0           'None (empty)
    vbPicTypeBitmap = 1         'Bitmap type of StdPicture object
    vbPicTypeMetafile = 2       'Metafile type of StdPicture object
    vbPicTypeIcon = 3           'Icon type of StdPicture object
    vbPicTypeEMetafile = 4      'Enhanced metafile type of StdPicture object
End Enum
Private Enum OleObjType
    OT_LINK = 1                 'The OLE item is a link.
    OT_EMBEDDED = 2             'The OLE item is embedded.
    OT_STATIC = 3               'The OLE item is static, that is, it contains only
                                'presentation data, not native data, and thus cannot be edited.
End Enum
Private Enum PICTYPE            'StdPicture object type
    PICTYPE_UNINITIALIZED = -1
    PICTYPE_NONE = 0            'None (empty)
    PICTYPE_BITMAP = 1          'Bitmap
    PICTYPE_METAFILE = 2        'Metafile
    PICTYPE_ICON = 3            'Icon
    PICTYPE_ENHMETAFILE = 4     'Enhanced metafile
End Enum
Private Enum CLIPFORMAT         'Predefined Clipboard Formats
    CF_TEXT = 1                 'Text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data. Use this format for ANSI text.
    CF_BITMAP = 2               'A handle to a bitmap (HBITMAP).
    CF_METAFILEPICT = 3         'Handle to a metafile picture format as defined by the METAFILEPICT structure. When passing a CF_METAFILEPICT handle by means of DDE, the application responsible for deleting hMem should also free the metafile referred to by the CF_METAFILEPICT handle.
    CF_SYLK = 4                 'Microsoft Symbolic Link (SYLK) format.
    CF_DIF = 5                  'Software Arts Data Interchange Format.
    CF_TIFF = 6                 'Tagged-image file format.
    CF_OEMTEXT = 7              'Text format containing characters in the OEM character set. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data.
    CF_DIB = 8                  'A memory object containing a BITMAPINFO structure followed by the bitmap bits.
    CF_PALETTE = 9              'Handle to a color palette. Whenever an application places data in the clipboard that depends on or assumes a color palette, it should place the palette on the clipboard as well.
                                'If the clipboard contains data in the CF_PALETTE (logical color palette) format, the application should use the SelectPalette and RealizePalette functions to realize (compare) any other data in the clipboard against that logical palette.
                                'When displaying clipboard data, the clipboard always uses as its current palette any object on the clipboard that is in the CF_PALETTE format.
    CF_PENDATA = 10             'Data for the pen extensions to the Microsoft Windows for Pen Computing.
    CF_RIFF = 11                'Represents audio data more complex than can be represented in a CF_WAVE standard wave format.
    CF_WAVE = 12                'Represents audio data in one of the standard wave formats, such as 11 kHz or 22 kHz PCM.
    CF_UNICODETEXT = 13         'Unicode text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data.
    CF_ENHMETAFILE = 14         'A handle to an enhanced metafile (HENHMETAFILE).
    CF_HDROP = 15               'A handle to type HDROP that identifies a list of files. An application can retrieve information about the files by passing the handle to the DragQueryFile function.
    CF_LOCALE = 16              'The data is a handle (HGLOBAL) to the locale identifier (LCID) associated with text in the clipboard. When you close the clipboard, if it contains CF_TEXT data but no CF_LOCALE data, the system automatically sets the CF_LOCALE format to the current input language. You can use the CF_LOCALE format to associate a different locale with the clipboard text.
                                'An application that pastes text from the clipboard can retrieve this format to determine which character set was used to generate the text.
                                'Note that the clipboard does not support plain text in multiple character sets. To achieve this, use a formatted text data type such as RTF instead.
                                'The system uses the code page associated with CF_LOCALE to implicitly convert from CF_TEXT to CF_UNICODETEXT. Therefore, the correct code page table is used for the conversion.
    CF_DIBV5 = 17               'A memory object containing a BITMAPV5HEADER structure followed by the bitmap color space information and the bitmap bits.
    CF_MAX = 17
    CF_OWNERDISPLAY = &H80      'Owner-display format. The clipboard owner must display and update the clipboard viewer window, and receive the WM_ASKCBFORMATNAME, WM_HSCROLLCLIPBOARD, WM_PAINTCLIPBOARD, WM_SIZECLIPBOARD, and WM_VSCROLLCLIPBOARD messages. The hMem parameter must be NULL.
    CF_DSPTEXT = &H81           'Text display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in text format in lieu of the privately formatted data.
    CF_DSPBITMAP = &H82         'Bitmap display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in bitmap format in lieu of the privately formatted data.
    CF_DSPMETAFILEPICT = &H83   'Metafile-picture display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in metafile-picture format in lieu of the privately formatted data.
    CF_DSPENHMETAFILE = &H8E    'Enhanced metafile display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in enhanced metafile format in lieu of the privately formatted data.
'Private formats don't get GlobalFree()
    CF_PRIVATEFIRST = &H200     'Start of a range of integer values for private clipboard formats. The range ends with CF_PRIVATELAST. Handles associated with private clipboard formats are not freed automatically; the clipboard owner must free such handles, typically in response to the WM_DESTROYCLIPBOARD message.
    CF_PRIVATELAST = &H2FF      'See CF_PRIVATEFIRST.
'GDIOBJ formats do get DeleteObject()
    CF_GDIOBJFIRST = &H300      'Start of a range of integer values for application-defined GDI object clipboard formats. The end of the range is CF_GDIOBJLAST.
                                'Handles associated with clipboard formats in this range are not automatically deleted using the GlobalFree function when the clipboard is emptied. Also, when using values in this range, the hMem parameter is not a handle to a GDI object, but is a handle allocated by the GlobalAlloc function with the GMEM_MOVEABLE flag.
    CF_GDIOBJLAST = &H3FF       'See CF_GDIOBJFIRST.
'Registered formats
    CF_RegisteredFIRST = &HC000&
    CF_RegisteredLAST = &HFFFF&
End Enum
'This Enum is needed to set the "Mapping" property for EMF images
Public Enum MMETRIC
    MM_HIMETRIC = 3
    MM_LOMETRIC = 2
    MM_LOENGLISH = 4
    MM_ISOTROPIC = 7
    MM_HIENGLISH = 5
    MM_ANISOTROPIC = 8
    MM_ADLIB = 9
End Enum
Public Enum DeviceCapIndex
    HORZSIZE = 4           '  Horizontal size in millimeters
    VERTSIZE = 6           '  Vertical size in millimeters
    HORZRES = 8            '  Horizontal width in pixels
    VERTRES = 10           '  Vertical width in pixels
    LOGPIXELSX = 88        '  Logical pixels/inch in X
    LOGPIXELSY = 90        '  Logical pixels/inch in Y
End Enum
'Private Enum EnumConversionScale
'    cvscPixels
'    cvscPoints
'    cvscTwips
'End Enum
Private Enum eDirection
    DIRECTION_HORIZONTAL = 0
    DIRECTION_VERTICAL = 1
End Enum
'
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

' ��������� ��� ����������
Private Const SBS_HORZ = &H0&
Private Const SBS_VERT = &H1&
Private Const SBS_SIZEBOX = &H8&
Private Const SB_CTL = 2
Private Const SB_THUMBPOSITION = 4
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115

' === Declare Types ===
'--------------------------------------------------------------------------------
' SAFEARRAY
'--------------------------------------------------------------------------------
Private Const FADF_AUTO As Long = (&H1)
Private Const FADF_FIXEDSIZE As Long = (&H10)
Private Type SAFEARRAYBOUND         ' 8 bytes
    cElements As Long               ' +0 ���������� ��������� � �����������
    lLbound As Long                 ' +4 ������ ������� �����������
End Type
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Private Type SAFEARRAY
    cDims           As Integer      ' +0  ����� ������������
    fFeatures       As Integer      ' +2  ����, ������������ ��������� SafeArray
    cbElements      As Long         ' +4  ������ ������ �������� � ������
    cLocks          As LongLong     ' +8  C������ ������, ����������� ���������� ����������, ���������� �� ������.
    pvData          As LongPtr      ' +16(x86) ��������� �� ������
    rgSAbound As SAFEARRAYBOUND     ' ����������� ��� ������ ����������� (������ = n*8 bytes, n- ���-�� ������������ �������)
                                    ' +24 rgSAbound.cElements (Long) - ���������� ��������� � �����������
                                    ' +28 rgSAbound.lLbound (Long)   - ������ ������� �����������
End Type
Private Type SAFEARRAY2D
    cDims           As Integer
    fFeatures       As Integer
    cbElements      As Long
    cLocks          As LongLong
    pvData          As LongPtr
    rgSAbound(0 To 1) As SAFEARRAYBOUND
End Type
#Else                   '<WIN32>
Private Type SAFEARRAY
    cDims           As Integer      ' +0  ����� ������������
    fFeatures       As Integer      ' +2  ����, ������������ ��������� SafeArray
    cbElements      As Long         ' +4  ������ ������ �������� � ������
    cLocks          As Long         ' +8  C������ ������, ����������� ���������� ����������, ���������� �� ������.
    pvData          As Long         ' +12 ��������� �� ������
    rgSAbound As SAFEARRAYBOUND     ' ����������� ��� ������ ����������� (������ = n*8 bytes, n- ���-�� ������������ �������)
                                    ' +16 rgSAbound.cElements (Long) - ���������� ��������� � �����������
                                    ' +20 rgSAbound.lLbound (Long)   - ������ ������� �����������
End Type
Private Type SAFEARRAY2D
    cDims           As Integer
    fFeatures       As Integer
    cbElements      As Long
    cLocks          As Long
    pvData          As Long
    rgSAbound(0 To 1) As SAFEARRAYBOUND
End Type
#End If                 '<WIN32>
Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As LongPtr
    hbmColor As LongPtr
End Type
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Private Type BITMAP             '28 bytes
    bmType As Long              '4
    bmWidth As Long             '4
    bmHeight As Long            '4
    bmWidthBytes As Long        '4
    bmPlanes As Integer         '2
    bmBitsPixel As Integer      '2
    bmBits As LongPtr           '8 bytes
End Type
Private Type PICTDESC
    Size      As Long
    Type      As PICTYPE
    hImage    As LongPtr
    Reserved1 As Long
    Reserved2 As Long
End Type
Private Type SHFILEINFO
    hIcon As LongPtr                    '  out: icon
    iIcon As Long                       '  out: icon index
    dwAttributes As Long                '  out: SFGAO_ flags
    szDisplayName As String * 260       '  out: display name (or path)
    szTypeName As String * 80           '  out: type name
End Type
Private Type METAFILEPICT
    mm As Long
    xExt As Long
    yExt As Long
    Hmf As LongPtr
End Type
#Else                   '<WIN32>
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type PICTDESC
    Size      As Long
    Type      As PICTYPE
    hImage    As Long
    Reserved1 As Long
    Reserved2 As Long
End Type
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type
Private Type METAFILEPICT
    mm As Long
    xExt As Long
    yExt As Long
    Hmf As Long
End Type
#End If                 '<WIN32>

Private Const BITMAPFILEHEADERSIZE = &HE&
Private Type BITMAPFILEHEADER   ' BITMAPFILEHEADER � 14-������� ���������.
    bfType As Integer           ' WORD  ��������� &h4D42
    bfSize As Long              ' DWORD (little-endian)  ������ ����� � ������
    bfReserved1 As Integer      ' WORD    ��������������� � ������ ��������� ����
    bfReserved2 As Integer      ' WORD    ��������������� � ������ ��������� ����
    bfOffset As Long            ' DWORD (little-endian) ��������� ���������� ������ ������������ ������ ������ ��������� (� ������).
End Type
Private Const BITMAPINFOHEADERSIZE = &H28&
Private Type BITMAPINFOHEADER   ' BITMAPINFOHEADER (v.3) 40 bytes
    biSize As Long              ' DWORD   ������ ������ ��������� � ������, ����������� ����� �� ������ ��������� (����� ������ ���� �������� 12).
    biWidth As Long             ' LONG    ������ ������ � ��������. ����������� ����� ������ ��� �����.
    biHeight As Long            ' LONG    ������ ������ � ��������.
    biPlanes As Integer         ' WORD    1 - � BMP ��������� ������ �������� 1
    biBitCount As Integer       ' WORD    ���������� ��� �� �������
    biCompression As Long       ' DWORD   0 - ��������� �� ������ �������� ��������
    biSizeImage As Long         ' DWORD   0 - ������ ���������� ������ � ������. ����� ���� �������� ���� �������� �������������� ��������� ��������.
    biXPelsPerMeter As Long     ' LONG    0 - ���������� �������� �� ���� �� �����������
    biYPelsPerMeter As Long     ' LONG    0 - ���������� �������� �� ���� �� ���������
    biClrUsed As Long           ' DWORD   0 - ������ ������� ������ � �������.
    biClrImportant As Long      ' DWORD   0 - ���������� ����� �� ������ ������� ������ �� ��������� ������������ (������� � ����).Stop
End Type
#If ObjectDataType = 0 Then     ' FI
#ElseIf ObjectDataType = 1 Then ' LV
Private Type RGBQUAD '
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type
Private Type RGBTRIPLE
   rgbtBlue As Byte
   rgbtGreen As Byte
   rgbtRed As Byte
End Type
#End If                          ' ObjectDataType


Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiBacks(0 To 255) As RGBQUAD
End Type
Private Type DIBSECTION
    dsBm As BITMAP
    dsBmih As BITMAPINFOHEADER
    dsBitFields(0 To 2) As Long
    dshSection As Long
    dsOffset As Long
End Type
Private Type POINTINT
    x As Integer
    y As Integer
End Type
Private Type OLEOBJECTHEADER ' information about object
    Signature As Integer     ' ��������� &h1C15
    HeaderSize As Integer    ' ������ ��������� (SizeOf(Struct OLEOBJECTHEADER)+cchName+cchClass).
    ObjectType As Long 'OleObjType ' ��� ���� OLE Object (OT_STATIC,OT_LINKED, OT_EMBEDDED).
    FriendlyNameLen As Integer  ' ���������� �������� � OLE Object Name (CchSz(szName) + 1).
    ClassNameLen As Integer     ' ���������� �������� � �lass Name (CchSz(szClass) + 1).
    FriendlyNameOffset As Integer    ' Offset of object name in structure (sizeof(OleObjectHeader)).
    ClassNameOffset As Integer   ' Offset of class name in structure (ibName + cchName).
    ObjectSize As POINTINT     ' Original size of Object (MM_HIMETRIC)
'    FriendlyName As Byte()   ' ��� �������
'    ClassName  As Byte()     ' ����� �������
End Type
Private Type OLEHEADER
    OleVersion As Long
    Format As CLIPFORMAT 'Long
    ObjectTypeNameLen As Long
'    ObjectTypeName As Byte()
End Type

' METAFILES
Private Const MEMORYMETAFILE As Integer = &H1
Private Const METAVERSION300  As Integer = &H300 '(DIBs are supported) defines the metafile version

Private Const META_EOF As Integer = &H0
Private Const META_SETMAPMODE As Integer = &H103
Private Const META_SETWINDOWORG As Integer = &H20B
Private Const META_SETWINDOWEXT As Integer = &H20C
Private Const META_DIBSTRETCHBLT As Integer = &HB41

'Private Type ENHMETAHEADER
'    iType As Long
'    nSize As Long
'    rclBounds As RECTL
'    rclFrame As RECTL
'    dSignature As Long
'    nVersion As Long
'    nBytes As Long
'    nRecords As Long
'    nHandles As Integer
'    sReserved As Integer
'    nDescription As Long
'    offDescription As Long
'    nPalEntries As Long
'    szlDevice As SIZEL
'    szlMillimeters As SIZEL   '#if(WINVER >= 0x0400)
'    cbPixelFormat As Long
'    offPixelFormat As Long
'    bOpenGL As Long            '#endif /* WINVER >= 0x0400 */'#if(WINVER >= 0x0500)
'    szMicrometers As SIZEL     '#endif /* WINVER >= 0x0500 */
'End Type
'Private Type ENHMETARECORD
'    iType As Long
'    nSize As Long
'    dParm(1) As Long
'End Type
'Private Type METAFILEPICT
'    mm As Long
'    xExt As Long
'    yExt As Long
'    hMF As LongPtr
'End Type
Private Type METAHEADER
    MetaType        As Integer
    HeaderSize      As Integer
    Version         As Integer
    SizeLow         As Integer
    SizeHigh        As Integer
    NumberOfObjects As Integer
    MaxRecord       As Long
    NumberOfMembers As Integer
End Type
'Private Type METAHEADER
'    mtType As Integer
'    mtHeaderSize As Integer
'    mtVersion As Integer
'    mtSize As Long
'    mtNoObjects As Integer
'    mtMaxRecord As Long
'    mtNoParameters As Integer
'End Type
Private Type METARECORD
    RecordSize      As Long
    RecordFunction  As Integer
    'Data
End Type
'Private Type METARECORD
'    rdSize As Long
'    rdFunction As Integer
'    rdParm(1) As Integer
'End Type
'Private Type METASETMAPMODE
'    RecordSize      As Long
'    RecordFunction  As Integer
'    MapMode         As Integer
'End Type
'Private Type METASETWINDOW
'    RecordSize      As Long
'    RecordFunction  As Integer
'    y               As Integer
'    x               As Integer
'End Type
'Private Type METADIBSTRETCHBLT
'    RecordSize      As Long
'    RecordFunction  As Integer
'    RasterOperation As Long '
'    SrcHeight       As Integer
'    SrcWidth        As Integer
'    YSrc            As Integer
'    XSrc            As Integer
'    DestHeight      As Integer
'    DestWidth       As Integer
'    YDest           As Integer
'    XDest           As Integer
'    'Target          As LongPtr
'End Type

' FONTS
Private Const LF_FACESIZE = 32
Private Const LF_FACESIZEW As Long = LF_FACESIZE * 2

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
Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
'Private Type PictDesc
'    Size As Long
'    Type As Long
'    hHandle As Long
'    hPal As Long
'End Type
'
' === Declare Functions ===
'--------------------------------------------------------------------------------
' MSVBA
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Private Declare PtrSafe Function VarPtrArray Lib "VBE7.dll" Alias "VarPtr" (ByRef Ptr() As Any) As LongPtr
#ElseIf VBA7 Then       '<WIN32 & OFFICE2010+>
Private Declare Function VarPtrArray Lib "VBE7.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
'#Else                   '<OFFICE2003-2010>
'Private Declare Function VarPtrArray Lib "VBE6.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
#Else                   '<OFFICE2000-2003>
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
'#Else                   '<OFFICE97-2000>
'Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
#End If                 '<WIN32>

'--------------------------------------------------------------------------------
' KERNEL32
'--------------------------------------------------------------------------------
#If VBA7 Then           '<OFFICE2010+>
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal bFill As Byte)
Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
' Used for workaround of VB not exposing IStream interface
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
' used to see if DLL exported function exists
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
' Unicode File operations
Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare PtrSafe Function DeleteFileW Lib "kernel32" (ByVal lpFileName As LongPtr) As Long
Private Declare PtrSafe Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare PtrSafe Function SetFileAttributesW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwFileAttributes As Long) As Long
Private Declare PtrSafe Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare PtrSafe Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As LongPtr) As Long
#Else                   '<WIN32>
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal bFill As Byte)
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
' Used for workaround of VB not exposing IStream interface
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
' used to see if DLL exported function exists
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32" (ByVal lpFileName As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function SetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
' USER32
'--------------------------------------------------------------------------------
Public Type RECT   ' Store rectangle coordinates.
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type POINT    ' aka Point
    x As Long
    y As Long
End Type
Public Type SIZEL    ' aka Size
    cX As Long
    cY As Long
End Type
Public Type POINTF   ' aka PointF
    x As Single
    y As Single
End Type
Public Type RECTF    ' aka RectF
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type
Public Type SIZEF    ' aka SizeF
    cX As Single
    cY As Single
End Type

#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Private Declare PtrSafe Function GetWindow Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function ClientToScreen Lib "user32.dll" (ByVal hwnd As LongPtr, ByRef lpPoint As POINT) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long 'Ptr
Private Declare PtrSafe Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare PtrSafe Function IsWindowUnicode Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function DestroyIcon Lib "user32" (ByVal hIcon As LongPtr) As Long
#Else                   '<OFFICE97-2010>
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
' GDI32: Fonts/Text functions
'--------------------------------------------------------------------------------
Private Const SYSTEM_FONT = 13        ' ��������� ����� (������������ ��� ����������� �������� � Windows)
#If VBA7 Then           '<OFFICE2010+>
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32.dll" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetObject Lib "gdi32.dll" Alias "GetObjectW" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
Private Declare PtrSafe Function SetWindowExtEx Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpSize As Any) As Long
Private Declare PtrSafe Function SetViewportExtEx Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal nX As Long, ByVal nY As Long, lpSize As Any) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpvBits As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
'Private Declare PtrSafe Function SelectBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal hBitmap As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateDIBitmap Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal lpInfoHeader As LongPtr, ByVal dwUsage As Long, lpInitBits As Any, ByVal lpInitInfo As LongPtr, ByVal wUsage As Long) As LongPtr
Private Declare PtrSafe Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal pbmi As LongPtr, ByVal iUsage As Long, ByRef ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal dwOffset As Long) As LongPtr
Private Declare PtrSafe Function GetBitmapBits Lib "gdi32.dll" (ByVal hBitmap As LongPtr, ByVal cb As Long, ByVal lpBits As LongPtr) As Long
Private Declare PtrSafe Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, ByVal lpBits As LongPtr, ByVal lpBI As LongPtr, ByVal wUsage As Long) As Long
Private Declare PtrSafe Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
#If ObjectDataType = 1 Then ' LV
Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As IUnknown) As Long
Private Declare PtrSafe Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As Any, RefIID As Any, ByVal fPictureOwnsHandle As Long, iPic As stdole.IPictureDisp) As Long
Private Declare PtrSafe Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare PtrSafe Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
#End If                     ' ObjectDataType
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
Private Declare PtrSafe Function GetTextMetrics Lib "gdi32.dll" Alias "GetTextMetricsW" (ByVal hdc As LongPtr, lpMetrics As TEXTMETRIC) As Long
Private Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32W" (ByVal hdc As LongPtr, ByVal lpsz As LongPtr, ByVal cbString As Long, ByRef lpSize As POINT) As Long
#Else                   '<WIN32>
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectW" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function SetWindowExtEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpvBits As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function SelectBitmap Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long) As Long
Private Declare Function CreateDIBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal lpInfoHeader As Long, ByVal dwUsage As Long, lpInitBits As Any, ByVal lpInitInfo As Long, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, ByVal pbmi As Long, ByVal iUsage As Long, ByRef ppvBits As Long, ByVal hSection As Long, ByVal dwOffset As Long) As Long
Private Declare Function GetBitmapBits Lib "gdi32.dll" (ByVal hBitmap As Long, ByVal cb As Long, ByVal lpBits As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByVal lpBits As Long, ByVal lpBI As Long, ByVal wUsage As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
#If ObjectDataType = 1 Then         ' LV
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As IUnknown) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As Any, RefIID As Any, ByVal fPictureOwnsHandle As Long, iPic As stdole.IPictureDisp) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
#End If                 ' ObjectDataType
' used to create the checkerboard pattern on demand
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
' used to create font
Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" (ByRef lpLogFont As LOGFONT) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
' used to text output
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutW" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32.dll" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32.dll" Alias "GetTextMetricsW" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32W" (ByVal hdc As Long, ByVal lpsz As Long, ByVal cbString As Long, ByRef lpSize As POINT) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
' OLE32
'--------------------------------------------------------------------------------
#If VBA7 Then           ' <OFFICE2010+>         use: PtrSafe with LongPtr only (LongLong not avaible)
Private Declare PtrSafe Function OleLoadPicture Lib "olepro32.dll" (ByVal pStream As IUnknown, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvPic As IPictureDisp) As Long
Private Declare PtrSafe Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As OLE_COLOR, ByVal hPal As Long, ByRef lpcolorref As Long) As Long
#Else                   ' <OFFICE97-2007>       use: Long only
Private Declare Function OleLoadPicture Lib "olepro32.dll" (ByVal pStream As IUnknown, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvPic As IPictureDisp) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As OLE_COLOR, ByVal hPal As Long, ByRef lpcolorref As Long) As Long
#End If                 ' <VBA7 & WIN64>
'--------------------------------------------------------------------------------
' COMDLG32
'--------------------------------------------------------------------------------
#If VBA7 Then           ' <OFFICE2010+>
Private Type ChooseColor
    lStructSize               As Long
    hwndOwner                 As LongPtr
    hInstance                 As LongPtr
    rgbResult                 As Long
    lpCustColors              As LongPtr
    Flags                     As Long
    lCustData                 As LongPtr
    lpfnHook                  As LongPtr
    lpTemplateName            As String
End Type
Private Type ChooseFont
    lStructSize As Long
    hwnd As LongPtr
    hdc As LongPtr
    lpLogFont As LongPtr
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
    hInstance As LongPtr
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type
#Else                   ' <OFFICE97-2007>
Private Type ChooseColor
    lStructSize               As Long
    hwndOwner                 As Long
    hInstance                 As Long
    rgbResult                 As Long
    lpCustColors              As Long
    Flags                     As Long
    lCustData                 As Long
    lpfnHook                  As Long
    lpTemplateName            As String
End Type
Private Type ChooseFont
    lStructSize As Long
    hwnd As Long
    hdc As Long
    lpLogFont As Long
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type
#End If                 ' <VBA7 & WIN64>
 
Private Const CC_ANYCOLOR = &H100
'Private Const CC_ENABLEHOOK = &H10
'Private Const CC_ENABLETEMPLATE = &H20
'Private Const CC_ENABLETEMPLATEHANDLE = &H40
Private Const CC_FULLOPEN = &H2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_RGBINIT = &H1
'Private Const CC_SHOWHELP = &H8
'Private Const CC_SOLIDCOLOR = &H80
 
Private Const CF_APPLY = &H200&
Private Const CF_ANSIONLY = &H400&
Private Const CF_TTONLY = &H40000
Private Const CF_EFFECTS = &H100&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLETEMPLATEHANDLE = &H20&
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_NOSTYLESEL = &H100000
Private Const CF_NOSIZESEL = &H200000
Private Const CF_NOSIMULATIONS = &H1000&
Private Const CF_NOVECTORFONTS = &H800&
Private Const CF_NOVERTFONTS = &H1000000
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_SCREENFONTS = &H1
Private Const CF_SCRIPTSONLY = CF_ANSIONLY
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_SHOWHELP = &H4&
Private Const CF_USESTYLE = &H80&
Private Const CF_WYSIWYG = &H8000
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS
 
#If VBA7 Then           ' <OFFICE2010+>
Private Declare PtrSafe Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare PtrSafe Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
#Else                   ' <OFFICE97-2007>
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChooseFont As ChooseFont) As Long
#End If                 ' <VBA7 & WIN64>
'--------------------------------------------------------------------------------
' Access internal functions
'--------------------------------------------------------------------------------
Public Const c_strAppDibIdPref = "#"  ' ������� ��� ����� DIB ������ ���������� � Access
Public Enum tbdibPictures
' Id ���� ���������� ������
    tbdibOpen = 23      ' �������
    tbdibSave = 3       ' ���������
    tbdibPrint = 4      ' ������            ' StrID=5570747
    tbdibCopy = 19      ' ����������        ' StrID=5570800
    tbdibCut = 21       ' ��������          ' StrID=5570799
    tbdibPaste = 22     ' ��������          ' StrID=5570801
    tbdibMail = 24      ' ������
    tbdibMagnify = 25   ' �������������� ������
    tbdibSearch = 46    ' �������           ' StrID=5570790 ' �����
    tbdibClear = 47     ' ������
    tbdibHelp = 49      ' ������ ������
    tbdibCalc = 50      ' �����������
    tbdibRecycle = 67   ' �������
    tbdibView = 109     ' �������� ���������
    tbdibProperties = 222 ' ��������
    tbdibDelete = 358   ' ������� X
    tbdibFunc = 385     ' ������� f(x)
    tbdibColor = 417    ' �������
    tbdibZoomUp = 444   ' ���������
    tbdibZoomDown = 445 ' ���������
    tbdibTools = 548    ' ���������
    tbdibFilter = 601   ' ������
    tbdibUser = 607     ' ������������
    tbdibCalendar = 612 ' ������������
    tbdibFont = 253 '2823 '6505    ' �����
    tbdibFontColor = 2611   ' ���� ������
    tbdibAlignLeft = 120    ' ��������� ����� �����
    tbdibAlignCenter = 122  ' ��������� ����� �� ������
    tbdibAlignRight = 121   ' ��������� ����� ������
End Enum
#If VBA7 Then           ' <OFFICE2010+>
' ' Private Declare PtrSafe Function accGetTbDIB Lib "msaccess.exe" Alias "MSAU_OfficeGetTcDIB@12" (ByVal lngBmp As Long, ByVal fLarge As Long, abytBuffer() As Byte) As Long' for Access 2003 and 2007
Private Declare PtrSafe Function accGetTbDIB Lib "msaccess.exe" Alias "#60" (ByVal lngBmp As Long, ByVal fLarge As Long, abytBuffer() As Byte) As Long
Private Declare PtrSafe Sub accChooseColor Lib "msaccess.exe" Alias "#53" (ByVal hwnd As LongPtr, rgb As Long)
#Else                   ' <OFFICE97-2007>
' ' Private Declare PtrSafe Function accGetTbDIB Lib "msaccess.exe" Alias "MSAU_OfficeGetTcDIB@12" (ByVal lngBmp As Long, ByVal fLarge As Long, abytBuffer() As Byte) As Long' for Access 2003 and 2007
Private Declare Function accGetTbDIB Lib "msaccess.exe" Alias "#60" (ByVal lngBmp As Long, ByVal fLarge As Long, abytBuffer() As Byte) As Long
Private Declare Sub accChooseColor Lib "msaccess.exe" Alias "#53" (ByVal hwnd As Long, rgb As Long)
#End If                 ' <VBA7 & WIN64>
'' AppLoadString(TextStringId as Long) as String
'--------------------------------------------------------------------------------
' Module types/vars/consts/enums
'--------------------------------------------------------------------------------
Public Const eContainer = &H10000   ' ������� ����������
Public Enum eControlType            ' ���� ���������, �������������� PictureData_SetToControl
    eCtrlTypeNone = &H0&            '
    eCtrlTypeUndef = &HFFFF&        ' ��� �����������
    eCtrlAccEmf = acImage&          ' ������� Access �� ��������� PictureData ������� EMF (Image,Page,Form,Report)
    eCtrlAccDib = acCommandButton&  ' ������� Access �� ��������� PictureData ������� DIB (CommandButton,ToggleButton)
    eCtrlPicture = &H1F000          ' ������ �� ��������� Picture ���� StdPicture (BMP)
    eCtrlPicEmf = &H1F001           ' ������ �� ��������� Picture ���� StdPicture (EMF)
    eStdPicture = &HF000&            ' StdPicture (BMP)
    eCtrlSwf = &H1FFFE              ' ShockWave Flash
End Enum
#If ObjectDataType = 1 Then         ' LV
Public Enum LaVolpe_IMAGE_FORMAT
    LVF_UNKNOWN = -1
    LVF_BMP = 0
    LVF_ICO = 1
    LVF_JPEG = 2
    LVF_PNG = 13
    LVF_TIFF = 18
    LVF_GIF = 25
    LVF_J2K = 30
    LVF_JP2 = 31
End Enum
#End If                             ' ObjectDataType
Public Const eObjectPic = &H10000   ' �����������
Public Const eObjectOmf = &H20000   ' ������ ����������
Public Const eObjectDoc = &H30000   ' ���������
Public Const eObjectDbs = &H40000   ' ���� ������
Public Const eObjectArc = &H50000   ' ������
Public Const eObjectOth = &HF0000   ' ������

Public Enum eObjectDataType         ' ���� ��������� ������, �������������� PictureData_SetToControl
    eObjectDataNone = 0             '
' ����������� (��������)
#If ObjectDataType = 0 Then         ' FI
    eObjectDataUndef = FIF_UNKNOWN  ' ��� �����������
    eObjectDataBMP = eObjectPic + FIF_BMP   ' BMP, DIB - Windows (or device-independent) bitmap image
    eObjectDataPNG = eObjectPic + FIF_PNG   ' Portable Network Graphics file
    eObjectDataGIF = eObjectPic + FIF_GIF   ' Graphics interchange format file
    eObjectDataJPG = eObjectPic + FIF_JPEG  ' JPEG, JPG - JPEG/JFIF graphics file
    eObjectDataICO = eObjectPic + FIF_ICO   ' Windows icon file
    eObjectDataTIF = eObjectPic + FIF_TIFF  ' TIF, TIFF - Tagged Image File Format file
#ElseIf ObjectDataType = 1 Then     ' LV
    eObjectDataUndef = LVF_UNKNOWN  ' ��� �����������
    eObjectDataBMP = eObjectPic + LVF_BMP   ' BMP, DIB - Windows (or device-independent) bitmap image
    eObjectDataPNG = eObjectPic + LVF_PNG   ' Portable Network Graphics file
    eObjectDataGIF = eObjectPic + LVF_GIF   ' Graphics interchange format file
    eObjectDataJPG = eObjectPic + LVF_JPEG  ' JPEG, JPG - JPEG/JFIF graphics file
    eObjectDataICO = eObjectPic + LVF_ICO   ' Windows icon file
    eObjectDataTIF = eObjectPic + LVF_TIFF  ' TIF, TIFF - Tagged Image File Format file
#End If                 ' ObjectDataType
' ����������� (������)
    eObjectDataEMF = eObjectPic + 91        ' Extended (Enhanced) Windows Metafile Format
    eObjectDataWMF = eObjectPic + 92        ' Windows graphics metafile
    eObjectDataEPS = eObjectPic + 93        ' Adobe encapsulated PostScript file
    eObjectDataWPG = eObjectPic + 94        ' WordPerfect text and graphics file
    eObjectDataWIM = eObjectPic + 95        ' Microsoft Windows Imaging Format file
    eObjectDataCUR = eObjectPic + 96        ' Windows cursor file
' ������ ����������
    eObjectDataSWF = eObjectOmf + 201       ' Macromedia Shockwave Flash player file
' ���������
    eObjectDataRTF = eObjectDoc + 1         ' Rich Text Format
    eObjectDataDOC20 = eObjectDoc + 2       ' Word 2.0 file
    eObjectDataDOC = eObjectDoc + 3         ' Word 97-2003 file
    eObjectDataDOCX = eObjectDoc + 4        ' Word 2007+ file
    eObjectDataXLS = eObjectDoc + 5         ' Excel 97-2003 file
    eObjectDataXLSX = eObjectDoc + 6        ' Excel 2007+ file
    eObjectDataPPT = eObjectDoc + 7         ' PowerPoint 97-2003 file
    eObjectDataPPTX = eObjectDoc + 8        ' PowerPoint 2007+ file
    eObjectDataPS = eObjectDoc + 9          ' PostScript �����
    eObjectDataPDF = eObjectDoc + 10        ' PDF, FDF, AI  Adobe Portable Document Format, Forms Document Format, and Illustrator graphics files
    eObjectDataDJV = eObjectDoc + 11        ' DJV, DjVu �����
' ���� ������
    eObjectDataMDB = eObjectDbs + 1         ' Standard Jet db MDB Microsoft Access file
    eObjectDataACCDB = eObjectDbs + 2       ' Standard ACE db ACCDB Microsoft Access 2007+ file
' ������
    eObjectDataZIP = eObjectArc + 1         ' ZIP
    eObjectDataRAR = eObjectArc + 2         ' RAR
    eObjectData7Z = eObjectArc + 3          ' 7zip
' ������
    eObjectDataXML = eObjectOth + 1         ' XML,XUL - User Interface Language file
End Enum
'Public Enum ePosition               ' ��������� ��������� �� ������� (��� ��������������� ������������ � ������������)
'    ePosUndef = 0                   ' �� ������
'    eLeft = 1                       ' �� ������ ����
'    eRight = 2                      ' �� ������� ����
'    eTop = 4                        ' �� �������� ����
'    eBottom = 8                     ' �� ������� ����
'    eCenterHorz = eLeft + eRight    ' ����� �� �����������
'    eCenterVert = eTop + eBottom    ' ����� �� ���������
'    eCascade = 256                  ' ���������� (������ ��� ����� ??)
'End Enum
'Public Enum eAlign                  ' ������������ ������ �������
'    eAlignUndef = 0                         ' �� ������
'    ' 2 ����������� �� 3 ��������� ����� �������
'    ' �����: 3x3 = 9 ����� ������������.
'    eAlignLeftTop = eLeft + eTop                ' �� ������ �������� ����
'    eAlignRightTop = eRight + eTop              ' �� ������� �������� ����
'    eAlignLeftBottom = eLeft + eBottom          ' �� ������ ������� ����
'    eAlignRightBottom = eRight + eBottom        ' �� ������� ������� ����
'    eCenterHorzTop = eCenterHorz + eTop         ' �� �������� ���� ������������ �� �����������
'    eCenterHorzBottom = eCenterHorz + eBottom   ' �� ������� ���� ������������ �� �����������
'    eCenterVertLeft = eLeft + eCenterVert       ' �� ������ ���� ������������ �� ���������
'    eCenterVertRight = eRight + eCenterVert     ' �� ������� ���� ������������ �� ���������
'    eCenter = eCenterHorz + eCenterVert         ' ������������ ���������� �������
'End Enum
'Public Enum ePlace              ' ���������� Obj2 ������������ Obj1
'    ' 2 ������� �� 9 ����� �������� �� ������: LT,LC,LB,CB,RB,RC,RT,CT,CC
'    ' �����: 9x9 = 81 ������� ��������.
'    ' ����������� �� ��� ������������, ������� �������� ��� ��� ���,
'    ' �� �������� ����� ���������� �������� �� �����:
'    ' =H2+V2+H1+V1, ���:
'    ' Obj1 (� �������� �����������) - ���� 0-3:  L1=1,  R1=2,  T1=4,  B1=8
'    '   H1 - ��������� �� ����������� ����� �������� �� Obj1
'    '       ={eLeft|eRight|eCenterHorz}
'    '   V1 - ��������� �� ��������� ����� �������� �� Obj1
'    '       ={eTop|eBottom|eCenterVert}
'    ' Obj2 (������� �����������)    - ���� 4-8:  L2=16, R2=32, T2=64, B2=128
'    '   H2 - ��������� �� ����������� ����� �������� �� Obj2
'    '       ={eLeft|eRight|eCenterHorz} * 16
'    '   V2 - ��������� �� ��������� ����� �������� �� Obj2
'    '       ={eTop|eBottom|eCenterVert} * 16
'    ePlaceUndef = 0     ' ��-��������� = 222 -> ePlaceOnRight - ������� ������ �� ������
'' ������ �� ������
'    ePlaceCenter = eCenter * 16 + eCenter                           ' �� ������ (������)
'    ePlaceToLeft = eCenterVertLeft * 16 + eCenterVertLeft           ' ������ ����� �� ������
'    ePlaceToRight = eCenterVertRight * 16 + eCenterVertRight        ' ������ ������ �� ������
'    ePlaceToTop = eCenterHorzTop * 16 + eCenterHorzTop              ' ������ �� ������ ������
'    ePlaceToBottom = eCenterHorzBottom * 16 + eCenterHorzBottom     ' ������ �� ������ �����
'' ������� �� ������
'    ePlaceOnLeft = eCenterVertRight * 16 + eCenterVertLeft          ' ������� ����� �� ������
'    ePlaceOnRight = eCenterVertLeft * 16 + eCenterVertRight         ' ������� ������ �� ������
'    ePlaceOnTop = eCenterHorzBottom * 16 + eCenterHorzTop           ' ������� �� ������ ������
'    ePlaceOnBottom = eCenterHorzTop * 16 + eCenterHorzBottom        ' ������� �� ������ �����
'' ������ �� ����
'    ePlaceToLeftTop = eAlignLeftTop * 16 + eAlignLeftTop            ' ������ ����� ������
'    ePlaceToRightTop = eAlignRightTop * 16 + eAlignRightTop         ' ������ ������ ������
'    ePlaceToLeftBottom = eAlignLeftBottom * 16 + eAlignLeftBottom   ' ������ ����� �����
'    ePlaceToRightBottom = eAlignRightBottom * 16 + eAlignRightBottom ' ������ ������ �����
'' ������� �� ����
'    ePlaceOnLeftToTop = eAlignRightTop * 16 + eAlignLeftTop         ' ������� ����� � �������� ����
'    ePlaceOnLeftToBottom = eAlignRightBottom * 16 + eAlignLeftBottom ' ������� ����� � ������� ����
'    ePlaceOnRightToTop = eAlignLeftTop * 16 + eAlignRightTop        ' ������� ������ � �������� ����
'    ePlaceOnRightToBottom = eAlignLeftBottom * 16 + eAlignRightBottom ' ������� ������ � ������� ����
'    ePlaceOnTopToLeft = eAlignLeftBottom * 16 + eAlignLeftTop       ' ������� � ������ ���� ������
'    ePlaceOnTopToRight = eAlignRightBottom * 16 + eAlignRightTop    ' ������� � ������� ���� ������
'    ePlaceOnBottomToLeft = eAlignLeftTop * 16 + eAlignLeftBottom    ' ������� � ������ ���� �����
'    ePlaceOnBottomToRight = eAlignRightTop * 16 + eAlignRightBottom ' ������� � ������� ���� �����
'' ���������� (������ ��� ����� ??)
'    eCascadeFromLeftTop = eCascade + ePlaceToLeftTop                ' ���������� �������� ������-����
'    eCascadeFromRightTop = eCascade + ePlaceToRightTop              ' ���������� �������� �����-����
'    eCascadeFromLeftBottom = eCascade + ePlaceToLeftBottom          ' ���������� �������� ������-�����
'    eCascadeFromRightBottom = eCascade + ePlaceToRightBottom        ' ���������� �������� �����-�����
'End Enum
'
'Public Enum eObjSizeMode                    ' ��������������� ��������
'    apObjSizeZoomDown = -1                  '-1 - ���������������� ��������������� (������ ����������)
'    apObjSizeClip = acOLESizeClip           ' 0 - �� ������ ������. ���� ������ ������ ������� ������ - �������
'    apObjSizeStretch = acOLESizeStretch     ' 1 - ������/���������� (�������� ���������)
'    'apObjSizeAutoSize = acOLESizeAutoSize   ' 2 - ???
'    apObjSizeZoom = acOLESizeZoom           ' 3 - ���������������� ���������������
'End Enum

'--------------------------------------------------------------------------------
' Module functions
'--------------------------------------------------------------------------------
#If ObjectDataType = 1 Then     'LV
Public Function ConvertColor(ByVal Color As Long) As Long
' This helper function converts a VB-style color value (like vbRed), which
' uses the ABGR format into a RGBQUAD compatible color value, using the ARGB
   ConvertColor = ((Color And &HFF000000) Or ((Color And &HFF&) * &H10000) Or ((Color And &HFF00&)) Or ((Color And &HFF0000) \ &H10000))
End Function
Public Function ConvertOleColor(ByVal Color As OLE_COLOR) As Long
' This helper function converts an OLE_COLOR value (like vbButtonFace), which uses the BGR format into a RGBQUAD compatible color value, using the ARGB format, needed by FreeImage.
' generally ingnores the specified color's alpha value but, in contrast to ConvertColor, also has support for system colors, which have the format &H80bbggrr.
' in ARGB format into a VB-style ABGR color value. Use function ConvertColor instead.
Dim lColorRef As Long: If (OleTranslateColor(Color, 0, lColorRef) = 0) Then ConvertOleColor = ConvertColor(lColorRef)
End Function
#End If                         'ObjectDataType
#If ObjectDataType = 0 Then     'FI
Public Function PictureData_LoadFromEx(ObjectData As Variant, _
    Optional ByVal Width As Long, Optional ByVal Height As Long, _
    Optional ObjectType As eObjectDataType) As LongPtr
If Not FreeImage_IsLoaded Then FreeImage_LoadLibrary
Dim fiPict As LongPtr, lIF As FREE_IMAGE_FORMAT
#ElseIf ObjectDataType = 1 Then 'LV
Public Function PictureData_LoadFromEx(ObjectData As Variant, _
    Optional ByVal Width As Long, Optional ByVal Height As Long, _
    Optional ObjectType As eObjectDataType) As clsPictureData
Dim lvPict As New clsPictureData, lIF As LaVolpe_IMAGE_FORMAT
#End If                         'ObjectDataType
' ��������� � ������ ������ ����������� �� ObjectData
Const c_strProcedure = "PictureData_LoadFromEx"
' ObjectData - �������� ������ (�������� ������, ���/���) ������� (�����������) ��� ��������
' DestWidth/DestHeight  - ������� ������� �����������
'Static col As Collection ' ����������� ��������� ��� �������� ����������� �����������
'    On Error Resume Next
'    col.Count: If Err Then Set col = New Collection: Err.Clear
    On Error GoTo HandleError
' Load FreeImage Library to memory
' ��������� ���������� ������
Dim aObjData() As Byte, aTransparency() As Byte
Dim Ret
    '#If ObjectDataType = 0 Then     'FI
    'If Is???(ObjectData) Then   fiPict = ObjectData                        ' ������� fibitmap
    '#ElseIf ObjectDataType = 1 Then 'LV
    'If TypeOf ObjectData Is clsPictureData Then Set lvPict = bjectData     ' ������� clsPictureData
    '#End If                         'ObjectDataType
    If IsArray(ObjectData) Then
' ������� ������
        aObjData() = ObjectData
    #If ObjectDataType = 0 Then     'FI
        fiPict = FreeImage_LoadBitmapFromMemoryEx(aObjData(), Width:=Width, Height:=Height, Format:=lIF)
    #ElseIf ObjectDataType = 1 Then 'LV
        Call lvPict.LoadPicture_Stream(aObjData, ByVal Width, ByVal Height, SaveFormat:=True)
    #End If                         'ObjectDataType
    ElseIf TypeOf ObjectData Is AccessField Then
' �������� ���� ���������� ������
        aObjData() = ObjectData.Value
    #If ObjectDataType = 0 Then     'FI
        fiPict = FreeImage_LoadBitmapFromMemoryEx(aObjData(), Width:=Width, Height:=Height, Format:=lIF)
    #ElseIf ObjectDataType = 1 Then 'LV
        Call lvPict.LoadPicture_Stream(aObjData, ByVal Width, ByVal Height, SaveFormat:=True)
    #End If                         'ObjectDataType
    ElseIf Trim(ObjectData) = vbNullString Then
' �������� �� ������
        Err.Raise vbObjectError + 512
' � ObjectData ������ - ��� ������ ����������/��� �� ������� ��� ��� �����
    ElseIf ByteArray_ReadFromApp(ObjectData, aObjData, aTransparency) = NOERROR Then
' ������� ��� ���������� � ���������� DIB ��������
    #If ObjectDataType = 0 Then     'FI
        fiPict = FreeImage_CreateFromPictureData(aObjData()): lIF = FIF_BMP 'fiPict = FreeImage_LoadBitmapFromMemoryEx(aObjData(), width:=DestWidth, height:=DestHeight)
        Call FreeImage_SetTransparencyTableEx(fiPict, aTransparency) ' ������������� ������������ ��� ������ ������ � �������
    ' ����������� � 32 bit � Alpha
        Dim fiTemp As LongPtr: fiTemp = fiPict: fiPict = FreeImage_ConvertTo32Bits(fiTemp): FreeImage_Unload (fiTemp)
    #ElseIf ObjectDataType = 1 Then 'LV
        If p_ArrayDibToBmp(aObjData) Then _
        Call lvPict.LoadPicture_Stream(aObjData, ByVal Width, ByVal Height, SaveFormat:=True): lIF = LVF_BMP
    #End If                         'ObjectDataType
    ElseIf ByteArray_ReadFromTable(ObjectData, aObjData) = NOERROR Then
' �������� ������� ��� �������� �� ������� SysObjects
    #If ObjectDataType = 0 Then     'FI
        fiPict = FreeImage_LoadBitmapFromMemoryEx(aObjData(), Width:=Width, Height:=Height, Format:=lIF)
    #ElseIf ObjectDataType = 1 Then 'LV
        Call lvPict.LoadPicture_Stream(aObjData, ByVal Width, ByVal Height, SaveFormat:=True)
    #End If                         'ObjectDataType
    ElseIf ByteArray_ReadFromFile(ObjectData, aObjData) = NOERROR Then
' ������� ���� � �����
    #If ObjectDataType = 0 Then     'FI
    fiPict = FreeImage_LoadBitmapFromMemoryEx(aObjData(), Width:=Width, Height:=Height, Format:=lIF)
    #ElseIf ObjectDataType = 1 Then 'LV
        Call lvPict.LoadPicture_Stream(aObjData, ByVal Width, ByVal Height, SaveFormat:=True)
    #End If                         'ObjectDataType
'    ElseIf FreeImage_GetImageType(ObjectData) <> FIT_UNKNOWN Then
'    ' ��� ��������� ������������� fiPict �� ������� Access ??
    Else: Err.Raise vbObjectError + 512
    End If
    'lFIF = FreeImage_GetFIFFromMemory(aObjData): If lFIF = FIF_UNKNOWN Then Err.Raise vbObjectError + 512
    #If ObjectDataType = 0 Then     'FI
    If fiPict = 0 Then Err.Raise vbObjectError + 512
    #ElseIf ObjectDataType = 1 Then 'LV
    If lvPict.Handle = 0 Then Err.Raise vbObjectError + 512
    #End If                         'ObjectDataType
'' test
    ObjectType = lIF + eObjectPic
    #If ObjectDataType = 0 Then     'FI
HandleExit:  PictureData_LoadFromEx = fiPict: Exit Function
HandleError: fiPict = 0: Err.Clear: Resume HandleExit
    #ElseIf ObjectDataType = 1 Then 'LV
HandleExit:  Set PictureData_LoadFromEx = lvPict: Exit Function
HandleError: lvPict.DestroyDIB: Err.Clear: Resume HandleExit
    #End If                         'ObjectDataType
End Function
Public Function PictureData_SetIcon(ByRef FormHwnd As LongPtr, Optional ByRef ObjectData As Variant) As Long
' ������������� ������ ��� �����/������
Const c_strProcedure = "PictureData_SetIcon"
' FormHwnd   - hwnd ������� ��� �������� ��������������� ������
' ObjectData - �������� ������ (�������� ������, ���/���) ������� (�����������)������� ���� ���������� ��� ������
'              ���� �������� ������ - ������� ������
'-------------------------
Dim Result As Long: Result = NOERROR 'False 'NOERROR
If FormHwnd = 0 Then Err.Raise vbObjectError + 512
Dim hIcon As LongPtr
#If ObjectDataType = 0 Then     'FI
Dim fiPict As LongPtr: fiPict = PictureData_LoadFromEx(ObjectData, 16)
'#If VBA7 Then
'Debug.Print "Error: FreeImage_GetIcon creates empty hIcon on x64"
'#End If
    If (fiPict <> 0) Then hIcon = FreeImage_GetIcon(fiPict, UnloadSource:=True)
#ElseIf ObjectDataType = 1 Then 'LV
'Dim lvPict As clsPictureData
'    Set lvPict = PictureData_LoadFromEx(ObjectData, 16)
Debug.Print "Error: LaVolpe_GetIcon function not released"
Stop
#End If                         'ObjectDataType
' ������������� ������ �����
Dim bolErase As Boolean: bolErase = (hIcon = 0)
    hIcon = SendMessage(FormHwnd, WM_SETICON, ICON_SMALL, hIcon)
    If bolErase Then If hIcon <> 0 Then DestroyIcon hIcon
HandleExit:  PictureData_SetIcon = Result: Exit Function
HandleError: Result = Err: Err.Clear: Resume HandleExit
End Function
Public Function PictureData_GetControlType(ByRef ObjectControl As Object, _
    Optional Offsize As Long, Optional Offpos As Long, Optional BackColor As Long) As eControlType
' ���������� ��� � ���.�������� ��������
' ��� ����� �������� ����������� ������� ���� ��� ������� ���������������� ��������
' Offsize - (px) ���������� ������� ���������� ������� �������� �� �������� ��������
' Offpos  - (px) �������� ����� ������� �������� ��������
' BackColor - ������� ���� �������� (��� ������ �������� �������� - ������� ������� ����� ������������ ������������ ��������� ���������)
Dim Result As Long
    On Error GoTo HandleError
    Offsize = 0: Offpos = 0
    BackColor = 0 'Or &HFF000000 ': BackColor = ConvertColor(vbBlack)
    If (TypeOf ObjectControl Is Access.CommandButton) Or _
        (TypeOf ObjectControl Is Access.ToggleButton) Then
    ' ������� Access ������������ PictureData (DIB)
        Result = eCtrlAccDib ' And &HFFFF
        BackColor = ConvertOleColor(vbButtonFace) Or &HFF000000
    ' � MSO2003 � ����� - ��� ������� BorderStyle/BorderWidth, �� ���� ������� = 3pt
    ' ������� ������ ������� ��:
    ' ������� ������� ������ = 2px, ��������� ����������� ��������� ��������
                        Offpos = 2: Offsize = 2 * Offpos
    ' ������� ���������      = 1px, ������������� �� ��������. ��������� � ��� ��� ������ �� ���������� ������������
                        Offpos = Offpos + 1: Offsize = Offsize + 2
    ElseIf (TypeOf ObjectControl Is Access.Image) Or _
        (TypeOf ObjectControl Is Access.Form) Or _
        (TypeOf ObjectControl Is Access.Page) Or _
        (TypeOf ObjectControl Is Access.Report) Then
    ' ������� Access ������������ PictureData (EMF)
        Result = eCtrlAccEmf ' And &HFFFF
        On Error Resume Next
        With ObjectControl
            If .BorderStyle = 0 Then GoTo HandleExit
            ' ������� ������� �������� ������� �������� �� 1px
            Select Case .BorderWidth
            Case 0, 1:  Offpos = 1: Offsize = 1
            Case 2:     Offpos = 2: Offsize = 3
            Case 3:     Offpos = 2: Offsize = 4
            Case 4:     Offpos = 3: Offsize = 5
            Case 5:     Offpos = 3: Offsize = 6
            Case 6:     Offpos = 4: Offsize = 7
            Case Else:  Err.Raise vbObjectError + 512
            End Select
        End With
    ElseIf TypeOf ObjectControl Is StdPicture Then
    ' StdPicture (bmp)
        Result = eStdPicture ' And &HFFFF
        BackColor = ConvertOleColor(vbButtonFace) Or &HFF000000
    ElseIf TypeOf ObjectControl Is CommandBarButton Then
    ' StdPicture (bmp)
        Result = eCtrlPicture ' And &HFFFF
    Else '
        On Error Resume Next
Dim tmp As Object: Set tmp = ObjectControl.Picture: If Err Then Err.Raise vbObjectError + 512
    ' ActiveX with Property .Picture; TypeOf Picture Is StdPicture (bmp)'(emf)
        Result = eCtrlPicture 'eCtrlPicEmf '
        BackColor = ConvertOleColor(vbButtonFace) Or &HFF000000
        With ObjectControl
        ' ������� MSForms.Image
            Select Case .Object.BorderStyle
            Case 0:     Offpos = 0: Offsize = 0 'None
            Case 1:     Offpos = 1: Offsize = 2 'Single
            End Select
        ' ������� ���������� ActiveX (CustomControl)
            If .BorderStyle = 0 Then GoTo HandleExit
            Select Case .BorderWidth
            Case 0, 1:  Offpos = Offpos + 1: Offsize = Offsize + 1
            Case 2:     Offpos = Offpos + 2: Offsize = Offsize + 3
            Case 3:     Offpos = Offpos + 2: Offsize = Offsize + 4
            Case 4:     Offpos = Offpos + 3: Offsize = Offsize + 5
            Case 5:     Offpos = Offpos + 4: Offsize = Offsize + 7
            Case 6:     Offpos = Offpos + 4: Offsize = Offsize + 8
            Case Else:  Err.Raise vbObjectError + 512
            End Select
        End With
    End If
    On Error Resume Next
#If VBA7 Then
'       '.Top/Left/Right/BottomPadding ' �� ���� ��� ���
#End If
    Err.Clear
HandleExit:  PictureData_GetControlType = Result: Exit Function
HandleError: Result = eCtrlTypeUndef: Err.Clear: Resume HandleExit
End Function
Public Function PictureData_SetToControl( _
    ByRef ObjectControl As Object, Optional ByRef ObjectData As Variant, _
    Optional Alignment As eAlign = eCenter, _
    Optional Description As String, Optional Comment As String, _
    Optional PictSizeMode As eObjSizeMode = apObjSizeZoomDown, _
    Optional PictLeft, Optional PictTop, Optional PictWidth, Optional PictHeight, _
    Optional PictAngle As Single = 0!, Optional PictOpacity As Single = 100!, _
    Optional GrayScale As Boolean = False, _
    Optional TextString As String, _
    Optional TextPlacement As ePlace = ePlaceOnRight, _
    Optional TextAlignment As eAlignText = TA_LEFT, _
    Optional TextLeft, Optional ByRef TextTop, Optional ByRef TextWidth, Optional ByRef TextHeight, _
    Optional TextAngle As Single = 0!, Optional TextOpacity As Single = 100!, _
    Optional FontName, Optional FontSize, Optional FontColor, Optional FontWeight, Optional FontItalic, Optional FontUnderline, Optional FontStrikeOut, _
    Optional RotateWithText As Boolean = True, Optional TestGrid As Boolean = False, _
    Optional RetXp As Long, Optional RetYp As Long, Optional RetWp As Long, Optional RetHp As Long, _
    Optional RetXt As Long, Optional RetYt As Long, Optional RetWt As Long, Optional RetHt As Long, _
    Optional ObjectType As eObjectDataType _
    ) As Long
' ��������� ������ ������� �� ��������� ������� � ������ �������� ���������� ��������
Const c_strProcedure = "PictureData_SetToControl"
' ObjectControl - ������� � ������� ����������� ������ (�����������)
' ObjectData - �������� ������ (�������� ������ ��� ��� �� �������) ������� (�����������) ��� ��������
' Description - �������� �������
' Comment - ����������� �������
' Alignment - ������������ ������������ ������ ��������
' PictSizeMode - ����� ��������������� ������� ������������ �������� ��������
' PictLeft/PictTop - (+/-) ���������� �������� ������� ������� ������� ������������ �������� ����� �������� � ��������
' PictWidth/PictHeight - ������/������ ������� �� �������� (���� = 0 ������������ ����������� ���������)
' PictAngle - ���� ������� ������� �� �������� (� ��������, (+) ������ ������, (-) ������ �����)
' PictOpacity - �������������� ����������� � % (0-��������� ���������/100%-��������� �����������)
' GrayScale - ����� ����������� (� ������) � �������� ������
' TextString - ����� �������� ��������� ������ � ��������
' TextPlacement - ������������ ������� ������ ������������ ������� (�����������)
' TextAlignment - ������������ ������ ������������ ������� ������
' TextLeft/TextTop - (+/-) ���������� �������� ������� ������� ������ ������������ �������� ����� �������� � ��������
' TextWidth/TextHeight - ������/������ ������� ������ �� �������� (���� = 0 ������������ ����������� ���������)
' TextAngle - ���� ������� ������ � ������� ������ (� ��������, (+) ������ ������, (-) ������ �����)
' TextOpacity - �������������� ������ � % (0-��������� ���������/100%-��������� �����������)
' FontName/Size/Color/... - ��������� ������ ��� ������
' RotateWithText (�� �����������)
'   = True  - ����������� �������������� ������ c ������� (����� �������������� �� TextAngle, ���������������� � ����������� � �������������� �� PictAngle)
'   = False - ����������� �������������� ���������� �� ������ (����������� �������������� �� PictAngle, ����� �������������� �� TextAngle � ���������������� � �����������)
' TestGrid  - �������� ������� ����� ��� �������������� ��������
'(��������� ��� �������)
' Ret[X|Y|W|H][p|t] - ���������� ������������ �������. (����� ����� �������)
' ObjectType        - ���������� ��� ��������
'-------------------------
' v.1.4.1       : 12.12.2023 - ������������ ��� �������� ������������� � ������������ LaVolpe ������
' v.1.4.0       : 08.12.2023 - ����� ��������� ����������. ���������� ��� ����������� ������ � �������
' v.1.3.0       : 02.09.2021 - �������� ������ ��� ���������� FreeImage
' v.1.1.2       : 06.06.2019 - ���������� ������ ���������������� ������ ��� ���������� �����������
' v.1.1.1       : 26.12.2018 - ��������� ����������. ���������� ������������, ����������� ����������� �������/����������
'-------------------------
' ToDo: ��� �������� ������ ��� �������� ����� ������ ��������� ������ ������� ��������, � ����������� ���� � ����������� ����� �������� �� ������� �� ������� ��������
' - !! FI: ��� ���������� � ������� � ������ � Access.Image ��������� ������ ���
' - !! ������ ������ (���� ����� ���������/������������ ����� �������� ����� ������)
' + ObjectData = FIBITMAP ??? ��� ��������� ������� �������� �� ��� ���������� �� �������� FreeImage ???
'-------------------------
' �������������� ���������:
Dim Offsize As Long             '(px) ������� ����� �������� �������� � �������� ����������� �� ��.
Dim Offset As Long              '(�� ������������) '(px) �������� ������� �������� �� ������� ��������. ������� �������� ����� ����������� ����������� ���� �������� ������� �����������������������
Dim Indent As Long: Indent = 3  '(px) �������� ������� ������ �� ��������
IsDebug = 0

Dim Result As Long: Result = NOERROR 'False 'NOERROR
On Error GoTo HandleError
    ObjectType = eObjectDataUndef
    If ObjectControl Is Nothing Then Err.Raise vbObjectError + 512
#If ObjectDataType = 0 Then         'FI
    If Not FreeImage_IsAvailable Then FreeImage_LoadLibrary (True) 'False
#End If                             'ObjectDataType
'-------------
' 0. ��������� ���������� ������� � ����������� �������
'-------------
Dim lBackColor As Long, bUseBackColor As Boolean  '
' lBackColor - ���� ��� ��������� ����������� ���� (����� � Access.CommandButton � StdPicture)
'   � Access.CommandButton �������� ����� ����: vbWhite (&hFFFFFF) -> (&hFCFCFC) - ��������� � ����� vbButtonFace (&hF0F0F0)
' ��� ����� ��������� � ��� ���� ������� �� �����
Dim oControl As Object:         Set oControl = ObjectControl
Dim eCtrlType As eControlType:  eCtrlType = PictureData_GetControlType(oControl, Offsize, BackColor:=lBackColor)
'-------------
Dim Trans As New clsTransform       ' ����� �������������

Dim Wb As Single, Hb As Single      ' ������� ��������
Dim Wp As Single, Hp As Single      ' ������� �������� �� ��������
Dim Xp As Single, Yp As Single      ' ������� ������ �� �������� ������� �������� ����� ��������
Dim Wt As Single, Ht As Single      ' ������� ������ �� ��������
Dim Xt As Single, Yt As Single      ' ������� ������ �� �������� ������� ������ ����� ��������
Dim dXp As Single, dYp As Single    ' �������� ������� ������ �������� ������������ ����� �������� � ��������
Dim dXt As Single, dYt As Single    ' �������� ������� ������ ������ ������������ ����� �������� � ������
Dim oXp As Single, oYp As Single    ' ���������� ����� �������� ��������
Dim oXt As Single, oYt As Single    ' ���������� ����� �������� ������

'Dim lColorOptions As FREE_IMAGE_COLOR_OPTIONS
   
#If ObjectDataType = 0 Then     'FI
Dim fiBack As LongPtr, fiTemp As LongPtr
#ElseIf ObjectDataType = 1 Then 'LV
Dim lvBack As New clsPictureData: Set lvBack = New clsPictureData
#End If                         'ObjectDataType
'-------------
' 1. �������������� �������� (Back)
'-------------
HandleBack:
' �������� ���������������� ������� ����� �������� � ��������
    ' �.�. ��� ����������� (���������� ��������), �� �� ���� ����� 9 ��������� ���������
    ' ���������������� ������� ����� �������� ������� � �������� ����� ��������� (LT-LT;CC-CC � �.�.)
Dim rXb As Single, rYb As Single: Call p_GetAlignPoint(Alignment, rXb, rYb) ' �� �������� � ��������
Dim rXo As Single, rYo As Single: rXo = rXb: rYo = rYb                      ' �� �������� � ��������
' ������� ����������� �� ������� �������� � �������������� ��� ��� ��������
    Select Case eCtrlType
    Case eCtrlAccDib, eCtrlAccEmf
    ' Access control with PictureData property
            AccControlLocation oControl, , , Wb, Hb, ClientAreaPos:=True
    Case eCtrlPicture, eCtrlPicEmf
    ' Control with Picture property
        If (TypeOf oControl Is CustomControl) Then
    ' ActiveX with Picture property
            AccControlLocation oControl, , , Wb, Hb, ClientAreaPos:=True
        Else
    #If ObjectDataType = 0 Then     'FI
            fiBack = FreeImage_CreateFromOlePicture(oControl.Picture): If fiBack <> 0 Then Wb = FreeImage_GetWidth(fiBack): Hb = FreeImage_GetHeight(fiBack)
    #ElseIf ObjectDataType = 1 Then 'LV
            lvBack.LoadPicture_StdPicture oControl.Picture: If lvBack.Handle <> 0 Then Wb = lvBack.Width: Hb = lvBack.Height
    #End If                         'ObjectDataType
        End If
    Case eStdPicture
    ' StdPicture
    If IsNumeric(PictWidth) Then If PictWidth > 0 Then Wb = Abs(PictWidth)
    If IsNumeric(PictHeight) Then If PictHeight > 0 Then Hb = Abs(PictHeight)
    If Wb < 1 Then Wb = p_Max(Hb, 16) '
    If Hb < 1 Then Hb = Wb
    #If ObjectDataType = 0 Then     'FI
            fiBack = FreeImage_CreateFromOlePicture(oControl): If fiBack <> 0 Then Wb = FreeImage_GetWidth(fiBack): Hb = FreeImage_GetHeight(fiBack)
    #ElseIf ObjectDataType = 1 Then 'LV
            lvBack.LoadPicture_StdPicture oControl: If lvBack.Handle <> 0 Then Wb = lvBack.Width: Hb = lvBack.Height
    #End If                         'ObjectDataType
    Case Else
    ' Somthing unknown
        Err.Raise vbObjectError + 512: GoTo HandleError
    End Select
'If Offset Then Stop
    ' ������ �������� �� ������ ������� ��������� (��. PictureData_GetControlType)
    Wb = Wb - Offsize: Hb = Hb - Offsize
' ���� ��� �� ������� - ������ ������ ��������
    #If ObjectDataType = 0 Then     'FI
    If (fiBack = 0) Then fiBack = FreeImage_Allocate(Wb, Hb, 32): fiTemp = FreeImage_Composite(fiBack, 0&, lBackColor, 0&): FreeImage_Unload (fiBack): fiBack = fiTemp
    #ElseIf ObjectDataType = 1 Then 'LV
    If lvBack.Handle = 0 Then lvBack.InitializeDIB Wb, Hb, lBackColor
    #End If                         'ObjectDataType
' ��� ������������� ������ � �������� ���� ����������� �����
    #If ObjectDataType = 0 Then     'FI
    If TestGrid Then fiBack = FreeImage_CreateCheckerBoard(Wb, Hb)
    #ElseIf ObjectDataType = 1 Then 'LV
    If TestGrid Then lvBack.CreateCheckerBoard , vbWhite - 10
    #End If
    'ObjectDataType
'-------------
' 2. �������������� ����������� (Pict)
'-------------
HandlePict:
With Trans
' �������� ������� ����������� �� ��������, ���� �� ������ - ���� ������� ��������
    If Not IsNumeric(PictWidth) Then Wp = Wb Else If PictWidth = 0 Then Wp = Wb Else Wp = IIf(Abs(PictWidth) > 1, PictWidth, PictWidth * Wb)
    If Not IsNumeric(PictHeight) Then Hp = Hb Else If PictHeight = 0 Then Hp = Hb Else Hp = IIf(Abs(PictHeight) > 1, PictHeight, PictHeight * Hb)
    If IsNumeric(PictLeft) Then dXp = IIf(Abs(PictLeft) >= 1, PictLeft, PictLeft * Wb)
    If IsNumeric(PictTop) Then dYp = IIf(Abs(PictTop) >= 1, PictTop, PictTop * Hb)
' �������� Bitmap ���������� ��������
    #If ObjectDataType = 0 Then     'FI
Dim fiPict As LongPtr
    fiPict = PictureData_LoadFromEx(ObjectData, Abs(Wp), Abs(Hp), ObjectType): If fiPict = 0 Then GoTo HandleText ' :Wp = 0: Hp = 0
    #ElseIf ObjectDataType = 1 Then 'LV
' �������� LavVolpe Bitmap ���������� ��������
Dim lvPict As New clsPictureData: Set lvPict = New clsPictureData
    Set lvPict = PictureData_LoadFromEx(ObjectData, Abs(Wp), Abs(Hp), ObjectType): If lvPict.Handle = 0 Then GoTo HandleText ' :Wp = 0: Hp = 0
    #End If                         'ObjectDataType
' ��� ������������� ��������
    #If ObjectDataType = 0 Then     'FI
    If (Wp < 0) Then Call FreeImage_FlipHorizontal(fiPict)
    If (Hp < 0) Then Call FreeImage_FlipVertical(fiPict) ':   Hp = -Hp
    #ElseIf ObjectDataType = 1 Then 'LV
    If ((Wp < 0) Or (Hp < 0)) Then Call lvPict.MirrorImage((Wp < 0), (Hp < 0))
    #End If                         'ObjectDataType
' ��� ������������� �����
    ' ��� �������������� ������������ � ����� ���������� ��������� ��� �������
    ' ������ ���� ����� ������ ����� ��������������� ���� ������������ ��������������� <1
    ' �� �� ���� ����� ������ �� �����
    #If ObjectDataType = 0 Then     'FI
    If GrayScale Then fiPict = FreeImage_ConvertToAlphaGreyScale(fiPict, True) 'FreeImage_ConvertColorDepth(fiPict, FICF_GREYSCALE, True)
    #ElseIf ObjectDataType = 1 Then 'LV
    If GrayScale Then lvPict.MakeGrayScale gsclNTSCPAL
    #End If                         'ObjectDataType
' ������������ ����������� � ������������ � PictSizeMode �� �������� ��������
Dim dXp0 As Single, dYp0 As Single
Dim Wp1 As Single, Hp1 As Single:    Wp1 = Wp: Hp1 = Hp                         ' Wp1/Hp1   - ��������� ������� ��������
    #If ObjectDataType = 0 Then     'FI
    Wp = FreeImage_GetWidth(fiPict): Hp = FreeImage_GetHeight(fiPict)           ' Wp/Hp     - ������� ������� ��������
    #ElseIf ObjectDataType = 1 Then 'LV
    Wp = lvPict.Width: Hp = lvPict.Height
    #End If                         'ObjectDataType
    Call p_GetSizeFactor(PictSizeMode, Wp, Hp, Abs(Wp1), Abs(Hp1), dXp0, dYp0)  ' dXp0/dYp0 - ������������ ��������������� �������� � ����������� ��������
    #If ObjectDataType = 0 Then     'FI
    If ((dXp0 <> 1) Or (dYp0 <> 1)) Then fiPict = FreeImage_RescaleEx(fiPict, dXp0, dYp0, False): Wp = FreeImage_GetWidth(fiPict): Hp = FreeImage_GetHeight(fiPict)
    If ((Wp <= 0) Or (Hp <= 0)) Then FreeImage_Unload (fiPict): GoTo HandleText ': Wp = 0: Hp = 0
    #ElseIf ObjectDataType = 1 Then 'LV
    If ((dXp0 <> 1) Or (dYp0 <> 1)) Then Call lvPict.Resize(dXp0 * Wp, dYp0 * Hp): Wp = lvPict.Width: Hp = lvPict.Height
    If ((Wp <= 0) Or (Hp <= 0)) Then lvPict.DestroyDIB: GoTo HandleText ': Wp = 0: Hp = 0
    #End If                         'ObjectDataType
' ������������ �����������
Dim pAngle As Single: pAngle = IIf(IsNumeric(PictAngle), PictAngle, 0)
    .Angle = pAngle
    Call .TransformSize(Wp, Hp, Wp1, Hp1)                                       ' Wp1/Hp1   - ������/������ ������� �������� (����� ��������)
    #If ObjectDataType = 0 Then     'FI
    ' ���������� ���������������� FreeImage_RotateExEx, ����� �� ��������������� ������� ����� ��������
    If .Angle <> 0 Then fiPict = p_FreeImage_RotateExEx(fiPict, Wp, Hp, Wp1, Hp1, .Angle, &H0&)
    #ElseIf ObjectDataType = 1 Then 'LV
    ' ��� LaVolpe ������� ����� ������ ��� ������������  ����������, ���� ������ ������������ ������� ������
    #End If                         'ObjectDataType
' �������� ������� ������ ������� ����������� ����� �������� �� �������� �������� ��������
    Call .GetDelta(Wp, Hp, dXp0, dYp0)                                          ' dXp0/dYp0 - �������� ����� �������� ����� �������� ������������ ����� 0 �������� �� ��������
    dXp = rXb * Wb - rXo * Wp1 + dXp0 + dXp                                     ' dXp/dYp   - �������� ����� �������� ����� �������� ������������ ����� 0 ��������
    dYp = rYb * Hb - rYo * Hp1 + dYp0 + dYp
' �������� ������� ������ �������� (Xp,Yp) �� ��������
    Call .Transform(0.5, 0.5, Xp, Yp, Wp, Hp)                                   ' Xp/Yp - ���������� ������ ������� ��������
    #If ObjectDataType = 0 Then     'FI
    Xp = Xp - 0.5 * Wp1 + dXp + Offset: Yp = Yp - 0.5 * Hp1 + dYp + Offset    ' Xp/Yp - ���������� ������ �������� ���� ������� ��������� �������� �� ��������
    #ElseIf ObjectDataType = 1 Then 'LV
    Xp = Xp - 0.5 * Wp + dXp + Offset: Yp = Yp - 0.5 * Hp + dYp + Offset       ' Xp/Yp - ���������� ������ �������� ���� �������� �� �������� �� ��������
    #End If                         'ObjectDataType
'-------------
' 3. �������������� ����� (Text)
'-------------
HandleText:
    If (Len(Trim$(TextString)) <= 0) Then GoTo HandleComposite
' ������ �����
    If IsNumeric(FontSize) Then FontSize = Abs(FontSize): If FontSize < 1 Then FontSize = FontSize * p_Min(Wb, Hb)  ' /pt
Dim hFont As LongPtr: hFont = CreateHFont(FontName, FontSize, FontWeight, FontItalic, FontUnderline, FontStrikeOut) ': If hFont=0 then GoTo HandleComposite
' �������� ������� ������� ������ �� ��������, ���� �� ������ - ������������ ������ �� ���������� ����� � ������������
    If IsNumeric(TextWidth) Then Wt = IIf(Abs(TextWidth) > 1, TextWidth, TextWidth * Wb)
    If IsNumeric(TextHeight) Then Ht = IIf(Abs(TextHeight) > 1, TextHeight, TextHeight * Hb)
    If IsNumeric(TextLeft) Then dXt = IIf(Abs(TextLeft) >= 1, TextLeft, TextLeft * Wb)
    If IsNumeric(TextTop) Then dYt = IIf(Abs(TextTop) >= 1, TextTop, TextTop * Hb)
' ���������� ����� �������� �� ������ ����� �������� (� ����������� ��������).
    #If ObjectDataType = 0 Then     'FI
    If fiPict = 0 Then
    #ElseIf ObjectDataType = 1 Then 'LV
    If lvPict.Handle = 0 Then
    #End If                         'ObjectDataType
    ' ���� �������� ��� - ����� �������� � ���������� �������� ��������
        TextPlacement = Alignment Or (Alignment * &H10): RotateWithText = False
    End If
    If Indent Then
    ' ���� ���� ������ ������ ��� ������ � ����������� �� ����� ��������
    Select Case TextPlacement And &H30
    Case &H30: dXt = dXt + 0.5 * Indent ' CenterHorz
    Case &H20: dXt = dXt - Indent ' OnLeft | ToRight
    Case &H10: dXt = dXt + Indent ' ToLeft | OnRight
    End Select
    Select Case TextPlacement And &HC0
    Case &HC0: dYt = dYt + 0.5 * Indent ' CenterVert
    Case &H80: dYt = dYt - Indent ' OnTop | ToBottom
    Case &H40: dYt = dYt + Indent ' ToTop | OnBottom
    End Select
    End If
' �������� ���������������� ������� ����� �������� ������ � ��������
Dim tAngle As Single: tAngle = IIf(IsNumeric(TextAngle), TextAngle, 0)
Dim rXp As Single, rYp As Single: Call p_GetAlignPoint(TextPlacement Mod &H10, rXp, rYp)    ' �� �������� � ������
Dim rXt As Single, rYt As Single: Call p_GetAlignPoint(TextPlacement \ &H10, rXt, rYt)      ' �� ������ � ��������
    #If ObjectDataType = 0 Then     'FI
    If fiPict = 0 Then
    #ElseIf ObjectDataType = 1 Then 'LV
    If lvPict.Handle = 0 Then
    #End If                         'ObjectDataType
    ' ���� �������� ��� - ���� ����������� (rXp, rYp) ���������������� ������� ����� �������� �� �������� � ������:
        ' �������� �������� ���������� ����� �������� �������� �� ��������� �������
        ' ������������ �� �� tAngle � �������� ����� �������� ������ �� ��������� �������
        pAngle = tAngle
'If tAngle <> 0 Then Stop
    End If
    ' � ����������� �� RotateWithText ���������� ��������� ����� �������� ������ � ��������
    Call p_GetAnchorPoint(oXt, oYt, rXp, rYp, Wp, Hp, Trans, RotateWithText, pAngle, dXt, dYt)
    oXt = oXt + dXp: oYt = oYt + dYp
' ������������ ������� ������
    If RotateWithText Then .Angle = .Angle + tAngle Else .Angle = tAngle
' �������� ������������ ��������� ������� ������ ��� �������� ���� �������� � ����� ��������
    Call p_GetMaxRect(Wt, Ht, Wb, Hb, oXt, oYt, rXt, rYt, Trans)        ' Wt/Ht     - ��������� ������� ������� ��� ������ (�� ��������)
    If Indent Then
    ' ���� ���� ������ ������ ��� ������ � ����������� �� ����� ��������
    If TextPlacement And &H30 Then Wt = Wt - Indent ' Horz
    If TextPlacement And &HC0 Then Ht = Ht - Indent ' Vert
    End If
' �������� ������ �������� ��������
Dim aText() As String                                                   ' aText     - ����� �������� �� ������
    Call TextToArrayByHFont(TextString, hFont, Wt, Ht, , aText)         ' Wt/Ht     - ��������� ������� ��������� ������ (�� ��������)
Dim lFontColor As Long: If Not IsMissing(FontColor) Then lFontColor = GetColorFromText(FontColor) Else lFontColor = vbBlack
    #If ObjectDataType = 0 Then     'FI
Dim fiText As LongPtr
    fiText = FreeImage_DrawText(aText, hFont, Wt, Ht, TextAlignment, lFontColor) ', lBackColor)
    Wt = FreeImage_GetWidth(fiText): Ht = FreeImage_GetHeight(fiText)   ' Wt/Ht     - �������� ������� ������
    'If GrayScale Then fiText = FreeImage_ConvertToAlphaGreyScale(fiText, True) 'FreeImage_ConvertColorDepth(fiText, FICF_GREYSCALE, True)
    #ElseIf ObjectDataType = 1 Then 'LV
Dim lvText As New clsPictureData: Set lvText = New clsPictureData
    Call lvText.DrawText_hFont(hFont, aText, TextAlignment, DestWidth:=Wt, DestHeight:=Ht, ForeColor:=lFontColor): If lvText.Handle = 0 Then GoTo HandleComposite
    Wt = lvText.Width: Ht = lvText.Height                               ' Wt/Ht     - �������� ������� ������
    'If GrayScale Then lvText.MakeGrayScale gsclNTSCPAL                  ' ���� ���� - ������ �����
    #End If             'ObjectDataType
' ������������ �����
Dim Wt1 As Single, Ht1 As Single
    Call .TransformSize(Wt, Ht, Wt1, Ht1)                               ' Wt1/Ht1   - ������/������ ������� ������ (����� ��������)
    #If ObjectDataType = 0 Then     'FI
    If .Angle <> 0 Then fiText = FreeImage_Rotate(fiText, .Angle, &H0&)
    #ElseIf ObjectDataType = 1 Then 'LV
    ' ��� LaVolpe ������� ����� ������ ��� ������������  ����������, ���� ������ ������������ ������� ������
    #End If                         'ObjectDataType
' !!! �������� !!!  ����� ��������� �������� ��������� ������ ���������� ����������� ����� �������� � ��������
    Call .Transform(rXt, rYt, Xt, Yt, Wt, Ht)                           ' Xt/Yt - ���������� ����� �������� ������ �� ������ ������������ �0 (������� A) ������
    dXt = oXt - Xt: dYt = oYt - Yt
' ������� ������ ������
    Call .Transform(0.5, 0.5, Xt, Yt, Wt, Ht)                           ' Xt/Yt - ���������� ������ ������� ������ ������������ �0 (������� A) ������
    #If ObjectDataType = 0 Then     'FI
    Xt = Xt - 0.5 * Wt1 + dXt: Yt = Yt - 0.5 * Ht1 + dYt                ' Xt/Yt - ���������� ������ �������� ���� ������� ���������� ������ �� ��������
    #ElseIf ObjectDataType = 1 Then 'LV
    Xt = Xt - 0.5 * Wt + dXt: Yt = Yt - 0.5 * Ht + dYt: tAngle = .Angle ' Xt/Yt - ���������� ������ �������� ���� ������ �� �������� �� ��������
    #End If                         'ObjectDataType
End With
'-------------
' 4. ��������� ����������� � ����� � ������ ������������/��������
'-------------
HandleComposite:
'If IsDebug Then
    #If ObjectDataType = 0 Then     'FI
'FreeImage_Save FIF_BMP, fiBack, CurrentProject.path & "\fiBack.bmp"
'FreeImage_Save FIF_BMP, fiText, CurrentProject.path & "\fiText.bmp"
'FreeImage_Save FIF_BMP, fiPict, CurrentProject.path & "\fiPict.bmp"
'End If
' ��������� ������ � ������ ������� �����������, ������ � ���������� ��������� ������������
    If fiBack = 0 Then Err.Raise vbObjectError + 512
' ���� ���� ����������� ������ �� �������� ����������� � �������� �������������
    If fiPict <> 0 Then Call FreeImage_CompositeWithAlpha(fiBack, fiPict, Xp, Yp, Alpha:=(PictOpacity * 255 / 100)): FreeImage_Unload (fiPict)
' ���� ���� ����� ������ �� �������� ����� � �������� �������������
    If fiText <> 0 Then Call FreeImage_CompositeWithAlpha(fiBack, fiText, Xt, Yt, Alpha:=(TextOpacity * 255 / 100)): FreeImage_Unload (fiText)
'If IsDebug Then
'FreeImage_Save FIF_BMP, fiBack, CurrentProject.path & "\fiComposite.bmp": Stop
'End If
    #ElseIf ObjectDataType = 1 Then 'LV
' ��������� ������ � ������ ������� �����������, ������ � ���������� ��������� ������������
    If lvBack.Handle = 0 Then Err.Raise vbObjectError + 512
' ���� ���� ����������� ������ �� �������� ����������� � �������� �������������
    If lvPict.Handle <> 0 Then lvPict.Render 0, Xp, Yp, Wp, Hp, DestHostDIB:=lvBack, Angle:=-pAngle, Opacity:=PictOpacity: Set lvPict = Nothing
' ���� ���� ����� ������ �� �������� ����� � �������� �������������
    If lvText.Handle <> 0 Then lvText.Render 0, Xt, Yt, Wt, Ht, DestHostDIB:=lvBack, Angle:=-tAngle, Opacity:=TextOpacity: Set lvText = Nothing
    #End If                         'ObjectDataType
'-------------
' 5. ��������� �������� ����������� (��������+�������+�����)
'-------------
HandleSetControl:
    With oControl
' ����������� ����������� � �������� �� ������-�������� ���� (Top-Left)
    ' ����� ����������� � ����� ����������������� �� �������� ��� ��������
    ' �.�. ������ �������� �� ������� ��������
    ' � ��������� �������� ������������ �� ���� ��������
        On Error Resume Next: .PictureAlignment = 0: Err.Clear    '
        On Error GoTo HandleError
' � ����������� �� ���� �������� ����������� ���������� ����������� � �����. ������
    #If ObjectDataType = 0 Then     'FI
        Select Case eCtrlType
        Case eCtrlAccDib:  .PictureData = FreeImage_GetPictureData(fiBack)
        Case eCtrlAccEmf:  .PictureData = FreeImage_GetPictureDataEMF(fiBack)
        Case eCtrlPicture: .Picture = FreeImage_GetOlePicture(fiBack)
        Case eCtrlPicEmf:  .Picture = FreeImage_GetOlePicture(fiBack)        'FreeImage_GetOlePictureEMF(fiBack)         '.Picture = FreeImage_GetOlePicture(fiBack) '������������ �����������
        Case eStdPicture:  Set ObjectControl = FreeImage_GetOlePicture(fiBack)  '.Picture = FreeImage_GetOlePicture(fiBack) '������������ �����������
        Case Else
        End Select
'If IsDebug Then
'FreeImage_Save FIF_BMP, fiBack, CurrentProject.path & "\fiBack+Pict.bmp"
'End If
    Call FreeImage_Unload(fiBack)
    #ElseIf ObjectDataType = 1 Then 'LV
Dim aObjData() As Byte
' � ����������� �� ���� �������� ����������� ���������� ����������� � �����. ������
        Select Case eCtrlType
        Case eCtrlAccDib: ' �� ������ ����������� �����:' vbWhite (&hFFFFFF) -> (&hFCFCFC) - ��������� � ����� vbButtonFace (&hF0F0F0)
            If Not lvBack.SaveToStream_PictureData(aObjData, picDIB) Then Err.Raise vbObjectError + 512
            .PictureData = aObjData: Erase aObjData:  Set lvBack = Nothing
        Case eCtrlAccEmf: ' vbWhite (&hFFFFFF) -> (&hFDFDFD/&hFEFEFE)
            If Not lvBack.SaveToStream_PictureData(aObjData, picEMF) Then Err.Raise vbObjectError + 512 '
            .PictureData = aObjData: Erase aObjData:  Set lvBack = Nothing
        Case eCtrlPicture, eCtrlPicEmf:
            If Not lvBack.SaveToStream_BMP(aObjData) Then Err.Raise vbObjectError + 512
            .Picture = ArrayToPicture(aObjData, 0&, 1& + UBound(aObjData)): Erase aObjData: Set lvBack = Nothing
        Case eStdPicture:
            If Not lvBack.SaveToStream_BMP(aObjData) Then Err.Raise vbObjectError + 512
            Set ObjectControl = ArrayToPicture(aObjData, 0&, 1& + UBound(aObjData)): Erase aObjData: Set lvBack = Nothing
        Case Else
            Set lvBack = Nothing: Err.Raise vbObjectError + 512
        End Select
    'Call lvBack.EraseDIB
    #End If                         'ObjectDataType
    End With
    Result = NOERROR ' False 'NOERROR
HandleExit:
'' ���������� ������� / ������� ��������� ����������� �� ��������
    RetXp = Xp: RetYp = Yp: RetWp = Wp: RetHp = Hp
    RetXt = Xt: RetYt = Yt: RetWt = Wt: RetHt = Ht
    PictureData_SetToControl = Result: Exit Function
HandleError:
    Select Case Err.Number
    Case 2004: Err.Clear: Stop              ' ������������ ������
    Case Else: Result = False: Err.Clear: Resume HandleExit
    End Select
    Err.Clear: Resume HandleExit
End Function
Private Function p_GetSizeFactor(SizeMode As eObjSizeMode, _
    W0 As Single, H0 As Single, _
    W1 As Variant, H1 As Variant, _
    fW As Single, fH As Single)
' ���������� ������������ ��������������� � ����������� �� ���������� ������ ���������������
'---------------------
' ��������:
'   SizeMode - ����� ���������������
'   W0/H0    - ������/������ ��������
'   W1/H1    - ������/������ ���������
' ����������:
'   fW/fH    - ������������ ���������������
'---------------------
    Select Case SizeMode
    Case apObjSizeZoomDown: fW = W1 / W0: fH = H1 / H0  '-1 - ���������������� ��������������� (������ ����������)
                            If fW < fH Then fH = fW Else fW = fH
                            If fW > 1 Then fW = 1: fH = 1
    Case apObjSizeZoom:     fW = W1 / W0: fH = H1 / H0  ' 3 - ���������������� ���������������
                            If fW < fH Then fH = fW Else fW = fH
    Case apObjSizeStretch:  fW = W1 / W0: fH = H1 / H0  ' 1 - ������/���������� (�������� ���������)
    'Case apObjSizeClip:     fW = 1: fH = 1                  ' 0 - �� ������ ������ ���� ������ ������ ������� ������ - �������
    Case Else:              fW = 1: fH = 1                  ' 0 - �� ������ ������ ���� ������ ������ ������� ������ - �������
    End Select
End Function
Private Function p_GetAlignPoint(Alignment As eAlign, _
    cX As Single, cY As Single)
' ���������� ���������������� ���������� ����� �������� � ����������� �� ��������� ������ ������������
'---------------------
' ��������:
'   Alignment - ����� ������������
' ����������:
'   cX,cY     - ������� ����� �������� ����� �������������
'---------------------
    ' Horz region anchor point position
    Select Case (Alignment And eCenterHorz)
    Case eLeft:         cX = 0            ' Left-to-Left
    Case eRight:        cX = 1            ' Right-to-Right
    Case eCenterHorz:   cX = 1 / 2        ' CenterHorz-to-CenterHorz
    End Select
    ' Vert region anchor point position
    Select Case (Alignment And eCenterVert)
    Case eTop:          cY = 0            ' Top-to-Top
    Case eBottom:       cY = 1            ' Bottom-to-Bottom
    Case eCenterVert:   cY = 1 / 2        ' CenterVert-to-CenterVert
    End Select
End Function
Private Function p_PolarRadiusForRect(gAngle As Single, w As Single, h As Single, Optional gTilt As Single) As Single
' ���������� �������� ������ ����� �� ����������� �������������� ��������� ��������� w/h
' gAngle - �������� ���� (� ��������)
' w/h    - ������/������ ��������������
' gTilt  - ������ �������������� ������������ ��� (��������� �.�. ����� ����� � gAngle)
Dim p As Single: p = ((gAngle + gTilt) Mod 360): If p < 0 Then p = p + 360          ' �������� ���� � �������� (0-360)
Dim d As Single: If w <> 0 Then d = Atn(h / w) * 180 / Pi Else d = 90               ' ���� ������� ��������� �������������� (0-360)
    Select Case p
    Case 0, 180:                 p_PolarRadiusForRect = 0.5 * w
    Case 90, 270:                p_PolarRadiusForRect = 0.5 * h
    Case (d) To (180 - d):       p_PolarRadiusForRect = 0.5 * h / Sin(p * Pi / 180)   ' AB
    Case (180 - d) To (180 + d): p_PolarRadiusForRect = -0.5 * w / Cos(p * Pi / 180)  ' BC
    Case (180 + d) To (360 - d): p_PolarRadiusForRect = -0.5 * h / Sin(p * Pi / 180)  ' CD
    Case Else:                   p_PolarRadiusForRect = 0.5 * w / Cos(p * Pi / 180)   ' AD
    End Select
End Function
Private Function p_GetAnchorPoint( _
    oXc As Single, oYc As Single, _
    rXh As Single, rYh As Single, _
    Wh As Single, Hh As Single, _
    Trans As clsTransform, _
    Optional RotClient As Boolean = False, _
    Optional RotAngle As Single, _
    Optional dXc As Single, Optional dYc As Single _
    )
' ������������ ���������� ����� ��������
'---------------------
' ��������:
'   rXh/rYh   - ���������������� ���������� ����� �������� �� ������� �������������� (������������ ��� ������/������)
'   Wh/Hh     - ������� ��������� ������� � ������� ���� ���������� �������������
'   Trans     - ������ �� ����� ������������� ��������� �������
'   RotClient - ������� ���� ��� ����������� ������� �������������� ������ � ���������
'   RotAngle  - ���� �������� ��������� �������
'   dXc/dYc   - �������� ��������� �.�������� ����������� �������
' ����������:
'   oXc/oYc   - ���������� ����� �������� �� ��������� ������� (� ����������� ��������� �������)
'---------------------
    On Error GoTo HandleError
With Trans
    ' � ����������� �� RotClient ���������� ��������� ����� �������� ����������� ������� � ���������
    If RotClient Then
    ' ���� RotClient = True - ����� �������� ��� �������� ������� �� ��������� ������� � ��� �� ������� ��� � ����, �������� ����������� ������������ ����������� ��������, ���� �������� ������ ����� TextAngle+PictAngle
        Call .Transform(rXh * Wh + dXc, rYh * Hh + dYc, oXc, oYc)   ' oXc/oYc   - ���������� ����� �������� ����������� ������� �� ���������
    Else
    ' ���� RotClient = False - ����� �������� ��� �������� ��������, �������� ����������� ������������ �����������, ���� �������� ����������� ������� ��������
        ' ������������� ��������� ���������� ����� �������� �� �������� � �������� ������������ ������ ��������� �������
        ' �������� ���������������� �������� ���������� �������� ����� �������� ������������ ��������� ������� �������������� � ���� ����������� (Phi)
        ' ������� ������������� ���������������� �������� ���������� ����� �������� � ���������, �� ��� ��� ���� � ������ ������� �������������� (Phi+RotAngle)
Dim Phi As Single, Rho As Single, rRho As Single
        oXc = (0.5 - rXh) * Wh: oYc = (0.5 - rYh) * Hh              ' oXc/oYc   - ��������� ���������� ����� �������� ����������� ������� � ����������� ��������� ������� ������������ ������ ��������� �������
        Rho = Sqr(oXc ^ 2 + oYc ^ 2)                                ' Rho       - �������� ������ ����� �������� �� ��������� �������
        Phi = p_ATan2(oXc, oYc) * 180 / Pi                          ' Phi       - �������� ���� ����� �������� �� ��������� �������
        rRho = Rho / p_PolarRadiusForRect(Phi, Wh, Hh)              ' rRho      - ���������������� �������� ������ ����� �������� ������������ ����������� ��������������
        Rho = rRho * p_PolarRadiusForRect(Phi + RotAngle, Wh, Hh)   ' Rho       - �������� ������ ����� �������� �� ��������� ������� ����� ��������
        Call .Transform(0.5, 0.5, oXc, oYc, Wh, Hh)                 ' oXc/oYc   - ���������� ������ ��������� �������
        oXc = -Rho * Cos(Phi * Pi / 180) + oXc + dXc                ' oXc/oYc   - ���������� ����� �������� �� ��������� ������� (� ����������� ��������� �������)
        oYc = -Rho * Sin(Phi * Pi / 180) + oYc + dYc
    End If
End With
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Private Function p_GetMaxRect( _
    Wc As Single, Hc As Single, _
    Wh As Single, Hh As Single, _
    oXc As Single, oYc As Single, _
    rXc As Single, rYc As Single, _
    Trans As clsTransform, _
    Optional dXc As Single, Optional dYc As Single _
    )
' ������������ ������������ ��������� ������������� ��� �������� ����� �������� � ���� ��������
'---------------------
' ��������:
'   Wh/Hh     - ������� ��������� ������� � ������� ���� ���������� �������������
'   oXc/oYc   - ���������� ����� �������� �� ��������� ������� (� ����������� ��������� �������)
'   rXc/rYc   - ���������������� ���������� ����� �������� �� ������� �������������� (������������ ��� ������/������)
'   Trans     - ������ �� ����� ������������� ����������� �������
' ����������:
'   Wc/Hc     - ������������ ������� ��������������
'   dXc/dYc   - �������� �������� ��������� �.�������� ������������ ���������
'---------------------
Const dErr = 0.0001     ' ���������� ����������� ������� ������� �����
' ! ������� ����������� !
    On Error GoTo HandleError
' ��� ���� �������� ���� Wc=0 b Hc=0
    If Wc > 0 And Hc > 0 Then Exit Function
    If Wc > 0 Then Exit Function
    'If Hc > 0 Then Stop: Exit Function
' ���� Wc>0 � Hc>0  -> ������ �� ��������� - ������� ������ �����������
' ���� Wc>0         -> ������ ����� �� ��������� - ������ ������ ��� ������ ��� ��� ��� ����� ������������ �������� ������
' ���� Hc>0         -> ����� ������ ��� ������ ������ ��� ������� ������� ����� ������������

With Trans
' ���������� ���� ������� ������ � ������� ������ ���������� ��� ��������
Dim radB As Single: radB = (.Angle Mod 90) * Pi / 180                   ' ���� � �������� � ��������� [0-Pi/2]
Dim SinB As Single, CosB As Single:     SinB = Sin(radB): CosB = Cos(radB)
Dim Sin2B As Single, Cos2B As Single:   Sin2B = Sin(2 * radB): Cos2B = Cos(2 * radB)
' ����� ��������� �������
Dim c:  c = Array(oXc, oYc, Wh - oXc, Hh - oYc)     ' c0/c1     (~)
Dim kX: kX = Array(rXc, 1 - rXc, 1 - rXc, rXc)      ' kX0/kX1   (=)
Dim kY: kY = Array(rYc, rYc, 1 - rYc, 1 - rYc)      ' kY0/kY1   (=)

Dim s As Long
Dim k0  As Byte, k1  As Byte
Dim c0  As Byte, c1  As Byte
Dim f0 As Single, f1 As Single
Dim r As Single
Dim Wc0  As Single, Hc0 As Single
Dim Wc1  As Single, Hc1 As Single
Dim Area As Single
Dim dXc0 As Single, dYc0 As Single
Dim oXc0 As Single, oYc0 As Single
' ������������ ������� ������������� �������������� ������������ �������� ����� ��������
    ' ���� Sin2B=0 ���������� ������ ���� (���� 0-3)
Dim n As Byte: n = .Angle \ 90   ' �������� � ����������� �� ���� ������� ����� ��������� �������
        If Sin2B = 0 Then s = 4: k1 = k0 + 1: c0 = (k0 + 4 - n) Mod 4
        For s = s To 9
    ' C��������� ����� 10 ��������� ���������� ��� ������� ����������� ������������� ����� ���� ������ � ��������� �������:
            Wc0 = 0: Hc0 = 0
            If s Mod 2 Then     ' 1,3,5,7,9
                f0 = CosB: f1 = SinB
            Else                ' 0,2,4,6,8
                f0 = SinB: f1 = CosB
            End If
            If k1 = 0 Then
    ' 0-3: ���� ������� ������� ������� ����� �� ������� ��������, ���������� ���� ������� (Chi) ����� ���� ������� ������� ������� (TextAngle)
                If ((kX(k0) <> 0) And (kY(k0) <> 0)) Then
        ' ������� ��������������� �������
                    c0 = (k0 + 4 - n) Mod 4: r = c(c0) / Sin2B:
        ' ������� �������
                    If Wc <= 0 Then Wc0 = r * (f0 / kX(k0)) Else Wc0 = Wc
                    If Hc <= 0 Then Hc0 = r * (f1 / kY(k0)) Else Hc0 = Hc
                End If
                k0 = k0 + 1: If k0 = 4 Then k0 = 0: k1 = k0 + 1: c0 = (k0 + 4 - n) Mod 4
            Else
    ' 4-9: ��� ������� ������� ������� ����� �� �������� ��������
        ' ������� ��������������� �������
            ' ������� ������� ������� ������� ����� �� �������� �������� (AB,CD,AD)
            ' ������� ������� ������� ������� ����� �� �������� �������� (BC)
            ' ��������������� ������� ������� ������� ����� �� �������� �������� (AC,BD)
                c1 = (k1 + 4 - n) Mod 4
                Select Case s
                Case 4, 6, 9:   r = (kX(k0) * kY(k1) - kX(k1) * kY(k0)) + (kX(k0) * kY(k1) + kX(k1) * kY(k0)) * Cos2B       ' AB,CD,AD
                Case 7:         r = (kX(k0) * kY(k1) - kX(k1) * kY(k0)) - (kX(k0) * kY(k1) + kX(k1) * kY(k0)) * Cos2B       ' BC
                Case 5, 8:      r = (kX(k0) - kY(k0))                                                                       ' AC,BD
                End Select
                If (r <> 0) Then
        ' ������� �������
            ' ������� ������� ������� ������� ����� �� �������� �������� (AB,CD,BC)
            ' ������� ������� ������� ������� ����� �� �������� �������� (AD)
            ' ��������������� ������� ������� ������� ����� �� �������� �������� (AC,BD)
                Select Case s
                Case 4, 6, 7:   If Wc <= 0 Then Wc0 = (2 * (c(c0) * kY(k1) * f1 - c(c1) * kY(k0) * f0)) / r Else Wc0 = Wc   ' AB,CD,BC
                                If Hc <= 0 Then Hc0 = -(2 * (c(c0) * kX(k1) * f0 - c(c1) * kX(k0) * f1)) / r Else Hc0 = Hc
                Case 9:         If Wc <= 0 Then Wc0 = (2 * (c(c0) * kY(k1) * f0 - c(c1) * kY(k0) * f1)) / r Else Wc0 = Wc   ' AD
                                If Hc <= 0 Then Hc0 = -(2 * (c(c0) * kX(k1) * f1 - c(c1) * kX(k0) * f0)) / r Else Hc0 = Hc
                Case 5, 8:
                                If (f0 <> 0) And (f1 <> 0) Then
                                If Wc <= 0 Then Wc0 = ((c(c0) * (1 - kY(k0)) - c(c1) * kY(k0)) / f0) / r Else Wc0 = Wc      ' AC,BD
                                If Hc <= 0 Then Hc0 = -((c(c0) * (1 - kX(k0)) - c(c1) * kX(k0)) / f1) / r Else Hc0 = Hc
                                End If
                End Select
                End If
                k1 = k1 + 1: If k1 = 4 Then k0 = k0 + 1: k1 = k0 + 1: c0 = (k0 + 4 - n) Mod 4
            End If
' ��������� ������������ ����������
        ' ��������� ������������ ���������� (����� ����� ������� - ����� ����������� ������)
            If Wc0 <= 0 Or Hc0 <= 0 Then GoTo HandleNext
        ' ��������� ����� ����� �� ������� ��� ����������� ��������
            ' ������ �������� �� ����� ��������
            Call .Transform(rXc, rYc, oXc0, oYc0, Wc0, Hc0) ' oXc0, oYc0 - ���������� ����� �������� ������������ ������� At (�.0)
            dXc0 = oXc - oXc0: dYc0 = oYc - oYc0            ' dXc0, dYc0 - �������� �������� ��������� �.�������� ������������ ���������
        ' �������� ������                                   ' oXc0, oYc0 - ����� � ����� ����������� ��� �������� ��������� ��������� ������
        ' 1 A (0,0)     - ������� A (�.0)                   ' ������ ������� A ��������� �.�. � ����������� �������� ��� ������� �������� (dXc, dYc)
            If (dXc0 < -dErr) Or (dXc0 > (Wh + dErr)) Then GoTo HandleNext
            If (dYc0 < -dErr) Or (dYc0 > (Hh + dErr)) Then GoTo HandleNext
        ' 2 B (1,0)     - ������� B
            Call .Transform(1, 0, oXc0, oYc0, Wc0, Hc0)
            oXc0 = oXc0 + dXc0: If (oXc0 < -dErr) Or (oXc0 > (Wh + dErr)) Then GoTo HandleNext
            oYc0 = oYc0 + dYc0: If (oYc0 < -dErr) Or (oYc0 > (Hh + dErr)) Then GoTo HandleNext
        ' 3 C (1,1)     - ������� C
            Call .Transform(1, 1, oXc0, oYc0, Wc0, Hc0)
            oXc0 = oXc0 + dXc0: If (oXc0 < -dErr) Or (oXc0 > (Wh + dErr)) Then GoTo HandleNext
            oYc0 = oYc0 + dYc0: If (oYc0 < -dErr) Or (oYc0 > (Hh + dErr)) Then GoTo HandleNext
        ' 4 D (0,1)     - ������� D
            Call .Transform(0, 1, oXc0, oYc0, Wc0, Hc0)
            oXc0 = oXc0 + dXc0: If (oXc0 < -dErr) Or (oXc0 > (Wh + dErr)) Then GoTo HandleNext
            oYc0 = oYc0 + dYc0: If (oYc0 < -dErr) Or (oYc0 > (Hh + dErr)) Then GoTo HandleNext
' �������� ��������� �� �������� ������������ �������
        ' �������� � ����������, ���� ������ - ��������� ���������
            If Wc0 * Hc0 > Area Then Wc1 = Wc0: Hc1 = Hc0: dXc = dXc0: dYc = dYc0: Area = Wc1 * Hc1
HandleNext:
        Next s
        Wc = Wc1: Hc = Hc1
End With
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function ObjData_ReadFromFile( _
    FileName As String, _
    ObjKey As String, _
    Optional ObjType As eObjectDataType = eObjectDataNone)
' ������ ���� FileName � ��������� ��� � ������� c_strObjectTable � ������ � ������ ObjectKey � ���� c_strObjectKey (���� OLE Data) � ���� BLOB ������
' ObjKey - ��������� ���� - ������� ��� �������
' FileName - ���� � ������������ �������
' ObjType - ��� ������ ������������ ������
Dim dlgFilePath As FileDialog
Dim strFilePath As String
Dim arrByte() As Byte
Dim rst As DAO.Recordset
Dim Ret As Long
    strFilePath = Trim$(FileName)
    If Len(strFilePath) = 0 Then
' ��������� ���������� ���� ��� ������ ����� � ���������� ���� � ����
        Set dlgFilePath = Application.FileDialog(msoFileDialogFilePicker)
        dlgFilePath.AllowMultiSelect = False
        If Not dlgFilePath.Show Then GoTo HandleExit
        strFilePath = dlgFilePath.SelectedItems.Item(1)
    End If
' ���� �� ������ - ����������� ������� ��� �������
    Do While Trim$(ObjKey) = vbNullString
        ObjKey = InputBox("������� ������� ��� �������")
    Loop
' ������ ������� � ������ ������
    Ret = ByteArray_ReadFromFile(strFilePath, arrByte): If Ret <> NOERROR Then GoTo HandleExit
    Ret = ByteArray_WriteToTable(arrByte, ObjKey): If Ret <> NOERROR Then GoTo HandleExit
HandleExit: ObjData_ReadFromFile = Ret: Exit Function
End Function
Public Function ObjData_WriteToFile( _
    ObjectKey As String, _
    Optional FileName As String) As Long
' ��������� � ���� FileName � ������ �� ������� c_strObjectTable
' �� ������ � ������ ObjectKey ������ ���� c_strObjectKey (���� OLE Data)
' � ���� BLOB ������
' ���������� ���������� ����������� ����
Dim strMessage As String, strTitle As String
Dim dlgFilePath As FileDialog
Dim strFilePath As String ', strFileName As String
Dim ObjectData() As Byte, ObjectType As eObjectDataType, ObjectExt As String
Dim Result As Long: Result = NOERROR
    strFilePath = Trim$(FileName)
    Do While Trim$(ObjectKey) = vbNullString
        ObjectKey = InputBox("������� ������� ��� �������")
    Loop
' ������ ������ �� ������� � ��������� � ����

    Result = ByteArray_ReadFromTable(ObjectKey, ObjectData)
    ObjectType = GetDataTypeBySig(ObjectData, ObjectExt)
    Select Case ObjectType
    Case eObjectDataNone    '
    Case eObjectDataUndef   ' ��� ����������� - �������� PictureData
        strTitle = "������ �������� PictureData?"
        strMessage = "���� ������ �������� PictureData" & vbCrLf & _
            "��� ����� ������������� � �����������" & vbCrLf & _
            "� ���������� ������� ������ ���������� " & vbCrLf & _
            "����� ������ ����� ��������� � ��� ���� � bin ����"
        Select Case MsgBox(strMessage, vbYesNo Or vbQuestion Or vbMsgBoxRtlReading Or vbMsgBoxRight, strTitle)
        Case vbYes
        ' ������������� PictureData � ����������� � ���������
            ' �������� � ObjectData ������� ��������� BMP ������� &h0E
'            If Not p_CreateBitmapFromDibPictureData(ObjectData, ObjectData) Then GoTo HandleExit 'Err.Raise (vbObjectError + 512)
'            ObjectExt = "bmp"
Stop
        Case vbNo
        ' ��������� ��� bin
            ObjectExt = "bin"
        End Select
    End Select
'��������� ���������� ���� ��� ������ ����� � ���������� ���� � ����
Dim strFilename As String
    strFilename = ObjectKey & "." & ObjectExt
    If Len(strFilePath) = 0 Then
        strFilePath = Access.CurrentProject.path & "\" & c_strPathRes
        Set dlgFilePath = Application.FileDialog(msoFileDialogSaveAs)
        With dlgFilePath
            .AllowMultiSelect = False
            .InitialFileName = strFilename
            If Not .Show Then Exit Function           '�����, ���� ������ ��������'
            strFilePath = .SelectedItems.Item(1)
        End With
    End If
'    strFilePath = strFilePath & strFileName
'
    Result = ByteArray_WriteToFile(ObjectData, strFilePath)
HandleExit:  ObjData_WriteToFile = Result: Exit Function
HandleError: Result = Err.Number: Err.Clear: Resume HandleExit
End Function
Public Function GetDataTypeBySig( _
    ByRef ByteStream As Variant, _
    Optional Extention As String, _
    Optional Details As String, _
    Optional Params _
    ) As eObjectDataType
' ���������� ��� ����� �� ��������� �����������
' ByteStream - ������ ������� ���������� ����������������
' Extention - (out) ���������� ����
' Details - (out)
' Params - (out) �������������� ��������� �����
' Params(0) = ���������� �����, ��������� � ������������ �� ����
Dim ObjectData() As Byte
Dim strSignature As String
Dim lngOffset As Long
Dim ParamName As String, paramValue As String
Dim i As Integer, iMax As Integer
Dim Result As eObjectDataType

    Result = eObjectDataUndef ' �������� ��� ����������� ��� Access.Picture BLOB ������
    Extention = "UNDEF": Details = ""
    ObjectData = ByteStream
HandleMediaFormats:
'---------------------------------------------------------------------------------------------------------------------------------
' ����������� � ��.����������
'---------------------------------------------------------------------------------------------------------------------------------
    If p_CmpArrays(ObjectData, StrConv("BM", vbFromUnicode)) Then
' BMP, DIB - Windows (or device-independent) bitmap image (��. https://ru.wikipedia.org/wiki/BMP)
        Result = eObjectDataBMP: Extention = "BMP": i = i + 1:
        'If IsMissing(Params) Then GoTo HandleExit
    Dim Size As Long, Offset As Long, bih As BITMAPINFOHEADER, adibData() As Byte
'[FILEHEADER] � 14-������� ���������.
        'CopyMemory fh, ObjectData(0), &HE
        CopyMemory Size, ObjectData(&H2), 4
        CopyMemory Offset, ObjectData(&HA), 4
'[INFOHEADER]
        CopyMemory bih, ObjectData(&HE), &H28
'[PIXELDATA] - ���������� ������. ���������� � ������ OffBits
        ReDim adibData(0 To Size - Offset) '- 1)
        CopyMemory adibData(0), ObjectData(Offset), Size - Offset
    ElseIf p_CmpArrays(ObjectData, ChrB(&H0) & ChrB(&H0) & ChrB(&H1) & ChrB(&H0)) _
        Or p_CmpArrays(ObjectData, ChrB(&H0) & ChrB(&H0) & ChrB(&H2) & ChrB(&H0)) Then
' ICO,CUR - Windows icon/cursor file
'   [HEADER] � 6-������� ���������.
'   00  Reserved1   WORD    0
'   02  Type        WORD    ��� �����: 1 ��� �������(.ICO) ��� 2 ��� ��������(.CUR). ���� �������� �����������.
        Select Case p_WordRead(ObjectData, &H2)  ' Type
         Case 1
            Result = eObjectDataICO: Extention = "ICO" 'type=1 ��� �������(.ICO)
         Case 2
            Result = eObjectDataCUR: Extention = "CUR" 'type=2 ��� ��������(.CUR)
        End Select
        If Not IsMissing(Params) Then
' E��� ��������� �������������� ��������� - ������:
'   04  Count       WORD    ���������� ����������� � �����, ������� 1.
'        Dim i As Long, iMax As Long ', ii As Byte, iiMax As Byte
            iMax = p_WordRead(ObjectData, &H4) - 1
            ReDim Params(0 To iMax, 0 To 7)
            For i = 0 To iMax
'   [SUBHEADER] � 16-������� ��������.  ���������� �� ������������
'   ���������������� ������ �������������� ������� (16 ����), ��������� ���� �� ������. ���������� ������� ������������ ����� count ���������.
                Params(i, 0) = p_ByteRead(ObjectData, (&H6 + i * &H10 + &H0)) '00  Width       BYTE    ������ ����������� � ������ �� 0 �� 255. ���� 0, �� ������ = 256 �����.
                Params(i, 1) = p_ByteRead(ObjectData, (&H6 + i * &H10 + &H1)) '01  Height      BYTE    ������ ����������� � ������ �� 0 �� 255. ���� 0, �� ������ = 256 �����.
                Params(i, 2) = p_ByteRead(ObjectData, (&H6 + i * &H10 + &H2)) '02  Colors      BYTE    ���������� ������ � ������� ��� TrueColor �.�. =0.
                Params(i, 3) = p_ByteRead(ObjectData, (&H6 + i * &H10 + &H3)) '03  Reserved    BYTE    ���������������. �.�.= 0 (�������� MSDN), ������ ������, �������� .NET (System.Drawing.Icon.Save) �������� 255
                Params(i, 4) = p_WordRead(ObjectData, (&H6 + i * &H10 + &H4)) '04  Planes      WORD    � .ICO - ���������� ����������. = 0 ��� 1.
                '� .CUR - ���.���������� "������� �����" � �������� ���-�� ������ ���� �����������.
                Params(i, 5) = p_WordRead(ObjectData, (&H6 + i * &H10 + &H6)) '06  bpp         WORD    � .ICO - ���-�� ����� �� ������� (bits-per-pixel). ��� �������� ����� ���� 0, ��� ��� ����� ���������� �� ������ ������;
                '��������, ���� ����������� �� �������� � ������� PNG, �����
                '�������������� �� ������ ���������� � ������� ������, � ����� ��� ������ � ������.
                '���� �� ����������� �������� � ������� PNG, �� ��������������� ���������� �������� � ����� PNG.
                '������ ��������� � ���� ���� 0 �� �������������, �.�. ������ ������ ���������� ����������� � ��������� ������� Windows ����������.
                '� .CUR - ����.���������� "������� �����" � �������� ���-�� �������� ���� �����������.
                Params(i, 6) = p_DWordRead(ObjectData, (&H6 + i * &H10 + &H8)) '08  Size        DWORD   ��������� ������ ������ � ������
                Params(i, 7) = p_DWordRead(ObjectData, (&H6 + i * &H10 + &HC)) '12  Offset      DWORD   ��������� ���������� �������� ������ � �����.
            Next i
        End If
    ElseIf p_CmpArrays(ObjectData, StrConv("�PNG", vbFromUnicode) & ChrB(&HD) & ChrB(&HA) & ChrB(&H1A) & ChrB(&HA)) _
     And p_CmpArrays(ObjectData, StrConv("IEND�B`�", vbFromUnicode), , True) Then
' PNG - Portable Network Graphics file (��. https://www.w3.org/TR/PNG/)
'   [HEADER]    8-������� ���������.
'   00  Signature   8*BYTE  HEX: 89 50 4E 47 0D 0A 1A 0A
'   08  Chunks              ����� ������. ������ ���� ������� �� 4 ������
'   [CHUNK] ������� �� ������ �����:
'   00  Length      DWORD   ����� ���� ������ �����
'   04  Type        DWORD   ���� ���� �����: IHDR/IDAT/IEND...
'   08  Data                ���� ������ (���������� �����). �.�. ������� �����
'   nn  CRC         DWORD   ���� � ����������� ����� CRC-32 ��� ����� Type, Data
'   ����������� PNG-���� ������ ��������� ��������� � ��� chunk'a: IHDR, IDAT, IEND
'   [CHUNKIHDR] - Type=IHDR - 12-������� ��������� � ��������� ����� (������ ����)
'   00  Width       DWORD   ������ � �������� (<>0)
'   04  Height      DWORD   ������ � �������� (<>0)
'   08  BitDepth    BYTE    Bit depth. the number of bits per sample or per palette index (not per pixel) (=1,2,4,8,16)
'   09  ColourType  BYTE    Colour type =0,1,2,4,6
'       =0 - Greyscale  BitDepth=1,2,4,8,16 Each pixel is a greyscale sample
'       =1 - Truecolour BitDepth=8,16       Each pixel is an R,G,B triple
'       =2 - IdxColour  BitDepth=1,2,4,8    Each pixel is a palette index; a PLTE chunk shall appear.
'       =4 - Grey&Alpha BitDepth=8,16       Each pixel is a greyscale sample followed by an alpha sample.
'       =6 - True&Alpha BitDepth=8,16       Each pixel is an R,G,B triple followed by an alpha sample.
'   10  Compress    BYTE    =0 (deflate/inflate) Compression method
'   11  Filter      BYTE    =0 (adaptive filtering with five basic filter types) Filter method indicates the preprocessing method applied to the image data before compression.
'   12  Interlace   BYTE    =0 (no interlace) or =1 (Adam7 interlace). Interlace method indicates the transmission order of the image data.
'   [CHUNKIDAT] - Type=IDAT - nn-������� ��������� � ��������, ����������, �����������.
'       � ����� �.�. ���� �� ���� IDAT ����.
'       �������� ������ �����������, ���������� � ���������� ���������� Compress � Filter
'       ���� ����� ��������� ��������� IDAT ������, � ����� ������ ��� ������������� ��������������� � �� ����������� ������� �������
'       ��� ��������� ����������� ���������� ��������������� ���������� IDAT ����� � ��������� ���������� ������� Compress
'   [CHUNKIEND] - Type=IEND - 0-������� ��������� � ����������� ���� (��������� � �����)
        Result = eObjectDataPNG: Extention = "PNG"
    ElseIf p_CmpArrays(ObjectData, StrConv("GIF87a", vbFromUnicode)) _
     And p_CmpArrays(ObjectData, ChrB(&H0) & ChrB(&H3B), , True) Then
' GIF - Graphics interchange format file
        Result = eObjectDataGIF: Extention = "GIF"
    ElseIf p_CmpArrays(ObjectData, StrConv("GIF89a", vbFromUnicode)) _
     And p_CmpArrays(ObjectData, ChrB(&H0) & ChrB(&H3B), , True) Then
        Result = eObjectDataGIF: Extention = "GIF"
    ElseIf p_CmpArrays(ObjectData, StrConv("���", vbFromUnicode)) _
     And p_CmpArrays(ObjectData, StrConv("��", vbFromUnicode), , True) Then
' JPEG/JFIF graphics file
        Result = eObjectDataJPG: Extention = "JPG": i = i + 1
        If Not IsMissing(Details) Then
    ' ���� �������� �������������� �������� ��������� ��� ���������
        ' ��������� ���� ���������� ��� JPEG �����������
            lngOffset = &H6 ' �� �������� 6 �������������� ������������
            If p_CmpArrays(ObjectData, ChrB(&HDB), 3) Then
        ' DB - Samsung D807 JPEG file
                Details = "Samsung D807 JPEG file"
            ElseIf p_CmpArrays(ObjectData, ChrB(&HE0), 3) _
             And p_CmpArrays(ObjectData, StrConv("JFIF", vbFromUnicode) & ChrB(&H0), lngOffset) Then
        ' E0 � Standard JPEG/JFIF file.
                Details = "Standard JPEG/JFIF file"
            ElseIf p_CmpArrays(ObjectData, ChrB(&HE1), 3) _
             And p_CmpArrays(ObjectData, StrConv("Exif", vbFromUnicode) & ChrB(&H0), lngOffset) Then
        ' E1 � Standard JPEG/Exif file. '   FF D8 FF E1 xx xx 45 78 69 66 00
            ' Digital camera JPG using Exchangeable Image File Format (EXIF)
            ' See "Using Extended File Information (EXIF) File Headers in Digital"
            ' Evidence Analysis" (P. Alvarez, IJDE, 2(3), Winter 2004) and ExifTool Tag Names
                Details = "Standard JPEG/EXIF file"
            ElseIf p_CmpArrays(ObjectData, ChrB(&HE2), 3) Then
        ' E2 � Canon EOS-1D JPEG file.
                Details = "Canon EOS-1D JPEG file"
            ElseIf p_CmpArrays(ObjectData, ChrB(&HE3), 3) Then
        ' E3 � Samsung D500 JPEG file.
                Details = "Samsung D500 JPEG file"
            ElseIf p_CmpArrays(ObjectData, ChrB(&HE8), 3) _
             And p_CmpArrays(ObjectData, StrConv("SPIFF", vbFromUnicode) & ChrB(&H0), lngOffset) Then
        ' E8 � Still Picture Interchange File Format (SPIFF).
                Details = "Standard JPEG/SPIFF file"
            End If
        End If
    ElseIf p_CmpArrays(ObjectData, StrConv("I I", vbFromUnicode)) Then
' TIF,TIFF - Tagged Image File Format file
        Result = eObjectDataTIF: Extention = "TIF"
    ElseIf p_CmpArrays(ObjectData, StrConv("II*", vbFromUnicode) & ChrB(&H0)) Then
    ' little endian, i.e., LSB first in the byte; Intel
        Result = eObjectDataTIF: Extention = "TIF" ': Details = "BigTIFF file (>4 GB)"
    ElseIf p_CmpArrays(ObjectData, StrConv("MM", vbFromUnicode) & ChrB(&H0) & ChrB(&H2A)) Then
    ' big endian, i.e., LSB last in the byte; Motorola
        Result = eObjectDataTIF: Extention = "TIF" ': Details = "BigTIFF file (>4 GB)"
    ElseIf p_CmpArrays(ObjectData, StrConv("MM", vbFromUnicode) & ChrB(&H0) & ChrB(&H2B)) Then
    ' BigTIFF files; Tagged Image File Format files >4 GB
        Result = eObjectDataTIF: Extention = "TIF": Details = "BigTIFF file (>4 GB)"
'    ElseIf p_CmpArrays(ObjectData, StrConv("����", vbFromUnicode)) _
'     Or p_CmpArrays(ObjectData, StrConv("����", vbFromUnicode), &H1E) Then
'' EPS - Adobe encapsulated PostScript file
'    ' If this signature is not at the immediate beginning of the file,
'    ' it will occur early in the file, commonly at byte offset 30 [0x1E])
'        Result = eObjectDataEPS: Extention = "EPS"
'    ElseIf p_CmpArrays(ObjectData, StrConv("%!PS-Adobe-3.0 EPSF-3 ", vbFromUnicode)) Then
'' EPS - Encapsulated PostScript file
'        Result = eObjectDataEPS: Extention = "EPS"
'    ElseIf p_CmpArrays(ObjectData, StrConv("�WPC", vbFromUnicode)) Then
'' WPG - WordPerfect text and graphics file
'        Result = eObjectDataWPG: Extention = "WPG"
'    ElseIf p_CmpArrays(ObjectData, ChrB(&H1) & ChrB(&H0) & ChrB(&H0) & ChrB(&H0)) Then
'' EMF - Extended (Enhanced) Windows Metafile Format,
'' printer spool file (0x18-17 & 0xC4-36 is Win2K/NT; 0x5C0-1 is WinXP)
'        Result = eObjectDataEMF: Extention = "EMF" ':  i = i + 1
'    ElseIf p_CmpArrays(ObjectData, StrConv("��ƚ", vbFromUnicode)) Then
'' WMF - Windows graphics metafile
'        Result = eObjectDataWMF: Extention = "WMF"
'    ElseIf p_CmpArrays(ObjectData, StrConv("MSWIM", vbFromUnicode)) Then
'' Microsoft Windows Imaging Format file
'        Result = eObjectDataWIM: Extention = "WIM"
    ElseIf p_CmpArrays(ObjectData, StrConv("CWS", vbFromUnicode)) Then
' Macromedia Shockwave Flash player file
' See SWF File Format Specification: http://wwwimages.adobe.com/content/dam/Adobe/en/devnet/swf/pdf/swf-file-format-spec.pdf
     ' CWS - zlib compressed SWF 6 and later
        Result = eObjectDataSWF: Extention = "SWF": Details = "zlib compressed SWF 6 and later"
    ElseIf p_CmpArrays(ObjectData, StrConv("FWS", vbFromUnicode)) Then
     ' FWS - uncompressed SWF
        Result = eObjectDataSWF: Extention = "SWF": Details = "uncompressed SWF"
    ElseIf p_CmpArrays(ObjectData, StrConv("ZWS", vbFromUnicode)) Then
     ' ZWS - LZMA compresse SWF 13 and later
        Result = eObjectDataSWF: Extention = "SWF": Details = "LZMA compresse SWF 13 and later"
     Else
' ���� �� ���� ������ �� ������� ��������� ���������
        GoTo HandleOfficeFormats
    End If
    GoTo HandleExit ' ���� ���������
'---------------------------------------------------------------------------------------------------------------------------------
' ���������:
'---------------------------------------------------------------------------------------------------------------------------------
HandleOfficeFormats:
    If p_CmpArrays(ObjectData, StrConv("<?xml version=""1.0""", vbFromUnicode)) Then
' User Interface Language file (XML,XUL)
    ' ���������� ��������� "<?xml version=""1.0""?>"
    ' �� MS Office ��������� � XML �����: <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        Result = eObjectDataXML: Extention = "XML" '
    ElseIf p_CmpArrays(ObjectData, StrConv("{\rtf1", vbFromUnicode)) _
        And p_CmpArrays(ObjectData, StrConv("\par }}", vbFromUnicode), , True) Then
' Rich text format word processing file
        Result = eObjectDataRTF: Extention = "RTF"
    ElseIf p_CmpArrays(ObjectData, StrConv("ۥ-", vbFromUnicode) & ChrB(&H0)) Then
' Word 2.0 file
        Result = eObjectDataDOC20: Extention = "DOC"
    ElseIf p_CmpArrays(ObjectData, StrConv("��ࡱ�", vbFromUnicode)) Then
' Microsoft Office document
' An Object Linking and Embedding (OLE) Compound File (CF) (i.e., OLECF) file format,
' known as Compound Binary File format by Microsoft, used by Microsoft Office 97-2003
' applications (Word, Powerpoint, Excel, Wizard).
    ' [See also Excel, Outlook, PowerPoint, and Word "subheaders" at byte offset 512 (0x200).]
    ' There appear to several subheader formats and a dearth of documentation.
    ' There have been reports that there are different subheaders for Windows and Mac versions of MS Office but I cannot confirm that.]
    ' Password-protected DOCX, XLSX, and PPTX files also use this signature those files are saved as OLECF files.
    ' [Note the similarity between D0 CF 11 E0 and the word "docfile"!]
        lngOffset = &H200 ' MS Office subheaders [512 (0x200) byte offsets]
        If p_CmpArrays(ObjectData, StrConv("��", vbFromUnicode) & ChrB(&H0), lngOffset) Then
' Word document subheader (MS Office)
            Result = eObjectDataDOC: Extention = "DOC"
            GoTo HandleExit
        ElseIf p_CmpArrays(ObjectData, ChrB(&H9) & ChrB(&H8) & ChrB(&H10) & ChrB(&H0) & ChrB(&H0) & ChrB(&H6) & ChrB(&H5) & ChrB(&H0), lngOffset) _
         Or p_CmpArrays(ObjectData, StrConv("����", vbFromUnicode), lngOffset) Then
        ' 09 08 10 00 00 06 05 00
        ' FD FF FF FF nn 00 - Excel spreadsheet subheader (MS Office)
        ' FD FF FF FF nn 02 - Excel spreadsheet subheader (MS Office)
        ' FD FF FF FF 20 00 00 00 - Excel spreadsheet subheader (MS Office) or Developer Studio File Workspace Options subheader (MS Office)
    ' Excel spreadsheet subheader (MS Office)
            Result = eObjectDataXLS: Extention = "XLS"
        Else
        ' ������ MS Office 97-2003 ��������
        End If
    ElseIf p_CmpArrays(ObjectData, ChrB(&H0) & ChrB(&H1) & ChrB(&H0) & ChrB(&H0) & StrConv("Standard Jet DB", vbFromUnicode)) Then
' Standard Jet db MDB         Microsoft Access file
        Result = eObjectDataMDB: Extention = "MDB"
    ElseIf p_CmpArrays(ObjectData, ChrB(&H0) & ChrB(&H1) & ChrB(&H0) & ChrB(&H0) & StrConv("Standard ACE DB", vbFromUnicode)) Then
' Microsoft Access 2007 file
        Result = eObjectDataACCDB: Extention = "ACCDB"
    ElseIf p_CmpArrays(ObjectData, StrConv("%PDF", vbFromUnicode)) _
     And p_CmpArrays(ObjectData, ChrB(&HA) & StrConv("%%EOF", vbFromUnicode), , True) Then
' PDF, FDF, AI  Adobe Portable Document Format, Forms Document Format, and Illustrator graphics files
    ' There may be multiple end-of-file marks within the file. When carving, be sure to get the last one.
    ' Trailer: 0A 25 25 45 4F 46
        Result = eObjectDataPDF: Extention = "PDF"
     ElseIf p_CmpArrays(ObjectData, StrConv("%PDF", vbFromUnicode)) _
      And p_CmpArrays(ObjectData, ChrB(&HA) & StrConv("%%EOF", vbFromUnicode) & ChrB(&HA), , True) Then
    ' Trailer: 0A 25 25 45 4F 46 0A
        Result = eObjectDataPDF: Extention = "PDF"
     ElseIf p_CmpArrays(ObjectData, StrConv("%PDF", vbFromUnicode)) _
      And p_CmpArrays(ObjectData, ChrB(&HD) & StrConv("%%EOF", vbFromUnicode) & ChrB(&HD), , True) Then
    ' Trailer: 0D 25 25 45 4F 46 0D
        Result = eObjectDataPDF: Extention = "PDF"
     ElseIf p_CmpArrays(ObjectData, StrConv("%PDF", vbFromUnicode)) _
      And p_CmpArrays(ObjectData, ChrB(&HD) & ChrB(&HA) & StrConv("%%EOF", vbFromUnicode) & ChrB(&HD) & ChrB(&HA), , True) Then
    ' Trailer: 0D 0A 25 25 45 4F 46 0D 0A
        Result = eObjectDataPDF: Extention = "PDF"
    ElseIf p_CmpArrays(ObjectData, StrConv("%!PS", vbFromUnicode)) Then
' PostScript document
        Result = eObjectDataPS: Extention = "PS"
    ElseIf p_CmpArrays(ObjectData, StrConv("AT&TFORM", vbFromUnicode)) Then
' DjVu document djvu, djv
    '41 54 26 54 46 4F 52 4D nn nn nn nn 44 4A 56
    'The following byte is either 55 (U) for single-page or 4D (M) for multi-page documents.     0
        Result = eObjectDataDJV: Extention = "djvu"
'   ElseIf p_CmpArrays(ObjectData, StrConv("PK", vbFromUnicode) & ChrB(&H3) & ChrB(&H4)) _
'        And p_CmpArrays(ObjectData, StrConv("PK", vbFromUnicode) & ChrB(&H5) & ChrB(&H6), 18, True) Then
'' Microsoft Office Open XML Format (OOXML) Document
'    ' 50 4B 03 04               Microsoft Office Open XML Format (OOXML) Document
'    ' 50 4B 03 04 14 00 06 00   MS Office 2007 documents
'    ' 50 4B 03 04 14 00 00 00
'' Trailer: Look for 50 4B 05 06 followed by 18 additional bytes at the end of the file.
'Debug.Print "Microsoft Office Open XML Format (OOXML) Document"
'    ' There is no subheader for MS OOXML files as there is with DOC, PPT, and XLS files.
'    ' To better understand the format of these files, rename any OOXML file to have a .ZIP extension and then unZIP the file;
'    ' look at the resultant file named [Content_Types].xml to see the content types.
'    ' In particular, look for the <Override PartName= tag, where you will find word, ppt, or xl, respectively.
''        ParamName = "Override PartName" ' ������ �������� ��������� �� [Content_Types].xml
''        Select Case ParamValue
''         Case "word"
''            Result = eObjectDataDOCX:Extention = "DOCX"
''         Case "xl":     Ext = "XLSX"
''            Result = eObjectDataXLSX:Extention = "XLSX"
''         Case "ppt":    Ext = "PPTX"
''            Result = eObjectDataPPTX:Extention = "PPTX"
''         Case Else
''            GoTo HandleExit
''        End Select
     Else
        GoTo HandleArchiveFormats
    End If
    GoTo HandleExit
'---------------------------------------------------------------------------------------------------------------------------------
' ������
'---------------------------------------------------------------------------------------------------------------------------------
HandleArchiveFormats:
    If p_CmpArrays(ObjectData, StrConv("Rar!", vbFromUnicode) & ChrB(&H7) & ChrB(&H0)) Then
' WinRAR compressed archive
    ' 52 61 72 21 1A 07 00 - RAR archive version 1.50 onwards
        Result = eObjectDataRAR: Extention = "RAR": Details = "RAR archive version 1.50 onwards"
    ElseIf p_CmpArrays(ObjectData, StrConv("Rar!", vbFromUnicode) & ChrB(&H7) & ChrB(&H1) & ChrB(&H0)) Then
    ' 52 61 72 21 1A 07 01 00 - RAR archive version 5.0 onwards
        Result = eObjectDataRAR: Extention = "RAR": Details = "RAR archive version 5.0 onwards"
    ElseIf p_CmpArrays(ObjectData, StrConv("7z��'", vbFromUnicode)) Then
' 7-Zip compressed file
        Result = eObjectData7Z: Extention = "7Z"
        GoTo HandleExit
    ElseIf p_CmpArrays(ObjectData, StrConv("PK", vbFromUnicode)) Then
' ZIP archive
        Result = eObjectDataZIP: Extention = "ZIP"
    ' ���� ��������� �������������� ���������
        If p_CmpArrays(ObjectData, ChrB(&H3) & ChrB(&H4), 2) Then
    ' PKZIP archive file
        ' http://members.tripod.com/~petlibrary/ZIP.HTM
        ' http://www.pkware.com/documents/casestudies/APPNOTE.TXT
        ' Trailer: (filename PK 17 characters ...)
        ' 50 4B 03 04 is used to show filename structure
        ' Local file header:
            'local file header signature     4 bytes  (0x04034b50)
            Details = "PKZIP archive file"
'            If Not IsMissing(Params) Then
'                ReDim Params(0 To 9)
'Stop
'                Params(0) = p_WordRead(ObjectData, 4) 'version needed to extract   2 bytes
'                Params(1) = p_WordRead(ObjectData, 6) 'general purpose bit flag    2 bytes
'                Params(2) = p_WordRead(ObjectData, 8) 'compression method  2 bytes
'                Params(3) = p_WordRead(ObjectData, 10) 'last mod file time 2 bytes
'                Params(4) = p_WordRead(ObjectData, 12) 'last mod file date 2 bytes
'                Params(5) = p_DWordRead(ObjectData, 14) 'crc-32    4 bytes
'                Params(6) = p_DWordRead(ObjectData, 18) 'compressed size   4 bytes
'                Params(7) = p_DWordRead(ObjectData, 22) 'uncompressed size 4 bytes
'                Params(8) = p_WordRead(ObjectData, 24) 'file name length   2 bytes
'                Params(9) = p_WordRead(ObjectData, 26) 'extra field length 2 bytes
'                'file name (variable size)
'                'extra field (variable size)
'            End If
    'PKLITE archive
        ' Offset:  &H1E
        ' Bytes:    50 4B 4C 49 54 45   ' "PKLITE"
    'PKSFX self-extracting archive
        ' Offset:  &H20E
        ' Bytes:    50 4B 53 70 58      ' "PKSFX"
    'WinZip compressed archive
        ' Offset:  &H71E0
        ' Bytes:    57 69 6E 5A 69 70   ' "WinZip"
        ElseIf p_CmpArrays(ObjectData, ChrB(&H5) & ChrB(&H6), 2) Then
    ' PKZIP empty and multivolume archive file, respectively
        ' 50 4B 05 06 is used to show  the  end  of the central directory
            Details = "PKZIP empty and multivolume archive file": i = i + 1
        ElseIf p_CmpArrays(ObjectData, ChrB(&H7) & ChrB(&H8), 2) Then
    ' PKZIP empty and multivolume archive file, respectively
            Details = "PKZIP empty and multivolume archive file": i = i + 1
        ElseIf p_CmpArrays(ObjectData, ChrB(&H1) & ChrB(&H2), 2) Then
        ' 50 4B 01 02 is used  to  signify  the  beginning  of  the  central directory while the byte sequence
'
        End If
     Else
        GoTo HandleOtherFormats
    End If
    GoTo HandleExit
'---------------------------------------------------------------------------------------------------------------------------------
' ������ �������
'---------------------------------------------------------------------------------------------------------------------------------
HandleOtherFormats:
' EF BB BF - Unicode Text
'---------------------------------------------------------------------------------------------------------------------------------
HandleExit:  GetDataTypeBySig = Result: Exit Function
HandleError: Err.Clear: Result = eObjectDataNone: Resume HandleExit
End Function
Public Function ByteArray_ReadFromFile( _
    ByVal FileName As String, ByRef ByteArray() As Byte, _
    Optional BytesRead As Long _
    ) As Long
'    Optional ByRef ObjectType As eObjectDataType = eObjectDataUndef, _
'    Optional ByRef ObjectTypeExtention As String, _
'    Optional ByRef ObjectTypeDetails As String _
'    ) As Long
Const c_strProcedure = "ByteArray_ReadFromFile"
' ������ ���� � �������� ������
Dim nFile As Integer
Dim Result As Long:   Result = 0
    On Error GoTo HandleError
'    ObjectTypeExtention = "UNDEF": ObjectTypeDetails = ""
'    FileLen FileName '������, ���� ����� �� ����������
    If Len(FileName) > 0 Then GoTo HandleOpenFile
' ��������� ���������� ���� ��� ������ ����� � ���������� ���� � ����
Dim dlgFilePath As Object: Set dlgFilePath = Application.FileDialog(msoFileDialogFilePicker)
    dlgFilePath.AllowMultiSelect = False: If Not dlgFilePath.Show Then Err.Raise vbObjectError + 512
    FileName = dlgFilePath.SelectedItems.Item(1)
HandleOpenFile: nFile = FreeFile(): Open FileName For Binary Access Read As #nFile
    BytesRead = LOF(nFile): If BytesRead <= 0 Then Err.Raise 75& ' Path/File access error
' ������ ��� �������� ������
    ReDim ByteArray(1 To BytesRead)
    Get #nFile, , ByteArray
    'ObjectType = GetDataTypeBySig(ByteArray, ObjectTypeExtention, ObjectTypeDetails)
    Result = NOERROR
HandleExit: If nFile > 0 Then Close #nFile
    ByteArray_ReadFromFile = Result: Exit Function
HandleError: Result = Err.Number: Err.Clear: Resume HandleExit
End Function
Public Function ByteArray_ReadFromTable( _
    ByVal ObjectKey As String, ByRef ByteArray() As Byte, _
    Optional dbs As DAO.Database, Optional wks As DAO.Workspace _
    ) As Long ' , _
    Optional ByRef ObjectDesc As String, Optional ByRef ObjectComm As String, _
    Optional ByRef ObjectType As eObjectDataType, _
' ������ ������ ������� SysObjs � �������� ������
Const c_strProcedure = "ByteArray_ReadFromTable"
' ObjectKey - ��������� ����� ��������� �������
' ByteArray - �������� ������ � ������� ���������� ������ �������
' ObjectDesc, ObjectComm - ��������� ������� �� ������� (��������/����������� ...)
' ObjectType -
' ObjectTypeExtention -
' ObjectTypeDetails -
' dbs, wks - ������ �� ���� ������ ������� ������������ � ������� ��������� ������� ��������
'-------------------------
Dim Result As Long: Result = NOERROR
'Dim ObjectTypeExtention As String, ObjectTypeDetails As String
'    ObjectTypeExtention = "UNDEF": ObjectTypeDetails = ""
    On Error GoTo HandleError
' �������� DAO �����������
    If dbs Is Nothing And wks Is Nothing Then
        Set dbs = CurrentDb ':Set wks = DBEngine.Workspaces(0)
    ElseIf Not dbs Is Nothing Then
        'Set wks = DBEngine.Workspaces(0)
    ElseIf dbs Is Nothing And wks.Databases.Count > 0 Then
        Set dbs = wks.Databases(0)
    Else
        Set dbs = CurrentDb ': Set wks = DBEngine.Workspaces(0)
    End If
Dim strSQL As String, strWhere As String
Dim rst As DAO.Recordset
    ' ��������� ObjectKey
    If IsNumeric(ObjectKey) Then
    ' ���� ����� - ���� �� ���� ID,
        strWhere = c_strKey & sqlEqual & ObjectKey
    ElseIf Len(ObjectKey) > 0 Then
    ' ���� ������ - �� ObjectKey
        strWhere = c_strObjectKey & sqlEqual & """" & ObjectKey & """"
    Else
        Err.Raise vbObjectError + 512
    End If
    strSQL = sqlSelectAll & c_strObjectTable & sqlWhere & strWhere
    Set rst = dbs.OpenRecordset(strSQL, dbOpenForwardOnly)
    With rst
        If .BOF And .EOF Then Err.Raise vbObjectError + 512
        ByteArray = .Fields(c_strObjectData).Value
'        ObjectName = Nz(.Fields(c_strObjectKey).Value, vbNullString)
'        ObjectDesc = Nz(.Fields(c_strObjectDesc).Value, vbNullString)
'        ObjectComm = Nz(.Fields(c_strObjectComm).Value, vbNullString)
    End With
'    ObjectType = GetDataTypeBySig(ByteArray, ObjectTypeExtention, ObjectTypeDetails)
    Result = NOERROR
HandleExit:  ByteArray_ReadFromTable = Result: Exit Function
HandleError: Result = Err.Number: Err.Clear: Resume HandleExit
End Function
Public Function ByteArray_ReadFromApp( _
    ByVal DibId As Variant, ByRef ByteArray() As Byte, _
    Optional ByRef Transparency) As Long ', _
    Optional RetType As Long = 0) As Long ', Optional UseLargeBitmapSize As Boolean = False) As Long ', Optional ReturnAsBitmap As Boolean = False) As Long
' ���������� ������ ������������ � ���������� (����� ���������� � ����������� �� ������)
'-------------------------
' DibId         - ��������� ����� ��������� ������� (���������� � �������� c_strAppDibIdPref)
' ByteArray     - �������� ������ � ������� ���������� ������ �������
' Transparency  - ������ �������� ������������ ������
'' RetType       - 0 ���������� DIB, 1 ���������� BITMAP
'-------------------------
Dim Result As Long: Result = NOERROR
    On Error GoTo HandleError
    If Left(DibId, Len(c_strAppDibIdPref)) <> c_strAppDibIdPref Then Err.Raise 13
Dim lDibId As Long: lDibId = Mid(DibId, Len(c_strAppDibIdPref) + 1) ': If Not IsNumeric(lDibId) Then Err.Raise 13
Const cColors = 16, cPalSize = &H40&    ' Palette = 16 colors * 4 bytes per color
Dim lColSize As Long                    ' Colors = 4bit color => 2 Color index per byte
Const bLarge = 0                        ' must be 0. 1 is for large icons but i don't know how it works
    Select Case bLarge
    Case 0: lColSize = &H80&            ' Small DIB (16x16x4bit)
    Case 1: lColSize = &H120&           ' Large DIB (24x24x4bit)
    Case Else: Err.Raise vbObjectError + 512
    End Select
Dim aData() As Byte, ldibSize As Long
    ldibSize = (BITMAPINFOHEADERSIZE + cPalSize + lColSize): ReDim aData(0 To (ldibSize - 1))
    If accGetTbDIB(lDibId, bLarge, aData()) = 0 Then Err.Raise vbObjectError + 512  ' get DIB with iternal access function
'    Select Case RetType
'    Case 1:    Dim lfilSize As Long: lfilSize = BITMAPFILEHEADERSIZE + ldibSize     ' return BMP
'               ReDim ByteArray(0 To (lfilSize - 1))
'               CopyMemory ByteArray(0), &H4D42, 2                                   ' BM
'               CopyMemory ByteArray(2), lfilSize, 4                                 ' bfSize
'               CopyMemory ByteArray(&HA), &H76&, 4                                  ' bfOffBits
'               CopyMemory ByteArray(BITMAPFILEHEADERSIZE), aData(0), ldibSize       ' DIB data
'    Case Else: ByteArray() = aData()                                                ' return DIB
'    End Select
'' Replace back color with fully transparent 'Const cBack = &HC0C0C0  ' value of back color in DIBs
Dim cBack As Long       '
Const iBack1 = &HA      '&HA - 1st BackColor Palette Index
Const iBack2 = &HD      '&HD - 2nd BackColor Palette Index
#If ObjectDataType = 0 Then     'FI
    cBack = &H0&
    CopyMemory aData(BITMAPINFOHEADERSIZE + iBack1 * LenB(cBack)), cBack, LenB(cBack)
    CopyMemory aData(BITMAPINFOHEADERSIZE + iBack2 * LenB(cBack)), cBack, LenB(cBack)
#End If
    ByteArray() = aData()
' Create transparency table
    If IsMissing(Transparency) Then Exit Function
Dim aTrans() As Byte: ReDim aTrans(0 To cColors - 1) As Byte: FillMemory aTrans(0), cColors, &HFF: aTrans(iBack1) = 0: aTrans(iBack2) = 0
' 0 - is for transparent color, &hff - opaque
'    'aTrans() = Array(&HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &H0, &HFF, &HFF, &H0, &HFF, &HFF)

    Transparency = aTrans
HandleExit:  ByteArray_ReadFromApp = Result: Exit Function
HandleError: Result = Err.Number: Err.Clear: Resume HandleExit
End Function
Public Function ByteArray_WriteToFile( _
    ByteArray() As Byte, _
    ByVal FileName As String, _
    Optional Overwrite As Boolean = True _
    ) As Long
' ��������� �������� ������ � ����
Const c_strProcedure = "ByteArray_WriteToFile"
'-------------------------
' ByteArray     - �������� ������ ���������� ������ ������� �.�. ���������
' FileName      - ��� ����� � ������� ����� ��������� ������
' Overwrite     - ���� ������������ ����� �� ����������� ������������ ���� ��� ������ �����
' BytesWrite    -
'-------------------------
Dim Result As Long: Result = False
On Error GoTo HandleError

Dim nFile As Integer
Dim strFilePath As String, dlgFilePath As Object
    If Len(FileName) = 0 Then
        strFilePath = Access.CurrentProject.path & "\" & c_strPathRes
        Set dlgFilePath = Application.FileDialog(msoFileDialogSaveAs)
        With dlgFilePath
            .AllowMultiSelect = False
            '.InitialFileName = strFilePath
            If Not .Show Then GoTo HandleExit    '�����, ���� ������ ��������'
            FileName = .SelectedItems.Item(1)
        End With
    End If
    ' Delete if exists
    If dir(FileName, vbHidden + vbSystem) <> vbNullString Then
    ' if exists
        If Overwrite Then
        ' delete if exists
            SetAttr FileName, vbNormal: Kill FileName
        Else
        ' add TimeStamp to FileName
            FileName = p_NewFileName(FileName)
        End If
    End If
    ' Open new for write binary data
    nFile = FreeFile
    Open FileName For Binary Access Write As #nFile
    Put #nFile, , ByteArray: Result = LOF(nFile)
HandleExit:
    If nFile > 0 Then Close #nFile
    ByteArray_WriteToFile = Result
    Exit Function
HandleError:
'    Select Case Err.Number
'    Case 70:  If Overwrite Then Err.Clear:  Resume Next     ' Permisission denied
'    Case Else:
'    End Select
    Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_NewFileName(Optional FileName As String) As String
Dim Result As String:
Dim i As Long, j As Long
    On Error GoTo HandleError
    i = InStrRev(FileName, "\")
    j = InStrRev(FileName, "."): If j < i Then j = 0 Else j = Len(FileName) - j + 1
    Result = Left(FileName, i)
    Result = Result & Mid(FileName, i + 1, Len(FileName) - i - j)
    Result = Result & "_" & Format$(Now, "yyyymmdd_hhnnss") ' add timestamp
    Result = Result & Right(FileName, j)
HandleExit:  p_NewFileName = Result: Exit Function
HandleError: Result = FileName: Err.Clear: Resume HandleExit
End Function
Public Function ByteArray_WriteToTable( _
    ByRef ByteArray() As Byte, _
    ByVal ObjectKey As String, _
    Optional ByRef ObjectDesc As String, _
    Optional ByRef ObjectComm As String _
    ) As Long
' ���������� � ������� SysObjs �������� ������
Const c_strProcedure = "ByteArray_WriteToTable"
Dim Result As Long: Result = NOERROR
Dim strSQL As String, strWhere As String
Dim rst As DAO.Recordset
    On Error GoTo HandleError
    Result = 0
    ' ��������� ObjectKey
    If IsNumeric(ObjectKey) Then
    ' ���� ����� - ���� �� ���� ID,
'        strWhere = c_strKey & c_strObjectKey & sqlEqual & ObjectKey
        Err.Raise vbObjectError + 512
    ElseIf Len(ObjectKey) > 0 Then
    ' ���� ������ - �� ObjectKey
        strWhere = c_strObjectKey & sqlEqual & """" & ObjectKey & """"
    Else
        Err.Raise vbObjectError + 512
    End If
    strSQL = sqlSelectAll & c_strObjectTable & sqlWhere & strWhere
    Set rst = CurrentDb.OpenRecordset(strSQL) ', dbOpenForwardOnly)
    With rst
        If .BOF And .EOF Then
        ' ���� ��� ������ � ����� ������� ������ - �������
            .AddNew
            .Fields(c_strKey).Value = DMax(c_strKey, c_strObjectTable) + 1
            .Fields(c_strObjectKey).Value = ObjectKey
         Else
         ' ���� ���� - ��������
            .Edit
        End If
    ' ��������� ������ � ����
        .Fields(c_strObjectData).Value = ByteArray
        .Update
    End With
HandleExit:  ByteArray_WriteToTable = Result: Exit Function
HandleError: Result = Err.Number: Err.Clear: Resume HandleExit
End Function
#If ObjectDataType = 0 Then     'FI
Public Function PictureData_CreateOleObject(fiPict As LongPtr, Optional BackColor As Long)
' ������� � ���������� OLE Object Picture �� FIBITAMAP
'-------------------------
' fiPict - ��������� �� ����������� FreeImage Bitmap, ������� �.�. ������������ � OLE Object
' BackColor - ���� ��� ������ ����������� ���� �������� (����� ����� ���������� ���� ���� �� ������� ��� �.�. ��������� ����� ����������� ������������)
'-------------------------
#ElseIf ObjectDataType = 1 Then 'LV
Public Function PictureData_CreateOleObject(lvPict As clsPictureData, Optional BackColor As Long)
' ������� � ���������� OLE Object Picture �� clsPictureData
'-------------------------
' lvPict - ��������� �� ����������� clsPictureData, ������� �.�. ������������ � OLE Object
' BackColor - ���� ��� ������ ����������� ���� �������� (����� ����� ���������� ���� ���� �� ������� ��� �.�. ��������� ����� ����������� ������������)
'-------------------------
#End If             'ObjectDataType
'-------------------------
' !!! ������ ��� ������������� !!!
'-------------------------
' ���������� �������� ������ ������ � ������� OLE Object (PBrush 24 bit wo Alpha) ��������� ��� ������� � ObjectData ��������� Access
' ��������� ��� �������������� �� �������� - ��� �������� - �� ����
'-------------------------
Dim Result As Long: Result = NOERROR
    On Error GoTo HandleError
#If ObjectDataType = 0 Then     'FI
    If fiPict = 0 Then Err.Raise vbObjectError + 512
#ElseIf ObjectDataType = 1 Then 'LV
    If lvPict.Handle = 0 Then Err.Raise vbObjectError + 512
#End If             'ObjectDataType
' ������ ���������� ������������ ������ ���
' translate ole colors
Dim vbColor As Long: If (OleTranslateColor(BackColor, 0, vbColor) <> 0) Then vbColor = BackColor
Dim lrgbColor As Long:  lrgbColor = ((vbColor And &HFF000000) Or ((vbColor And &HFF&) * &H10000) Or ((vbColor And &HFF00&)) Or ((vbColor And &HFF0000) \ &H10000))

Dim aData() As Byte
#If ObjectDataType = 0 Then     'FI
' composite �������� ������ �� 8 bit � 32 bit FIBITMAP
Dim fiTemp As LongPtr: fiTemp = fiPict '
' Load FreeImage Library to memory
If Not FreeImage_IsLoaded Then FreeImage_LoadLibrary
    If FreeImage_IsTransparent(fiPict) Then fiPict = FreeImage_Composite(fiTemp, 0, lrgbColor): FreeImage_Unload (fiTemp)
    aData = FreeImage_GetPictureData(fiPict, True)
#ElseIf ObjectDataType = 1 Then 'LV
Dim lvTemp As New clsPictureData: lvTemp.InitializeDIB lvPict.Width, lvPict.Height, lrgbColor
    lvPict.Render 0, DestHostDIB:=lvTemp: Set lvPict = lvTemp
    Call lvPict.SaveToStream_PictureData(aData, picDIB)
#End If             'ObjectDataType
    PictureData_CreateOleObject = p_OleObject_GetPBrush(aData)
HandleExit:  Exit Function
HandleError: Result = Err.Number: Err.Clear: Resume HandleExit
End Function

Public Function OleObject_BackColor(Color As Long, Optional Width As Long = 1, Optional Height As Long = 1)
' ������� �������� ������� 1x1 ��������� ����� � ���� OLE Object Data
On Error GoTo HandleError
'-------------------------
' ��� ��� ������� ������ ����� ����������� � ����������� OLE data ������� ��� �����.
' ������� ��������:
' - �������� ����������� ����������� ����� ���������, �� ����� �������� � � ������
'-------------------------
' ����� ������������ ��� ��������� ������ � ������ � ��������� �����
' ��� ����� �������� BoundObjectFrame � ���������� ��������� OLE � .ControlSource
Const BitCount = 24
Const ByteCount = BitCount / 8
'Const nColors = 0

' bmih.biCompression = 0 =>  Calculate bytes per line and padding required
Dim BytesPerLine As Long: BytesPerLine = Width * ByteCount
Dim BytesPadding As Long: BytesPadding = 4 - (BytesPerLine Mod 4): If BytesPadding = 4 Then BytesPadding = 0

' translate ole colors
Dim vbColor As Long: If (OleTranslateColor(Color, 0, vbColor) <> 0) Then vbColor = Color
Dim lrgbColor As Long:  lrgbColor = ((vbColor And &HFF000000) Or ((vbColor And &HFF&) * &H10000) Or ((vbColor And &HFF00&)) Or ((vbColor And &HFF0000) \ &H10000))

Dim bmih As BITMAPINFOHEADER
'Create byte array
Dim i As Long
Dim lb As Long, ub As Long
    ub = Len(bmih) + Height * (BytesPerLine + BytesPadding) - 1  '+Len(bmfh) '+ nColors * 4
' Create DIB in memory
    With bmih
        .biSize = &H28
        .biWidth = Width
        .biHeight = Height
        .biPlanes = &H1
        .biBitCount = BitCount
        '.biClrUsed  = nColors
    End With

'Fill Headers
Dim aData() As Byte: ReDim aData(lb To ub): CopyMemory aData(lb), bmih, Len(bmih): lb = lb + Len(bmih)
'Fill Color Table (no need here) ' If BitCount<=8
    'For i = 1 To nColor: CopyMemory aData(LB), 4, aColor(i): LB = LB + 4: Next i
'Fill Pixel Array
    'LB = Len(bmih)
    ' create 1st line from color bytes, then clone it into the rest lines
    For i = 1 To Width:  CopyMemory aData(lb), lrgbColor, ByteCount: lb = lb + ByteCount: Next i: lb = lb + BytesPadding
    For i = 2 To Height: CopyMemory aData(lb), aData(LBound(aData) + Len(bmih)), BytesPerLine: lb = lb + BytesPerLine + BytesPadding: Next i
    'OleObject_BackColor = p_OleObject_GetDIB(aData)
    OleObject_BackColor = p_OleObject_GetPBrush(aData)
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Private Function p_OleObject_GetPBrush(adibData() As Byte)
' create OLE Object (PBrush) picture from byte array (DIB) (24 bit wo Alpha)
On Error GoTo HandleError
'-------------------------
'[MS-OLEDS] https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oleds/85583d21-c1cf-4afe-a35f-d6701c5fbb6f
'[MS-WMF]   https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-wmf/4813e7fd-52d0-4f42-965f-228c8b7488d2
'-------------------------
' Name constants
Const cFName = "Picture", cCName = ""

Dim aoleData() As Byte, lb As Long, ub As Long
Dim aObject() As Byte, lSize As Long
Dim fRet As Long
' create Structures
Dim oo As OLEOBJECTHEADER
    lSize = Len(oo) + Len(cFName) + Len(cCName) + 2
    ub = lSize - 1
    ReDim aoleData(lb To ub)
'----------------------------------------------
' OLEOBJECTHEADER + FriendlyName + ClassName
    With oo
        .Signature = &H1C15
        .HeaderSize = lSize
        .ObjectType = OT_STATIC       'OT_EMBEDDED 'OT_STATIC       '&H2
        .FriendlyNameLen = Len(cFName) + 1
        .ClassNameLen = Len(cCName) + 1
        .FriendlyNameOffset = Len(oo)                   '&H14
        .ClassNameOffset = Len(oo) + .FriendlyNameLen
        .ObjectSize.x = &HFFFF  ' Original size of Object (MM_HIMETRIC)
        .ObjectSize.y = &HFFFF  ' Original size of Object (MM_HIMETRIC)
        '.FriendlyName = StrConv(cFName, vbFromUnicode)
        '.ClassName = StrConv(cCName, vbFromUnicode)
    End With
    ' << oo + StrConv(cFName, vbFromUnicode) + StrConv(cCName, vbFromUnicode)
    CopyMemory aoleData(lb), oo, Len(oo): lb = lb + Len(oo)
    CopyMemory aoleData(lb), ByVal StrPtr(StrConv(cFName & vbNullChar, vbFromUnicode)), Len(cFName) + 1: lb = lb + Len(cFName) + 1
    CopyMemory aoleData(lb), ByVal StrPtr(StrConv(cCName & vbNullChar, vbFromUnicode)), Len(cCName) + 1: lb = lb + Len(cCName) + 1
' Create and add PBrush Presentation
    lSize = p_OleObject_CreateObject(aObject(), adibData(), "PBrush", CF_BITMAP)
    If lSize > 0 Then
    ' << PBrush Object
        ub = ub + lSize: ReDim Preserve aoleData(LBound(aoleData) To ub)
        CopyMemory aoleData(lb), aObject(LBound(aObject)), lSize: lb = lb + lSize
    End If
' Create and add METAFILEPICT Presentation
    lSize = p_OleObject_CreateObject(aObject(), adibData(), "METAFILEPICT")
    If lSize > 0 Then
    ' << METAFILEPICT Object
        ub = ub + lSize: ReDim Preserve aoleData(LBound(aoleData) To ub)
        CopyMemory aoleData(lb), aObject(LBound(aObject)), lSize: lb = lb + lSize
    End If
' CHECKSUM_SIGNATURE (the end of data) ' ???
    lSize = CHECKSUM_STRING_SIZE '4
    If lSize > 0 Then
    ' << CHECKSUM_SIGNATURE
        ub = ub + lSize: ReDim Preserve aoleData(LBound(aoleData) To ub)
        ''CopyMemory aoleData(LB), OleObject_CheckSum8(aOleData) Or CHECKSUM_SIGNATURE, lSize ': LB = LB + lSize
        CopyMemory aoleData(lb), CHECKSUM_SIGNATURE, lSize ': LB = LB + lSize
    End If
''----------------------------------------------
HandleExit:  p_OleObject_GetPBrush = aoleData: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Private Function p_OleObject_CreateObject(ObjectData() As Byte, SourceData() As Byte, Optional ObjectTypeName As String, Optional FormatID As Long) As Long
' Creates array containing OLE Embedded Presentation Object
Dim Result As Long: Result = False
' ObjectData        - reference byte array that will contain created Object data
' SourceData        - byte array the source of object data
' ObjectTypeName    - one of the standart (METAFILEPICT|BITMAP|DIB) or registered class name
' ObjectType        - 5 for standart or 0 for clipboard/registered classes
    On Error GoTo HandleError
Const cAlign = &H119 ' ??? ��������� ������� �����
Dim lSize As Long ', lOffset As Long
    lSize = UBound(SourceData) - LBound(SourceData) + 1: If lSize <= 0 Then Err.Raise vbObjectError + 512
' when working with PictureData arrays.
Dim aObject() As Byte, lb As Long, ub As Long
Const SIZE_OF_BITMAPINFOHEADER = &H28
Const SIZE_OF_BITMAPFILEHEADER = &HE
Dim bmih As BITMAPINFOHEADER
Dim pxWidth As Long, pxHeight As Long
Dim hmWidth As Long, hmHeight As Long
    
' Source is a DIB
    Call CopyMemory(bmih, SourceData(LBound(SourceData)), BITMAPINFOHEADERSIZE)
    pxWidth = bmih.biWidth:     hmWidth = ConvPixelsToHimetrics(pxWidth, DIRECTION_HORIZONTAL)
    pxHeight = bmih.biHeight:   hmHeight = ConvPixelsToHimetrics(pxHeight, DIRECTION_VERTICAL)

' create OLE Object
Const �OleVersion = &H501&
Const �StandardPresentationObject = &H5&
Const �GenericPresentationObject = &H0&
Dim oh As OLEHEADER:       ' Const OLEHEADERSIZE = &HC          ' + oh.ObjectTypeNameLen
    
    Select Case ObjectTypeName
    Case "METAFILEPICT"
'StandardPresentationObjects [MS-OLEDS], 2.2.2
'MetaFilePresentationObject  [MS-OLEDS], 2.2.2.1
Dim mh As METAHEADER, mr As METARECORD
'Dim mrSetMM As METASETMAPMODE, mrSetWin As METASETWINDOW, mrDibStrBlt As METADIBSTRETCHBLT
Dim NDS As Long, RDS As Long
    'get the ladgest record (META_DIBSTRETCHBLT) DataSize
        RDS = &H1A + lSize              ' META_DIBSTRETCHBLT record + BMP wo BITMAPFILEHEADER
    'get NativeDataSize
        NDS = 8                         ' unknown NativeData before META_HEADER
        NDS = NDS + Len(mh)             ' + META_HEADER
        NDS = NDS + 8                   ' + META_SETMAPMODE record
        NDS = NDS + &H14                ' + META_SETWINDOWEXT + META_SETWINDOWORG records
        NDS = NDS + RDS                 ' + META_DIBSTRETCHBLT record + BMP wo BITMAPFILEHEADER
        NDS = NDS + Len(mr)             ' + META_EOF record
    'get result array sizes
        ub = Len(oh) + (Len(ObjectTypeName))
        ub = ub + &HC                   ' + MetaFilePresentationDataWidth& + MetaFilePresentationDataHeight& + NativeDataSize&
        ub = ub + NDS                   ' + NativeDataSize
        ReDim aObject(lb To ub)
'OLEHEADER
        oh.OleVersion = �OleVersion: oh.Format = �StandardPresentationObject
        oh.ObjectTypeNameLen = Len(ObjectTypeName) + 1
    '<< OLEHEADER (CF_BITMAP) & ObjectTypeName
        CopyMemory aObject(lb), oh, Len(oh):        lb = lb + Len(oh)
        CopyMemory aObject(lb), ByVal StrPtr(StrConv(UCase(ObjectTypeName) & vbNullChar, vbFromUnicode)), oh.ObjectTypeNameLen:    lb = lb + oh.ObjectTypeNameLen
'RegisteredClipboardFormatPresentationObject [MS-OLEDS], 2.2.2
    '<< {MetaFilePresentationDataWidth&, MetaFilePresentationDataHeight&}
        CopyMemory aObject(lb), hmWidth, 4:         lb = lb + 4 'Len(hmWidth)
        CopyMemory aObject(lb), -hmHeight, 4:       lb = lb + 4 'Len(hmHeight)
'EmbeddedObject [MS-OLEDS], 2.2.5
    '<< NativeDataSize
        CopyMemory aObject(lb), NDS, 4:             lb = lb + 4 'Len(hmHeight)
        'NativeDataSize = &h0000013E& = SizeOf(PresentationData)
'NativeData/PresentationData
    'unknown NativeData before META_HEADER (8 bytes)
    '<< {0x0008, 0x00D4, 0x00D4, 0x0000}
        CopyMemory aObject(lb), &H8, 2:             lb = lb + 2 '&h0008% ' META_HEADER Offset (bytes) or .MapMode = MM_ANISOTROPIC as in METAFILEPICT struc ???
        CopyMemory aObject(lb), hmWidth, 2:         lb = lb + 2 '&h00D4% ' Original size of Object (MM_HIMETRIC)
        CopyMemory aObject(lb), hmHeight, 2:        lb = lb + 2 '&h00D4% ' Original size of Object (MM_HIMETRIC)
        lb = lb + 2 '0x0000
'METAHEADER
        With mh             'META_HEADER
            .MetaType = MEMORYMETAFILE              'the type of metafile
            .HeaderSize = Len(mh) \ 2               '&H9
            .Version = METAVERSION300               '= METAVERSION300 (DIBs are supported) defines the metafile version. It MUST be a value in the MetafileVersion Enumeration (section 2.1.1.19).<54>
            .SizeLow = p_WordLo((NDS - 8) \ 2)        '= SUMM(.HeaderSize(META_Records))
            .SizeHigh = p_WordHi((NDS - 8) \ 2)
            .NumberOfObjects = &H0                  'the number of graphics objects that are defined in the entire metafile. These objects include brushes, pens, and the other objects specified in section 2.2.1.
            .MaxRecord = RDS \ 2                    'the size of the largest record (META_DIBSTRETCHBLT+DATA) used in the metafile (in 16-bit elements).
        End With
    '<< META_HEADER record
        CopyMemory aObject(lb), mh, Len(mh):        lb = lb + Len(mh)
'METARECORDS
    'META_SETMAPMODE
        With mr: .RecordFunction = META_SETMAPMODE: .RecordSize = &H4&: End With
    '<< META_SETMAPMODE record
        CopyMemory aObject(lb), mr, Len(mr):    lb = lb + Len(mr)
        CopyMemory aObject(lb), MM_ANISOTROPIC, 2: lb = lb + 2 '.MapMode = MM_ANISOTROPIC
    'META_SETWINDOWEXT
        With mr: .RecordFunction = META_SETWINDOWEXT: .RecordSize = &H5&: End With
    '<< META_SETWINDOWEXT record
        CopyMemory aObject(lb), mr, Len(mr):    lb = lb + Len(mr)
        CopyMemory aObject(lb), -pxHeight, 2:   lb = lb + 2 '.y = -pxHeight
        CopyMemory aObject(lb), pxWidth, 2:     lb = lb + 2 '.x = pxWidth
    'META_SETWINDOWORG
        With mr: .RecordFunction = META_SETWINDOWORG: .RecordSize = &H5&: End With
    '<< META_SETWINDOWORG record
        CopyMemory aObject(lb), mr, Len(mr):    lb = lb + Len(mr)
        'CopyMemory aObject(LB), 0, 2:          LB = LB + 2 '.y
        'CopyMemory aObject(LB), 0, 2:          LB = LB + 2 '.x
                                                lb = lb + 4
    ' META_DIBSTRETCHBLT with bitmap [MS-WMF], 2.3.1.3.1
        'If RecordSize > SHR(META_DIBSTRETCHBLT,8)+3 then META_DIBSTRETCHBLT with bitmap
        With mr: .RecordFunction = META_DIBSTRETCHBLT: .RecordSize = RDS \ 2: End With
    '<< META_DIBSTRETCHBLT record +
        CopyMemory aObject(lb), mr, Len(mr):    lb = lb + Len(mr)  ' ???
        CopyMemory aObject(lb), SRCCOPY, 4:     lb = lb + 4 '.RasterOperation& = SRCCOPY   '&HCC0020       '
        CopyMemory aObject(lb), pxHeight, 2:    lb = lb + 2 '.SrcHeight = pxHeight
        CopyMemory aObject(lb), pxWidth, 2:     lb = lb + 2 '.SrcWidth = pxWidth
        'CopyMemory aObject(LB), 0, 2:          LB = LB + 2 '.YSrc
        'CopyMemory aObject(LB), 0, 2:          LB = LB + 2 '.XSrc
                                                lb = lb + 4
        CopyMemory aObject(lb), -pxHeight, 2:   lb = lb + 2 '.DestHeight = -pxHeight                '-8 ???
        CopyMemory aObject(lb), pxWidth, 2:     lb = lb + 2 '.DestWidth = pxWidth
        'CopyMemory aObject(LB), 0, 2:          LB = LB + 2 '.YDest
        'CopyMemory aObject(LB), 0, 2:          LB = LB + 2 '.XDest
                                                lb = lb + 4
    '<< + BITMAP wo BITMAPFILEHEADER
        '.Target            = BITMAPINFOHEADER + RGB DATA (SizeOf(.Target) = 0xE8)
        CopyMemory aObject(lb), SourceData(LBound(SourceData)), (lSize): lb = lb + (lSize)
    'META_EOF
        With mr: .RecordFunction = META_EOF: .RecordSize = &H3: End With
    '<< META_EOF record
        CopyMemory aObject(lb), mr, Len(mr) ':    LB = LB + Len(mr)
    ' NativeData/PresentationData END
' METADATA END
    Case "DIB"          ' [MS-OLEDS], 2.2.2.3 - DIBPresentationObject
Stop
'        UB = Len(oh) + (Len(ObjectTypeName))
'        UB = UB + &HC           ' + Width& + Height& + NativeDataSize&
'        UB = UB + lSize         ' + DIB
'        If lSize Mod cAlign <> 0 Then UB = UB + (cAlign - lSize Mod cAlign)
'        ReDim aObject(LB To UB)
'    'OLEHEADER
'        oh.OleVersion = �OleVersion: oh.Format = �StandardPresentationObject
'        oh.ObjectTypeNameLen = Len(ObjectTypeName) + 1
'        'ObjectTypeName = StrConv(UCase$(ObjectTypeName), vbFromUnicode) & vbNullChar
'    '<< OLEHEADER (CF_DIB) + ObjectTypeName
'        CopyMemory aObject(LB), oh, Len(oh):  LB = LB + Len(oh)
'        CopyMemory aObject(LB), ByVal StrPtr(StrConv(ObjectTypeName & vbNullChar, vbFromUnicode)), oh.ObjectTypeNameLen:    LB = LB + oh.ObjectTypeNameLen
'    '<< {DIBPresentationDataWidth&, DIBPresentationDataHeight&}
'        CopyMemory aObject(LB), hmWidth, 4:         LB = LB + 4 'Len(hmWidth)
'        CopyMemory aObject(LB), -hmHeight, 4:       LB = LB + 4 'Len(hmHeight)
'' PresentationData
'    ' DIB data [MS-WMF], 2.2.2.9
'    '<< PresentationDataSize + PresentationData(DIB)
'        CopyMemory aObject(LB), lSize, 4: LB = LB + 4
'        CopyMemory aObject(LB), SourceData(LBound(SourceData)), lSize ': LB = LB + lSize
    Case "BITMAP" ' BitmapPresentationObject [MS-OLEDS], 2.2.2.2
Stop
'    'OLEHEADER
'        oh.OleVersion = �OleVersion: oh.Format = �StandardPresentationObject
'        oh.ObjectTypeNameLen = Len(ObjectTypeName) + 1
'        'ObjectTypeName = StrConv(UCase$(ObjectTypeName), vbFromUnicode) & vbNullChar
'        'DIBPresentationDataWidth
'        'DIBPresentationDataHeight
    Case Else
'GenericPresentationObject [MS-OLEDS], 2.2.3
    ' 1.StandardClipboardFormatPresentationObject   [MS-OLEDS], 2.2.3.2
    ' 2.RegisteredClipboardFormatPresentationObject [MS-OLEDS], 2.2.3.3
        ub = Len(oh) + (Len(ObjectTypeName))
        ub = ub + &HC           ' + 0x00000000 + 0x00000000 + NativeDataSize&
        ub = ub + BITMAPFILEHEADERSIZE 'Len(bmfh)' + BITMAPFILEHEADER
        ub = ub + lSize         ' + DIB
        If lSize Mod cAlign <> 0 Then ub = ub + (cAlign - lSize Mod cAlign)
        ReDim aObject(lb To ub)
'OLEHEADER (CF_BITMAP)  + ObjectTypeName
        oh.OleVersion = �OleVersion
        oh.Format = FormatID 'CF_BITMAP
        oh.ObjectTypeNameLen = Len(ObjectTypeName) + 1
        'oh.ObjectTypeName = StrConv(ObjectTypeName, vbFromUnicode) & vbNullChar
    '<< OLEHEADER (CF_BITMAP) + ObjectTypeName
        CopyMemory aObject(lb), oh, Len(oh):  lb = lb + Len(oh)
        CopyMemory aObject(lb), ByVal StrPtr(StrConv(ObjectTypeName & vbNullChar, vbFromUnicode)), oh.ObjectTypeNameLen:    lb = lb + oh.ObjectTypeNameLen
    ' RegisteredClipboardFormatPresentationObject [MS-OLEDS], 2.2.2
    '<< 0x00000000, 0x00000000 (8 bytes)
        lb = lb + &H8       ' ObjectTypeName=PBrush: all zeros
' PresentationData
    '<< PresentationDataSize + PresentationData(BITMAP)
        CopyMemory aObject(lb), (ub - lb - 3), 4: lb = lb + 4
        ' Source is DIB - add BITMAPFILEHEADER
        Dim bmfh As BITMAPFILEHEADER: With bmfh
            .bfType = &H4D42
            .bfOffset = &H36 ' + PaletteSize
            .bfSize = .bfOffset + lSize
        End With
        'Write 'BM' from BITMAPFILEHEADER
        CopyMemory aObject(lb), bmfh.bfType, Len(bmfh.bfType) ': LB = LB + Len(bmfh.bfType)
        'Write the rest of BITMAPFILEHEADER
        CopyMemory aObject(lb + Len(bmfh.bfType)), bmfh.bfSize, Len(bmfh) - Len(bmfh.bfType): lb = lb + BITMAPFILEHEADERSIZE 'Len(bmfh)
        CopyMemory aObject(lb), SourceData(LBound(SourceData)), lSize ': LB = LB + lSize
    End Select
    ObjectData = aObject ': Erase aObject()
    Result = UBound(aObject) - LBound(aObject) + 1
HandleExit:  p_OleObject_CreateObject = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_CmpArrays( _
    Data1, _
    Data2, _
    Optional Position As Long = 0, _
    Optional LookFromEnd As Boolean = False _
    ) As Boolean
' ���������� ��� �������� ������� � ���������� ���������
' Position - ������� � ������� ���������� �������� Arr1
' LookFromEnd = False - ��������� � ����� ������ ������� Arr1 (j = ������ ������� + Position)
' LookFromEnd = True - ��������� � ������ ������ ������� Arr1 (����� ������� - Position)
Dim Result As Boolean: Result = False
On Error GoTo HandleError
Dim i As Long, iMax As Long
Dim j As Long, jMax As Long
Dim Arr1() As Byte, Arr2() As Byte
    Arr1() = Data1: Arr2() = Data2
    i = LBound(Arr2): iMax = UBound(Arr2)
    jMax = UBound(Arr1)
    If LookFromEnd Then
    ' ��������� ������ (������)
        j = jMax - Position - (iMax - i)    '- 1
     Else
    ' ��������� ����� (�������)
        j = LBound(Arr1) + Position '- 1
    End If
    If jMax - j < iMax - i Then GoTo HandleExit ' +1?
    ' ���� ����� ���������������� ������� (Arr1) ������ ������� (Arr2)
    ' ��� ������ ���������� - �� ��������� - �������
    Do While i <= iMax
        If Arr1(j) <> Arr2(i) Then GoTo HandleExit
        i = i + 1: j = j + 1
    Loop
    Result = True
HandleExit:
    p_CmpArrays = Result: Exit Function
HandleError:
    Result = False: Resume HandleExit
End Function
Public Function GetFileHandle(ByVal FileName As String, bOpen As Boolean, Optional ByVal useUnicode As Boolean = False) As LongPtr
' Function uses APIs to read/create files with unicode support
Dim pZero As LongPtr

Const GENERIC_READ As Long = &H80000000
Const OPEN_EXISTING = &H3
Const FILE_SHARE_READ = &H1
Const GENERIC_WRITE As Long = &H40000000
Const FILE_SHARE_WRITE As Long = &H2
Const CREATE_ALWAYS As Long = 2
Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Const FILE_ATTRIBUTE_READONLY As Long = &H1
Const FILE_ATTRIBUTE_SYSTEM As Long = &H4

Dim Flags As Long, Access As Long
Dim Disposition As Long, Share As Long

    If useUnicode = False Then useUnicode = (Not (IsWindowUnicode(GetDesktopWindow) = 0&))
    If bOpen Then
        Access = GENERIC_READ
        Share = FILE_SHARE_READ
        Disposition = OPEN_EXISTING
        Flags = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL _
                Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM
    Else
        Access = GENERIC_READ Or GENERIC_WRITE
        Share = 0&
        If useUnicode Then
            Flags = GetFileAttributesW(StrPtr(FileName))
        Else
            Flags = GetFileAttributes(FileName)
        End If
        If Flags < 0& Then Flags = FILE_ATTRIBUTE_NORMAL
        ' CREATE_ALWAYS will delete previous file if necessary
        Disposition = CREATE_ALWAYS
    End If

    If useUnicode Then
        GetFileHandle = CreateFileW(StrPtr(FileName), Access, Share, ByVal pZero, Disposition, Flags, pZero)
    Else
        GetFileHandle = CreateFile(FileName, Access, Share, ByVal pZero, Disposition, Flags, pZero)
    End If
End Function
Public Function FileDelete(FileName As String, Optional ByVal useUnicode As Boolean = False) As Boolean
' Function uses APIs to delete files :: unicode supported
    If useUnicode = False Then useUnicode = (Not (IsWindowUnicode(GetDesktopWindow) = 0&))
    If useUnicode Then
        If Not (SetFileAttributesW(StrPtr(FileName), FILE_ATTRIBUTE_NORMAL) = 0&) Then
            FileDelete = Not (DeleteFileW(StrPtr(FileName)) = 0&)
        End If
    Else
        If Not (SetFileAttributes(FileName, FILE_ATTRIBUTE_NORMAL) = 0&) Then
            FileDelete = Not (DeleteFile(FileName) = 0&)
        End If
    End If
End Function
Public Function FileExists(FileName As String, Optional ByVal useUnicode As Boolean) As Boolean
' test to see if a file exists
Const INVALID_HANDLE_VALUE = -1&
    If useUnicode = False Then useUnicode = (Not (IsWindowUnicode(GetDesktopWindow) = 0&))
    If useUnicode Then
        FileExists = Not (GetFileAttributesW(StrPtr(FileName)) = INVALID_HANDLE_VALUE)
    Else
        FileExists = Not (GetFileAttributes(FileName) = INVALID_HANDLE_VALUE)
    End If
End Function
'=========================
' �������������� ��������
'=========================
Public Function ConvTwipsToPixels(ByVal lngTwips As Long, Optional Direction As Long = DIRECTION_HORIZONTAL) As Long
' convert Twips to Pixels for the current screen resolution
Const c_strProcedure = "ConvTwipsToPixels"
' lngTwips - the number of twips to be converted
' lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
' Returns the number of pixels corresponding to the given twips
Dim Result As Long: Result = False
    On Error GoTo HandleError
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Dim hdc As LongPtr, pZero As LongPtr
#Else
Dim hdc As Long, pZero As Long
#End If
    hdc = GetDC(pZero)
Dim lngPixelsPerInch As Long
    If Direction = DIRECTION_HORIZONTAL Then
        lngPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSX)
    Else
        lngPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSY)
    End If
    hdc = ReleaseDC(pZero, hdc)
    Result = lngTwips / TwipsPerInch * lngPixelsPerInch
HandleExit:  ConvTwipsToPixels = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function ConvPixelsToTwips(ByVal lngPixels As Long, Optional Direction As Long = DIRECTION_HORIZONTAL) As Long
' convert Pixels to Twips for the current screen resolution
Const c_strProcedure = "ConvPixelsToTwips"
' lngPixels - the number of pixels to be converted
' lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
' Returns the number of twips corresponding to the given pixels
Dim Result As Long: Result = False
    On Error GoTo HandleError
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Dim hdc As LongPtr, pZero As LongPtr
#Else
Dim hdc As Long, pZero As Long
#End If
    hdc = GetDC(pZero)
Dim lngPixelsPerInch As Long
    If Direction = DIRECTION_HORIZONTAL Then
        lngPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSX)
     Else
        lngPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSY)
    End If
    ConvPixelsToTwips = lngPixels * TwipsPerInch / lngPixelsPerInch
    Result = lngPixels * TwipsPerInch / lngPixelsPerInch
HandleExit:  ConvPixelsToTwips = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Function ConvPixelsToHimetrics(ByVal lngPixels As Long, Optional Direction As Long = DIRECTION_HORIZONTAL) As Long
' convert Pixels to Himetrics
Const c_strProcedure = "ConvPixelsToHimetrics"
Dim Result As Long: Result = False
    On Error GoTo HandleError
    If Direction = DIRECTION_HORIZONTAL Then
        Result = lngPixels * (HimetricPerInch / TwipsPerInch) * TwipsPerPixels(LOGPIXELSX) 'Screen.TwipsPerPixelX
    Else
        Result = lngPixels * (HimetricPerInch / TwipsPerInch) * TwipsPerPixels(LOGPIXELSY) 'Screen.TwipsPerPixelY
    End If
HandleExit:  ConvPixelsToHimetrics = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function TwipsPerPixels(Optional ByVal Dimension As Long = LOGPIXELSX) As Long
' can be replaced with ConvPixelsToTwips (1,DIRECTION_HORIZONTAL|DIRECTION_VERTICAL)
    On Error GoTo HandleError
    TwipsPerPixels = TwipsPerInch / GetDeviceCaps(GetDC(Application.hWndAccessApp), Dimension)
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
'=========================
Public Function CreateHFont(Optional FontName, Optional FontSize, Optional FontWeight, Optional FontItalic, Optional FontUnderline, Optional FontStrikeOut) As LongPtr
' ������ �����
Const CName = "Arial", cSize = 10 ', cWeight = 0, cItalic = False, cUnderline = False, cStrikeOut = False
Dim hFont As LongPtr
    On Error GoTo HandleError
Dim hdc As LongPtr: hdc = GetDC(0): If (hdc = 0) Then Err.Raise vbObjectError + 512
' ��������� ������ ��-���������
Dim nHeight As Long             ' ������� ������ �������
'Dim nWidth As Long              ' ������� ������ �������
'Dim nEscapement As Long         ' ���� �������, � ������� �������, ����� �������� ������� � ���� X ����������. ������ ������� ���������� �������� ����� ���� ������
'Dim nOrientation As Long        ' ���� ���������� �������� �����, � ������� �������, ����� �������� ������ ������� ������� � ���� X ����������
Dim fnWeight As Long            ' ������� ������
Dim fdwItalic  As Long          ' ��������� ��������� ���������� ������
Dim fdwUnderline As Long        ' ��������� ��������� �������������
Dim fdwStrikeOut  As Long       ' ��������� ��������� ������������
'Dim fdwCharSet As Long          ' ������������� ������ ��������
'Dim fdwOutputPrecision As Long  ' �������� ������
'Dim fdwClipPrecision As Long    ' �������� ���������
'Dim fdwQuality As Long          ' �������� ������
'Dim fdwPitchAndFamily As Long   ' ��� ����� ��������� ������ � ���������
Dim lpszFace As String * LF_FACESIZE      ' ��� ��������� ������
Dim nSize As Long

    If Not IsMissing(FontName) Then lpszFace = Left$(FontName, LF_FACESIZE - 1) & vbNullChar Else lpszFace = CName
    If Not IsMissing(FontSize) Then nSize = FontSize Else nSize = cSize
    If Not IsMissing(FontItalic) Then fdwItalic = FontItalic
    If Not IsMissing(FontUnderline) Then fdwUnderline = FontUnderline
    If Not IsMissing(FontStrikeOut) Then fdwStrikeOut = FontStrikeOut
    If Not IsMissing(FontWeight) Then fnWeight = FontWeight
    'fnWeight =                              ' FW_DONTCARE | FW_THIN | FW_EXTRALIGHT | FW_LIGHT | FW_NORMAL | FW_MEDIUM | FW_SEMIBOLD | FW_BOLD | FW_EXTRABOLD | FW_HEAVY | FW_BLACK | FW_DEMIBOLD | FW_REGULAR | FW_ULTRABOLD | FW_ULTRALIGHT
    'nEscapement = nDegrees * 10             '
    'nOrientation =                          '
    'fdwCharSet = DEFAULT_CHARSET            ' DEFAULT_CHARSET | SYMBOL_CHARSET | RUSSIAN_CHARSET | OEM_CHARSET | SHIFTJIS_CHARSET | HANGEUL_CHARSET | CHINESEBIG5_CHARSET
    'fdwOutputPrecision = OUT_DEFAULT_PRECIS ' OUT_CHARACTER_PRECIS | OUT_DEFAULT_PRECIS | OUT_DEVICE_PRECIS
    'fdwClipPrecision = CLIP_DEFAULT_PRECIS  ' CLIP_DEFAULT_PRECIS | CLIP_CHARACTER_PRECIS | CLIP_STROKE_PRECIS
    'fdwQuality = PROOF_QUALITY              ' DEFAULT_QUALITY | DRAFT_QUALITY | PROOF_QUALITY | NONANTIALIASED_QUALITY | ANTIALIASED_QUALITY | CLEARTYPE_QUALITY
    'fdwPitchAndFamily = DEFAULT_PITCH       ' DEFAULT_PITCH | FIXED_PITCH | VARIABLE_PITCH
'   ������ ������
    '> 0 �������� ����������� ����������� ������ � ���������� ����������� ��� �������� � ������� ��������� ���������� (�������) � ������������� ��� � ����������� �� ������ ����� �������� ��������� �������.
    '= 0 �������� ����������� ����������� ������ � ���������� ���������� �������� �� ��������� �������� ������, ����� �� ���� ������������ �������.
    '< 0 �������� ����������� ����������� ������ � ���������� ����������� ��� �������� � ������� ��������� ���������� (�������) � ������������� ��� ���������� �������� � ����������� �� ������ ������� ��������� �������.
    'nSize = -(FontSize * PT / TwipsPerPixels)'
    nHeight = -MulDiv(nSize, GetDeviceCaps(hdc, 90), PointsPerInch)   ' pt -> px ' LOGPIXELSY = 90
    hFont = CreateFont(nHeight, 0, 0, 0, fnWeight, fdwItalic, fdwUnderline, fdwStrikeOut, _
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH, _
        lpszFace)
'Stop
'' �������� ������� ����������� ������
'Dim LF As LOGFONT: GetObject hFont, LenB(LF), LF
'    With LF
'        FontName = StrZ(.lfFaceName())
'        FontSize = -MulDiv(.lfHeight, PointsPerInch, GetDeviceCaps(hDC, 90))
'        FontWeight = .lfWeight
'        FontItalic = CBool(.lfItalic)
'        FontUnderline = CBool(.lfUnderline)
'        FontStrikeOut = CBool(.lfStrikeOut)
'    End With
HandleExit: If (hdc <> 0) Then Call ReleaseDC(0, hdc)
    CreateHFont = hFont: Exit Function
HandleError: hFont = 0: Err.Clear: Resume HandleExit
End Function
Public Function TextToArrayByHFont(TextString As String, Optional ByVal hFont As LongPtr, _
    Optional ByRef WidthInPix, Optional ByRef HeightInPix, _
    Optional Separators As String = " �.,;:!?()[]{}�+-*/\|" & vbTab & vbCrLf, _
    Optional ByRef OutLines, _
    Optional Overhang As Long, Optional OutDelimiter = vbCrLf) As String
' , Optional OutLineWidth, Optional OutLineHeight
' ����� ����� �� ����� � ����������� �� ������ � ������
Const c_strProcedure = "TextToArrayByHfont"
' TextString - ������ ������ ������� ���������� �������
' hFont - hFont ������ ��� �������� ������������ ���������
' WidthInPix   - �� ����� - ������������ ������ ��������� ������,
'                �� ������ - �������� ������ ��������� ������
' HeightInPix  - �� ������ - �������� ������ ��������� ������
' Separators   - ������ ������������ �� ������� ����� ���� ����� ���� ������� ������ ������� �������� - � �������� ������ ����� ������
' OutLines     - ������ ����� ��������� ������
'' OutLineWidth, OutLineHeight - ������� �������� ����� ��������� ������
' Overhang     - �������� ��� ������������� ������� ��� ���������, ������ � ��. �������
' OutDelimiter - ����������� ����� � �������� ������
Dim Result As String: Result = vbNullString
    On Error GoTo HandleError
    If WidthInPix < 0 Then GoTo HandleExit
Dim tWidth As Long, tHeight As Long: tWidth = 0: tHeight = 0
Dim hdc As LongPtr: hdc = GetDC(0): If (hdc = 0) Then Err.Raise vbObjectError + 512
' split text into the parts
Dim aWords() As String: Call Tokenize(TextString, aWords, Separators)
Dim i As Long, iMax As Long, ii As Long, w As Long
Dim spLen As Long, spPos As Integer, spPosNext As Integer
    i = LBound(aWords): iMax = UBound(aWords)
    ii = 0: spLen = 1
    ' �������: vbCrLf ������ �� vbCr ����� ������ ������� ������ ������
Dim strText As String, strRest As String, strTemp As String
    strRest = Replace$(TextString, vbCrLf, vbCr)
' If hFont is null then get default font
    If hFont = 0 Then hFont = GetStockObject(SYSTEM_FONT)
Dim hOldFont As LongPtr: hOldFont = SelectObject(hdc, hFont) ' select hFont into DC
' go throw the parts while line width is less then WidthInPix
' and begin new line when length is above
Dim aText() As String ', aWidth() As Long, aHeight() As Long
Dim sz As POINT
    Do
        w = 0
        strText = vbNullString
        Do
        ' ���������� ����� ������
            If i < iMax Then
                strTemp = aWords(i)
                spPos = Len(strTemp) + 1
                spPosNext = InStr(spPos, strRest, aWords(i + 1))
                spLen = spPosNext - spPos
            Else
                strTemp = strRest
                spPos = Len(strTemp) + 1
                spPosNext = spPos
                spLen = 0
            End If
            ' ������� ������ �������� Chr(&HAD)
            strTemp = Replace$(strText & strTemp, Chr(&HAD), vbNullString)
            ' ������ ����� ���������� ������ + ������� �������� + ������� �����������
            strTemp = strTemp & Mid$(strRest, spPos, spLen)
            spLen = Len(Trim$(strTemp)): If spLen = 0 Then spLen = 1
' get text size with GetTextExtentPoint32 function
            GetTextExtentPoint32 hdc, StrPtr(strTemp), spLen, sz
            If sz.x <= WidthInPix Or w = 0 Then
        ' �������: w=0, - ��� ���� �������
            ' ���� ������ ����� � ������ ������ ������� ������ - �� ����� ����,
            ' ����� �������� � ������� �����
                If sz.x > tWidth Then tWidth = sz.x
                strRest = Mid$(strRest, spPosNext)
                strText = strTemp
                i = i + 1
                w = w + 1
            End If
        Loop Until i > iMax Or sz.x > WidthInPix
        ReDim Preserve aText(ii): aText(ii) = Trim$(strText)
'        ReDim Preserve aWidth(ii): aWidth(ii) = sz.x
'        ReDim Preserve aHeight(ii): aHeight(ii) = sz.y
        tHeight = tHeight + sz.y
'        Result = Result & OutDelimiter & strText
    ' ���� �������� ����� - �������
        If Len(strRest) = 0 Then Exit Do
        ii = ii + 1
    Loop
' �������� �������� ������ � �� ������ � ��������
    WidthInPix = tWidth: HeightInPix = tHeight
    Result = Join(aText, OutDelimiter)
    OutLines = aText:           Erase aText
'    OutLineWidth = aWidth:      Erase aWidth
'    OutLineHeight = aHeight:    Erase aHeight
Dim tm As TEXTMETRIC: GetTextMetrics hdc, tm
    Overhang = tm.tmOverhang ' ������� ��� ��������� � ������� �������
    ' Destroy the new font.
    SelectObject hdc, hOldFont
HandleExit:
    If (hdc <> 0) Then Call ReleaseDC(0, hdc)
    TextToArrayByHFont = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function CommonDlg_ChooseColor(ByRef theColor As Long) As Boolean
' Show Common Dialog ChooseColor and return result
Dim Result As Boolean: Result = False
'https://www.devhut.net/vba-choosecolor-api-x32-x64/
    On Error GoTo HandleError
Dim cc As ChooseColor
' Some predefined color, there are 16 slots available for predefined colors
' You don't have to defined any, if you don't want to!
Static CustomColors(16)   As Long
    CustomColors(0) = vbWhite
    CustomColors(1) = vbBlack
    CustomColors(2) = vbRed
    CustomColors(3) = vbGreen
    CustomColors(4) = vbBlue
' Fill structure
    With cc
        .lStructSize = LenB(cc)
        .hwndOwner = Application.hWndAccessApp
        .Flags = CC_ANYCOLOR Or CC_FULLOPEN Or CC_PREVENTFULLOPEN Or CC_RGBINIT
        .rgbResult = theColor      'Set the initial color of the dialog
        .lpCustColors = VarPtr(CustomColors(0))
    End With
' Call dialog and get result
    If ChooseColor(cc) = 0 Then Err.Raise vbObjectError + 512 'Cancelled by the user
    theColor = cc.rgbResult
HandleExit:  CommonDlg_ChooseColor = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function CommonDlg_ChooseFont(FontName As String, Optional FontSize As Single, Optional FontWeight As Long, Optional FontColor As Long, _
    Optional FontBold As Boolean, Optional FontItalic As Boolean, Optional FontUnderline As Boolean, Optional FontStrikeOut As Boolean) As Boolean
' Show Common Dialog ChooseFont and return selection in theFont
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
Dim LF As LOGFONT, CF As ChooseFont
#If VBA7 Then           ' <OFFICE2010+>
Dim pLF As LongPtr, hMem As LongPtr
#Else                   ' <OFFICE97-2007>
Dim pLF As Long, hMem As Long
#End If                 ' <VBA7 & WIN64>
Dim sText As String
    With LF
        .lfHeight = -MulDiv(CLng(FontSize), GetDeviceCaps(GetDC(hWndAccessApp), LOGPIXELSY), PointsPerInch)
        .lfWeight = IIf(FontBold And (FontWeight = FW_DONTCARE), FW_BOLD, FontWeight)
        .lfItalic = Abs(FontItalic)
        .lfUnderline = Abs(FontUnderline)
        .lfStrikeOut = Abs(FontStrikeOut)
        Call p_Str2Bytes(FontName, .lfFaceName())
    End With
    With CF
        .lStructSize = LenB(CF)
        .hwnd = Application.hWndAccessApp ' to be modal must be valid Hwnd
        .rgbColors = FontColor
    
        hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, LenB(LF)):  If hMem = 0 Then Err.Raise vbObjectError + 512
        pLF = GlobalLock(hMem):  If pLF = 0 Then Err.Raise vbObjectError + 512
        CopyMemory ByVal pLF, LF, LenB(LF)
        .lpLogFont = pLF
        .Flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
' This had better be the address of a public function in a standard module, or you're going down!
' Use the adhFnPtrToLong procedure to convert from AddressOf to long.
        'If .Flags And cdlCFEnableHook Then .lpfnHook = Callback
    End With
'
    If ChooseFont(CF) = 0 Then Err.Raise vbObjectError + 512
'
    CopyMemory LF, ByVal pLF, LenB(LF)
    With LF
        FontWeight = .lfWeight
        FontBold = (.lfWeight >= FW_BOLD)
        FontItalic = CBool(.lfItalic)
        FontUnderline = CBool(.lfUnderline)
        FontStrikeOut = CBool(.lfStrikeOut)
        FontName = p_BytesToStr(.lfFaceName())
    End With
    With CF
        FontSize = CLng(.iPointSize / 10)
        FontColor = .rgbColors
    End With
    Result = True
HandleExit:  CommonDlg_ChooseFont = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

'=========================
' ��������������� �������
'=========================
Private Function p_BytesToStr(aBytes() As Byte) As String
Dim i As Long: i = LBound(aBytes)
    While i <= UBound(aBytes)
Dim lVal As Long:       lVal = aBytes(i): If lVal = 0 Then GoTo HandleExit
Dim szOut As String:    szOut = szOut & Chr$(lVal)
        i = i + 1
    Wend
'    CopyMemory ByVal StrPtr(sText), ByVal VarPtr(.lfFaceName(1)), ByVal LF_FACESIZE
'    FontName = StrConv(StrZ(sText), vbUnicode)
HandleExit: p_BytesToStr = szOut
End Function
Private Sub p_Str2Bytes(InString As String, ByteArray() As Byte)
Dim iMin As Long: iMin = LBound(ByteArray)
Dim iMax As Long: iMax = UBound(ByteArray)
Dim lLen As Long: lLen = Len(InString)
Dim i As Long
    If lLen > iMax - iMin Then lLen = iMax - iMin
    For i = 1 To lLen:  ByteArray(i - 1 + iMin) = Asc(Mid(InString, i, 1)): Next i
'   CopyMemory ByVal VarPtr(ByteArray(iMin)), ByVal StrPtr(sText), ByVal LenB(sText)
End Sub
Private Function p_WordLo(DWord As Long) As Integer
    If DWord And &H8000& Then p_WordLo = DWord Or &HFFFF0000 Else p_WordLo = DWord And &HFFFF&
End Function
Private Function p_WordHi(DWord As Long) As Integer: p_WordHi = (DWord And &HFFFF0000) \ &H10000: End Function
Public Function ByteAlignOnWord(ByVal BitDepth As Byte, ByVal Width As Long) As Long
' function to align any bit depth on dWord boundaries
    ByteAlignOnWord = (((Width * BitDepth) + &H1F&) And Not &H1F&) \ &H8&
End Function
Private Function p_DWordRead(ByteArray() As Byte, Pos As Long, Optional LittleEndian As Boolean = True) As Long
' ������ ������� ����� �� ��������� �������
' � ������� Pos
' ���� LittleEndian = True  - ������� ������ �������� (������� �������)
' ���� LittleEndian = False - ������� ������ ������ (������� �������)
Dim Result As Long
Dim i As Long:      i = Pos + LBound(ByteArray)
Dim j As Byte
Dim jMin As Byte:   jMin = 0
Dim jMax As Byte:   jMax = 3
Dim hexDWord As String * 8: hexDWord = String$(8, "0")
Dim hexByte As String '* 2
Dim p As Byte
    For j = jMin To jMax
        hexByte = Hex$(ByteArray(i + j))
        If LittleEndian Then p = 2 * (jMax - j) + (3 - Len(hexByte)) Else p = 2 * j + (3 - Len(hexByte))
        Mid$(hexDWord, p, 2) = hexByte
    Next j
    Result = CLng("&H" & hexDWord)
HandleExit: p_DWordRead = Result: Exit Function
HandleError:
End Function
Private Function p_WordRead(ByteArray() As Byte, Pos As Long, Optional LittleEndian As Boolean = True) As Long
' ������ ����� �� ��������� ������� ByteArray � ������� Pos
' ���� LittleEndian = True  - ������� ������ �������� (������� �������)
' ���� LittleEndian = False - ������� ������ ������ (������� �������)
Dim Result As Long
Dim i As Long:      i = Pos + LBound(ByteArray)
Dim j As Byte
Dim jMin As Byte:   jMin = 0
Dim jMax As Byte:   jMax = 1
Dim hexWord As String * 4: hexWord = String$(4, "0")
Dim hexByte As String '* 2
Dim p As Byte
    For j = jMin To jMax
        hexByte = Hex$(ByteArray(i + j))
        If LittleEndian Then p = 2 * (jMax - j) + (3 - Len(hexByte)) Else p = 2 * j + (3 - Len(hexByte))
        Mid$(hexWord, p, 2) = hexByte
    Next j
    Result = CLng("&H" & hexWord)
HandleExit: p_WordRead = Result: Exit Function
HandleError:
End Function
Private Function p_ByteRead(ByteArray() As Byte, Pos As Long)
' ������ ���� �� ��������� ������� ByteArray � ������� Pos
    p_ByteRead = ByteArray(Pos + LBound(ByteArray))
End Function
' ��������������
Private Function p_Min(ParamArray Values()) As Variant
If UBound(Values) < LBound(Values) Then Exit Function
Dim i As Long
    p_Min = Values(LBound(Values))
    For i = LBound(Values) + 1 To UBound(Values)
        If Values(i) < p_Min Then p_Min = Values(i)
    Next i
End Function
Private Function p_Max(ParamArray Values()) As Variant
If UBound(Values) < LBound(Values) Then Exit Function
Dim i As Long
    p_Max = Values(LBound(Values))
    For i = LBound(Values) + 1 To UBound(Values)
        If Values(i) > p_Max Then p_Max = Values(i)
    Next i
End Function
Private Function p_ATan2(x As Single, y As Single) As Single
    Select Case x
    Case Is > 0: p_ATan2 = Atn(y / x)
    Case Is < 0: p_ATan2 = Atn(y / x) + Pi * Sgn(y): If y = 0 Then p_ATan2 = p_ATan2 + Pi
    Case Is = 0: p_ATan2 = Pi / 2 * Sgn(y)
    End Select
End Function
#If ObjectDataType = 0 Then         'FI
'=========================
' ������� ����������� ��� FreeImage
'=========================
Private Function p_FreeImage_RotateExEx(ByVal BITMAP As LongPtr, _
    Optional ByRef cX0 As Single = 0&, Optional ByRef cY0 As Single = 0&, _
    Optional ByRef cX1 As Single = 0&, Optional ByRef cY1 As Single = 0&, _
    Optional ByRef Angle As Single = 0&, _
    Optional ByRef Color As Long = 0&) As LongPtr ', _
' ���������� ����������� ��������������� �� ������� ������������� ���a���� � ������
'---------------------
'   BITMAP   - ��������� �� FreeImage bitmap
'---------------------
    On Error GoTo HandleExit
    If BITMAP = 0 Then Exit Function
Dim mAngle As Single: mAngle = Angle - Fix(Angle / 360) * 360: If mAngle < 0 Then mAngle = 360 + mAngle ' ����������� ����
    If cX0 = 0 Then cX0 = FreeImage_GetWidth(BITMAP)
    If cY0 = 0 Then cY0 = FreeImage_GetHeight(BITMAP)
    If cX1 = 0 Or cY1 = 0 Then
Dim cTrans As New clsTransform
        With cTrans
            .Angle = mAngle: Call .TransformSize(cX0, cY0, cX1, cY1)
        End With
    End If
Dim sx As Long, sy As Long                          ' �������� �������� �� ���� ��� ������������� ��������
Dim dl As Long, dt As Long, dr As Long, db As Long  ' �������� ����������/������ ������� ��� �������� ��������� ��������
' �������� ������������ ����������/������ ������� �����������
    If (cX1 > cX0) Then
        dl = (cX1 - cX0) / 2: dr = dl: If dl = 0 Then dr = 1
    ElseIf ((mAngle = 180) Or (mAngle = 270)) Then
' ������� ��� ������������� ������� ������� (��� ���� ����� ������ ����� ��� ��������)
    ' ����������� ��� �������� ������� �� ������� ������ �� 2, �� ��� ���� ����� ������ �����
        dl = 1: dr = 1 ' ����� ����� ������ �������
    End If
    If (cY1 > cY0) Then
        dt = (cY1 - cY0) / 2: db = dt: If dt = 0 Then db = 1
    ElseIf ((mAngle = 90) Or (mAngle = 180)) Then
        dt = 1: db = 1  ' ����� ����� ������ �������
    End If
' ��������� ������� �����������
Dim fiPict As LongPtr, fiTemp As LongPtr: fiTemp = BITMAP
    If ((dr > 0) Or (db > 0)) Then fiPict = FreeImage_EnlargeCanvas(fiTemp, dl, dt, dr, db, Color, FI_COLOR_IS_RGBA_COLOR): Call FreeImage_Unload(fiTemp): fiTemp = fiPict
' ������������ �����������
    fiPict = FreeImage_RotateEx(fiTemp, mAngle, sx, sy, FreeImage_GetWidth(fiTemp) / 2, FreeImage_GetHeight(fiTemp) / 2, 1): Call FreeImage_Unload(fiTemp): fiTemp = fiPict
' �������� ������������ ������ ������� ����� �������� (������� ������)
    dl = 0: dt = 0: dr = 0: db = 0
    If (cX1 < cX0) Then
        dl = (cX1 - cX0) / 2: dr = dl: If dl = 0 Then dl = -1
    ElseIf ((mAngle = 180) Or (mAngle = 270)) Then
    ' ����������� ��� �������� ������� �� ������� ������ �� 2, �� ��� ���� ����� ������ �����
        dl = -2: dr = 0     ' ����� ����������� �������������
    End If
    If (cY1 < cY0) Then
        dt = (cY1 - cY0) / 2: db = dt: If dt = 0 Then dt = -1  ': sy = -1
    ElseIf ((mAngle = 90) Or (mAngle = 180)) Then
        dt = -2: db = 0     ' ����� ����������� �������������
    End If
' �������� ������� �����������
    If ((dl < 0) Or (dt < 0)) Then fiPict = FreeImage_EnlargeCanvas(fiTemp, dl, dt, dr, db, Color, FI_COLOR_IS_RGBA_COLOR): FreeImage_Unload (fiTemp)
    p_FreeImage_RotateExEx = fiPict
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
#ElseIf ObjectDataType = 1 Then     'LV
'=========================
' ������� ����������� ��� clsPictureData
'=========================
Public Function ValidateDLL(ByVal DllName As String, ByVal dllProc As String) As Boolean
' Test a DLL for a specific function.
Dim hLib As LongPtr, pProc As LongPtr
    hLib = LoadLibrary(DllName) 'attempt to open the DLL to be checked
    If hLib Then pProc = GetProcAddress(hLib, dllProc): FreeLibrary hLib 'if so, retrieve the address of one of the function calls
    ValidateDLL = (Not (hLib = 0 Or pProc = 0))
End Function
Public Function CreateShapedRegion(cHost As clsPictureData, regionStyle As eRegionStyles) As Long
'*******************************************************
' FUNCTION RETURNS A HANDLE TO A REGION IF SUCCESSFUL.
' If unsuccessful, function retuns zero.
' The fastest region from bitmap routines around, custom
' designed by LaVolpe. This version modified to create
' regions from alpha masks.
'*******************************************************
' Note: See clsPictureData.CreateRegion for description of the regionStyle parameter

' declare bunch of variables...
Dim rgnRects() As RECT ' array of rectangles comprising region
Dim rectCount As Long ' number of rectangles & used to increment above array
Dim rStart As Long ' pixel that begins a new regional rectangle
Dim x As Long, y As Long, z As Long ' loop counters
Dim bDib() As Byte  ' the DIB bit array
Dim tSA As SAFEARRAY2D ' array overlay
Dim rtnRegion As Long ' region handle returned by this function
Dim Width As Long, Height As Long
Dim lScanWidth As Long ' used to size the DIB bit array
    ' Simple sanity checks
    If cHost.Alpha = AlphaNone Then
        CreateShapedRegion = CreateRectRgn(0&, 0&, cHost.Width, cHost.Height)
        Exit Function
    End If
    Width = cHost.Width
    If Width < 1& Then Exit Function
    Height = cHost.Height
    If Height < 1& Then Exit Function
    On Error GoTo Cleanup
    lScanWidth = Width * 4& ' how many bytes per bitmap line?
    With tSA                ' prepare array overlay
        .cbElements = 1     ' byte elements
        .cDims = 2          ' two dim array
        .pvData = cHost.BitsPointer  ' data location
        .rgSAbound(0).cElements = Height
        .rgSAbound(1).cElements = lScanWidth
    End With
    ' overlay now
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
    If regionStyle = regionShaped Then
        ReDim rgnRects(0 To Width * 3&) ' start with an arbritray number of rectangles
        ' begin pixel by pixel comparisons
        For y = Height - 1 To 0& Step -1&
            ' the alpha byte is every 4th byte
            For x = 3& To lScanWidth - 1& Step 4&
                ' test to see if next pixel is 100% transparent
                If bDib(x, y) = 0 Then
                    If Not rStart = 0& Then ' we're currently tracking a rectangle,
                        ' so let's close it, but see if array needs to be resized
                        If rectCount + 1& = UBound(rgnRects) Then _
                            ReDim Preserve rgnRects(0 To UBound(rgnRects) + Width * 3&)
                         
                         ' add the rectangle to our array
                         SetRect rgnRects(rectCount + 2&), rStart \ 4, Height - y - 1&, x \ 4 + 1&, Height - y
                         rStart = 0&                    ' reset flag
                         rectCount = rectCount + 1&     ' keep track of nr in use
                    End If
                
                Else
                    ' non-transparent, ensure start value set
                    If rStart = 0& Then rStart = x  ' set start point
                End If
            Next x
            If Not rStart = 0& Then
                ' got to end of bitmap without hitting another transparent pixel
                ' but we're tracking so we'll close rectangle now
               
               ' see if array needs to be resized
               If rectCount + 1& = UBound(rgnRects) Then _
                   ReDim Preserve rgnRects(0 To UBound(rgnRects) + Width * 3&)
                   
                ' add the rectangle to our array
                SetRect rgnRects(rectCount + 2&), rStart \ 4, Height - y - 1&, Width, Height - y
                rStart = 0&                     ' reset flag
                rectCount = rectCount + 1&      ' keep track of nr in use
            End If
        Next y
    ElseIf regionStyle = regionEnclosed Then
        ReDim rgnRects(0 To Width * 3&) ' start with an arbritray number of rectangles
        ' begin pixel by pixel comparisons
        For y = Height - 1 To 0& Step -1&
            ' the alpha byte is every 4th byte
            For x = 3& To lScanWidth - 1& Step 4&
                ' test to see if next pixel has any opaqueness
                If Not bDib(x, y) = 0 Then
                    ' we got the left side of the scan line, check the right side
                    For z = lScanWidth - 1 To x + 4& Step -4&
                        ' when we hit a non-transparent pixel, exit loop
                        If Not bDib(z, y) = 0 Then Exit For
                    Next
                    ' see if array needs to be resized
                    If rectCount + 1& = UBound(rgnRects) Then _
                        ReDim Preserve rgnRects(0 To UBound(rgnRects) + Width * 3&)
                     
                     ' add the rectangle to our array
                     SetRect rgnRects(rectCount + 2&), x \ 4, Height - y - 1&, z \ 4 + 1&, Height - y
                     rectCount = rectCount + 1&     ' keep track of nr in use
                     Exit For
                End If
            Next x
        Next y
    ElseIf regionStyle = regionBounds Then
        ReDim rgnRects(0 To 0) ' we will only have 1 regional rectangle
        ' set the min,max bounding parameters
        SetRect rgnRects(0), Width * 4, Height, 0, 0
        With rgnRects(0)
            ' begin pixel by pixel comparisons
            For y = Height - 1 To 0& Step -1&
                ' the alpha byte is every 4th byte
                For x = 3& To lScanWidth - 1& Step 4&
                    ' test to see if next pixel has any opaqueness
                    If Not bDib(x, y) = 0 Then
                        ' we got the left side of the scan line, check the right side
                        For z = lScanWidth - 1 To x + 4& Step -4&
                            ' when we hit a non-transparent pixel, exit loop
                            If Not bDib(z, y) = 0 Then Exit For
                        Next
                        rStart = 1& ' flag indicating we have opaqueness on this line
                        ' resize our bounding rectangle's left/right as needed
                        If x < .Left Then .Left = x
                        If z > .Right Then .Right = z
                        Exit For
                    End If
                Next x
                If rStart = 1& Then
                    ' resize our bounding rectangle's top/bottom as needed
                    If y < .Top Then .Top = y
                    If y > .Bottom Then .Bottom = y
                    rStart = 0& ' reset flag indicating we do not have any opaque pixels
                End If
            Next y
        End With
        If rgnRects(0).Right > rgnRects(0).Left Then
            rtnRegion = CreateRectRgn(rgnRects(0).Left \ 4, Height - rgnRects(0).Bottom - 1&, rgnRects(0).Right \ 4 + 1&, _
                                     (rgnRects(0).Bottom - rgnRects(0).Top) + (Height - rgnRects(0).Bottom))
        End If
    End If
    ' remove the array overlay
    CopyMemory ByVal VarPtrArray(bDib()), 0&, 4&
    On Error Resume Next
    ' check for failure & engage backup plan if needed
    If Not rectCount = 0 Then
        ' there were rectangles identified, try to create the region in one step
        rtnRegion = p_CreatePartialRegion(rgnRects(), 2&, rectCount + 1&, 0&, Width)
        
        ' ok, now to test whether or not we are good to go...
        ' if less than 2000 rectangles, region should have been created & if it didn't
        ' it wasn't due to O/S restrictions -- failure
        If rtnRegion = 0& Then
            If rectCount > 2000& Then
                ' Win98 has limitation of approximately 4000 regional rectangles
                ' In cases of failure, we will create the region in steps of
                ' 2000 vs trying to create the region in one step
                rtnRegion = p_CreateWin98Region(rgnRects, rectCount + 1&, 0&, Width)
            End If
        End If
    End If
Cleanup:
    Erase rgnRects()
    If Err Then ' failure; probably low on resources
        If Not rtnRegion = 0& Then DeleteObject rtnRegion
        Err.Clear
    Else
        CreateShapedRegion = rtnRegion
    End If
End Function
Public Function HandleToStdPicture(ByVal hImage As Long, ByVal imgType As Long) As IPicture
' function creates a stdPicture object from an image handle (bitmap or icon)
Dim lpPictDesc As PICTDESC, aGUID(0 To 3) As Long
    With lpPictDesc
        .Size = Len(lpPictDesc)
        .Type = imgType
        .hImage = hImage
        .Reserved1 = 0
        .Reserved2 = 0
    End With
    ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    ' create stdPicture
    Call OleCreatePictureIndirect(lpPictDesc, aGUID(0), True, HandleToStdPicture)
End Function
Public Sub ValidateAlphaChannel(inStream() As Byte, bPreMultiply As Boolean, bIsAlpha As AlphaTypeEnum, imgType As Long)
' Purpose: Modify 32bpp DIB's alpha bytes depending on whether or not they are used
' Parameters
' inStream(). 2D array overlaying the DIB to be checked
' bPreMultiply. If true, image will be premultiplied if not already
' bIsAlpha. Returns whether or not the image contains transparency
' imgType. If passed as -1 then image is known to be not alpha, but will have its alpha values set to 255
'          When routine returns, imgType is either imgBmpARGB, imgBmpPARGB or imgBitmap
Dim x As Long, y As Long
Dim lPARGB As Long, zeroCount As Long, opaqueCount As Long
Dim bPARGB As Boolean, bAlpha As AlphaTypeEnum
    ' see if the 32bpp is premultiplied or not and if it is alpha or not
    ' �� ��������� ���������� alpha ����� � PNG
    If Not imgType = -1 Then
        For y = 0 To UBound(inStream, 2)
            For x = 3 To UBound(inStream, 1) Step 4
                Select Case inStream(x, y)
                Case 0
                    If lPARGB = 0 Then
                        ' zero alpha, if any of the RGB bytes are non-zero, then this is not pre-multiplied
                        If Not inStream(x - 1, y) = 0 Then
                            lPARGB = 1 ' not premultiplied
                        ElseIf Not inStream(x - 2, y) = 0 Then
                            lPARGB = 1
                        ElseIf Not inStream(x - 3, y) = 0 Then
                            lPARGB = 1
                        End If
                        ' but don't exit loop until we know if any alphas are non-zero
                    End If
                    zeroCount = zeroCount + 1 ' helps in decision factor at end of loop
                Case 255
                    ' no way to indicate if premultiplied or not, unless...
                    If lPARGB = 1 Then
                        lPARGB = 2    ' not pre-multiplied because of the zero check above
                        Exit For
                    End If
                    opaqueCount = opaqueCount + 1
                Case Else
                    ' if any Exit For's below get triggered, not pre-multiplied
                    If lPARGB = 1 Then
                        lPARGB = 2: Exit For
                    ElseIf inStream(x - 3, y) > inStream(x, y) Then
                        lPARGB = 2: Exit For
                    ElseIf inStream(x - 2, y) > inStream(x, y) Then
                        lPARGB = 2: Exit For
                    ElseIf inStream(x - 1, y) > inStream(x, y) Then
                        lPARGB = 2: Exit For
                    Else
                        bAlpha = AlphaComplex
                    End If
                End Select
            Next
            If lPARGB = 2 Then Exit For
        Next
        ' if we got all the way thru the image without hitting Exit:For then
        ' the image is not alpha unless the bAlpha flag was set in the loop
        If zeroCount = (x \ 4) * (UBound(inStream, 2) + 1) Then ' every alpha value was zero
            bPARGB = False: bAlpha = AlphaNone ' assume RGB, else 100% transparent ARGB
            ' also if lPARGB=0, then image is completely black
        ElseIf opaqueCount = (x \ 4) * (UBound(inStream, 2) + 1) Then ' every alpha is 255
            bPARGB = False: bAlpha = AlphaNone
        Else
            Select Case lPARGB
                Case 2: bPARGB = False ' 100% positive ARGB
                Case 1: bPARGB = False ' now 100% positive ARGB
                Case 0: bPARGB = True
            End Select
            If bAlpha = AlphaNone Then bAlpha = AlphaSimple
        End If
    End If
    ' see if caller wants the non-premultiplied alpha channel premultiplied
    If bAlpha Then
        If bPARGB Then ' else force premultiplied
            imgType = imgBmpPARGB
        Else
            imgType = imgBmpARGB
            If bPreMultiply = True Then
                bAlpha = AlphaSimple
                For y = 0 To UBound(inStream, 2)
                    For x = 3 To UBound(inStream, 1) Step 4
                        If inStream(x, y) = 0 Then
                            CopyMemory inStream(x - 3, y), 0&, 4&
                        ElseIf Not inStream(x, y) = 255 Then
                            bAlpha = AlphaComplex
                            For lPARGB = x - 3 To x - 1
                                inStream(lPARGB, y) = ((0& + inStream(lPARGB, y)) * inStream(x, y)) \ &HFF
                            Next
                        End If
                    Next
                Next
            End If
        End If
    Else
        imgType = imgBitmap
        If bPreMultiply = True Then
            For y = 0 To UBound(inStream, 2)
                For x = 3 To UBound(inStream, 1) Step 4
                    inStream(x, y) = 255
                Next
            Next
        End If
    End If
    bIsAlpha = bAlpha
End Sub
Public Function FindColor(ByRef PaletteItems() As Long, ByVal Color As Long, ByVal Count As Long, ByRef isNew As Boolean) As Long
' MODIFIED BINARY SEARCH ALGORITHM -- Divide and conquer.
' Binary search algorithms are about the fastest on the planet, but
' its biggest disadvantage is that the array must already be sorted.
' Ex: binary search can find a value among 1 million values between 1 and 20 iterations

' [in] PaletteItems(). Long Array to search within. Array must be 1-bound, sorted ascending
' [in] Color. A value to search for.
' [in] Count. Number of items in PaletteItems() to compare against
' [out] isNew. If Color not found, isNew is True else False
' [out] Return value: The Index where Color was found or where the new Color should be inserted

Dim ub As Long, lb As Long
Dim newIndex As Long
    
    If Count = 0& Then FindColor = 1&: isNew = True: Exit Function
    ub = Count: lb = 1&
    Do Until lb > ub
        newIndex = lb + ((ub - lb) \ 2&)
        If PaletteItems(newIndex) = Color Then
            Exit Do
        ElseIf PaletteItems(newIndex) > Color Then ' new color is lower in sort order
            ub = newIndex - 1&
        Else ' new color is higher in sort order
            lb = newIndex + 1&
        End If
    Loop
    If lb > ub Then  ' color was not found
        If Color > PaletteItems(newIndex) Then newIndex = newIndex + 1&
        isNew = True
    Else
        isNew = False
    End If
    FindColor = newIndex
End Function
Public Sub GrayScaleRatios(Formula As eGrayScaleFormulas, r As Single, g As Single, b As Single)
' note: when adding your own formulas, ensure they add up to 1.0 or less;
' else unexpected colors may be calculated. Exception: non-grayscale are always: 1,1,1
    Select Case Formula
    Case gsclNone:          r = 1: g = 1: b = 1             ' no grayscale
    Case gsclNTSCPAL:       r = 0.299: g = 0.587: b = 0.114 ' standard weighted average
    Case gsclSimpleAvg:     r = 0.333: g = 0.334: b = r     ' pure average
    Case gsclCCIR709:       r = 0.213: g = 0.715: b = 0.072 ' Formula.CCIR 709, Default
    Case gsclRedMask:       r = 0.8: g = 0.1: b = g         ' personal preferences: could be r=1:g=0:b=0 or other weights
    Case gsclGreenMask:     r = 0.1: g = 0.8: b = r         ' personal preferences: could be r=0:g=1:b=0 or other weights
    Case gsclBlueMask:      r = 0.1: g = r: b = 0.8         ' personal preferences: could be r=0:g=0:b=1 or other weights
    Case gsclBlueGreenMask: r = 0.1: g = 0.45: b = g        ' personal preferences: could be r=0:g=.5:b=.5 or other weights
    Case gsclRedGreenMask:  r = 0.45: g = r: b = 0.1        ' personal preferences: could be r=.5:g=.5:b=0 or other weights
    Case Else:              r = 0.299: g = 0.587: b = 0.114 ' use gsclNTSCPAL
    End Select
End Sub
Public Function ArrayToPicture(inArray() As Byte, Optional Offset, Optional Size) As IPictureDisp
' function creates a stdPicture from the passed array
' Note: The array was already validated as not empty when calling class' LoadStream was called
Dim lOffset As Long, lSize As Long
Dim o_hMem  As Long
Dim o_lpMem  As Long
Dim aGUID(0 To 3) As Long
Dim IIStream As IUnknown
    
    If IsMissing(Offset) Then lOffset = LBound(inArray) Else lOffset = CLng(Offset)
    If IsMissing(Size) Then lSize = UBound(inArray) - LBound(inArray) + 1 Else lSize = CLng(Size)
    
    aGUID(0) = &H7BF80981 ' &H7BF80980  ' GUID for stdPicture
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    o_hMem = GlobalAlloc(&H2&, lSize)
    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)
        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, inArray(lOffset), lSize
            Call GlobalUnlock(o_hMem)
            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                  Call OleLoadPicture(IIStream, 0&, 0&, aGUID(0), ArrayToPicture)
            End If
        End If
    End If
End Function
Public Function ArrayProps( _
    ByVal arrayPtr As Long, _
    Optional Dimensions As Long, _
    Optional FirstElementPtr As Long) As Long
' Function returns the overall size of the array in bytes or returns zero
' if the array is uninitialized or contains no elements

' Parameters
'   ArrayPtr :: result of call from GetArrayPointer
'   Dimensions [out] :: number of dimensions for the array
'   FirstElementPtr [out] :: basically VarPtr(first element of array)
Dim tSA As SAFEARRAY2D
Dim lBounds() As Long
Dim x As Long, totalSize As Long
    
    If arrayPtr = 0& Then Exit Function
    CopyMemory arrayPtr, ByVal arrayPtr, 4&
    If arrayPtr = 0& Then Exit Function             ' uninitialized array
    
    CopyMemory ByVal VarPtr(tSA), ByVal arrayPtr, 16&     ' safe array structure minus bounds info
    Dimensions = tSA.cDims
    FirstElementPtr = tSA.pvData
    ReDim lBounds(1 To 2, 1 To Dimensions)
    CopyMemory lBounds(1, 1), ByVal arrayPtr + 16&, Dimensions * 8&
    
    totalSize = 1
    For x = 1 To Dimensions
        totalSize = totalSize * lBounds(1, x)
    Next
    ArrayProps = totalSize * tSA.cbElements
End Function
Public Sub OverlayHost_2DbyHost(aOverlay() As Byte, ptrSafeArray As LongPtr, Host As clsPictureData)
' Routine overlays a BYTE array on top of some memory address. Passing incorrect values will crash the app and maybe the system
' NOTE: Multidimensional arrays are passed right to left. If aOverlay(0 to 9, 0 to 99) were desired: pass ElemCount_Dim1=100:ElemCount_Dim2=10

' aOverlay() is an uninitialized, dynamic Byte array. If initialized, call Erase on the array before passing it
' ptrSafeArray is passed as VarPtr(mySafeArray_Variable). It must point to a structure/array that contains at least 32bytes. Not used if memPtr=0
' nrDims must be 1 or 2. Not used if memPtr=0
' ElemCount_Dim1 is number of array elements in 1st dimension of array. Not used if memPtr=0
' ElemCount_Dim2 is number of array elements in 2nd dimension of array. Not used if memPtr=0 or nrDims=1
' memPtr is the memory address to overlay the array onto
    If ptrSafeArray = 0& Then
        CopyMemory ByVal VarPtrArray(aOverlay), ptrSafeArray, 4&      ' remove overlay
    ElseIf Not Host Is Nothing Then
        If Host.Handle Then
Dim tSA As SAFEARRAY2D
            With tSA
                .cbElements = 1               '1=byte
                .pvData = Host.BitsPointer    'memory address
                .cDims = 2                    'nr of dimensions
                .rgSAbound(0).cElements = Host.Height  'number array items (1st dimension)
                .rgSAbound(1).cElements = Host.ScanWidth 'number array items (2nd dimension)
            End With
            CopyMemory ByVal ptrSafeArray, tSA, 32&    ' copy the safeArray structure to passed pointer
            CopyMemory ByVal VarPtrArray(aOverlay), ptrSafeArray, 4&    ' overlay the array onto the memory address
        End If
    End If
End Sub
Public Sub OverlayHost_Byte(aOverlay() As Byte, ptrSafeArray As LongPtr, nrDims As Long, ElemCount_Dim1 As Long, ElemCount_Dim2 As Long, ByVal memPtr As LongPtr)
' Routine overlays a BYTE array on top of some memory address. Passing incorrect values will crash the app and maybe the system
' NOTE: Multidimensional arrays are passed right to left. If aOverlay(0 to 9, 0 to 99) were desired: pass ElemCount_Dim1=100:ElemCount_Dim2=10

' aOverlay() is an uninitialized, dynamic Byte array. If initialized, call Erase on the array before passing it
' ptrSafeArray is passed as VarPtr(mySafeArray_Variable). It must point to a structure/array that contains at least 32bytes. Not used if memPtr=0
' nrDims must be 1 or 2. Not used if memPtr=0
' ElemCount_Dim1 is number of array elements in 1st dimension of array. Not used if memPtr=0
' ElemCount_Dim2 is number of array elements in 2nd dimension of array. Not used if memPtr=0 or nrDims=1
' memPtr is the memory address to overlay the array onto
    If memPtr = 0& Then
        CopyMemory ByVal VarPtrArray(aOverlay), memPtr, 4&      ' remove overlay
    Else
Dim tSA As SAFEARRAY2D
        With tSA
            .cbElements = 1     '1=byte
            .pvData = memPtr    'memory address
            .cDims = nrDims     'nr of dimensions
            If nrDims = 2 Then
                .rgSAbound(0).cElements = ElemCount_Dim1  'number array items (1st dimension)
                .rgSAbound(1).cElements = ElemCount_Dim2  'number array items (2nd dimension)
            Else
                .rgSAbound(0).cElements = ElemCount_Dim1  'number array items (only one dimension)
            End If
            ' Note: the .LBound members of .rgSABound are always zero. Set them on routine's return if needed. Remember right to left order
        End With
        CopyMemory ByVal ptrSafeArray, tSA, 32&                     ' copy the safeArray structure to passed pointer
        CopyMemory ByVal VarPtrArray(aOverlay), ptrSafeArray, 4&    ' overlay the array onto the memory address
    End If
End Sub
Public Sub OverlayHost_Long( _
    aOverlay() As Long, _
    ptrSafeArray As Long, nrDims As Long, _
    ElemCount_Dim1 As Long, ElemCount_Dim2 As Long, _
    ByVal memPtr As Long)
' Routine overlays a LONG array on top of some memory address. Passing incorrect values will crash the app and maybe the system
' NOTE: Multidimensional arrays are passed right to left. If aOverlay(0 to 9, 0 to 99) were desired: pass ElemCount_Dim1=100:ElemCount_Dim2=10

' aOverlay() is an uninitialized, dynamic Long array. If initialized, call Erase on the array before passing it
' ptrSafeArray is passed as VarPtr(mySafeArray_Variable). It must point to a structure/array that contains at least 32bytes. Not used if memPtr=0
' nrDims must be 1 or 2. Not used if memPtr=0
' ElemCount_Dim1 is number of array elements in 1st dimension of array. Not used if memPtr=0
' ElemCount_Dim2 is number of array elements in 2nd dimension of array. Not used if memPtr=0 or nrDims=1
' memPtr is the memory address to overlay the array onto
    If memPtr = 0& Then
        CopyMemory ByVal VarPtrArray(aOverlay), memPtr, 4&      ' remove overlay
    Else
Dim tSA As SAFEARRAY2D
        With tSA
            .cbElements = 4     '4=long
            .pvData = memPtr    'memory address
            .cDims = nrDims     'nr of dimensions
            If nrDims = 2 Then
                .rgSAbound(0).cElements = ElemCount_Dim1  'number array items (1st dimension)
                .rgSAbound(1).cElements = ElemCount_Dim2  'number array items (2nd dimension)
            Else
                .rgSAbound(0).cElements = ElemCount_Dim1  'number array items (only one dimension)
            End If
            ' Note: the .LBound members of .rgSABound are always zero. Set them on routine's return if needed. Remember right to left order
        End With
        CopyMemory ByVal ptrSafeArray, tSA, 32&    ' copy the safeArray structure to passed pointer
        CopyMemory ByVal VarPtrArray(aOverlay), ptrSafeArray, 4&    ' overlay the array onto the memory address
    End If
End Sub
Private Function p_CreatePartialRegion(rgnRects() As RECT, lIndex As Long, uIndex As Long, leftOffset As Long, cX As Long) As Long
' Helper function for CreateShapedRegion & p_CreateWin98Region
' Called to create a region in its entirety or stepped (see p_CreateWin98Region)
    On Error Resume Next
    ' Note: Ideally contiguous rectangles of equal height & width should be combined
    ' into one larger rectangle. However, thru trial & error I found that Windows
    ' does this for us and taking the extra time to do it ourselves
    ' is too cumbersome & slows down the results.
    
    ' the first 32 bytes of a region is the header describing the region.
    ' Well, 32 bytes equates to 2 rectangles (16 bytes each), so I'll
    ' cheat a little & use rectangles to store the header
    With rgnRects(lIndex - 2) ' bytes 0-15
        .Left = 32&                     ' length of region header in bytes
        .Top = 1&                       ' required cannot be anything else
        .Right = uIndex - lIndex + 1&   ' number of rectangles for the region
        .Bottom = .Right * 16&          ' byte size used by the rectangles; can be zero
    End With
    With rgnRects(lIndex - 1&) ' bytes 16-31 bounding rectangle identification
        .Left = leftOffset                  ' left
        .Top = rgnRects(lIndex).Top         ' top
        .Right = leftOffset + cX            ' right
        .Bottom = rgnRects(uIndex).Bottom   ' bottom
    End With
    ' call function to create region from our byte (RECT) array
    p_CreatePartialRegion = ExtCreateRegion(ByVal 0&, (rgnRects(lIndex - 2&).Right + 2&) * 16&, rgnRects(lIndex - 2&))
    If Err Then Err.Clear
End Function
Private Function p_CreateWin98Region(rgnRects() As RECT, rectCount As Long, leftOffset As Long, cX As Long) As Long
' Fall-back routine when a very large region fails to be created.
' Win98 has problems with regional rectangles over 4000
' So, we'll try again in case this is the prob with other systems too.
' We'll step it at 2000 at a time which is stil very quick
Dim x As Long, y As Long ' loop counters
Dim win98Rgn As Long     ' partial region
Dim rtnRegion As Long    ' combined region & return value of this function
Const RGN_OR As Long = 2&
Const scanSize As Long = 2000&
    ' we start with 2 'cause first 2 RECTs are the header
    For x = 2& To rectCount Step scanSize
    
        If x + scanSize > rectCount Then
            y = rectCount
        Else
            y = x + scanSize
        End If
        
        ' attempt to create partial region, scanSize rects at a time
        win98Rgn = p_CreatePartialRegion(rgnRects(), x, y, leftOffset, cX)
        
        If win98Rgn = 0& Then    ' failure
            ' cleaup combined region if needed
            If Not rtnRegion = 0& Then DeleteObject rtnRegion
            Exit For ' abort; system won't allow us to create the region
        Else
            If rtnRegion = 0& Then ' first time thru
                rtnRegion = win98Rgn
            Else ' already started
                ' use combineRgn, but only every scanSize times
                CombineRgn rtnRegion, rtnRegion, win98Rgn, RGN_OR
                DeleteObject win98Rgn
            End If
        End If
    Next
    ' done; return result
    p_CreateWin98Region = rtnRegion
End Function
Public Function BlendImageToColor(cHost As clsPictureData, ByVal FillColor As Long, outStream() As Byte, Optional ByVal bitDepth24 As Boolean = False) As Boolean
' Function basically renders an image over a solid bkg color
' Function called from SaveToFlle/Stream_JPG & BMP when the image to be saved
' has premultiplied pixels.
Dim x As Long, y As Long
Dim r As Byte, g As Byte, b As Byte
Dim pAlpha As Byte, dAlpha As Long
Dim tSA As SAFEARRAY2D, srcBytes() As Byte
Dim ScanWidth As Long, DestX As Long
    ScanWidth = cHost.ScanWidth ' cache vs recalculating each scan line
    OverlayHost_2DbyHost srcBytes(), VarPtr(tSA), cHost
    ' extract individual RGB values & convert FillColor
    r = (FillColor And &HFF)
    g = ((FillColor \ &H100) And &HFF)
    b = ((FillColor \ &H10000) And &HFF)
    FillColor = r * &H10000 Or (FillColor And &HFF00&) Or b
    
    If bitDepth24 = False Then ' requesting to blend & save as 32bpp (GDI+ JPG routine calls this)
        ReDim outStream(0 To ScanWidth - 1, 0 To cHost.Height - 1)
        For y = 0 To cHost.Height - 1
            For x = 0 To ScanWidth - 1& Step 4&
                pAlpha = srcBytes(x + 3&, y)
                If pAlpha = 255 Then
                    CopyMemory outStream(x, y), srcBytes(x, y), 4&
                ElseIf pAlpha = 0 Then
                    CopyMemory outStream(x, y), FillColor, 4&
                Else ' blend to backcolor
                    dAlpha = &HFF& - pAlpha
                    outStream(x, y) = (dAlpha * b) \ &HFF + srcBytes(x, y)
                    outStream(x + 1&, y) = (dAlpha * g) \ &HFF + srcBytes(x + 1&, y)
                    outStream(x + 2&, y) = (dAlpha * r) \ &HFF + srcBytes(x + 2&, y)
                    outStream(x + 3&, y) = 255 ' indicate pixel is fully opaque now
                End If
            Next
        Next
    Else ' requesting to blend as a 24bpp image
        ReDim outStream(0 To ByteAlignOnWord(24, cHost.Width) - 1, 0 To cHost.Height - 1)
        For y = 0 To cHost.Height - 1
            DestX = 0&
            For x = 0 To ScanWidth - 1& Step 4&
                pAlpha = srcBytes(x + 3&, y)
                If pAlpha = 255 Then
                    CopyMemory outStream(DestX, y), srcBytes(x, y), 3&
                ElseIf pAlpha = 0 Then
                    CopyMemory outStream(DestX, y), FillColor, 3&
                Else ' blend to backcolor
                    dAlpha = &HFF& - pAlpha
                    outStream(DestX, y) = (dAlpha * b) \ &HFF + srcBytes(x, y)
                    outStream(DestX + 1&, y) = (dAlpha * g) \ &HFF + srcBytes(x + 1&, y)
                    outStream(DestX + 2&, y) = (dAlpha * r) \ &HFF + srcBytes(x + 2&, y)
                End If
                DestX = DestX + 3&
            Next
        Next
    End If
    OverlayHost_2DbyHost srcBytes(), 0&, Nothing
End Function
Public Function IsArrayEmptyP(FarPointer As Long) As Long
' test to see if an array has been initialized
    CopyMemory IsArrayEmptyP, ByVal FarPointer, 4&
End Function
Public Function IsArrayEmpty(Arr As Variant) As Boolean
Dim lb As Long, ub As Long
    Err.Clear
    On Error Resume Next
    If IsArray(Arr) = False Then IsArrayEmpty = True
    ub = UBound(Arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
        Err.Clear: lb = LBound(Arr)
        IsArrayEmpty = lb > ub
    End If
End Function
Private Function p_ArrayDibToBmp(aData() As Byte) As Boolean
' make Bitmap byte array from Dib byte array
Dim Result As Boolean ': Result = False
    On Error GoTo HandleError
Dim ldibSize As Long: ldibSize = UBound(aData) - LBound(aData) + 1
    If ldibSize < BITMAPINFOHEADERSIZE Then Err.Raise vbObjectError + 512
Dim lfilSize As Long: lfilSize = BITMAPFILEHEADERSIZE + ldibSize     ' return BMP
Dim aTemp() As Byte: ReDim aTemp(0 To (lfilSize - 1))
    CopyMemory aTemp(0), &H4D42, 2                                   ' BM
    CopyMemory aTemp(2), lfilSize, 4                                 ' bfSize
    CopyMemory aTemp(&HA), &H76&, 4                                  ' bfOffBits
    CopyMemory aTemp(BITMAPFILEHEADERSIZE), aData(0), ldibSize       ' DIB data
    aData = aTemp
    Result = True
HandleExit:  p_ArrayDibToBmp = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
#End If             'ObjectDataType


Attribute VB_Name = "modForms"
Option Explicit
Option Compare Database
'=========================
Private Const c_strModule As String = "modForms"
'=========================
' Описание      : Функции для работы с формами/отчетами и крнтролами Access
' Версия        : 1.0.20.453846943
' Дата          : 02.04.2024 16:39:48
' Автор         : Кашкин Р.В. (KashRus@gmail.com)
' Примечание    :
' v.1.0.20      : 08.08.2021 - изменения в AccControlLocation - добавлена возможность возвращать клиентские/экранные координаты
' v.1.0.19      : 18.07.2019 - добавлена CreateHFontByControl - создание hFont на основе параметров контрола
' v.1.0.18      : 16.04.2019 - добавлена AccControlLocation - получения позиции контрола на экране или в клиентской  области окна (в пикселях)
' v.1.0.17      : 19.10.2012 -
'=========================
' ToDo: очень много древнего мусора который просто неохота перебирать
'-------------------------
#Const UseModPictureData = 1    ' признак использования модуля вывода изображений в контрол ModPictureData
#Const UseContextMenu = 1       ' признак использования классов вызова контекстного меню clsContextMenu
#Const UseAccObjectExClass = 0  ' для хранения доп свойств, прочитанных из тэгов используется:
'                            0 - тип   typAccObjectEx
'                            1 - класс clsAccObjectEx
Public Const twMinLim = 10  ' значения меньше данного считаются процентами, больше - твипами

Public Const pxWinGap = 16                          ' смещение окон в px при размещении каскадом
' для перевода Width/Height >> InsideWidth/InsideHeight
Public Const pxScrBar = 17                          ' 17px высота ScrollBars
Public Const pxNavBut = 18                          ' 18px высота NavigationButtons
Public Const pxRecSel = 19                          ' 19px ширина RecordSelectors
'-------------------------
' константы имён полей служебных таблиц
'-------------------------
Public Const c_strKey = "ID"            '
Public Const c_strParent = "PARENT"     '
Public Const c_strName = "NAME"         '
Public Const c_strCName = "CNAME"       '
Public Const c_strFaceKey = "FaceId"    ' имя поля кода картинки для PictureData_SetToControl используется в частности при создании меню с помощью CreateItemsFromSQL
'-------------------------
' константы преобразования размеров
'-------------------------
Public Const DotsPerInch = 96
Public Const PointsPerInch = 72
Public Const TwipsPerInch = 1440
Public Const CentimitersPerInch = 2.54              '1 дюйм = 127 / 50 см
Public Const HimetricPerInch = 2540                 '1 дюйм = 1000 * 127/50 himetrix
'
Public Const cm = TwipsPerInch / CentimitersPerInch '1 см = 566.929133858 twips
Public Const inch = TwipsPerInch                    '1 дюйм = 1440 twips
Public Const pt = TwipsPerInch / PointsPerInch      '1 пункт = 20 twips
Public Const px = TwipsPerInch / DotsPerInch        '1 пиксель = 15 twips
'-------------------------
Public Enum eDirection
    DIRECTION_HORIZONTAL = 0
    DIRECTION_VERTICAL = 1
End Enum
Public Enum acFormState         ' состояние формы
    acObjStateClosed = 0            ' Closed
    acObjStateOpen = 1              ' Open
    acObjStateDirty = 2             ' Changed but not saved
    acObjStateNew = 4               ' New
End Enum
Public Enum ePosition           ' константы положения на области (для конструирования выравнивания и расположения)
    ePosUndef = 0                   ' не задано
    eLeft = 1                       ' по левому краю
    eRight = 2                      ' по правому краю
    eTop = 4                        ' по верхнему краю
    eBottom = 8                     ' по нижнему краю
    eWidth = eLeft + eRight         ' по ширине (центр по горизонтали)
    eCenterHorz = eWidth
    eHeight = eTop + eBottom        ' по высоте (центр по вертикали)
    eCenterVert = eHeight
    eCascade = &H100                ' композиция (только для формы ??)
End Enum
Public Enum eAlign              ' выравнивание внутри области
    eAlignUndef = 0                 ' не задано
    ' 2 направления по 3 положения вдоль каждого
    ' ИТОГО: 3x3 = 9 точек выравнивания.
    eAlignLeftTop = eLeft + eTop                ' по левому верхнему краю
    eAlignRightTop = eRight + eTop              ' по правому верхнему краю
    eAlignLeftBottom = eLeft + eBottom          ' по левому нижнему краю
    eAlignRightBottom = eRight + eBottom        ' по правому нижнему краю
    eCenterHorzTop = eCenterHorz + eTop         ' по верхнему краю центрировать по горизонтали
    eCenterHorzBottom = eCenterHorz + eBottom   ' по нижнему краю центрировать по горизонтали
    eCenterVertLeft = eLeft + eCenterVert       ' по левому краю центрировать по вертикали
    eCenterVertRight = eRight + eCenterVert     ' по правому краю центрировать по вертикали
    eCenter = eCenterHorz + eCenterVert         ' центрировать посередине области
End Enum
Public Enum ePlace              ' размещение Obj2 относительно Obj1
    ' 2 области по 9 точек привязки на каждой: LT,LC,LB,CB,RB,RC,RT,CT,CC
    ' ИТОГО: 9x9 = 81 вариант привязки.
    ' Большинство из них бессмысленны, поэтому констант под них нет,
    ' но возможны путем комбинации констант по схеме:
    ' =H2+V2+H1+V1, где:
    ' Obj1 (к которому привязывают) - биты 0-3:  L1=1,  R1=2,  T1=4,  B1=8
    '   H1 - положение по горизонтали точки привязки на Obj1
    '       ={eLeft|eRight|eCenterHorz}
    '   V1 - положение по вертикали точки привязки на Obj1
    '       ={eTop|eBottom|eCenterVert}
    ' Obj2 (который привязывают)    - биты 4-8:  L2=16, R2=32, T2=64, B2=128
    '   H2 - положение по горизонтали точки привязки на Obj2
    '       ={eLeft|eRight|eCenterHorz} * 16
    '   V2 - положение по вертикали точки привязки на Obj2
    '       ={eTop|eBottom|eCenterVert} * 16
    ePlaceUndef = 0     ' по-умолчанию = 222 -> ePlaceOnRight - снаружи справа по центру
' внутри по центру
    ePlaceCenter = eCenter * 16 + eCenter                           ' по центру (внутри)
    ePlaceToLeft = eCenterVertLeft * 16 + eCenterVertLeft           ' внутри слева по центру
    ePlaceToRight = eCenterVertRight * 16 + eCenterVertRight        ' внутри справа по центру
    ePlaceToTop = eCenterHorzTop * 16 + eCenterHorzTop              ' внутри по центру сверху
    ePlaceToBottom = eCenterHorzBottom * 16 + eCenterHorzBottom     ' внутри по центру снизу
' снаружи по центру
    ePlaceOnLeft = eCenterVertRight * 16 + eCenterVertLeft          ' снаружи слева по центру
    ePlaceOnRight = eCenterVertLeft * 16 + eCenterVertRight         ' снаружи справа по центру
    ePlaceOnTop = eCenterHorzBottom * 16 + eCenterHorzTop           ' снаружи по центру сверху
    ePlaceOnBottom = eCenterHorzTop * 16 + eCenterHorzBottom        ' снаружи по центру снизу
' внутри по краю
    ePlaceToLeftTop = eAlignLeftTop * 16 + eAlignLeftTop            ' внутри слева сверху
    ePlaceToRightTop = eAlignRightTop * 16 + eAlignRightTop         ' внутри справа сверху
    ePlaceToLeftBottom = eAlignLeftBottom * 16 + eAlignLeftBottom   ' внутри слева снизу
    ePlaceToRightBottom = eAlignRightBottom * 16 + eAlignRightBottom ' внутри справа снизу
' снаружи по краю
    ePlaceOnLeftToTop = eAlignRightTop * 16 + eAlignLeftTop         ' снаружи слева к верхнему краю
    ePlaceOnLeftToBottom = eAlignRightBottom * 16 + eAlignLeftBottom ' снаружи слева к нижнему краю
    ePlaceOnRightToTop = eAlignLeftTop * 16 + eAlignRightTop        ' снаружи справа к верхнему краю
    ePlaceOnRightToBottom = eAlignLeftBottom * 16 + eAlignRightBottom ' снаружи справа к нижнему краю
    ePlaceOnTopToLeft = eAlignLeftBottom * 16 + eAlignLeftTop       ' снаружи к левому краю сверху
    ePlaceOnTopToRight = eAlignRightBottom * 16 + eAlignRightTop    ' снаружи к правому краю сверху
    ePlaceOnBottomToLeft = eAlignLeftTop * 16 + eAlignLeftBottom    ' снаружи к левому краю снизу
    ePlaceOnBottomToRight = eAlignRightTop * 16 + eAlignRightBottom ' снаружи к правому краю снизу
' композиция (только для формы ??)
    eCascadeFromLeftTop = eCascade + ePlaceToLeftTop                ' размещение каскадом вправо-вниз
    eCascadeFromRightTop = eCascade + ePlaceToRightTop              ' размещение каскадом влево-вниз
    eCascadeFromLeftBottom = eCascade + ePlaceToLeftBottom          ' размещение каскадом вправо-вверх
    eCascadeFromRightBottom = eCascade + ePlaceToRightBottom        ' размещение каскадом влево-вверх
End Enum

Public Enum eObjSizeMode                    ' масштабирование объектов
    apObjSizeZoomDown = -1                  '-1 - пропорциональное масштабирование (только уменьшение)
    apObjSizeClip = acOLESizeClip           ' 0 - не меняем размер. если размер больше области вывода - обрезка
    apObjSizeStretch = acOLESizeStretch     ' 1 - сжатие/растяжение (нарушает пропорции)
    'apObjSizeAutoSize = acOLESizeAutoSize   ' 2 - ???
    apObjSizeZoom = acOLESizeZoom           ' 3 - пропорциональное масштабирование
End Enum

Public Enum eControlScale   ' масштабирование контрола
    csYes = -1                  ' масштабировать контрол
    csNo = 0                    ' не масштабировать контрол
    csDefault = 1               ' использовать значение по-умолчанию для формы
End Enum
Public Enum eScaleWhen      ' когда масштабировать контрол
    scNo = 0                    ' никогда
    scYes = -1                  ' при изменении размеров
    scAtLoad = 1                ' только при первоначальной загрузке
End Enum
Public Enum eScaleType      ' тип масштабирования формы
    sfNo = 0                    ' не перестраивать форму при изменении окна формы
    sfSequent = -1              ' последовательно изменять размеры секций при изменении окна формы
    sfProp = 1                  ' изменять размеры секций пропорционально изменению окна формы
End Enum
Public Enum eControlSize    ' тип изменения размеров контрола
    czNone = 0                  ' не растягивать
    czRight = 1                 ' растягивать вправо
    czBottom = 2                ' растягивать вниз
    czBoth = 3                  ' растягивать вправо-вниз
End Enum
Public Enum eControlFloat   ' тип привязки контрола
    cfNone = 0                  ' нет привязки (привязка к левому верхнему)
    cfRight = 1                 ' привязка к правому (верхнему) краю
    cfBottom = 2                ' привязка к (левому) нижнему краю
    cfBoth = 3                  ' привязка к правому-нижнему краю
End Enum

Public Enum eObjectStyle   ' стиль отображения контрола контрола
    lsNone = 0                  ' позиция и размер не меняются (нет необходимости в использовании)
' основные стили привязки
    lsLeft = 1                  ' привязан к левой границе области
    lsRight = 2                 ' привязан к правой стороне области
    lsTop = 4                   ' привязан к верхей границе области
    lsBottom = 8                ' привязан к нижней границе области
' дополнительные стили привязки
    lsHorz = 3                  ' зафиксирован между правыми и левыми границами области (растягивается горизонтально)
    lsLeftRight = 3
    lsLeftTop = 5               ' привязан к левому верхнему углу области
    lsRightTop = 6              ' привязан к правому верхнему углу области
    lsHorzTop = 7               ' привязан к верхней границе и зафиксирован между правыми и левыми границами области (растягивается горизонтально)
    lsLeftBottom = 9            ' привязан к левому нижнему углу области
    lsRightBottom = 10          ' привязан к правому нижнему углу области
    lsHorzBottom = 11           ' привязан к нижней границе и зафиксирован между правыми и левыми границами области (растягивается горизонтально)
    lsVert = 12                 ' зафиксирован между верхней и нижней границами области (растягивается вертикально)
    lsTopBottom = 3
    lsVertLeft = 13             ' привязан к левой границе и зафиксирован между верхней и нижней границами области (растягивается вертикально)
    lsVertRight = 14            ' привязан к правой границе и зафиксирован между верхней и нижней границами области (растягивается вертикально)
    lsFull = 15                 ' растягивается вертикально и горизонтально по правой и нижней границе области
    lsLeftRightTopBottom = 15
' стили относительных размеров
    lsXProp = 16                ' позиция Left зависит от ширины области (распределение по ширине)
    lsRProp = 32                ' позиция Right зависит от высоты области (распределение по высоте)
    lsYProp = 64                ' позиция Top зависит от высоты области (распределение по высоте)
    lsBProp = 128               ' позиция Bottom зависит от ширины области (распределение по ширине)
    lsWProp = lsXProp + lsRProp ' ширина пропорциональна ширине области
    lsHProp = lsYProp + lsBProp ' высота пропорциональна высоте области
' стили вывода изображения/текста
    lsShowIcon = 1024           ' выводить только иконку
    lsShowText = 2048           ' выводить только подпись
    lsShowIconText = 3072       ' выводить иконку и подпись
End Enum
Public Enum eControlSplit   ' тип разделителя
    cdNone = 0                  ' не является сплиттером
    cdVert = 1                  ' вертикальный сплиттер
    cdHorz = 2                  ' горизонтальный сплиттер
'    cdBoth = 3                 ' ??? это как ???
End Enum
Public Enum eAlignText ' выравнивание текста
    TA_LEFT = 0                 ' Опорная точка находится на левой кромке рабочего прямоугольника.
    TA_RIGHT = 2                ' Опорная точка находится на правой кромке рабочего прямоугольника.
    TA_CENTER = 6               ' Опорная точка выравнивается горизонтально по центру рабочего прямоугольника.
    TA_TOP = 0                  ' Опорная точка на верхней кромке рабочего прямоугольника.
    TA_BOTTOM = 8               ' Опорная точка на нижней кромке рабочего прямоугольника.
    TA_BASELINE = 24            ' Опорная точка находится на базовой линии текста.
    TA_RTLREADING = 256         ' Редакция Windows на языках Ближнего Востока: Текст размечается для порядка чтения справа налево , в противоположность порядку чтения по умолчанию слева направо. Это применяется только тогда, когда шрифт, выбранный в контекст устройства предназначен или для Еврейского или для Арабского языка.
'    TA_NOUPDATECP               ' Текущая позиция не модифицируется после каждого вызова вывода текста.
'    TA_UPDATECP                 ' Текущая позиция модифицируется после каждого вызова вывода текста.
'    TA_MASK  = (TA_BASELINE + TA_CENTER + TA_UPDATECP)
End Enum
'Public Enum eObjSizeMode    ' режим изменения размеров изображения
'    apObjSizeClip = acOLESizeClip           ' 0 - не меняем размер. если размер больше области вывода - обрезка
'    apObjSizeStretch = acOLESizeStretch     ' 1 - сжатие/растяжение (нарушает пропорции)
'    ' 2 'acOLESizeAutoSize
'    apObjSizeZoom = acOLESizeZoom           ' 3 - пропорциональное масштабирование
'    apObjSizeZoomDown = -1                  '-1 - пропорциональное масштабирование, только уменьшаем
'End Enum
Private Type tControlEx
' положение
    ' привязка контрола
    BondLeft As Object
    BondTop As Object
    BondRight As Object
    BondBottom As Object
    BondWidth As Object
    BondHeight As Object
    ' положение по-умолчанию
    DefLeft As Single
    DefTop As Single
    DefRight As Single
    DefBottom As Single
    DefWidth As Single
    DefHeight As Single
    ' предельные размеры
    MinWidth As Single
    MaxWidth As Single
    MinHeight As Single
    MaxHeight As Single
' стиль оформления
    TextColor As Long
    BackColor As Long
    BorderColor As Long
' изображение/текст
    PictName As String
    PictSize As Long
    PictAngle As Single
    PictGray As Boolean
    PictMode As eObjSizeMode
    PictAlign As eAlign
    PictText As String
    TextAngle As Single
    TextAlign As eAlign
    TextPlace As ePlace
    FontName As String
    FontSize As Long
    IconName As String
' стиль отображения                                             (adhcSizeIt/adhcFloatIt/adhcStyleIt)
    StyleIt As eObjectStyle ' растяжение/смещение
    ScaleIt As eScaleType   ' масштабирование
' поведение
    SplitIt As eControlSplit
    Action As String ' заданное для контрола действие
End Type
''--------------------------------------------------------------------------------
Public Enum eFieldFormat
    vbInteger = 2
    vbLong = 3
    vbByte = 17
    vbDate = 7
    vbDateTimeJ = 77
    vbSingle = 4
    vbDouble = 5
    vbCurrency = 6
    vbString = 8
    vbBoolean = 11
End Enum

''--------------------------------------------------------------------------------
'' POINTER
''--------------------------------------------------------------------------------
'#If VBA7 = 0 Then       'LongPtr trick by @Greedo (https://github.com/Greedquest)
'Private Enum LongPtr
'    [_]
'End Enum
'#End If
'#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>  PtrSafe, LongPtr and LongLong
'Private Const PTR_LENGTH As Long = 8
'Private Const VARIANT_SIZE As Long = 24
'#Else                   '<OFFICE97-2010>        Long
'Private Const PTR_LENGTH As Long = 4
'Private Const VARIANT_SIZE As Long = 16
'#End If                 '<WIN32>

Private Type RECT ' Store rectangle coordinates.
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINT
    x As Long
    y As Long
End Type
'--------------------------------------------------------------------------------
' Управление курсором
'--------------------------------------------------------------------------------
Public Const IDC_ARROW = 32512&        ' Основной режим
'Public Const IDC_IBEAM = 32513&        ' Выделение текста
'Public Const IDC_WAIT = 32514&         ' Система недоступна
'Public Const IDC_CROSS = 32515&        ' Графическое выделение
'Public Const IDC_UPARROW = 32516&      ' Специальное выделение
'
'Public Const IDC_SIZE = 32640&         '
'Public Const IDC_ICON = 32641&         '
'Public Const IDC_SIZENWSE = 32642&     ' Изменение размеров по диагонали1
'Public Const IDC_SIZENESW = 32643&     ' Изменение размеров по диагонали2
Public Const IDC_SIZEWE = 32644&       ' Изменение горизонтальных размеров
Public Const IDC_SIZENS = 32645&       ' Изменение вертикальных размеров
Public Const IDC_SIZEALL = 32646&      ' Перемещение
'Public Const IDC_NO = 32648&           ' Операция невозможна
'Public Const IDC_HAND = 32649&         ' Изображение руки (гиперссылка)
'Public Const IDC_APPSTARTING = 32650&  ' Фоновый режим (песочные часы)

#If Win64 Then
Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Public Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
Public Declare PtrSafe Function LoadCursorByNum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As Long) As LongPtr
#Else
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
Public Declare Function LoadCursorByNum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As Long) As LongPtr
#End If

'--------------------------------------------------------------------------------
' KERNEL32
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>      PtrSafe, LongPtr and LongLong
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
#Else                   '<WIN32>                    Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
' USER32
'--------------------------------------------------------------------------------
'Private Type SCROLLBARINFO
'    cbSize As Long
'    rcScrollBar As RECT
'    dxyLineButton As Long
'    xyThumbTop As Long
'    xyThumbBottom As Long
'    reserved As Long
'    rgstate(0 To 5) As Long
'End Type
'Private Type SCROLLINFO
'    cbSize As Long
'    fMask As Long
'    nMin As Long
'    nMax As Long
'    nPage As Long
'    nPos As Long
'    nTrackPos As Long
'End Type
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>  PtrSafe, LongPtr and LongLong
Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hWndLock As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function ClientToScreen Lib "user32.dll" (ByVal hwnd As LongPtr, ByRef lpPoint As POINT) As Long
Private Declare PtrSafe Function ScreenToClient Lib "user32.dll" (ByVal hwnd As LongPtr, ByRef lpPoint As POINT) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Boolean) As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As apiGetWindowLongIndex) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpTextString As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long
'Функции управления скроллбаром
Private Declare PtrSafe Function GetScrollPos Lib "user32" (ByVal hwnd As LongPtr, ByVal nBar As Long) As Long
'Private Declare PtrSafe Function SetScrollPos Lib "user32" (ByVal hWnd As LongPtr, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
'Private Declare PtrSafe Function GetScrollBarInfo Lib "user32" (ByVal hWnd As LongPtr, ByVal idObject As Long, psbi As SCROLLBARINFO) As Long
'Private Declare PtrSafe Function GetScrollInfo Lib "user32" (ByVal hWnd As LongPtr, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
'Private Declare PtrSafe Function SetScrollInfo Lib "user32" (ByVal hWnd As LongPtr, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
#Else                   '<OFFICE97-2010>        Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINT) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Boolean) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As apiGetWindowLongIndex) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpTextString As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long
'Функции управления скроллбаром
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
'Private Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
'Private Declare Function GetScrollBarInfo Lib "user32" (ByVal hWnd As Long, ByVal idObject As Long, psbi As SCROLLBARINFO) As Long
'Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
'Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
' GDI32
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>  PtrSafe, LongPtr and LongLong
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal e As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal q As Long, ByVal PAF As Long, ByVal f As String) As LongPtr
#Else                   '<OFFICE97-2010>        Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal e As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal U As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal cp As Long, ByVal q As Long, ByVal PAF As Long, ByVal f As String) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
Private Enum apiDeviceCapability
    HORZSIZE = 4
    VERTSIZE = 6
    HORZRES = 8
    VERTRES = 10
    LOGPIXELSX = 88        '  Logical pixels/inch in X
    LOGPIXELSY = 90        '  Logical pixels/inch in Y
End Enum

Public Const HWND_DESKTOP = &H0
Public Const SPI_GETWORKAREA = &H30

Public Const ANSI_CHARSET As Long = &H0
Public Const DEFAULT_CHARSET As Long = &H1
Public Const SYMBOL_CHARSET As Long = &H2
Public Const RUSSIAN_CHARSET As Long = &HCC
Public Const OEM_CHARSET As Long = &HFF

Public Const DEFAULT_QUALITY As Long = 0
Public Const DRAFT_QUALITY  As Long = 1
Public Const PROOF_QUALITY  As Long = 2
Public Const NONANTIALIASED_QUALITY  As Long = 3
Public Const ANTIALIASED_QUALITY As Long = 4
Public Const CLEARTYPE_QUALITY As Long = 5

'------------------------------
' классы окон Access
'------------------------------
Public Const accClass = "OMain"                 ' класс окна Access
Public Const accClassChild = "MDIClient"        ' клиентская область окна Access
Public Const accClassBD = "ODb"                 ' окно базы данных
Public Const accClassFormWindow = "OForm"       ' класс окна формы Access
Public Const accClassFormClient = "OFormSub"    ' класс формы Access
Public Const accClassFormPopup = "OFormPopup"   ' класс всплывающей формы Access
Public Const accClassFormChild = "OFormChild"   ' класс подчиненной формы Access
Public Const accClassFormNoClose = "OFormNoClose"
Public Const accClassFormClientChild = "OFEDT"  ' класс подчиненного окна формы Access
Public Const accClassTableClientChild = "OGNUM" ' класс подчиненного окна табличной формы Access
Public Const accClassRecordSlector = "OSUI"     ' класс селектора записей табличной формы Access
Public Const accClassTextbox = "OKttbx"         ' класс текстбокса формы Access

Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2
Public Const MAX_LEN = 255
'The window position
Public Const SWP_NOSIZE = &H1 ' Retains the current size (ignores the cx and cy members).
Public Const SWP_NOMOVE = &H2 ' Retains the current position (ignores the x and y members).
Public Const SWP_NOZORDER = &H4 ' Retains the current Z order (ignores the hwndInsertAfter member).
Public Const SWP_NOREDRAW = &H8 ' Does not redraw changes. If this flag is set, no repainting of any kind occurs. This applies to the client area, the nonclient area (including the title bar and scroll bars), and any part of the parent window uncovered as a result of the window being moved. When this flag is set, the application must explicitly invalidate or redraw any parts of the window and parent window that need redrawing.
Public Const SWP_NOACTIVATE = &H10 ' Does not activate the window. If this flag is not set, the window is activated and moved to the top of either the topmost or non-topmost group (depending on the setting of the hwndInsertAfter member).
Public Const SWP_DRAWFRAME = &H20 ' Draws a frame (defined in the window's class description) around the window. Same as the SWP_FRAMECHANGED flag.
Public Const SWP_FRAMECHANGED = &H20 ' Sends a WM_NCCALCSIZE message to the window, even if the window's size is not being changed. If this flag is not specified, WM_NCCALCSIZE is sent only when the window's size is being changed.
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80 ' Hides the window.
Public Const SWP_NOCOPYBITS = &H100 ' Discards the entire contents of the client area. If this flag is not specified, the valid contents of the client area are saved and copied back into the client area after the window is sized or repositioned.
Public Const SWP_NOOWNERZORDER = &H200 ' Does not change the owner window's position in the Z order.
Public Const SWP_NOREPOSITION = &H200  ' Does not change the owner window's position in the Z order. Same as the SWP_NOOWNERZORDER flag.
Public Const SWP_NOSENDCHANGING = &H400 ' Prevents the window from receiving the WM_WINDOWPOSCHANGING message.

' константы для скролбаров
Public Const SBS_HORZ = &H0&
Public Const SBS_VERT = &H1&
Public Const SBS_SIZEBOX = &H8&
Public Const SB_CTL = 2
Public Const SB_THUMBPOSITION = 4
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10

Enum apiGetWindowLongIndex
    GWL_WNDPROC = -4&
    GWL_HINSTANCE = -6&
    GWL_HWNDPARENT = -8&
    GWL_ID = -12&
    GWL_STYLE = -16&
    GWL_EXSTYLE = -20&
    GWL_USERDATA = -21&
End Enum
Enum apiMouseKeys
    MK_LBUTTON = &H1
    MK_RBUTTON = &H2
    MK_SHIFT = &H4
    MK_CONTROL = &H8
    MK_MBUTTON = &H10
End Enum
Enum apiWheelDeltaConst
    WHEEL_DELTA = 120
End Enum
' ==================================
' Константы для тэгов дополнительных свойств
' ----------------------------------
Public Enum eObjectProps   ' типы дополнительных параметров, которые д.б. прочитаны/записаны из/в Tag
    ltDefault = 0                   ' по-умолчанию сохраняются/извлекаются все доступные доп свойства
    ltAll = &HFFFFFFFF              ' все
' ----------------------------------
' позиция/размеры
' ----------------------------------
    ltLeft = 1                      ' позиция левой границы                                         (adhcDefLeft)
    ltRight = 2                     ' позиция правой границы                                         ...
    ltTop = 4                       ' позиция верхней границы
    ltBottom = 8                    ' позиция нижней границы
    ltWidth = &H10                  ' ширина
    ltHeight = &H20                 ' высота
    ltSizeAll = &H3F                ' все размеры
    
    ltLeftRight = ltLeft + ltRight  ' заданы одновременно левая и правая границы
    ltLeftWidth = ltLeft + ltWidth  ' заданы одновременно левая граница и ширина
    ltRightWidth = ltRight + ltWidth ' заданы одновременно правая граница и ширина
    
    ltBadHorz = ltLeft + ltRight + ltWidth ' не могут быть заданы одновременно Left, Right and Width
    
    ltTopBottom = ltTop + ltBottom  ' заданы одновременно верхняя и нижняя границы
    ltTopHeight = ltTop + ltHeight  ' заданы одновременно верхняя граница и высота
    ltBottomHeight = ltBottom + ltHeight ' заданы одновременно нижняя граница и высота
    
    ltBadVert = ltTop + ltBottom + ltHeight ' не могут быть заданы одновременно Top, Bottom and Height
' ----------------------------------
' стиль/оформление
' ----------------------------------
    ltStyle = &H40                  ' стиль отображения                                             (adhcSizeIt/adhcFloatIt/adhcStyleIt)
    ltColors = &H80                 ' цветовое оформление (Back/Fore/Font/TextColor)                (adhcColor/adhcBackColor/etc)
'' ----------------------------------
'' относительная (пропорциональная) позиция/размеры
'    ' сохраняется относительная координата (дробное число)
'    ' может сохраняться имя контрола относительно позиции/размеров которого заданы позиция/размер
'    ' если имя не задано в качестве базы берутся размеры формы/секции
'' ----------------------------------
'    ltPropLeft = ltLeft * &H100     '                                                               (adhcDefLeft  [& adhcBoundLeft])
'    ltPropRight = ltRight * &H100   '                                                                ...
'    ltPropTop = ltTop * &H100
'    ltPropBottom = ltBottom * &H100
'    ltPropWidth = ltWidth * &H100
'    ltPropHeight = ltHeight * &H100
'' ----------------------------------
'' относительная (смещение) позиция/размеры
'    ' сохраняется относительная координата (+- целое число)
'    ' может сохраняться имя контрола относительно позиции/размеров которого заданы позиция/размер
'    ' если имя не задано в качестве базы берутся размеры формы/секции
'' ----------------------------------
'    ltShiftLeft = ltLeft * &H101    ' = ltPropLeft + ltLeft                                         (adhcDefLeft  [& adhcBoundLeft])
'    ltShiftRight = ltRight * &H101  '                                                                ...
'    ltShiftTop = ltTop * &H101
'    ltShiftBottom = ltBottom * &H101
'    ltShiftWidth = ltWidth * &H101
'    ltShiftHeight = ltHeight * &H101
' ----------------------------------
' действие
' ----------------------------------
    ltSplit = &H4000                ' признак разделителя                                           (adhcSplitIt)
    ltAction = &H8000               ' заданное для контрола действие                                (adhcAction)
' ----------------------------------
' свойства выводимого изображения
' ----------------------------------
    ltPictAll = &HFFFF0000
    ltPictName = &H10000            '                                                               (adhcPictName)
    ltPictSize = &H20000            '                                                               (adhcPictSize)
    ltPictMode = &H40000            '                                                               (adhcPictMode)
    ltPictPlace = &H80000           '                                                               (adhcPictAlign)
    ltPictAngle = &H100000          '                                                               (adhcPictAngle)
    ltPictGray = &H200000           '                                                               (adhcPictGray)
    ltPictText = &H1000000          '                                                               (adhcText)
    ltTextAlign = &H4000000         '                                                               (adhcTextAlign)
    ltTextPlace = &H8000000         '                                                               (adhcTextPlace)
    ltTextAngle = &H10000000        '                                                               (adhcTextAngle)
    ltFontName = &H40000000         '                                                               (adhcDefFontName)
    ltFontSize = &H80000000         '                                                               (adhcDefFontSize)
    ltPictShow = ltPictName + ltPictText
' ----------------------------------
End Enum
' ----------------------------------
' Текстовые константы
' ----------------------------------
Public Const adhcNone = "None"
Public Const adhcDefault = "Default", adhcDef = "Def"
Public Const adhcLeft = "Left", adhcLeft1 = "L", adhcLeft2 = "X"
Public Const adhcRight = "Right", adhcRight1 = "R"
Public Const adhcTop = "Top", adhcTop1 = "T", adhcTop2 = "Y"
Public Const adhcBottom = "Bottom", adhcBottom1 = "B"
Public Const adhcWidth = "Width", adhcWidth1 = "W"
Public Const adhcHeight = "Height", adhcHeight1 = "H"                       '!!! adhcHeight1 совпадает с adhcHorz2 - следить за контекстом
Public Const adhcHorz = "Horizontal", adhcHorz1 = "Horz", adhcHorz2 = "H"
Public Const adhcVert = "Vertical", adhcVert1 = "Vert", adhcVert2 = "V"
Public Const adhcFull = "Full", adhcFull1 = "F"
Public Const adhcBoth = "Both"
Public Const adhcYes = "Yes", adhcNo = "No"
Public Const adhcTrue = "True", adhcFalse = "False"
Public Const adhcOff = "Off"
Public Const adhcCenter = "Center", adhcCenter1 = "C"                       ' выравнивание по центру
Public Const adhcTo = "To"                                                  ' для указания в Place размещения внутри (к..)
Public Const adhcOn = "On"                                                  ' для указания в Place размещения снаружи (на..), может таже использоваться для флагов в значении True
Public Const adhcCascade = "Cascade"                                        ' для указания в Place направления композиции (только для окон)
Public Const adhcFrom = "From"                                              ' для композиции (только для окон)
Public Const adhcProp = "Prop"                                              ' для признака стиля пропорционального размера
Public Const adhcMin = "Min", adhcMax = "Max"
Public Const adhcFont = "Font"
Public Const adhcSize = "Size"
Public Const adhcScale = "Scale"
Public Const adhcMode = "Mode"
Public Const adhcAlign = "Align"
Public Const adhcPlace = "Place"
Public Const adhcAngle = "Angle"
Public Const adhcName = "Name"
Public Const adhcBack = "Back"
Public Const adhcBorder = "Border"
Public Const adhcFore = "Fore"
Public Const adhcAction = "Action"                                          ' действие
' единицы измерения
Public Const adhcCm = "cm", adhcCm1 = "см"                                  ' сантиметры
Public Const adhcInch = "in", adhcInch1 = "'"                               ' дюймы
Public Const adhcPoints = "pt"
Public Const adhcPixels = "px"
Public Const adhcTwips = "tw"
' ----------------------------------
Public Const adhcStyleIt = "StyleIt"            ' cтиль контрола. значения можно объединять.
Public Const c_strStyleDelims = "+&,"           ' допустимые разделители значений свойства StyleIt

Public Const adhcLeftTop = adhcLeft & adhcTop, adhcLeftTop1 = adhcLeft1 & adhcTop1                    ' 5  Left-Top
Public Const adhcRightTop = adhcRight & adhcTop, adhcRightTop1 = adhcRight1 & adhcTop1                ' 6  Right-Top
Public Const adhcLeftBottom = adhcLeft & adhcBottom, adhcLeftBottom1 = adhcLeft1 & adhcBottom1        ' 9  Left-Bottom
Public Const adhcRightBottom = adhcRight & adhcBottom, adhcRightBottom1 = adhcRight1 & adhcBottom1    ' 10 Right-Bottom
Public Const adhcHorTop = adhcHorz & adhcTop, adhcHor1Top = adhcHorz1 & adhcTop, adhcHor2Top1 = adhcHorz2 & adhcTop1 ' 7  Horz-Top
Public Const adhcHorBottom = adhcHorz & adhcBottom, adhcHor1Bottom = adhcHorz1 & adhcBottom, adhcHor2Bottom1 = adhcHorz2 & adhcBottom1 ' 11 Horz-Bottom
Public Const adhcVerLeft = adhcVert & adhcLeft, adhcVer1Left = adhcVert1 & adhcLeft, adhcVer2Left1 = adhcVert2 & adhcLeft1 ' 13 Left-Vert
Public Const adhcVerRight = adhcVert & adhcRight, adhcVer1Right = adhcVert1 & adhcRight, adhcVer2Right1 = adhcVert2 & adhcRight1  ' 14 Right-Vert

Public Const adhcCenterHor = adhcCenter & adhcHorz, adhcCenterHor1 = adhcCenter & adhcHorz1, adhcCenterHor2 = adhcCenter1 & adhcHorz2   ' горизонтальное выравнивание по центру
Public Const adhcCenterHorTop = adhcCenterHor & adhcTop, adhcCenterHor1Top = adhcCenterHor1 & adhcTop, adhcCenterHor2Top1 = adhcCenterHor2 & adhcTop1
Public Const adhcCenterHorBottom = adhcCenterHor & adhcBottom, adhcCenterHor1Bottom = adhcCenterHor1 & adhcBottom, adhcCenterHor2Bottom1 = adhcCenterHor2 & adhcBottom1
Public Const adhcCenterVer = adhcCenter & adhcVert, adhcCenterVer1 = adhcCenter & adhcVert1, adhcCenterVer2 = adhcCenter1 & adhcVert2   ' вертикальное выравнивание по центру
Public Const adhcCenterVerLeft = adhcCenterVer & adhcLeft, adhcCenterVer1Left = adhcCenterVer1 & adhcLeft, adhcCenterVer2Left1 = adhcCenterVer2 & adhcLeft1
Public Const adhcCenterVerRight = adhcCenterVer & adhcRight, adhcCenterVer1Right = adhcCenterVer1 & adhcRight, adhcCenterVer2Right1 = adhcCenterVer2 & adhcRight1

Public Const adhcXProp = adhcLeft2 & adhcProp, adhcLProp = adhcLeft1 & adhcProp   '16  позиция слева зависит от ширины области (распределение по ширине)
Public Const adhcRProp = adhcRight1 & adhcProp                                         '32  позиция справа зависит от ширины области (распределение по ширине)
Public Const adhcYProp = adhcTop2 & adhcProp, adhcTProp = adhcTop1 & adhcProp     '64  позиция сверху зависит от высоты области (распределение по высоте)
Public Const adhcBProp = adhcBottom1 & adhcProp                                        '128 позиция снизу зависит от высоты области (распределение по высоте)
Public Const adhcWProp = adhcWidth1 & adhcProp                                         '48  ширина пропорциональна ширине области
Public Const adhcHProp = adhcHeight1 & adhcProp                                        '192 высота пропорциональна высоте области

Public Const adhcIcon = "Icon"                          '
Public Const adhcPict = "Pict"                          '1024 выводить только картинку
Public Const adhcText = "Text"                          '2048 выводить только текст
Public Const adhcPictAndText = adhcPict & adhcText      '3072 выводить картинку и текст
' ----------------------------------
Public Const adhcSizeIt = "SizeIt"              ' изменить размер (фиксируем левую-правую/верхнюю-нижнюю границы)
Public Const adhcSizeRight = adhcRight          ' вправо по краю формы
Public Const adhcSizeBottom = adhcBottom        ' вниз по краю формы
Public Const adhcSizeBoth = adhcBoth            ' вправо и вниз по краю формы
Public Const adhcSizeNone = adhcNone
' ----------------------------------
Public Const adhcFloatIt = "FloatIt"            ' перемещать (фиксируем правую/нижнюю границу и размер)
Public Const adhcFloatRight = adhcRight         ' вправо по краю формы
Public Const adhcFloatBottom = adhcBottom       ' вниз по краю формы
Public Const adhcFloatBoth = adhcBoth           ' вправо и вниз по краю формы
Public Const adhcFloatNone = adhcNone
' ----------------------------------
Public Const adhcScaleIt = "ScaleIt"            ' масштабировать
' ----------------------------------
Public Const adhcSplitNone = adhcNone           ' не является разделителем
Public Const adhcSplitIt = "SplitIt"            ' разделитель (свободно перемещаемый) (изменяемая граница области)
Public Const adhcDelimIt = "Delimit"            ' разделитель (фиксированный) (границы области для привязки контролов)
Public Const adhcSplitVer = adhcVert, adhcSplitVer1 = adhcVert1, adhcSplitVer2 = adhcVert2 ' вертикальный разделитель
Public Const adhcSplitHor = adhcHorz, adhcSplitHor1 = adhcHorz1, adhcSplitHor2 = adhcHorz2 ' горизонтальный разделитель
'Public Const adhcSplitBoth = adhcBoth           ' оба (вертикальный и горизантальный) разделитель
'Public Const adhcDefault = adhcDefault          ' по-умолчанию - не является разделителем
' ----------------------------------
' для размеров/позиции контрола
' ----------------------------------
' adhcDef &
Public Const adhcDefLeft = adhcLeft, adhcDefLeft1 = adhcLeft1, adhcDefLeft2 = "X"   ' положение контрола по-умолчанию (м.б. заданы в абсолютных значениях относительно края, либо в долях относительно размеров формы/секции)
Public Const adhcDefTop = adhcTop, adhcDefTop1 = adhcTop1, adhcDefTop2 = "Y"
Public Const adhcDefRight = adhcRight, adhcDefRight1 = "R"
Public Const adhcDefBottom = adhcBottom, adhcDefBottom1 = "B"
Public Const adhcDefWidth = adhcWidth, adhcDefWidth1 = adhcWidth1                   ' размеры контрола по-умолчанию (м.б. заданы в абсолютных значениях, либо в долях относительно размеров формы/секции)
Public Const adhcDefHeight = adhcHeight, adhcDefHeight1 = adhcHeight1
'
Public Const adhcMinWidth = adhcMin & adhcWidth, adhcMinWidth1 = adhcMin & adhcWidth1, adhcMinWidth2 = adhcWidth1 & "0"     ' предельные размеры (м.б. заданы в абсолютных значениях, либо в долях относительно размеров контрола по-умолчанию)
Public Const adhcMaxWidth = adhcMax & adhcWidth, adhcMaxWidth1 = adhcMax & adhcWidth1, adhcMaxWidth2 = adhcWidth1 & "1"
Public Const adhcMinHeight = adhcMin & adhcHeight, adhcMinHeight1 = adhcMin & adhcHeight1, adhcMinHeight2 = adhcHeight1 & "0"
Public Const adhcMaxHeight = adhcMax & adhcHeight, adhcMaxHeight1 = adhcMax & adhcHeight1, adhcMaxHeight2 = adhcHeight1 & "1"
' ----------------------------------
Private Const adhcBond = "Bond"           ' привязка контрола (имя контрола относительно границ которого определены координаты)
Public Const adhcBondLeft = adhcBond & adhcDefLeft, adhcBondLeft1 = "BL", adhcBondLeft2 = "BX"
Public Const adhcBondTop = adhcBond & adhcDefTop, adhcBondTop1 = "BT", adhcBondTop2 = "BY"
Public Const adhcBondRight = adhcBond & adhcDefRight, adhcBondRight1 = "BR"
Public Const adhcBondBottom = adhcBond & adhcDefBottom, adhcBondBottom1 = "BB"
Public Const adhcBondWidth = adhcBond & adhcDefWidth, adhcBondWidth1 = "BW"
Public Const adhcBondHeight = adhcBond & adhcDefHeight, adhcBondHeight1 = "BH"
' ----------------------------------
' для масштабирования изображений на контролах
' ----------------------------------
Public Const adhcClip = "Clip" ' 0 - не меняем размер. если размер больше области вывода - обрезка
Public Const adhcStretch = "Stretch" ' 1 - сжатие/растяжение (нарушает пропорции)
Public Const adhcZoom = "Zoom" ' 3 - пропорциональное масштабирование
Public Const adhcDown = "Down", adhcDown1 = adhcZoom & adhcDown ' -1 - пропорциональное масштабирование, только уменьшаем
' ----------------------------------
' для задания относительного размещения изображения и текста
' ----------------------------------
' внутри по центру
Public Const adhcPlaceCenter = adhcCenter                       ' по центру (внутри)
Public Const adhcPlaceToLeft = adhcTo & adhcLeft                ' внутри слева по центру
Public Const adhcPlaceToRight = adhcTo & adhcRight              ' внутри справа по центру
Public Const adhcPlaceToTop = adhcTo & adhcTop                  ' внутри по центру сверху
Public Const adhcPlaceToBottom = adhcTo & adhcBottom            ' внутри по центру снизу
' снаружи по центру
Public Const adhcPlaceOnLeft = adhcOn & adhcLeft                ' снаружи слева по центру
Public Const adhcPlaceOnRight = adhcOn & adhcRight              ' снаружи справа по центру
Public Const adhcPlaceOnTop = adhcOn & adhcTop                  ' снаружи по центру сверху
Public Const adhcPlaceOnBottom = adhcOn & adhcBottom            ' снаружи по центру снизу
' внутри по краю
Public Const adhcPlaceToLeftTop = adhcTo & adhcLeftTop          ' внутри слева сверху
Public Const adhcPlaceToRightTop = adhcTo & adhcRightTop        ' внутри справа сверху
Public Const adhcPlaceToLeftBottom = adhcTo & adhcLeftBottom    ' внутри слева снизу
Public Const adhcPlaceToRightBottom = adhcTo & adhcRightBottom  ' внутри справа снизу
' снаружи по краю
Public Const adhcPlaceOnLeftToTop = adhcOn & adhcLeft & adhcTo & adhcTop          ' снаружи слева к верхнему краю
Public Const adhcPlaceOnLeftToBottom = adhcOn & adhcLeft & adhcTo & adhcBottom    ' снаружи слева к нижнему краю
Public Const adhcPlaceOnRightToTop = adhcOn & adhcRight & adhcTo & adhcTop        ' снаружи справа к верхнему краю
Public Const adhcPlaceOnRightToBottom = adhcOn & adhcRight & adhcTo & adhcBottom  ' снаружи справа к нижнему краю
Public Const adhcPlaceOnTopToLeft = adhcOn & adhcTop & adhcTo & adhcLeft          ' снаружи к левому краю сверху
Public Const adhcPlaceOnTopToRight = adhcOn & adhcTop & adhcTo & adhcRight        ' снаружи к правому краю сверху
Public Const adhcPlaceOnBottomToLeft = adhcOn & adhcBottom & adhcTo & adhcLeft    ' снаружи к левому краю снизу
Public Const adhcPlaceOnBottomToRight = adhcOn & adhcBottom & adhcTo & adhcRight  ' снаружи к правому краю снизу
'' композиция (только для формы ??)
Public Const adhcCascadeFromLeftTop = adhcCascade & adhcFrom & adhcLeftTop               ' размещение каскадом вправо-вниз
Public Const adhcCascadeFromRightTop = adhcCascade & adhcFrom & adhcRightTop             ' размещение каскадом влево-вниз
Public Const adhcCascadeFromLeftBottom = adhcCascade & adhcFrom & adhcLeftBottom         ' размещение каскадом вправо-вверх
Public Const adhcCascadeFromRightBottom = adhcCascade & adhcFrom & adhcRightBottom       ' размещение каскадом влево-вверх
' ----------------------------------
' для раскраски контролов
' ----------------------------------
Public Const adhcColor = "Color"
Public Const adhcBackColor = adhcBack & adhcColor
Public Const adhcForeColor = adhcFore & adhcColor
Public Const adhcFontColor = adhcFont & adhcColor
Public Const adhcTextColor = adhcText & adhcColor
Public Const adhcBorderColor = adhcBorder & adhcColor
Public Const adhcColorBlack = "Black"
Public Const adhcColorWhite = "White"
Public Const adhcColorGray = "Gray"
Public Const adhcColorDark = "Dark"        ' =appColorDark
Public Const adhcColorDark2 = adhcColorDark & 2
Public Const adhcColorDark3 = adhcColorDark & 3
Public Const adhcColorBright = "Bright"    ' =appColorBright
Public Const adhcColorBright2 = adhcColorBright & 2
Public Const adhcColorBright3 = adhcColorBright & 3
Public Const adhcColorLight = "Light"      ' =appColorLight
Public Const adhcColorLight2 = adhcColorLight & 2
Public Const adhcColorLight3 = adhcColorLight & 3
' ----------------------------------
' для присвоения изображений
' ----------------------------------
Public Const adhcIconName = adhcIcon 'PictName
Public Const adhcPictName = adhcPict 'PictName
Public Const adhcPictSize = adhcPict & adhcSize 'PictSize
Public Const adhcPictMode = adhcPict & adhcMode, adhcPictMode1 = adhcPictSize & adhcMode 'PictMode
Public Const adhcPictAlign = adhcPict & adhcPlace 'PictAlign
Public Const adhcPictAngle = adhcPict & adhcAngle 'PictAngle
Public Const adhcPictText = adhcText 'TextString
Public Const adhcTextAlign = adhcText & adhcAlign
Public Const adhcTextPlace = adhcText & adhcPlace 'TextPlace
Public Const adhcTextAngle = adhcText & adhcAngle 'TextAngle

Public Const adhcPictGray = adhcColorGray, adhcPictGray1 = adhcPictGray & adhcScale 'GrayScale
'
Public Const adhcFontName = adhcFont & adhcName
Public Const adhcFontSize = adhcFont & adhcSize
' ----------------------------------
Public Const COLORNOTSET = &HFFFFFFFF

' ==================================
Public Function AccWinHide()
On Error Resume Next
    DoCmd.SelectObject acTable, , True
    DoCmd.RunCommand acCmdWindowHide
    Err.Clear
End Function
Public Function AccWinShow()
On Error Resume Next
    DoCmd.SelectObject acTable, , True
    'DoCmd.RunCommand acCmdWindowUnhide
    Err.Clear
End Function

Public Function IsActiveControl(ctl As Control) As Boolean
Const PTR_SIZE = 4
Dim o As Object
Dim ptrAct As Long 'active control ptr
Dim ptrCtl As Long 'control ptr
Dim Ret As Boolean
    On Error Resume Next
    Set o = Screen.ActiveControl
    If Not (o Is Nothing) Then
        CopyMemory ptrAct, o, PTR_SIZE
        CopyMemory ptrCtl, ctl, PTR_SIZE
        If ptrAct = ptrCtl Then Ret = True
    End If
    'IsActiveControl = Screen.ActiveControl Is ctl
    IsActiveControl = Ret
End Function

Public Function IsSubForm(frm As Form) As Boolean
' Проверяет открыта ли форма как субформа
    On Error Resume Next
Dim strName As String: strName = frm.PARENT.Name
    IsSubForm = (Err.Number = 0): Err.Clear
End Function

Public Function IsSubformFocus() As Boolean
' Проверяет фокус в субформе
    On Error Resume Next
Dim ctl As Object: Set ctl = Screen.ActiveControl.PARENT
    If Not TypeOf ctl Is Access.Form Then Set ctl = ctl.PARENT
    Set ctl = ctl.PARENT
    IsSubformFocus = Not CBool(Err.Number)
    Err.Clear
End Function

Public Function IsFormExists(FormName As String) As Boolean
' проверяет существует ли форма
Dim Result As Boolean:  Result = False
    On Error GoTo HandleError
    Result = (CurrentProject.AllForms(FormName).Name = FormName) '
HandleExit:  IsFormExists = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function IsFormLoaded(FormName As String) As Boolean
' проверяет загружена ли форма
Dim Result As Boolean:  Result = False
    On Error GoTo HandleError
    If SysCmd(acSysCmdGetObjectState, acDefault, FormName) = acObjStateClosed Then Err.Raise vbObjectError + 512
    Result = (Application.Forms(FormName).CurrentView <> 0)  ' <> Design
HandleExit:  IsFormLoaded = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function IsFormOpened(FormName As String, _
    Optional View As AcFormView, Optional state As acFormState) As Boolean
' проверяет открыта ли форма
' State =   0 - Closed
'           1 - (acObjStateOpen)    Open
'           2 - (acObjStateDirty)   Changed but not saved
'           4 - (acObjStateNew)     New
' View  =   0 - Normal View
'           1 - Design View
'           3 - Datasheet View
Dim Result As Boolean:  Result = False
    On Error GoTo HandleError
    state = SysCmd(acSysCmdGetObjectState, acDefault, FormName): If state = acObjStateClosed Then Err.Raise vbObjectError + 512
    Select Case Application.Forms(FormName).CurrentView
    Case 0: View = acDesign
    Case 1: View = acNormal
    Case 2: View = acFormDS
    Case Else: View = -1
    End Select
HandleExit:  IsFormOpened = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function IsReportExists(FormName As String) As Boolean
' проверяет существует ли отчёт
Dim Result As Boolean:  Result = False
    On Error GoTo HandleError
    Result = (CurrentProject.AllReports(FormName).Name = FormName) '
HandleExit:  IsReportExists = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function IsReportLoaded(ReportName As String) As Boolean
' проверяет загружен ли отчёт
Dim Result As Boolean:  Result = False
    On Error GoTo HandleError
    If SysCmd(acSysCmdGetObjectState, acDefault, ReportName) = acObjStateClosed Then Err.Raise vbObjectError + 512
    Result = (Application.Reports(ReportName).CurrentView <> 0) ' <> Design
HandleExit:  IsReportLoaded = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function IsReportOpened(ReportName As String, _
    Optional View As AcFormView, Optional state As acFormState) As Boolean
' проверяет открыт ли отчёт
' State =   0 - Closed
'           1 - Open
' View  =   0 - Normal View
'           1 - Design View
Dim Result As Boolean:  Result = False
    On Error GoTo HandleError
    state = SysCmd(acSysCmdGetObjectState, acDefault, ReportName): If state = acObjStateClosed Then Err.Raise vbObjectError + 512
    Select Case Application.Reports(ReportName).CurrentView
    Case 0: View = acDesign
    Case 1: View = acNormal
    Case Else: View = -1
    End Select
HandleExit:  IsReportOpened = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function FormOpen( _
    FormName As String, _
    Optional View As AcFormView = acNormal, _
    Optional WhereCondition, _
    Optional DataMode As AcFormOpenDataMode = acFormPropertySettings, _
    Optional WindowMode As AcWindowMode = acWindowNormal, _
    Optional OpenArgs, _
    Optional PARENT, _
    Optional x, Optional y, _
    Optional Placement As ePlace = eCascadeFromLeftTop, _
    Optional Icon, _
    Optional Visible As Boolean = True, _
    Optional NewForm As Access.Form _
    ) As Boolean
'    Optional FilterName,
' открывает форму с указаным именем
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
'----------------------------------
' обрабатываем режим открытия
'----------------------------------
'        Select Case View
'        Case acNormal: .DefaultView = 0            ' нормальный
'        Case acFormDS: .DefaultView = 2            ' режим таблицы
'        Case acFormPivotTable: .DefaultView = 3    ' режим сводной таблицы
'        Case acFormPivotChart: .DefaultView = 4    ' режим сводной диаграммы
'        Case acPreview ' пока не поддерживается    ' режим просмотра печати
'        Case acDesign  ' пока не поддерживается    ' режим просмотра печати
'        End Select
    ' проверяем текущее состояние формы
Dim acState As acFormState: acState = SysCmd(acSysCmdGetObjectState, acForm, FormName)
    If (acState <> acObjStateClosed) Then
    ' режим конструктора можно задать или изменить только при открытии
        Set NewForm = Application.Forms(FormName)
        If (View = acDesign) Xor (NewForm.CurrentView = 0) Then
    ' форма открыта, но заданный режим, либо текущий - режим конструктора
        ' режим конструктора можно задать или изменить только при открытии - закрываем
            DoCmd.Close acForm, FormName, Save:=acSaveYes: acState = acObjStateClosed
        End If
    End If
    If acState = acObjStateClosed Then
' открываем форму по имени в нужном режиме
Dim tmpWinMode As AcWindowMode: tmpWinMode = acHidden '  WindowMode
        DoCmd.OpenForm FormName, View, _
                WhereCondition:=WhereCondition, DataMode:=DataMode, _
                WindowMode:=tmpWinMode, OpenArgs:=OpenArgs ',FilterName:=FilterName
    End If
'----------------------------------
' возвращаем ссылку на объект
'----------------------------------
Dim i As Long
    For i = Application.Forms.Count - 1 To 0 Step -1
        Result = (Application.Forms(i).Name = FormName):   If Result Then Set NewForm = Application.Forms(i):   Exit For
    Next i
    If Not Result Then: Err.Raise vbObjectError + 512
'----------------------------------
' задаём свойства объекта
'----------------------------------
    Result = AccObjectSet(NewForm, _
        View:=View, _
        WhereCondition:=WhereCondition, _
        DataMode:=DataMode, _
        WindowMode:=WindowMode, _
        OpenArgs:=OpenArgs, _
        x:=x, y:=y, _
        ObjectParent:=PARENT, _
        Icon:=Icon, _
        Visible:=Visible, _
        Placement:=Placement)
    NewForm.SetFocus
    'WindowUnFreeze  ' разблокируем обновление
HandleExit:  DoCmd.Echo True: FormOpen = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function FormOpenDrop( _
    FormName As String, _
    Optional FormVal, _
    Optional PARENT, _
    Optional x, Optional y, _
    Optional Placement As ePlace = ePlaceOnBottomToLeft, _
    Optional Icon, _
    Optional Visible As Boolean = True, _
    Optional ByRef NewForm As Object _
    )
Const c_strProcedure = "FormOpenDrop"
' открывает выпадающую форму типа FormMode
' в NewForm возвращает ссылку на открытую форму ()
Dim tmpPoint As POINT ' координаты опорной точки
Dim Result

    Result = False
    On Error GoTo HandleError
    If IsMissing(PARENT) Then
        GetCursorPos tmpPoint: x = tmpPoint.x: y = tmpPoint.y ' по умолчанию выводим в координаты курсора
    ElseIf PARENT Is Nothing Then
        GetCursorPos tmpPoint: x = tmpPoint.x: y = tmpPoint.y ' по умолчанию выводим в координаты курсора
    End If
    Result = FormOpen(FormName, _
        WindowMode:=acDialog, x:=x, y:=y, Icon:=Icon, Placement:=Placement, _
        PARENT:=PARENT, NewForm:=NewForm, _
        Visible:=False): If Result = 0 Then Err.Raise vbObjectError + 512
    On Error Resume Next ':GoTo HandleError
    With NewForm
Dim bolModal As Boolean: bolModal = (.ModalResult = .ModalResult)
        If IsMissing(FormVal) Then Result = PARENT.Value Else Result = FormVal
        .Value = Result
        Err.Clear
'        On Error GoTo HandleError
' выводим на экран
        .Visible = Visible: If Visible Then .SetFocus
' гоняем в цикле пока не дождёмся выбора пользователя
    If bolModal Then
        Do While .Visible: DoEvents: Loop: If .ModalResult = vbOK Then Result = .Value
        DoCmd.Close acForm, NewForm.Name, acSaveNo: Set NewForm = Nothing ' если открывали модальную - после получения ответа закрываем
    End If
    End With
' возвращаем результат в исходное поле
    If TypeOf PARENT Is Access.Control Then PARENT.Value = Result
    Err.Clear
HandleExit:  FormOpenDrop = Result: Exit Function
HandleError: Result = False:  Err.Clear: Resume HandleExit
End Function
#If UseContextMenu Then
Public Function FormOpenContext( _
    ContextData As String, _
    Optional ByRef ContextMenu As Object, _
    Optional ByRef ContextVal, _
    Optional ByRef ContextName As String, _
    Optional PARENT, _
    Optional x, Optional y, _
    Optional Arrange As eAlign = eAlignLeftTop, _
    Optional Visible As Boolean = True)
' создает и открывает контекстное меню
Const c_strProcedure = "FormOpenContext"
' ContextData - перечень элементов меню или имя запроса источника элементов
' ContextMenu - ссылка на открываемое (созданое) меню
' ContextVal - значение контекстного меню (отображаемое по-умолчанию или возвращаемое)
' ContextName - имя создаваемого меню (по-умолчанию - "~tmpContextMenu")
' Parent  - ссылка на родительский объект
' X, Y - координаты вывода меню
' Visible - определяет создавать контекстное меню видимым или не видимым
' Arrange - тип выравнивания меню отн-но координат (по-умолчанию координаты задают верхний левый угол меню)
Const cstrContextName = "~tmpContextMenu"
Dim mnu As clsContextMenu
Dim strWhere As String
Dim Ret As Long
    ContextName = Trim$(ContextName)
    'If Len(ContextName) = 0 Then ContextName = cstrContextName
' создаем контекстное меню
    Set mnu = New clsContextMenu 'Set ContextMenu = Application.CommandBars.Add(Name:=ContextName, Position:=msoBarPopup)
    With mnu ' ContextMenu
        .CreateContextMenu ContextName
    ' проверяем ContextData
'Stop
On Error Resume Next
    ' код контекстного меню по таблице SysMenu
        If IsNumeric(ContextData) Then Ret = .CreateItemsFromSQL(c_strTableMenu, WhereCond:=c_strParent & sqlEqual & ContextData): GoTo HandleShow
Dim dbs As DAO.Database: Set dbs = CurrentDb
Dim rst As DAO.Recordset
Dim strSQL As String
    ' кодовое имя контекстного меню по таблице SysMenu
        strSQL = sqlSelectAll & c_strTableMenu & sqlWhere & c_strCName & sqlEqual & "'" & ContextData & "'"
        Set rst = dbs.OpenRecordset(strSQL): If Err Then Err.Clear Else Ret = .CreateItemsFromSQL(c_strTableMenu, WhereCond:=c_strParent & sqlEqual & rst.Fields(c_strKey)): GoTo HandleShow
    ' имя таблицы/запроса/связаной таблицы
        Set rst = dbs.OpenRecordset(ContextData): If Err Then Err.Clear Else Ret = .CreateItemsFromSQL(ContextData): GoTo HandleShow
    ' список элементов
        Ret = .CreateItemsFromString(ContextData)
HandleShow:
' выводим и ждём выбора
        If Ret Then
            .ShowMenu x, y ', Arrange ' перемещаем окно
            ContextVal = .Value
        Else
            .RemoveContextMenu ContextName: Set mnu = Nothing
        End If
    End With
    'Result = ContextVal
    Set ContextMenu = mnu
    'DoCmd.Echo True' Включаем отображение на экране
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
#End If
Public Function ReportOpen( _
    ReportName As String, _
    Optional View As AcFormView = acViewPreview, _
    Optional WhereCondition, _
    Optional WindowMode As AcWindowMode = acWindowNormal, _
    Optional OpenArgs, _
    Optional PARENT, _
    Optional Placement As ePlace = eCascadeFromLeftTop, _
    Optional NewReport As Access.Report _
    ) As Boolean
' открывает отчёт с заданными параметрами
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
'----------------------------------
' обрабатываем режим открытия
'----------------------------------
'        Select Case View
'        Case acNormal: .DefaultView = 0            ' нормальный
'        Case acFormDS: .DefaultView = 2            ' режим таблицы
'        Case acFormPivotTable: .DefaultView = 3    ' режим сводной таблицы
'        Case acFormPivotChart: .DefaultView = 4    ' режим сводной диаграммы
'        Case acPreview ' пока не поддерживается    ' режим просмотра печати
'        Case acDesign  ' пока не поддерживается    ' режим просмотра печати
'        End Select
    ' проверяем текущее состояние формы
Dim acState As acFormState: acState = SysCmd(acSysCmdGetObjectState, acReport, ReportName)
    If (acState <> acObjStateClosed) Then
    ' режим конструктора можно задать или изменить только при открытии
        Set NewReport = Application.Reports(ReportName)
        If (View = acDesign) Xor (NewReport.CurrentView = 0) Then
    ' отчёт открыт, но заданный режим, либо текущий - режим конструктора
        ' режим конструктора можно задать или изменить только при открытии - закрываем
            DoCmd.Close acReport, ReportName, Save:=acSaveYes: acState = acObjStateClosed
        End If
    End If
    If acState = acObjStateClosed Then
' открываем отчёт по имени в нужном режиме
        DoCmd.OpenReport ReportName, View, _
            WhereCondition:=WhereCondition, _
            WindowMode:=WindowMode, OpenArgs:=OpenArgs ',FilterName:=FilterName
    End If
'----------------------------------
' возвращаем ссылку на объект
'----------------------------------
Dim i As Long
    For i = Application.Reports.Count - 1 To 0 Step -1
        Result = (Application.Reports(i).Name = ReportName): If Result Then Set NewReport = Application.Reports(i): Exit For
    Next i
'----------------------------------
' задаём свойства объекта
'----------------------------------
    If Not Result Then: Err.Raise vbObjectError + 512
    Result = AccObjectSet(NewReport, _
        View:=View, _
        WhereCondition:=WhereCondition, _
        WindowMode:=WindowMode, _
        OpenArgs:=OpenArgs, _
        ObjectParent:=PARENT, _
        Placement:=Placement)
    'WindowUnFreeze  ' разблокируем обновление
HandleExit:  DoCmd.Echo True: ReportOpen = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function AccObjectSet(accObject As Object, _
    Optional View As AcFormView = acNormal, _
    Optional WhereCondition, _
    Optional DataMode As AcFormOpenDataMode = acFormPropertySettings, _
    Optional WindowMode As AcWindowMode = acWindowNormal, _
    Optional OpenArgs, _
    Optional ObjectParent, _
    Optional x, Optional y, _
    Optional Placement As ePlace = eCascadeFromLeftTop, _
    Optional Icon, _
    Optional Visible As Boolean = True, _
    Optional ModalResult As VbMsgBoxResult, _
    Optional FormResult) As Boolean
' задаёт свойства объекта Access (форму или отчёт)
'---------------------
'Возвращает ссылку на открытый объект
'accObject - ссылка на объект Access форму или отчёт
'View - (необязательный) режим просмотра формы
'FilterName -
'WhereCondition - (необязательный) строковое выражение соответствующее SQL WHERE без WHERE.
'DataMode - (необязательный) режим доступа к данным формы. Перезаписывает значения свойств AllowEdits, AllowDeletions, AllowAdditions и DataEntry
'WindowMode - (необязательный) режим открытия окна
'OpenArgs - (необязательный) строка параметров передаваемых открываемой форме
'ObjectParent - (необязательный) ссылка на родительскую форму для данной формы
'ModalResult - (необязательный) возвращает значение (VbMsgBoxResult) нажатой кнопки в форме открытой как окно диалога
'ObjectResult - (необязательный) результат возвращаемый формой после завершения работы
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    
    If TypeOf accObject Is Access.Form Then
    ElseIf TypeOf accObject Is Access.Report Then
    Else: Err.Raise vbObjectError + 512
    End If
    
'    'WindowFreeze Application.hWndAccessApp
    DoCmd.Echo False
    On Error Resume Next
    With accObject '
    ' здесь надо все сильно переделать, а пока так:
    ' выводим на экран, передаем параметры и делаем активным окно формы
'----------------------------------
' наполняем и настраиваем форму
'----------------------------------
        '.Parametres = OpenArgs  ' заполняем форму (передаем строку параметров)
        If Not IsMissing(WhereCondition) Then If Len(WhereCondition) > 0 Then .Filter = WhereCondition: .FilterOn = True
        Err.Clear
'----------------------------------
' обрабатываем режим вывода данных
'----------------------------------
        Select Case DataMode
        Case acFormPropertySettings    ' определяется свойствами формы
        Case acFormEdit        ' режим редактирования данных
        '.DataEntry = False
            .RecordsetType = 0  ' Dynaset
            .AllowEdits = True: .AllowAdditions = True: .AllowDeletions = True
        Case acFormAdd         ' режим ввода данных
            .DataEntry = True   ' чистая запись
            .RecordsetType = 1  ' Dynaset
            .AllowEdits = False: .AllowAdditions = True: .AllowDeletions = False
        Case acFormReadOnly    ' режим просмотра данных
        '.DataEntry = False
            .RecordsetType = 2  ' Snapshot
            .AllowEdits = False: .AllowAdditions = False: .AllowDeletions = False
        End Select
'----------------------------------
' обрабатываем режим открытия окна
'----------------------------------
        Select Case WindowMode
        Case acWindowNormal:    WindowSet .hwnd, SW_SHOWNORMAL: .Visible = True ' обычное окно          '.SetFocus
        Case acIcon:            WindowSet .hwnd, SW_SHOWMINIMIZED               ' свернутое в значок    '.Visible = True: .SetFocus
        Case acDialog:          .Visible = False: .PopUp = True: .Modal = True  ' диалоговое окно       '.Visible = Visible ': .SetFocus:WindowSet .hWnd, SW_SHOWNORMAL'.BorderStyle = 3   ' только в конструкторе
        Case acHidden:          .Visible = False                                ' скрытое               'WindowSet .hWnd, SW_HIDE
        End Select
'----------------------------------
' обрабатываем диалоговое окно
'----------------------------------
        If WindowMode = acDialog Then   ' диалоговое окно
        '
        End If
HandlePlaceForm:
'----------------------------------
' позиционируем окно относительно заданного родителя
'----------------------------------
        On Error GoTo HandleObjectDesign
' получаем точки привязки и параметры размещения созданной формы относительно родителя
Dim Cascade As Boolean: Cascade = ((Placement And eCascade) = eCascade)
Dim rXpar As Single, rYpar As Single: Call GetAlignPoint(Placement Mod &H10, rXpar, rYpar)    ' на родителе к клиенту
Dim rXсli As Single, rYсli As Single: Call GetAlignPoint(Placement \ &H10, rXсli, rYсli)      ' на клиенте к родителю
Dim Xcli As Long, Ycli As Long, dX As Long, dY As Long
' получаем координаты/размеры созданной формы
Dim cliRect As RECT: GetWindowRect accObject.hwnd, cliRect
' получаем координаты/размеры родительского объекта
Dim parRect As RECT:
        If IsMissing(ObjectParent) Then
    ' родитель не задан - размещаем относительно рабочей области Access
            GetWindowRect FindWindowEx(hWndAccessApp, 0&, accClassChild, vbNullString), parRect ' MDIClient
        ElseIf ObjectParent Is Nothing Then
            GetWindowRect FindWindowEx(hWndAccessApp, 0&, accClassChild, vbNullString), parRect ' MDIClient
        ElseIf TypeOf ObjectParent Is Access.Form Or TypeOf ObjectParent Is Access.Report Then
    ' родитель форма или отчёт - размещаем относительно позиции окна заданной формы/отчёта
            GetWindowRect ObjectParent.hwnd, parRect                                            ' Access Form/Report
        ElseIf TypeOf ObjectParent Is Access.Control Then
    ' родитель контрол - размещаем относительно позиции заданного контрола
            With parRect: AccControlLocation ObjectParent, .Left, .Top, .Right, .Bottom: .Right = .Left + .Right: .Bottom = .Top + .Bottom: End With
        Else: Err.Raise vbObjectError + 512
        End If
    ' добавляем смещение для каскадного размещения окон
        If Cascade Then dX = pxWinGap: dY = pxWinGap
    ' x, y рассматриваем как смещение от точки привязки
        If Not IsMissing(x) Then dX = dX + x
        If Not IsMissing(y) Then dY = dY + y
    ' в зависимости от Placement размещаем форму относительно родителя
        Xcli = parRect.Left + rXpar * (parRect.Right - parRect.Left) - rXсli * (cliRect.Right - cliRect.Left) + dX
        Ycli = parRect.Top + rYpar * (parRect.Bottom - parRect.Top) - rYсli * (cliRect.Bottom - cliRect.Top) + dY
' Надо отловить почему ошибочно позиционирует окно за пределами области видимости
        FormMove accObject, Xcli, Ycli ' позиционируем
'----------------------------------
' задаём оформление формы
'----------------------------------
HandleObjectDesign:
' ...
    On Error Resume Next
    Call TagsRead(accObject, Init:=True)
#If UseModPictureData Then
    If Not IsMissing(Icon) Then Call PictureData_SetIcon(accObject.hwnd, Icon)
#End If
'----------------------------------
    End With
    Result = True
    'WindowUnFreeze 'AccObject.hwnd ' разблокируем обновление
HandleExit:  DoCmd.Echo True: AccObjectSet = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function AccControlSet( _
    ctl As Access.Control, _
    Optional x, Optional y, _
    Optional w, Optional h, _
    Optional r, Optional b, _
    Optional Show, _
    Optional Init As Boolean = False, _
    Optional Col, _
    Optional frm As Access.Form) As Boolean
' выводит контрол в заданную позицию, и настраивает его внешний вид
Const c_strProcedure = "AccControlSet"
' ctl       - настраиваемый контрол
' x/y/r/b/w/h - позиция/размеры для вывода контрола контрола
' Relative  - признак относительных позиции/размеров контрола
'             имеет смысл только при Show=False - когда необходимо сохранить текущую позицию/размеры в Тэг перед обнулением
' Show      - признак видимости контрола
' Init      - признак инициализации контрола
' col       - ссылка на коллекцию имен контролов содержащих читаемые тэги
' frm       - ссылка на родительскую форму
'-------------------------
' v.1.0.0       : 17.07.2023 - исходная версия
'-------------------------
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
Const cShift = px \ 3       ' величина смещения границы если контрол не влазит по высоте
Dim bTry As Byte: bTry = 0  ' счетчик попыток вместить контрол
    If frm Is Nothing Then Set frm = GetTopParent(ctl, True)  ' ???
Dim sec As Access.Section: Set sec = frm.Section(ctl.Section)
'Dim bInit As Boolean: bInit = bInit And Show
    
' читаем и разбираем доп свойства контрола сохранённые в тэге
Dim xCtl As Long, yCtl As Long, wCtl As Long, hCtl As Long
    If Init Then If Len(ctl.Tag) > 0 Then Call TagsRead(ctl)  'Control.Tag -> Control
    On Error GoTo HandleError
' задаём позицию и размеры контрола
    With ctl
        If Show Then
        ' получаем требуемые размеры контрола
        ' если  >1      - берём как абсолютное занчение в tw,
        ' если [0..1]   - берём относительно размеров формы/секции контрола
            If IsMissing(x) Then xCtl = .Left Else If x > twMinLim Then xCtl = x Else xCtl = frm.Width * x
            If IsMissing(w) Then wCtl = .Width Else If w > twMinLim Then wCtl = w Else wCtl = frm.Width * w
            'If IsMissing(r) Then wCtl = .Width Else If w > twMinLim Then wCtl = w Else wCtl = frm.Width * w
            If IsMissing(y) Then yCtl = .Top Else If y > twMinLim Then yCtl = y Else yCtl = sec.Height * y
            If IsMissing(h) Then hCtl = .Height Else If h > twMinLim Then hCtl = h Else hCtl = sec.Height * h
            'If IsMissing(b) Then hCtl = .Height Else If h > twMinLim Then hCtl = h Else hCtl = sec.Height * h
        ' если контрол не влазит изменяем размеры формы/секции
            If xCtl + wCtl > frm.Width Then frm.Width = xCtl + wCtl
            If yCtl + hCtl > sec.Height Then sec.Height = yCtl + hCtl
        Else
' !!! косяк - если скроем таким способом контрол сплиттер потом при его восстановлении все посыпется
'  наверное надо запоминать при скрытии последнии данные в Тэг, а потом восстанавливать
            xCtl = 0: yCtl = 0: wCtl = 0: hCtl = 0
        End If
        ' задаём видимость контрола
        If Not IsMissing(Show) Then .Visible = Show
        ' задаём размеры контрола
        .Width = wCtl: .Height = hCtl: .Left = xCtl:
HandleNextTry: .Top = yCtl: On Error Resume Next
    End With
    Result = True
HandleExit:  AccControlSet = Result: Exit Function
HandleError: ' если элемент не может быть размещён в указанном месте - пробуем увеличить высоту
    If Err.Number = 2100 And bTry < 3 Then Err.Clear: sec.Height = sec.Height + cShift: bTry = bTry + 1: Resume HandleNextTry
    Result = False: Err.Clear: Resume HandleExit
End Function

Public Function FormMove( _
    AccForm, _
    ByRef x As Long, ByRef y As Long, _
    Optional Width, Optional Height, _
    Optional Arrange As eAlign = eAlignLeftTop, _
    Optional Inscribe As Boolean = True) As Long
' устанавливает акцесcовскую форму в экранные координаты
' с поправкой на координаты клиентской области окна Access
Const c_strProcedure = "FormMove"
Dim accHwnd As LongPtr ', frmHwnd As LongPtr
Dim accPoint As POINT 'accX As Long, accY As Long
Dim frmRect As RECT, accRect As RECT

    On Error GoTo HandleError
' получаем координаты области окна
    With AccForm
        GetWindowRect .hwnd, frmRect
    ' проверяем нужна ли поправка на координаты рабочей области Access
        ' всплывающей форме поправка не нужна, - она выводится поверх Access
        If Not .PopUp Then
'Stop    ' может быть всплывающей но не иметь аттрибута PopUp - как? - надо проверить
        ' делаем поправку на координаты рабочей области Access
            ' ищем клиентскую область окна Access
            accHwnd = FindWindowEx(hWndAccessApp, 0&, accClassChild, vbNullString) ' MDIClient
            GetWindowRect accHwnd, accRect
            ' получаем экранные координаты верхнего левого угла клиентской части области рабочей области Access
            accPoint.x = 0: accPoint.y = 0
            ClientToScreen accHwnd, accPoint
            ' координаты внутри Access для точного позиционирования на экране
            x = x - accPoint.x: y = y - accPoint.y
'        Else
        End If
    ' если не указаны размеры формы используем текущие
        If IsMissing(Height) Or IsMissing(Width) Then
            GetWindowRect .hwnd, frmRect
            With frmRect
                Width = .Right - .Left
                Height = .Bottom - .Top
            End With
        End If
    ' относительно x (разместить горизонтально)
        If (Arrange And eCenterHorz) = eCenterHorz Then
            x = x + Width \ 2 ' центрировать
         ElseIf (Arrange And eLeft) = eLeft Then
            'x = x           ' выровнять по левому (разместить справа от точки)
         ElseIf (Arrange And eRight) = eRight Then
            x = x - Width     ' выровнять по правому (разместить слева от точки)
        End If
    ' относительно y (разместить вертикально)
        If (Arrange And eCenterVert) = eCenterVert Then
            y = y + Height \ 2 ' центрировать
         ElseIf (Arrange And eTop) = eTop Then
            'y = y           ' выровнять по верхнему (разместить ниже точки)
         ElseIf (Arrange And eBottom) = eBottom Then
            y = y - Height     ' выровнять по нижнему (разместить выше точки)
        End If
        If Inscribe Then
        ' если указано вписать в рабочую область
        ' пытаемся изменить Arrange чтобы форма вся помещалась в экран
        End If
    ' позиционируем форму
        MoveWindow .hwnd, x, y, Width, Height, 1
    End With

HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function PosAccForm2Client( _
    AccForm As Access.Form, _
    ByRef x As Long, ByRef y As Long, _
    Optional Section As AcSection = acDetail)
' переводит координаты формы Access в координаты клиетской области (в пикселях)
' ранее AccFormCoords2Client
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
Const lSelectorWidth = 18
Dim v As Long, h As Long
Dim lBorder As Long
Dim accPoint As POINT
Dim hwnd As Long
' переводим координаты контрола в форме в твипах в пиксели
    With AccForm
        hwnd = .hwnd
        Select Case .DefaultView
        Case 1, 2 'ленточная или табличная
            x = x + .CurrentSectionLeft ', DIRECTION_HORIZONTAL)
            y = y + .CurrentSectionTop ', DIRECTION_VERTICAL)
        Case Else
        End Select
    End With
    accPoint.x = TwipsToPixels(x, DIRECTION_HORIZONTAL) '(X) / lTwipsPerInch * lLogPixelPerInchX
    accPoint.y = TwipsToPixels(y, DIRECTION_VERTICAL)   '(Y) / lTwipsPerInch * lLogPixelPerInchY

    'учтем скроллбар
    ScrollbarGetPos AccForm.hwnd, v, h
    accPoint.x = accPoint.x - h
    If Section = acDetail Then accPoint.y = accPoint.y - v  'вертикальный скроллбар влияет только на Detail
    
    x = accPoint.x: y = accPoint.y
    Result = True
HandleExit:  PosAccForm2Client = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function AccControlLocation(accObject As Variant, _
    Optional ByRef x, Optional ByRef y, Optional ByRef w, Optional ByRef h, _
    Optional ParentObject, _
    Optional ClientAreaPos As Boolean) As Boolean
' возвращает размер и позицию (экранные/клиентские) объекта Access (в пикселях)
'----------------
' ctl - ссылка на проверяемый объект Access
' x/y - возвращаемые координаты контрола (в пикселях) в зависимости от RetType экранные/клиентские
' w/h - возвращаемые ширина/высота контрола (в пикселях)
' ParentObject  - ссылка на старшую родительскую форму/отчет которым принадлежит контрол
' ClientAreaPos - определяет тип возвращаемых координат контрола
'       False - экранные (по-умолчанию)
'       True  - координаты в клиентской области формы
'----------------
' v.0.2.2       : 08.08.2021 - добавлена возможность возвращать экранные/клиентски координаты
' v.0.2.1       : 15.04.2019 - исправлено позиционирование в секциях формы
'----------------
' в большинстве случаев отлично cработает ctl.accLocation X, Y, W, H, varChild
' НО - в ленточной форме (покрайней мере в Access 2003) при расчете позиции по вертикали
' похоже он к каждой строке добавляет высоту секции заголовка
' из-за чего смещение от строки к строке возрастает
' поэтому считаем по-другому:
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
Dim lpPoint As POINT
'Dim frm As Access.Form
Dim cX As Long, cY As Long, cW As Long, ch As Long
Dim dX As Long, dY As Long, dW As Long, dH As Long
' определяем родительскую форму и координаты/размеры переданного контрола
    If TypeOf accObject Is Access.Control Then
        Set ParentObject = GetTopParent(accObject)
        With accObject: cX = .Left: cY = .Top: cW = .Width: ch = .Height: End With
        With ParentObject
        dX = .CurrentSectionLeft
        If accObject.Section <> acHeader Then
    ' если это не заголовок формы
        ' для обычной формы добавляем высоты вышестоящих секций
        ' для ленточной формы добавляем расстояние от верхнего края формы
            On Error Resume Next: dY = .Section(acHeader).Height: Err.Clear: On Error GoTo HandleError
            Select Case accObject.Section
            Case acDetail: If (.DefaultView = 1 Or .DefaultView = 2) Then dY = .CurrentSectionTop
            Case acFooter: dY = dY + .Section(acDetail).Height
            End Select
        End If
        End With
    ElseIf TypeOf accObject Is Access.Form Or TypeOf accObject Is Access.Report Then
        Set ParentObject = accObject
        cX = 0: cY = 0
        '' размеры окна формы
        'cW = frm.InsideWidth: cH = frm.InsideHeight: End With
        '' размеры формы/ если есть заголовок и примечание добавляем их высоту к высоте секции Detail
        With ParentObject
            On Error Resume Next
            dH = .Section(acHeader).Height + .Section(acFooter).Height: Err.Clear
            On Error GoTo HandleError
            cW = .Width: ch = dH + .Section(acDetail).Height
        End With
    Else: Err.Raise vbObjectError + 512
    End If
' если надо - получаем точку отсчета (верхний-левый угол клиентской части формы)
    If Not ClientAreaPos Then ClientToScreen ParentObject.hwnd, lpPoint    ' клиентской области формы в экранных координатах
' прибавляем расстояние от угла формы до контрола
    If Not IsMissing(x) Then x = lpPoint.x + TwipsToPixels(cX + dX, DIRECTION_HORIZONTAL)
    If Not IsMissing(y) Then y = lpPoint.y + TwipsToPixels(cY + dY, DIRECTION_VERTICAL)
' получаем размеры контрола
    If Not IsMissing(w) Then w = TwipsToPixels(cW, DIRECTION_HORIZONTAL)
    If Not IsMissing(h) Then h = TwipsToPixels(ch, DIRECTION_VERTICAL)
    Result = True
HandleExit:  AccControlLocation = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function PosAccForm2Screen( _
    AccForm As Access.Form, _
    ByRef x As Long, ByRef y As Long, _
    Optional Section As AcSection = acDetail)
' переводит координаты формы Access в координаты экрана (в пикселях)
' ранее AccFormCoords2Screen
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
Const lSelectorWidth = 18
Dim v As Long, h As Long
Dim lBorder As Long
Dim accPoint As POINT
Dim hwnd As Long
' переводим координаты контрола в форме в твипах в пиксели
' надо сделать с учетом смещения в ленточной форме
    With AccForm
        hwnd = .hwnd
        Select Case .DefaultView
         Case 1, 2 'ленточная или табличная
            x = x + .CurrentSectionLeft ', DIRECTION_HORIZONTAL)
            y = y + .CurrentSectionTop ', DIRECTION_VERTICAL)
         Case Else
        End Select
    End With
    accPoint.x = TwipsToPixels(x, DIRECTION_HORIZONTAL) '(X) / lTwipsPerInch * lLogPixelPerInchX
    accPoint.y = TwipsToPixels(y, DIRECTION_VERTICAL)   '(Y) / lTwipsPerInch * lLogPixelPerInchY

    'если есть рекордселектор - надо прибавить его ширину
    'If frm.RecordSelectors Then cltPoint.X = cltPoint.X + lSelectorWidth
    
    ClientToScreen& hwnd, accPoint 'переводим клиентские координаты в экранные
    
    'am 030407_10:39:00  --begin-- **************
    'учтем скроллбар
    ScrollbarGetPos AccForm.hwnd, v, h
    accPoint.x = accPoint.x - h
    If Section = acDetail Then accPoint.y = accPoint.y - v  'вертикальный скроллбар влияет только на Detail
    
    x = accPoint.x: y = accPoint.y
    Result = True
HandleExit:  PosAccForm2Screen = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function FormGetClientSize( _
    ByRef aForm As Access.Form, _
    ByRef cliWidth As Long, ByRef cliHeight As Long, _
    Optional sec As AcSection = -1, _
    Optional ByVal InTwips As Boolean = True) _
    As Boolean
' возвращает видимые размеры рабочей области формы Access/секции формы в пикселях или твипах
Const c_strProcedure = "FormGetClientSize"
' Sec = acHeader, acDetail, acFooter - высота секции, любое др значение - высота клиентской области формы

' размер клиентской области формы включает ненужные нам:
' RecordSelectors(w=19px) и NavigationButtons (h=18px), ScrollBars (h/w=17px)
' поэтому вместо размеров клиентской области формы для расчетов будем брать сумму размеров секций
' высота клиентской области формы = сумме высот секций + (кол-во видимых секций-1)*1px
' секции это окна типа "OFormSub" подчиненные окну формы hWnd в следующем порядке:
' 0-Header; 1-Detail; 2-Footer. Отсутствующие секции всё равно имеют окно высотой = 0
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    cliWidth = 0: cliHeight = 0
Dim hwnd As LongPtr: hwnd = aForm.hwnd
Dim hSec As LongPtr: hSec = FindWindowEx(hwnd, 0, accClassFormClient, vbNullString): If hSec = 0 Then Err.Raise vbObjectError + 512
Dim s As Byte, sMax As Byte: s = 0: sMax = 2
Dim hRect As RECT, tmp As Long
    With hRect
        Do
            GetWindowRect hSec, hRect: tmp = (.Bottom - .Top)
            Select Case sec
            Case acHeader: If s = 0 Then cliHeight = tmp: Exit Do
            Case acDetail: If s = 1 Then cliHeight = tmp: Exit Do
            Case acFooter: If s = 2 Then cliHeight = tmp: Exit Do
            Case Else
                If cliHeight > 0 And tmp > 0 Then tmp = tmp + 1 '1px - между секциями?
                cliHeight = cliHeight + tmp
            End Select
            hSec = FindWindowEx(hwnd, hSec, accClassFormClient, vbNullString): If hSec = 0 Then Exit Do
            s = s + 1
        Loop Until s > sMax
        cliWidth = (.Right - .Left)
    End With
' переводим пиксели в твипы
    If InTwips Then cliWidth = cliWidth * px: cliHeight = cliHeight * px
    Result = True
HandleExit:  FormGetClientSize = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Sub GetScreenInfo( _
    Optional ScrResX As Long, Optional ScrResY As Long, _
    Optional ScrDpiX As Long, Optional ScrDpiY As Long, _
    Optional DesResX As Long, Optional DesResY As Long, _
    Optional DesDpiX As Long, Optional DesDpiY As Long)
' читает размеры/разрешение текущего экрана
    On Error GoTo HandleError
Dim lpDC As LongPtr: lpDC = GetDC(HWND_DESKTOP): If lpDC = 0 Then GoTo HandleExit
Dim rcScr As RECT: Call SystemParametersInfo(SPI_GETWORKAREA, 0, rcScr, 0)
' размер рабочей области экрана в пикселях (текущее)
    With rcScr: ScrResX = (.Right - .Left): ScrResY = (.Bottom - .Top): End With
' разрешение экрана в пикселях (текущее)
    If Not IsMissing(ScrDpiX) Then ScrDpiX = GetDeviceCaps(lpDC, LOGPIXELSX)
    If Not IsMissing(ScrDpiY) Then ScrDpiY = GetDeviceCaps(lpDC, LOGPIXELSY)
    Call ReleaseDC(HWND_DESKTOP, lpDC)
'    ScrTppX = TwipsPerInch / ScrDpiX: ScrTppY = TwipsPerInch / ScrDpiY
On Error Resume Next
Dim tmpString As String
' размер рабочей области экрана в пикселях (при разработке)
    If Not IsMissing(DesResX) Or Not IsMissing(DesResY) Then
        CustomPropertyGet c_strDesignRes, tmpString, CurrentProject  ', dbText
        DesResX = Split(tmpString, c_strResDelim)(0): If DesResX = 0 Then DesResX = ScrResX
        DesResY = Split(tmpString, c_strResDelim)(1): If DesResY = 0 Then DesResY = ScrResY
    End If
' разрешение экрана в пикселях (при разработке)
    If Not IsMissing(DesDpiX) Or Not IsMissing(DesDpiY) Then
        CustomPropertyGet c_strDesignDpi, tmpString, CurrentProject ', dbText
        DesDpiX = Split(tmpString, c_strResDelim)(0): If DesDpiX = 0 Then DesDpiX = ScrDpiX
        DesDpiY = Split(tmpString, c_strResDelim)(1): If DesDpiY = 0 Then DesDpiY = ScrDpiY
    End If
    Err.Clear
HandleExit:  Exit Sub
HandleError: Err.Clear: Resume HandleExit
End Sub
Public Sub FixScreenInfo()
' сохраняет размеры/разрешение текущего экрана в свойства приложения как параметры разработки
    On Error GoTo HandleError
Dim lpDC As LongPtr: lpDC = GetDC(HWND_DESKTOP): If lpDC = 0 Then GoTo HandleExit
Dim tmpString As String
' сохраняем разрешение экрана в пикселях (текущее)
Dim rcScr As RECT: Call SystemParametersInfo(SPI_GETWORKAREA, 0, rcScr, 0)
    With rcScr
    tmpString = (.Right - .Left) & c_strResDelim & (.Bottom - .Top)
    CustomPropertySet c_strDesignRes, tmpString, CurrentProject ', dbText
    End With
' сохраняем размер рабочей области экрана в пикселях (текущее)
    tmpString = GetDeviceCaps(lpDC, LOGPIXELSX) & c_strResDelim & GetDeviceCaps(lpDC, LOGPIXELSY)
    CustomPropertySet c_strDesignDpi, tmpString, CurrentProject ', dbText
    Call ReleaseDC(HWND_DESKTOP, lpDC)
HandleExit:  Exit Sub
HandleError: Err.Clear: Resume HandleExit
End Sub
Public Function ScreenRes(Optional xRes, Optional yRes)
' Проверка разрешения экрана
Dim hdc As LongPtr: hdc = GetDC(GetDesktopWindow()): If hdc = 0 Then Exit Function
    xRes = GetDeviceCaps(hdc, HORZRES)
    yRes = GetDeviceCaps(hdc, VERTRES)
End Function
Public Function ScreenSize(Optional xSize, Optional ySize)
' Проверка размера экрана
Dim hdc As LongPtr: hdc = GetDC(GetDesktopWindow()): If hdc = 0 Then Exit Function
    xSize = GetDeviceCaps(hdc, HORZSIZE)
    ySize = GetDeviceCaps(hdc, VERTSIZE)
End Function

Public Function FormGetSectionHwnd(hWndForm As LongPtr, Seciton As Long, Optional twWidth As Long, Optional twHeight As Long) As LongPtr
'*********************************************************
'Назначение:Форма Access состоит из главного окна,
'дескриптор которого доступен через свойство hwnd формы.
'у этого окна существуют дочерние окна - по одному на каждый из
'Section формы. Имя класса этих окон - OFormSub
'как отличить одну секцию от другой точно не знаю - поэтому зашьюсь
'на их расположение- то есть самый верхний -хидер,
'потом детайл - потом футер
'Входы:hwndForm - hwnd формы (главного окна формы)
'Seciton - номер секции, может быть равен 0,1,2
'am v1.0.0_030407_11:47:16
'http://am.rusimport.ru
'appto:a_mitin@app.ru
'*********************************************************
Const c_strProcedure = "FormGetSectionHwnd"
On Error GoTo HandleError
Dim hwnd As LongPtr, sec(2) As LongPtr
Dim Result As LongPtr
Dim Top(2) As Long
Dim minVal As Long, maxVal As Long
Dim minIdx As Long, maxIdx As Long
Dim avgIdx As Long
Dim i As Long
Dim r As RECT
    minIdx = -1
    hwnd = GetWindow(hWndForm, GW_CHILD)
    Do While hwnd <> 0
        If StrComp(GetWinClass(hwnd), accClassFormClient, vbTextCompare) = 0 Then
            sec(i) = hwnd
            GetWindowRect hwnd, r
            Top(i) = r.Top
            If minIdx = -1 Or minVal >= Top(i) Then
                minIdx = i
                minVal = Top(i)
                maxIdx = i
                maxVal = Top(i)
            End If
            If maxVal < Top(i) Then
                maxIdx = i
                maxVal = Top(i)
            End If
            i = i + 1
        End If
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop
    avgIdx = 3 - (maxIdx + minIdx)
    Select Case Seciton
    Case acDetail:  Result = sec(avgIdx)
    Case acFooter:  Result = sec(maxIdx)
    Case acHeader:  Result = sec(minIdx)
    Case Else:      Err.Raise vbObjectError + 512, , "Для секции " & Seciton & " не поддерживается!"
    End Select
    GetWindowRect Result, r
    twWidth = PixelsToTwips(r.Right - r.Left, DIRECTION_HORIZONTAL)
    twHeight = PixelsToTwips(r.Bottom - r.Top, DIRECTION_VERTICAL)
HandleExit: FormGetSectionHwnd = Result: Exit Function
HandleError: Result = 0: Err.Clear: Resume HandleExit
End Function

Public Function AdjustHeight(obj As Object) As Boolean
' выравнивает отображаемую высоту формы/отчета по размерам секций
Const c_strProcedure = "AdjustHeight"
Dim o As Object
    On Error GoTo HandleError
    If TypeOf obj Is Access.Form Or TypeOf obj Is Access.Report Then
        Set o = obj
    ElseIf TypeOf obj.PARENT Is Access.Form Or TypeOf obj.PARENT Is Access.Report Then
        Set o = obj.PARENT
    Else
        GoTo HandleExit
    End If
    
Dim tmp As Long: tmp = 0
Dim i As Byte
    On Error Resume Next
    For i = acDetail To acGroupLevel2Footer:  tmp = tmp + o.Section(i).Height:  Next i
    o.InsideHeight = tmp
    Err.Clear
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Function FormGetInsideHeight(frm As Form)
' получает внутренний размер формы
    On Error Resume Next
Dim i&, Result&
    Result = 0
    With frm
        For i = acDetail To acGroupLevel2Footer
            Result = Result + .Section(i).Height
            Err.Clear
        Next i
    End With
    FormGetInsideHeight = Result
End Function

Public Function FormSectionsCount(frm As Access.Form) As Long
Dim i&, Result&, s$
    Result = 0
    With frm
        For i = acDetail To acGroupLevel2Footer
            Err.Clear
            s = .Section(i).Name
            If Err = 0 Then Result = Result + 1
        Next i
    End With
    FormSectionsCount = Result
End Function

Public Function GetOpenedObjects(Optional ListOfObjectNames As String)
' возвращает массив открытых объектов Access
'Const c_strProcedure = "GetOpenedObjects"
'Dim arrNames() As String
'Dim i As Long
'Dim obj, nam
'Dim ObjName As String
'Dim bolCheckList As Boolean
'Dim bolAddName As Boolean
'Dim Result() As Access.Form
'
'    On Error GoTo HandleError
'    bolCheckList = Len(Trim(ListOfObjectNames))
'    ' ищем только объекты из списка
'    If bolCheckList Then arrNames = Split(ListOfObjectNames, ";")
'    For Each obj In Application.Forms
'' проверяем открытые формы
'        bolAddName = False
'        ObjName = obj.Name
''        If IsFormOpened(strFrmName) Then
''            IsFormLoaded = Forms(strFrmName).CurrentView <> conDesignView
''        End If
'
'        If IsFormLoaded(ObjName) Then
'            If bolCheckList Then
'    ' проверяем форму по списку
'                For Each nam In arrNames
'                    bolAddName = ObjName = nam
'                    If bolAddName Then Exit For
'                Next nam
'            Else
'                bolAddName = True
'            End If
'        End If
'        If bolAddName Then Call InsertElementIntoArray(Result, obj) 'ObjName)
'    Next obj
'    For Each obj In Application.Reports
'' проверяем открытые отчеты
'        bolAddName = False
'        ObjName = obj.Name
'        If IsReportLoaded(ObjName) Then
'            If bolCheckList Then
'    ' проверяем отчет по списку
'                For Each nam In arrNames
'                    bolAddName = ObjName = nam
'                    If bolAddName Then Exit For
'                Next nam
'            Else
'                bolAddName = True
'            End If
'        End If
'        If bolAddName Then Call InsertElementIntoArray(Result, obj) 'ObjName)
'    Next obj
'HandleExit:
'    GetOpenedObjects = Result
'    Exit Function
'HandleError:
'    'Result = Null
''    Dbg.Error Err.Number, Err.Description, Err.Source, Erl(), c_strModule & "." & c_strProcedure
'    Err.Clear
'    Resume HandleExit
End Function

Public Sub ListboxUnSelectAll(lst As ListBox)
'Precondition: MultiSelect > 0
'None: 0 (Default) Multiple selection isn't allowed.
'Simple: 1
'Extended: 2
Dim varItem As Variant

    For Each varItem In lst.ItemsSelected
        lst.Selected(varItem) = False
    Next
End Sub

Public Sub ListboxSelectAll(lst As ListBox)
'Precondition: MultiSelect > 0
'None: 0 (Default) Multiple selection isn't allowed.
'Simple: 1
'Extended: 2
    Dim lngRow As Long

    For lngRow = 0 To lst.ListCount - 1
        lst.Selected(lngRow) = True
    Next
End Sub

Public Function ListboxToInStr( _
    ListCtrl As ListBox, _
    Optional dtFormat As eFieldFormat = vbLong, _
    Optional ListDelim As String = ",") As String
'Возвращает список значений списка с мультивыделением
'разделенный запятыми внутри выражения In(...) или Not In (...)
'  в зависимости от соотношения ВЫДЕЛЕННОГО и НЕ ВЫДЕЛЕННОГО
' а также
'   - "True" если отобраны все значения
'   - "False" если нет отобраных знчений
Const c_strProcedure = "ListboxToInStr"

'Dim UseSelected As Boolean  'Что собираем - ВЫДЕЛЕННОЕ или НЕ ВЫДЕЛЕННОЕ
'Dim strIn As String         'выражение - In (...) или Not In (...)
'Dim i As Integer, iMax As Integer, iSel As Integer 'индекс значения списка
'Dim tmpVal
'Dim Result As String
'On Error GoTo HandleError
''Проверка на наличие выбранных элем. списка
'    strIn = sqlIn
'    Result = vbNullString 'sqlFalse
'    UseSelected = True
'    With ListCtrl
'        i = 0: iMax = .ListCount: iSel = .ItemsSelected.Count
'        If iSel = 0 Then
'            Result = sqlFalse
'            GoTo HandleExit
'         ElseIf iSel = iMax Then
'            Result = sqlTrue
'            GoTo HandleExit
'        End If
'    'Определяем чего больше ВЫДЕЛЕННОГО или НЕ ВЫДЕЛЕННОГО
'    ' (не вижу смысла - список может быть отфильтрован (неполным)
'    ' тогда полученый по Not IN () фильтр не будет соответвовать выбору)
''        UseSelected = (iSel < CInt(0.5*iMax ))
'    'Сборка строки
'        Do While i < iMax
'            If .Selected(i) = UseSelected Then
'                tmpVal = .ItemData(i)
'                If IsNull(tmpVal) Then
'                    Result = Result & sqlNull & ListDelim
'                'ElseIf Len(tmpVal) = 0 Then
'                '    Result = Result & sqlNull & ListDelim
'                 Else
'                    Select Case dtFormat
'                     Case vbString  'текстовые значения
'                        Result = Result & "'" & Replace(CStr(tmpVal), """", """""") & "'" & ListDelim
'                     Case vbInteger, vbLong, vbByte 'числовые значения
'                        Result = Result & CStr(tmpVal) & ListDelim
'                     Case vbDate 'date значения
'                        Result = Result & DateToSQL(CDate(tmpVal)) & ListDelim
'                     Case vbDateTimeJ 'dateTime значения
'                        Result = Result & DateTimeToSQL(CDate(tmpVal)) & ListDelim
'                     Case vbBoolean 'boolean значения
'                        Result = Result & IIf(CBool(tmpVal), sqlTrue, sqlFalse) & ListDelim    '& "#"
'                     Case vbSingle, vbDouble, vbCurrency 'с десятичной точкой
'                        Result = Result & Replace(CStr(tmpVal), ",", ".") & ListDelim
'                     Case Else
'                        Result = vbNullString '" Конструкция In (...) для такого типа данных не предусмотрена"
'                        GoTo HandleExit
'                    End Select
'                End If
'            End If
'            i = i + 1
'        Loop
'    End With
'    If Right$(Result, Len(ListDelim)) = ListDelim Then Result = Trim$(Left$(Result, Len(Result) - Len(ListDelim)))
'    If Len(Result) = 0 Then GoTo HandleExit
'    If Not UseSelected Then strIn = sqlNot & LTrim$(strIn)
'    Result = strIn & "(" & Result & ")"
'HandleExit:
'    ListboxToInStr = Result
'    Exit Function
'HandleError:
'    Dbg.Error Err.Number, Err.Description, Err.Source, Erl(), c_strModule & "." & c_strProcedure
'    Err.Clear
'    Result = vbNullString
'    Err.Clear
'    Resume HandleExit
End Function

Public Function IsControlExists(frm As Form, ctlName As String) As Boolean
    Const c_strProcedure = "IsControlExists"
    On Error Resume Next
Dim strValue As String
    ' If you can retrieve the value, the such control exists.
    strValue = frm.Controls(ctlName).Name
    IsControlExists = (Err.Number = 0)
    Err.Clear
End Function

Public Function GetTopParent(accObject, Optional AllowSubForms As Boolean) As Object 'Access.Form
' возвращает ссылку на верхнюю родительскую форму(отчёт) для объекта
' AllowSubForms - если True будет останавливаться на форме/отчёте в том числе открытых как субформа(-отчёт),
'                 иначе будет подниматься выше пока не найдёт форму (отчёт) без родителей
Dim Result As Object
    On Error GoTo HandleError
' перебираем родителей контрола пока не доберётся до верхнего (формы или отчёта)
    Set Result = accObject
    Do While (TypeOf Result Is Access.Control): Set Result = Result.PARENT: Loop
    If AllowSubForms Then GoTo HandleExit
    On Error Resume Next
' перебираем родителей формы или отчёта пока не до берёмся до верхней формы (или отчёта)
Dim strName As String
    Do
        strName = Result.ParentName
        If Err Then Exit Do
        Set Result = Result.PARENT
    Loop
    Err.Clear
HandleExit:  Set GetTopParent = Result: Exit Function
HandleError: Resume HandleExit
End Function

Public Function ActiveControlHwndGet() As LongPtr
Dim CtlHwnd As LongPtr
' Windows координаты активного контрола
' сохраняет hWnd в глобальной переменной app.AccessCtlHwnd
' нужным контролам в событие GotFocus вставить вызов этой функции
    CtlHwnd = GetFocus() ' hWnd появляется только в момент получения фокуса
    'app.AccessCtlHwnd = ctlHwnd
    ActiveControlHwndGet = CtlHwnd
End Function

Public Function ControlHwndGet(ctl As Access.Control) As LongPtr
' Windows координаты контрола
    ctl.SetFocus ' контрол Access имеет hWnd только если в нем фокус
    ControlHwndGet = ActiveControlHwndGet
End Function

Public Function ControlRectGet(ctl As Access.Control) As RECT
' Windows координаты контрола
Dim ControlRect As RECT
    'ctl.SetFocus
    'GetWindowRect ControlHwndGet(ctl), ControlRect
    ' контролам Access hWnd присваивается лишь
    ' в момент события GotFocus.
    GetWindowRect ControlHwndGet(ctl), ControlRect
    ControlRectGet = ControlRect
End Function

'=================================
Public Function WindowMove(hwnd As LongPtr, Optional x, Optional y, Optional Arrange As eAlign = eAlignLeftTop, Optional Inscribe As Boolean = True) As Long
' Позиционирование окна на экране
' frm - позиционируемая форма
' x,y - экранные координаты точки относительно которой позиционируется форма
' Arrange - стиль позиционирования
' Inscribe - если установлен, - при позиционировании
'   стиль подбирается так чтобы форма помещалась в видимой области Access
Dim accHwnd As LongPtr
Dim accRect As RECT, accHeight As Long, accWidth As Long    ' размеры клиентской области Access
Dim winRect As RECT, winHeight As Long, winWidth As Long    ' размеры выпадающей формы
Dim xPos As Long, yPos As Long
    On Error GoTo HandleError
    
' ищем клиентскую область окна Access
    accHwnd = FindWindowEx(hWndAccessApp, 0&, accClassChild, vbNullString) ' MDIClient
    GetWindowRect accHwnd, accRect
    With accRect
        accHeight = (.Bottom - .Top)
        accWidth = (.Right - .Left)
    End With
' получаем координаты области окна
    GetWindowRect hwnd, winRect
    With winRect
        winHeight = (.Bottom - .Top)
        winWidth = (.Right - .Left)
    End With
' относительно x (разместить горизонтально)
    If (Arrange And eCenterHorz) = eCenterHorz Then
    ' центрировать
        xPos = x + winWidth \ 2
    ElseIf (Arrange And eLeft) = eLeft Then
    ' выровнять по левому (разместить справа от точки)
        xPos = x
    ElseIf (Arrange And eRight) = eRight Then
    ' выровнять по правому (разместить слева от точки)
        xPos = x - winWidth
    End If
' относительно y (разместить вертикально)
    If (Arrange And eCenterVert) = eCenterVert Then
    ' центрировать
        yPos = y + winHeight \ 2
    ElseIf (Arrange And eTop) = eTop Then
    ' выровнять по верхнему (разместить ниже точки)
        yPos = y
    ElseIf (Arrange And eBottom) = eBottom Then
    ' выровнять по нижнему (разместить выше точки)
        yPos = y - winHeight
    End If
' пытаемся изменить чтобы форма вся помещалась в экран
    If Inscribe Then
    ' если указано вписать в рабочую область
        If xPos < accRect.Left Then xPos = accRect.Left
        If xPos + winWidth > accRect.Right Then xPos = accRect.Right - winWidth
        If yPos < accRect.Top Then yPos = accRect.Top
        If yPos + winHeight > accRect.Bottom Then yPos = accRect.Bottom - winHeight
    End If
    MoveWindow hwnd, xPos, yPos, winWidth, winHeight, 1
HandleExit:  Exit Function 'x = xPos: y = yPos
HandleError: Resume HandleExit
End Function
Public Function WindowSet(hwnd As LongPtr, nCmdShow As Long) As Long
' устанавливает стиль окна
Dim loX  As Long
Dim loForm As Object
Dim Result As Boolean: Result = False
    On Error Resume Next
    If hwnd = 0 Then GoTo HandleExit
    Set loForm = Screen.ActiveForm
    If Err <> 0 Then 'нет активной формы
        If nCmdShow = SW_HIDE Then
            MsgBox "Не могу скрыть окно пока отсутствует форма на экране"
        Else
            loX = ShowWindow(hwnd, nCmdShow): Err.Clear
        End If
     Else
        If nCmdShow = SW_SHOWMINIMIZED And loForm.Modal = True Then
            MsgBox "Не могу минимизировать окно" & (loForm.Caption + " ") & "пока форма на экране"
        ElseIf nCmdShow = SW_HIDE And loForm.PopUp <> True Then
           MsgBox "Не могу скрыть окно " & (loForm.Caption + " ") & "пока форма на экране"
        Else
            loX = ShowWindow(hwnd, nCmdShow)
        End If
    End If
    Result = (loX <> 0)
HandleExit:  WindowSet = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
'==============================
' Преобразование размеров
'==============================
Public Function AccClientSize(Optional Width As Long, Optional Height As Long) As RECT
' ищем клиентскую область окна Access
Dim tmpHwnd As LongPtr
Dim tmpRect As RECT
    tmpHwnd = FindWindowEx(hWndAccessApp, 0&, accClassChild, vbNullString)
    GetWindowRect tmpHwnd, tmpRect
'    GetClientRect hWndAccessApp, tmpRect
    With tmpRect
        Width = PixelsToTwips(.Right - .Left, DIRECTION_HORIZONTAL)
        Height = PixelsToTwips(.Bottom - .Top, DIRECTION_VERTICAL)
    End With
    AccClientSize = tmpRect
End Function

Public Function ScrollbarGetPos(FormHwnd As LongPtr, ByRef VSB_Pos As Long, ByRef HSB_Pos As Long)
'Назначение:получаем текущее расположение скролбара (если есть)
On Error GoTo HandleError
Dim hWndVSB As LongPtr, hWndHSB As LongPtr
    hWndHSB = ScrollbarGetHwnd(FormHwnd, SBS_HORZ): If hWndHSB = 0 Then HSB_Pos = 0 Else HSB_Pos = GetScrollPos(hWndHSB, SB_CTL)
    hWndVSB = ScrollbarGetHwnd(FormHwnd, SBS_VERT): If hWndVSB = 0 Then VSB_Pos = 0 Else VSB_Pos = GetScrollPos(hWndVSB, SB_CTL)
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Function ScrollbarSetPos(FormHwnd As LongPtr, VSB_Pos As Long, HSB_Pos As Long)
' устанавливаем положение скролбара (если есть)
On Error GoTo HandleError
Dim hWndVSB As LongPtr, hWndHSB As LongPtr
    hWndHSB = ScrollbarGetHwnd(FormHwnd, SBS_HORZ)
    hWndVSB = ScrollbarGetHwnd(FormHwnd, SBS_VERT)
    If hWndHSB <> 0 And HSB_Pos >= 0 Then Call SendMessage&(FormHwnd, WM_HSCROLL, (HSB_Pos * 2 ^ 16) Or SB_THUMBPOSITION, 0) 'SetScrollPos hwndHSB, SB_CTL, HSB_Pos, 1
    If hWndVSB <> 0 And VSB_Pos >= 0 Then Call SendMessage&(FormHwnd, WM_VSCROLL, (VSB_Pos * 2 ^ 16) Or SB_THUMBPOSITION, 0)  'SetScrollPos hwndVSB, SB_CTL, VSB_Pos, 1
HandleExit:  Exit Function
HandleError: Resume HandleExit
End Function
Public Function ScrollbarGetHwnd(FormHwnd As LongPtr, ScrollbarType As Integer) As LongPtr
'Назначение:получаем hwnd скролбара формы
On Error GoTo HandleError
Dim hWndChild As LongPtr
Dim s As String
Dim Style&
    'находим всех детей формы - и ищем скроллбары
    hWndChild = GetWindow(FormHwnd, GW_CHILD)
    If hWndChild = 0 Then
        ScrollbarGetHwnd = 0
    Else
        Do
            s = GetWinClass(hWndChild)
            If StrComp(s, "SCROLLBAR", vbTextCompare) = 0 Then
                'это скролбар - проверим тип
                Style& = GetWindowLong&(hWndChild, GWL_STYLE)
                If (Style& And SBS_SIZEBOX) = False _
                        And (Style& And &H1) = SBS_HORZ Then
                    'горизонтальный
                    If ScrollbarType = 0 Then
                        'нашли
                        ScrollbarGetHwnd = hWndChild ' GetScrollPos(hwndChild, SB_CTL)
                        Exit Function
                    End If
                End If
                If (Style& And &H1) = SBS_VERT Then
                    'вертикильный
                    If ScrollbarType = 1 Then
                        'нашли
                        ScrollbarGetHwnd = hWndChild 'GetScrollPos(hwndChild, SB_CTL)
                        Exit Function
                    End If
                End If
            End If
            hWndChild = GetWindow(hWndChild, GW_HWNDNEXT)
        Loop While hWndChild <> 0
    End If
HandleExit:  Exit Function
HandleError: Resume HandleExit
End Function
Private Function StrZ(Par As String) As String
Dim nSize As Long, i As Long ', Rez As String
   nSize = Len(Par)
   i = InStr(1, Par, Chr(0)) - 1
   If i > nSize Then i = nSize
   If i < 0 Then i = nSize
   StrZ = Mid$(Par, 1, i)
End Function
Public Function GetWinClass(hwnd As LongPtr) As String
' получаем класс окна
Dim Buff As String, BuffSize As Long
    Buff = Space(255)
    BuffSize = GetClassName(hwnd, Buff, 255)
    GetWinClass = StrZ(Buff)
End Function
Public Function GetWinText(hwnd As LongPtr) As String
' получаем класс окна
Dim Buff As String, BuffSize As Long
    BuffSize = GetWindowTextLength(hwnd) + 1            ' длина заголовка
    Buff = Space(BuffSize)                              ' резервируем текстовой буфер
    BuffSize = GetWindowText(hwnd, Buff, BuffSize)      ' получаем заголовок
    GetWinText = StrZ(Buff)
End Function
Public Function ChildWindowProc(ByVal hwnd As LongPtr, ByVal lp As Long) As LongPtr
' вызывается функцией EnumChildWindows
' параметры hWnd и lp передаются функцией EnumChildWindows
Dim RetVal As Long
'Static i As Long
'Dim strClass As String
'Dim strText As String
    
    RetVal = 1
'    i = i + 1
'
'    strClass = GetWinClass(hWnd) ' имя класса окна
'    strText = GetWinText(hWnd)
'    Select Case strClass
'     Case accClassFormWindow = "OForm"       ' класс окна формы Access
'     Case accClassFormClient = "OFormSub"    ' класс формы Access
'        ' область секции формы либо, подформы
'     Case accClassFormPopup = "OFormPopup"   ' класс всплывающей формы Access
'     Case accClassFormChild = "OFormChild"   ' класс подчиненной формы Access
'     Case accClassFormNoClose = "OFormNoClose"
'     Case accClassFormClientChild = "OFEDT"  ' класс подчиненного окна формы Access
'     Case accClassTableClientChild = "OGNUM" ' класс подчиненного окна табличной формы Access
'     Case accClassRecordSlector = "OSUI"     ' класс формы селектора записей Access
'     Case accClassTextbox = "OKttbx"         ' класс текстбокса формы Access
'    End Select
' Debug.Print i; "Window: "; hWnd; " Class: "; strClass; " Text: "; strText
    ChildWindowProc = RetVal ' 1 - продолжить перечисление 0 - прервать
End Function

'=========================
' прообразование размеров
'=========================
Public Function TwipsToPixels(ByVal lngTwips As Long, lngDirection As Long) As Long
Const c_strProcedure = "PixelsToTwips"
On Error GoTo HandleError
'   Function to convert Twips to pixels for the current screen resolution
'   Accepts:
'       lngTwips - the number of twips to be converted
'       lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
'   Returns:
'       the number of pixels corresponding to the given twips
Dim hdc As LongPtr
Dim lngPixelsPerInch As Long
    hdc = GetDC(0)
    If lngDirection = DIRECTION_HORIZONTAL Then
        lngPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSX)
     Else
        lngPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSY)
    End If
    hdc = ReleaseDC(0, hdc)
    TwipsToPixels = lngTwips / TwipsPerInch * lngPixelsPerInch
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function PixelsToTwips(ByVal lngPixels As Long, Optional lngDirection As Long = DIRECTION_HORIZONTAL) As Long
'   Function to convert pixels to twips for the current screen resolution
'   Accepts:
'       lngPixels - the number of pixels to be converted
'       lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
'   Returns:
'       the number of twips corresponding to the given pixels
On Error GoTo HandleError
Dim hdc As LongPtr
Dim lngPixelsPerInch As Long
    hdc = GetDC(0)
    If lngDirection = DIRECTION_HORIZONTAL Then
        lngPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSX)
     Else
        lngPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSY)
    End If
    hdc = ReleaseDC(0, hdc)
    PixelsToTwips = lngPixels * TwipsPerInch / lngPixelsPerInch
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function HimetrixToPixels(ByVal lngHiMetrix As Long, lngDirection As Long) As Long
Dim Result As Long
    On Error GoTo HandleError
    ' conversion from Himetrics to Pixels when ScaleX/Y is not available
    If lngDirection = DIRECTION_HORIZONTAL Then
        Result = lngHiMetrix * TwipsPerInch / 2540 / TwipsPerPixel(LOGPIXELSX) 'Screen.TwipsPerPixelX
    Else
        Result = lngHiMetrix * TwipsPerInch / 2540 / TwipsPerPixel(LOGPIXELSY) 'Screen.TwipsPerPixelY
    End If
HandleExit:  HimetrixToPixels = Result: Exit Function
HandleError: Result = False: Resume HandleExit
End Function
Public Function TwipsPerPixel(Optional ByVal Dimension As Long = LOGPIXELSX) As Long
Dim hdc As LongPtr
    On Error GoTo HandleError
    hdc = GetDC(Application.hWndAccessApp) 'DesktopDC = GetDC(HWND_DESKTOP)
    TwipsPerPixel = TwipsPerInch / GetDeviceCaps(hdc, Dimension)
HandleError:
    Call ReleaseDC(Application.hWndAccessApp, hdc) 'Call ReleaseDC(HWND_DESKTOP, DesktopDC)
  'Call Exception.RaiseAgain
End Function
Public Function CreateHFontByControl(Optional ctl As Variant, _
    Optional FontName, Optional FontSize, Optional FontColor, Optional FontWeight, Optional FontUnderline, Optional FontStrikeOut, Optional FontItalic, _
    Optional hdc As LongPtr) As LongPtr
' создает hFont из параметров контрола
Const c_strProcedure = "CreateHFontbyControl"
Dim Result As LongPtr ': Result = 0
On Error Resume Next
Dim tDC As LongPtr
    'If Not TypeOf ctl Is Access.Control Then Err.Raise vbObjectError + 512
    If hdc = 0 Then tDC = GetDC(0) Else tDC = hdc
' создаём шрифт
Dim fName As String:    fName = IIf(IsMissing(FontName), ctl.FontName, FontName): If Err Then fName = "Arial": Err.Clear
Dim fSize As Long:      fSize = IIf(IsMissing(FontSize), ctl.FontSize, FontSize): If Err Then fSize = 10: Err.Clear
Dim fColor As Long:     fColor = IIf(IsMissing(FontColor), ctl.FontColor, FontColor): If Err Then fColor = vbBlack: Err.Clear
Dim fWeight As Long:    fWeight = IIf(IsMissing(FontWeight), ctl.FontWeight, FontWeight): If Err Then fWeight = 0: Err.Clear
Dim fItalic As Long:    fItalic = IIf(IsMissing(FontItalic), ctl.FontItalic, FontItalic): If Err Then fItalic = False: Err.Clear
Dim fUnderline As Long: fUnderline = IIf(IsMissing(FontUnderline), ctl.FontUnderline, FontUnderline): If Err Then fUnderline = False: Err.Clear
Dim fStrikeOut As Long: fStrikeOut = IIf(IsMissing(FontStrikeOut), ctl.FontStrikeOut, FontStrikeOut): If Err Then fStrikeOut = False: Err.Clear
On Error GoTo HandleError
    'FontSize = -(FontSize * PT / TwipsPerPixel)
    'fSize = -Int(fSize * GetDeviceCaps(tDC, LOGPIXELSY) / 72)
    fSize = -MulDiv(fSize, GetDeviceCaps(tDC, LOGPIXELSY), 72)
    Result = CreateFont(fSize, 0, 0, 0, _
        fWeight, fItalic, fUnderline, fStrikeOut, _
        RUSSIAN_CHARSET, 0, 0, ANTIALIASED_QUALITY, 0, fName)  ' PROOF_QUALITY | CLEARTYPE_QUALITY | ANTIALIASED_QUALITY
    If hdc = 0 Then ReleaseDC 0, tDC
HandleExit:  CreateHFontByControl = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
'=========================
' Функции чтения/записи свойств из тэга
'=========================
Public Function TagsRead(accObject As Object, Optional Init As Boolean, Optional Col As Collection) As Boolean
' читает доп свойства из Tag объекта в дополнительные свойства
'-------------------------
' accObject   - ссылка на объект (Form,Section,Control) из тэга которого читаются дополнительные свойства
' Init      - если True читаются все свойства
'-------------------------
Dim Result As Boolean ': Result = False
    On Error GoTo HandleError
'Stop
Dim i As Long, iMax As Long
Dim ctl As Control, tmp ': If col Is Nothing Then Set col = New Collection
    If TypeOf accObject Is Access.Form Or TypeOf accObject Is Access.Report Then
' Если объект - форма/отчёт
    ' пробегаем все секции
        Set Col = New Collection
        For i = acDetail To acFooter:
        Call TagsRead(accObject.Section(i), Init, Col):
        Next i
        Set Col = Nothing
    Else
    On Error Resume Next 'GoTo HandleError
        i = accObject.Controls.Count
        If Err = 0 And i > 0 Then
' объекта c подчинёнными контролами в коллекции Controls (Section, Tabs, Groups, Switchboard etc)
    ''' !!! надо отдельно обрабатывать встроенные формы и группы контролов
    '' Если объект - контрол у которого есть родитель или привязка - рекурсивно читаем их
    '' т.к. прежде чем выстраивать данный контрол надо задать позиции объектов относительно которых он задан
            For Each ctl In accObject.Controls
            ' если контрол ещё не обработан - обрабатываем если уже обработан - пропускаем
                Set tmp = Col(ctl.Name): If Err Then Err.Clear: Call TagsRead(ctl, Init, Col)
            Next ctl
        Else
            Err.Clear: If TypeOf accObject Is Access.SubForm Then Call TagsRead(accObject.Form, Init) ' у субформы будет своя коллекция проверенных контролов
        End If
    End If
'----------------------------------
Dim tmpVal As String
Dim ctlEx As tControlEx: ctlEx = p_TypeControlExInit
    With ctlEx 'objTo
    If Len(Trim(accObject.Tag)) = 0 Then If TypeOf accObject Is Access.Section Then GoTo HandleColor Else Exit Function
    On Error Resume Next 'GoTo HandleError
Dim Tags As New Collection, Keys As Variant
    Call TaggedString2Collection(accObject.Tag, Tags, Keys): i = LBound(Keys): iMax = UBound(Keys)
'    If (objTo Is Nothing) Then Set objTo = accObject
' ----------------------------------
' чтение доп свойств из тэга в типизированную переменную / объект (пока не решил)
' ----------------------------------
' ToDo: для некоторых элементов оформления можно сделать так:
'       если не задан элемент но задан для родителя - читать по родителю.
'       например м.б. актуально для субформы - перенос стиля оформления секции на субформу
    For i = LBound(Keys) To UBound(Keys) 'Each sKey In Keys
        tmpVal = Tags(i)
        Select Case Keys(i) 'tmpVal
' стиль оформления
    ' цветовое оформление
        Case adhcBackColor:                                 If Len(tmpVal) > 0 Then .BackColor = GetColorFromText(UCase(tmpVal))
        Case adhcTextColor, adhcFontColor, adhcForeColor:   If Len(tmpVal) > 0 Then .TextColor = GetColorFromText(UCase(tmpVal))
        Case adhcBorderColor:                               If Len(tmpVal) > 0 Then .BorderColor = GetColorFromText(UCase(tmpVal))
' предельные размеры
        Case adhcMinWidth, adhcMinWidth1:                   If Len(tmpVal) > 0 Then .MinWidth = GetSizeFromText(tmpVal)
        Case adhcMaxWidth, adhcMaxWidth1:                   If Len(tmpVal) > 0 Then .MaxWidth = GetSizeFromText(tmpVal)
        Case adhcMinHeight, adhcMinHeight1:                 If Len(tmpVal) > 0 Then .MinHeight = GetSizeFromText(tmpVal)
        Case adhcMaxHeight, adhcMaxHeight1:                 If Len(tmpVal) > 0 Then .MaxHeight = GetSizeFromText(tmpVal)
' положение
        Case adhcDefLeft, adhcDefLeft1, adhcDefLeft2:       If Len(tmpVal) > 0 Then .DefLeft = GetSizeFromText(tmpVal)
        Case adhcDefTop, adhcDefTop1, adhcDefTop2:          If Len(tmpVal) > 0 Then .DefTop = GetSizeFromText(tmpVal)
        Case adhcDefWidth, adhcDefWidth1:                   If Len(tmpVal) > 0 Then .DefWidth = GetSizeFromText(tmpVal)
        Case adhcDefHeight, adhcDefHeight1:                 If Len(tmpVal) > 0 Then .DefHeight = GetSizeFromText(tmpVal)
        Case adhcDefRight, adhcDefRight1:                   If Len(tmpVal) > 0 Then .DefRight = GetSizeFromText(tmpVal)
        Case adhcDefBottom, adhcDefBottom1:                 If Len(tmpVal) > 0 Then .DefBottom = GetSizeFromText(tmpVal)
' привязка контрола (получаем из имени контрола относительно границ которого определены координаты ссылку на него)
        Case adhcBondLeft, adhcBondLeft1, adhcBondLeft2:    If Len(tmpVal) > 0 Then Set .BondLeft = accObject.Form.Controls(tmpVal): Call TagsRead(.BondLeft, Init, Col)
        Case adhcBondTop, adhcBondTop1, adhcBondTop2:       If Len(tmpVal) > 0 Then Set .BondTop = accObject.Form.Controls(tmpVal): Call TagsRead(.BondTop, Init, Col)
        Case adhcBondWidth, adhcBondWidth1:                 If Len(tmpVal) > 0 Then Set .BondWidth = accObject.Form.Controls(tmpVal): Call TagsRead(.BondWidth, Init, Col)
        Case adhcBondHeight, adhcBondHeight1:               If Len(tmpVal) > 0 Then Set .BondHeight = accObject.Form.Controls(tmpVal): Call TagsRead(.BondHeight, Init, Col)
        Case adhcBondRight, adhcBondRight1:                 If Len(tmpVal) > 0 Then Set .BondRight = accObject.Form.Controls(tmpVal): Call TagsRead(.BondRight, Init, Col)
' стиль вывода (adhcSizeIt/adhcFloatIt/adhcStyleIt)
        Case adhcStyleIt:                                   .StyleIt = GetStyleFromText(tmpVal)
        Case adhcSizeIt: Select Case tmpVal             ' растяжение (перекрывается StyleIt)
        Case adhcSizeRight:                                 .StyleIt = .StyleIt Or lsHorz           ' растягивать по горизонтали (вправо)
        Case adhcSizeBottom:                                .StyleIt = .StyleIt Or lsVert           ' растягивать по вертикали (вниз)
        Case adhcSizeBoth:                                  .StyleIt = .StyleIt Or lsFull           ' растягивать по горизонтали и вертикали (вправо-вниз)
        End Select
        Case adhcFloatIt: Select Case tmpVal            ' смещение (перекрывается StyleIt)
        Case adhcFloatRight:                                .StyleIt = .StyleIt Or lsRight          ' смещать по горизонтали (вправо)
        Case adhcFloatBottom:                               .StyleIt = .StyleIt Or lsBottom         ' смещать по вертикали (вправо)
        Case adhcFloatBoth:                                 .StyleIt = .StyleIt Or lsRightBottom    ' смещать по горизонтали и вертикали (вправо-вниз)
        End Select
        Case adhcScaleIt:                                   .ScaleIt = GetBoolFromText(tmpVal)      ' масштабировать
' изображение/текст
        Case adhcIconName:                                  .IconName = tmpVal
        ' свойства выводимого изображения
        Case adhcPictName:                                  .PictName = tmpVal
        Case adhcPictSize:                                  If Len(tmpVal) > 0 Then .PictSize = tmpVal
        Case adhcPictAngle:                                 If Len(tmpVal) > 0 Then .PictAngle = tmpVal
        Case adhcPictMode:                                  If Len(tmpVal) > 0 Then .PictMode = GetPictModeFromText(tmpVal) 'Else .PictMode = apObjSizeZoomDown
        Case adhcPictAlign:                                 If Len(tmpVal) > 0 Then .PictAlign = GetPositionFromText(tmpVal) 'Else .PictAlign = eCenter
        Case adhcPictGray:                                  If Len(tmpVal) > 0 Then .PictGray = GetBoolFromText(tmpVal)
        ' свойства выводимого текста
        Case adhcPictText:                                  .PictText = tmpVal
        Case adhcTextAngle:                                 If Len(tmpVal) > 0 Then .TextAngle = tmpVal
        Case adhcTextAlign:                                 If Len(tmpVal) > 0 Then .TextAlign = GetAlignFromText(tmpVal) 'Else .TextAlign = TA_LEFT
        Case adhcTextPlace:                                 If Len(tmpVal) > 0 Then .TextPlace = GetPlaceFromText(tmpVal) 'Else .TextPlace = ePlaceOnRight
        ' свойства шрифта выводимого текста
        Case adhcFontName:                                  If Len(tmpVal) > 0 Then .FontName = tmpVal
        Case adhcFontSize:                                  If Len(tmpVal) > 0 Then .FontSize = tmpVal
' поведение
        Case adhcSplitIt:                                   .SplitIt = GetSplitFromText(tmpVal)     ' признак разделителя (сплиттера)
        'Case adhcAction:                                    .Action = tmpVal                        ' заданное для контрола действие
        Case adhcBondBottom, adhcBondBottom1:               If Len(tmpVal) > 0 Then Set .BondBottom = accObject.Form.Controls(tmpVal): Call TagsRead(.BondBottom, Init, Col)
        End Select
HandleNext: Err.Clear
    Next
' ----------------------------------
' применяем прочитанные свойства на объект
' ----------------------------------
HandleSet: Err.Clear
' позиция/размеры
    '' положение по-умолчанию
    '' если нет привязки позиция берётся относительно родителя
    '' для расчёта позиции:
    '' 1) проверить наличие привязки - если привязка не задана берём привязку по родителю
    '' 2) определить к какой границе связанного контрола привязан текущий
    ''    если контрол лежит внутри контрола к которому привязан - левая граница привязана к левой,- правая к правой  (+ - внутрь - - наружу)
    ''    если правая граница контрола левее левой границы родителя - и правая и левая заданы относительно левой границы родителя
    ''    если левая граница контрола правее правой границы родителя - и правая и левая заданы относительно правой границы родителя
    '' 3) если задан стиль вывода - получаем положение из стиля
    ''    если попложение конфликтует со стилем вывода,- предпочитаем положение
    
    'If .DefLeft Then
    'If .DefTop Then
    'If .DefWidth Then
    'If .DefHeight Then
    'If .DefRight Then
    'If .DefBottom Then
    ' предельные размеры
    '  .MinWidth = GetSizeFromText(tmpVal) ', .GetBoundInTwips(eWidth))
    'Case adhcMaxWidth, adhcMaxWidth1:                   .MaxWidth = GetSizeFromText(tmpVal) ', .GetBoundInTwips(eWidth))
    'Case adhcMinHeight, adhcMinHeight1:                 .MinHeight = GetSizeFromText(tmpVal) ', .GetBoundInTwips(eHeight))
    'Case adhcMaxHeight, adhcMaxHeight1:                 .MaxHeight = GetSizeFromText(tmpVal) ', .GetBoundInTwips(eHeight))
' стиль отображения                                             (adhcSizeIt/adhcFloatIt/adhcStyleIt)
        'If .StyleIt <> lsUndef Then   ' стиль
        
        'If .ScaleIt <> sfNo Then      ' масштабировать
' поведение
        'If .SplitIt <>cdNone Then     ' разделитель
' изображение/текст
#If UseModPictureData Then
    If TypeOf accObject Is Access.Control Then If Len(.PictName) = 0 And Len(.IconName) > 0 Then .PictName = .IconName
    If Len(.PictName) > 0 Or Len(.PictText) > 0 Then Call PictureData_SetToControl(accObject, _
                ObjectData:=.PictName, Alignment:=.PictAlign, PictSizeMode:=.PictMode, PictWidth:=.PictSize, PictAngle:=.PictAngle, GrayScale:=.PictGray, _
                TextString:=.PictText, TextPlacement:=.TextPlace, TextAlignment:=.TextAlign, TextAngle:=.TextAngle)
    If TypeOf accObject Is Access.Form Or TypeOf accObject Is Access.Report Then If Len(.IconName) > 0 Then Call PictureData_SetIcon(accObject.hwnd, ObjectData:=.IconName)
#End If
' цветовое оформление
HandleColor:
    If (.BackColor = COLORNOTSET) Then If TypeOf accObject Is Access.Section Then tmpVal = TaggedStringGet(accObject.PARENT.Tag, adhcBackColor): If Len(tmpVal) > 0 Then .BackColor = GetColorFromText(UCase(tmpVal))
    If (.BackColor <> COLORNOTSET) Then accObject.BackColor = .BackColor
    If (.TextColor <> COLORNOTSET) Then accObject.ForeColor = .TextColor: If Err Then Err.Clear: accObject.FontColor = .TextColor: If Err Then Err.Clear: accObject.TextColor = .TextColor
    If (.BorderColor <> COLORNOTSET) Then accObject.BorderColor = .BorderColor
    End With
' ----------------------------------
    ' добавляем контрол в коллекцию уже обработанных для контроля
    If TypeOf accObject Is Access.Control Then Col.Add accObject, accObject.Name
    Result = True
HandleExit:  TagsRead = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_TypeControlExInit() As tControlEx
' значения элементов типа по-умолчанию
    With p_TypeControlExInit
        .BackColor = COLORNOTSET
        .TextColor = COLORNOTSET
        .BorderColor = COLORNOTSET
        .PictMode = apObjSizeZoomDown
        .PictAlign = eCenter
        .TextPlace = ePlaceOnRight
    End With
End Function
Public Function TagsSave(objFrom As Object, Optional objTo As Object) As Boolean '(Optional TagTypes As eObjectProps = ltAll) As Boolean
' сохраняет дополнительные свойства в тэг контрола
'-------------------------
' objFrom   - ссылка на объект дополнительные свойства которого сохраняются
' objTo     - ссылка на объект в свойство Tag которого сохраняются дополнительные свойства
' TagTypes  - типы доп свойств которые необходимо сохранить (можно и нужно комбинировать)
'-------------------------
Dim Result As Boolean: Result = False
    On Error Resume Next 'GoTo HandleError
'    With Control
'Dim strTag As String, strVal As String
'' ----------------------------------
'' читаем свойства контрола в дополнительные свойства
'' ----------------------------------
'    If GetFromCtrl Then
'' ----------------------------------
'' стиль/оформление
'' ----------------------------------
'' цветовое оформление (Back/Fore/Font/TextColor)                (adhcColor/adhcBackColor/etc)
'' ----------------------------------
'
'' ----------------------------------
'' поведение
'' ----------------------------------
'    'If (SplitIt <> cdNone) Then      ' инициализация сплиттера
'' ----------------------------------
'' положение
'' ----------------------------------
'    'If () Then      '
'' ----------------------------------
'' изображение/текст
'' ----------------------------------
'    'If () Then      '
'' ----------------------------------
'    End If
''' ----------------------------------
''' Сохраняем их в Tag
''' ----------------------------------
''' ----------------------------------
''' стиль отображения                                             (adhcSizeIt/adhcFloatIt/adhcStyleIt)
''' ----------------------------------
''    If (TagTypes And ltStyle) = ltStyle Then
'''   сохраняем ScaleIt
'        strTag = adhcScaleIt: strVal = Choose(ScaleIt + 2, adhcYes, adhcNo, vbNullString): Call TaggedStringSet(.Tag, strTag, strVal)
'''   сохраняем StyleIt как SizeIt/FloatIt
'''        If SizeIt <> czNone Then Call TaggedStringSet(.Tag, adhcSizeIt, Choose(SizeIt, adhcSizeRight, adhcSizeBottom, adhcSizeBoth))
'''        If FloatIt <> cfNone Then Call TaggedStringSet(.Tag, adhcFloatIt, Choose(SizeIt, adhcFloatRight, adhcFloatBottom, adhcFloatBoth))
'''   сохраняем StyleIt
''    ' если значение свойства размера меньше порога, он задано как отностельное - устанавливаем соответствующий стиль
'        StyleIt = SetBits(StyleIt, lsXProp, Abs(DefLeft) <= twMinLim)
'        StyleIt = SetBits(StyleIt, lsНProp, Abs(DefTop) <= twMinLim)
'        StyleIt = p_SetBits(StyleIt, lsRProp, Abs(DefRight) <= twMinLim)
'        StyleIt = p_SetBits(StyleIt, lsBProp, Abs(DefBottom) <= twMinLim)
'        StyleIt = p_SetBits(StyleIt, lsWProp, Abs(DefWidth) <= twMinLim)
'        StyleIt = p_SetBits(StyleIt, lsHProp, Abs(DefHeight) <= twMinLim)
'        strTag = adhcStyleIt: strVal = GetStyleText(StyleIt): Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''' ----------------------------------
''' оформление
''' ----------------------------------
''' цветовое оформление (Back/Fore/Font/TextColor)                (adhcColor/adhcBackColor/etc)
''    If (TagTypes And ltColors) = ltColors Then
'        strTag = adhcBackColor: strVal = GetColorText(BackColor, False): Call TaggedStringSet(.Tag, strTag, strVal)
'        strTag = adhcForeColor: strVal = GetColorText(ForeColor, False): Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''' ----------------------------------
''' поведение
''' ----------------------------------
''' признак разделителя (сплиттера)
''    If (TagTypes And ltSplit) = ltSplit Then
'        strTag = adhcSplitIt: strVal = GetSplitText(SplitIt): Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''' заданное для контрола действие
''    If (TagTypes And ltAction) = ltAction Then
'        strTag = adhcAction: strVal = Trim(Action): Call TaggedStringSet(.Tag, strTag, strVal) ', adhcSplitBoth))
''    End If
''' ----------------------------------
''' положение
''' ----------------------------------
'Dim BondName As String
''    If (TagTypes And ltLeft) = ltLeft Then
'        strTag = adhcDefLeft1: If DefLeft = 0 Then strVal = vbNullString Else strVal = GetSizeText(DefLeft, adhcCm, p_GetBoundInTwips(eLeft, BondName), ((StyleIt And lsXProp) = lsXProp))
'        Call TaggedStringSet(.Tag, strTag, strVal): If Len(strVal) > 0 Then strVal = BondName
'        strTag = adhcBondLeft1: Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltTop) = ltTop Then
'        strTag = adhcDefTop1: If DefTop = 0 Then strVal = vbNullString Else strVal = GetSizeText(DefTop, adhcCm, p_GetBoundInTwips(eTop, BondName), ((StyleIt And lsYProp) = lsYProp))
'        Call TaggedStringSet(.Tag, strTag, strVal): If Len(strVal) > 0 Then strVal = BondName
'        strTag = adhcBondTop1: Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltRight) = ltRight Then
'        strTag = adhcDefRight1: If DefRight = 0 Then strVal = vbNullString Else strVal = GetSizeText(DefRight, adhcCm, p_GetBoundInTwips(eRight, BondName), ((StyleIt And lsRProp) = lsRProp))
'        Call TaggedStringSet(.Tag, strTag, strVal): If Len(strVal) > 0 Then strVal = BondName
'        strTag = adhcBondRight1: Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltBottom) = ltBottom Then
'        strTag = adhcDefBottom1: If DefBottom = 0 Then strVal = vbNullString Else strVal = GetSizeText(DefBottom, adhcCm, p_GetBoundInTwips(eBottom, BondName), ((StyleIt And lsBProp) = lsBProp))
'        Call TaggedStringSet(.Tag, strTag, strVal): If Len(strVal) > 0 Then strVal = BondName
'        strTag = adhcBondBottom1: Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltWidth) = ltWidth Then
'        strTag = adhcDefWidth1: If DefWidth = 0 Then strVal = vbNullString Else strVal = GetSizeText(DefWidth, adhcCm, p_GetBoundInTwips(eWidth, BondName), ((StyleIt And lsWProp) = lsWProp))
'        Call TaggedStringSet(.Tag, strTag, strVal): If Len(strVal) > 0 Then strVal = BondName
'        strTag = adhcBondWidth1: Call TaggedStringSet(.Tag, strTag, strVal)
'
'        strTag = adhcMinWidth1: If MinWidth = 0 Then strVal = vbNullString Else strVal = GetSizeText(MinWidth, adhcCm, p_GetBoundInTwips(eWidth), ((StyleIt And lsWProp) = lsWProp))
'        Call TaggedStringSet(.Tag, strTag, strVal)
'
'        strTag = adhcMaxWidth1: If MaxWidth = 0 Then strVal = vbNullString Else strVal = GetSizeText(MaxWidth, adhcCm, p_GetBoundInTwips(eWidth), ((StyleIt And lsWProp) = lsWProp))
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltHeight) = ltHeight Then
'        strTag = adhcDefHeight1: If DefHeight = 0 Then strVal = vbNullString Else strVal = GetSizeText(DefHeight, adhcCm, p_GetBoundInTwips(eHeight, BondName), ((StyleIt And lsHProp) = lsHProp))
'        Call TaggedStringSet(.Tag, strTag, strVal): If Len(strVal) > 0 Then strVal = BondName
'        strTag = adhcBondHeight1: Call TaggedStringSet(.Tag, strTag, strVal)
'
'        strTag = adhcMinHeight1: If MinHeight = 0 Then strVal = vbNullString Else strVal = GetSizeText(MinHeight, adhcCm, p_GetBoundInTwips(eHeight), ((StyleIt And lsHProp) = lsHProp))
'        Call TaggedStringSet(.Tag, strTag, strVal)
'
'        strTag = adhcMaxHeight1: If MaxHeight = 0 Then strVal = vbNullString Else strVal = GetSizeText(MaxHeight, adhcCm, p_GetBoundInTwips(eHeight), ((StyleIt And lsHProp) = lsHProp))
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''' ----------------------------------
''' свойства выводимого изображения
''' ----------------------------------
''    If (TagTypes And ltPictName) = ltPictName Then
'        strTag = adhcPictName: strVal = Trim(ObjName): Call TaggedStringSet(.Tag, strTag, strVal) ' имя/код выводимого объекта
''    End If
'        If Len(strVal) > 0 Then
''    If (TagTypes And ltPictSize) = ltPictSize Then
'        strTag = adhcPictSize: If ObjSize > 0 Then strVal = ObjSize Else strVal = vbNullString    ' размер выводимого объекта в пикселях
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltPictMode) = ltPictMode Then
'        strTag = adhcPictMode: strVal = GetPictModeText(ObjMode)                                  ' режим масштабирования объекта
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltPictPlace) = ltPictPlace Then
'        strTag = adhcPictAlign: strVal = GetPositionText(ObjAlign)                                   ' режим выравнивания объекта
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltPictAngle) = ltPictAngle Then
'        strTag = adhcPictAngle: If ObjAngle = 0 Then strVal = vbNullString Else strVal = ObjAngle ' угол поворота изображения
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltPictGray) = ltPictGray Then
'        strTag = adhcPictGray: If Not ObjGray Then strVal = vbNullString Else strVal = GetBoolText(ObjGray, 1)   ' серое/цветное изображение
'        Call TaggedStringSet(.Tag, strTag, strVal)
'        End If
''    End If
''' ----------------------------------
''' управление текстом на контроле
''' ----------------------------------
''    If (TagTypes And ltPictText) = ltPictText Then
'        strTag = adhcText: strVal = Trim(ObjText): Call TaggedStringSet(.Tag, strTag, strVal) ' текст выводимый вместе с изображением
''    End If
'        If Len(strVal) > 0 Then
''    If (TagTypes And ltTextPlace) = ltTextPlace Then
'        strTag = adhcTextPlace: strVal = GetPlaceText(TxtPlace)                               ' размещение текста относительно изображения
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltTextAlign) = ltTextAlign Then
'        strTag = adhcTextAlign: strVal = GetPositionText(TxtAlign)                               ' выравнивание текста
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltTextAngle) = ltTextAngle Then
'        strTag = adhcTextAngle: If TxtAngle = 0 Then strVal = vbNullString Else strVal = TxtAngle ' угол наклона текста
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltFontName) = ltFontName Then
'        strTag = adhcDefFontName: strVal = Trim(DefFontName)                                        ' шрифт текста
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltFontSize) = ltFontSize Then
'        strTag = adhcDefFontSize: If DefFontSize > 0 Then strVal = DefFontSize Else strVal = vbNullString ' размер шрифта
'        Call TaggedStringSet(.Tag, strTag, strVal)
'        End If
''    End If
''' ----------------------------------
'    End With
    Err.Clear: Result = True
HandleExit:  TagsSave = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function GetUnitByName(Unit As String, Optional UnitName As String) As Single
' возвращает множитель для единицы измерения размера
    Select Case Unit
    Case adhcCm, adhcCm1:       GetUnitByName = cm:     UnitName = adhcCm
    Case adhcInch, adhcInch1:   GetUnitByName = inch:   UnitName = adhcInch
    Case adhcPoints:            GetUnitByName = pt:     UnitName = adhcPoints
    Case adhcPixels:            GetUnitByName = px:     UnitName = adhcPixels
    'Case adhcTwips:             GetUnitByName = 1:      UnitName = vbNullString
    Case Else:                  GetUnitByName = 1:      UnitName = vbNullString
    End Select
End Function
Public Function GetSizeText(twSize, Optional Unit As String, _
        Optional twBase As Long, Optional Prop As Boolean = False) As String
' переводит размеры заданные числом в строку
Const cDecimails = 3 ' количество знаков после запятой в результате
' twSize - размер/позиция (twip) или пропорциональная величина
' Unit - текстовое обозначение единиц измерения результата
' twBase - размер/позиция (twip) базы для относительных величин
' Prop - признак необходимости вернуть результат ввиде пропорциональной величины к twBase
Dim Result As Single
    If Abs(twSize) <= twMinLim Then
' передан относительный размер (доли), требуется вывод относительного (пропорционального) размера
        Result = twSize: Unit = vbNullString            ' оставляем как есть
    ElseIf Prop And twBase <> 0 Then
' передан абсолютный размер (twip), требуется вывод относительного (пропорционального) размера
        Result = twSize / twBase: Unit = vbNullString   ' возвращаем отношение переданного к базе
    Else
' передан абсолютный размер (twip), требуется вывод размера в абсолютных величинах (Unit)
Dim sngUnit As Single: sngUnit = GetUnitByName(Unit, Unit)
        Result = (twSize - twBase) / sngUnit              ' возвращаем положение переданного относительно базы в указанных единицах
    End If
' форматируем результат
    Result = Round(Result, cDecimails)
    GetSizeText = Format(Result, "#0" & IIf(Result <> Fix(Result), "." & String$(cDecimails, "#"), vbNullString))
    If Len(Unit) > 0 Then GetSizeText = GetSizeText & Unit  ' добавляем еденицу измерения
End Function
Public Function GetSizeFromText(SizeText As String, Optional strBase As String) As Single
' переводит размеры заданные строкой в целое число (твипы), либо дробь (относительную величину)
' если задано значение базы - целое определяет как смещение, дробное как отношение к базе и переводит в абсолютное (твипы)
' twMinLim - порог ниже которого значения считаются отношением
' SizeText - строка размера/позиции (может использовать единицы измерения)
' strBase - строка размера/позиции относительно которых задано значение
'-------------------------
' ! десятичный разделитель в соответствии с региональными настройками (,)
'-------------------------
' IsNumeric проверяет число в соответствии с региональными настройками
' cDbl преобразует в соответствии с региональными настройками
' Val игнорирует региональные настройки
Dim sngSize As Single
Dim twBase As Long: If Len(strBase) > 0 Then twBase = GetSizeFromText(strBase)
    If IsNumeric(SizeText) Then
' число (без единиц измерения)
        sngSize = CSng(SizeText)
    ' число - относительная величина
        If Abs(sngSize) <= twMinLim Then If twBase > 0 Then GetSizeFromText = sngSize * twBase: Exit Function
    ' число - абсолютная величина заданная в twip
        GetSizeFromText = twBase + sngSize: Exit Function
    End If
' число - абсолютная величина в единицах измерения
' выполняем преобразование по заданным правилам
Dim arrRules: arrRules = Array("см", "cm", "in", "'", "pt", "px", "tw")
Dim Pos As Long, Sym As String, tmp As String
Dim i As Long
    For i = LBound(arrRules) To UBound(arrRules)
        Pos = InStrRev(SizeText, arrRules(i))
        If Pos > 1 Then
            tmp = Left$(SizeText, Pos - 1)
            If IsNumeric(tmp) Then
                sngSize = CSng(tmp) * GetUnitByName(CStr(arrRules(i)))
                GetSizeFromText = twBase + sngSize: Exit Function
            End If
        End If
    Next i
End Function
Public Function GetColorText(vbСolor As Long, Optional SchemeColorsOnly As Boolean = False) As String
    On Error GoTo HandleError
    GetColorText = vbNullString
' цвет задан именем
    Select Case vbСolor
' условные названия цветов (настраиваются в цветовой схеме)
' например Цвет1, Цвет2, Цвет3
    Case appColorDark:      GetColorText = StrConv(adhcColorDark, vbProperCase)  ' темный цвет
    Case appColorDark2:     GetColorText = StrConv(adhcColorDark2, vbProperCase)
    Case appColorDark3:     GetColorText = StrConv(adhcColorDark3, vbProperCase)
    Case appColorBright:    GetColorText = StrConv(adhcColorBright, vbProperCase) ' яркий цвет
    Case appColorBright2:   GetColorText = StrConv(adhcColorBright2, vbProperCase)
    Case appColorBright3:   GetColorText = StrConv(adhcColorBright3, vbProperCase)
    Case appColorLight:     GetColorText = StrConv(adhcColorLight, vbProperCase)  ' светлый цвет
    Case appColorLight2:    GetColorText = StrConv(adhcColorLight2, vbProperCase)
    Case appColorLight3:    GetColorText = StrConv(adhcColorLight3, vbProperCase)
    End Select
    If SchemeColorsOnly Then Err.Raise vbObjectError + 512
    Select Case vbСolor
' цвета заданные именем (основные)
    Case vbBlack:           GetColorText = adhcColorBlack   ' черный (&H0)
    Case vbWhite:           GetColorText = adhcColorWhite   ' белый (&HFFFFFF)
    Case &H808080:          GetColorText = adhcColorGray    ' серый
    Case &H333333:          GetColorText = adhcColorDark & adhcColorGray     ' темно серый
    Case &HC0C0C0:          GetColorText = "Silver"         ' серебряный
    Case vbRed:             GetColorText = "Red"            ' красный (&HFF)
    Case vbBlue:            GetColorText = "Blue"           ' синий (&HFF0000)
    Case &H8000:            GetColorText = "Green"          ' зеленый
    Case vbYellow:          GetColorText = "Yellow"         ' желтый (&HFFFF)
    Case vbMagenta:         GetColorText = "Magenta"        ' розовый/фуксия (&HFF00FF)
    Case &H80:              GetColorText = "Navy"           ' глубокий синий
    Case &H808000:          GetColorText = "Teal"           ' сине-зелёный
    Case vbCyan:            GetColorText = "Cyan"           ' зелено-голубой/аквамарин (&HFFFF00)
    Case &HFF00:            GetColorText = "Lime"           ' желто-зеленый (лимонный)
    Case &H8080:            GetColorText = "Olive"          ' тёмный желтовато-зелёный (оливковый)
    Case &H800000:          GetColorText = "Maroon"         ' тёмно-бордовый (каштан)
    Case &H800080:          GetColorText = "Purple"         ' пурпурный
    End Select
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function GetColorFromText(ColorName)  'As Long
    On Error GoTo HandleError
    GetColorFromText = Null
' цвет задан числом
    If IsNumeric(ColorName) Then GetColorFromText = CLng(ColorName): GoTo HandleExit
' цвет задан именем
    Select Case ColorName
' условные названия цветов (настраиваются в цветовой схеме)
' например Цвет1, Цвет2, Цвет3
    Case adhcColorDark:     GetColorFromText = appColorDark   ' темный цвет
    Case adhcColorDark2:    GetColorFromText = appColorDark2
    Case adhcColorDark3:    GetColorFromText = appColorDark3
    Case adhcColorBright:   GetColorFromText = appColorBright ' яркий цвет
    Case adhcColorBright2:  GetColorFromText = appColorBright2
    Case adhcColorBright3:  GetColorFromText = appColorBright3
    Case adhcColorLight:    GetColorFromText = appColorLight  ' светлый цвет
    Case adhcColorLight2:   GetColorFromText = appColorLight2
    Case adhcColorLight3:   GetColorFromText = appColorLight3
' Цвета заданные именем HTML: https://colorscheme.ru/html-colors.html?ysclid=lrrntfgkb9889651755
' Основные цвета:
    ' HTML Color Name       ' VB Color (BGR)                        ' HTML Color (RGB)
    Case "Black":           GetColorFromText = vbBlack              ' #000000   ' чёрный
    Case "White":           GetColorFromText = vbWhite              ' #FFFFFF   ' белый
    Case "Gray", "Grey":    GetColorFromText = rgb(128, 128, 128)   ' #808080   ' серый
    Case "Silver":          GetColorFromText = rgb(192, 192, 192)   ' #C0C0C0   ' серебряный
    Case "Fuchsia", "Magenta": GetColorFromText = vbMagenta         ' #FF00FF   ' розовый/фуксия
    Case "Purple":          GetColorFromText = rgb(128, 0, 128)     ' #800080   ' пурпурный
    Case "Red":             GetColorFromText = vbRed                ' #FF0000   ' красный
    Case "Maroon":          GetColorFromText = rgb(128, 0, 0)       ' #800000   ' тёмно-бордовый (каштан)
    Case "Yellow":          GetColorFromText = rgb(255, 255, 0)     ' #FFFF00   ' жёлтый
    Case "Olive":           GetColorFromText = rgb(128, 128, 0)     ' #808000   ' тёмный желтовато-зелёный (оливковый)
    Case "Lime":            GetColorFromText = rgb(0, 255, 0)       ' #00FF00   ' желто-зеленый (лимонный)
    Case "Green":           GetColorFromText = rgb(0, 128, 0)       ' #008000   ' зеленый
    Case "Aqua", "Cyan":    GetColorFromText = rgb(0, 255, 255)     ' #00FFFF   ' зелено-голубой/аквамарин
    Case "Teal":            GetColorFromText = rgb(0, 128, 128)     ' #008080   ' сине-зелёный
    Case "Blue":            GetColorFromText = vbBlue               ' #0000FF   ' синий
    Case "Navy":            GetColorFromText = rgb(0, 0, 128)       ' #000080   ' глубокий синий
' Красные тона:
    Case "IndianRed":       GetColorFromText = rgb(205, 92, 92)     ' #CD5C5C
    Case "LightCoral":      GetColorFromText = rgb(240, 128, 128)   ' #F08080
    Case "Salmon":          GetColorFromText = rgb(250, 128, 114)   ' #FA8072
    Case "DarkSalmon":      GetColorFromText = rgb(233, 150, 122)   ' #E9967A
    Case "LightSalmon":     GetColorFromText = rgb(255, 160, 122)   ' #FFA07A
    Case "Crimson":         GetColorFromText = rgb(220, 20, 60)     ' #DC143C
    'Case "Red":             GetColorFromText = rgb(255, 0, 0)       ' #FF0000
    Case "FireBrick":       GetColorFromText = rgb(178, 34, 34)     ' #B22222
    Case "DarkRed":         GetColorFromText = rgb(139, 0, 0)       ' #8B0000
' Розовые тона:
    Case "Pink":            GetColorFromText = rgb(255, 192, 203)   ' #FFC0CB
    Case "LightPink":       GetColorFromText = rgb(255, 182, 193)   ' #FFB6C1
    Case "HotPink":         GetColorFromText = rgb(255, 105, 180)   ' #FF69B4
    Case "DeepPink":        GetColorFromText = rgb(255, 20, 147)    ' #FF1493
    Case "MediumVioletRed": GetColorFromText = rgb(199, 21, 133)    ' #C71585
    Case "PaleVioletRed":   GetColorFromText = rgb(219, 112, 147)   ' #DB7093
' Оранжевые тона:
    Case "LightSalmon":     GetColorFromText = rgb(255, 160, 122)   ' #FFA07A
    Case "Coral":           GetColorFromText = rgb(255, 127, 80)    ' #FF7F50
    Case "Tomato":          GetColorFromText = rgb(255, 99, 71)     ' #FF6347
    Case "OrangeRed":       GetColorFromText = rgb(255, 69, 0)      ' #FF4500
    Case "DarkOrange":      GetColorFromText = rgb(255, 140, 0)     ' #FF8C00
    Case "Orange":          GetColorFromText = rgb(255, 165, 0)     ' #FFA500
' Жёлтые тона:
    Case "Gold":            GetColorFromText = rgb(255, 215, 0)     ' #FFD700
    'Case "Yellow":          GetColorFromText = rgb(255, 255, 0)     ' #FFFF00
    Case "LightYellow":     GetColorFromText = rgb(255, 255, 224)   ' #FFFFE0
    Case "LemonChiffon":    GetColorFromText = rgb(255, 250, 205)   ' #FFFACD
    Case "LightGoldenrodYellow":  GetColorFromText = rgb(250, 250, 210) ' #FAFAD2
    Case "PapayaWhip":      GetColorFromText = rgb(255, 239, 213)   ' #FFEFD5
    Case "Moccasin":        GetColorFromText = rgb(255, 228, 181)   ' #FFE4B5
    Case "PeachPuff":       GetColorFromText = rgb(255, 218, 185)   ' #FFDAB9
    Case "PaleGoldenrod":   GetColorFromText = rgb(238, 232, 170)   ' #EEE8AA
    Case "Khaki":           GetColorFromText = rgb(240, 230, 140)   ' #F0E68C
    Case "DarkKhaki":       GetColorFromText = rgb(189, 183, 107)   ' #BDB76B
' Фиолетовые тона:
    Case "Lavender":        GetColorFromText = rgb(230, 230, 250)   ' #E6E6FA
    Case "Thistle":         GetColorFromText = rgb(216, 191, 216)   ' #D8BFD8
    Case "Plum":            GetColorFromText = rgb(221, 160, 221)   ' #DDA0DD
    Case "Violet":          GetColorFromText = rgb(238, 130, 238)   ' #EE82EE
    Case "Orchid":          GetColorFromText = rgb(218, 112, 214)   ' #DA70D6
    'Case "Fuchsia", "Magenta": GetColorFromText = rgb(255, 0, 255)  ' #FF00FF
    Case "MediumOrchid":    GetColorFromText = rgb(186, 85, 211)    ' #BA55D3
    Case "MediumPurple":    GetColorFromText = rgb(147, 112, 219)   ' #9370DB
    Case "BlueViolet":      GetColorFromText = rgb(138, 43, 226)    ' #8A2BE2
    Case "DarkViolet":      GetColorFromText = rgb(148, 0, 211)     ' #9400D3
    Case "DarkOrchid":      GetColorFromText = rgb(153, 50, 204)    ' #9932CC
    Case "DarkMagenta":     GetColorFromText = rgb(139, 0, 139)     ' #8B008B
    Case "Purple":          GetColorFromText = rgb(128, 0, 128)     ' #800080
    Case "Indigo":          GetColorFromText = rgb(75, 0, 130)      ' #4B0082
    Case "SlateBlue":       GetColorFromText = rgb(106, 90, 205)    ' #6A5ACD
    Case "DarkSlateBlue":   GetColorFromText = rgb(72, 61, 139)     ' #483D8B
' Коричневые тона:
    Case "Cornsilk":        GetColorFromText = rgb(255, 248, 220)   ' #FFF8DC
    Case "BlanchedAlmond":  GetColorFromText = rgb(255, 235, 205)   ' #FFEBCD
    Case "Bisque":          GetColorFromText = rgb(255, 228, 196)   ' #FFE4C4
    Case "NavajoWhite":     GetColorFromText = rgb(255, 222, 173)   ' #FFDEAD
    Case "Wheat":           GetColorFromText = rgb(245, 222, 179)   ' #F5DEB3
    Case "BurlyWood":       GetColorFromText = rgb(222, 184, 135)   ' #DEB887
    Case "Tan":             GetColorFromText = rgb(210, 180, 140)   ' #D2B48C
    Case "RosyBrown":       GetColorFromText = rgb(188, 143, 143)   ' #BC8F8F
    Case "SandyBrown":      GetColorFromText = rgb(244, 164, 96)    ' #F4A460
    Case "Goldenrod":       GetColorFromText = rgb(218, 165, 32)    ' #DAA520
    Case "DarkGoldenRod":   GetColorFromText = rgb(184, 134, 11)    ' #B8860B
    Case "Peru":            GetColorFromText = rgb(205, 133, 63)    ' #CD853F
    Case "Chocolate":       GetColorFromText = rgb(210, 105, 30)    ' #D2691E
    Case "SaddleBrown":     GetColorFromText = rgb(139, 69, 19)     ' #8B4513
    Case "Sienna":          GetColorFromText = rgb(160, 82, 45)     ' #A0522D
    Case "Brown":           GetColorFromText = rgb(165, 42, 42)     ' #A52A2A
    'Case "Maroon":          GetColorFromText = rgb(128, 0, 0)       ' #800000
' Зелёные тона:
    Case "GreenYellow":     GetColorFromText = rgb(173, 255, 47)    ' #ADFF2F
    Case "Chartreuse":      GetColorFromText = rgb(127, 255, 0)     ' #7FFF00
    Case "LawnGreen":       GetColorFromText = rgb(124, 252, 0)     ' #7CFC00
    'Case "Lime":            GetColorFromText = rgb(0, 255, 0)       ' #00FF00
    Case "LimeGreen":       GetColorFromText = rgb(50, 205, 50)     ' #32CD32
    Case "PaleGreen":       GetColorFromText = rgb(152, 251, 152)   ' #98FB98
    Case "LightGreen":      GetColorFromText = rgb(144, 238, 144)   ' #90EE90
    Case "MediumSpringGreen": GetColorFromText = rgb(0, 250, 154)   ' #00FA9A
    Case "SpringGreen":     GetColorFromText = rgb(0, 255, 127)     ' #00FF7F
    Case "MediumSeaGreen":  GetColorFromText = rgb(60, 179, 113)    ' #3CB371
    Case "SeaGreen":        GetColorFromText = rgb(46, 139, 87)     ' #2E8B57
    Case "ForestGreen":     GetColorFromText = rgb(34, 139, 34)     ' #228B22
    'Case "Green":           GetColorFromText = rgb(0, 128, 0)       ' #008000
    Case "DarkGreen":       GetColorFromText = rgb(0, 100, 0)       ' #006400
    Case "YellowGreen":     GetColorFromText = rgb(154, 205, 50)    ' #9ACD32
    Case "OliveDrab":       GetColorFromText = rgb(107, 142, 35)    ' #6B8E23
    'Case "Olive":           GetColorFromText = rgb(128, 128, 0)     ' #808000
    Case "DarkOliveGreen":  GetColorFromText = rgb(85, 107, 47)     ' #556B2F
    Case "MediumAquamarine": GetColorFromText = rgb(102, 205, 170)  ' #66CDAA
    Case "DarkSeaGreen":    GetColorFromText = rgb(143, 188, 143)   ' #8FBC8F
    Case "LightSeaGreen":   GetColorFromText = rgb(32, 178, 170)    ' #20B2AA
    Case "DarkCyan":        GetColorFromText = rgb(0, 139, 139)     ' #008B8B
    'Case "Teal":            GetColorFromText = rgb(0, 128, 128)     ' #008080
' Синие тона:
    'Case "Aqua", "Cyan":    GetColorFromText = rgb(0, 255, 255)     ' #00FFFF
    Case "LightCyan":       GetColorFromText = rgb(224, 255, 255)   ' #E0FFFF
    Case "PaleTurquoise":   GetColorFromText = rgb(175, 238, 238)   ' #AFEEEE
    Case "Aquamarine":      GetColorFromText = rgb(127, 255, 212)   ' #7FFFD4
    Case "Turquoise":       GetColorFromText = rgb(64, 224, 208)    ' #40E0D0
    Case "MediumTurquoise": GetColorFromText = rgb(72, 209, 204)    ' #48D1CC
    Case "DarkTurquoise":   GetColorFromText = rgb(0, 206, 209)     ' #00CED1
    Case "CadetBlue":       GetColorFromText = rgb(95, 158, 160)    ' #5F9EA0
    Case "SteelBlue":       GetColorFromText = rgb(70, 130, 180)    ' #4682B4
    Case "LightSteelBlue":  GetColorFromText = rgb(176, 196, 222)   ' #B0C4DE
    Case "PowderBlue":      GetColorFromText = rgb(176, 224, 230)   ' #B0E0E6
    Case "LightBlue":       GetColorFromText = rgb(173, 216, 230)   ' #ADD8E6
    Case "SkyBlue":         GetColorFromText = rgb(135, 206, 235)   ' #87CEEB
    Case "LightSkyBlue":     GetColorFromText = rgb(135, 206, 250)  ' #87CEFA
    Case "DeepSkyBlue":     GetColorFromText = rgb(0, 191, 255)     ' #00BFFF
    Case "DodgerBlue":      GetColorFromText = rgb(30, 144, 255)    ' #1E90FF
    Case "CornflowerBlue":  GetColorFromText = rgb(100, 149, 237)   ' #6495ED
    Case "MediumSlateBlue": GetColorFromText = rgb(123, 104, 238)   ' #7B68EE
    Case "RoyalBlue":       GetColorFromText = rgb(65, 105, 225)    ' #4169E1
    'Case "Blue":            GetColorFromText = rgb(0, 0, 255)       ' #0000FF
    Case "MediumBlue":      GetColorFromText = rgb(0, 0, 205)       ' #0000CD
    Case "DarkBlue":        GetColorFromText = rgb(0, 0, 139)       ' #00008B
    'Case "Navy":            GetColorFromText = rgb(0, 0, 128)       ' #000080
    Case "MidnightBlue":    GetColorFromText = rgb(25, 25, 112)     ' #191970
' Белые тона:
    'Case "White":           GetColorFromText = rgb(255, 255, 255)   ' #FFFFFF
    Case "Snow":            GetColorFromText = rgb(255, 250, 250)   ' #FFFAFA
    Case "Honeydew":        GetColorFromText = rgb(240, 255, 240)   ' #F0FFF0
    Case "MintCream":       GetColorFromText = rgb(245, 255, 250)   ' #F5FFFA
    Case "Azure":           GetColorFromText = rgb(240, 255, 255)   ' #F0FFFF
    Case "AliceBlue":       GetColorFromText = rgb(240, 248, 255)   ' #F0F8FF
    Case "GhostWhite":      GetColorFromText = rgb(248, 248, 255)   ' #F8F8FF
    Case "WhiteSmoke":      GetColorFromText = rgb(245, 245, 245)   ' #F5F5F5
    Case "Seashell":        GetColorFromText = rgb(255, 245, 238)   ' #FFF5EE
    Case "Beige":           GetColorFromText = rgb(245, 245, 220)   ' #F5F5DC
    Case "OldLace":         GetColorFromText = rgb(253, 245, 230)   ' #FDF5E6
    Case "FloralWhite":     GetColorFromText = rgb(255, 250, 240)   ' #FFFAF0
    Case "Ivory":           GetColorFromText = rgb(255, 255, 240)   ' #FFFFF0
    Case "AntiqueWhite":    GetColorFromText = rgb(250, 235, 215)   ' #FAEBD7
    Case "Linen":           GetColorFromText = rgb(250, 240, 230)   ' #FAF0E6
    Case "LavenderBlush":   GetColorFromText = rgb(255, 240, 245)   ' #FFF0F5
    Case "MistyRose":       GetColorFromText = rgb(255, 228, 225)   ' #FFE4E1
' Серые тона:
    Case "Gainsboro":       GetColorFromText = rgb(220, 220, 220)   ' #DCDCDC
    Case "LightGrey", "LightGray": GetColorFromText = rgb(211, 211, 211) ' #D3D3D3
    'Case "Silver":          GetColorFromText = rgb(192, 192, 192)   ' #C0C0C0
    Case "DarkGray", "DarkGrey": GetColorFromText = rgb(169, 169, 169) ' #A9A9A9
    'Case "Gray", "Grey":         GetColorFromText = rgb(128, 128, 128) ' #808080
    Case "DimGray", "DimGrey": GetColorFromText = rgb(105, 105, 105) ' #696969
    Case "LightSlateGray", "LightSlateGrey":   GetColorFromText = rgb(119, 136, 153) ' #778899
    Case "SlateGray", "SlateGrey": GetColorFromText = rgb(112, 128, 144) ' #708090
    Case "DarkSlateGray", "DarkSlateGrey": GetColorFromText = rgb(47, 79, 79) ' #2F4F4F
    'Case "Black":           GetColorFromText = rgb(0, 0, 0)         ' #000000
    End Select
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function GetStyleText(lngStyle As eObjectStyle, Optional Short As Boolean = False) As String
' преобразует числовое значение стиля/выравнивания/размещения в строковое
Dim strVal As String, strDelim As String * 1:   strDelim = c_strStyleDelims  ' значения стилей могут объединяться знаками '+&,'
Dim tmp As Long: tmp = lngStyle
' фиксация размеров/положения
    If (tmp And lsFull) = lsFull Then strVal = strVal & IIf(Short, adhcFull1, adhcFull): tmp = tmp And Not lsFull
    If (tmp And lsHorz) = lsHorz Then strVal = strVal & IIf(Short, adhcHorz2, adhcHorz1): tmp = tmp And Not lsHorz
    If (tmp And lsVert) = lsVert Then strVal = strVal & IIf(Short, adhcVert2, adhcVert1): tmp = tmp And Not lsVert
    If (tmp And lsLeft) = lsLeft Then strVal = strVal & IIf(Short, adhcLeft1, adhcLeft): tmp = tmp And Not lsLeft
    If (tmp And lsRight) = lsRight Then strVal = strVal & IIf(Short, adhcRight1, adhcRight): tmp = tmp And Not lsRight
    If (tmp And lsTop) = lsTop Then strVal = strVal & IIf(Short, adhcTop1, adhcTop): tmp = tmp And Not lsTop
    If (tmp And lsBottom) = lsBottom Then strVal = strVal & IIf(Short, adhcBottom1, adhcBottom): tmp = tmp And Not lsBottom
' признаки относительных размеров
    If (tmp And lsWProp) = lsWProp Then strVal = strVal & strDelim & adhcWProp: tmp = tmp And Not lsWProp
    If (tmp And lsHProp) = lsHProp Then strVal = strVal & strDelim & adhcHProp: tmp = tmp And Not lsHProp
    If (tmp And lsXProp) = lsXProp Then strVal = strVal & strDelim & adhcXProp: tmp = tmp And Not lsXProp
    If (tmp And lsYProp) = lsYProp Then strVal = strVal & strDelim & adhcYProp: tmp = tmp And Not lsYProp
    If (tmp And lsRProp) = lsRProp Then strVal = strVal & strDelim & adhcRProp: tmp = tmp And Not lsRProp
    If (tmp And lsBProp) = lsBProp Then strVal = strVal & strDelim & adhcBProp: tmp = tmp And Not lsBProp
' вывод изображения/подписи
    Select Case (tmp And lsShowIconText)
    Case lsShowIconText:    strVal = strVal & strDelim & adhcPictAndText: tmp = tmp And Not lsShowIconText
    Case lsShowIcon:        strVal = strVal & strDelim & adhcPict: tmp = tmp And Not lsShowIcon
    Case lsShowText:        strVal = strVal & strDelim & adhcText: tmp = tmp And Not lsShowText
    End Select
    If Left(strVal, Len(strDelim)) = strDelim Then strVal = Mid$(strVal, Len(strDelim) + 1)
    GetStyleText = strVal
End Function
Public Function GetStyleFromText(strStyle As String) As eObjectStyle
' преобразует строковое значение стиля/выравнивания/размещения в числовое
    ' значения стилей могут объединяться знаками '+&,'
Dim strVal As String, arrVal() As String, j As Long, tmp As eObjectStyle
    GetStyleFromText = lsNone
    Call Tokenize(strStyle, arrVal, c_strStyleDelims)     ' SplitMany(strVal, c_strStyleDelims)
    For j = LBound(arrVal) To UBound(arrVal)
        strVal = UCase(Trim(arrVal(j)))
        tmp = lsNone ' позиция и размер зафиксированы (нет необходимости в использовании)
        Select Case strVal
        Case lsNone, lsLeft, lsRight, lsHorz, lsTop, lsLeftTop, _
            lsRightTop, lsHorzTop, lsBottom, lsLeftBottom, lsRightBottom, _
            lsHorzBottom, lsVert, lsVertLeft, lsVertRight, lsFull, _
            lsShowIcon, lsShowText, lsShowIconText, _
            lsXProp, lsYProp, lsRProp, lsBProp, lsWProp, lsHProp ', _
            'lsCenterHorz, lsCenterVert, lsCenter,
                                          tmp = CLng(strVal)                  ' задано числом
        Case adhcNone:                    tmp = lsNone                        '
' стиль выравнивания позиции (основные)
        Case adhcLeft, adhcLeft1:         tmp = lsLeft                        ' привязан к левой границе области
        Case adhcRight, adhcRight1:       tmp = lsRight                       ' привязан к правой стороне области
        Case adhcTop, adhcTop1:           tmp = lsTop                         ' привязан к верхей границе области
        Case adhcBottom, adhcBottom1:     tmp = lsBottom                      ' привязан к нижней границе области
        Case adhcVert, adhcVert1, adhcVert2: tmp = lsVert                   ' зафиксирован между верхней и нижней границами области (растягивается вертикально)
        Case adhcHorz, adhcHorz1, adhcHorz2: tmp = lsHorz                   ' зафиксирован между правыми и левыми границами области (растягивается горизонтально)
' стиль выравнивания позиции (производные)
        Case adhcLeftTop, adhcLeftTop1:   tmp = lsLeftTop                     ' привязан к левому верхнему углу области
        Case adhcRightTop, adhcRightTop1: tmp = lsRightTop                    ' привязан к правому верхнему углу области
        Case adhcHorTop, adhcHor1Top, adhcHor2Top1: tmp = lsHorzTop      ' привязан к верхней границе и зафиксирован между правыми и левыми границами области (растягивается горизонтально)
        Case adhcLeftBottom, adhcLeftBottom1: tmp = lsLeftBottom              ' привязан к левому нижнему углу области
        Case adhcRightBottom, adhcRightBottom1: tmp = lsRightBottom           ' привязан к правому нижнему углу области
        Case adhcHorBottom, adhcHor1Bottom, adhcHor2Bottom1: tmp = lsHorzBottom ' привязан к нижней границе и зафиксирован между правыми и левыми границами области (растягивается горизонтально)
        Case adhcVerLeft, adhcVer1Left, adhcVer2Left1:   tmp = lsVertLeft ' привязан к левой границе и зафиксирован между верхней и нижней границами области (растягивается вертикально)
        Case adhcVerRight, adhcVer1Right, adhcVer2Right1: tmp = lsVertRight ' привязан к правой границе и зафиксирован между верхней и нижней границами области (растягивается вертикально)
        Case adhcFull, adhcFull1:         tmp = lsFull                        ' растягивается вертикально и горизонтально по правой и нижней границе области
' стили относительных размеров
        Case adhcXProp:                   tmp = lsXProp                       ' Left зависит от ширины области (распределение по ширине)
        Case adhcYProp:                   tmp = lsYProp                       ' Top зависит от высоты области (распределение по высоте)
        Case adhcRProp:                   tmp = lsRProp                       ' Right зависит от ширины области (распределение по ширине)
        Case adhcBProp:                   tmp = lsBProp                       ' Bottom зависит от высоты области (распределение по высоте)
        Case adhcWProp:                   tmp = lsWProp                       ' Width пропорциональна ширине области
        Case adhcHProp:                   tmp = lsHProp                       ' Height пропорциональна высоте области
' стили вывода изображения/текста
        Case adhcPict:                    tmp = lsShowIcon                    ' выводить иконку
        Case adhcPictText:                    tmp = lsShowText                    ' выводить подпись
        Case adhcPictAndText:             tmp = lsShowIconText                ' выводить иконку и подпись
        Case Else 'Exit for' если не предусмотренное значение сбрасываем стиль и выходим
        End Select
        GetStyleFromText = GetStyleFromText Or tmp ' суммируем стили
    Next j
End Function
Public Function GetSplitFromText(strSplit As String) As eControlSplit
    Select Case strSplit
    Case adhcSplitVer, adhcSplitVer1: GetSplitFromText = cdVert
    Case adhcSplitHor, adhcSplitHor1: GetSplitFromText = cdHorz
    'Case adhcSplitBoth: GetSplitFromText = cdBoth
    Case Else: GetSplitFromText = cdNone
    End Select
End Function
Public Function GetSplitText(lngSplit As eControlSplit) As String
    Select Case lngSplit
    Case cdVert: GetSplitText = adhcSplitVer1
    Case cdHorz: GetSplitText = adhcSplitHor1
    'Case cdBoth: GetSplitText = adhcSplitBoth
    End Select
End Function
Public Function GetBoolFromText(strBool As String) As Boolean
    Select Case strBool
    Case adhcTrue, adhcYes, adhcOn: GetBoolFromText = True
    Case Else: GetBoolFromText = False
    End Select
End Function
Public Function GetBoolText(bBool As Boolean, Optional Mode As Long) As String
    If bBool Then
        GetBoolText = Choose(Mode + 1, adhcTrue, adhcYes, adhcOn)
    Else
        GetBoolText = Choose(Mode + 1, adhcFalse, adhcNo, adhcOff)
    End If
End Function
Public Function GetAlignText(lngAlign As eAlignText, Optional Short As Boolean = False) As String
' выравнивания текста
    Select Case lngAlign
    Case TA_RIGHT:      GetAlignText = IIf(Short, adhcRight1, adhcRight)    ' Опорная точка находится на правой кромке рабочего прямоугольника.
    Case TA_CENTER:     GetAlignText = IIf(Short, adhcCenter1, adhcCenter)  ' Опорная точка выравнивается горизонтально по центру рабочего прямоугольника.
    Case TA_BOTTOM:     GetAlignText = IIf(Short, adhcBottom1, adhcBottom)  ' Опорная точка на нижней кромке рабочего прямоугольника.
''    Case TA_BASELINE:   GetAlignText = IIf(Short, , ) ' Опорная точка находится на базовой линии текста.
''    Case TA_RTLREADING: GetAlignText = IIf(Short, , ) ' Редакция Windows на языках Ближнего Востока: Текст размечается для порядка чтения справа налево , в противоположность порядку чтения по умолчанию слева направо. Это применяется только тогда, когда шрифт, выбранный в контекст устройства предназначен или для Еврейского или для Арабского языка.
''   Case TA_NOUPDATECP               ' Текущая позиция не модифицируется после каждого вызова вывода текста.
''   Case TA_UPDATECP                 ' Текущая позиция модифицируется после каждого вызова вывода текста.
''   Case TA_MASK  = (TA_BASELINE + TA_CENTER + TA_UPDATECP)

'    Case TA_LEFT:      GetAlignText = vbNullString ' 0         ' Опорная точка находится на левой кромке рабочего прямоугольника.
'    Case TA_TOP:       GetAlignText = vbNullString ' 0         ' Опорная точка на верхней кромке рабочего прямоугольника.
    Case Else:          GetAlignText = vbNullString '
    End Select
End Function
Public Function GetAlignFromText(strAlign As String) As eAlignText
' выравнивания текста
    Select Case strAlign
    Case adhcRight1, adhcRight:         GetAlignFromText = TA_RIGHT         ' Опорная точка находится на правой кромке рабочего прямоугольника.
    Case adhcCenter1, adhcCenter:       GetAlignFromText = TA_CENTER        ' Опорная точка выравнивается горизонтально по центру рабочего прямоугольника.
    Case adhcBottom1, adhcBottom:       GetAlignFromText = TA_BOTTOM        ' Опорная точка на нижней кромке рабочего прямоугольника.
    Case Else:                          GetAlignFromText = TA_LEFT '
    End Select
End Function
Public Function GetPositionText(lngAlign As eAlign, Optional Short As Boolean = False) As String
    If lngAlign = eAlignUndef Then Exit Function
'    If lngAlign = eAlignLeftTop Then Exit Function
Dim strVal As String
Dim tmp As Long: tmp = lngAlign
' стиль выравнивания позиции
    If (tmp And lsFull) = lsFull Then strVal = strVal & IIf(Short, adhcCenter1, adhcCenter): tmp = tmp And Not lsFull
    If (tmp And lsHorz) = lsHorz Then strVal = strVal & IIf(Short, adhcHorz2, adhcCenterHor1): tmp = tmp And Not lsHorz
    If (tmp And lsVert) = lsVert Then strVal = strVal & IIf(Short, adhcVert2, adhcCenterVer1): tmp = tmp And Not lsVert
    If (tmp And lsLeft) = lsLeft Then strVal = strVal & IIf(Short, adhcLeft1, adhcLeft): tmp = tmp And Not lsLeft
    If (tmp And lsRight) = lsRight Then strVal = strVal & IIf(Short, adhcRight1, adhcRight): tmp = tmp And Not lsRight
    If (tmp And lsTop) = lsTop Then strVal = strVal & IIf(Short, adhcTop1, adhcTop): tmp = tmp And Not lsTop
    If (tmp And lsBottom) = lsBottom Then strVal = strVal & IIf(Short, adhcBottom1, adhcBottom): tmp = tmp And Not lsBottom
    GetPositionText = strVal
End Function
Public Function GetPositionFromText(strAlign As String) As eAlign
' стиль выравнивания позиции
    GetPositionFromText = eAlignUndef: If Len(Trim(strAlign)) = 0 Then Exit Function
Dim strVal As String, arrVal() As String, j As Long, tmp As eAlign
    Call Tokenize(strAlign, arrVal, c_strStyleDelims)     ' SplitMany(strVal, c_strStyleDelims)
    For j = LBound(arrVal) To UBound(arrVal)
        strVal = UCase(Trim(arrVal(j)))
        tmp = eAlignUndef
        Select Case strVal
        Case eAlignUndef, eLeft, eRight, eTop, eBottom, eCenterHorz, eCenterVert, _
            eAlignLeftTop, eAlignRightTop, eAlignLeftBottom, eAlignRightBottom, _
            eCenterHorzTop, eCenterHorzBottom, eCenterVertLeft, eCenterVertRight, eCenter, eCascade
                                        tmp = CLng(strVal)                  ' задано числом
' стиль выравнивания позиции (основные)
        Case adhcLeft, adhcLeft1:       tmp = eLeft                         ' по левому краю
        Case adhcRight, adhcRight1:     tmp = eRight                        ' по правому краю
        Case adhcTop, adhcTop1:         tmp = eTop                          ' по верхнему краю
        Case adhcBottom, adhcBottom1:   tmp = eBottom                       ' по нижнему краю
        Case adhcHorz, adhcHorz1, adhcHorz2, _
             adhcCenterHor, adhcCenterHor1, adhcCenterHor2: tmp = eCenterHorz        ' центрировать по горизонтали
        Case adhcVert, adhcVert1, adhcVert2, _
             adhcCenterVer, adhcCenterVer1, adhcCenterVer2: tmp = eCenterVert         ' центрировать по вертикали
' стиль выравнивания позиции (производные)
'    ' 2 направления по 3 положения вдоль каждого
'    ' ИТОГО: 3x3 = 9 точек выравнивания.
        Case adhcLeftTop, adhcLeftTop1:   tmp = eAlignLeftTop                 ' по левому верхнему краю
        Case adhcRightTop, adhcRightTop1: tmp = eAlignRightTop                ' по правому верхнему краю
        Case adhcLeftBottom, adhcLeftBottom1: tmp = eAlignLeftBottom          ' по левому нижнему краю
        Case adhcRightBottom, adhcRightBottom1: tmp = eAlignRightBottom       ' по правому нижнему краю
        Case adhcHorTop, adhcHor1Top, _
             adhcCenterHorTop, adhcCenterHor1Top: tmp = eCenterHorzTop                  ' по верхнему краю центрировать по горизонтали
        Case adhcHorBottom, adhcHor1Bottom, adhcHor2Bottom1, _
             adhcCenterHorBottom, adhcCenterHor1Bottom, adhcCenterHor2Bottom1: tmp = eCenterHorzBottom ' по нижнему краю центрировать по горизонтали
        Case adhcVerLeft, adhcVer1Left, adhcVer2Left1, _
             adhcCenterVerLeft, adhcCenterVer1Left, adhcCenterVer2Left1: tmp = eCenterVertLeft ' по левому краю центрировать по вертикали
        Case adhcVerRight, adhcVer1Right, adhcVer2Right1, _
             adhcCenterVerRight, adhcCenterVer1Right, adhcCenterVer2Right1: tmp = eCenterVertRight ' по правому краю центрировать по вертикали
        Case adhcCenter, adhcCenter1, adhcFull, adhcFull1: tmp = eCenter      ' центрировать посередине области
        Case adhcCascade:       tmp = eCascade                      ' композиция (только для формы ??)
        Case eAlignUndef:       tmp = eAlignUndef                   '
        End Select
        GetPositionFromText = GetPositionFromText Or tmp ' суммируем
    Next j
End Function
Public Function GetPlaceText(lngPlace As ePlace) As String
    Select Case lngPlace
    ' внутри по центру
    Case ePlaceCenter: GetPlaceText = adhcPlaceCenter                      ' по центру (внутри)
    Case ePlaceToLeft: GetPlaceText = adhcPlaceToLeft                      ' внутри слева по центру
    Case ePlaceToRight: GetPlaceText = adhcPlaceToRight                    ' внутри справа по центру
    Case ePlaceToTop: GetPlaceText = adhcPlaceToTop                        ' внутри по центру сверху
    Case ePlaceToBottom: GetPlaceText = adhcPlaceToBottom                  ' внутри по центру снизу
    ' снаружи по центру
    Case ePlaceOnLeft: GetPlaceText = adhcPlaceOnLeft                      ' снаружи слева по центру
    Case ePlaceOnRight: GetPlaceText = adhcPlaceOnRight                    ' снаружи справа по центру
    Case ePlaceOnTop: GetPlaceText = adhcPlaceOnTop                        ' снаружи по центру сверху
    Case ePlaceOnBottom: GetPlaceText = adhcPlaceOnBottom                  ' снаружи по центру снизу
    ' внутри по краю
    Case ePlaceToLeftTop: GetPlaceText = adhcPlaceToLeftTop                ' внутри слева сверху
    Case ePlaceToRightTop: GetPlaceText = adhcPlaceToRightTop              ' внутри справа сверху
    Case ePlaceToLeftBottom: GetPlaceText = adhcPlaceToLeftBottom          ' внутри слева снизу
    Case ePlaceToRightBottom: GetPlaceText = adhcPlaceToRightBottom        ' внутри справа снизу
    ' снаружи по краю
    Case ePlaceOnLeftToTop: GetPlaceText = adhcPlaceOnLeftToTop            ' снаружи слева к верхнему краю
    Case ePlaceOnLeftToBottom: GetPlaceText = adhcPlaceOnLeftToBottom      ' снаружи слева к нижнему краю
    Case ePlaceOnRightToTop: GetPlaceText = adhcPlaceOnRightToTop          ' снаружи справа к верхнему краю
    Case ePlaceOnRightToBottom: GetPlaceText = adhcPlaceOnRightToBottom    ' снаружи справа к нижнему краю
    Case ePlaceOnTopToLeft: GetPlaceText = adhcPlaceOnTopToLeft            ' снаружи к левому краю сверху
    Case ePlaceOnTopToRight: GetPlaceText = adhcPlaceOnTopToRight          ' снаружи к правому краю сверху
    Case ePlaceOnBottomToLeft: GetPlaceText = adhcPlaceOnBottomToLeft      ' снаружи к левому краю снизу
    Case ePlaceOnBottomToRight: GetPlaceText = adhcPlaceOnBottomToRight    ' снаружи к правому краю снизу
    ' композиция (только для формы ??)
    Case eCascadeFromLeftTop: GetPlaceText = adhcCascadeFromLeftTop          ' размещение каскадом вправо-вниз
    Case eCascadeFromRightTop: GetPlaceText = adhcCascadeFromRightTop        ' размещение каскадом влево-вниз
    Case eCascadeFromLeftBottom: GetPlaceText = adhcCascadeFromLeftBottom    ' размещение каскадом вправо-вверх
    Case eCascadeFromRightBottom: GetPlaceText = adhcCascadeFromRightBottom  ' размещение каскадом влево-вверх
    Case Else: GetPlaceText = vbNullString 'adhcUndef
'Dim StrVal As String
'    StrVal = GetPositionText(lngPlace And eCenter, Short)
'    StrVal = StrVal & IIf(Short, adhcTo1, adhcTo)
'    StrVal = StrVal & GetPositionText((lngPlace And eCenter * &H10) / &H10, Short)
'    GetPlaceText = StrVal
    End Select
End Function
Public Function GetPlaceFromText(strPlace As String) As ePlace
' стиль размещения одного объекта относительно другого (OnLeftToTop, Center, etc)
    Select Case strPlace
    ' внутри по центру
    Case adhcPlaceCenter: GetPlaceFromText = ePlaceCenter                       ' по центру (внутри)
    Case adhcPlaceToLeft: GetPlaceFromText = ePlaceToLeft                       ' внутри слева по центру
    Case adhcPlaceToRight: GetPlaceFromText = ePlaceToRight                     ' внутри справа по центру
    Case adhcPlaceToTop: GetPlaceFromText = ePlaceToTop                         ' внутри по центру сверху
    Case adhcPlaceToBottom: GetPlaceFromText = ePlaceToBottom                   ' внутри по центру снизу
    ' снаружи по центру
    Case adhcPlaceOnLeft: GetPlaceFromText = ePlaceOnLeft                       ' снаружи слева по центру
    Case adhcPlaceOnRight: GetPlaceFromText = ePlaceOnRight                     ' снаружи справа по центру
    Case adhcPlaceOnTop: GetPlaceFromText = ePlaceOnTop                         ' снаружи по центру сверху
    Case adhcPlaceOnBottom: GetPlaceFromText = ePlaceOnBottom                   ' снаружи по центру снизу
    ' внутри по краю
    Case adhcPlaceToLeftTop: GetPlaceFromText = ePlaceToLeftTop                 ' внутри слева сверху
    Case adhcPlaceToRightTop: GetPlaceFromText = ePlaceToRightTop               ' внутри справа сверху
    Case adhcPlaceToLeftBottom: GetPlaceFromText = ePlaceToLeftBottom           ' внутри слева снизу
    Case adhcPlaceToRightBottom: GetPlaceFromText = ePlaceToRightBottom         ' внутри справа снизу
    ' снаружи по краю
    Case adhcPlaceOnLeftToTop: GetPlaceFromText = ePlaceOnLeftToTop             ' снаружи слева к верхнему краю
    Case adhcPlaceOnLeftToBottom: GetPlaceFromText = ePlaceOnLeftToBottom       ' снаружи слева к нижнему краю
    Case adhcPlaceOnRightToTop: GetPlaceFromText = ePlaceOnRightToTop           ' снаружи справа к верхнему краю
    Case adhcPlaceOnRightToBottom: GetPlaceFromText = ePlaceOnRightToBottom     ' снаружи справа к нижнему краю
    Case adhcPlaceOnTopToLeft: GetPlaceFromText = ePlaceOnTopToLeft             ' снаружи к левому краю сверху
    Case adhcPlaceOnTopToRight: GetPlaceFromText = ePlaceOnTopToRight           ' снаружи к правому краю сверху
    Case adhcPlaceOnBottomToLeft: GetPlaceFromText = ePlaceOnBottomToLeft       ' снаружи к левому краю снизу
    Case adhcPlaceOnBottomToRight: GetPlaceFromText = ePlaceOnBottomToRight     ' снаружи к правому краю снизу
    ' композиция (только для формы ??)
    Case adhcCascadeFromLeftTop: GetPlaceFromText = eCascadeFromLeftTop             ' размещение каскадом вправо-вниз
    Case adhcCascadeFromRightTop: GetPlaceFromText = eCascadeFromRightTop       ' размещение каскадом влево-вниз
    Case adhcCascadeFromLeftBottom: GetPlaceFromText = eCascadeFromLeftBottom   ' размещение каскадом вправо-вверх
    Case adhcCascadeFromRightBottom: GetPlaceFromText = eCascadeFromRightBottom ' размещение каскадом влево-вверх
    Case Else: GetPlaceFromText = ePlaceUndef
    'Dim Result As eAlign: Result = ePlaceUndef
    '    On Error GoTo HandleError
    '' если числовое
    '    If IsNumeric(strPlace) Then Result = CLng(strPlace): GoTo HandleExit
    '' если несовпадает с готовым значением пробуем разобрать как составное
    '' разбить strPlace по adhcTo на 2 части
    'Dim Pos As Long, i As Long, arrDelim()
    '    arrDelim = Array(adhcTo, adhcTo1)
    '    For i = LBound(arrDelim) To UBound(arrDelim)
    '        Pos = InStr(1, strPlace, arrDelim(i)): If Pos > 1 Then Exit For
    '    Next i
    '    If Pos < 1 Then Err.Raise vbObjectError + 512
    '' каждую прочесть через GetPositionFromText и объединить со сдвигом &h10
    '    Result = GetPositionFromText(Left$(strPlace, Pos - 1))                                ' читаем левую часть
    '    Result = Result Or &H10 * GetPositionFromText(Mid$(strPlace, Pos + Len(arrDelim(i))))    ' читаем правую часть
    End Select
'HandleExit:  GetPlaceFromText = Result: Exit Function
'HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function GetPictModeText(ObjMode As eObjSizeMode) As String
    Select Case ObjMode
    Case acOLESizeStretch:      GetPictModeText = adhcStretch             ' сжатие/растяжение (нарушает пропорции)
    Case acOLESizeZoom:         GetPictModeText = adhcZoom                ' пропорциональное масштабирование
    'Case acOLESizeAutoSize:     GetPictModeText = adhcAuto
    Case acOLESizeClip:         GetPictModeText = adhcClip  ' не меняем размер. если размер больше области вывода - обрезка
    'Case apObjSizeZoomDown:     GetPictModeText = adhcDown                ' пропорциональное масштабирование, только уменьшаем
    Case Else:                  GetPictModeText = vbNullString 'adhcClip  ' не меняем размер. если размер больше области вывода - обрезка
    End Select
End Function

Public Function GetPictModeFromText(ObjModeText As String) As eObjSizeMode
    Select Case ObjModeText
    Case adhcStretch:           GetPictModeFromText = acOLESizeStretch    ' сжатие/растяжение (нарушает пропорции)
    Case adhcZoom:              GetPictModeFromText = acOLESizeZoom       ' пропорциональное масштабирование
    'Case adhcAuto, adhcAuto1:   GetPictModeFromText = acOLESizeAutoSize  '
    Case adhcClip:              GetPictModeFromText = acOLESizeClip      ' не меняем размер. если размер больше области вывода - обрезка
    'Case adhcDown, adhcDown1:   GetPictModeFromText = apObjSizeZoomDown   ' пропорциональное масштабирование, только уменьшаем
    Case Else:                  GetPictModeFromText = apObjSizeZoomDown   ' не меняем размер. если размер больше области вывода - обрезка
    End Select
End Function
' --------------------
Private Function CustomPropertyGet( _
    PropName As String, _
    PropValue As Variant, _
    Optional PropObject As Object _
    ) As Boolean
' читает пользовательское свойство произвольного объекта
Const c_strProcedure = "CustomPropertyGet"
' PropName      - имя свойства
' PropValue     - значение свойства
' PropObject    - объект к которому добавляется свойство
Dim prp As Property
Dim Result As Boolean
    Result = False
    On Error GoTo HandleError
    If PropObject Is Nothing Then Set PropObject = CurrentProject ' по-умолчанию
    PropValue = PropObject.Properties(PropName)
    Result = True
HandleExit: CustomPropertyGet = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Function CustomPropertySet( _
    PropName As String, _
    PropValue As Variant, _
    Optional PropObject As Object, _
    Optional PropType As DataTypeEnum = dbText _
    ) As Boolean
' добавляет пользовательское свойство к объекту DAO или AccessObject
Const c_strProcedure = "CustomPropertySet"
' PropName      - имя свойства
' PropValue     - значение свойства
' PropObject    - объект к которому добавляется свойство
' PropType      - тип данных свойства
    On Error Resume Next
    If PropObject Is Nothing Then Set PropObject = CurrentProject ' по-умолчанию
    ' пытаемся записать свойство
    PropObject.Properties(PropName) = PropValue
    If Err.Number = 0 Then GoTo HandleExit
    Err.Clear: On Error GoTo HandleExit
    ' нет такого свойства - добавляем
    If TypeOf PropObject.Properties Is DAO.Properties Then
    ' добавляем DAO свойство
    Dim daoProp As DAO.Property
        Set daoProp = PropObject.CreateProperty(PropName, PropType, PropValue)
        PropObject.Properties.Append daoProp
    ElseIf TypeOf PropObject.Properties Is AccessObjectProperties Then
    ' добавляем AccessObject свойство
        PropObject.Properties.Add PropName, PropValue
    Else
        Err.Raise vbObjectError + 512
    End If
HandleExit: CustomPropertySet = Err.Number = 0: Err.Clear: Exit Function
End Function
Public Function WindowFreeze(hwnd As LongPtr): LockWindowUpdate (hwnd): End Function
Public Function WindowUnFreeze(): LockWindowUpdate (0): End Function
'---------------------
Public Function GetAlignPoint(Alignment As eAlign, _
    cX As Single, cY As Single, Optional Cascade As Boolean)
' возвращает пропорциональные координаты точки привязки в зависимости от заданного режима выравнивания
'---------------------
' Получает:
'   Alignment - режим выравнивания
' Возвращает:
'   cX,cY     - позиция точки привязки после трансформации
'---------------------
    '
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

Private Function p_IsExist(Key As String, Coll As Collection, Optional ByRef Value) As Boolean
' проверяет наличие элемента в коллекции
'-------------------------
    On Error GoTo HandleError
    Value = Coll(Key)
HandleExit:  p_IsExist = True:  Exit Function
HandleError: p_IsExist = False: Err.Clear
End Function


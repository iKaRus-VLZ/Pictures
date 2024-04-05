Attribute VB_Name = "modStrings"
Option Base 0
Option Explicit
#Const APPTYPE = 0  ' 0=ACCESS, 1=EXCEL
'=========================
Private Const c_strModule As String = "modStrings"
'=========================
' Описание      : Функции для работы со строками
' Автор         : Кашкин Р.В. (KashRus@gmail.com)
' Версия        : 1.1.30.453854194
' Дата          : 03.04.2024 10:03:56
' Примечание    : сделано под Access x86, адаптировано под x64, но толком не тестировалось. _
'               : для работы с Excel сделать APPTYPE=1
' v.1.1.30      : 12.03.2024 - изменения в GroupsGet - первая попытка переделать скобки под шаблоны (чтобы получить возможность разбирать двух- и более -звенные выражения вроде If .. Then .. End If)
' v.1.1.29      : 21.12.2022 - изменения в GroupsGet - исправлены многочисленные ошибки. (всё еще сильно экспериментальная)
' v.1.1.27      : 09.08.2022 - изменения в DelimStringSet - добавлен параметр SetUnique для контроля уникальности вставляемых значений
' v.1.1.26      : 11.02.2022 - изменения в TaggedStringSet/Del - исправлены ошибки возникающие если TagDelim текстовое выражение, зависящее от регистра
' v.1.1.25      : 16.07.2020 - переписана PlaceHoldersGet - прошлая версия приводила к неправильному результату если первое совпадение окажется впоследствии забракованным. добавлена поддержка множественных вхождений (ReplaceExisting)
' v.1.1.24      : 24.03.2020 - добавлена функция GroupGet - для извлечения из текста групп (выражений заключенных в скобки)
' v.1.1.23      : 04.02.2020 - изменения в PlaceHoldersGet - расширен синтаксис шаблона за счет выражений VBA.Like и VBS.RegExp
' v.1.1.22      : 03.02.2020 - добавлена PlaceHoldersGet - позволяет извлекать в коллекцию значения переменных из строки по шаблону; _
                               переименованы функции группы работы с подстановочными значениями для большего единообразия
' v.1.1.19      : 30.01.2020 - изменения в PlaceHoldersSet - добавлена возможность использования модификаторов (формат см. p_TermModify)
' v.1.1.17      : 29.11.2019 - добавлены функции для работы с подстановочными переменными PlaceHoldersSetByIndex и PlaceHoldersSet
' v.1.1.16      : 31.10.2019 - изменения в Tokenize - добавлен необязательный параметр Positions - возвращающий массив позиций найденных токенов в исходной строке _
                               и добавлены функции TokenString[Get","Set","Del] извлечения/вставки/удаления токенов из строки, аналогичные имеющимся для DelimString и TaggedString _
                               для всех функций работы с подстроками добавлены параметры sBeg, sEnd возвращающие позицию подстроки в строке
' v.1.1.12      : 24.09.2019 - обновлены функции работы со строками с разделителями DelimString[Get","Set","Del], _
                               добавлена возможность работы с отрицательными позициями (от конца строки) _
                               и добавлены функции работы со строками именных параметров TaggedString[Get","Set","Del]
' v.1.1.10      : 16.08.2019 - добавлены функции фонетического сравнения: SoundEx, PolyPhone _
                               и функции определения фонетического расстояния: Наибольшая общая подпоследовательность, Levenshtein, Dice, коэффициент Jaro etc. _
                               некоторые алгоритмы пытался воспроизвести по текстовым описаниям - возможно реализации не вполне корректны.
' v.1.1.9       : 18.07.2019 - добавлена NumToWords - преобразование числа в текст и склонение результата по падежам
' v.1.1.8       : 13.07.2019 - внесены исправления в функции разбиения строки на части по контролам/шрифту
' v.1.1.5       : 12.12.2018 - добавлена HyphenateWord - для расстановки переносов в словах. Источник: http://www.cyberforum.ru/vba/thread792944.html
' v.1.1.4       : обновлена DeclineWords - добавлена возможность пропуска слов при склонении словосочетаний
' v.1.1.2       : переписан DeclineWord - функция склонения по падежам. Добавлена возможность склонять по числам.
'=========================
' ToDo: DeclineWord - переписать для работы с правилами описанными как шаблоны (продумать вид шаблонов), для лучшей читаимости и удобства настройки
' + добавить возможность эскапирования символов в функциях с разделителями и пропуск разделителей внутри кавычек
' + PlaceHoldersGet - проверять соответствие извлекаемого значения, заданным параметрам переменной, например принадлежность списку допустимых значений
' - NumToWords      - неправильно склоняет знаменатели натуральных дробей >10^6
'=========================
'Private Const c_strEsc = "\" ' эскейп символ - для указания что следующий знак следует рассматривать как обычный символ (не специальный, не разделитель)
' русский алфавит
    ' проверить например глухой звук: If iInStr (c_strSymbRusConsonDeaf,sChar) Then
Private Const c_strSymbRusConson = "йбвгджзклмнпрстфхцчшщ"  ' согласные звуки
'Private Const c_strSymbRusConsonVoicPaired = "бвгджз", c_strSymbRusConsonVoicOnly = "лмнр" ' звонкие парные/непарные
'Private Const c_strSymbRusConsonDeafPaired = "пфктшс", c_strSymbRusConsonDeafOnly = "хцчщ" ' глухие парные/непарные
'Private Const c_strSymbRusConsonVoic = c_strSymbRusConsonVoicPaired & c_strSymbRusConsonVoicOnly ' звонкие все
'Private Const c_strSymbRusConsonDeaf = c_strSymbRusConsonDeafPaired & c_strSymbRusConsonDeafOnly ' глухие все
'Private Const c_strSymbRusConsonHardSoft = "бвгдзклмнпрстфх" ' парные твёрдые/мягкие в зависимости от гласной
'Private Const c_strSymbRusConsonHardOnly = "жшц", c_strSymbRusConsonSoftOnly = "йчщ" ' только твёрдые/мягкие
'Private Const c_strSymbRusConsonHissing = "жшчщ", c_strSymbRusConsonWhistling = "сзц"  ' шипящие/свистящие
'Private Const c_strSymbRusConsonSonar = "йлмнр", c_strSymbRusConsonNoisy = "кпстфхцчшщбвгджз"  ' сонорные/шумные
Private Const c_strSymbRusVowel = "аеёиоуыэюя"   ' гласные звуки
'Private Const c_strSymbRusVowelYotated = "еёюя" ' йотированные (двойные) гласные звуки
'Private Const c_strSymbRusVowelSoft = "еёюяи"   ' смягчающие гласные звуки
'Private Const c_strSymbRusVowelHard = "эоуаы"   ' парные им не являющиеся смягчающими
Private Const c_strSymbRusSign = "ьъ"            ' знаки
'Private Const c_strSymbRusSignSoft = "ь"        ' мягкий знак
'Private Const c_strSymbRusSignHard = "ъ"        ' твёрдый знак
' английский алфавит
Private Const c_strSymbRusAll = c_strSymbRusVowel & c_strSymbRusConson & c_strSymbRusSign
Private Const c_strSymbEngVowel = "aeiouy", c_strSymbEngConson = "bcdfghjklmnpqrstvwxz", c_strSymbEngSign = "" '"'`"
Private Const c_strSymbEngAll = c_strSymbEngVowel & c_strSymbEngConson & c_strSymbEngSign
' цифры и символы
Private Const c_strSymbDigits = "0123456789", c_strSymbMath = "+-*/\^|=", c_strSymbPunct = ".,?!:;-()" ' & "…"
Private Const c_strSymbCommas = "'""", c_strSymbParenth = "()[]{}<>", c_strSymbOthers = "_&@#$%~`"
Private Const c_strSymbSpaces = " " & vbCr & vbLf & vbNewLine & vbTab & vbVerticalTab
' для преобразования имен
Private Const c_strHexPref = "&H"
Private Const c_strOthers = " -~_"

Private Const c_idxPref = "i" ' префикс имен элементов коллекции
' size convertion constants
Private Const PointsPerInch = 72
Private Const TwipsPerInch = 1440
Private Const CentimitersPerInch = 2.54                 '1 дюйм = 127 / 50 см
Private Const HimetricPerInch = 2540                    '1 дюйм = 1000 * 127/50 himetrix
'
Private Const inch = TwipsPerInch                       '1 дюйм = 1440 twips
Private Const pt = TwipsPerInch / PointsPerInch         '1 пункт = 20 twips
Private Const cm = TwipsPerInch / CentimitersPerInch    '1 см = 566.929133858 twips
'--------------------------------------------------------------------------------
Public Enum DeclineCase         ' падеж
    DeclineCaseUndef = 0
    DeclineCaseImen = 1         ' им.п. (кто/что)       Nominative
    DeclineCaseRod = 2          ' р.п.  (кого/чего)     Genitive
    DeclineCaseDat = 3          ' д.п.  (кому/чему)     Dative
    DeclineCaseVin = 4          ' в.п.  (кого/что)      Accusative
    DeclineCaseTvor = 5         ' т.п.  (кем/чем)       Ablative
    DeclineCasePred = 6         ' п.п.  (о ком/о чём)   Prepositional
End Enum
Public Enum DeclineGend         ' род ("м|ж|ср")
    DeclineGendUndef = 0
    DeclineGendMale = 1         ' м.р.
    DeclineGendFem = 2          ' ж.р.
    DeclineGendNeut = 3         ' с.р.
End Enum
Public Enum DeclineNumb         ' число ("ед|мн")
    DeclineNumbUndef = 0
    DeclineNumbSingle = 1       ' ед.ч.
    DeclineNumbPlural = 2       ' мн.ч.
End Enum
Public Enum SpeechPartType      ' часть речи
    SpeechPartTypeUndef = 0
    SpeechPartTypeNoun = 1      ' существительное
    SpeechPartTypeAdject = 2    ' прилагательное
    SpeechPartTypeNumeral = 3   ' числительное
    SpeechPartTypeVerb = 4      ' глагол
    SpeechPartTypeAdverb = 5    ' наречие
    SpeechPartTypePronoun = 6   ' местоимение
    SpeechPartTypePreposition = 7 ' предлог
End Enum
Public Enum NumeralType         ' тип числительных ("количественное|порядковое")
    NumeralUndef = 0
    NumeralOrdinal = 1          ' количественное
    NumeralCardinal = 2         ' порядковое
End Enum
Public Enum SymbolType          ' тип символа
    SymbolTypeUndef = 0
    SymbolTypeVowel = 1         ' гласные
    SymbolTypeCons = 2          ' согласные
    SymbolTypeSign = 3          ' знаки алфавита
    SymbolTypeNumb = 4          ' цифры
End Enum
Public Enum AlphabetType        ' тип алфавита
    AlphabetTypeUndef = 0
    AlphabetTypeLatin = 1       ' латинский
    AlphabetTypeCyrilic = 2     ' кириллический
End Enum
Public Type GroupExpr           ' тип для хранения выражений содержимого групп в текстовых выражениях (см. GroupsGet)
    Text As String              ' внутренний текст подстроки (без скобок)
    TextBeg As Long             ' позиция начала подстроки в исходной строке (включая скобки)
    TextEnd As Long             ' позиция конца подстроки в исходной строке (включая скобки)
    Bracket As Long             ' вид скобки/маркера группы (достаточно хранить индекс типа скобки в массиве)
    Level As Long               ' уровень вложенности (0-вне скобок, 1-внешние скобки, ... n-скобки n-уровня)
End Type
'--------------------------------------------------------------------------------
' POINTER
'--------------------------------------------------------------------------------
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
'--------------------------------------------------------------------------------
' SAFEARRAY
'--------------------------------------------------------------------------------
Private Const FADF_AUTO As Long = (&H1)
Private Const FADF_FIXEDSIZE As Long = (&H10)
Private Const FADF_HAVEVARTYPE As Long = (&H80)
Private Type SAFEARRAYBOUND         ' 8 bytes
    cElements As Long               ' +0 Количество элементов в размерности
    lLbound As Long                 ' +4 Нижняя граница размерности
End Type
'Private Type SAFEARRAY
'    cDims           As Integer      ' +0 Число размерностей
'    fFeatures       As Integer      ' +2 Флаг, используется функциями SafeArray
'    cbElements      As Long         ' +4 Размер одного элемента в байтах
'    cLocks          As Long         ' +8 Cчетчик ссылок, указывающий количество блокировок, наложенных на массив.
'    dummyPadding    As Long         ' +8 (x64 only!)
'    pvData          As Long         ' +12(x86) Указатель на данные
'                    As LongLong     ' +16(x64)
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
    cLocks As Long
#If Win64 Then
    dummyPadding As Long
    pvData As LongLong
#Else
    pvData As Long
#End If
'    rgsabound0 As SAFEARRAYBOUND
   cElements    As Long
   lLbound      As Long
End Type
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
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Private Declare PtrSafe Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoA" (ByVal locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare PtrSafe Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
#Else                   '<WIN32>
Private Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
' USER32
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Private Declare PtrSafe Function GetDC Lib "user32.dll" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
#Else                   '<WIN32>
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
' GDI32
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" (ByVal hdc As LongPtr, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32.dll" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function GetTextMetrics Lib "gdi32.dll" Alias "GetTextMetricsA" (ByVal hdc As LongPtr, ByRef lpMetrics As TEXTMETRIC) As Long
Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal e As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal q As Long, ByVal PAF As Long, ByVal f As String) As LongPtr
#Else                   '<WIN32>
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32.dll" Alias "GetTextMetricsA" (ByVal hdc As Long, ByRef lpMetrics As TEXTMETRIC) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal e As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal U As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal cp As Long, ByVal q As Long, ByVal PAF As Long, ByVal f As String) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
' OLEAUT32
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Private Declare PtrSafe Function SysReAllocStringLen Lib "oleaut32.dll" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr, Optional ByVal Length As Long) As Long
Private Declare PtrSafe Function SysAllocStringByteLen Lib "oleaut32.dll" (ByVal pBSTR As LongPtr, ByVal bLen As Long) As LongPtr
#Else                   '<WIN32>
Private Declare Function SysReAllocStringLen Lib "oleaut32.dll" (ByVal oleStr As Long, ByVal psz As Long, ByVal cch As Long) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32.dll" (ByVal oleStr As Long, ByVal bLen As Long) As Long
#End If                 '<WIN32>
'Private Type TPROC
'    hMem As Long
'    vtPtr As Long
'End Type
'Private aProc() As TPROC
'
'Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Const LOCALE_SLIST = &HC        ' разделитель элементов списка
Private Const LOCALE_SDECIMAL = &HE     ' десятичный разделитель
Private Const LOCALE_STHOUSAND = &HF    ' разделитель разрядов

Const LOCALE_USER_DEFAULT = &H400
Const LOCALE_SYSTEM_DEFAULT = &H800

Private Type Size
    cX As Long
    cY As Long
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
Private Enum apiDeviceCapability
    HORZSIZE = 4
    VERTSIZE = 6
    HORZRES = 8
    VERTRES = 10
    LOGPIXELSX = 88        '  Logical pixels/inch in X
    LOGPIXELSY = 90        '  Logical pixels/inch in Y
End Enum

' ==================
Public Function RegEx() As Object
' возвращает объект RegExp для работы с регулярными выражениями
' по RegEx см. https://regex.sorokin.engineer/ru/latest/regular_expressions.htmlStatic oRegEx As Object: If oRegEx Is Nothing Then Set oRegEx = CreateObject("VBScript.RegExp")
Static soRegEx As Object: If soRegEx Is Nothing Then Set soRegEx = CreateObject("VBScript.RegExp")
    Set RegEx = soRegEx
End Function
' ==================
' Замены стандартных строковых функций и другие быстрые строковые функции с
' http://www.xbeat.net/vbspeed/ см. также варианты на http://www.vbforums.com/showthread.php?540323-VB6-Faster-Split-amp-Join-(development)
' ==================
Public Sub xSplit(Expression As String, Result() As String, Optional Delimiter As String = " ") 'As Long
' Returns a zero-based, one-dimensional array containing a specified number of substrings.
'-------------------------
' Expression  - Required. String expression containing substrings and delimiters. If expression is a zero-length string, xSplit returns a single-element array containing a zero-length string.
' asToken()   - Required. One-dimensional string array that will hold the returned substrings. Does not have to be bound before calling xSplit, and is guaranteed to hold at least one element (zero-based) on return.
' Delimiter   - Optional. String character used to identify substring limits. If omitted, the space character (" ") is assumed to be the delimiter. If delimiter is a zero-length string, a single-element array containing the entire expression string is returned.
' returns number of elements
' стандартный Split работает быстрее на коротких строках, xSplit - на длинных
'-------------------------
' v.1.0.0       : 08.12.2001 - original SplitB04 by Chris Lucas, cdl1051@earthlink.net from http://www.xbeat.net/vbspeed/c_Split.htm#SplitB04
'-------------------------
Dim c As Long, sLen As Long, DelLen As Long, tmp As Long, Results() As Long
'Dim lCount As Long
    sLen = LenB(Expression) \ 2: DelLen = LenB(Delimiter) \ 2
    If sLen = 0 Or DelLen = 0 Then ReDim Preserve Result(0 To 0): Result(0) = Expression: Exit Sub ': xSplit = 1: Exit Function     ' пустая строка
' считаем разделители и запоминаем их позиции
    ReDim Preserve Results(0 To sLen): tmp = InStr(Expression, Delimiter)
    Do While tmp
        Results(c) = tmp: c = c + 1
        tmp = InStr(Results(c - 1) + 1, Expression, Delimiter)
    Loop
' наполняем массив
    ReDim Preserve Result(0 To c)
    If c = 0 Then Result(0) = Expression: Exit Sub ': xSplit = 1: Exit Function      ' нет разделителей
    Result(0) = Left$(Expression, Results(0) - 1)
    For c = 0 To c - 2
        Result(c + 1) = Mid$(Expression, Results(c) + DelLen, Results(c + 1) - Results(c) - DelLen)
    Next c
    Result(c + 1) = Right$(Expression, sLen - Results(c) - DelLen + 1)
    'xSplit = c + 2
End Sub
Public Function xJoin(Arr() As String, _
    Optional Delimiter As String = " ", _
    Optional ByVal Count As Long = -1) As String
' замена стандартного Join
'-------------------------
' v.1.0.0       : 01.10.2000 - original Join08 by by Matt Curland, mattcur@microsoft.com, www.PowerVB.com from http://www.xbeat.net/vbspeed/c_Join.htm#Join08
'-------------------------
' стабильно медленнее оригинальной, быстрые варианты без ASM инъекций никак
' Works with VB- or typelib-declared CopyMemory (pass strings with StrPtr)
'-------------------------
Dim Lower As Long
Dim Upper As Long
Dim cbDelim As Long
Dim cbTotal As Long
Dim i As Long
Dim cbCur As Long
Dim pCurDest As LongPtr, pDelim As LongPtr
    Lower = LBound(Arr)
    If Count = -1 Then
        Upper = UBound(Arr)
    Else
        Upper = Lower + Count - 1
    End If
    For i = Lower To Upper
        cbTotal = cbTotal + LenB(Arr(i))
    Next i
    cbDelim = LenB(Delimiter)
    If cbDelim Then cbTotal = cbTotal + cbDelim * (Upper - Lower)

    'Use API to avoid useless initialization
    CopyMemory ByVal VarPtr(xJoin), SysAllocStringByteLen(0, cbTotal), PTR_LENGTH
    'Use this instead if APIs are typelib-declared:
    ''xJoin = SysAllocStringByteLen(0, cbTotal)

    'Now, split into two different paths
    'a) No delimiter
    'b) Delimiter
    pCurDest = StrPtr(xJoin)
    If cbDelim = 0 Then
        For i = Lower To Upper
            cbCur = LenB(Arr(i))
            CopyMemory ByVal pCurDest, ByVal StrPtr(Arr(i)), cbCur
            pCurDest = pCurDest + cbCur
        Next i
    Else
        pDelim = StrPtr(Delimiter)
        For i = Lower To Upper - 1
            cbCur = LenB(Arr(i))
            CopyMemory ByVal pCurDest, ByVal StrPtr(Arr(i)), cbCur
            pCurDest = pCurDest + cbCur
            CopyMemory ByVal pCurDest, ByVal pDelim, cbDelim
            pCurDest = pCurDest + cbDelim
        Next i
        CopyMemory ByVal pCurDest, ByVal StrPtr(Arr(i)), LenB(Arr(i))
    End If
End Function
Public Function xReplace(ByRef Expression As String, _
    ByRef Find As String, ByRef ReplaceWith As String, _
    Optional ByVal Start As Long = 1, _
    Optional ByVal Count As Long = 2147483647, _
    Optional ByVal Compare As VbCompareMethod = vbBinaryCompare _
  ) As String
' Returns a string in which a specified substring has been replaced with another substring a specified number of times.
'-------------------------
' Expression  - Required. String expression containing substring to replace.
' Find    - Required. Substring being searched for. If Find is zero-length, xReplace returns a copy of Expression.
' ReplaceWith - Required. Replacement substring. If ReplaceWith is zero-length, xReplace returns a copy of Expression with all occurences of Find removed.
' Start   - Optional. Position within expression where substring search is to begin. If omitted, 1 is assumed. Must be used in conjunction with count.
' Count   - Optional. Number of substring substitutions to perform. If omitted, the default value is -1, which means make all possible substitutions. Must be used in conjunction with start.
' Compare - Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. If omitted, the default value is 0, which means perform a binary comparison.
'-------------------------
' v.1.0.0       : 18.12.2000 - original Replace09 by Jost Schwider, jost@schwider.de from http://www.xbeat.net/vbspeed/c_Replace.htm#Replace09
'-------------------------
    If LenB(Find) Then
        If Compare = vbBinaryCompare Then
            p_ReplaceBin xReplace, Expression, Expression, Find, ReplaceWith, Start, Count
        Else
            p_ReplaceBin xReplace, Expression, LCase$(Expression), LCase$(Find), ReplaceWith, Start, Count
        End If
    Else
        xReplace = Expression
    End If
End Function
Private Static Sub p_ReplaceBin(ByRef Result As String, _
    ByRef Text As String, ByRef Search As String, _
    ByRef sOld As String, ByRef sNew As String, _
    ByVal Start As Long, ByVal Count As Long)
' by Jost Schwider, jost@schwider.de, 20001218
'-------------------------
Dim TextLen As Long, OldLen As Long, NewLen As Long
Dim ReadPos As Long, WritePos As Long, CopyLen As Long
Dim Buffer As String, BufferLen As Long
Dim BufferPosNew As Long, BufferPosNext As Long
  
    If Start < 2 Then
        Start = InStrB(Search, sOld)
    Else
        Start = InStrB(Start + Start - 1, Search, sOld)
    End If
    If Start Then
        OldLen = LenB(sOld): NewLen = LenB(sNew)
        Select Case NewLen
        Case OldLen
            Result = Text
            For Count = 1 To Count
                MidB$(Result, Start) = sNew
                Start = InStrB(Start + OldLen, Search, sOld)
                If Start = 0 Then Exit Sub
            Next Count
            Exit Sub
        Case Is < OldLen
            TextLen = LenB(Text)
            If TextLen > BufferLen Then Buffer = Text: BufferLen = TextLen
            ReadPos = 1: WritePos = 1
            If NewLen Then
                For Count = 1 To Count
                    CopyLen = Start - ReadPos
                    If CopyLen Then
                        BufferPosNew = WritePos + CopyLen
                        MidB$(Buffer, WritePos) = VBA.MidB$(Text, ReadPos, CopyLen)
                        MidB$(Buffer, BufferPosNew) = sNew
                        WritePos = BufferPosNew + NewLen
                    Else
                        MidB$(Buffer, WritePos) = sNew: WritePos = WritePos + NewLen
                    End If
                    ReadPos = Start + OldLen: Start = InStrB(ReadPos, Search, sOld)
                    If Start = 0 Then Exit For
                Next Count
            Else
                For Count = 1 To Count
                    CopyLen = Start - ReadPos
                    If CopyLen Then MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen): WritePos = WritePos + CopyLen
                    ReadPos = Start + OldLen
                    Start = InStrB(ReadPos, Search, sOld)
                    If Start = 0 Then Exit For
                Next Count
            End If
            If ReadPos > TextLen Then
                Result = LeftB$(Buffer, WritePos - 1)
            Else
                MidB$(Buffer, WritePos) = MidB$(Text, ReadPos)
                Result = LeftB$(Buffer, WritePos + LenB(Text) - ReadPos)
            End If
            Exit Sub
        Case Else
            TextLen = LenB(Text): BufferPosNew = TextLen + NewLen
            If BufferPosNew > BufferLen Then Buffer = Space$(BufferPosNew): BufferLen = LenB(Buffer)
            ReadPos = 1: WritePos = 1
            For Count = 1 To Count
                CopyLen = Start - ReadPos
                If CopyLen Then
                    BufferPosNew = WritePos + CopyLen: BufferPosNext = BufferPosNew + NewLen
                    If BufferPosNext > BufferLen Then Buffer = Buffer & Space$(BufferPosNext): BufferLen = LenB(Buffer)
                    MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen)
                    MidB$(Buffer, BufferPosNew) = sNew
                Else
                    BufferPosNext = WritePos + NewLen
                    If BufferPosNext > BufferLen Then Buffer = Buffer & Space$(BufferPosNext): BufferLen = LenB(Buffer)
                    MidB$(Buffer, WritePos) = sNew
                End If
                WritePos = BufferPosNext: ReadPos = Start + OldLen
                Start = InStrB(ReadPos, Search, sOld)
                If Start = 0 Then Exit For
            Next Count
            If ReadPos > TextLen Then
                Result = LeftB$(Buffer, WritePos - 1)
            Else
                BufferPosNext = WritePos + TextLen - ReadPos
                If BufferPosNext < BufferLen Then
                    MidB$(Buffer, WritePos) = MidB$(Text, ReadPos)
                    Result = LeftB$(Buffer, BufferPosNext)
                Else
                    Result = LeftB$(Buffer, WritePos - 1) & MidB$(Text, ReadPos)
                End If
            End If
            Exit Sub
        End Select
    Else
        Result = Text
    End If
End Sub
Public Function ReplaceMany(Source As String, _
    ParamArray ReplacePairs()) As String
' производит множественную замену в строке значением из выражения
'-------------------------
' Source - исходная строка
' ReplacePairs  - подстановочные значения в виде "OldText=NewText"
'-------------------------
Const cDelim = ";", cTagDelim = "="
Dim Result As String
    On Error GoTo HandleError
    Result = Source: If Len(Result) = 0 Then GoTo HandleExit
Dim i As Long: i = 1 'LBound(Terms)             ' начинаем с [%1%]
Dim sTerms As String:         sTerms = Join(ReplacePairs, cDelim)
Dim cTerms As New Collection, aKeys, sKey 'As String
    Call TaggedString2Collection(sTerms, cTerms, aKeys, cDelim, cTagDelim)
    For Each sKey In aKeys: Result = Replace(Result, sKey, cTerms(sKey)): Next 'sKey
HandleExit:     ReplaceMany = Result: Exit Function
HandleError:    Err.Clear: Resume HandleExit
End Function

Public Function Tokenize(Source As String, _
    Tokens() As String, _
    Optional Delims As String = c_strSymbSpaces & c_strSymbPunct & c_strSymbParenth & c_strSymbMath, _
    Optional Positions, _
    Optional IncEmpty As Boolean = False _
    ) As Long
' разбивает строку на подстроки по набору разделителей
'-------------------------
' Source    - исходная строка
' Tokens()  - на выходе содержит массив строк выделенных из исходной строки по набору разделителей
' Delims    - набор возможных разделителей
' Positions - (необязательный) массив позиций начала элементов в исходной строке (необходимы для замены/удаления элементов строки)
' IncEmpty  = False - пропуск пустых элементов - последовательные разделители будут рассматриваться как один
'           = True  - результирующий массив будет включать пустые элементы между последовательными разделителями
'-------------------------
' v.1.1.0       : 31.10.2019 - добавлен необязательный параметр Positions - возвращающий массив позиций найденных токенов в исходной строке
' v.1.0.0       : 24.11.2000 - original Tokenize02 by Donald, donald@xbeat.net modified by G.Beckmann, G.Beckmann@NikoCity.de from http://www.xbeat.net/vbspeed/c_Tokenize.htm#Tokenize04
'-------------------------
Const ARR_CHUNK& = 1024
Dim cExp As Long, ubExpr As Long
Dim cDel As Long, ubDelim As Long
Dim aExpr() As Integer, aDelim() As Integer
Dim sa1 As SAFEARRAY1D, sa2 As SAFEARRAY1D
Dim cTokens As Long, iPos As Long
Dim bPos As Boolean
Dim Result As Long

    Result = -1
    On Error GoTo HandleError
    ubExpr = Len(Source): ubDelim = Len(Delims)
    ' готовим SAFEARRAY для исходной строки
    sa1.cbElements = 2:   sa1.cElements = ubExpr
    sa1.cDims = 1:        sa1.pvData = StrPtr(Source)
    ' заполняем массив символов исходной строки
    CopyMemory ByVal VarPtrArray(aExpr), VarPtr(sa1), PTR_LENGTH ' 4
    ' готовим SAFEARRAY для строки разделителей
    sa2.cbElements = 2:   sa2.cElements = ubDelim
    sa2.cDims = 1:        sa2.pvData = StrPtr(Delims)
    ' заполняем массив символов разделителей
    CopyMemory ByVal VarPtrArray(aDelim), VarPtr(sa2), PTR_LENGTH ' 4
    
    ' инициализируем результирующие массивы
    If IncEmpty Then ReDim Preserve Tokens(ubExpr) Else ReDim Preserve Tokens(ubExpr \ 2)
    ' проверяем необходимость возвращать позиции найденных элементов
    bPos = Not IsMissing(Positions): If bPos Then bPos = IsArray(Positions)
    If bPos Then If IncEmpty Then ReDim Preserve Positions(ubExpr) Else ReDim Preserve Positions(ubExpr \ 2)
    
    ubDelim = ubDelim - 1
    For cExp = 0 To ubExpr - 1
    ' перебираем все символы исходной строки
        For cDel = 0 To ubDelim
    ' перебираем все символы строки разделителей
            If aExpr(cExp) = aDelim(cDel) Then
                If cExp > iPos Then
        ' если текущий символ исходной строки совпадает с разделителем
            ' и предыдущией не был разделителем
            ' (если предыдущий символ также был разделителем было бы cExp=iPos)
                ' сохраняем фрагмент строки
                    Tokens(cTokens) = Mid$(Source, iPos + 1, cExp - iPos)
                    If bPos Then Positions(cTokens) = iPos + 1
                    cTokens = cTokens + 1
                ElseIf IncEmpty Then
            ' или если выводим пустые строки
                ' сохраняем пустую строку
                    Tokens(cTokens) = vbNullString
                    If bPos Then Positions(cTokens) = iPos + 1
                    cTokens = cTokens + 1
                End If
        ' сохраняем позицию начала следующего символа строки
                iPos = cExp + 1: Exit For
            End If
        Next cDel
    Next cExp
    ' если после последнего разделителя остались символы или указано выводить пустые строки
    ' добавляем в остаток
    If (cExp > iPos) Or IncEmpty Then
        Tokens(cTokens) = Mid$(Source, iPos + 1)
        If bPos Then Positions(cTokens) = iPos + 1
        cTokens = cTokens + 1
    End If
    ' обрезаем результирующие массивы по количеству найденных элементов
    If cTokens = 0 Then Erase Tokens() Else ReDim Preserve Tokens(cTokens - 1)
    If bPos Then If cTokens = 0 Then Erase Positions() Else ReDim Preserve Positions(cTokens - 1)
    ' возвращаем количество найденных элементов
    Result = cTokens '- 1
    ' очищаем вспомогательные массивы
    ZeroMemory ByVal VarPtrArray(aExpr), PTR_LENGTH '4
    ZeroMemory ByVal VarPtrArray(aDelim), PTR_LENGTH '4
HandleExit:  Tokenize = Result: Exit Function
HandleError: Result = -1: Err.Clear: Resume HandleExit
End Function
Public Function PlaceHoldersSetByIndex(Source As String, ParamArray Terms()) As String
' заменяет подстановочные шаблоны вида [%n%] (где n - номер параметра), значениями массива параметров
'-------------------------
' Source - исходная строка
' Terms - подстановочные значения
Const LBr = "[%", RBr = "%]"                ' константы левой/правой скобок шаблона
Dim Result As String
    On Error GoTo HandleError
    Result = Source: If Len(Result) = 0 Then GoTo HandleExit
Dim i As Long: i = 1 'LBound(Terms)             ' начинаем с [%1%]
Dim Term, sTemp As String
    If IsArray(Terms(0)) Then GoTo HandleArray
' передан массив параметров
    For Each Term In Terms
        sTemp = LBr & CStr(i) & RBr             ' создаем шаблон
        Result = Replace(Result, sTemp, Term)   ' замена
        i = i + 1
    Next
    GoTo HandleExit
HandleArray:
' на тот случай если в параметре передали обычный массив
    For Each Term In Terms(0)
        sTemp = LBr & CStr(i) & RBr             ' создаем шаблон
        Result = Replace(Result, sTemp, Term)   ' замена
        i = i + 1
    Next
HandleExit:  PlaceHoldersSetByIndex = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function PlaceHoldersSetByNames(Source As String, _
    ParamArray NamedTerms()) As String
' заменяет подстановочные шаблоны вида [%Param1%], значениями массива параметров
Const c_strProcedure = "PlaceHoldersSetByNames"
' Source - исходная строка
' NamedTerms  - подстановочные значения в виде "Param1=Value1"
Const LBr = "[%", RBr = "%]"                ' константы левой/правой скобок шаблона
Const cDelim = ";"
Dim Result As String
    On Error GoTo HandleError
    Result = Source: If Len(Result) = 0 Then GoTo HandleExit
Dim i As Long: i = 1 'LBound(Terms)             ' начинаем с [%1%]
Dim sTerms As String:         sTerms = Join(NamedTerms, cDelim)
Dim cTerms As New Collection: Call TaggedString2Collection(sTerms, cTerms, Delim:=cDelim)
    Result = PlaceHoldersSet(Result, cTerms, False, LBr, RBr)
HandleExit:     PlaceHoldersSetByNames = Result: Exit Function
HandleError:    Err.Clear: Resume HandleExit
End Function
Public Function PlaceHoldersSet(ByRef Source As String, _
    ByRef NamedTerms As Collection, _
    Optional AskMissing As Boolean = False, _
    Optional LBr As String = "[%", Optional RBr As String = "%]") As String
' заменяет в строке именные подстановочные шаблоны вида [%Param1%] значениями из коллекции
'-------------------------
' Source    - исходная строка
' NamedTerms - коллекция подставляемых значений. в качестве имени подстановочной переменно берётся ключ элемента коллекции
' AskMissing - запрашивать значения отсутствующих в коллекции элементов
' LBr/RBr   - левая/правая скобки отмечающие границы имени переменной (напр [%Param1%])
'-------------------------
' v.1.1.1       : 30.01.2020 - добавлена возможность использования модификаторов (формат см. p_TermModify)
' v.1.1.0       : 20.01.2020 - изменён алгоритм замены переменных. Теперь можно в качестве значения в коллекции подавать выражения
'-------------------------
' ToDo: шаблон должен обязательно содержать условный оператор
'-------------------------
Dim Result As String
    On Error GoTo HandleError
    Result = Source: If Len(Result) = 0 Then GoTo HandleExit
Dim Term As String, Xpr As String, Key As String, Value As String
Dim i As Long: i = 1
' для доп.модификаторов. формат модификаторов см. p_TermModify
Const c_ModLBr = "{", c_ModRBr = "}"
Dim Par As String, Pos As Long
    ' ищем именную переменную в выражении
    Do While p_FindNamedPlaceHolder(Result, Xpr, i, , LBr, RBr)
    ' найдена именная переменная
        ' анализируем строку на наличие дополнительных модификаторов
        Key = p_TermModify(Xpr, Par, Operation:=1)
        ' если не нужно - просто сделать: Key = Xpr
    ' получаем ее значение из коллекции (или запрашиваем при отсутствии)
        If p_IsExist(Key, NamedTerms, Term) Then
        ' нашли в коллекции
        ElseIf AskMissing Then
        ' запрашиваем значение отсутствующей в коллекции переменной
            Term = InputBox("Укажите значение переменной " & vbCrLf & Key & ":", "Переменная не найдена!")
        ' ??? и добавляем её в набор для повторного использования
            NamedTerms.Add Term, Key
        Else
        ' если не нашли и не запрашиваем - оставляем как есть
            ' можно создать список отсутствующих переменных
            ' или другим способом передать управление пользователю
            ' чтобы он мог приемлемым способом задать значения отсутствующего параметра,
            ' но чтобы не усложнять больше необходимого оставим так
            Term = LBr & Xpr & RBr: GoTo HandleNext
        End If
    ' рекурсивно проверяем (и заменяем) полученное значение на наличие в нем именных переменных
        Term = PlaceHoldersSet(Term, NamedTerms, AskMissing, LBr, RBr)
    ' если есть модификаторы - применяем их к Term
        If Len(Par) > 0 Then Term = p_TermModify(Term, Par, Operation:=0)
    '' предполагалось, что это могло улучшить скорость вычисления выражений
        '    If EvalExpres Then
    '' если указано вычислять выражения в процесе подстановки
        '    ' вычисляем значение выражения и
        '    ' заменяем вычислимое выражение в коллекции его значением
        '        If p_IsEvalutable(Term, Value) Then
        '            Term = Value: With NamedTerms: .Remove (Key): .Add Key, Term: End With
        '        End If
        '    End If
    ' производим замену именной переменной полученным значением по всему выражению
        Result = Left$(Result, i - 1) & Replace(Result, LBr & Xpr & RBr, Term, i) ' замена
HandleNext:  i = i + Len(Term) 'If i > Len(Result) Then Exit Do ' повторяем пока в выражении есть именные переменные
    Loop
    'Erase Keys()
HandleExit:  PlaceHoldersSet = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Function PlaceHoldersGet(ByRef Source As String, ByRef Template As String, _
    Optional ByRef NamedTerms As Collection, Optional Keys, _
    Optional LBr As String = "[%", Optional RBr As String = "%]", _
    Optional ReplaceExisting As Integer = False, _
    Optional Method As Integer = 0, _
    Optional MultiSfx As String) As Boolean
' проверяет строку на соответствие шаблону и извлекает из строки коллекцию значений именованых параметров
'-------------------------
' Source    - исходная строка
' Template  - шаблон строки, может содержать подстановочные переменные вида [%Param1%]
' NamedTerms - (возвращаемое значение) коллекция значений подстановочных переменных вида [%Param1%], извлечённых из исходной строки
' Keys      - (возвращаемое значение) массив имён переменных вида [%Param1%], извлечённых из исходной строки (являются ключами NamedTerms)
' LBr/RBr   - левая/правая скобки отмечающие границы имени переменной (напр [%Param1%])
' ReplaceExisting - определяет сохраняемые в коллекцию значения в случае
' если шаблон содержит несколько ссылок на переменную с одним и тем же именем
'   0 - будет сохранено первое значение
'  -1 - будет сохранено последнее значение
'   1 - будут сохранены все значения в переменных с добавлением суффикса
' Method    - способ сравнения подстрок шаблона
'   0 - по простой подстроке (InStr)
'   1 - по Like подстроке (InStrLike)
'   2 - по RegEx выражению (InStrRegEx)
' MultiSfx - признак суффикса для повторяющихся имен (при ReplaceExisting=1) д.б. что-то заведомо отсутствующее в именах переменных
'-------------------------
' v.1.0.2       : 16.07.2020 - исходный вариант себя не оправдал - функция переписана. добавлена поддержка множественных вхождений (ReplaceExisting)
' v.1.0.1       : 04.02.2020 - расширен синтаксис шаблона за счет выражений VBA.Like и VBS.RegExp
' v.1.0.0       : 03.02.2020 - исходная версия
'-------------------------
' ToDo: шаблон должен обязательно содержать: условный оператор, список допустимых значений, якоря для определения позиции
' - при ReplaceExisting = -1 - ошибка при добавлении в NamedTerms
'-------------------------
Const cSfx = "~&#" ' признак суффикса для имени переменной при множественных значениях (по-умолчанию)
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If Len(Source) = 0 Then GoTo HandleExit Else If Len(Template) = 0 Then GoTo HandleExit
Dim Xpr As String, Key As String, Item As String
Dim Part As String, Params As String, Found As String
Dim bLen As Integer: bLen = Len(LBr) + Len(RBr)
Dim bKeys As Boolean, aKeys() As String: bKeys = Not IsMissing(Keys)
Dim rBeg As Long, rEnd As Long, cBeg As New Collection
Dim tBeg As Long, tEnd As Long
Dim i As Long, j As Long, x As Long
Dim sKey As String
Dim bOK As Long
    i = 0
    rBeg = 1: tBeg = 1: tEnd = tBeg
    Set NamedTerms = New Collection
    If ReplaceExisting = 1 And Len(MultiSfx) = 0 Then MultiSfx = cSfx
    Do Until tBeg > Len(Template)
' ищем именную переменную
        Found = vbNullString
    ' если найдена именная переменная работаем с результатми поиска
    ' если не найдена, но до конца строки еще остались символы - проверить остаток до конца строки
        ' (осталась не сохранена предыдущая переменная)
        If p_FindNamedPlaceHolder(Template, Xpr, tEnd, , LBr, RBr) Then Else tEnd = Len(Template) + 1
    ' берём кусок из шаблона от конца прошлой именной переменной до начала найденной
        Part = Mid$(Template, tBeg, tEnd - tBeg)
        If j = 0 Then ' ???
    ' если это первый кусок шаблона (не именованная переменная)- получаем все вхождения фрагмента в массив
'            aBeg = InStrAll(rBeg, Source, Part, aFound, Method): jMax = UBound(aBeg)
            Do While rBeg <= Len(Source)
                j = j + 1
                Select Case Method
                Case 1:     rEnd = InStrLike(rBeg, Source, Part, Found)
                Case 2:     rEnd = InStrRegEx(rBeg, Source, Part, Found)
                Case Else:  rEnd = InStr(rBeg, Source, Part): Found = Part
                End Select
                If rEnd = 0 Then Exit Do
                rBeg = rEnd + Len(Found)
                'sKey = c_idxPref & j     ' ??? сохраняем имя ключа - понадобится для индексов переменных при поиске множественных результатов
                cBeg.Add rBeg ', sKey
            Loop
        End If
    ' если ищем с конца - берем последнее вхождение, иначе - начинаем с первого
        If ReplaceExisting = -1 Then j = cBeg.Count Else j = 1
    ' если это первый проход найдена только левая ганица (перед переменной) - ищем правую
        If Len(Key) = 0 Then GoTo HandleNext
        i = i + 1   ' увеличиваем индекс переменной в шаблоне
        Do Until j > cBeg.Count
' перебираем возможные совпадения
            If Len(Part) > 0 Then
' сравниваем проверяемую строку с текущим куском шаблона
    ' если перед именованной переменной есть фрагмент шаблона для распознавания
' !!! непонятно как разделять две и более идущие подряд именованные переменные
    ' вероятно надо извлекать текст в первую потом по мере анализа Params отсекать хвост
    ' постепенно распределяя его по оставшимся.
' пока иметь ввиду это как ограничение допустимых шаблонов
            Select Case Method
            Case 1:     rEnd = InStrLike(cBeg(j), Source, Part, Found)
            Case 2:     rEnd = InStrRegEx(cBeg(j), Source, Part, Found)
            Case Else:  rEnd = InStr(cBeg(j), Source, Part): If rEnd > 0 Then Found = Part Else Found = vbNullString
            End Select
            End If
        ' если кусок не найден - вычеркиваем фрагмент из дальнейшей обработки и переходим к следующему
            bOK = rEnd > 0
            If Not bOK Then GoTo HandleNotOk
        ' текущий вариант подтвержден - извлекаем значение именованной переменной
            Item = Mid$(Source, cBeg(j), rEnd - cBeg(j))
'' <<< здесь можно проверить соответствие Item заданному в Params
'            If Len(Params) > 0 Then
'    '         ' если Item не соотв Params - текущий фрагмент не значение переменной,
'    '         ' а возможный фрагмент шаблона - ??? подумать как реагировать
'Stop
'                Item = p_TermModify(Item, Params, Operation:=0)
'                bOk = ??? '
'            End If
'            If Not bOk then GoTo HandleNotOk
    ' сохраняем позицию найденого фрагмента для дальнейшего использования
            rBeg = rEnd + Len(Found)
            cBeg.Remove (j): If j <= cBeg.Count Then cBeg.Add rBeg, Before:=j Else cBeg.Add rBeg, After:=cBeg.Count
    ' формируем имя в коллекции для извлекаемого элемента
        ' если сохраняем единственный результат - берём извлечённое имя переменной из шаблона
        ' если сохраняем все результаты - формируем на основе имени переменной из шаблона и индекса элемента в коллекции
            sKey = Key
            If ReplaceExisting = 1 And j > 1 Then sKey = sKey & MultiSfx & (j - 1)
            ' имеет смысл сохранять результаты в NamedTerms так,
            ' чтобы все имеющие отношение к одному индексу хранились подряд
            ' это облегчит удаление если ветка впоследствии окажется забракованной
'Stop
            If NamedTerms.Count = 0 Then NamedTerms.Add Item, sKey Else NamedTerms.Add Item, sKey, After:=j * i - 1    ', Before:=
            ' увеличиваем индекс совпадения в результате
            If ReplaceExisting = 1 Then j = j + 1 Else Exit Do
            GoTo HandleNextVar
HandleNotOk:
    ' текущий вариант не подтвержден
        ' удаляем его из коллекции совпадений для предотвращения дальнейшего разбора этой ветки
            cBeg.Remove (j)
        ' также надо исключить найденые именованные переменные неподтвержденного варианта
            If ReplaceExisting <> 1 Then
            ' для поиска единичных совпадений - просто очистить NamedTerms и aKeys
                Set NamedTerms = New Collection
                If ReplaceExisting = -1 Then j = cBeg.Count
            Else
            ' для множественных - надо убрать предыдущие результаты
            ' не совпавшей (текущей) ветки по индексу j в NamedTerms
                For x = 1 To i - 1: NamedTerms.Remove (j * i - x): Next x
            End If
HandleNextVar:
        Loop
HandleNext:
        If tEnd > Len(Template) Then Exit Do
' извлекаем из именной переменной собственно имя без дополнительных модификаторов
        Key = p_TermModify(Xpr, Params, Operation:=1)
' переводим указатель в строке шаблона на символ после найденной именной переменной
        tBeg = tEnd + Len(Xpr) + bLen
        tEnd = tBeg
    Loop
'HandleExitDo:
    If bKeys Then Keys = p_GetCollKeys(NamedTerms)
HandleExit:  PlaceHoldersGet = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_FindNamedPlaceHolder(ByRef Source As String, _
    Optional ByRef Name As String, _
    Optional ByRef sBeg As Long = 1, Optional ByRef sEnd As Long = 0, _
    Optional LBr As String = "[%", Optional RBr As String = "%]") As Boolean
' ищет в строке именную переменную, ограниченную разделителями, возвращает её имя и границы
'-------------------------
' Source - строка в которой производится поиск
' Name - имя найденной переменной (без скобок)
' sBeg - позиция начала найденной переменной в строке (включая скобки)
' sEnd - позиция конца найденной переменной в строке (включая скобки)
' LBr/RBr - левая/правая скобки отмечающие границы имени переменной (напр %Param1%)
'-------------------------
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If Len(Source) = 0 Then GoTo HandleExit
Dim pBeg As Long, pEnd As Long
    ' ищем в выражении левую скобку
    pBeg = InStr(sBeg, Source, LBr): If pBeg = 0 Then GoTo HandleExit Else sBeg = pBeg: pBeg = pBeg + Len(LBr)
    ' ищем в выражении правую скобку
    pEnd = InStr(pBeg, Source, RBr): If pEnd = 0 Then GoTo HandleExit Else sEnd = pEnd + Len(RBr)
    ' получаем строку между скобками
    Name = Mid$(Source, pBeg, pEnd - pBeg)
    Result = True 'Len(Name) > 0
HandleExit:  p_FindNamedPlaceHolder = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Private Function p_TermModify(ByVal Term As String, ByRef Params As String, _
    Optional Operation = 0, Optional ReplaceExisting As Integer = False) As String
' обрабатывает строку с модификаторами
'-------------------------
' Term      - обрабатываемое значение
' Params    - строка параметров
' Operation - параметр определяющий тип обработки
'   0 - применяет очищенную строку модификаторов Params к значению Term и возвращает в Result
'   1 - определяет наличие в строке модификаторов, возвращает:
'       в Params - очищенную строку модификаторов, в Result - имя параметра
' сделано в одной функции чтобы все элементы описания формата модификаторов хранились в одном месте
' ReplaceExisting - флаг определяющий поведение при обнаружении параметров с одинаковым именем при разборе параметров
'   0 - сработает первый, последующие будут игнорироваться
'  -1 - параметры одного типа будут заменяться - сработает последний
'-------------------------
' v.1.0.3       : 05.02.2020 - для удобства настройки выделены в отдельные функции распознавание допустимых имен и параметров модификаторов
'-------------------------
' Формат строки модификаторов: {Модификатор1:Параметр1-1,...,Параметр1-X1;...;МодификаторN:ПараметрN-1,...,ПараметрN-XN}
Const c_ModLBr = "{", c_ModRBr = "}" ' скобки выделяющие строку модификаторов в выражении
Const cXprDelim = ";" ' разделитель выражений модификаторов в строке
Const cNamDelim = ":" ' разделитель имени/параметров модификатора
Const cParDelim = "," ' разделитель параметров модификатора
Dim Pos As Long
Dim Result As String: Result = Term
    On Error GoTo HandleError
    If Len(Term) = 0 Then GoTo HandleExit
    Select Case Operation
    Case 0
' обработка Term и применение параметров модификатора
        If Len(Params) = 0 Then GoTo HandleExit
        Dim Xpr, Par As String
    ' получаем массив модфикаторов с параметрами
        For Each Xpr In Split(Params, cXprDelim)
    ' перебираем наборы модификаторов
        ' получаем тип модификатора и список его параметров
            Pos = InStr(1, Xpr, cNamDelim)
            If Pos > 0 Then Par = Mid$(Xpr, Pos + Len(cNamDelim)): Xpr = Left$(Xpr, Pos - 1)
    ' применяем модификатор
            Result = p_TermModifyXprGet(Result, Xpr, Par, cParDelim, ReplaceExisting)
        Next Xpr
    Case Else
' извлечение из Term имени параметра, определение наличия в ней модификаторов
    ' раскрываем скобки и извлекаем строку модификаторов
        Result = Term: Params = vbNullString
        Pos = InStr(1, Result, c_ModLBr): If Pos = 0 Then GoTo HandleExit
        If Right$(Result, Len(c_ModRBr)) <> c_ModRBr Then GoTo HandleExit
        Term = Left$(Result, Pos - 1): Pos = Pos + Len(c_ModLBr)
        Params = Mid$(Result, Pos, Len(Result) - Pos - Len(c_ModRBr) + 1)
        Result = Term
    End Select
HandleExit:  p_TermModify = Result: Exit Function
HandleError: Result = Term: Err.Clear: Resume HandleExit
End Function
Private Function p_TermModifyXprGet(Term As String, _
    Modificator, Optional Params As String, _
    Optional ParDelim As String, _
    Optional ReplaceExisting As Integer = False) As String
' применяет модификатор к значению и возвращает результат
'-------------------------
' Term - обрабатываемое значение
' Modificator - имя модификатора применяемого к значению
' Params - набор параметров модификатора
' ParDelim - разделитель параметров в списке
' ReplaceExisting - флаг определяющий способ реакции на однотипные параметры
'-------------------------
' v.1.0.2       : 31.01.2020 - изменён способ передачи аргументов функциям модификаторов на более удобный
'-------------------------
Dim Result As String: Result = Term
    On Error GoTo HandleError
    If Len(Modificator) = 0 Then GoTo HandleExit
Dim sFun As String, sKey As String, sVal
Dim cPar As Collection
    Select Case LCase(Modificator)
' <<< здесь нужно описать допустимые имена модификаторов и формат вызываемых ими функций
' !!! следить за порядком аргументов в пользовательских функциях !!! - по имени передавать не получается, пропускать параметры тоже нельзя
    Case "верхрег", "ucase":    Result = UCase(Result)
    Case "нижрег", "lcase":     Result = LCase(Result)
    Case "первверхрег", "pcase": Result = StrConv(Result, vbProperCase)
    Case "склонять", "decline":   sFun = "DeclineWords('" & Result & "'"
                ' получаем параметры модификатора
                    Set cPar = p_TermModifyParGet(Params, ParDelim, ReplaceExisting)
                ' заполняем параметры функции
                    sKey = "NewCase": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "NewNumb": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "NewGend": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "Animate": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "IsFio": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    'sKey = "SkipWords": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, vbNullString)
                    sFun = sFun & ")"
                    If p_IsEvalutable(sFun, Result) Then Else Err.Raise vbObjectError + 512
    Case "числовтекст", "numtowords": sFun = "NumToWords(" & Result
                ' получаем параметры модификатора
                    Set cPar = p_TermModifyParGet(Params, ParDelim, ReplaceExisting)
                ' заполняем параметры функции
                    sKey = "NewCase": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "NewNumb": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "NewGend": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "Animate": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "NewType": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    'sKey = "Unit": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, vbNullString)
                    'sKey = "SubUnit": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, vbNullString)
                    'sKey = "DecimalPlaces": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    'sKey = "TranslateFrac": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sFun = sFun & ")"
                    If p_IsEvalutable(sFun, Result) Then Else Err.Raise vbObjectError + 512
    Case "всписке", "in" ' проверка принадлежности значения списку допустимых значений, заданному в Params
        ' из PlaceHolderGet - извлечение значения из строки - проверяем вероятное значение по списку - при соответствии - возвращаем, при несоответствии продолжаем поиск в строке
        ' из PlaceHolderSet - установка значения в строке - проверяем подставляемое значение на соответствие списку - если соответствует - подставляем, а если нет - ???
                ' получаем параметры модификатора (список допустимых значений)
                    Result = vbNullString
                    Set cPar = p_TermModifyParGet(Params, ParDelim, ReplaceExisting)
                    For Each sVal In cPar
                        If Left$(Term, Len(sVal)) = sVal Then Result = sVal: Exit For
                    Next sVal
    Case "тип", "is"     ' проверка соответствия типу, заданному в Params
                    Result = vbNullString
                    Set cPar = p_TermModifyParGet(Params, ParDelim, ReplaceExisting)
                    Dim l As Long
                    For Each sVal In cPar
                        For l = Len(Term) To 1 Step -1
                            sKey = Left$(Term, l)
                            Select Case sVal
                            Case "число", "num":    If IsNumeric(sKey) Then Result = sKey
                            Case "дата", "date":    If IsDate(sKey) Then Result = sKey
                            Case "слово", "word":   sFun = "^[a-zA-Zа-яА-ЯёЁ]*$": RegEx.Pattern = sFun: If RegEx.Test(sKey) Then Result = sKey
                            Case "текст", "text":   Result = sKey ' под текст подходит всё
                            Case "словорус", "rus": sFun = "^[а-яА-ЯёЁ]*$": RegEx.Pattern = sFun: If RegEx.Test(sKey) Then Result = sKey
                            Case "словоанг", "eng": sFun = "^[a-zA-Z]*$": RegEx.Pattern = sFun: If RegEx.Test(sKey) Then Result = sKey
                            'Case "имя", "var":   sFun = "^[a-zA-Zа-яА-ЯёЁ][_a-zA-Zа-яА-ЯёЁ0-9]*$": RegEx.Pattern = sFun: If RegEx.Test(sKey) Then Result = sKey
                            Case Else: GoTo HandleExit ' можно сделать проверку на какой-то определённый набор символов
                            End Select
                            If Len(Result) > 0 Then Exit For
                        Next l
                        If Len(Result) > 0 Then Exit For
                    Next sVal
'    Case "если","if"     ' проверка условия возвращает вариант соответствующий условию
'    Case "выбор","choose"    ' проверка нескольких условий возвращает вариант соответствующий первому истинному условию
    Case Else
    End Select
HandleExit:  p_TermModifyXprGet = Result: Exit Function
HandleError: Result = Term: Err.Clear: Resume HandleExit
End Function
Private Function p_TermModifyParGet(Params As String, _
    Optional ParDelim As String, _
    Optional ReplaceExisting As Integer = False) As Collection
' сопоставляет имена параметров модификаторов именам и значениям параметров функций и фозвращает в виде коллекции
'-------------------------
' Params - набор параметров модификатора
' ParDelim - разделитель параметров в списке
' ReplaceExisting - флаг определяющий способ реакции на однотипные параметры
Dim sKey As String, sVal 'As String
Dim cPar As New Collection, Par
    For Each Par In Split(Params, ParDelim)
        Select Case LCase(Par)
' <<< здесь нужно описать допустимые имена параметров функций, вызываемых модификаторами и их значения
        Case "колич":   sKey = "NewType": sVal = NumeralOrdinal
        Case "поряд":   sKey = "NewType": sVal = NumeralCardinal
        Case "им":      sKey = "NewCase": sVal = DeclineCaseImen
        Case "род":     sKey = "NewCase": sVal = DeclineCaseRod
        Case "дат":     sKey = "NewCase": sVal = DeclineCaseDat
        Case "вин":     sKey = "NewCase": sVal = DeclineCaseVin
        Case "тв":      sKey = "NewCase": sVal = DeclineCaseTvor
        Case "пред":    sKey = "NewCase": sVal = DeclineCasePred
        Case "ед":      sKey = "NewNumb": sVal = DeclineNumbSingle
        Case "мн":      sKey = "NewNumb": sVal = DeclineNumbPlural
        Case "муж":     sKey = "NewGend": sVal = DeclineGendMale
        Case "жен":     sKey = "NewGend": sVal = DeclineGendFem
        Case "cр":      sKey = "NewGend": sVal = DeclineGendNeut
        Case "одуш":    sKey = "Animate": sVal = True
        Case "фио":     sKey = "IsFio":   sVal = True
        Case Else:      sKey = c_idxPref & Par: sVal = Par ' прочие просто добавляем как есть может они зачем-то нужны
        'Case Else:      GoTo HandleNext                    ' неизвестные пропускаем
        End Select
    ' добавляем параметр в коллекцию
        If p_IsExist(sKey, cPar) Then If Not ReplaceExisting Then GoTo HandleNext Else cPar.Remove sKey
        cPar.Add sVal, sKey
HandleNext: Next Par: Set p_TermModifyParGet = cPar
End Function

Public Function GroupsGet(Source As String, _
    ByRef cGroups As Collection, _
    Optional UsePlaceHolders As Boolean = False, _
    Optional Templates, Optional TermDelim = "@", Optional TempDelim = ";", _
    Optional aGroups) As Boolean
' возвращает коллекцию групп содержащихся в строке (выражений заключенных в скобки)
'-------------------------
' Source    - выражение содержащее скобки
' cGroups   - (возвращаемое) коллекция содержимого скобок индекс элемента соответствует уровню скобки в порядке разбора
'             именованные индексы соответствуют порядковому номеру с префиксом Br
'             коллекция нужна для возможности использования результата соместно с функциями PlaceHoldersGet/Set
' UsePlaceHolders - если True в Text, будут возвращены выражения содержащее подстановочные ссылки на соотв элементы массива результата
'             вида: ([%1%])+([%2%]), где 1,2.. - индексы элементов коллекции cGroups хранящей содержимое скобок
'             иначе - полное текстовое выражение содержащееся в скобках.
' Templates - строка или массив строк содержащий шаблоны допустимых групп
'             т.к. скобки проверяются прямым перебором слева-направо, для более-менее корректной работы
'             надо чтобы шаблоны были упорядочены по мере усложнения в порядке возможного срабатывания
' TermDelim - разделитель элементов (замещающий символ для обозначения извлекаемого элемента группы) в строке шаблона
' TempDelim - разделитель шаблонов в строке
' aGroups   - (возвращаемое) массив позиций элементов строки (уровней групп/границ групп/содержимого групп) нужно только если хотите отслеживать позиции в строке
' возвращает: True  - если выражение успешно разобрано,
'             False - если выражение содержит незакрытые скобки или не корректно
'-------------------------
' v.1.0.2       : 12.03.2024 - первая попытка переделать скобки под шаблоны
' v.1.0.1       : 21.12.2022 - исправлены многочисленные ошибки. (всё еще сильно экспериментальная)
' v.1.0.0       : 24.03.2020 - исходная (очень кривая и глючная) версия
'-------------------------
' Примеры:
' 1) strText = "Do: If True Then 1 Else 0 End If: Loop"
'    strTemp = "If @ Then @ Else @ End If;Do: @: Loop"
'    Call GroupsGet(strText, cGroup, True, strTemp)
' 2) strText = "((5+2)+3*(4+5)^4)-97"
'    Call GroupsGet(strText, cGroup, True)
'-------------------------
' ToDo: - алгоритм срабатывает на первый подходящий шаблон, но элементы могут встречаться в разных шаблонах и правильным может оказаться не первый, - надо предусмотреть
'         возможные решения:
'           1) "ошибка-возврат" - при разборе по сработавшему шаблону после ошибки возвращаемся в стеке к началу выражения в строке и продолжаем разбор со следующего шаблона
'               минус - ногократные проверки одних и тех же фрагментов
'           2) "выбраковка" - при разборе формируем коллекцию всех сработавших на фрагмент шаблонов и вычёркиваем их по мере продвижения по шаблону
'               минус - придется в стек загонять коллекции из пар номер шаблона/элемента
'           3) "упорядочивание" - перед началом работы расположить шаблоны по мере их усложнения в порядке возможного срабатывания
'               минус - непонятно по каким критериям (отобрать начинающиеся одинаково, расположить по возростанию длины элемента, .. ещё???)
'               и не факт что это возможно для конкретного набора шаблонов, возможны неоднозначности из-за повторов и необязательных элементов
'           думаю оптимально будет гибрид из 3 и 1.
'           сначала элементарное упорядочивание правил (до разбора), чтобы уменьшить количество возвратов,
'           а то что не исключили сортировкой отловим повторами при разборе
'       - распознавание повторяющихся и необязательных элементов в шаблонах например (@[,@]) м.б. (@);(@,@),(@,@,@) и т.п.
'-------------------------
#Const TestErr = False          ' проверять ошибку несогласованных скобок
Const cPref = "Br"              ' префикс именованного элемента коллекции
Const errUnclosedExp = vbObjectError + 511
Const errIncompleteExp = vbObjectError + 512
Dim Result As Boolean ': Result = False
On Error GoTo HandleError
    If Len(Source) = 0 Then Result = True: GoTo HandleExit
' задаем допустимые шаблоны групп (скобок)
Dim sTerm, sName As String
Dim aTemp                       ' массив массивов элементов шаблона
Dim t As Long                   ' индекс шаблона в массиве
    If IsMissing(Templates) Then
' не задано - берём набор скобок по-умолчанию
        aTemp = Array(Array("(", ")"), Array("[", "]"), Array("{", "}"), Array("<", ">"), Array("%", "%"), Array("'", "'"), Array("""", """"))
    ElseIf IsArray(Templates) Then
' задано одномерным массивом (не проверяется)
        ReDim aTemp(LBound(aTemp), UBound(aTemp)): For t = LBound(aTemp) To UBound(aTemp): aTemp(t) = Split(Templates(t), TermDelim): Next t
    Else
' задано строкой
        ReDim aTemp(0 To 0): For Each sTerm In Split(Templates, TempDelim): ReDim Preserve aTemp(0 To t): aTemp(t) = Split(sTerm, TermDelim): t = t + 1: Next sTerm
    End If
Dim l As Long           ' индекс проверяемого элемента шаблона = LBound(aTemp(t)) - открывающая скобка, =UBound(aTemp(t)) - закрывающая скобка, остальное - промежуточные скобки
Dim i As Long           ' позиция символа в разбираемой строке
Dim j As Long           ' индекс стека
Dim g As Long           ' индекс элемента массива для хранения результата разбора
Dim aStack() As Long    ' имитируем стек скобок
Const sStep = 3         ' шаг элементов массива стека
                ' +0    '(.TempNum) номер шаблона по массиву
                ' +1    '(.TempItm) номер элемента шаблона
                ' +2    '(.TempBeg) позиция начала подстроки в исходной строке (включая скобки)
'Dim aGroups() As Long ' массив для хранения результата разбора
Const gStep = 7         ' шаг элементов массива позиций элементов
                ' +0    '(.TextLev) уровень вложенности (0-вне скобок, 1-внешние скобки, ... n-скобки n-уровня)
                ' +1    '(.TextBeg) позиция начала текста подстроки в исходной строке (после открывающей скобки)
                ' +2    '(.TextEnd) позиция конца текста подстроки в исходной строке (до закрывающей скобки)
                ' +3    '(.TempBeg) позиция начала подстроки в исходной строке (включая скобки)
                ' +4    '(.TempEnd) позиция конца подстроки в исходной строке (включая скобки)
                ' +5    '(.TempNum) номер шаблона по массиву
                ' +6    '(.TempItm) номер элемента в шаблоне
Dim iBeg As Long, iEnd As Long, iLen As Long
    ReDim aGroups(1 To 1) As Long
    Set cGroups = New Collection
' ищем парные скобки в строке используя стек
    i = 1
    Do Until i > Len(Source)
' проверяем символы в текущей позиции
        iLen = 1        ' просматриваем строку посимвольно
        If j > 0 Then
    ' если в стеке есть незакрытые скобки
            ' проверяем следующий элемент для шаблона с вершины стека
            t = aStack(j - 2)                       '(.TempNum) номер шаблона по массиву
            l = aStack(j - 1)                       '(.TempItm) номер элемента шаблона
            sTerm = aTemp(t)(l + 1)
            If sTerm = Mid$(Source, i, Len(sTerm)) Then
            ' если совпадает - извлекаем фрагмент строки в результат
            ' заносим фрагмент в результат
                g = g + gStep: ReDim Preserve aGroups(1 To g) 'As Long
                iLen = Len(aTemp(t)(l))             ' длина подстроки предыдущего элемента шаблона
                iBeg = aStack(j - 0)                ' позиция начала подстроки предыдущего элемента шаблона в исходной строке
                iEnd = iBeg + iLen                  ' позиция конца подстроки предыдущего элемента шаблона в исходной строке
                iLen = Len(sTerm)                   ' длина подстроки текущего элемента шаблона
                l = l + 1                           ' переходим к следующему элементу текщего шаблона
                
                aGroups(g - 6) = j \ sStep        '(.TextLev) уровень вложенности (0-вне скобок, 1-внешние скобки, ... n-скобки n-уровня)
                aGroups(g - 5) = iEnd             '(.TextBeg) позиция начала текста подстроки в исходной строке (после открывающей скобки)
                aGroups(g - 4) = i                '(.TextEnd) позиция конца текста подстроки в исходной строке (до закрывающей скобки)
                aGroups(g - 3) = iBeg             '(.TempBeg) позиция начала подстроки в исходной строке (включая скобки)
                aGroups(g - 2) = i + iLen         '(.TempEnd) позиция конца подстроки в исходной строке (включая скобки)
                aGroups(g - 1) = t                '(.TempNum) номер шаблона по массиву
                aGroups(g - 0) = l                '(.TempItm) номер элемента в шаблоне
            ' заносим содержимое скобки в результирующую коллекцию
                sTerm = Mid$(Source, iEnd, i - iEnd)                    ' очищенный от внешних скобок фрагмент
                sName = cPref & (g \ gStep): cGroups.Add sTerm, sName   ' добавляем в коллекцию
                If l = UBound(aTemp(t)) Then
            ' если текущий элемент закрывающий - уменьшаем вершину стека
                    j = j - sStep: If j > 0 Then ReDim Preserve aStack(1 To j) Else Erase aStack   ' уменьшаем стек
                Else
            ' иначе увеличиваем в стеке уровень элемента шаблона и его позицию
                    aStack(j - 1) = l               '(.TempLev) номер элемента шаблона
                    aStack(j - 0) = i               '(.TempBeg) позиция начала подстроки в исходной строке (включая скобки)
                End If
                GoTo HandleNextSym                  ' фрагмент найден и разобран - переход к следующему символу
            End If
        End If
' проверяем первый (открывающий) элемент всех шаблонов
        For t = LBound(aTemp) To UBound(aTemp)
            l = LBound(aTemp(t)): sTerm = aTemp(t)(l)
            If sTerm = Mid(Source, i, Len(sTerm)) Then
            ' если открывающий элемент совпадает с текущим фрагментом строки
                If l < UBound(aTemp(t)) Then
                ' если текущий элемент не закрывающий (простые разделители сразу и открывающие и закрывающие) - заносим его в стек
                    j = j + sStep: ReDim Preserve aStack(1 To j) ' увеличиваем стек
                    aStack(j - 2) = t               '(.TempNum) номер шаблона по массиву
                    aStack(j - 1) = l               '(.TempItm) номер элемента шаблона
                    aStack(j - 0) = i               '(.TempBeg) позиция начала подстроки в исходной строке (включая скобки)
                End If
                iLen = Len(sTerm)                   ' смещаем позицию в строке на длину найденного фрагмента
                GoTo HandleNextSym ': Exit For      ' фрагмент найден и разобран - переход к следующему символу
            End If
        Next t
#If TestErr Then
' можно дополнительно выполнить проверку незакрытых скобок - проверить соответствие все м остальным скобкам не являющимся открывающими и вернуть позицию ошибки
        For t = LBound(aTemp) To UBound(aTemp)
            For l = LBound(aTemp(t)) + 1 To UBound(aTemp(t))
            sTerm = aTemp(t)(l): If sTerm = Mid$(Source, i, Len(sTerm)) Then sName = aTemp(t)(l - 1): Err.Raise errIncompleteExp
        Next l: Next t
#End If
' текущий фрагмент строки - содержимое скобки - просто переходим к следующему символу
HandleNextSym: i = i + iLen    ' смещаем указатель в строке на следующий после проанализированного символ
    Loop
'#If TestErr Then
'' проверяем наличие в стеке незакрытых скобок
    If j <> 0 Then sTerm = Join(aTemp(aStack(j - 2)), "..."): i = aStack(j): Err.Raise errUnclosedExp
'#End If
    Erase aStack
' если UsePlaceHolders=False необходимости добавлять внешнее выражение в результат нет,
' т.к. оно совпадает с Source, но для единообразия - сделаем.
    ' добавляем внешний уровень массива результата для исходной строки
' заносим фрагмент в результат
    g = g + gStep: ReDim Preserve aGroups(1 To g) 'As tTerm
    'aGroups(g - 6) = 0                '(.TextLev) уровень вложенности (0-вне скобок)
    aGroups(g - 5) = 1                '(.TextBeg) позиция начала текста подстроки в исходной строке (после открывающей скобки)
    aGroups(g - 4) = Len(Source) + 1  '(.TextEnd) позиция конца текста подстроки в исходной строке (до закрывающей скобки)
    aGroups(g - 3) = 1                '(.TempBeg) позиция начала подстроки в исходной строке (включая скобки)
    aGroups(g - 2) = Len(Source) + 1  '(.TempEnd) позиция конца подстроки в исходной строке (включая скобки)
    aGroups(g - 1) = -1               '(.TempNum) номер шаблона по массиву
    aGroups(g - 0) = -1               '(.TempItm) номер элемента в шаблоне
    sTerm = Source                      '
    sName = cPref & (g \ gStep): cGroups.Add sTerm, sName     ' добавляем в коллекцию
    Result = True:
' если надо создавать шаблоны разбора выражений в скобках - делаем это
    If UsePlaceHolders Then Call p_GroupsPlaceHoldersSet(cGroups, aGroups)
HandleExit:     GroupsGet = Result: Exit Function
HandleError:    Select Case Err
    Case errUnclosedExp:    Debug.Print "Ошибка! Незавершённое выражение """ & sTerm & """ в позиции " & i & " в строке: """ & Source & """"
    Case errIncompleteExp:  Debug.Print "Ошибка! """ & sTerm & """ без """ & sName & """ в позиции " & i & " в строке: """ & Source & """"
    Case Else: Stop: Resume 0
    End Select
    Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_GroupsPlaceHoldersSet(ByRef cGroups As Collection, aGroups) As Boolean
' пересортирует коллекцию элементов разобранных групп заменяя элементы символами подстановки
Dim Result As Boolean ': Result = False
' вынесено в отдельную функцию для упрощения читаемости
On Error GoTo HandleError
Const cLBr = "[%", cRBr = "%]"  ' скобки для ссылок на элементы массива результата.
Const cPref = "Br"              ' префикс именованного элемента коллекции
Const gStep = 7                 ' шаг элементов массива позиций элементов
Dim g As Long                   ' индекс элемента массива для хранения результата разбора
Dim i As Long, j As Long, iMin As Long
Dim iBeg As Long, iEnd As Long
Dim jBeg As Long, jEnd As Long
Dim iLvl As Long, jLvl As Long
Dim sTerm, sName As String
    i = UBound(aGroups) \ gStep: iMin = 1
    Do While i > iMin 'For i = i To 2 Step -1
' разбираем все нижестоящие элементы массива в обратном порядке
    ' элементы в массиве отсортированы так что наружные (те у которых Level меньше) будут выше
    ' проверяем - если границы проверяемого (j) элемента лежат внутри разбираемого (i)
    ' заменяем содержимое проверяемого элемента в разбираемом на символьный указатель
    ' смещаем позицию границы проверки в разбираемом элементе до границ неразобранного фрагмента
            j = iMin
        ' текст фрагмента до разбора
            sName = cPref & i
            sTerm = cGroups(sName)
        ' уровень и границы разбираемого фрагмента в исходной строке
            g = (i - 1) * gStep + 1
            iLvl = aGroups(g + 0)
            iBeg = aGroups(g + 1)
            iEnd = aGroups(g + 2)
        Do While j < i 'For j = 1 To i - 1
    ' проверяем все нижестоящие элементы массива в прямом порядке по разбираемому
        ' уровень и границы разбираемого фрагмента в исходной строке
            g = (j - 1) * gStep + 1
        ' проверяемый элемент должен принадлежать предыдущему уровню вложенности относительно разбираемого
            jLvl = aGroups(g + 0): If iLvl <> (jLvl - 1) Then GoTo HandleNextJ    'iLvl > jLvl -> Next j
        ' границы проверяемого элемента должены лежать в пределах границ разбираемого
            jBeg = aGroups(g + 1): If iBeg > jBeg Then GoTo HandleNextJ           'iBeg > jBeg -> Next j
            jEnd = aGroups(g + 2): If iEnd < jEnd Then GoTo HandleNextJ           'iEnd < jEnd -> Next j
        ' если прямой перебор j
            ' нужно считать от конца строки (начало фрагмента меняется)
            sTerm = Left$(sTerm, Len(sTerm) - (iEnd - jBeg)) & cLBr & j & cRBr & Right$(sTerm, iEnd - jEnd)
        ' и сдвигать нижнюю границу просматриваемого фрагмента
            iBeg = aGroups(g + 3) + 1     ' смещаем позицию на начало неразобранного фрагмента на позицию после найденной
            If j = iMin Then iMin = j + 1   ' смещаем указатель нижней границы проверяемых элементов массива
                                            ' (найденный нижний элемент уже не встретится нет смысла его проверять снова)
        '' если обратный перебор j
        '    ' нужно считать от начала строки (конец фрагмента меняется)
        '    ??? 'sTerm = Left$(sTerm, (iEnd - jBeg + 1)) & cLBr & j & cRBr & Right$(sTerm, Len(sTerm) - iBeg - jEnd)
        '    ' и сдвигать верхнюю границу просматриваемого фрагмента
        '    iEnd = aGroups(g + 4) - 1      ' смещаем позицию на конец неразобранного фрагмента на позицию до найденной
        '    ??? 'If j = iMin Then iMin = j   '
        ' проверяем полностью ли разобран фрагмент
            If iEnd <= iBeg Then Exit Do   ' разбор текущего фрагмента окончен
            If i = iMin Then iMin = i + 1   ' ??? смещаем указатель нижней границы проверяемых элементов массива
                                            ' (полностью разобранный нижний элемент уже не встретится нет смысла его проверять снова)
HandleNextJ: j = j + 1
    Loop 'Next j
    ' возвращаем результат разбора в коллекцию
        With cGroups: .Remove sName: .Add sTerm, sName, After:=i - 1: End With
HandleNextI: i = i - 1
    Loop 'Next i
HandleExit:  p_GroupsPlaceHoldersSet = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function GroupText(Source As String, idx As Long, _
    Optional Templates, Optional TermDelim = "@", Optional TempDelim = ";" _
    ) As String
' возвращает содержимое группы символов из строки (выражение заключенное в скобки)
'-------------------------
' Source    - выражение содержащее скобки
' Idx       - индекс группы содержимое которой необходимо
' Templates - строка или массив строк содержащий шаблоны допустимых групп
'             т.к. скобки проверяются прямым перебором составные скобки надо ставить вначале
' TermDelim  - разделитель элементов (замещающий символ для обозначения извлекаемого элемента группы) в строке шаблона
' TempDelim  - разделитель шаблонов в строке
'-------------------------
On Error GoTo HandleError
Dim cGroups As Collection: Call GroupsGet(Source, cGroups, , Templates, TermDelim, TempDelim)    'If Not .. Then Err.Raise vbObjectError + 512 'Exit Function
    GroupText = cGroups(idx) ' Техт
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function InStrLike(Start As Long, _
    String1 As String, String2 As String, _
    Optional Found As String, _
    Optional Compare As VbCompareMethod = vbTextCompare) As Long
' InStr позволяющий искать совпадения по маске выражений Like
'-------------------------
' Start     - начальная позиция
' String1   - строка в которой производим поиск
' String2   - строка содержащая строку маски поиска
' Found     - (возвращаемое) найденая по маске подстрока
' Compare   - способ сравнения
' возвращает позицию первого вхождения String2 в String1 начиная с позиции Start
'-------------------------
' v.1.0.1       : 05.02.2020 - переписал для лучшего понимания и компактности кода
' v.1.0.0       : 23.02.2003 - original by VictorB212 from http://www.vbforums.com/showthread.php?232259-InStrLike-(debugging-help-required)
'-------------------------
Const cSymBeg = "^" ' символ привязки к началу (как в RegEx) - бессмысленен т.к результат всегда будет Start или 0
Const cSymEnd = "$" ' символ привязки к концу
Dim Result As Long: Result = False
    On Error GoTo HandleError
    If Start <= 0 Then Start = 1
    Found = vbNullString
Dim S1 As String, S2 As String
    S1 = Mid$(String1, Start): S2 = String2
' формируем шаблон с учетом якорей
Dim bBeg As Boolean: bBeg = Left$(String2, Len(cSymBeg)) = cSymBeg:  If bBeg Then S2 = Mid$(S2, Len(cSymBeg) + 1) Else S2 = "*" & S2
Dim bEnd As Boolean: bEnd = Right$(String2, Len(cSymEnd)) = cSymEnd: If bEnd Then S2 = Left$(S2, Len(S2) - Len(cSymBeg)) Else S2 = S2 & "*"
    If Compare = vbTextCompare Then S1 = UCase$(S1): S2 = UCase$(S2)
' предварительная проверка соответствия строки шаблону
Dim iSgn As Integer: iSgn = S1 Like S2: If Not iSgn Then GoTo HandleExit
Dim lLen As Long, lPos As Long
' ищем правую границу
HandleRightBound:
    lLen = Len(S1): lPos = lLen
    If bEnd Then GoTo HandleLeftBound           ' если привязка к правому краю - правая граница известна - ищем левую
    Do
        If Not iSgn Then iSgn = 1 Else If Not (Left$(S1, lPos - 1) Like S2) Then Exit Do
        lLen = lLen \ 2: If lLen < 1 Then lLen = 1  ' делим диапазон пополам
        lPos = lPos + iSgn * lLen                   ' уменьшаем границу если совпадает, иначе - увеличиваем
        iSgn = Left$(S1, lPos) Like S2              ' проверяем совпадение
    Loop
' ищем левую границу
HandleLeftBound:
    S1 = Left$(S1, lPos)                            ' обрезаем справа по найденой границе
    lLen = Len(S1)
    If bBeg Then lPos = 0: GoTo HandleResult        ' если привязка к левому краю - левая граница известна - получаем результат
    lPos = lLen                     '
    Do
        If Not iSgn Then iSgn = 1 Else If Not (Right$(S1, lPos - 1) Like S2) Then Exit Do
        lLen = lLen \ 2: If lLen < 1 Then lLen = 1  ' делим диапазон пополам
        lPos = lPos + iSgn * lLen                   ' уменьшаем границу если совпадает, иначе - увеличиваем
        iSgn = Right$(S1, lPos) Like S2             ' проверяем совпадение
    Loop
    lLen = lPos: lPos = Len(S1) - lPos
' возвращаем результат
HandleResult:
    Result = Start + lPos: If Result Then Found = Mid$(String1, Result, lLen)
HandleExit:  InStrLike = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function InStrRegEx( _
    Start As Long, _
    String1 As String, String2 As String, _
    Optional Found As String, _
    Optional Compare As VbCompareMethod = vbTextCompare) As Long
' InStr позволяющий искать совпадения по маске выражений RegEx
'-------------------------
' Start     - начальная позиция
' String1   - строка в которой производим поиск
' String2   - строка содержащая строку маски поиска
' Found     - (возвращаемое) найденая по маске подстрока
' возвращает позицию первого вхождения String2 в String1 начиная с позиции Start
'-------------------------
' v.1.0.0       : 25.08.2010 - original by BC_Programmer https://www.computerhope.com/forum/index.php?topic=109171.msg736986#msg736986
'-------------------------
Dim Result As Long: Result = False
    On Error GoTo HandleError
    If Start <= 0 Then Start = 1
    Found = vbNullString
Dim S1 As String: S1 = Mid$(String1, Start)   'shortened version of String1
Dim oMatches As Object, oMatch As Object
    ' вызов RegExp и передача ему маски
    With RegEx: .IgnoreCase = (Compare = vbTextCompare): .Pattern = String2: Set oMatches = .Execute(S1): End With
    If oMatches.Count = 0 Then GoTo HandleError
    For Each oMatch In oMatches
        Result = oMatch.FirstIndex + Start: Found = oMatch
        Exit For
    Next
HandleExit:  InStrRegEx = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function InStrAll( _
    Start As Long, _
    String1 As String, String2 As String, _
    Optional Found, _
    Optional Method As Integer = 0, _
    Optional Compare As VbCompareMethod = vbTextCompare)
' InStr позволяющий искать все совпадения по подстроке
Const c_strProcedure = "InStrAll"
' Start     - начальная позиция
' String1   - строка в которой производим поиск
' String2   - строка содержащая строку маски поиска
' Found     - (возвращаемое) массив найденных по маске подстрок
' Method    - способ сравнения подстрок шаблона
'   0 - по простой подстроке (InStr)
'   1 - по Like подстроке (InStrLike)
'   2 - по RegEx выражению (InStrRegEx)
' возвращает массив позиций вхождения String2 в String1 начиная с позиции Start
'-------------------------
' v.1.0.0       : 16.07.2020 -
'-------------------------
    On Error GoTo HandleError
    If Start <= 0 Then Start = 1
Dim sFound As String
Dim bFound As Boolean: bFound = Not IsMissing(Found)
Dim S1 As String: S1 = Mid$(String1, Start)   'shortened version of String1
Dim lPos As Long, lLen As Long: lLen = Len(S1): lPos = 1
Dim aResult() As Long, aFound() As String
'Dim cResult As New Collection, cFound As New Collection
Dim i As Long: i = 0
    Do While lPos <= lLen
        Select Case Method
        Case 1:     lPos = InStrLike(lPos, S1, String2, sFound)
        Case 2:     lPos = InStrRegEx(lPos, S1, String2, sFound)
        Case Else:  lPos = InStr(lPos, S1, String2): sFound = String2
        End Select
        If lPos = 0 Then Exit Do
        i = i + 1
        ReDim Preserve aResult(1 To i): aResult(i) = lPos: If bFound Then ReDim Preserve aFound(1 To i): aFound(i) = sFound
        'cResult.Add lPos: If bFound Then cFound.Add sFound
        lPos = lPos + Len(sFound)
    Loop
'HandleExit:  Set InStrAll = cResult: If bFound Then Set Found = cFound: Exit Function
HandleExit:  InStrAll = aResult: If bFound Then Found = aFound: Exit Function
HandleError: Erase aResult: Err.Clear: Resume HandleExit
End Function
Public Static Function InStrCount( _
    ByRef Text As String, _
    ByRef Find As String, _
    Optional ByVal Start As Long = 1, _
    Optional ByVal Compare As VbCompareMethod = vbBinaryCompare _
    ) As Long
' возвращает количество подстрок в строке
Const c_strProcedure = "InStrCount"
' Text - текст в котором производится поиск
' Find - искомая подстрока подсчёт количества вхождений которой производится
' Start - начальная позиция поиска
' Compare - тип сравнения
'-------------------------
' v.1.0.1       : 21.11.2001 - original InStrCount04 by Jost Schwider, jost@schwider.de from http://www.xbeat.net/vbspeed/c_InStrCount.htm#InStrCount04
'-------------------------
Const MODEMARGIN = 8
Dim TextAsc() As Integer
Dim TextData As LongPtr, TextPtr As LongPtr
Dim FindAsc(0 To MODEMARGIN) As Integer
Dim FindLen As Long
Dim FindChar1 As Integer
Dim FindChar2 As Integer
Dim i As Long

    If Compare = vbBinaryCompare Then
' двоичное сравнение
        FindLen = Len(Find)
        If FindLen Then
    ' проверка первого совпадения
            If Start < 2 Then Start = InStrB(Text, Find) Else Start = InStrB(Start + Start - 1, Text, Find)
            If Start Then
        ' ищем последующие вхождения
                InStrCount = 1
                If FindLen <= MODEMARGIN Then
            ' для длин искомой подстроки до MODEMARGIN - быстрый способ
                ' подготовка текстового массива
                    If TextPtr = 0 Then ReDim TextAsc(1 To 1): TextData = VarPtr(TextAsc(1)):
                    CopyMemory TextPtr, ByVal VarPtrArray(TextAsc), PTR_LENGTH: TextPtr = TextPtr + 8 + PTR_LENGTH
                ' инициализация массива
                    CopyMemory ByVal TextPtr, ByVal VarPtr(Text), PTR_LENGTH            'pvData
                    CopyMemory ByVal TextPtr + PTR_LENGTH, Len(Text), 4 ' PTR_LENGTH    'nElements
                    Select Case FindLen
                    Case 1 ' в буфере один знак
                        FindChar1 = AscW(Find)
                        For Start = Start \ 2 + 2 To Len(Text)
                            If TextAsc(Start) = FindChar1 Then InStrCount = InStrCount + 1
                        Next Start
                    Case 2 ' в буфере два знака
                        FindChar1 = AscW(Find): FindChar2 = AscW(Right$(Find, 1))
                        For Start = Start \ 2 + 3 To Len(Text) - 1
                            If TextAsc(Start) = FindChar1 Then
                                If TextAsc(Start + 1) = FindChar2 Then
                                    InStrCount = InStrCount + 1: Start = Start + 1
                                End If
                            End If
                        Next Start
                    Case Else ' в буфере больше двух знаков
                        CopyMemory ByVal VarPtr(FindAsc(0)), ByVal StrPtr(Find), FindLen + FindLen
                        FindLen = FindLen - 1
                        ' первые два знака
                        FindChar1 = FindAsc(0): FindChar2 = FindAsc(1)
                        For Start = Start \ 2 + 2 + FindLen To Len(Text) - FindLen
                            If TextAsc(Start) = FindChar1 Then
                                If TextAsc(Start + 1) = FindChar2 Then
                                    For i = 2 To FindLen
                                        If TextAsc(Start + i) <> FindAsc(i) Then Exit For
                                    Next i
                                    If i > FindLen Then
                                        InStrCount = InStrCount + 1: Start = Start + FindLen
                                    End If
                                End If
                            End If
                        Next Start
                    End Select
                ' восстанавливаем значения из массива
                    CopyMemory ByVal TextPtr, TextData, PTR_LENGTH 'pvData
                    CopyMemory ByVal TextPtr + PTR_LENGTH, 1&, 4 'PTR_LENGTH  'nElements
                Else
            ' для больших длин - обычный способ
                    FindLen = FindLen + FindLen
                    Start = InStrB(Start + FindLen, Text, Find)
                    Do While Start
                        InStrCount = InStrCount + 1
                        Start = InStrB(Start + FindLen, Text, Find)
                    Loop
                End If 'FindLen <= MODEMARGIN
            End If 'Start
        End If 'FindLen
    Else
' текстовое сравнение
    ' игнорируем верхний регистр
        InStrCount = InStrCount(LCase$(Text), LCase$(Find), Start)
    End If
End Function
Public Function Replicate(ByVal Number As Long, Pattern As String) As String
' Returns a pattern replicated in a string a specified number of times.
'-------------------------
' Number  - Required. Number of replications desired.
' Pattern - Required. Character pattern to replicate.
'-------------------------
' v.1.0.2       : 06.12.2000 - original Replicate05 by Donald, donald@xbeat.net from http://www.xbeat.net/vbspeed/c_Replicate.htm#Replicate05
'-------------------------
Dim lp As Long
    If Number > 0 Then
        lp = Len(Pattern)
        Select Case lp
        Case Is > 1:    Replicate = Space$(Number * lp): Mid$(Replicate, 1, lp) = Pattern: If Number > 1 Then Mid$(Replicate, lp + 1) = Replicate
        Case 1:         Replicate = String$(Number, Pattern)
        End Select
    End If
End Function
Public Function WordWrap( _
    ByRef Text As String, _
    ByVal Width As Long, _
    Optional ByRef CountLines As Long) As String
' разбивает строку на заданное количество символов с использованием символа разрыва строки.
'-------------------------
' Text  - разбиваемая строка
' Width - длина строки в символах
'-------------------------
' v.1.0.0       : 13.09.2004 - original WordWrap01 by Donald, donald@xbeat.net from http://www.xbeat.net/vbspeed/c_WordWrap.htm#WordWrap01
'-------------------------
Dim i As Long, lenLine As Long
Dim posBreak As Long, cntBreakChars As Long
Dim abText() As Byte, abTextOut() As Byte
Dim ubText As Long

    If Width <= 0 Then CountLines = 0: Exit Function
    If Len(Text) <= Width Then CountLines = 1: WordWrap = Text: Exit Function
    abText = StrConv(Text, vbFromUnicode): ubText = UBound(abText)
    ReDim abTextOut(ubText * 3)     'dim to potential max
    For i = 0 To ubText
        Select Case abText(i)
        Case 32, 45: posBreak = i   ' пробел и перенос
        Case Else
        End Select
        abTextOut(i + cntBreakChars) = abText(i)
        lenLine = lenLine + 1
        If lenLine > Width Then
            If posBreak > 0 Then
                If posBreak = ubText Then Exit For ' don't break at the very end
                ' разрыв после пробела или переноса
                abTextOut(posBreak + cntBreakChars + 1) = 13  'CR
                abTextOut(posBreak + cntBreakChars + 2) = 10  'LF
                i = posBreak: posBreak = 0
            Else ' cut word
                abTextOut(i + cntBreakChars) = 13     'CR
                abTextOut(i + cntBreakChars + 1) = 10 'LF
                i = i - 1
            End If
            cntBreakChars = cntBreakChars + 2: lenLine = 0
        End If
    Next
    CountLines = cntBreakChars \ 2 + 1
    ReDim Preserve abTextOut(ubText + cntBreakChars)
    WordWrap = StrConv(abTextOut, vbUnicode)
End Function
Public Function Compress( _
    sExpression As String, _
    sCompress As String, _
    Optional Compare As VbCompareMethod = vbBinaryCompare) As String
' Returns a string where multiple adjacent occurrences of a specified substring are compressed to just one occurrence. For example, the function will compress multiple spaces within a string down to single spaces.
'-------------------------
' sExpression - Required. String expression containing substring sequences to be compressed. If sExpression is a zero-length string, Compress returns a zero-length string as well.
' sCompress - Required. The single string whereof sequences are to be compressed. Read sCompress as "compress multiples of". Can be longer than one char. If sCompress is a zero-length string, the function returns sExpression unmodified.
' Compare - Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. If omitted, the default value is 0, which means a binary comparison is performed.
'-------------------------
' v.1.0.0       : 04.11.2001 - original Compress06 by Tom Winters, tom@interplanetary.freeserve.co.uk from http://www.xbeat.net/vbspeed/c_Compress.htm#Compress06
'-------------------------
Dim sExp$, sFind$, lLenCompress&, lLenExpression&
Dim lChrPosition&

    lLenExpression = Len(sExpression)
    If lLenExpression = 0 Then Exit Function
    lLenCompress = Len(sCompress)
    If lLenCompress <> 0 Then
        If lLenCompress = 1 Then
    ' искомая строка состоиит из одного символа
            If lLenExpression < 10 Then
        ' сжимаемая строка короткая
                sFind = sCompress + sCompress ' ищем два одинаковых символа подряд
                Compress = sExpression
                ' повтоторяем пока есть совпадения
                ' этот метод хорошо работает на коротких строках
                Do
                    lChrPosition = InStr(1, Compress, sFind, Compare)
                    If lChrPosition = 0 Then Exit Function
                    sExp = Left$(Compress, lChrPosition)
                    Compress = Right$(Compress, Len(Compress) - Len(sExp) - lLenCompress)
                    Compress = sExp + Compress
                Loop
            Else
    ' для длинных строк поиска
        ' проверяем повторения строки поиска
            ' Ideally we'd check the the entire string for segment matches,
            ' but if we do that we'll be here for ever
            ' So, we'll use a reasonable compromise..
        ' проверка первых 12 символов строки  дает 2/3 совпадений
            Dim sNewSearchString As String
                sExp = Left$(sExpression, 12)
                sNewSearchString = String$(8, sCompress)
                lChrPosition = InStr(1, sExp, sNewSearchString, Compare)
                ' если у нас длинная повторяющаяся строка поиска
                If lChrPosition > 0 Then
        ' сжимаем односимвольные повторяющиеся строки в длинном выражении
                Dim lLenNewSearchString As Long, lLenFind2 As Long, lStringSizeCounter As Long
                    lLenFind2 = lLenCompress + lLenCompress
                    lStringSizeCounter = (lLenExpression - lLenFind2)
                    ' Make new search string divisible by 2
                    lStringSizeCounter = lStringSizeCounter + (lStringSizeCounter And 1)
                    ' создаем новую строку поиска
                    sNewSearchString = String$(lStringSizeCounter, sCompress)
                    lLenNewSearchString = Len(sNewSearchString)
                    lStringSizeCounter = 0
                    Compress = sExpression
                    sFind = sCompress + sCompress
                ' если мы ищем длинную строку быстрее искать сначала большие последовательности, затем - меньшие
                ' поэтому для больших строк будем производить поиск постепенно уменьшая длину последовательности пока совпадение не будет найдено
                ' этот метод показывает лучшую производительность с длинными последовательностями
                    Do
                        Do
                            lChrPosition = InStr(1, Compress, sNewSearchString, Compare)
                            If lChrPosition = 0 Then Exit Do
                            sExp = Left$(Compress, lChrPosition)
                            Compress = Right$(Compress, Len(Compress) - Len(sExp) - lLenNewSearchString + lLenCompress)
                            Compress = sExp + Compress
                            ' Reset quantity removed from search string
                            lStringSizeCounter = 0
                        Loop
                        lChrPosition = InStr(1, Compress, sFind, Compare)
                        If lChrPosition = 0 Then Exit Function
                        ' Increase size of search string counter
                        lStringSizeCounter = lStringSizeCounter + lLenCompress
                        ' Make search string smaller and try again
                        sNewSearchString = Right$(sNewSearchString, Len(Compress) - lStringSizeCounter)
                        lLenNewSearchString = Len(sNewSearchString)
                    Loop
                End If
            End If
        End If
' проверяем Unicode
    ' We can save a lot of work by not passing Unicode strings to our byte arrays
    ' so we test for Unicode first and if true, use another Instring method
    Dim lCharacter As Long, lAsciiValue As Long
        For lCharacter = 1 To lLenCompress
            lAsciiValue = Asc(Mid$(sCompress, lCharacter, 1))
            If lAsciiValue > 127 Then
                ' If we have a Unicode multiple search character
                Dim bGo As Boolean, lPosition&
                sExp = sExpression
                ' This method is fast with more complex expressions
                Do While Len(sExp) > 0
                    bGo = False
                    lPosition = InStr(1, sExp, sCompress, Compare)
                    If Mid$(sExp, lPosition + lLenCompress, lLenCompress) = sCompress Then
                        If lPosition = 1 Then
                            bGo = True
                        End If
                    End If
                    If bGo Then
                        sExp = Right$(sExp, Len(sExp) - lLenCompress)
                    Else
                        Compress = Compress + Left$(sExp, 1)
                        sExp = Right$(sExp, Len(sExp) - 1)
                    End If
                Loop
                Exit Function
            End If
        Next
' сжатие с использованием байтового массива
    Dim bMatch As Boolean, bMatchResult1 As Boolean, bMatchResult2 As Boolean
    Dim lLenExpressionArray&, lLenCompressArray&, lbytePosition&, lNewCounter&
    Dim byExpressionArray() As Byte, byNewArray() As Byte, byCompressArray() As Byte
    Dim lNearEndofExpression&, lExpCounter&, lLenCompressArrayplus1&
    ' Set case according to status of comparison
        If Compare = vbTextCompare Then sExpression = LCase$(sExpression): sCompress = LCase$(sCompress)
        ' преобразуем строку в байтовый массив
        byExpressionArray = sExpression: byCompressArray = sCompress
        ' получаем размер байтового массива из ранее найденной длины строки (немного быстрее чем UBound)
        lLenExpressionArray = lLenExpression + lLenExpression - 1: lLenCompressArray = lLenCompress + lLenCompress - 1
        ReDim byNewArray(lLenExpressionArray): lNewCounter = 0
        bMatch = Left$(sExpression, 1) = sCompress
' Cжимаем односимвольные поисковые строки при помощи байтового массива
        ' Equate character match in boolean terms, applying logic "And"
        If Not bMatch And (lLenCompressArray = 1) Then
            For lbytePosition = 1 To lLenExpressionArray
                ' For/Next loops work faster advancing from 1
                ' But we want lbytePosition to start at byte zero so;
                lbytePosition = lbytePosition - 1
                If byExpressionArray(lbytePosition) <> byCompressArray(0) Then
                    ' Set this byte of byNewArray() equal to byExpressionArray()
                    byNewArray(lNewCounter) = byExpressionArray(lbytePosition): lNewCounter = lNewCounter + 2
                Else
                    ' Set this byte of byNewArray() equal to byCompressArray()
                    If byExpressionArray(lbytePosition - 2) <> byCompressArray(0) Then
                        byNewArray(lNewCounter) = byCompressArray(0): lNewCounter = lNewCounter + 2
                    End If
                End If
                ' As we want to advance 2 Hex places on each loop cycle
                ' (the range of 1 byte being 0-255 (&H0 to &HFF, 2^8 or 8-bits))
                ' we add 2 to lbytePosition, (a touch faster than using "Step 2")
                lbytePosition = lbytePosition + 2
            Next
        Else
' Сжимаем многосимвольную поисковую строку при помощи байтового массива
    ' For a 3 character compress string (could be any length here)
    ' if the expression contains "xbz" and the findstring is "abc"
    ' bMatchResult1 of ("x" And "a") = False
    ' bMatchResult2 of ("b" And "b") = True
    ' We "And" these two results so that (in this example);
    ' (bMatchResult1 And bMatchResult2) = (False And True) = False
    ' Pass this result to bMatchResult1 and pass a new test to bMatchResult2
    ' bMatchResult2 result: ("z" And "c") = False
    ' -> (bMatchResult1 And bMatchResult2) = (False And False) = False!
    
    ' We only get a result of True, if all characters in the stream match
            lNewCounter = 0
            lLenCompressArrayplus1 = lLenCompressArray + 1
            lNearEndofExpression = lLenExpressionArray - (lLenCompressArray - 1)
            bMatchResult1 = True
            ' lbytePosition is at start of expression, so we init subsequent loop by
            ' obtaining a match result for bFindArray() in bExpressionArray()
            For lbytePosition = 1 To lLenCompressArrayplus1
                lbytePosition = lbytePosition - 1
                ' Equate character match in boolean terms
                bMatchResult2 = byExpressionArray(lbytePosition) = byCompressArray(lbytePosition)
                ' Colate results to date and save
                bMatchResult1 = bMatchResult1 And bMatchResult2
                ' Provided no match was found in stream
                If Not bMatchResult1 Then
                    ' Reset byNewArray() counter to zero and exit loop
                    lNewCounter = 0
                    Exit For
                End If
                ' ..or advance both array counters and check next character
                byNewArray(lNewCounter) = byCompressArray(lbytePosition)
                lNewCounter = lNewCounter + 2
                lbytePosition = lbytePosition + 2
            Next
            ' With this result for first term done, we loop to build bNewArray()
            For lExpCounter = 1 To lLenExpressionArray
                lExpCounter = lExpCounter - 1
                ' Reset current result to False
                bMatch = False
                ' If we're not at end of the expression
                If lExpCounter < lNearEndofExpression Then
                    bMatch = True
                    ' For each element in byCompressArray(), check for match with expression
                    For lbytePosition = 1 To lLenCompressArray
                        lbytePosition = lbytePosition - 1
                        bMatchResult2 = byExpressionArray(lExpCounter + lbytePosition) = byCompressArray(lbytePosition)
                        bMatch = bMatch And bMatchResult2
                        lbytePosition = lbytePosition + 2
                    Next
                End If
                ' If no match found, set just one byte in byNewArray()
                ' to current byExpressionArray() byte value..
                If Not bMatch Then
                    byNewArray(lNewCounter) = byExpressionArray(lExpCounter)
                    ' ..and advance both counters to next byte
                    lNewCounter = lNewCounter + 2
                    lExpCounter = lExpCounter + 2
                ' Provided the previous pass through loop did not achieve a match
                ElseIf Not bMatchResult1 Then
                    ' Work through find string stream [ bFindArray() ]
                    For lbytePosition = 1 To lLenCompressArray
                        lbytePosition = lbytePosition - 1
                        ' Set current byte in byNewArray() = byCompressArray()
                        ' and advance all counters
                        byNewArray(lNewCounter) = byCompressArray(lbytePosition)
                        lNewCounter = lNewCounter + 2
                        lExpCounter = lExpCounter + 2
                        lbytePosition = lbytePosition + 2
                    Next
                Else
                    ' If no match was found
                    ' just advance lExpressionCounter beyond length of search string
                    lExpCounter = lExpCounter + lLenCompressArrayplus1
                End If
                ' ..accumulating results thus far
                bMatchResult1 = bMatch
            ' ..and try again
            Next
        End If
' преобразуем byNewArray() в строку и изменяем размер строки на правильный (немного быстрее чем ReDim Preserve с байтовым массивом)
        Compress = byNewArray: Compress = Left$(Compress, lNewCounter * 0.5)
        Exit Function
    Else
' Handle Error
    ' если искомая последовательность была нулевой длины возвращаем исходное выражение
        Compress = sExpression
    End If
End Function
Public Function IsAscII(txt As String) As Boolean
' проверяет является ли строка ASCII строкой
'-------------------------
    If Len(txt) = LenB(txt) Then IsAscII = True: Exit Function
Dim i As Long
    For i = 1 To Len(txt)
        If Asc(MidB$(txt, 2 * i, 1) & vbNullChar) <> 0 Then Exit Function  ' False
    Next i
    IsAscII = True
End Function
' ==================
' Функции для разбиения/модификации/формирования строк
' ==================
Public Function DelimStringGet(ByRef Source As String, _
    ByVal Pos As Long, _
    Optional Delim As String = " ", _
    Optional sBeg As Long, Optional sEnd As Long _
    ) As String
' возвращает фрагмент строки с разделителями с указанным индексом
'-------------------------
' Source    - исходная строка
' Pos       - позиция извлекаемой подстроки
' Delim     - разделитель
' sBeg,sEnd - возвращает позицию начала и окончания извлекаемой подстроки в исходной
'-------------------------
Dim Result As String: Result = vbNullString
    If Len(Source) = 0 Then GoTo HandleExit
'    If Pos < 1 then Goto HandleExit
'' при помощи Split - красивый, но медленный вариант
    'Result = Split(Source, Delim)(Pos - 1)
'' при помощи InStr - чуть длиннее, но сильно быстрее
    'Dim i As Long: i = 1: sBeg = 1
    'Do
    '    sEnd = InStr(sBeg, Source, Delim)
    '    If sEnd = 0 Then sEnd = Len(Source) + 1: Exit Do
    '    i = i + 1: If i > Pos Then Exit Do
    '    sBeg = sEnd + Len(Delim)
    'Loop
    'Result = Mid$(Source, sBeg, sEnd - sBeg)
' вариант с единой функцией поиска подстроки - производительность близка к оригинальному Split, на больших строках незначительно его превосходит
    ' позволяет использовать отрицательные позиции (с конца строки)
    Call p_GetSubstrBounds(Source, Pos, sBeg, sEnd, Delim)
    Result = Mid$(Source, sBeg, sEnd - sBeg)
HandleExit:  DelimStringGet = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function DelimStringDel(Source As String, _
    ByVal Pos As Long, _
    Optional Delim As String = " ", _
    Optional sBeg As Long, Optional sEnd As Long _
    ) As String
' возвращает фрагмент строки с разделителями без элемента с указанным индексом
'-------------------------
' Source    - исходная строка
' Pos       - позиция удаляемой подстроки
' Delim     - разделитель
' sBeg,sEnd - возвращает позицию удаленной подстроки
'-------------------------
Dim Result As String: Result = Source
    On Error GoTo HandleError
    If Len(Source) = 0 Then GoTo HandleExit
'    If Pos < 1 then Goto HandleExit
'' при помощи Split
'Dim arr() As String
'    arr = Split(Result, Delim): arr(Pos - 1) = vbNullString
'    Result = Replace(Join(arr, Delim), Delim & Delim, Delim): Erase arr()
'' при помощи InStr аналогично DelimStringGet
' вариант с единой функцией поиска подстроки - можно использовать отрицательные позиции (с конца строки)
    Call p_GetSubstrBounds(Source, Pos, sBeg, sEnd, Delim)
    If sBeg = 1 Then sEnd = sEnd + Len(Delim) Else sBeg = sBeg - Len(Delim)
    Result = Left$(Source, sBeg - 1) & Mid$(Source, sEnd)
HandleExit:  sEnd = sBeg: DelimStringDel = Result: Exit Function
HandleError: Result = Source: Err.Clear: Resume HandleExit
End Function
Public Function DelimStringSet(Source As String, _
    ByVal Pos As Long, ByVal Data As String, _
    Optional Delim As String = " ", _
    Optional SetUnique As Integer = False, _
    Optional Overwrite As Boolean = False, _
    Optional sBeg As Long, Optional sEnd As Long _
    ) As String
' вставляет в строку с разделителями элемент в позицию с указанным индексом
'-------------------------
' Source    - исходная строка
' Pos       - позиция вставки элемента
' Data      - вставляемая строка
' Delim     - разделитель
' SetUnique = False - вставка подстроки независимо от её наличия в исходной
'           = True  - вставка будет произведена только если подстрока отсутствует в исходной, иначе - исходная строка останется без изменений
'           = 1     - вставка будет произведена только если подстрока отсутствует в исходной, если подстрока уже присутствует в исходной, - подстрока будет удалена из исходной, а затем добавлена в указанную позицию
' Overwrite = False - вставка со сдвигом (для Pos>0 вставка перед указанной позицией, для Pos<0 - после. т.о. Pos=1 - вставка вначало, а Pos=-1 - в конец строки)
'           = True  - вставка с заменой  элемента строки в указанной позиции,
'             (!)     в совокупности с SetUnique<>0 может приводить к неожиданным результатам
' sBeg,sEnd - возвращает позицию начала и окончания вставленной подстроки в исходной
'-------------------------
' v.1.0.1       : 09.08.2022 - добавлен параметр SetUnique для контроля уникальности вставляемых значений
'-------------------------
Dim Result As String: Result = Source
    On Error GoTo HandleError
    If Len(Result) = 0 Then Result = Data: GoTo HandleExit
    ' проверить исходную строку на наличие вхождений подстроки,
    If SetUnique Then
        Select Case SetUnique
        Case 1      ' если вставляемый элемент уже есть в строке - удалить найденные вхождения и продолжить
            If Result = Data Then GoTo HandleExit
            Result = Replace(Result, Delim & Data & Delim, Delim)
            If Left$(Result, Len(Data & Delim)) = Data & Delim Then Result = Mid$(Result, Len(Data & Delim) + 1)
            If Right$(Result, Len(Delim & Data)) = Delim & Data Then Result = Left$(Result, Len(Result) - Len(Data & Delim))
        Case True   ' если вставляемый элемент уже есть в строке - выход
            If Result = Data Then GoTo HandleExit
            sBeg = InStr(1, Result, Delim & Data & Delim): If sBeg Then sBeg = sBeg + Len(Delim): GoTo HandleExit
            If Left$(Result, Len(Data & Delim)) = Data & Delim Then sBeg = 1: GoTo HandleExit
            If Right$(Result, Len(Delim & Data)) = Delim & Data Then sBeg = Len(Result) - Len(Data) + 1: GoTo HandleExit
        End Select
    End If
    ' проверить позицию вставки
    Select Case Pos
    Case 1:     If Not Overwrite Then sBeg = 1:           Result = Data & Delim & Result: GoTo HandleExit
    Case -1:    If Not Overwrite Then sBeg = Len(Result): Result = Result & Delim & Data: GoTo HandleExit
    End Select
'    If Pos < 1 then Goto HandleExit
'' при помощи Split
'    Dim arr() As String:arr = Split(Result, Delim)
'    If Overwrite Then arr(Pos - 1) = Data Else arr(Pos - 1) = Data & Delim & arr(Pos - 1)
'    Result = Join(arr, Delim): Erase arr()
'' при помощи InStr аналогично DelimStringGet
' вариант с единой функцией поиска подстроки - можно использовать отрицательные позиции (с конца строки)
    If Overwrite Then
    ' вставка с заменой
        ' получаем позицию от начала, границы
        Call p_GetSubstrBounds(Result, Pos, sBeg, sEnd, Delim)
    Else
    ' вставка со сдвигом
        ' получаем позицию от начала, границы, проверяем выход за пределы строки
        ' и сравниваем его с направлением просмотра строки (Sgn(Pos) = -1) д.б. до вызова ф-ции (переопределяет Pos)
        If (Sgn(Pos) = -1) = p_GetSubstrBounds(Result, Pos, sBeg, sEnd, Delim) Then
        ' есть выход за пределы и направление от начала    (bRes = False; bDir = False)
        ' или нет выхода за пределы и направление от конца (bRes = True;  bDir = True)
        ' вставка после:  "PrevVal & Delim & NewVal"
            Data = Mid$(Result, sBeg, sEnd - sBeg) & Delim & Data: Pos = Pos + 1
        Else
        ' нет выхода за пределы и направление от начала    (bRes = True;  bDir = False)
        ' или есть выход за пределы и направление от конца (bRes = False; bDir = True)
        ' вставка перед: "NewVal & Delim & PrevVal"
            Data = Data & Delim & Mid$(Result, sBeg, sEnd - sBeg)
        End If
    End If
    Result = Left$(Result, sBeg - 1) & Data & Mid$(Result, sEnd)
HandleExit:  sEnd = sBeg + Len(Data): DelimStringSet = Result: Exit Function
HandleError: Result = Source: Err.Clear: Resume HandleExit
End Function
Public Function DelimStringShrink(ByVal Source As String, _
    Optional Delim As String = " " _
    ) As String
' удаляет из строки с разделителями повторяющиеся элементы оставляя только первое вхождение
'-------------------------
' Source - исходная строка
' Delim - разделитель
'-------------------------
Dim Arr() As String, Col As Collection
Dim Result As String

    On Error GoTo HandleError
    Result = Trim$(Source)
    If Len(Result) = 0 Then GoTo HandleExit
' разбираем строку
    'Call xSplit(Result, Arr, Delim)
    Arr = Split(Result, Delim)
    Set Col = New Collection
Dim i As Long, iMax As Long: i = LBound(Arr): iMax = UBound(Arr)
    On Error Resume Next
' переносим в коллекцию
Dim Itm
    Do Until i > iMax
        Itm = Trim$(Arr(i))
        Col.Add Itm, c_idxPref & Itm: Err.Clear: i = i + 1
    Loop
    On Error GoTo HandleError
' собираем строку
    Result = vbNullString
    For Each Itm In Col
        Result = Result & Delim & Itm
    Next Itm
    Result = Mid$(Result, Len(Delim) + 1)
HandleExit:  Erase Arr: Set Col = Nothing
             DelimStringShrink = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function DelimStringSimile(ByVal String1 As String, ByVal String2 As String, _
    Optional Delim As String = " ", _
    Optional Compare As VbCompareMethod = vbTextCompare _
    ) As Boolean
' ищет в строках с разделителями совпадающие элементы, возвращает True при нахождении первого совпадения
'-------------------------
' String1, String2  - строки с разделителями элементы которых будут сравниваться между собой
' Delim             - разделитель
' Compare           - метод сравнения
'-------------------------
Dim Result As Boolean ':Result = False
    On Error GoTo HandleError
    String1 = Trim$(String1): If Len(String1) = 0 Then Err.Raise vbObjectError + 512
    String2 = Trim$(String2): If Len(String2) = 0 Then Err.Raise vbObjectError + 512
    If Compare = vbTextCompare Then String1 = UCase(String1): String2 = UCase(String2)
Dim a1() As String: a1 = Split(String1, Delim) 'Call xSplit(String1, a1, Delim)
Dim a2() As String: a2 = Split(String2, Delim) 'Call xSplit(String2, a2, Delim)
Dim i As Long, j As Long: i = LBound(a1): j = LBound(a2)
    Do
        If i > UBound(a1) Then i = LBound(a1): j = j + 1: If j > UBound(a2) Then Exit Do
        Result = Trim$(a1(i)) = Trim$(a2(j)): i = i + 1
    Loop Until Result
HandleExit:  DelimStringSimile = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function TokenStringGet(Source As String, _
    ByVal Pos As Long, _
    Optional Delims As String = c_strSymbSpaces & c_strSymbPunct & c_strSymbParenth & c_strSymbMath, _
    Optional IncEmpty As Boolean = False, _
    Optional DelimsLeft As String, Optional DelimsRight As String, _
    Optional sBeg As Long, Optional sEnd As Long _
    ) As String
' возвращает фрагмент строки со множественными разделителями с указанным индексом
'-------------------------
' Source    - исходная строка
' Pos       - позиция извлекаемого элемента
' Delims    - набор разделителей для разбиения исходной строки
' IncEmpty  = False - пропуск пустых элементов - последовательные разделители будут рассматриваться как один
'           = True  - результирующий массив будет включать пустые элементы между последовательными разделителями
' DelimsLeft/DelimsRight    - возвращают разделители расположенные слева и справа от извлекаемого токена
' sBeg,sEnd - возвращает позицию начала и окончания извлекаемой подстроки (токена) в исходной
'-------------------------
Dim Result As String: Result = Source
    
    On Error GoTo HandleError
    If Len(Source) = 0 Then Result = vbNullString: GoTo HandleExit
Dim aData() As String, aPos() As Long
Dim aMin As Long: aMin = 1
Dim aMax As Long: aMax = Tokenize(Source, aData(), Delims, aPos(), IncEmpty)
    If Pos < 0 Then Pos = aMax + Pos + 1
    If Pos < aMin Then Pos = aMin Else If Pos > aMax Then Pos = aMax
    ' блок разделителей до выбранного токена (слева)
    sEnd = aPos(Pos - 1): If Pos = 1 Then sBeg = 1 Else sBeg = aPos(Pos - 2) + Len(aData(Pos - 2))
    If sEnd > sBeg Then DelimsLeft = Mid$(Source, sBeg, sEnd - sBeg) Else DelimsLeft = vbNullString
    ' блок разделителей после выбранного токена (справа)
    sBeg = aPos(Pos - 1) + Len(aData(Pos - 1)): If Pos = aMax Then sEnd = Len(Source) Else sEnd = aPos(Pos)
    If sEnd > sBeg Then DelimsRight = Mid$(Source, sBeg, sEnd - sBeg) Else DelimsRight = vbNullString
    ' значение токена
    Result = aData(Pos - 1): sBeg = aPos(Pos - 1): sEnd = sBeg + Len(Result)
HandleExit:  TokenStringGet = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function TokenStringSet(Source As String, _
    ByVal Pos As Long, ByVal Data As String, _
    Optional Delims As String = c_strSymbSpaces & c_strSymbPunct & c_strSymbParenth & c_strSymbMath, _
    Optional IncEmpty As Boolean = False, _
    Optional Overwrite As Boolean = False, _
    Optional NewDelim As String = " ", _
    Optional sBeg As Long, Optional sEnd As Long _
    ) As String
' вставляет в строку со множественными разделителями элемент в позицию с указанным индексом
'-------------------------
' Source    - исходная строка
' Pos       - позиция вставки элемента
' Data      - вставляемая строка
' Delims    - набор разделителей для разбиения исходной строки
' IncEmpty  = False - пропуск пустых элементов - последовательные разделители будут рассматриваться как один
'           = True  - результирующий массив будет включать пустые элементы между последовательными разделителями
' Overwrite = False - вставка со сдвигом (для Pos>0 вставка перед указанной позицией, для Pos<0 - после. т.о. Pos=1 - вставка вначало, а Pos=-1 - в конец строки)
'           = True  - вставка с заменой  элемента строки в указанной позиции
' NewDelim - добавляемые разделители (обязательно при Overwrite=False) вставляются после Data перед следующим токеном
' sBeg,sEnd - возвращает позицию начала и окончания вставленной подстроки (токена) в исходной
'-------------------------
Dim Result As String: Result = Source
    
    On Error GoTo HandleError
    If Len(Source) = 0 Then Result = Data: Pos = 1: GoTo HandleExit
    Select Case Pos
    Case 1:     If Not Overwrite Then sBeg = 1:           Result = Data & NewDelim & Source: GoTo HandleExit
    Case -1:    If Not Overwrite Then sBeg = Len(Source): Result = Source & NewDelim & Data: GoTo HandleExit
    End Select
Dim aData() As String, aPos() As Long
Dim aMin As Long: aMin = 1
Dim aMax As Long: aMax = Tokenize(Source, aData(), Delims, aPos(), IncEmpty)
Dim sTemp As String: sTemp = Data
    If Pos < 0 Then Pos = aMax + Pos + 2
    If Pos < aMin Then Pos = aMin Else If Pos > aMax Then Pos = aMax
    sBeg = aPos(Pos - 1)
    If Overwrite Then
    ' замена текущего фрагмента строки заданным. имеющиеся разделители сохраняются
        sEnd = sBeg + Len(aData(Pos - 1))
    Else
    ' вставка строки перед указанной. добавить разделители (NewDelim)
        sEnd = sBeg: sTemp = sTemp & NewDelim
    End If
    Result = Left$(Result, sBeg - 1) & sTemp & Mid$(Result, sEnd)
HandleExit:  sEnd = sBeg + Len(Data): TokenStringSet = Result: Exit Function
HandleError: Result = Source: Err.Clear: Resume HandleExit
End Function
Public Function TokenStringDel(Source As String, _
    ByVal Pos As Long, _
    Optional Delims As String = c_strSymbSpaces & c_strSymbPunct & c_strSymbParenth & c_strSymbMath, _
    Optional IncEmpty As Boolean = False, _
    Optional SubDelims As Boolean = True, Optional NewDelim As String = " ", _
    Optional sBeg As Long, Optional sEnd As Long _
    ) As String
' удаляет токен с указанным номером из строки со множественными разделителями
'-------------------------
' Source    - исходная строка
' Pos       - позиция удаляемой подстроки
' Delims    - набор разделителей для разбиения исходной строки
' IncEmpty  = False - пропуск пустых элементов - последовательные разделители будут рассматриваться как один
'           = True  - результирующий массив будет включать пустые элементы между последовательными разделителями
' SubDelims = False - разделители вокруг удаляемого элемента остаются в итоговой строке,
'           = True  - разделители оставшиеся после удаления элемента будут заменены на EndDelims
' NewDelim - конечные разделители (обязательно при SubDelims = True) добавляются после Data перед следующим токеном
' sBeg,sEnd - возвращает позицию удаленной подстроки (токена)
'-------------------------
Dim Result As String: Result = Source
    
    On Error GoTo HandleError
    If Len(Source) = 0 Then Pos = 1: GoTo HandleExit
Dim aData() As String, aPos() As Long
Dim aMin As Long: aMin = 1
Dim aMax As Long: aMax = Tokenize(Source, aData(), Delims, aPos(), IncEmpty)
Dim sTemp As String ':  sTemp = vbNullString
    If Pos < 0 Then Pos = aMax + Pos + 1
    If Pos < aMin Then Pos = aMin Else If Pos > aMax Then Pos = aMax
    If SubDelims Then
    ' разделители слева и справа удаляются и, если вставка в середине строки,- заменяются на новый
        Select Case Pos
        Case aMin: sBeg = 1: sEnd = aPos(Pos)
        Case aMax: sBeg = aPos(Pos - 2) + Len(aData(Pos - 2)): sEnd = Len(Result) + 1
        Case Else: sBeg = aPos(Pos - 2) + Len(aData(Pos - 2)): sEnd = aPos(Pos): sTemp = NewDelim
        End Select
    Else
    ' имеющиеся разделители сохраняются
        sBeg = aPos(Pos - 1)
        sEnd = sBeg + Len(aData(Pos - 1))
    End If
    Result = Left$(Result, sBeg - 1) & sTemp & Mid$(Result, sEnd)
HandleExit:  sEnd = sBeg: TokenStringDel = Result: Exit Function
HandleError: Result = Source: Err.Clear: Resume HandleExit
End Function
Public Function TaggedStringGet(Source As String, _
    ByRef Tag As String, _
    Optional Delim As String = ";", Optional TagDelim As String = "=", _
    Optional Compare As VbCompareMethod = vbTextCompare, _
    Optional sBeg As Long, Optional sEnd As Long _
    ) As String
' возвращает значение элемента строки с разделителями с указанным именем
'-------------------------
' лучше для этих целей использовать класс вроде TaggedValues от Гетца: https://www.sql.ru/forum/661816/vdrug-u-kogo-est-dlya-obrabotki-v-vba-strok-svoystv
' но иногда и такой вариант бывает полезен
' функция никак не проверяет уникальность тэга, - будет возвращено первое вхождение
' Source    - строка элементов вида "Tag1=Val1;...TagN=ValN"
' Tag       - имя (Tag) элемента. если Tag не задан - элемент будет получен по Pos
' Delim     - разделитель пар имя (Tag) / значение (Val)
' TagDelim  - разделитель имени (Tag) и значение (Val) в паре
' sBeg,sEnd - возвращает позицию начала и окончания извлеченной подстроки (значения тэга) в исходной
' Compare   - тип сравнения (vbBinaryCompare/vbTextCompare)
' возвращает значение (Val) элемента с указанным именем (Tag)
'-------------------------
'' ! если нужен доступ по позиции фрагмента - надо раскомментировать соотв строки
'    и добавить Optional ByRef Pos As Long = 0, _
'' Pos -     на входе позиция возвращаемого элемента. (если указан Tag не используется)
''           >0 - позиция относительно начала строки
''           <0 - позиция относительно конца строки
''           на выходе позиция полученного элемента относительно начала строки
'-------------------------

'RegExp: Mask = Delim & Tag & TagDelim & "(.+?)" & Delim
'        With Regex.Execute(Delim & Source & Delim)(0)
'           Val = .SubMatches(0)
'           Pos = .Length + Len(Tag & TagDelim) - Len(Delim)
'        End With
Dim Result As String: Result = vbNullString
    On Error GoTo HandleError
    If Len(Source) = 0 Then GoTo HandleExit
    sBeg = 1
    If Len(Tag) > 0 Then
' ищем по Tag
Dim pLen As Long: pLen = Len(Tag & TagDelim): If Len(Source) < pLen Then GoTo HandleExit 'Pos = 0: GoTo HandleExit
        sBeg = sBeg + pLen
'        If StrComp(Mid$(Source, 1, pLen), Tag & TagDelim, Compare) = 0 ' If Mid$(tmpSrc, 1, pLen) = tmpTag & TagDelim Then
'            Pos = 1
'        Else
        If StrComp(Mid$(Source, 1, pLen), Tag & TagDelim, Compare) <> 0 Then ' If Mid$(tmpSrc, 1, pLen) <> tmpTag & TagDelim Then
'        ' если тэг с заданным именем не в начале строки
'        ' ищем подстроку с заданным именем начинающуюся с Delim и заканчивающуюся TagDelim
            sBeg = InStr(1, Source, Delim & Tag & TagDelim, Compare)
            If sBeg = 0 Then GoTo HandleExit 'Pos = 0: GoTo HandleExit
            sBeg = sBeg + Len(Delim) + pLen
'            Pos = InStrCount(Left$(Source, sBeg), Delim) + 1
        End If
'        ' сейчас в sBeg позиция начала значения тэга
'        ' ищем позицию конца значения тэга с позиции начала до следующего Delim
        sEnd = InStr(sBeg, Source, Delim, Compare): If sEnd = 0 Then sEnd = Len(Source) + 1
        Result = Mid$(Source, sBeg, sEnd - sBeg)
'    Else
'' ищем по Pos
'        Call p_GetSubstrBounds(Source, Pos, sBeg, sEnd, Delim)
'        Result = Mid$(Source, sBeg, sEnd - sBeg)
'        sEnd = InStr(Result, TagDelim): sBeg = sEnd + Len(TagDelim)
'        Tag = Left$(Result, sEnd - 1): Result = Mid$(Result, sBeg)
    End If
HandleExit:  TaggedStringGet = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function TaggedStringSet(Source As String, _
    ByRef Tag As String, Optional ByRef Data As String, _
    Optional Delim As String = ";", Optional TagDelim As String = "=", _
    Optional sBeg As Long, Optional sEnd As Long, _
    Optional Compare As VbCompareMethod = vbTextCompare _
    ) As String
' устанавливает значение элемента строки с разделителями с указанным именем
'-------------------------
' Source -  строка элементов вида "Tag1=Val1;...TagN=ValN"
' Tag -     имя (Tag) устанавливаемого элемента, если элемент отсутствует - будет добавлен
' Data -    значение (Val) устанавливаемого элемента
' Delim -   разделитель пар имя (Tag) / значение (Val)
' TagDelim - разделитель имени (Tag) и значения (Val) в паре
' sBeg,sEnd - возвращает позицию начала и окончания вставленной подстроки (значения тэга) в исходной
' Compare - тип сравнения (vbBinaryCompare/vbTextCompare)
' возвращает строку элементов с учетом добавленного элемента
'-------------------------
'' ! если нужен доступ по позиции фрагмента - надо раскомментировать соотв строки
'    и добавить Optional ByRef Pos As Long = 0, _
'' Pos -     на входе позиция вставки элемента. (порядковый номер позиции в результирующей строке)
''           0  - будет добавлен в позицию найденного элемента или конец при его отсутствии
''           >0 - позиция относительно начала строки, вставка перед указанной позицией
''           <0 - позиция относительно конца строки, вставка после указанной позиции
''           на выходе содержит позицию добавленного элемента относительно начала строки
'-------------------------
' ! функция никак не проверяет уникальность тэга, - будет заменено первое вхождение
'-------------------------
' v.1.0.3       : 11.02.2022 - Исправлены ошибки возникающие если TagDelim текстовое выражение, зависящее от регистра
'-------------------------
Dim Result As String: Result = Source
    On Error GoTo HandleError
' пустое имя тэга
    If Len(Tag) = 0 Then
        Result = Source
'' ищем по Pos и получаем его Tag
'   ' при получении по позиции - не даем выходить за пределы строки, - всегда берем имеющуюся подстроку
'        Call p_GetSubstrBounds(Result, Pos, sBeg, sEnd, Delim)
'        Tag = Split(Mid$(Result, sBeg, sEnd - sBeg), TagDelim)(0)
'        sTemp = Tag & TagDelim & Data   ' "Tag=Val" - тэг только что получен
        GoTo HandleExit
    End If
' пустое значение
    If Len(Data) = 0 Then Result = TaggedStringDel(Source, Tag): GoTo HandleExit
' пустая строка
    If Len(Source) = 0 Then If Len(Tag) > 0 Then Result = Tag & TagDelim & Data: GoTo HandleExit ': Pos = 1
Dim sTemp As String
Dim pLen As Long
'' ищем по Tag
'    ' поиск наличия тэга с таким именем
    sBeg = 0: sEnd = 0 'Len(Result)
    pLen = 1 'Len(Tag) + Len(TagDelim)
    sTemp = Tag & TagDelim & Data ' "Tag=Val"
    If StrComp(Left$(Source, Len(Tag) + Len(TagDelim)), Tag & TagDelim, Compare) = 0 Then  ' Left$(Source, Len(Tag) + Len(TagDelim)) = Tag & TagDelim Then
    ' имя тэга найдено в начале строки ("Tag=...")
        sBeg = 1
    Else
    ' ищем в середине строки имя тэга с разделителем тэг/значение ("...;Tag=...")
        sBeg = InStr(pLen + 1, Result, Delim & Tag & TagDelim, Compare)
        If sBeg > 0 Then
    ' имя тэга с разделителем тэг/значение найдено в середине строки ("...;Tag=...")
            pLen = pLen + Len(Delim)
        Else
    ' если не найдено - проверяем тэг без значения и разделителя тэг/значение
            pLen = Len(Tag) + Len(Delim)
            If StrComp(Left$(Source, pLen), Tag & Delim, Compare) = 0 Then 'Left$(Source, pLen) = Tag & Delim Then
        ' проверяем в начале строки ("Tag;...")
                sBeg = 1: sEnd = Len(Tag) + 1
            ElseIf StrComp(Source, Tag, Compare) = 0 Then  'tmpSource = tmpTag Then
        ' проверяем всю строку ("Tag")
                sBeg = 1: sEnd = Len(Result) + 1
            ElseIf StrComp(Right$(Source, pLen), Delim & Tag, Compare) = 0 Then  'Right$(tmpSource, pLen) = tmpDelim & tmpTag Then
        ' проверяем в конце строки ("...;Tag")
                sBeg = Len(Result) - Len(Tag): sEnd = Len(Result) + 1
            Else
        ' проверяем в середине строки ("...;Tag;...")
                sBeg = InStr(1, Result, Delim & Tag & Delim, Compare): sEnd = sBeg + pLen
            End If
        End If
    End If
' если начало найдено а конец не определен ищем конечный разделитель
    If sBeg > 0 And sEnd = 0 Then sEnd = InStr(sBeg + pLen + 1, Result, Delim, Compare): If sEnd = 0 Then sEnd = Len(Result) + 1
'        bFound = sBeg > 0

'        If Pos = 0 Then
'    ' вставка в позицию найденного элемента или в конец строки
'        ' корректируем границы фрагмента
'            ' найден не в начале вставка после разделителя
'            ' не найден - вставка в конец, добавляем перед вставкой разделитель
        If sBeg = 0 Then sBeg = Len(Result) + 1: sEnd = sBeg
'        ' получаем позицию фрагмента вставки
        If sBeg > 1 Then
            sTemp = Delim & sTemp
'                Pos = InStrCount(Left$(Result, sBeg), Delim) + 2
'            Else
'                Pos = 1
        End If
'        Else
'    ' удаление фрагмента в текущей позиции и вставка в указанную позицию
'        ' удаление в найденной позиции
'            If sBeg > 0 Then
'                If sBeg = 1 Then sEnd = sEnd + Len(Delim)
'                Result = Left$(Source, sBeg - 1) & Mid$(Source, sEnd)
'            End If
'        ' получаем позицию фрагмента вставки и формируем вставляемую строку
'            'Result = DelimStringSet(Result, Pos, sTemp, Delim, Overwrite:=False): GoTo HandleExit
'            Select Case Pos
'            Case 1:     Result = sTemp & Delim & Result: GoTo HandleExit
'            Case -1:    Result = Result & Delim & sTemp: Pos = InStrCount(Result, Delim) + 2: GoTo HandleExit
'            End Select
'            If (Sgn(Pos) = -1) = p_GetSubstrBounds(Source, Pos, sBeg, sEnd, Delim) Then
'                sTemp = Mid$(Result, sBeg, sEnd - sBeg) & Delim & sTemp: If Not bFound Then Pos = Pos + 1
'            Else
'                sTemp = sTemp & Delim & Mid$(Result, sBeg, sEnd - sBeg)
'            End If
'        End If

' вставка в указанную позицию
    Result = Left$(Result, sBeg - 1) & sTemp & Mid$(Result, sEnd)
HandleExit:  sBeg = sEnd - Len(Data): TaggedStringSet = Result: Exit Function
HandleError: Result = Source: Err.Clear: Resume HandleExit
End Function
Public Function TaggedStringDel(Source As String, _
    ByRef Tag As String, _
    Optional Delim As String = ";", Optional TagDelim As String = "=", _
    Optional sBeg As Long, Optional sEnd As Long, _
    Optional Compare As VbCompareMethod = vbTextCompare _
    ) As String
' удаляет элемент строки с разделителями с указанным именем
'-------------------------
' Source - строка элементов вида "Tag1=Val1;...TagN=ValN"
' Tag -     имя (Tag) элемента. если Tag не задан - элемент будет получен по Pos
' Delim -   разделитель пар имя (Tag) / значение (Val)
' TagDelim - разделитель имени (Tag) и значение (Val) в паре
' sBeg,sEnd - возвращает позицию начала и окончания вставленной подстроки в исходной
' Compare - тип сравнения (vbBinaryCompare/vbTextCompare)
' Возвращает строку элементов без указанного элемента
'-------------------------
'' ! если нужен доступ по позиции фрагмента - надо раскомментировать соотв строки
'    и добавить Optional ByRef Pos As Long = 0, _
'' Pos -     на входе позиция удаляемого элемента. (если указан Tag не используется)
''           >0 - позиция относительно начала строки
''           <0 - позиция относительно конца строки
''           на выходе позиция элемента предшествующего удалённому относительно начала строки или 1
'-------------------------
' ! функция никак не проверяет уникальность тэга, - будет удалено первое вхождение
'-------------------------
Dim Result As String: Result = Source
    On Error GoTo HandleError
    If Len(Source) = 0 Then GoTo HandleExit
    If Len(Tag) > 0 Then
'' ищем по Tag
'    ' поиск наличия тага с таким именем
'' ищем по Tag
'    ' поиск наличия тэга с таким именем
        sBeg = 0: sEnd = 0 'Len(Result)
Dim pLen As Long
        pLen = 1 'Len(Tag) + Len(TagDelim)
        'sTemp = Tag & TagDelim & Data   ' "Tag=Val"
        If StrComp(Left$(Source, Len(Tag) + Len(TagDelim)), Tag & TagDelim, Compare) = 0 Then ' Left$(tmpSource, Len(Tag) + Len(TagDelim)) = tmpTag & tmpTagDelim Then
        ' имя тэга найдено в начале строки ("Tag=...")
            sBeg = 1
        Else
        ' ищем в середине строки имя тэга с разделителем тэг/значение ("...;Tag=...")
            sBeg = InStr(pLen + 1, Result, Delim & Tag & TagDelim, Compare)
            If sBeg > 0 Then
        ' имя тэга с разделителем тэг/значение найдено в середине строки ("...;Tag=...")
                pLen = pLen + Len(Delim)
            Else
        ' если не найдено - проверяем тэг без значения и разделителя тэг/значение
                pLen = Len(Tag) + Len(Delim)
                If StrComp(Left$(Source, pLen), Tag & Delim, Compare) = 0 Then  ' Left$(tmpSource, pLen) = tmpTag & tmpDelim Then
            ' проверяем в начале строки ("Tag;...")
                    sBeg = 1: sEnd = Len(Tag) + 1
                ElseIf StrComp(Source, Tag, Compare) = 0 Then   ' tmpSource = tmpTag Then
            ' проверяем всю строку ("Tag")
                    sBeg = 1: sEnd = Len(Result) + 1
                ElseIf StrComp(Right$(Source, pLen) = Delim & Tag, Compare) = 0 Then   'Right$(tmpSource, pLen) = tmpDelim & tmpTag Then
            ' проверяем в конце строки ("...;Tag")
                    sBeg = Len(Result) - Len(Tag): sEnd = Len(Result) + 1
                Else
            ' проверяем в середине строки ("...;Tag;...")
                    sBeg = InStr(1, Result, Delim & Tag & Delim, Compare): sEnd = sBeg + pLen
                End If
            End If
        End If
        If sBeg = 0 Then GoTo HandleExit ' тэг не найден
'        ' если начало фрагмента найдено ищем его конец
        If sEnd = 0 Then sEnd = InStr(sBeg + pLen + 1, Result, Delim, Compare): If sEnd = 0 Then sEnd = Len(Result) + 1
'        ' получаем позицию фрагмента удаления
        If sBeg <= 1 Then
            sEnd = sEnd + Len(Delim) ': Pos = 1
'        Else
'            Pos = InStrCount(Left$(Result, sBeg), Delim) + 1
        End If
'    Else
'' ищем по Pos
'    ' при получении по позиции - не даем выходить за пределы строки, - всегда берем имеющуюся подстроку
'        Call p_GetSubstrBounds(Result, Pos, sBeg, sEnd, Delim)
'        'Tag = Split(Mid$(Result, sBeg, sEnd - sBeg), TagDelim)(0)
'        If sBeg > 1 Then sBeg = sBeg - Len(Delim) Else sEnd = sEnd + Len(Delim)
'        If Pos > 1 Then Pos = Pos - 1
    End If
' удаление указанной позиции
    Result = Left$(Result, sBeg - 1) & Mid$(Result, sEnd)
HandleExit:  sEnd = sBeg: TaggedStringDel = Result: Exit Function
HandleError: Result = Source: Err.Clear: Resume HandleExit
End Function
Public Function TaggedString2Collection(Source As String, _
    Optional Tags As Collection, Optional Keys, _
    Optional Delim As String = ";", Optional TagDelim As String = "=", _
    Optional ReplaceExisting As Integer = True _
    ) As Boolean
' преобразует строку именованных параметров в коллекцию значений с ключами соотв имени тэга
'-------------------------
' Source - строка элементов вида "Tag1=Val1;...TagN=ValN"
' Tags - возвращаемая коллекция значений тэгов. если на входе передана непустая коллекция новые значения будут добавлены к ней
' Keys - (если задано) возвращает массив ключей коллекции (Tag)
' Delim -   разделитель пар имя (Tag) / значение (Val)
' TagDelim - разделитель имени (Tag) и значение (Val) в паре
' ReplaceExisting - определяет поведение при обнаружении переменных с одинаковым именем
'   0 - в коллекцию будет добавлена первая найденная, последующие будут игнорироваться
'  -1 - переменные с одинаковым именем будут заменяться - в коллекции останется последняя найденная
'-------------------------
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    Do While Right(Source, Len(Delim)) = Delim: Source = Trim(Left$(Source, Len(Source) - Len(Delim))): Loop
    If Len(Source) = 0 Then GoTo HandleExit
    If Tags Is Nothing Then Set Tags = New Collection
Dim Arr() As String, Term 'As String
Dim sKey As String, vVal 'As String
    Arr() = Split(Source, Delim)
    For Each Term In Arr
        sKey = Split(Term, TagDelim)(0) ' получаем тэг
        If p_IsExist(sKey, Tags) Then If ReplaceExisting Then Tags.Remove sKey Else GoTo HandleNext
        vVal = Split(Term, TagDelim)(1) ' получаем значение
        Tags.Add vVal, sKey
HandleNext:
    Next
    Erase Arr
    If Not IsMissing(Keys) Then Keys = p_GetCollKeys(Tags)
    Result = True
HandleExit:  TaggedString2Collection = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function TaggedCollection2String(Tags As Collection, _
    Optional Delim As String = ";", Optional TagDelim As String = "=" _
    ) As String
' преобразует коллекцию значений с ключами соотв имени тэга в строку именованных параметров
'-------------------------
Dim Result As String: Result = vbNullString
    On Error GoTo HandleError
    If Tags Is Nothing Then GoTo HandleExit
    If Tags.Count = 0 Then GoTo HandleExit
Dim Keys() As String: Keys() = p_GetCollKeys(Tags)
Dim i As Long
    For i = 1 To Tags.Count
        Result = Result & Delim & Keys(i) & TagDelim & Tags.Item(i)
    Next i
    If Left$(Result, Len(Delim)) = Delim Then Result = Mid$(Result, Len(Delim) + 1)
HandleExit:  TaggedCollection2String = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Private Function p_GetSubstrBounds(ByRef Source As String, _
    ByRef Pos As Long, ByRef sBeg As Long, ByRef sEnd As Long, _
    Optional Delim As String = " ") As Boolean
' возвращает номер и позицию начала и конца подстроки в строке с разделителями
'-------------------------
' Source -  исходная строка
' Pos -     на входе позиция искомого элемента.
'           >0 - позиция относительно начала строки
'           <0 - позиция относительно конца строки
'           на выходе позиция элемента относительно начала строки
' sBeg, sEnd - возвращает границы подстроки в строке
' Delim -   разделитель
' возвращает True если заданная позиция в границах строки, иначе False
'-------------------------
Dim Result As Boolean
    On Error GoTo HandleError
Dim i As Long
    i = 1: sBeg = 1
    If Pos >= 0 Then
' позиция от начала
    ' пробегаем всю строку с начала по разделителям, проверяя номер подстроки
        If Pos = 0 Then Pos = 1
        Do
            sEnd = InStr(sBeg, Source, Delim)
            If sEnd = 0 Then sEnd = Len(Source) + 1:  Exit Do
            i = i + 1: If i > Pos Then Exit Do
            sBeg = sEnd + Len(Delim)
        Loop
        Result = (Pos <= i): If Not Result Then Pos = i  ' позиция выше верхней границы
    Else
' позиция от конца
    ' Вариант 1: пробегаем всю строку с конца по разделителям, проверяя номер подстроки
    ' Вариант 2: пробегаем всю строку с начала подсчитывая количество подстрок и формируя массив позиций разделителей
    Dim aPos() As Long
    ' подсчитываем количество фрагментов в строке, формируя массив позиций разделителей
        Do
            ReDim Preserve aPos(1 To i): aPos(i) = sBeg
            sBeg = InStr(sBeg, Source, Delim) '
            i = i + 1
            If sBeg > 0 Then sBeg = sBeg + Len(Delim) Else Exit Do
        Loop
    ' переводим позицию относительно конца строки в позицию от начала
        Pos = i + Pos
        Select Case Pos
        Case 1 To i: Result = True              ' позиция в пределах строки
        Case Is < 1: Result = False: Pos = 1    ' позиция ниже нижней границы
        'Case Is > i: Result = False: Pos = i    ' позиция выше верхней границы
        End Select
    ' берем границы фрагмента с указанной позицией из массива
        sBeg = aPos(Pos): If Pos < (i - 1) Then sEnd = aPos(Pos + 1) - Len(Delim) Else sEnd = Len(Source) + 1
        Erase aPos()
    End If
HandleExit:  p_GetSubstrBounds = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
' ==================
' Функции для разбиения/распределения элементов строки с разделителями
' ==================
#If APPTYPE = 0 Then ' только для Access
Public Sub TextToArrayByControl(TextString As String, _
    vControls As Variant, _
    Optional Separators As String = " ­.,;:!?()[]{}…+-*/\|" & vbTab & vbCrLf)
' разбивает текст с учётом допустимых разделителей на строки, соответствующие ширине полей, и распределяет её по полям
'-------------------------
' TextString - строка текста которую необходимо разбить
' vControls  - коллекция или массив полей в которых надо распределить текст
' Separators - список разделителей по которым можно бить текст если включен символ мягкого переноса - в выходной строке будет опущен
'-------------------------
Dim hFont As LongPtr, hOldFont As LongPtr
Dim WidthInPix As Long
Dim sz As Size, tm As TEXTMETRIC
Dim strText As String, strRest As String, strTemp As String
Dim aWords() As String
Dim spLen As Long, spPos As Integer, spPosNext As Integer
Dim tWidth As Long, tHeight As Long
Dim i As Long, iMax As Long, w As Long
Dim aCtl As Variant, ctl  As Variant

Dim strMessage As String, strComment As String
Dim Result As Boolean

    Result = False
    On Error GoTo HandleError
    If Len(TextString) = 0 Then GoTo HandleExit
    If IsArray(vControls) Then
        aCtl = vControls
    ElseIf TypeOf vControls Is Collection Then
        Set aCtl = vControls
    ElseIf TypeOf vControls Is Access.Control Then
        aCtl = Array(vControls)
    Else
        Err.Raise vbObjectError + 512
    End If
Dim tDC As LongPtr:           tDC = GetDC(0)
Dim PIXEL_PER_INCH_X As Long: PIXEL_PER_INCH_X = GetDeviceCaps(tDC, LOGPIXELSX)
' разбиваем строку
    Call Tokenize(TextString, aWords, Separators)
    i = LBound(aWords): iMax = UBound(aWords)
    spLen = 1
    'strRest = Text
    ' костыль: vbCrLf меняем на vbCr иначе делает двойной разрыв строки
    strRest = Replace(TextString, vbCrLf, vbCr)
    For Each ctl In aCtl
        w = 0
        strText = vbNullString
        Do
        ' перебираем куски текста
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
            ' убираем мягкие переносы Chr(&HAD)
            strTemp = Replace(strText & strTemp, Chr(&HAD), vbNullString)
            ' строка равна предыдущей строке + текущий фрагмент + текущий разделитель
            strTemp = strTemp & Mid$(strRest, spPos, spLen)
            spLen = Len(Trim$(strTemp)): If spLen = 0 Then spLen = 1
        ' получаем размер текста
            hFont = p_HFontByControl(ctl)
            hOldFont = SelectObject(tDC, hFont)
            GetTextExtentPoint32 tDC, strTemp, spLen, sz
            SelectObject tDC, hOldFont
            DeleteObject hFont
        ' условие: w=0, - ещё один костыль
            WidthInPix = ctl.Width * (PIXEL_PER_INCH_X / TwipsPerInch)
            If sz.cX <= WidthInPix Or w = 0 Then
            ' если первое слово в строке меньше области печати - всё равно берём,
            ' иначе зависает в мертвом цикле
                If sz.cX > tWidth Then tWidth = sz.cX
                strRest = Mid$(strRest, spPosNext)
                strText = strTemp
                i = i + 1
                w = w + 1
            End If
        ' сравниваем размер текста с размером контрола
        Loop Until (i > iMax) Or (WidthInPix < sz.cX) '(WidthInPix < (sz.cx * (1 + spLen) / spLen))
        ctl.Value = strText
        'tHeight = tHeight + sz.cy
    Next ctl
' получаем разбитую строку и ее высоту в пикселях
'    WidthInPix = tWidth: HeightInPix = tHeight
'    GetTextMetrics tDC, tm
'    Overhang = tm.tmOverhang ' добавка для наклонных и толстых шрифтов
HandleExit:  SelectObject tDC, hOldFont
             DeleteObject hFont: ReleaseDC 0, tDC
             Exit Sub
HandleError: Err.Clear: Resume HandleExit
End Sub
Public Function TextToArrayByWidth(TextString As String, WidthInPix As Long, Optional HeightInPix, _
    Optional Separators As String = " ­.,;:!?()[]{}…+-*/\|" & vbTab & vbCrLf, _
    Optional OutLines, Optional Overhang As Long, Optional OutDelimiter = vbCrLf, _
    Optional hFont As LongPtr, Optional hdc As LongPtr = 0) As String
' , Optional OutLineWidth, Optional OutLineHeight
' разбивает текст с учётом допустимых разделителей на строки, соответствующие размерам заданонй области и параметрам шрифта
'-------------------------
' TextString - строка текста которую необходимо разбить
' WidthInPix - на входе - максимальная ширина разбитого текста,
'              на выходе - реальная высота разбитого текста
' HeightInPix - на выходе - реальная высота разбитого текста
' Separators - список разделителей по которым можно бить текст.
'       если включен символ мягкого переноса - в выходной строке будет опущен
' OutLines - массив строк разбитого текста
' Overhang - смещение для корректировки размера для наклонных, жирных и пр. шрифтов
' OutDelimiter - разделитель строк в выходной строке
' hFont - hFont шрифта для которого рассчитываем разбиение
' hDC - hDC -области куда будет выводиться текст
'' OutLineWidth, OutLineHeight - массивы размеров строк разбитого текста
'-------------------------
Dim Result As String
Dim sz As Size, tm As TEXTMETRIC
Dim strText As String, strRest As String, strTemp As String
Dim aWords() As String, aText() As String ', aWidth() As Long, aHeight() As Long
Dim spLen As Long, spPos As Integer, spPosNext As Integer
Dim tWidth As Long, tHeight As Long
Dim i As Long, iMax As Long, ii As Long, w As Long

    On Error GoTo HandleError
    Result = vbNullString
    tWidth = 0: tHeight = 0
    If hFont = 0 Then hFont = p_HFontByControl() 'GoTo HandleExit
    If WidthInPix < 0 Then GoTo HandleExit
Dim WidthMax As Long: WidthMax = WidthInPix - 2 ' костылик )
Dim tDC As LongPtr, hOldFont As LongPtr
    If hdc = 0 Then tDC = GetDC(0) Else tDC = hdc

Dim PIXEL_PER_INCH_X As Long: PIXEL_PER_INCH_X = GetDeviceCaps(tDC, LOGPIXELSX)
    hOldFont = SelectObject(tDC, hFont) 'select font into the DC
    
    Call Tokenize(TextString, aWords, Separators)
    i = LBound(aWords): iMax = UBound(aWords)
    ii = 0: spLen = 1
    'strRest = Text
    ' костыль: vbCrLf меняем на vbCr иначе делает двойной разрыв строки
    strRest = Replace$(TextString, vbCrLf, vbCr)
    Do
        w = 0
        strText = vbNullString
        Do
        ' перебираем куски текста
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
            ' убираем мягкие переносы Chr(&HAD)
            strTemp = Replace$(strText & strTemp, Chr(&HAD), vbNullString)
            ' строка равна предыдущей строке + текущий фрагмент + текущий разделитель
            strTemp = strTemp & Mid$(strRest, spPos, spLen)
            spLen = Len(Trim$(strTemp)): If spLen = 0 Then spLen = 1
        ' получаем размер текста
            GetTextExtentPoint32 tDC, strTemp, spLen, sz
        ' условие: w=0, - ещё один костыль
            If sz.cX <= WidthMax Or w = 0 Then
            ' если первое слово в строке меньше области печати - всё равно берём,
            ' иначе зависает в мертвом цикле
                If sz.cX > tWidth Then tWidth = sz.cX
                strRest = Mid$(strRest, spPosNext)
                strText = strTemp
                i = i + 1
                w = w + 1
            End If
        Loop Until (i > iMax) Or (WidthMax < sz.cX) '(WidthMax < (sz.cx * ((1 + spLen) / spLen))
        ReDim Preserve aText(ii): aText(ii) = Trim$(strText)
'        ReDim Preserve aWidth(ii): aWidth(ii) = sz.CX
'        ReDim Preserve aHeight(ii): aHeight(ii) = sz.CY
        tHeight = tHeight + sz.cY
'        Result = Result & OutDelimiter & strText
    ' если достигли конца - выходим
        If Len(strRest) = 0 Then Exit Do
        ii = ii + 1
    Loop
' получаем разбитую строку и ее высоту в пикселях
    WidthInPix = tWidth: HeightInPix = tHeight
    Result = Join(aText, OutDelimiter)
    OutLines = aText:           Erase aText
'    OutLineWidth = aWidth:      Erase aWidth
'    OutLineHeight = aHeight:    Erase aHeight
    GetTextMetrics tDC, tm
    Overhang = tm.tmOverhang ' добавка для наклонных и толстых шрифтов
    
HandleExit:  SelectObject tDC, hOldFont: If hdc = 0 Then ReleaseDC 0, tDC
             TextToArrayByWidth = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
#End If
' ==================
' Функции проверки/преобразования строк/символов
' ==================
Public Function TextTranslit(ByVal Source As String, Optional Direction As Byte = 0) As String
' транслитерация кириллицы для русского алфавита по ГОСТ Р 52535.1-2006
'-------------------------
' Direction = 0 - транслитерация (рус > лат)
'             1 - обратная транслитерация (лат > рус)
'-------------------------
Dim Result As String: Result = vbNullString
    On Error GoTo HandleError
Dim i As Integer, j As Integer, c As String: i = 1
Const cSymbRus = "щжхцчшюяаеёиоуыэйбвгдзклмнпрстфьъ"
Dim TransLat(): TransLat = Array("shch", "zh", "kh", "tc", "ch", "sh", "iu", "ia", "a", "e", "e", "i", "o", "u", "y", "e", "i", "b", "v", "g", "d", "z", "k", "l", "m", "n", "p", "r", "s", "t", "f", "", "")
Dim cLen As Integer: cLen = 1
    If Direction = 0 Then   ' рус >> лат
        Do Until i > Len(Source)
            c = Mid$(Source, i, cLen)           ' текущий рус символ
            j = InStr(1, cSymbRus, LCase$(c))    ' номер элемента массива
            If j > 0 Then j = j - 1: If c = LCase$(c) Then c = TransLat(j) Else c = StrConv(TransLat(j), vbProperCase)
            Result = Result & c
            i = i + cLen
        Loop
    Else                    ' лат >> рус
Const cSymbLat = "chjqwx" ''"
Dim TransRus(): TransRus = Array("ц", "х", "дж", "к", "в", "кс") ', "ь")
        Do Until i > Len(Source)
        ' проверяем по массиву транслитерации по ГОСТ
            For j = 0 To UBound(TransLat)       ' номер элемента массива
                c = TransLat(j): cLen = Len(c): If cLen = 0 Then Exit For
                If LCase(Mid$(Source, i, cLen)) = c Then c = Mid$(cSymbRus, j + 1, 1): Exit For
            Next j
            If Len(c) = 0 Then
        ' если не найдены соотв лат символы ГОСТ проверяем не вошедшие символы
                cLen = 1: c = Mid$(Source, i, cLen) ' текущий рус символ
                j = InStr(1, cSymbLat, LCase$(c))    ' номер элемента массива
                If j > 0 Then c = TransRus(j - 1)   ' если найден - берём
            End If
            If UCase$(Mid$(Source, i, 1)) = Mid$(Source, i, 1) Then c = StrConv(c, vbProperCase)
            Result = Result & c
            i = i + cLen
        Loop
    End If
HandleExit:  TextTranslit = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function TextCorrectLatRus(Source As String, Optional Direction As Byte = 0) As String
' Исправляет опечатку Lat<=>Rus
'-------------------------
' Direction = 0 - заменяет латинский символ соответствующим русским
' Direction = 1 - заменяет русский символ соответствующим латинским
'-------------------------
Dim Result As String
    Result = Source
    If Len(Source) = 0 Then GoTo HandleExit
Const cSymbLat As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const cSymbRus As String = "АВСДЕФГНИЖКЛМНОРОРСТУВВХУЗ"
Dim strSearch As String, strReplace As String
    If Direction = 0 Then   'Lat=>Rus
        strSearch = cSymbLat: strReplace = cSymbRus
    Else                    'Rus=>Lat
        strSearch = cSymbRus: strReplace = cSymbLat
    End If
Dim Char As String * 1, bUpper As Boolean
Dim c As Long, cMax As Long, i As Byte: c = 1: cMax = Len(Source)
    Do Until c > cMax
        Char = VBA.Mid$(Result, c, 1): bUpper = Char = VBA.UCase$(Char)
        i = InStr(1, strSearch, VBA.UCase$(Char)): If i = 0 Then GoTo HandleNext
        Char = VBA.Mid$(strReplace, i, 1): If bUpper Then Else Char = VBA.LCase$(Char)
        Mid$(Result, c, 1) = Char
HandleNext: c = c + 1
    Loop
HandleExit:  TextCorrectLatRus = Result
HandleError: Result = Source: Err.Clear: Resume HandleExit
End Function
Public Function TextContainsAlpha(Source As String) As Boolean
' проверяет наличие в строке символов невходящих в список допустимых
'-------------------------
Dim c As Long, cMax As Long
Dim Char As String * 1
Dim Result As Boolean
    Result = False
    On Error GoTo HandleError
' задаем разрешенные символы
Dim PermissedSymb As String: PermissedSymb = VBA.UCase$(c_strSymbRusAll & c_strOthers)  '(c_strOthers)
' пробегаем все символы пока не найдем первый не из списка
    c = 1: cMax = Len(Source)
    Do Until c > cMax
        Char = VBA.UCase$(VBA.Mid$(Source, c, 1))
        Select Case Char
        Case "A" To "Z", "a" To "z", "0" To "9"
        Case Else: If InStr(1, PermissedSymb, Char) = 0 Then Result = True: Exit Do
        End Select
        c = c + 1
    Loop
HandleExit:  TextContainsAlpha = Result
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function TextAlpha2Code(Source As String, _
    Optional Encoding As Byte = 0, _
    Optional Prefix As String = "%") As String
' Заменяет символы не входящие в список допустимых их шестнадцатиричным кодом вида %XX
'-------------------------
' Source   - кодируемая строка
' Encoding - тип кодировки 0-cp1251, 1-UTF-8, 2-URL код (как в поисковых запросах)
' Prefix   - префикс кода символа: "%","\u","=" или др
'-------------------------
Dim Result As String
    Result = vbNullString
    On Error GoTo HandleError
Dim c As Long, cMax As Long, cLen As Byte: c = 1: cMax = Len(Source): If cMax = 0 Then GoTo HandleExit
' определяем параметры кодирования
    Select Case Encoding
    Case 0: cLen = 2  ' cp1251   Prefix = "%"
    Case 1: cLen = 4  ' UTF-8    Prefix = "\u"
    Case 2: cLen = 2  ' URL код  Prefix = "%" или "="
    Case Else: Err.Raise vbObjectError + 512
    End Select
' задаем дополнительные разрешенные символы помимо a-z,A-Z и 0-9. можно вынести в параметр функции
Dim PermissedSymb As String: PermissedSymb = Replace(VBA.UCase$(c_strOthers), " ", "") '(c_strSymbRusAll & c_strOthers)  '(c_strOthers)
' пробегаем все символы строки
Dim Char As String, Code As String
    Do Until c > cMax
        Char = VBA.Mid$(Source, c, 1)
        Select Case Char
        Case "A" To "Z", "a" To "z", "0" To "9"
        Case Else: If InStr(1, PermissedSymb, VBA.UCase$(Char)) > 0 Then GoTo HandleNext
            Select Case Encoding
            Case 0: Code = VBA.Hex$(Asc(Char)):  Char = Prefix & String(cLen - Len(Code), "0") & Code
            Case 1: Code = VBA.Hex$(AscW(Char)): Char = Prefix & String(cLen - Len(Code), "0") & Code
            Case 2: Code = VBA.Hex$(AscW(Char)): Code = String(2 * cLen - Len(Code), "0") & Code
                    Select Case VBA.Left$(Code, cLen)
                    Case "00": Char = Prefix & VBA.Mid$(Code, cLen + 1)
                    Case "04": Char = Prefix & Hex$(&HD0 + (CLng(c_strHexPref & VBA.Mid$(Code, cLen + 1)) \ &H40)) & Prefix & Hex$(&H80 + (CLng(c_strHexPref & VBA.Mid$(Code, cLen + 1)) Mod &H40))
                    Case Else: Char = Prefix & VBA.Left$(Code, cLen) & Prefix & VBA.Mid$(Code, cLen + 1)
                    End Select
            End Select
        End Select
HandleNext:  Result = Result & Char: c = c + 1
    Loop
HandleExit:  TextAlpha2Code = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function TextCode2Alpha(Source As String, _
    Optional Encoding As Byte = 0, _
    Optional Prefix As String = "%") As String
' Заменяет код вида %XX, символов не входящих в список допустимых, их значением
'-------------------------
' Source   - декодируемая строка
' Encoding - тип кодировки 0-cp1251, 1-UTF-8, 2-URL код (как в поисковых запросах)
' Prefix   - префикс кода символа: "%","\u","=" или др
'-------------------------
Dim Result As String
    On Error GoTo HandleError
    Result = vbNullString
    On Error GoTo HandleError
Dim c As Long, cMax As Long, cLen As Byte: c = 1: cMax = Len(Source): If cMax = 0 Then GoTo HandleExit
' определяем параметры кодирования
    Select Case Encoding
    Case 0: cLen = 2  ' cp1251   Prefix = "%"
    Case 1: cLen = 4  ' UTF-8    Prefix = "\u"
    Case 2: cLen = 2  ' URL код  Prefix = "%" или "="
    Case Else: Err.Raise vbObjectError + 512
    End Select
' пробегаем все символы строки
Dim Char As String, Code As String, Cod2 As String
    Do Until c > cMax
        If VBA.Mid$(Source, c, Len(Prefix)) <> Prefix Then
' разрешенный (незакодированный) символ
            Char = VBA.Mid$(Source, c, 1)
        Else
' если находим управляющий символ - расшифровываем код
            Code = VBA.UCase$(VBA.Mid$(Source, c + Len(Prefix), cLen))
            Select Case Encoding
            Case 0: Code = c_strHexPref & Code: If IsNumeric(Code) Then Char = VBA.Chr$(Val(Code)):  c = c + cLen + Len(Prefix) - 1
            Case 1: Code = c_strHexPref & Code: If IsNumeric(Code) Then Char = VBA.ChrW$(Val(Code)): c = c + cLen + Len(Prefix) - 1
            Case 2: Code = c_strHexPref & Code: c = c + cLen + Len(Prefix) - 1
            ' символы U+0000..U+00FF >> %00..%FF
            ' символы U+0400..U+04FF >> %D0%80..%D0%BF;%D1%80..%D1%BF;%D2%80..%D2%BF;%D3%80..%D3%BF
                    If IsNumeric(Code) Then
                        Select Case CLng(Code)
                        Case &HD0 To &HD3:
                            If VBA.Mid$(Source, c + 1, Len(Prefix)) = Prefix Then
                                Cod2 = c_strHexPref & VBA.UCase$(VBA.Mid$(Source, c + Len(Prefix) + 1, cLen))
                                If IsNumeric(Cod2) Then
                                    Select Case CLng(Cod2)
                                    Case &H80 To &HBF: Code = c_strHexPref & Hex(Val(&H400 + &H40 * (Val(Code) Mod &HD0) + Val(Cod2) - &H80)): c = c + cLen + Len(Prefix) '- 1
                                    Case Else:         Code = Code & Mid$(Cod2, Len(Prefix) + 1)
                                    End Select
                                End If
                            End If
                        Case Else:
                        End Select
                    End If
                    Char = VBA.ChrW$(Val(Code))  ' прочие символы
            End Select
        End If
HandleNext:  Result = Result & Char: c = c + 1
    Loop
HandleExit:  TextCode2Alpha = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function TextClearAlpha( _
    Source, _
    Optional ReplaceWith As String = "_", _
    Optional AllowedSymbols As String)
' Сжимает строку, заменяя символы не входящие в список допустимых указанным символом
'-------------------------
' работает также с массивами строк.
'-------------------------
Dim c As Long, cMax As Long
Dim i As Long, iMax As Long
Dim strSource As String, strResult As String
Dim Char As String, PrevChar As String
Dim Result

    On Error GoTo HandleError
    Result = vbNullString
    If Len(AllowedSymbols) = 0 Then AllowedSymbols = c_strSymbDigits & c_strSymbRusAll & c_strSymbEngAll
    AllowedSymbols = UCase$(AllowedSymbols)
    If IsArray(Source) Then
        i = LBound(Source): iMax = UBound(Source)
        If iMax < i Then GoTo HandleExit
        ReDim Result(i To iMax)
        strSource = Source(i)
    Else
        i = 1: iMax = 1
        strSource = CStr(Source)
    End If
    Do
        cMax = Len(strSource)
        strResult = vbNullString
        For c = 1 To cMax
            Char = Mid$(strSource, c, 1)
            If InStr(1, AllowedSymbols, UCase$(Char)) = 0 Then
                If PrevChar <> ReplaceWith Then Char = ReplaceWith Else Char = vbNullString
                PrevChar = ReplaceWith
             Else
                PrevChar = Char
            End If
            strResult = strResult & Char
        Next c
        If Right$(strResult, Len(ReplaceWith)) = ReplaceWith Then strResult = Left$(strResult, Len(strResult) - Len(ReplaceWith))
        If Left$(strResult, Len(ReplaceWith)) = ReplaceWith Then strResult = Right$(strResult, Len(strResult) - Len(ReplaceWith))
        strResult = Trim$(strResult)
        If Len(strResult) < 1 Then strResult = vbNullString 'c_strEmptyString
        If IsArray(Source) Then
            Result(i) = strResult
            If i >= iMax Then Exit Do
            i = i + 1: strSource = Source(i)
         Else
            Result = strResult
            Exit Do
        End If
    Loop
HandleExit:  TextClearAlpha = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function GetCharType(Letter As String, Optional AlphaType As Byte = 0) As Byte
' проверяет букву и возвращает результат
'-------------------------
' на выходе: 1-гласная,2-согласная,3-знак(ьъ),и т.д.,0-не определено
' AlphaType = 1-буква латинского алфавита, 2-буква русского алфавита, 3-цифра, 0-иной символ
'-------------------------
Dim Result As Byte: Result = 0 ': SymbType = 0
Dim sChar As String * 1: sChar = LCase$(Left$(Trim$(Letter), 1))

    If InStr(c_strSymbDigits, sChar) > 0 Then Result = SymbolTypeNumb:      AlphaType = AlphabetTypeUndef:      GoTo HandleExit
    If InStr(c_strSymbRusVowel, sChar) > 0 Then Result = SymbolTypeVowel:   AlphaType = AlphabetTypeCyrilic:    GoTo HandleExit
    If InStr(c_strSymbRusConson, sChar) > 0 Then Result = SymbolTypeCons:   AlphaType = AlphabetTypeCyrilic:    GoTo HandleExit
    If InStr(c_strSymbRusSign, sChar) > 0 Then Result = SymbolTypeSign:     AlphaType = AlphabetTypeCyrilic:    GoTo HandleExit
    If InStr(c_strSymbEngVowel, sChar) > 0 Then Result = SymbolTypeVowel:   AlphaType = AlphabetTypeLatin:      GoTo HandleExit
    If InStr(c_strSymbEngConson, sChar) > 0 Then Result = SymbolTypeCons:   AlphaType = AlphabetTypeLatin:      GoTo HandleExit
    If InStr(c_strSymbEngSign, sChar) > 0 Then Result = SymbolTypeSign:     AlphaType = AlphabetTypeLatin:      GoTo HandleExit
'    If InStr(c_strSymbMath, sChar)>0 Then Result = : AlphaType = AlphabetTypeUndef:GoTo HandleExit
'    If InStr(c_strSymbPunct, sChar)>0 Then Result = : AlphaType = AlphabetTypeUndef:GoTo HandleExit
'    If InStr(c_strSymbCommas, sChar)>0 Then Result = : AlphaType = AlphabetTypeUndef:GoTo HandleExit
'    If InStr(c_strSymbParenth, sChar)>0 Then Result = : AlphaType = AlphabetTypeUndef:GoTo HandleExit
'    If InStr(c_strSymbOthers, sChar)>0 Then Result = SymbolTypeUndef:       AlphaType = AlphabetTypeUndef:      GoTo HandleExit
HandleExit: GetCharType = Result
End Function
' ==================
' Функции для сравнения слов
' ==================
Public Function PolyPhone(ByVal Word As String, Optional FuzzyIdx As Boolean = False) 'As String
'Polyphon: An Algorithm for Phonetic String Matching in Russian Language (Paramonov V.V., Shigarov A O., Ruzhnikov G.M. )
'-------------------------
' FuzzyIdx - если True  - возвращает числовой код для нечеткого сравнения
'            если False - возвращает фонетический код
'-------------------------
'http://td.icc.ru/files/papers/Paramonov_ICIST2016.pdf
'https://cyberleninka.ru/article/n/obzor-algoritmov-foneticheskogo-kodirovaniya
'-------------------------
Dim Result As String, x As Long: Result = vbNullString
    On Error GoTo HandleError
Dim i As Long, sChar As String
    For i = 1 To Len(Word)
        sChar = Mid$(Word, i, 1)
        Select Case sChar
        Case "B":                     sChar = "В" ' Подстановка вместо латинских букв схожих букв русского алфавита:
        Case "M":                     sChar = "М"
        Case "H":                     sChar = "Н"
        Case "A", "a":                sChar = "А"
        Case "E", "e":                sChar = "Е"
        Case "O", "o":                sChar = "О"
        Case "C", "c":                sChar = "С"
        Case "X", "x":                sChar = "Х"
        Case "Ъ", "ъ", "Ь", "ь":      sChar = vbNullString  ' Удаление букв Ь, Ъ.
        Case "А" To "Я", "а" To "я":  sChar = UCase$(sChar)
        Case Else:                    sChar = vbNullString  ' Удаление всех букв, не принадлежащих алфавиту русского языка.
        End Select
        ' Замена двух одинаковых букв одной.
        If sChar = UCase$(Mid$(Word, i + 1, 1)) Then i = i + 1: GoTo HandleNext
        ' Замена одиночных букв
        Select Case sChar
        Case "А", "Е", "Ё", _
             "И", "О", "Ы", "Э", "Я": sChar = "А"
        Case "Б":                     sChar = "П"
        Case "В":                     sChar = "Ф"
        Case "Г":                     sChar = "К"
        Case "Д":                     sChar = "Т"
        Case "З":                     sChar = "С"
        Case "Щ":                     sChar = "Ш"
        Case "Ж":                     sChar = "Ш"
        Case "М":                     sChar = "Н"
        Case "Ю":                     sChar = "У"
        End Select
        ' Выполнение подстановок:
        If Len(Result) > 3 Then
            Select Case Right$(Result, 4) & sChar
            Case "ЛФСТФ": Mid$(Result, Len(Result) - 3, 4) = "ЛСТФ": sChar = vbNullString: GoTo HandleNext
            End Select
        End If
        If Len(Result) > 2 Then
            Select Case Right$(Result, 3) & sChar
            Case "НТСК": Mid$(Result, Len(Result) - 2, 3) = "НCК": sChar = vbNullString: GoTo HandleNext
            Case "ФСТФ": Mid$(Result, Len(Result) - 2, 3) = "CТФ": sChar = vbNullString: GoTo HandleNext
            End Select
        End If
        If Len(Result) > 1 Then
            Select Case Right$(Result, 2) & sChar
            Case "ТАТ": Mid$(Result, Len(Result) - 1, 2) = "Т ": Result = Trim$(Result): sChar = vbNullString: GoTo HandleNext
            Case "ТСА": Mid$(Result, Len(Result) - 1, 2) = "Ц ": Result = Trim$(Result): sChar = vbNullString: GoTo HandleNext
            Case "ТСЯ": Mid$(Result, Len(Result) - 1, 2) = "Ц ": Result = Trim$(Result): sChar = vbNullString: GoTo HandleNext
            Case "НАТ": Mid$(Result, Len(Result) - 1, 2) = "Н ": Result = Trim$(Result): sChar = vbNullString: GoTo HandleNext
            Case "ТАФ": Mid$(Result, Len(Result) - 1, 2) = "ТФ": sChar = vbNullString: GoTo HandleNext
            Case "ФАК": Mid$(Result, Len(Result) - 1, 2) = "ФК": sChar = vbNullString: GoTo HandleNext
            Case "СТЛ": Mid$(Result, Len(Result) - 1, 2) = "CЛ": sChar = vbNullString: GoTo HandleNext
            Case "СТН": Mid$(Result, Len(Result) - 1, 2) = "CН": sChar = vbNullString: GoTo HandleNext
            Case "НТА": Mid$(Result, Len(Result) - 1, 2) = "НA": sChar = vbNullString: GoTo HandleNext
            Case "НТК": Mid$(Result, Len(Result) - 1, 2) = "НК": sChar = vbNullString: GoTo HandleNext
            Case "НТС": Mid$(Result, Len(Result) - 1, 2) = "НC": sChar = vbNullString: GoTo HandleNext
            Case "ЛНЦ": Mid$(Result, Len(Result) - 1, 2) = "НЦ": sChar = vbNullString: GoTo HandleNext
            Case "НТЦ": Mid$(Result, Len(Result) - 1, 2) = "НЦ": sChar = vbNullString: GoTo HandleNext
            Case "НТШ": Mid$(Result, Len(Result) - 1, 2) = "НШ": sChar = vbNullString: GoTo HandleNext
            Case "ПАЛ": Mid$(Result, Len(Result) - 1, 2) = "ПЛ": sChar = vbNullString: GoTo HandleNext
            Case "РТЧ": Mid$(Result, Len(Result) - 1, 2) = "PЧ": sChar = vbNullString: GoTo HandleNext
            Case "РТЦ": Mid$(Result, Len(Result) - 1, 2) = "PЦ": sChar = vbNullString: GoTo HandleNext
            Case "АКA": Mid$(Result, Len(Result) - 1, 2) = "AФ": sChar = "A": GoTo HandleNext
            Case "ОКО": Mid$(Result, Len(Result) - 1, 2) = "ОФ": sChar = "О": GoTo HandleNext
            End Select
        End If
        If Len(Result) > 0 Then
            Select Case Right$(Result, 1) & sChar
            Case "АН": Mid$(Result, Len(Result), 1) = "Н": sChar = vbNullString: GoTo HandleNext
            Case "ЗЧ": Mid$(Result, Len(Result), 1) = "Ш": sChar = vbNullString: GoTo HandleNext
            Case "НТ": Mid$(Result, Len(Result), 1) = "Н": sChar = vbNullString: GoTo HandleNext
            Case "СЧ": Mid$(Result, Len(Result), 1) = "Ш": sChar = vbNullString: GoTo HandleNext
            Case "СШ": Mid$(Result, Len(Result), 1) = "Ш": sChar = vbNullString: GoTo HandleNext
            Case "ТЦ": Mid$(Result, Len(Result), 1) = "Ц": sChar = vbNullString: GoTo HandleNext
            Case "ТЧ": Mid$(Result, Len(Result), 1) = "Ч": sChar = vbNullString: GoTo HandleNext
            Case "ШЧ": Mid$(Result, Len(Result), 1) = "Ш": sChar = vbNullString: GoTo HandleNext
            Case "СП": Mid$(Result, Len(Result), 1) = "C": sChar = "Ф": GoTo HandleNext
            Case "ТС": Mid$(Result, Len(Result), 1) = "Т": sChar = "Ц": GoTo HandleNext
            End Select
        End If
HandleNext:
        If FuzzyIdx Then
        ' добавляем числовые значения выходных символов
            Select Case sChar
            Case "А": x = x + 2
            Case "П": x = x + 3
            Case "К": x = x + 5
            Case "Л": x = x + 7
            Case "М": x = x + 11
            Case "Н": x = x + 13
            Case "Р": x = x + 17
            Case "С": x = x + 19
            Case "Т": x = x + 23
            Case "У": x = x + 29
            Case "Ф": x = x + 31
            Case "Х": x = x + 37
            Case "Ц": x = x + 41
            Case "Ч": x = x + 43
            Case "Щ": x = x + 47
            Case "Э": x = x + 53
            Case "Я": x = x + 59
            End Select
        End If
        Result = Result & sChar
    Next i
HandleExit:  PolyPhone = IIf(FuzzyIdx, x, Result): Exit Function
HandleError: Result = vbNullString: x = 0: Err.Clear: Resume HandleExit
End Function
Public Function MetaPhoneRu1(ByVal Word As String) As String
'Первоначальный вариант — простой, но не оптимальный.
'-------------------------
'Источник: http://forum.aeroion.ru/topic461.html
Const alf$ = "ОЕАИУЭЮЯПСТРКЛМНБВГДЖЗЙФХЦЧШЩЫЁ", _
      cns1$ = "БЗДВГ", _
      cns2$ = "ПСТФК", _
      cns3$ = "ПСТКБВГДЖЗФХЦЧШЩ", _
      ch$ = "ОЮЕЭЯЁЫ", _
      ct$ = "АУИИАИА"
'alf - алфавит кроме исключаемых букв, cns1 и cns2 - звонкие и глухие
'согласные, cns3 - согласные, перед которыми звонкие оглушаются,
'ch, ct - образец и замена гласных
Dim s$, v$, i&, b&, c$
'S, V - промежуточные строки, i - счётчик цикла, B - позиция
'найденного элемента, c$ - текущий символ

'Переводим в верхний регистр, оставляем только символы из alf
'приписываем пробел с начала, копируем в S:
    Word = UCase$(Word): s = " "
    For i = 1 To Len(Word)
        c = Mid$(Word, i, 1)
        If InStr(alf, c) Then s = s & c
    Next i
    If Len(s) = 1 Then Exit Function
    'Заменяем окончания:
    Select Case Right$(s, 6)
    Case "ОВСКИЙ":      s = Left$(s, Len(s) - 6) & "@"
    Case "ЕВСКИЙ":      s = Left$(s, Len(s) - 6) & "#"
    Case "ОВСКАЯ":      s = Left$(s, Len(s) - 6) & "$"
    Case "ЕВСКАЯ":      s = Left$(s, Len(s) - 6) & "%"
    End Select
    
    Select Case Right$(s, 3)
    Case "ОВА", "ЕВА":  s = Left$(s, Len(s) - 3) & "9"
    Case "ИНА":         s = Left$(s, Len(s) - 3) & "1"
    Case "НКО":         s = Left$(s, Len(s) - 3) & "3"
    End Select
    
    Select Case Right$(s, 2)
    Case "ОВ", "ЕВ":    s = Left$(s, Len(s) - 2) & "4"
    Case "АЯ":          s = Left$(s, Len(s) - 2) & "6"
    Case "ИЙ", "ЫЙ":    s = Left$(s, Len(s) - 2) & "7"
    Case "ЫХ", "ИХ":    s = Left$(s, Len(s) - 2) & "5"
    Case "ИН":          s = Left$(s, Len(s) - 2) & "8"
    Case "ИК", "ЕК":    s = Left$(s, Len(s) - 2) & "2"
    Case "УК", "ЮК":    s = Left$(s, Len(s) - 2) & "0"
    End Select
    'Оглушаем последний символ, если он - звонкий согласный:
    b = InStr(cns1, Right$(s, 1))
    If b Then Mid$(s, Len(s), 1) = Mid$(cns2, b, 1)
    'Основной цикл:
    For i = 2 To Len(s)
        c = Mid$(s, i, 1)
        b = InStr(ch, c)
        If b Then Mid$(s, i, 1) = Mid$(ct, b, 1) 'Замена гласных
        If InStr(cns3, c) Then 'Оглушение согласных
            b = InStr(cns1, Mid$(s, i - 1, 1))
            If b Then Mid$(s, i - 1, 1) = Mid$(cns2, b, 1)
        End If
    Next i
    'Устраняем повторы, убираем первый пробел:
    For i = 2 To Len(s)
        c = Mid$(s, i, 1)
        If c <> Mid$(s, i - 1, 1) Then v = v & c
    Next i
    MetaPhoneRu1 = v
End Function
Public Function MetaPhoneRu2(ByVal Word As String) As String
'Второй вариант — пожалуй, лучший.
'-------------------------
'Источник: http://forum.aeroion.ru/topic461.html
'Заменяет ЙО, ЙЕ и др.; неплохо оптимизирован.
Const alf$ = "ОЕАИУЭЮЯПСТРКЛМНБВГДЖЗЙФХЦЧШЩЁЫ", _
      cns1$ = "БЗДВГ", _
      cns2$ = "ПСТФК", _
      cns3$ = "ПСТКБВГДЖЗФХЦЧШЩ", _
      ch$ = "ОЮЕЭЯЁЫ", _
      ct$ = "АУИИАИА"
'alf - алфавит кроме исключаемых букв, cns1 и cns2 - звонкие и глухие
'согласные, cns3 - согласные, перед которыми звонкие оглушаются,
'ch, ct - образец и замена гласных
Dim s$, v$, i&, b&, c$, old_c$
'S, V - промежуточные строки, i - счётчик цикла, B - позиция
'найденного элемента, c$ - текущий символ, c_old$ - предыдущий
'символ

'Переводим в верхний регистр, оставляем только
'символы из alf и копируем в S:
    Word = UCase$(Word)
    For i = 1 To Len(Word)
        c = Mid$(Word, i, 1)
        If InStr(alf, c) Then s = s & c
    Next i
    If Len(s) = 0 Then Exit Function
    'Сжимаем окончания:
    Select Case Right$(s, 6)
    Case "ОВСКИЙ":                  s = Left$(s, Len(s) - 6) & "@"
    Case "ЕВСКИЙ":                  s = Left$(s, Len(s) - 6) & "#"
    Case "ОВСКАЯ":                  s = Left$(s, Len(s) - 6) & "$"
    Case "ЕВСКАЯ":                  s = Left$(s, Len(s) - 6) & "%"
    Case Else
        If Right$(s, 4) = "ИЕВА" Or Right$(s, 4) = "ЕЕВА" Then
            s = Left$(s, Len(s) - 4) & "9"
        Else
            Select Case Right$(s, 3)
            Case "ОВА", "ЕВА":      s = Left$(s, Len(s) - 3) & "9"
            Case "ИНА":             s = Left$(s, Len(s) - 3) & "1"
            Case "ИЕВ", "ЕЕВ":      s = Left$(s, Len(s) - 3) & "4"
            Case "НКО":             s = Left$(s, Len(s) - 3) & "3"
            Case Else
                Select Case Right$(s, 2)
                Case "ОВ", "ЕВ":    s = Left$(s, Len(s) - 2) & "4"
                Case "АЯ":          s = Left$(s, Len(s) - 2) & "6"
                Case "ИЙ", "ЫЙ":    s = Left$(s, Len(s) - 2) & "7"
                Case "ЫХ", "ИХ":    s = Left$(s, Len(s) - 2) & "5"
                Case "ИН":          s = Left$(s, Len(s) - 2) & "8"
                Case "ИК", "ЕК":    s = Left$(s, Len(s) - 2) & "2"
                Case "УК", "ЮК":    s = Left$(s, Len(s) - 2) & "0"
                End Select
            End Select
        End If
    End Select
    'Оглушаем последний символ, если он - звонкий согласный:
    b = InStr(cns1, Right$(s, 1))
    If b Then Mid$(s, Len(s), 1) = Mid$(cns2, b, 1)
    old_c = " "
    'Основной цикл:
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        b = InStr(ch, c)
        If b Then   'Если гласная
            If old_c = "Й" Or old_c = "И" Then
                If c = "О" Or c = "Е" Then 'Буквосочетания с гласной
                    old_c = "И": Mid$(v, Len(v), 1) = old_c
                Else 'Если не буквосочетания с гласной, а просто гласная
                    If c <> old_c Then v = v & Mid$(ct, b, 1)
                End If
            Else    'Если не буквосочетания с гласной, а просто гласная
                If c <> old_c Then v = v & Mid$(ct, b, 1)
            End If
        Else        'Если согласная
            If c <> old_c Then 'для «Аввакумов»
                If InStr(cns3, c) Then 'Оглушение согласных
                    b = InStr(cns1, old_c)
                    If b Then old_c = Mid$(cns2, b, 1): Mid$(v, Len(v), 1) = old_c
                End If
                If c <> old_c Then v = v & c 'для «Шмидт»
            End If
        End If
        old_c = c
    Next i
    MetaPhoneRu2 = v
End Function

Public Function MetaPhoneRu3(ByVal Word As String) As Long
'Третий вариант — с переводом ключа в 24-ричное число.
'-------------------------
'Источник: http://forum.aeroion.ru/topic461.html
'Единственный недостаток первых двух процедур — значительное место, необходимое для хранения
'строки-ключа. В некоторых реализациях SoundEx эту проблему решают, переводя ключ в число.

'Строка с ключом MetaPhoneRu из шести символов занимает в файле базы данных минимум 7 байт
'с завершающим нулём (без сжатия Unicode — 14 байт). Её можно перевести в число Long,
'занимающее только 4 байта.

'Например, имеем 24 возможных буквы в каждой позиции ключа. Переведём их в 24-ричную
'систему счисления так, например, что:
'«А» даст «1», «К» — «8», «Н» — «11». тогда
'«КА» превратится в 24 * 8 + 1 = 193;
'«КАН» — в 24 * (24 * 8 + 1) +  11 = 4643.

'Недостаток здесь — в MetaPhoneRu слишком велико число возможных значений буквы — 24.
'В числе типа Long удастся сохранить только 6 букв, и фамилии придётся усекать: «Колоколова»
'и «Колокольникова» различаться не будут оба усекаются до «КОЛОКО».

'Частично этот недостаток устаняется сжатием окончаний: «Кузьмина» и «Кузьминова» будут иметь
'разные ключи, хотя первые шесть букв у этих фамилий те же. Код сжатого окончание хранится
'в знаке ключа и в последней «половинной» цифре, появляющейся из-за того, что 24 не является
'степенью двойки. Между 24^6 (число комбинаций из 24-х символов на шести позициях) и 2^32
'(число возможных состояний 4-байтной переменной Long) остаётся небольшая разница, как бы
'«половинка» 24-ричной цифры. Эта «половинка» и используется для хранения окончания —
'она может принимать 2^32 / 24^6 = 22 различных состояния, из которых в процедуре используется
'четырнадцать.

'Реально применять на практике этот вариант MetaPhoneRu вряд ли стоит — выигрыш в размере поля
'с ключом всего в log2(256) / log2(24) = в 1,7 раз, но все фамилии длиннее шести символов в основе
'усекаются. Кроме того, программу сложнее понять и поддерживать: строковой ключ можно вывести
'на экран во время отладки и легко прочитать, а числовой код требует дополнительной расшифровки.
'-------------------------
Const alf$ = "АИУПСТРКЛМНБВГДЖЗЙФХЦЧШЩЕОЁЫЭЮЯ", _
      cns1$ = "БЗДВГ", _
      cns2$ = "ПСТФК", _
      cns3$ = "ПСТКБВГДЖЗФХЦЧШЩ", _
      ch$ = "ОЮЕЭЯЁЫ", _
      ct$ = "АУИИАИА"
'alf - алфавит кроме исключаемых букв, cns1 и cns2 - звонкие и глухие
'согласные, cns3 - согласные, перед которыми звонкие оглушаются,
'ch, ct - образец и замена гласных
Dim s$, v&, i&, b&, c$, old_c$, new_c$
'S - промежуточная строка, V — ключ, который создаётся в ходе
'работы процедуры, i - счётчик цикла, B - позиция найденного
'элемента, c$ - текущий символ, c_old$ - предыдущий символ,
'new_c$ — преобразованный текущий символ.

'Переводим в верхний регистр, оставляем только
'символы из alf и копируем в S:
    Word = UCase$(Word)
    For i = 1 To Len(Word)
        c = Mid$(Word, i, 1)
        If InStr(alf, c) Then s = s + c
    Next i
    If Len(s) = 0 Then Exit Function
    'Сжимаем окончания:
    Select Case Right$(s, 6)
    Case "ОВСКИЙ": s = Left$(s, Len(s) - 6): v = -1
    Case "ЕВСКИЙ": s = Left$(s, Len(s) - 6): v = -2
    Case "ОВСКАЯ": s = Left$(s, Len(s) - 6): v = -3
    Case "ЕВСКАЯ": s = Left$(s, Len(s) - 6): v = -4
    Case Else
        Select Case Right$(s, 4)
        Case "ИЕВА", "ЕЕВА": s = Left$(s, Len(s) - 4): v = 9
        Case Else
            Select Case Right$(s, 3)
            Case "ОВА", "ЕВА": s = Left$(s, Len(s) - 3): v = 9
            Case "ИНА": s = Left$(s, Len(s) - 3): v = 1
            Case "ИЕВ", "ЕЕВ": s = Left$(s, Len(s) - 3): v = 4
            Case "НКО": s = Left$(s, Len(s) - 3): v = 3
            Case Else
                Select Case Right$(s, 2)
                Case "ОВ", "ЕВ": s = Left$(s, Len(s) - 2): v = 4
                Case "АЯ": s = Left$(s, Len(s) - 2): v = 6
                Case "ИЙ", "ЫЙ": s = Left$(s, Len(s) - 2): v = 7
                Case "ЫХ", "ИХ": s = Left$(s, Len(s) - 2): v = 5
                Case "ИН": s = Left$(s, Len(s) - 2): v = 8
                Case "ИК", "ЕК": s = Left$(s, Len(s) - 2): v = 2
                Case "УК", "ЮК": s = Left$(s, Len(s) - 2): v = -5
                End Select
            End Select
        End Select
    End Select
    'Оглушаем последний символ, если он - звонкий согласный:
    b = InStr(cns1, Right$(s, 1))
    If b Then Mid$(s, Len(s), 1) = Mid$(cns2, b, 1)
    old_c = " "
    s = Left$(s, 6)
    'Основной цикл:
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        b = InStr(ch, c)
        If b Then 'Если гласная
            If old_c = "Й" Or old_c = "И" Then
                If c = "О" Or c = "Е" Then 'Буквосочетания с гласной
                    old_c = "И"
                Else 'Если не буквосочетания с гласной, а просто гласная
                    If c <> old_c Then new_c = Mid$(ct, b, 1)
                End If
            Else 'Если не буквосочетания с гласной, а просто гласная
                If c <> old_c Then new_c = Mid$(ct, b, 1)
            End If
        Else 'Если согласная
            If c <> old_c Then 'для «Аввакумов»
                If InStr(cns3, c) Then 'Оглушение согласных
                    b = InStr(cns1, old_c)
                    If b Then old_c = Mid$(cns2, b, 1)
                End If
                If c <> old_c Then new_c = c 'для «Шмидт»
            End If
        End If
        old_c = c
        v = v * 24                'Новая цифра в 24-ричном числе V —
        v = v + InStr(alf, new_c) 'порядковый номер буквы new_c в alf$.
        
        'Первые 24 символа в alf — это возможные значения каждой
        'буквы ключа. После них в alf приведены те гласные, которые
        'заменяются всегда и в конечном ключе не присутствуют.
    Next i
    MetaPhoneRu3 = v
End Function
Public Function SoundEx(ByVal Word As String) As String
' русский Russell (NARA) Soundex, алгоритм с учетом cмежных W и H
'-------------------------
' Источник: http://forum.aeroion.ru/topic443.html
' также:    http://www.source-code.biz/snippets/vbasic/4.htm
'-------------------------
Dim s As String, l As Long
Dim Result As String
    s = Trim$(UCase$(Word)): l = Len(s)
    If l = 0 Then Result = String$(4, 0): GoTo HandleExit
Const RusTab = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЫЭЮЯЬЪ"  ' для упрощённой транслитерации
Const LatTab = "ABVGDEEGZIIKLMNOPRSTUFHCHHHIAWAXY"  '(для транслитерации по ГОСТ см.TextTranslit)
Const CodeTab = "01230120022455012623010202"        ' для кодирования по правилам Soundex
'               "ABCDEFGHIJKLNMOPQRSTUVWXYZ"        ' соответствующие кодам символы
'' альтернативная для кириллицы без транслитерации из: https://cyberleninka.ru/article/n/obzor-algoritmov-foneticheskogo-kodirovaniya
'Const CodeTab = "012460033074788019360235555000000"
''               "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЫЭЮЯЬЪ"
Dim h As Long:  h = 0   ' код текущего символа
Dim lH As Long: lH = -1 ' код предыдущего символа. -1 - первый символ (нет предыдущего)
Dim i As Long:  i = 1   ' текущая позиция во входной строке
Dim c As String         ' текущий символ
Dim p As Integer        ' позиция текущего символа в строке транслитерации
    Do Until i > l
        c = Mid$(s, i, 1)
    ' простая транслитерация
        p = InStr(1, RusTab, c): If p > 0 Then c = Mid$(LatTab, p, 1)
    ' проверка допустимых символов
        If InStr(1, LatTab, c) = 0 Then GoTo HandleNext
    ' получаем Soundex код символа
        h = Mid$(CodeTab, Asc(c) - 64, 1)
        If lH <> h Then
    ' cмежные символы, или символы, разделенные буквами H или W,
    ' входящие в одну и ту же группу, записываются как один
            If lH = -1 Then
    ' первый символ в строке
                Result = c: lH = h
            Else
                If h = 0 Then
    ' не альфа символ: "HWAEIOUY" - пропуск (в русском пропускаем "УЕЁЫАОЭЯИЮЪЬ")
                    If InStr(1, "HW", c) = 0 Then lH = h '
                    GoTo HandleNext
                End If
    ' запоминаем предыдущий Soundex код и добавляем полученный к результирующей строке
                lH = h: Result = Result & h
            End If
        End If
    ' если длина Soundex кода >4 - выходим
        If Len(Result) >= 4 Then Exit Do
HandleNext: i = i + 1    ' следующий символ
    Loop
    Result = Result & String$(4 - Len(Result), "0")
HandleExit: SoundEx = Result
End Function
Public Function SoundEx2(ByVal Word As String) As String
' русский Refined Soundex
'-------------------------
' Источник: https://habr.com/ru/post/114947/
'-------------------------
Dim s As String, l As Long
Dim Result As String
    s = Trim$(UCase$(Word)): l = Len(s)
    If l = 0 Then Result = String$(4, 0): GoTo HandleExit
Const RusTab = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЫЭЮЯЬЪ"  ' для транслитерации
Const LatTab = "ABVGDEEGZIIKLMNOPRSTUFHCHHHIAWAXY"
'               "ABCDEFGHIJKLNMOPQRSTUVWXYZ"
Const CodeTab = "01360240043788015936020505"
' схема кодирования по правилам Refined Soundex русских символов без транслитерации
' по https://cyberleninka.ru/article/n/obzor-algoritmov-foneticheskogo-kodirovaniya
'Const CodeTab = "01246003307478801936023555500000"
''               "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЬЭЮЯ"
Dim h As Long:  h = 0   ' код текущего символа
Dim lH As Long: lH = -1 ' код предыдущего символа. -1 - первый символ (нет предыдущего)
Dim i As Long:  i = 1   ' текущая позиция во входной строке
Dim c As String         ' текущий символ
Dim p As Integer        ' позиция текущего символа в строке транслитерации
    Do Until i > l
        c = Mid$(s, i, 1)
    ' простая транслитерация
        p = InStr(1, RusTab, c): If p > 0 Then c = Mid$(LatTab, p, 1)
    ' проверка допустимых символов
        If InStr(1, LatTab, c) = 0 Then GoTo HandleNext
    ' получаем Soundex код символа
        h = Mid$(CodeTab, Asc(c) - 64, 1)
    ' входящие в одну и ту же группу, записываются как один
        If lH = h Then GoTo HandleNext
    ' первый символ в строке
        If lH = -1 Then Result = c
    ' запоминаем предыдущий Soundex код и добавляем полученный к результирующей строке
        lH = h: Result = Result & h
HandleNext: i = i + 1  ' следующий символ
    Loop
HandleExit: SoundEx2 = Result
End Function
Public Function SoundExDM(ByVal Word As String) As String
' SoundEx - Daitch-Mokotoff
'-------------------------
' описание алгоритма: http://www.avotaynu.com/soundex.htm
' для проверки результатов онлайн: https://stevemorse.org/census/soundex.html
' есть разница для "rs":    "Halberstadt"   мой "587943 587433", эталон - "587943"
'                           "Peters"        мой "739400 734000", эталон - "739400"
' т.е. эталон выдает вариант только для формы "rtz"
' в тоже время в https://cyberleninka.ru/article/n/obzor-algoritmov-foneticheskogo-kodirovaniya
' для последнего примера даны результаты "739400 734000"
' также мой вариант рассматривает rs как коды 94 и 4 , а эталон как 9,4 (и 4)
' т.е мой при анализе повторов мой будет сравнивать следующий с 94, а эталон с 4
' и вообще требует оптимизации
'-------------------------
Const jMax = 6      ' максимальное количество символов в выходном коде
Const nMax = 7      ' максимальная длина распознаваемого паттерна
Const Delim = " "   ' разделитель альтернатив в результирующей строке
Dim Result As String: Result = vbNullString 'String$(jMax, "0")
    On Error GoTo HandleError
    Word = Trim$(Word): If Len(Word) = 0 Then Err.Raise vbObjectError + 512 'GoTo HandleExit
    
    Word = LCase$(Replace(TextTranslit(Word), " ", ""))     ' подготавливаем слово - транслитерация, убираем пробелы и нижний регистр
Dim bolFound As Boolean                                     ' признак найденного паттерна
Dim i As Integer, iMax As Integer: i = 1: iMax = Len(Word)  ' индекс символа текста
Dim jEnd As Integer, jBeg As Integer                        ' позиция в результирующей строке
Dim r As Integer, rMax As Integer: rMax = 1                 ' индекс варианта результата
Dim rOld As Integer: rOld = 1                               ' предыдущее значение количества результатов
Dim sPart As String, n As Integer                           ' распознаваемый фрагмент и его длина
Dim Code, sCode As String, с As Integer, cMax As Integer: cMax = 0  ' текущий(ие) коды
Dim Prev, sPrev As String, p As Integer, pMax As Integer: pMax = 0  ' предыдущий(ие) коды
    Prev = vbNullString
    Do Until (i > iMax) 'Or (jEnd > jMax)
        If Len(Result) = rMax * (jMax + 1) - 1 Then Exit Do ' отсекаем ненужные проверки, когда все коды уже достигли максимальной возможной длины. вообще-то больше быть не должно
        cMax = 0                                ' количество найденных вариантов кода -1
        bolFound = False: Code = vbNullString   ' сбрасываем найденый фрагмент
        n = iMax - i + 1: If n > nMax Then n = nMax   ' выбираем максимальную возможную длину фрагмента
        Do Until n < 1
    ' перебираем допустимые фрагменты по длине по убыванию
        ' при нахождении фрагмента:
            '1. устанавливаем флаг bolFound=True
            '2. определяем место фрагмента:
            '       [Н] вначале слова (i=1);
            '       [Г] перед гласной (InStr(1, c_strSymbEngVowel, Mid$(Word, i + n, 1)) > 0),
            '       [О] все сотальные случаи
            '3. в зависимости от места возвращаем код фрагмента (Code) результата
            ' для варианта с альтернативами (см. "ch","ck" и т.д.) массив альтернатив
            
            sPart = Mid$(Word, i, n)
            Select Case n
            Case 7: Select Case sPart
                Case "schtsch": bolFound = True: If i > 1 Then Code = "4" Else Code = "2"
                End Select
            Case 6: Select Case sPart
                Case "schtsh", "schtch": bolFound = True: If i > 1 Then Code = "4" Else Code = "2"
                End Select
            Case 5: Select Case sPart
                Case "shtch", "shtsh", "stsch", "zhdzh": bolFound = True: If i > 1 Then Code = "4" Else Code = "2"
                Case "ttsch": bolFound = True: Code = "4"
                End Select
            Case 4: Select Case sPart
                Case "shch", "stch", "strz", "strs", "stsh", "szcz", "szcs", "zdzh": bolFound = True: If i > 1 Then Code = "4" Else Code = "2"
                Case "scht", "schd": bolFound = True: If i > 1 Then Code = "43" Else Code = "2"
                Case "ttch", "tsch", "ttsz", "zsch": bolFound = True: Code = "4" '4,4,4
                End Select
            Case 3: Select Case sPart
                Case "zdz": bolFound = True: If i > 1 Then Code = "4" Else Code = "2"
                Case "sht", "szt", "shd", "szd", "zhd": bolFound = True: If i > 1 Then Code = "43" Else Code = "2"
                Case "csz", "czs", "drz", "drs", "dsh", "dsz", "dzh", "dzs", "sch", "tch", "trz", "trs", "tsh", "tts", "ttz", "tzs", "tsz", "zsh": bolFound = True: Code = "4"
                Case "chs": bolFound = True: If i > 1 Then Code = "54" Else Code = "5"
                End Select
            Case 2: Select Case sPart
                Case "ai", "aj", "ay", "ei", "ej", "ey", "oi", "oj", "oy", "ui", "uj", "uy", "ue": bolFound = True: If i = 1 Then Code = "0" Else If InStr(1, c_strSymbEngVowel, Mid$(Word, i + n, 1)) > 0 Then Code = "1" 'Else Code = vbNullString
                Case "ia", "ie", "io", "iu": bolFound = True: If i = 1 Then Code = "1" 'Else Code = vbNullString
                Case "dt", "th": bolFound = True: Code = "3"
                Case "sc": bolFound = True: If i > 1 Then Code = "4" Else Code = "2"
                Case "st", "sd", "zd": bolFound = True: If i > 1 Then Code = "43" Else Code = "2"
                Case "cz", "cs", "ds", "dz", "sh", "ts", "tc", "tz", "sz", "zh", "zs", "rz": bolFound = True: Code = "4"
                Case "rs", "rz": bolFound = True: Code = Array("94", "4"): cMax = UBound(Code)  '"rs">"rtz","zh"
                Case "ch": bolFound = True: Code = Array("5", "4"): cMax = UBound(Code)   '"ch">"kh","tch"
                Case "ck": bolFound = True: Code = Array("5", "45"): cMax = UBound(Code)  '"ck">"k","tsk"
                Case "kh": bolFound = True: Code = "5"
                Case "ks": bolFound = True: If i > 1 Then Code = "54" Else Code = "5"
                Case "mn", "nm": bolFound = True: Code = "66"
                Case "pf", "ph", "fb": bolFound = True: Code = "7"
                Case "au": bolFound = True: If i = 1 Then Code = "0" Else If InStr(1, c_strSymbEngVowel, Mid$(Word, i + n, 1)) > 0 Then Code = "7"  'Else Code = vbNullString
                Case "eu": bolFound = True: If i = 1 Or InStr(1, c_strSymbEngVowel, Mid$(Word, i + n, 1)) > 0 Then Code = "1" 'Else Code = vbNullString
                End Select
            Case 1: Select Case sPart
                Case "a", "e", "i", "o", "u": bolFound = True: If i = 1 Then Code = "0" 'Else Code = vbNullString
                Case "y": bolFound = True: If i = 1 Then Code = "1" 'Else Code = vbNullString
                Case "j": bolFound = True: If i = 1 Then Code = Array("4", "1"): cMax = UBound(Code) '"j" >"dzh","y"
                Case "d", "t": bolFound = True: Code = "3"
                Case "s", "z": bolFound = True: Code = "4"
                Case "h": bolFound = True: If i = 1 Or InStr(1, c_strSymbEngVowel, Mid$(Word, i + n, 1)) > 0 Then Code = "5" 'Else Code = vbNullString  '5,5,-1
                Case "g", "k", "q": bolFound = True: Code = "5"
                Case "c": bolFound = True: Code = Array("5", "4"): cMax = UBound(Code) '"c" >"tz", "k"
                Case "x": bolFound = True: If i > 1 Then Code = "54" Else Code = "5" '5,54,54
                Case "m", "n": bolFound = True: Code = "6"
                Case "b", "f", "p", "v", "w": bolFound = True: Code = "7"
                Case "l": bolFound = True: Code = "8"
                Case "r": bolFound = True: Code = "9"
                End Select
            Case Else: Err.Raise vbObjectError + 512
            End Select
'!!! все ниже написанное надо переписать заново !!!
' Запись кода в результат.
    ' используемый вариант со строкой результата упрощает его дублирование при обнаружении альтернатив
    ' (дублирование - сложением строк), но требует более сложного кода для контроля позиции
    ' и, возможно, устранения повторов в итоговой строке (не реализовано)
    ' проще использовать коллекцию или массив это позволит не искать каждый раз начало/конец фрагмента
    ' дублирование результатов при обнаружении альтернативных вариантов кода придется делать перебором
    ' коллекция, также, позволит вставлять элементы в нужную позицию при дублировании
    ' и исключить повторяющиеся элементы в результате отловом ошибки при присвоении индекса,
    ' но.. - мы не ищем лёгких путей)
            If bolFound Or ((n = 1) And (i = iMax)) Then Else GoTo HandleNext
    ' если фрагмент найден и распознан записываем его код в результат
    ' если это не распознанный символ в конце строки - дополняем нулями до длины кода (jMax)
        
        ' если предыдущий код имел единственный вариант берём его иначе берём первый из массива
            If pMax = 0 Then sPrev = Prev Else sPrev = Prev(p)
        ' если единственный вариант кода берём его и переходим к формированию результирующей строки
            If cMax = 0 Then sCode = Code: GoTo HandleMakeResult
        ' если есть альтернативные варианты кода символа
            sCode = Code(0)          ' берём первый из массива кодов
            rOld = rMax              ' запоминаем предыдущее количество результатов (нужно для определения конца старой последовательности)
            rMax = rOld * (cMax + 1) ' определяем новое максимальное количество альтернативных результатов
        ' клонируем результирующую строку и строку позиций
Dim c As Long: For c = 1 To cMax: Result = Result & Delim & Result: Next
HandleMakeResult:
        ' перебираем все результаты
            jBeg = 1: jEnd = 0  ' начало/конец текущего варианта кода в результирующей строке
            r = 1: c = 0: p = 0 ' индексы результата, кода и предыдущего кода
            Do
        ' получаем позицию окончания текущего варианта результата (окончание предыдущего + длина разделителя)
            ' если несколько результатов - ищем её по позиции следующего разделителя
            ' если один результат или разделитель не найден - равна длине строки+1
                If rMax > 1 Then jEnd = VBA.InStr(jBeg, Result, Delim) - 1
                If jEnd <= 0 Then jEnd = Len(Result) + 1
            ' пропускаем соседние повторяющиеся коды кроме MN/NM (66)
                If (sCode = sPrev) And (sCode <> "66") Then sCode = vbNullString
        ' формируем r-тый результат и обрезаем его по jMax
                If rMax = 1 Then sCode = Result & sCode Else sCode = Mid$(Result, jBeg, jEnd - jBeg + 1) & sCode
                sCode = Left$(sCode, jMax)
            ' если номер следующего символа слова больше длины слова (i+n>=iMax)
            ' дополняем текущий вариант результата нулями до необходимой длинны(jMax)
                If (i + n) > iMax Then sCode = sCode & String(jMax - Len(sCode), "0")
        ' в множественных результатах можно устранить повторы
            ' для этого делаем обратный просмотр Result от jBeg к началу по разделителями
            ' сравниаваем предыдущие коды с текущим фрагментом sCode
            ' если находим полностью совпадающий с текущим - удаляем текущий фрагмент,
            '!это потребует смены алгоритма обхода индексов вариантов!
            ' мы этого делать не будем
        ' записываем r-тый результат.
                If rMax = 1 Then Result = sCode Else Result = VBA.Left$(Result, jBeg - 1) & sCode & VBA.Mid$(Result, jEnd + 1)
                ' сдвигаем позицию начала текущего варианта результата
                jBeg = jBeg + Len(sCode) + Len(Delim) '- 1
' обход последовательностей - увеличиваем индексы вариантов и предыдущих вариантоы
    ' т.к. новые варианты образуются путем добавления старых в конец строки -
    ' при переборе чередование индексов будет следующее:
        ' последовательно перебирая варианты результата,
        ' для каждого текущего кода перебираем все варианты предыдущего,
        ' а затем переходим к следующему варианту текущего
        
            ' если достигли конца старой (дублированной) последовательности -
            ' увеличиваем индекс текущего кода (который сравниваем) и получаем его значение
                If r Mod rOld = 0 Then c = c + 1: c = IIf(c > cMax, 0, c): If cMax = 0 Then sCode = Code Else sCode = Code(c)    '
            ' увеличиваем индекс предыдущего кода (с которым сравниваем) и получаем его значение
                p = p + 1: p = IIf(p > pMax, 0, p): If pMax = 0 Then sPrev = Prev Else sPrev = Prev(p)
            ' увеличиваем индекс результата, если перебрали все - выходим
                r = r + 1: If r > rMax Then Exit Do
            Loop
            ' запоминаем предыдущий(ие) код(ы) и выходим из цикла
            Prev = Code: pMax = cMax: Exit Do
HandleNext: If n > 1 Then n = n - 1 Else Exit Do
        Loop
        i = i + n ' следующий символ
    Loop
HandleExit:  SoundExDM = Result: Exit Function
HandleError: Result = String$(jMax, "0"): Err.Clear: Resume HandleExit
End Function

Public Function SimilarityLCS(ByVal Word1 As String, ByVal Word2 As String) As Double
' Наибольшая общая подпоследовательность - позволяет только вставки и удаления, но не замену
'-------------------------
' Longest Common Subsequence by George Brown - 11/14/04
' http://forums.devarticles.com/microsoft-access-development-49/approximate-string-matching-fyi-46951.html
'-------------------------
Dim i As Integer, j As Integer, m As Integer, n As Integer
Dim c() As Integer, b() As Integer, x$(), y$()
Dim Len1 As Integer, Len2 As Integer
Dim Result As Double

    Result = False
    On Error GoTo HandleError
    
    Word1 = UCase$(Trim$(Word1)): Word2 = UCase$(Trim$(Word2))
    Len1 = Len(Word1): Len2 = Len(Word2)
    
    n = IIf(Len1 <= Len2, Len1, Len2) 'n = min(Len1, Len2)
    m = IIf(Len1 >= Len2, Len1, Len2) 'm = max(Len1, Len2)
    
    ReDim x(m): ReDim y(m)
    ReDim c(m, m): ReDim b(m, m)
    '
    If Len1 > Len2 Then
        For i = 1 To Len1: x(i) = Mid$(Word1, i, 1): Next i
        For i = 1 To Len2: y(i) = Mid$(Word2, i, 1): Next i
    Else
        For i = 1 To Len2: x(i) = Mid$(Word2, i, 1): Next i
        For i = 1 To Len1: y(i) = Mid$(Word1, i, 1): Next i
    End If
    
    For i = 1 To m
        For j = 1 To n
            If x(i - 1) = y(j - 1) Then
                c(i, j) = c(i - 1, j - 1) + 1:  b(i, j) = 1
            ElseIf c(i - 1, j) >= c(i, j - 1) Then
                c(i, j) = c(i - 1, j):          b(i, j) = 2
            Else
                c(i, j) = c(i, j - 1):          b(i, j) = 3
            End If
        Next j
    Next i
    'If c(m, n)>0 Then Result = c(m, n)  ' Наибольшая общая подпоследовательность
    If c(m, n) > 0 Then Result = 2 * c(m, n) / (Len1 + Len2) ' c(m, n) / IIf(Len1 > Len2, Len1, Len2)
    Erase x: Erase y: Erase c: Erase b
HandleExit:  SimilarityLCS = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function SimilarityLev(ByVal Word1 As String, ByVal Word2 As String) As Double
' Расстояние Левенштейна - позволяет вставку, удаление и замену символов
'-------------------------
' Источник: http://qaru.site/questions/390285/finding-similar-sounding-text-in-vba
' также:    https://ru.wikibooks.org/wiki/Реализации_алгоритмов/Расстояние_Левенштейна#Visual_Basic_6.0
Dim Result As Double

    Result = False
    On Error GoTo HandleError
    If Word1 = vbNullString Or Word2 = vbNullString Then GoTo HandleExit
    Word1 = UCase$(Trim$(Word1)): Word2 = UCase$(Trim$(Word2))
    If Word1 = Word2 Then Result = 1: GoTo HandleExit
    Dim m As Integer: m = Len(Word1)
    Dim n As Integer: n = Len(Word2)
    Dim d As Variant: ReDim d(m, n)
    
    Dim i As Integer, j As Integer, k As Integer
    For i = 1 To m: d(i, 0) = i: Next i
    For j = 1 To n: d(0, j) = j: Next j
    
    Dim a(2), r As Integer, cost As Integer
    For i = 1 To m
        For j = 1 To n
            If Mid$(Word1, i, 1) = Mid$(Word2, j, 1) Then cost = 0 Else cost = 1
            a(0) = d(i - 1, j) + 1          ' удаление
            a(1) = d(i, j - 1) + 1          ' вставка
            a(2) = d(i - 1, j - 1) + cost   ' подстановка
            ' выбираем наименьшее
            r = a(0)
            For k = 1 To 2
                If a(k) < r Then r = a(k)
            Next k
            d(i, j) = r
        Next j
    Next i
'    Result = d(m, n)                       ' расстояние Левенштейна
    Result = 1 - d(m, n) / IIf(m > n, m, n) ' расстояние нормированное по длине строки
HandleExit:  SimilarityLev = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function SimilarityDL(ByVal Word1 As String, ByVal Word2 As String) As Double
' Расстояние Дамерау-Левенштейна - позволяет вставку, удаление, замену и перестановку двух соседних символов
'-------------------------
' Источник: https://ru.wikipedia.org/wiki/Расстояние_Дамерау_—_Левенштейна
'-------------------------
Dim Result As Double ': Result = False
    On Error GoTo HandleError
    If Word1 = vbNullString Or Word2 = vbNullString Then GoTo HandleExit
    Word1 = UCase$(Trim$(Word1)): Word2 = UCase$(Trim$(Word2))
    If Word1 = Word2 Then Result = 1: GoTo HandleExit
Dim m As Integer: m = Len(Word1)
Dim n As Integer: n = Len(Word2)
Dim d As Variant: ReDim d(m, n)

Dim i As Integer, j As Integer, k As Integer
    For i = 1 To m: d(i, 0) = i: Next i
    For j = 1 To n: d(0, j) = j: Next j
    
Dim a(3), r As Integer, cost As Integer
    For i = 1 To m
        For j = 1 To n
            If Mid$(Word1, i, 1) = Mid$(Word2, j, 1) Then cost = 0 Else cost = 1
            a(0) = d(i - 1, j) + 1          ' удаление
            a(1) = d(i, j - 1) + 1          ' вставка
            a(2) = d(i - 1, j - 1) + cost   ' подстановка
                                            ' перестановка
            If i And j And Mid$(Word1, i + 1, 1) = Mid$(Word2, j, 1) _
                 And Mid$(Word1, i, 1) = Mid$(Word2, j + 1, 1) Then _
            a(3) = d(i - 2, j - 2) + cost Else a(3) = &H7FFF
            
            ' выбираем наименьшее
            r = a(0)
            For k = 1 To 3
                If a(k) < r Then r = a(k)
            Next k
            d(i, j) = r
        Next j
    Next i
'    Result = d(m, n) ' расстояние Дамерау-Левенштейна
    Result = 1 - d(m, n) / IIf(m > n, m, n)
HandleExit:  SimilarityDL = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function SimilarityDice(ByVal Word1 As String, ByVal Word2 As String) As Double
' Сходство Дайса
'-------------------------
' https://en.wikibooks.org/wiki/Algorithm_Implementation/Strings/Dice%27s_coefficient
'-------------------------
Const n = 2 ' используем биграммы
Dim Result As Double ': Result = False
    On Error GoTo HandleError
    If Word1 = vbNullString Or Word2 = vbNullString Then GoTo HandleExit
    Word1 = UCase$(Trim$(Word1)): Word2 = UCase$(Trim$(Word2))
    If Word1 = Word2 Then Result = 1: GoTo HandleExit
Dim i As Integer, j As Integer
Dim iMax As Integer: iMax = Len(Word1) - (n - 1)    ' n-грамм в Word1
Dim jMax As Integer: jMax = Len(Word2) - (n - 1)    ' n-грамм в Word2
    If (iMax < 1) Or (jMax < 1) Then GoTo HandleExit    ' слово меньше длины n-граммы
' формируем коллекцию индексов n-грамм для Word2 (не хочется использовать объекты, но из массива сложно удалять)
Dim Col As New Collection, с As Long: For j = 1 To jMax: Col.Add j: Next j
' сопоставляем n-граммы Word1 с Word2
Dim x As Integer: x = 0
    For i = 1 To iMax
        For j = 1 To Col.Count
' при совпадении n-грамм:
    ' увеличиваем счётчик совпадений,
    ' вычёркиваем индекс из перебора n-грамм Word2
    ' и переходим к следующей n-грамме Word1
            If Mid$(Word1, i, n) = Mid$(Word2, Col(j), n) Then x = x + 1: Col.Remove j: Exit For
    Next j, i
    Result = 2 * x / (iMax + jMax)  ' совпадения на среднее количество n-грамм в словах
HandleExit:  SimilarityDice = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function SimilarityJaro(ByVal Word1 As String, ByVal Word2 As String) As Double
' Сходство Джаро-Винклера - минимальное число односимвольных преобразований, которое необходимо для того, чтобы изменить одно слово в другое
'-------------------------
' By: Ernanie F. Gregorio Jr. (from psc cd)
' Источник: https://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=73978&lngWId=1
' https://blog.developpez.com/philben/p12207/vba-access/vba-distance-de-jaro-winkler
' https://ru.wikipedia.org/wiki/Сходство_Джаро_—_Винклера
'-------------------------
Dim Result As Double ': Result = False
    On Error GoTo HandleError
Dim iProximity As Integer ' set the number of adjacent characters to compare to
Dim i As Integer, x As Integer
Dim iFr As Integer, iTo As Integer
Dim iMatch As Integer, iTrans As Integer
Dim iJaro As Double
    Word1 = UCase$(Trim$(Word1)): Word2 = UCase$(Trim$(Word2))
    If Word1 = Word2 Then Result = 1: GoTo HandleExit ' check if the two words are the same
    If InStr(1, Word1, Word2) > 0 Then Result = 1: GoTo HandleExit
    ' compute for the proximity of character checking
    ' allows matching characters to be up to X number of characters away.
    If Len(Word1) >= Len(Word2) Then
        iProximity = (Len(Word1) / 2) - 1
    Else
        iProximity = (Len(Word2) / 2) - 1
    End If
    For i = 1 To Len(Word2)
    ' this is the index of the character to be compared to
        iTo = (i + iProximity) - 1
        ' get the left most side character based on the iProximity
        If i <= iProximity Then iFr = 1 Else iFr = i - iProximity + 1
        ' start the letter by letter comparison
        For x = iFr To iTo
            If Mid$(Word2, i, 1) = Mid$(Word1, x, 1) Then
                If i = x Then iMatch = iMatch + 1: GoTo HandleNext
                iMatch = iMatch + 1: iTrans = iTrans + 1: Exit For
            End If
        Next x
HandleNext:
    Next i
    
    iTrans = iTrans \ 2
    If iMatch <= 0 Then Result = 0: GoTo HandleExit
    x = 0
    For i = 1 To 4
        If Mid$(Word2, i, 1) = Mid$(Word1, i, 1) Then x = x + 1 Else Exit For
    Next i
    iJaro = ((iMatch / Len(Word2)) + (iMatch / Len(Word1)) + ((iMatch - iTrans) / iMatch)) / 3
    If x > 0 Then Result = iJaro + 0.1 * x * (1 - iJaro) Else Result = iJaro
HandleExit:  SimilarityJaro = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

'=========================
' Функции обработки текста
'=========================
Public Function GenPassword( _
    Optional PassLen As Integer = 12, _
    Optional Symbols As String = vbNullString, _
    Optional bRepeats As Boolean = True, _
    Optional NewSeed As Boolean = True _
    ) As String
' генерирует "случайный" текст указанной длины
'-------------------------
' PassLen - длина пароля
' Symbols - допустимый набор символов. если не задан будет сформирован
' bRepeats = True - разрешить одинаковые символы подряд
' NewSeed = True - вызывает Randomize для создания новой "случайной" последовательности
'-------------------------
' при генерации паролей в цикле, Randomize приводит к повтору последовательности после ~100-200 уникальных паролей
' поэтому при генерации единичного пароля лучше задавать True, при генерации серии паролей в цикле - False
'-------------------------
Const cMin = 0              ' минимальная допустимая длина итоговой строки (0-неограничено)
Dim sMax As Byte: sMax = 3  ' максимальное количество однотипных символов подряд (0-неограничено)
Const bDigits = True        ' добавлять цифры
Const bLatin = True         ' добавлять символы латинского алфавита
Const bCyrillic = False     ' добавлять символы кириллического алфавита
Const bOthers = False       ' добавлять символы из доп. набора
'Const bRepeats = True       ' разрешить одинаковые символы подряд
Const sCase = 0             ' регистр символов формируемой строки
                            ' 0-допустимы символы в верхнем и нижнем регистрах
                            ' 1-только в верхнем, 2-только в нижнем
Dim Result As String ': Result = vbNullString
    On Error GoTo HandleError
    If PassLen < cMin Then Err.Raise vbObjectError + 512
    If PassLen < 1 Then GoTo HandleExit
    If Len(Symbols) = 0 Then
' если не задана - формируем последовательность допустимых символов
'    ' идея для генерации "читаемых" паролей - вместо набора алфавитных символов
'    ' генерировать на основе массива слогов и чередовать с другими символами
'    ' естественно нужен будет словарь. разбить на слоги можно при помощи HyphenateWord
' в порядке бреда:
'    ' можно попробовать "улучшить" генератор используя вместо чтения слева (N*Rnd)
'    ' чтение произвольной части сгенерированного числа
'    ' что-то типа: Replace(cCur(Rnd*1E15),",","") - превратит Rnd в строку из 19 цифр
'    ' из которой можно брать произвольный фрагмент
        If bDigits Then Symbols = Symbols & c_strSymbDigits
        If bLatin Then If sCase = 0 Then Symbols = Symbols & UCase$(c_strSymbEngAll) & LCase$(c_strSymbEngAll) Else If sCase = 1 Then Symbols = Symbols & UCase$(c_strSymbEngAll) Else Symbols = Symbols & LCase$(c_strSymbEngAll)
        If bCyrillic Then If sCase = 0 Then Symbols = Symbols & UCase$(c_strSymbRusAll) & LCase$(c_strSymbRusAll) Else If sCase = 1 Then Symbols = Symbols & UCase$(c_strSymbRusAll) Else Symbols = Symbols & LCase$(c_strSymbRusAll)
        If bOthers Then Symbols = Symbols & c_strSymbDigits & c_strSymbMath & c_strSymbPunct & c_strSymbCommas & c_strSymbParenth & c_strSymbOthers
    End If
Dim sLen As Long: sLen = Len(Symbols)
    If sLen < 1 Then GoTo HandleExit
' проверить количество типов в последовательности чтоб не подвесить цикл на условии по количеству однотипных символов
    ' sType д.б. один символ иначе надо будет делать проверку через массив/коллекцию типов
Dim sTemp As String: sTemp = vbNullString
Dim sType As Integer
Dim i As Long
    For i = 1 To sLen
        sType = GetCharType(Mid$(Symbols, i, 1))
        If InStr(1, sTemp, sType) = 0 Then sTemp = sTemp & sType
    Next i
' проверяем ограничения
    If sLen = 1 Then bRepeats = True ' если в наборе допустимых всего один символ снимаем запрет повторов
    If Len(sTemp) <= 1 Then sMax = 0 ' если все символы набора одного типа снимаем условие на повторы однотипных символов
' собственно генератор
Dim sChar As String * 1
Dim sPrev As Integer, sCount As Integer
    ' создание новой "случайной" последовательности
    sType = 0: sPrev = -1
    If NewSeed Then Randomize Timer
    Do Until Len(Result) = PassLen
    ' выбираем символ
HandleNewSymb: sChar = VBA.Mid$(Symbols, CLng((sLen - 1) * Rnd) + 1, 1)
    ' проверяем соответствие дополнительным требованиям к формируемой строке:
        '1. не более sMax однотипных символов подряд
        If sMax > 0 Then sType = GetCharType(sChar): If sType <> sPrev Then sPrev = sType: sCount = 1 Else If sCount >= sMax Then GoTo HandleNewSymb Else sCount = sCount + 1
        '2. запретить в результирующей строке находящиеся подряд одинаковые символы
        If Not bRepeats Then If LCase$(sChar) = LCase$(Right$(Result, 1)) Then GoTo HandleNewSymb
    ' формируем результирующую строку
        Result = Result & sChar
    Loop
HandleExit:  GenPassword = Result: Exit Function
HandleError: Result = vbNullString
    Select Case Err.Number
    Case vbObjectError + 512: MsgBox "Слишком короткий пароль." & vbCrLf & "Должен быть не меньше " & cMin & " символов.", vbOKOnly Or vbExclamation, "Ошибка!"
    End Select
    Err.Clear: Resume HandleExit
End Function
Public Function HyphenateWord( _
    ByVal Text As String, _
    Optional Delimiter As String = "­") As String
' расставляет переносы в словах
'-------------------------
' Источник: http://www.cyberforum.ru/vba/thread792944.html
' Описание и обсуждение исходного алгоритма здесь https://habr.com/post/138088/
' правила из набора паттернов сильно приблизительно соответствуют правилам русского языка
Const cstrTemp = "xgg xgs xsg xss sggsg gsssssg gssssg gsssg sgsg gssg sggg sggs"
Dim sPattern() As String
Dim i As Long, j As Long, k As Long
Dim m As String, sText As String
    
    On Error GoTo HandleError
' массив допустимых символов по типам: 0-знаки(x), 1-гласные(g), 2-согласные(s)
    ' единственная поправка - для алгоритма необходимо, чтобы "й" была знаком, а не согласной
    ' внесем соответствующую поправку в наборы символов массива
Dim sArr: sArr = Array(c_strSymbRusSign & c_strSymbEngSign, c_strSymbRusVowel & c_strSymbEngVowel, c_strSymbRusConson & c_strSymbEngConson)
' массив распознаваемых паттернов в слове
Dim sTemp() As String: sTemp = Split(cstrTemp) 'Call xSplit(cstrTemp, sTemp)
' позиция разбиения соотв паттерна - номер символа паттерна (см. массив выше) после которого необходимо поставить разделитель
Dim sPos: sPos = Array(1, 1, 1, 1, 3, 3, 2, 2, 2, 2, 2, 2)

    sText = Text
' заменяем символы исходной строки их обозначениями в паттерне (x, g, s)
    For i = 1 To Len(Text)
        m = LCase$(Mid$(Text, i, 1))
        For j = 0 To UBound(sArr)
            If InStr(sArr(j), m) Then Mid$(Text, i, 1) = Mid$("xgs", j + 1, 1): Exit For
    Next j, i
    
' выявляем паттерны и вставляем разделитель в позицию разбиения
' в преобразованную и исходную строки. Замена в преобразованной строке
' нужна чтобы исключить уже отработанные шаблоны разбиения
    For i = 0 To UBound(sTemp)
        j = 0
        Do
            k = InStr(j + 1, Text, sTemp(i))
            If k Then
                j = k + sPos(i)
                Text = VBA.Left$(Text, j - 1) & Delimiter & Mid$(Text, j)
                sText = VBA.Left$(sText, j - 1) & Delimiter & Mid$(sText, j)
            End If
        Loop While k
    Next i
HandleExit:  HyphenateWord = sText: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function NumToWords( _
    ByVal Number As String, _
    Optional ByVal NewCase As DeclineCase = DeclineCaseImen, _
    Optional ByRef NewNumb As DeclineNumb = DeclineNumbSingle, _
    Optional ByRef NewGend As DeclineGend = DeclineGendUndef, _
    Optional ByVal Animate As Boolean = False, _
    Optional ByRef NewType As NumeralType = NumeralOrdinal, _
    Optional ByVal Unit As String, Optional ByVal SubUnit As String, _
    Optional ByVal DecimalPlaces As Byte = 2, _
    Optional ByVal TranslateFrac As Boolean = False _
    ) As String
' преобразование чисел в слова и склонение по падежам
'-------------------------
' Number  - преобразуемое число (целое число, десятичная или натуральная дробь, число в экспоненциальной форме не распознается)
' NewCase - падеж склонения (им","род","дат","вин","тв","пред)
' NewNumb - число ("ед","мн")
' NewGend - род ("м","ж") если заданы единицы измерения определяется по ним
' Animate - признак того, что обозначение единиц измерения надо склонять как одушевлённые
' NewType - тип числительного (количественное","порядковое) если задано дробное число - м.б. только количественное
' Unit    - единица измерения - обозначение целой части (ед.ч., им.п.)
' SubUnit - вспомогательная единица измерения - обозначение дробной части (ед.ч., им.п.)
' DecimalPlaces - размерность вспомогательной единицы (количество знаков после запятой в десятичной дроби)
' TranslateFrac - (пока не используется) если True также переводится в текст дробная часть
'-------------------------
' ToDo: !!! подчистить правила склонения !!! - код слишком запутанный надо пересмотреть
'-------------------------
' как это должно быть:
    ' 23,50     - двадцать три рубля пятьдесят копеек
    ' 23 1/2    - двадцать три целых одна вторая рубля
    '(вариант)  - двадцать три и одна вторая рубля
' SubUnit и DecimalPlaces в десятичной дроби должны соответствовать друг другу.
' т.е. если Unit = "рубль"     и SubUnit = "копейка",  DecimalPlaces д.б. = 2 (1/100 руб.),
'    а если Unit = "килограмм" и SubUnit = "грамм",    DecimalPlaces д.б. = 3 (1/1000 кг)
' !!! не выйдет фокус вроде: NumToWords(Day(Now), Unit:=LCase(Format(Now, "mmmm")), NewType:=NumeralCardinal, NewCase:=DeclineCaseImen, NewGend:=DeclineGendNeut)
'     на выходе мы ожидаем что-то вроде "первое января", а получим: "первый январь", потому что программа правильно посчитает, что мы пересчитываем количество январей,
'     а совсем не то, что вы наверное имели ввиду: "первое (число) января", чтобы получилось надо отдельно просклонять число, отдельно месяц (в род.пад.)
'     правильно надо: NumToWords(Day(Now), NewType:=NumeralCardinal, NewCase:=DeclineCaseImen, NewGend:=DeclineGendNeut) & " " & DeclineWord(LCase(Format(Now, "mmmm")),DeclineCaseRod)
' !!! контроль соответствия размерности обозначения дробной части знаменателю дроби не производится
    ' например, - если DecimalPlaces не определено (=0)
    ' все варианты типа: "0,1";"0,01" и "0,001" "рубль"/"копейка"
    ' будут выведены как "ноль рублей одна копейка"
' также получится ерунда если задать SubUnit и оставить пустым Unit:
    ' "четыре целых десять копеек" ??? - даже не знаю как это правильно можно обработать...
' или попробовать вывести дробь как порядковое
    ' "первый и двенадцать сотых рубля" ??? - это вообще как должно правильно звучать?
'-------------------------
Const сShowNullInWhole = True   ' выводить нулевую целую часть
'Const сShowOnesInDemom = True   ' выводить единичные разряды в знаменателе (одна однотысячная","тысячная)
Const сEmptyWholeUnit = "целая" ' обозначение целой части если не задана единица измерения
Const сEmptyWholeUnit2 = "и"    ' текстовой разделитель целой и дробной части дроби если не заданы единицы измерения

Const cWhlDelim = " "   ' Chr(32)  - разделитель целой/дробной части натуральной дроби
Const cNatDelim = "/"   ' Chr(47)  - разделитель числителя/знаменателя натуральной дроби
Dim cDecDelim As String * 1: cDecDelim = p_GetLocaleInfo(LOCALE_SDECIMAL)   ' Chr(44)  - разделитель целой/дробной части десятичной дроби
Dim cPosDelim As String * 1: cPosDelim = p_GetLocaleInfo(LOCALE_STHOUSAND)  ' Chr(160) - разделитель разрядов целой части

Dim strWhole As String, strNomin As String, strDenom As String
Dim bolWhole As Boolean, bolNomin As Boolean, bolDenom As Boolean
Dim bytStep As Byte     ' текущий шаг обработки: 1-целая часть,2-числитель,3-знаменатель,0-не определено
Dim Result As String
    Result = vbNullString
    bolWhole = False: bolNomin = False: bolDenom = False
    On Error GoTo HandleError
    Number = Trim$(Number): Unit = Trim$(Unit): SubUnit = Trim$(SubUnit)
' делим на целую и дробную часть
    ' очищаем компоненты числа от разделителей разрядов и пробелов
    ' Replace вместо CLng потому что исходная задача -
    ' обрабатывать числа в том числе выходящие за ограничения типа Long
Dim tmpSymPos As Long: tmpSymPos = Nz(InStrRev(Number, cDecDelim), 0)
    If (tmpSymPos > 0) Then
' десятичная дробь
    ' знаменатель десятичной дроби выводим если не задано обозначение дробной части
        bolDenom = Len(SubUnit) = 0
        strWhole = VBA.Left$(Number, tmpSymPos - 1)
        strNomin = VBA.Mid$(Number, tmpSymPos + Len(cDecDelim), Len(Number) - tmpSymPos)
    ' формируем знаменатель десятичной дроби
    ' количество десятичных знаков в знаменателе м.б. от 1 до макс известного программе 10^33. опреджеляем по индексу массива от i=37 (тысяча) и далее
        If bolDenom Then
        ' если размерность единицы не указана - считаем по количеству разрядов в числителе
            DecimalPlaces = Len(strNomin)
        Else
        ' если размерность единицы указана
            Select Case Len(strNomin)
            Case Is < DecimalPlaces
            ' если размерность вспомогательной единицы больше числа знаков после запятой -
                ' дополняем числитель нулями до соответствия размерностей
                strNomin = strNomin & String(DecimalPlaces - Len(strNomin), "0")
            Case Is > DecimalPlaces:
            ' если размерность вспомогательной единицы меньше числа знаков после запятой -
                ' будем отдельно брать целую часть и числитель
                Result = NumToWords(strWhole, NewType:=NewType, NewCase:=NewCase, Unit:=Unit, Animate:=Animate, DecimalPlaces:=0)     ' целая часть
                ' числитель превращаем в десятичную дробь по количеству допустимых разрядов
                strNomin = Left$(strNomin, DecimalPlaces) & cDecDelim & Mid$(strNomin, DecimalPlaces + 1)
                Result = Result & " " & NumToWords(strNomin, NewType:=NewType, NewCase:=NewCase, Unit:=SubUnit, Animate:=Animate, DecimalPlaces:=0)  ' дробная часть
                GoTo HandleExit
            End Select
        End If
        strDenom = 1 & String(DecimalPlaces, "0")
    Else
        tmpSymPos = Nz(InStrRev(Number, cNatDelim), 0)
        If tmpSymPos > 0 Then
' натуральная дробь
            bolDenom = True ' ставим признак что это натуральная дробь
            SubUnit = vbNullString ' для натуральной дроби вспомогательная единица измерения не имеет смысла
            ' знаменатель
            strDenom = VBA.Mid$(Number, tmpSymPos + Len(cNatDelim))
            Number = VBA.Left$(Number, tmpSymPos - 1): tmpSymPos = Nz(InStrRev(Number, cWhlDelim), 0)
            ' числитель
            strNomin = VBA.Mid$(Number, tmpSymPos + Len(cWhlDelim), Len(Number) - tmpSymPos)
            ' целая часть
            If tmpSymPos > 0 Then strWhole = VBA.Left$(Number, tmpSymPos - 1)
        Else
' целое число или не число
            strWhole = Number: strNomin = 0: strDenom = 1
        End If
    End If
' определяем необходимость вывода частей числа
    ' числитель выводим если он непустой
    ' знаменатель выводим если он непустой
        ' и ранее решено его выводить (натуральная дробь или десятичная без вспомогательной единицы)
    ' целую часть выводим если он непустой
        ' или если задан вывод нулевой целой части дроби
        ' или если не задан вывод числителя (числитель пустой)
    On Error Resume Next
    bolNomin = p_NumType(strNomin) > 0
    bolDenom = p_NumType(strDenom) > 0 And bolDenom
    bolWhole = p_NumType(strWhole) > 0 Or сShowNullInWhole
    On Error GoTo HandleError
    
' переменные для хранения строк частей числа
Dim strNumb As String   ' разбираемая часть числа (целая/числитель/знаменатель)
Dim strWord As String   ' слово текущей части разбираемого числа
Dim strDelim As String  ' разделитель слов (обычно = Chr(32)) в результате
' переменные для хранения параметров текущего триплета числа
Dim bytNumb As Byte     ' тип разбираемого числа (для склонения, см. p_NumType)
Dim intTrip As Integer  ' содержимое текущего триплета разбираемой части числа
Dim bytTrip As Byte     ' порядковый номер триплета числа (с конца, начиная с 0)
Dim bolNull As Boolean  ' признак первого выводимого триплета. (нужно для правильного склонения)
    ' True  - означает отсутствие вывода или вывод первого триплета ещё не завершён
    '   т.е. при True при необходимости выводим единицу измерения,
    '   также в случае порядкового при True для триплета >0 изменяем склонение всех слов триплета, для 1 только первого
    ' False - первый доступный триплет уже выведен, идёт вывод старших триплетов
' переменные для хранения параметров склонения strWord
Dim tmpType  As NumeralType, tmpCase As DeclineCase, tmpNumb As DeclineNumb, tmpGend As DeclineGend  ' уточнённые (непосредственно используются при склонении)
'Dim NewGend As DeclineGend: NewGend = DeclineGendUndef ' род обозначения
' собираем строку справа
    ' обработка в три прохода: для целой части, для числителя и для знаменателя
    bytStep = 2                     ' начинаем с знаменателя
    Do
        bytTrip = 0                 ' порядковый номер триплета (справа)
        bolNull = True              ' признак пустого вывода
        strWord = vbNullString      ' обозначение ед.изм/разряда или слово текущего числа
    ' если вся часть равна 0 и это число с дробной частью возможно следует опустить целую часть
    ' ед.изм. Unit должна ставиться после целой части если есть дробная часть, это десятичная дробь и задан SubUnit
    ' иначе ставится после знаменателя и склоняется относительно единицы (1)
    ' т.е. "один рубль пятьдесят копеек", но "одна целая(и) пятьдесят сотых рубля" и "одна целая(и) одна вторая рубля"
    '      "десять рублей пятьдесят копеек", но "десять целых(и) пятьдесят сотых рубля" и "десять целых(и) одна вторая рубля"
    ' или  "одна верста пятьсот саженей", но "одна целая(и) пятьсот тысячных версты" или "одна целая(и) одна вторая версты"
    '      "десять вёрст пятьсот саженей", но "десять целых(и) пятьсот тысячных версты" или "десять целых(и) одна вторая версты"
        Select Case bytStep
        Case 0: If bolWhole Then strNumb = strWhole: strWord = IIf((Len(Unit) = 0) Or bolDenom Or (bolNomin And (Len(SubUnit) = 0)), IIf(bolNomin, IIf(NewType = NumeralCardinal, сEmptyWholeUnit2, сEmptyWholeUnit), vbNullString), Unit): GoTo HandleBegin
        Case 1: If bolNomin Then strNumb = strNomin: strWord = IIf(bolDenom, vbNullString, SubUnit): GoTo HandleBegin
        Case 2: If bolDenom Then strNumb = strDenom: strWord = Unit: GoTo HandleBegin
        Case Else: Exit Do
        End Select
        GoTo HandleNextPart
HandleBegin:
' начало обработки части числа
    ' подготовка строки разбираемой части числа
        If Len(strNumb) = 0 Then strNumb = 0        ' пустая строка = 0
        ' очищаем от разделителей разрядов и пробелов
        strNumb = Replace$(Replace(strNumb, cPosDelim, vbNullString), Space(1), vbNullString)
        ' убираем нули вначале (кроме числа целиком состоящего из нулей)
        tmpSymPos = 1: Do While VBA.Mid$(strNumb, tmpSymPos, 1) = "0": tmpSymPos = tmpSymPos + 1: Loop: If (tmpSymPos > 1) Then If (tmpSymPos > Len(strNumb)) Then strNumb = "0" Else strNumb = VBA.Mid$(strNumb, tmpSymPos)
    ' определяем тип числа (см.p_NumType) необходимо для правильного склонения
        Select Case bytStep
        Case 0: bytNumb = p_NumType(strNumb):   tmpType = NewType
        Case 1: bytNumb = p_NumType(strNumb):   If bolDenom Then tmpType = NumeralOrdinal Else tmpType = NewType
        Case 2: bytNumb = p_NumType(strNomin):  tmpType = NumeralCardinal  ': tmpCase = DeclineCaseImen
        End Select
        
' начало обработки числа
        intTrip = Abs(CInt(VBA.Right$(strNumb, 3))) ' берём младший (0) триплет числа
HandleUnits:
' обозначение единицы измерения, получаем его род и добавляем его к результирующей строке части числа
    ' если вывода ещё не было (это первый не пустой триплет) в strWord сейчас единица измерения части или пусто
        tmpGend = DeclineGendUndef
        If Len(strWord) > 0 Then
    ' склоняем обозначение единицы измерения и определяем его род
        ' число слова.  ед.ч для чисел на 1 (кроме 11) и чисел на 2-4 (кроме 12-14) в им.,род. и вин.п., остальные - во мн.ч.
        ' падеж слова.  NewCase, кроме им. и вин.п. для чисел не заканчивающихся на 1 (кроме 11) они в род.п.
            tmpCase = NewCase: tmpNumb = DeclineNumbPlural
            If (bytStep = 2) Or (tmpType = NumeralCardinal) Then
                tmpNumb = DeclineNumbSingle: If (bytStep = 2) Then tmpCase = DeclineCaseRod
            Else
                If (intTrip = 0) Or (bytTrip <> 0) Then tmpCase = DeclineCaseRod
                Select Case bytNumb
                Case 1:     If (bytTrip = 0) Then tmpNumb = DeclineNumbSingle
                Case 2:     If (tmpCase = DeclineCaseImen) Or (tmpCase = DeclineCaseVin) Then tmpCase = DeclineCaseRod: If (bytTrip = 0) Then tmpNumb = DeclineNumbSingle
                Case Else:  If (tmpCase = DeclineCaseImen) Or (tmpCase = DeclineCaseVin) Then tmpCase = DeclineCaseRod
                End Select
            End If
            ' уточняем для обозначений - прилагательных
            If bytNumb = 2 Then If p_GetWordSpeechPartType(strWord) = SpeechPartTypeAdject Then tmpNumb = DeclineNumbPlural
        ' склоняем и запись в результат
            strWord = DeclineWord(strWord, tmpCase, tmpNumb, tmpGend, Animate): If tmpGend <> DeclineGendUndef Then NewGend = tmpGend
            If Len(Result) > 0 Then Result = strWord & strDelim & Result Else Result = strWord
        End If
        ' род числительного - по роду ед.измерения/разряда/если не определен - муж.род (двадцать один), жен.род (одна целая две сотых)
        If tmpGend = DeclineGendUndef Then If NewGend = DeclineGendUndef Then tmpGend = DeclineGendMale Else tmpGend = NewGend
        'If tmpGend = DeclineGendUndef Then If NewGend = DeclineGendUndef Then tmpGend = DeclineGendFem Else tmpGend = NewGend
        Do
' делим часть числа на триплеты разрядов (перебираем цифры справа по 3).
' первый уже получен и находится в intTrip, в bytTrip - порядковый номер текущего триплета
HandleThousands:
' обозначение разряда, получаем его род и добавляем его к результирующей строке части числа
            strDelim = Space(1)         ' разделитель элементов числа - пробел (кроме составных порядковых)
    ' пропускаем пустые триплеты (кроме младшего. 0 в младшем - возможно 0 целых)
            If intTrip = 0 Then If (Not bolWhole) Or (bytStep <> 0) Or (bytNumb <> 0) Then GoTo HandleNextTriplet
    ' у нулевого триплета нет своего обозначения разряда, его - пропускаем
            If bytTrip = 0 Then GoTo HandleDigits
    ' для старших триплетов получаем обозначение разряда, склоняем его и определяем его род
        ' обозначение разряда - порядковое, число соответствующее разряду
        ' род обозначения разряда используется при склонении старших триплетов числительного
    ' склоняем обозначение разряда и определяем его род
        ' число слова.  ед.ч для чисел на 1 (кроме 11) и чисел на 2-4 (кроме 12-14) в им.,род. и вин.п., остальные - во мн.ч.
        ' падеж слова.  NewCase, кроме им.,род. и вин.п. для чисел не заканчивающихся на 1 (кроме 11) они в род.п.
            strWord = vbNullString
            tmpCase = NewCase: tmpNumb = DeclineNumbPlural
            If ((bytStep = 2) Or (NewType = NumeralCardinal)) And Not bolNull Then tmpCase = DeclineCaseImen
            If (tmpType = NumeralCardinal) And (bytStep = 0) Then
                tmpNumb = DeclineNumbSingle
            Else
                Select Case bytNumb
                Case 1:     tmpNumb = DeclineNumbSingle
                Case 2:     If (bytStep <> 2) Then If (tmpCase = DeclineCaseImen) Or (tmpCase = DeclineCaseVin) Then tmpCase = DeclineCaseRod: tmpNumb = DeclineNumbSingle
                Case Else:  If (tmpCase = DeclineCaseImen) Or (tmpCase = DeclineCaseVin) Then tmpCase = DeclineCaseRod
                End Select
            End If
        ' склоняем и запись в результат
            strWord = p_NumDecline(intTrip, bytTrip, tmpCase, tmpNumb, tmpGend, tmpType, Animate)
            If Len(strWord) = 0 Then GoTo HandleDigits
            If Len(Result) > 0 Then Result = strWord & strDelim & Result Else Result = strWord
        ' если обозначение разряда делали порядковым, т.е. если младший разряд (единицы) пустой
            ' делаем значения разряда количественными
            ' делаем разделитель пустым т.к. порядковые на -тысячный/-милионный и т.п. пишутся слитно
            ' для знаменателя - в дальнейшем склоняем значения разряда в зависимости от типа числа в знаменателе
            If tmpType = NumeralCardinal Then tmpType = NumeralOrdinal: strDelim = vbNullString: If bytStep = 2 Then bytNumb = p_NumType(intTrip)
HandleDigits:
'    ' пропускаем пустые триплеты (кроме младшего. 0 в младшем - возможно 0 целых)
'            If intTrip = 0 Then If (Not bolWhole) Or (bytStep <> 0) Or (bytNumb <> 0) Then GoTo HandleNextTriplet
            Do
' перебираем последовательно элементы триплета:
    ' сотни, десятки (кроме второго), второй десяток (10-19), единицы и ноль
    ' проверяем состав триплета и уменьшаем остаток
    ' склоняем слово разбираемого триплета с учётом рода обозначения разряда/единицы измерения
                tmpCase = NewCase: If tmpCase = DeclineCaseUndef Then tmpCase = DeclineCaseImen
                Select Case bytNumb
                Case 0, 1:  tmpNumb = DeclineNumbSingle
                Case Else:  tmpNumb = DeclineNumbPlural
                End Select
                If tmpType = NumeralCardinal Then
                ' первое число в первом триплете склоняем как порядковое в указанном (NewCase) падеже,
                ' кроме им. и вин., их - в род.п. для чисел не на 1 и 2-4
                    Select Case bytStep
                    Case 0: tmpNumb = NewNumb
                    'Case 1: If Not bolDenom Then tmpNumb = DeclineNumbSingle
                    Case 2: If (bytNumb <> 1 And bytNumb <> 2) And ((tmpCase = DeclineCaseImen) Or (tmpCase = DeclineCaseVin)) Then tmpCase = DeclineCaseRod
                    End Select
                ElseIf ((NewType = NumeralCardinal) And (bytStep = 0)) Or (bytStep = 2) Then
                ' старшие разряды в первом триплете
                ' и старшие триплеты (кроме первого выводимого) склоняем в им.п.
                    tmpCase = DeclineCaseImen
                    If bolNull And bytTrip > 0 Then
                ' старшие триплеты (в первом выводимом) склоняем:
                '   все разряды склоняем в р.п., искл.: одно- и сто-
                '   -тысячный -миллионный и т.п., берём пустой разделитель
                        If (intTrip Mod 10 = 1) And ((intTrip Mod 100) \ 10 <> 1) Then
                            tmpGend = DeclineGendNeut ': tmpNumb = DeclineNumbSingle
                        ElseIf intTrip <> 100 Then
                            tmpCase = DeclineCaseRod
                        End If
                    End If
                End If
        ' склоняем и запись в результат
                strWord = p_NumDecline(intTrip, , tmpCase, tmpNumb, tmpGend, tmpType, Animate, NumbRest:=intTrip)
                If Len(strWord) > 0 Then If Len(Result) > 0 Then Result = strWord & strDelim & Result Else Result = strWord
            ' переопределяем тип числительного (только первый элемент составного числительного м.б. порядковым)
                tmpNumb = DeclineNumbUndef ': tmpGend = DeclineGendUndef
                If tmpType = NumeralCardinal Then tmpType = NumeralOrdinal
            Loop While intTrip > 0
            bolNull = False
HandleNextTriplet:
        ' обрезаем разобранный триплет
            If Len(strNumb) > 2 Then strNumb = VBA.Left$(strNumb, Len(strNumb) - 3) Else strNumb = vbNullString
        ' повторяем пока не достигнем старшего триплета
            If Len(strNumb) = 0 Then Exit Do
        ' берем очередной триплет части числа и (если был вывод) его тип
            intTrip = Abs(CInt(VBA.Right$(strNumb, 3))): If Not (bolNull And (tmpType = NumeralCardinal)) Then bytNumb = p_NumType(intTrip)
        ' увеличиваем счетчик разобранных триплетов
            bytTrip = bytTrip + 1
        Loop 'While Len(strNumb)>0
HandleNextPart:
        ' переходим к обработке следующей части дроби
        If bytStep = 0 Then Exit Do Else bytStep = bytStep - 1
    Loop While bytStep >= 0
'    ' добавляем минус
'    If VBA.Left$(Number, 1) = "-" And Len(Result)>0 Then Result = "минус " & Result
HandleExit:  NumToWords = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function

Private Function p_NumType(ByVal Numb As Variant, Optional InNomin As Boolean = False) As Byte
' для NumToWords. получает тип числа 0-999
'-------------------------
' InNomin - флаг устанавливаемый при проверке числителя для правильного склонения знаменателя
' в этом случае числа на 2 должны склоняться в ед.ч. во всех остальных - во мн.ч.
' возвращает:
'   0 - если Numb = 0
'   1 - если Numb = xx1 и <>x11         (при InNomin=True также xx2 и <>x12)
'   2 - если Numb = xx2-4 и <>x12-x14   (при InNomin=True кроме xx2 и <>x12)
'  255 - все остальное
'-------------------------
' для правильного склонения важна последняя цифра триплета
'-------------------------
Dim Result As Byte
    Result = 255
    On Error GoTo HandleError
    If Numb = 0 Then Result = 0:   GoTo HandleExit      ' 0
    If Numb < 0 Then Err.Raise vbObjectError + 512
    If Numb > 999 Then Numb = VBA.Right$(Numb, 3)
    If (Numb \ 10) Mod 10 = 1 Then GoTo HandleExit      ' 10-19
    Select Case Numb Mod 10
    Case 1:         Result = 1
    Case 2:         Result = IIf(InNomin, 1, 2)
    Case 3 To 4:    Result = 2
    End Select
HandleExit:  p_NumType = Result: Exit Function
HandleError: Result = 255: Err.Clear: Resume HandleExit
End Function


Public Function DeclineWord( _
    ByVal Word As String, _
    Optional NewCase As DeclineCase = DeclineCaseRod, _
    Optional NewNumb As DeclineNumb = DeclineNumbSingle, _
    Optional ByRef NewGend As DeclineGend = DeclineGendUndef, _
    Optional ByRef Animate As Boolean = False, _
    Optional IsFio As Byte = 0, _
    Optional ByRef SymbCase As Integer = 0, Optional ByRef Template As String _
    ) As String
' склонение слова. работает весьма условно
'-------------------------
' Крайне условно пытается различать род, число, часть речи, выделять окончание
' Никак не различает одушевленное/неодушевленное (из-за этого, в частности, неправильно склоняет в мн.ч. вин.п.)
' для лучших результатов пользуйтесь Morpher'ом с http://morpher.ru
' или Padej'ом http://www.delphikingdom.com/asp/viewitem.asp?catalogid=412
' Word - существительное в единственном числе именительном падеже
' NewCase - падеж ("р","д","в","т","п")
' NewNumb - число ("ед","мн")
' NewGend - род ("м","ж")
' Animate - признак (одушевленное","неодушевленное)
' IsFio - признак ФИО (0-не ФИО, 1-Фамилия, 2-Имя, 3-Отчество)
' SymbCase, Template - состояние регистра символов исходного слова и шаблон
'-------------------------
' v.1.0.1       : 06.07.2019 - склонение числительных вынесено в отдельную функцию
'-------------------------
Dim WordBeg As String, WordEnd As String
Dim WordType As SpeechPartType
Dim sChar As String '* 1
Dim i As Long, iMax As Long
Dim Result As String
' надо переписать на более вменяемо выглядящие правила
    On Error GoTo HandleError
    If NewCase = DeclineCaseUndef Then NewCase = DeclineCaseImen Else If NewCase > DeclineCasePred Or NewCase < DeclineCaseImen Then Err.Raise vbObjectError + 512
    If NewNumb = DeclineNumbUndef Then NewNumb = DeclineNumbSingle
    If IsFio Then Animate = True ' определять одушевленные не умеем, однако ФИО - однозначно одушевленные
' получаем шаблон регистра символов слова
    Word = Trim$(Word): If SymbCase = 0 Then SymbCase = p_GetSymbCase(Word, Template)
    
    Result = LCase$(Word)
' Обработка исключений
'' Шаблон исключения
'    Case ""
'        Select Case NewCase
'        Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = ""   ' им.п. мн.ч    (кто/что)       Nominative
'        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "", "")  ' р.п. ед/мн.ч  (кого/чего)     Genitive
'        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "", "")  ' д.п. ед/мн.ч  (кому/чему)     Dative
'        Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "", "")  ' в.п. ед/мн.ч  (кого/что)      Accusative
'        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "", "")  ' т.п. ед/мн.ч  (кем/чем)       Ablative
'        Case DeclineCasePred: WordEnd = Choose(NewNumb, "", "")  ' п.п. ед/мн.ч  (о ком/о чём)   Prepositional
'        'Case Else 'DeclineCaseUndef = 0
'        End Select
'        i = Len(Result)-2: GoTo HandleExit ' i = кол-во знаков от начала слова к которым добавляется окончание
' Пропускаем предлоги и пр. несклоняемые слова
    Select Case Result
    Case "в", "с", "к", "у", "а", "и", "о", "об", "на", "по", "во", "за", "от", "не", "ни", "ли", _
         "или", "еле", "над", "при", "под", "для", "через", "перед", "ввиду", "наподобие", "вроде", _
         "вблизи", "вглубь", "вдоль", "возле", "около", "среди", "вокруг", "внутри", "впереди", "после", _
         "насчет", "навстречу", "вслед", "вместо", "ввиду", "благодаря", "вследствие"
        i = Len(Result): GoTo HandleExit
    End Select
' прочие исключения
    If IsFio Then
    ' не склоняются фамилии на:
    Select Case Right(Result, 1)
    Case "о": i = Len(Result): GoTo HandleExit
    End Select
    End If
    Select Case Result
    ' замена гласной (е -> ь)
    Case "лев", "лёд", "лён": If NewCase <> DeclineCaseImen Or NewNumb = DeclineNumbPlural Then Mid$(Result, 2, 1) = "ь"
    ' выпадение гласной (2-я с конца)
    Case "павел", "угол", "конец", "лоб", "сон", "рот", "потолок": If Not (NewCase = DeclineCaseImen Or (Not Animate And NewCase = DeclineCaseVin)) Or NewNumb = DeclineNumbPlural Then i = Len(Result) - 2: Result = Left$(Result, i) & Mid(Result, i + 2)
    ' выпадение гласной (3-я с конца)
    Case "пень", "уголь", "день", "огонь": If Not (NewCase = DeclineCaseImen Or (Not Animate And NewCase = DeclineCaseVin)) Or NewNumb = DeclineNumbPlural Then i = Len(Result) - 3: Result = Left$(Result, i) & Mid(Result, i + 2)
    ' выпадение гласной (особые случаи)
    Case "ложь": NewNumb = DeclineNumbSingle: If Not (NewCase = DeclineCaseImen Or NewCase = DeclineCaseVin Or NewCase = DeclineCaseTvor) Then i = Len(Result) - 3: Result = Left$(Result, i) & Mid(Result, i + 2)
    ' местоимения
    Case "я": Animate = True
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewNumb, "я", "мы")
        Case DeclineCaseRod, DeclineCaseVin: WordEnd = Choose(NewNumb, "меня", "нас")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "мне", "нам")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "мной", "нами")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "мне", "нас")
        End Select
        i = Len(Result) - 1: GoTo HandleExit
    Case "ты": Animate = True
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewNumb, "ты", "вы")
        Case DeclineCaseRod, DeclineCaseVin: WordEnd = Choose(NewNumb, "тебя", "вас")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "тебе", "вам")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "тобой", "вами")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "тебе", "вас")
        End Select
        i = Len(Result) - 2: GoTo HandleExit
    Case "вы": Animate = True
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = "ы"
        Case DeclineCaseRod, DeclineCaseVin: WordEnd = "ас"
        Case DeclineCaseDat:  WordEnd = "ам"
        Case DeclineCaseTvor: WordEnd = "ами"
        Case DeclineCasePred: WordEnd = "ас"
        End Select
        i = Len(Result) - 1: GoTo HandleExit
    Case "он": 'Animate = True
        If NewGend = DeclineGendUndef Then NewGend = DeclineGendMale
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewNumb, Choose(NewGend, "он", "она", "оно"), "они")
        Case DeclineCaseRod, DeclineCaseVin: WordEnd = Choose(NewNumb, Choose(NewGend, "него", "неё", "них"), "их")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, Choose(NewGend, "нему", "ней", "ним"), "ним")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, Choose(NewGend, "ним", "ней", "ними"), "ними")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, Choose(NewGend, "нём", "ней", "них"), "них")
        End Select
        '"н" обязательно после предлогов как проверять пока не решил
        i = Len(Result) - 2: GoTo HandleExit
    Case "то": Animate = False: NewGend = DeclineGendNeut
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewNumb, "о", "е")
        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "ого", "ех")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "ому", "ем")
        Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "от", "ех")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ем", "еми")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "ом", "ех")
        End Select
        i = Len(Result) - 1: GoTo HandleExit
    Case "это": Animate = False: NewGend = DeclineGendNeut
        NewGend = DeclineGendNeut
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewNumb, "о", "и")
        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "ого", "их")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "ому", "им")
        Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "от", "их")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "им", "ими")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "ом", "их")
        End Select
        i = Len(Result) - 1: GoTo HandleExit
    Case "номер": Animate = False: NewGend = DeclineGendMale
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: If NewNumb = DeclineNumbPlural Then WordEnd = "а"
        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "а", "ов")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "у", "ам")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ом", "ами")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ах")
        End Select
        i = Len(Result): GoTo HandleExit
    End Select
    
'    SymbCase = p_GetSymbCase(Word, Template) ' получаем шаблон регистра
    iMax = Len(Result): WordEnd = vbNullString: i = iMax ' позиция начала окончания
' Определение части речи (очень приблизительно)
    WordType = p_GetWordSpeechPartType(LCase$(Word))
    Select Case WordType
    Case SpeechPartTypeNoun, _
         SpeechPartTypeAdject   ' существительные и прилагательные склоняем по правилам ниже
    Case SpeechPartTypeNumeral  ' числительные склоняем отдельно
        Result = p_NumDecline(Word, , NewCase, NewNumb, NewGend, , Animate): i = Len(Result): GoTo HandleExit
    Case Else: GoTo HandleExit  ' все остальные (необрабатываемые) - пропускаем
'    Case SpeechPartTypePronoun ' местоимения
'    Case SpeechPartTypeVerb    ' глаголы
'    Case SpeechPartTypePretext ' предлоги
'    Case SpeechPartTypeUndef   ' неизвестные
    End Select
' Определяем начало, окончание слова и букву перед окончанием
    ' возможо замену ё>е надо делать после определения окончания
    ' для обработки -ок, -ёк и -он, -ён
    'Result = Replace(Result, "ё", "е")
    Call p_GetWordParts(Result, WordBeg, WordEnd, Template)
    i = iMax - Len(WordEnd)                     ' позиция начала окончания
' Дополнительное ограниичение: если длина слова < 3 и последняя гласная - лучше не склонять
    If i < 3 And InStr(1, c_strSymbRusVowel, WordEnd) Then GoTo HandleExit
    sChar = LCase$(Mid$(Result, i, 1))          ' буква перед окончанием
' Определение рода (очень приблизительно)
    If NewGend = DeclineGendUndef Then _
       NewGend = p_GetWordGender(Word, WordEnd) 'And NewNumb <> DeclineNumbPlural
    If NewGend = DeclineGendNeut Then Animate = False ' ср.р считаем неодушевленным
' Обработка окончаний и склонение
    Select Case WordEnd
    'Case "ии"
    '' мн.число
    'Case "и"
    '' мн.число
    Case "а"
        Select Case sChar
        Case "к", "г", "х"
    ' -ка, -га, -ха
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "и"
            Case DeclineCaseRod: If NewNumb <> 2 Then WordEnd = "и": GoTo HandleExit
                    Select Case Mid$(WordBeg, i - 1, 1)
                    Case "й":                   WordEnd = "е" & sChar: i = i - 2 '-йка, -йга
                    Case "ч", "ш", "ж", "ц":    WordEnd = "е" & sChar: i = i - 1 '-чка, -шка
                    Case Else:                  WordEnd = ""
                    End Select
            Case DeclineCaseDat: WordEnd = Choose(NewNumb, "е", "ам")
            Case DeclineCaseVin: If NewNumb <> DeclineNumbPlural Then WordEnd = "у": GoTo HandleExit
                    If Animate Then
                    Select Case Mid$(WordBeg, i - 1, 1)
                    Case "й":                   WordEnd = "е" & sChar: i = i - 2 '-йка, -йга
                    Case "ч", "ш", "ж", "ц":    WordEnd = "е" & sChar: i = i - 1 '-чка, -шка
                    Case Else:                  WordEnd = ""
                    End Select
                    Else: WordEnd = "и"
                    End If
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ой", "ами")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ах")
            End Select
        Case "ч", "щ"
    ' -ча, -ща
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "и"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "и", "")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "е", "ам")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "у", IIf(Animate, "ей", "и"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ей", "ами")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ах")
            End Select
        Case "ц"
    ' -ца
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "ы"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "ы", "")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "е", "ам")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "у", IIf(Animate, "", "ы"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ей", "ами")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ах")
            End Select
        Case "в", "н"
            If IsFio = 1 And NewGend = 2 Then
    ' женские фамилии на -ва, -на
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "ы"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "ой", "ых")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "ой", "ым")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "у", "ых")
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ой", "ыми")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "ой", "ых")
            End Select
    ' прочие на -ва, -на
            Else
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "ы"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "ы", "ых")
            Case DeclineCaseDat, DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ым")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "у", IIf(Animate, "", "ы"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ой", "ыми")
            End Select
            End If
        Case Else
    ' прочие на -а
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "ы"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "ы", "")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "е", "ам")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "у", IIf(Animate, "", "ы"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ой", "ами")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ах")
            End Select
        End Select
    Case "я"
        Select Case sChar
        Case "м"
    ' на -мя
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "ена"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "ени", "ян")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "ени", "енам")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "я", IIf(Animate, "ян", "ена"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "енем", "енами")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "ени", "енах")
            End Select
        Case Else
    ' прочие на -я
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "и"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "и", "")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "е", "ям")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "ю", IIf(Animate, "", "и"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ей", "ями")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ях")
            End Select
        End Select
    Case "о"
    ' на -о
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "а"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "а", "")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "у", "ам")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "о", IIf(Animate, "", "а"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ом", "ами")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ах")
            End Select
    Case "е"
    ' на -е
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "я"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "я", "ей")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "ю", "ям")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "е", IIf(Animate, "ей", "я"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ем", "ями")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ях")
            End Select
    Case "ое"
    ' на -ое
            If InStr(1, "кгх", sChar) Then sChar = "и" Else sChar = "ы"
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = sChar & "е"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "ого", sChar & "х")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "ому", sChar & "м")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "ое", sChar & IIf(Animate, "х", "е")) '
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, sChar & "м", sChar & "ми")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "ом", sChar & "х")
            End Select
    Case "ее"
    ' на -ее
            If sChar <> "ц" Then sChar = "и" Else sChar = "ы"
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = sChar & "е"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "его", sChar & "х")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "ему", sChar & "м")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "ее", IIf(Animate, "их", sChar & "е"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, sChar & "м", sChar & "ми")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "ем", sChar & "х")
            End Select
    Case "ие"
    ' на -ие
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "ия"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "ия", "ий")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "ию", "иям")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "ие", IIf(Animate, "их", "ия"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ием", "иями")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "ии", "иях")
            End Select
    Case "ия"
    ' на -ия
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "ии"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "ии", "ий")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "ии", "иям")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "ию", IIf(Animate, "ий", "ии"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ией", "иями")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "ии", "иях")
            End Select
    Case "ая"
    ' на -ая
        If InStr(1, "цчшщж", sChar) Then
            If NewNumb = DeclineNumbPlural And sChar = "ц" Then sChar = "ы" Else sChar = "и"
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "ие"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "ей", sChar & "х")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "ей", sChar & "м")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "ую", IIf(Animate, sChar & "х", "ие"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ей", sChar & "ми")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "ей", sChar & "х")
            End Select
        Else
            If NewNumb = DeclineNumbPlural Then
            If InStr(1, "кгх", sChar) Then sChar = "и" Else sChar = "ы"
            End If
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = sChar & "е"
            Case DeclineCaseRod, DeclineCasePred: WordEnd = Choose(NewNumb, "ой", sChar & "х")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "ой", sChar & "м")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "ую", sChar & IIf(Animate, "х", "е"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ой", sChar & "ми")
            End Select
        End If
    Case "ин", "ын"
        i = i + 2
        If IsFio = 1 Then
    ' фамилии на -ин, -ын
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, vbNullString, "ы")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "а", "ых")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "у", "ым")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "а", "ых")
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ым", "ыми")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ых")
            End Select
        Else
    ' прочие на -ин, -ын
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = WordEnd & "ы"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "а", "")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "у", "ам")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "а", "")
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ом", "ами")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ах")
            End Select
        End If
    Case "ой", "ый"
            If (NewNumb = DeclineNumbPlural) And (WordEnd = "ой") And (sChar = "ш" Or sChar = "х") Then sChar = "и" Else sChar = "ы"
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, Choose(NewGend, WordEnd, "ая", "ое"), sChar & "е")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ого", "ой"), sChar & "х")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ому", "ой"), sChar & "м")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, IIf(Animate, "ого", WordEnd), "ую"), sChar & "х")
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ым", "ой"), sChar & "ми")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ом", "ой"), sChar & "х")
            End Select
    Case "ий"
        Select Case sChar
        Case "к"    ' -кий
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, Choose(NewGend, "ий", "ая", "ое"), "ие")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ого", "ой"), "их")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ому", "ой"), "им")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ого", "ой"), "их")
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "им", "ой"), "ими")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ом", "ой"), "их")
            End Select
        Case "р"    ' -рий
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, Choose(NewGend, "ий", "ия", "ие"), "ии")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ия", "ию"), "иев")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ию", "ии"), "иям")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ия", "ию"), "иев")
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ием", "ией"), "иями")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ие", "ии"), "иях")
            End Select
        Case "т"    ' -тий
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, Choose(NewGend, "ий", "ья", "ье"), "ьи")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ьего", "ьей"), "ьих")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ьему", "ьей"), "ьим")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ьего", "ьей"), "ьих")
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ьим", "ьей"), "ьими")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ьем", "ьей"), "ьих")
            End Select
        Case "ж", "ш", "н"   ' -жий и -ший
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, Choose(NewGend, "ий", IIf(sChar = "н", "я", "а") & "я", "ее"), "ие")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "его", "ей"), "их")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ему", "ей"), "им")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, Choose(NewGend, IIf(Animate, "его", "ий"), IIf(sChar = "н", "ю", "у") & "ю", "ее"), IIf(Animate, "их", "ие"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "им", "ей"), "ими")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "ем", "ей"), "их")
            End Select
        End Select
    Case "ай", "ей", "уй", "эй", "юй", "яй" ', "ий"
        If sChar <> "к" Then
    ' прочие на -ай, -ей, -уй, -эй, -юй, -яй ', -ий
            i = i + 1
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, "й", "и")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "я", "и")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "ю", "ям")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, "я", "й"), IIf(Animate, "ев", "и"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ем", "и")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ях")
            End Select
        End If
    Case "ав"
    ' на -ав
            i = i + 2
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, vbNullString, "ы")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "а", "ов")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "у", "ам")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, "а", ""), IIf(Animate, "ов", "ы"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ом", "ами")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ах")
            End Select
    Case "ок", "ёк"
    ' на -ок
            If Left$(WordEnd, 1) = "ё" Then sChar = "ь" Else sChar = ""
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = IIf(Left$(WordEnd, 1) = "ё", "ь", "") & "ки"
            Case DeclineCaseRod:  WordEnd = sChar & Choose(NewNumb, "ка", "ков")
            Case DeclineCaseDat:  WordEnd = sChar & Choose(NewNumb, "ку", "кам")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, sChar & "ка", WordEnd), IIf(Animate, sChar & "ков", sChar & "ки"))
            Case DeclineCaseTvor: WordEnd = sChar & Choose(NewNumb, "ком", "ками")
            Case DeclineCasePred: WordEnd = sChar & Choose(NewNumb, "ке", "ках")
            End Select
            i = Len(Result) - 2
'    Case "еёк" '>йка и т.д.
'    Case "их", "ых"
'    Case "ин", "ын"
    Case Else
    ' все остальное
        WordBeg = LCase$(Result): WordEnd = vbNullString: i = iMax
        sChar = Right$(WordBeg, 1)
        Select Case LCase$(Right$(Template, 1))
        Case "s" ' заканчивается на согласную
        ' женские фамилии на согласную оставляем как есть
            If IsFio = 1 And NewGend = 2 Then GoTo HandleExit
            Select Case sChar
            Case "й"
        ' на -й
                Select Case NewCase
                Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "и":
                Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "я", "ёв")
                Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "ю", "ям")
                Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, "я", WordEnd), IIf(Animate, "ев", "и"))
                Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ем", "ями")
                Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ях")
                End Select
                i = i - 1
            Case "ч", "ш", "щ", "ж"
        ' на -ч,-ш,-щ,-ж
                Select Case NewCase
                Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "и"
                Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "а", "ей")
                Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "у", "ам")
                Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, "а", WordEnd), IIf(Animate, "ей", "и"))
                Case DeclineCaseTvor: WordEnd = Choose(IsFio + 1, Choose(NewNumb, "ом", "ами"), Choose(NewNumb, "ем", "ами"), Choose(NewNumb, "ом", "ами"), Choose(NewNumb, "ем", "ами"))
                Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ах")
                End Select
            Case "з"
      ' на -з
                Select Case NewCase
                Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "ы"
                Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "а", "ов")
                Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "у", "ам")
                Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, "а", ""), IIf(Animate, "ов", "ы"))
                Case DeclineCaseTvor: WordEnd = Choose(IsFio + 1, Choose(NewNumb, "ом", "ами"), Choose(NewNumb, "ым", "ыми"), Choose(NewNumb, "ом", "ами"), Choose(NewNumb, "ем", "ами"))
                Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ах")
                End Select
            Case Else
      ' на остальные согласные
                Select Case NewCase
                Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = IIf(InStr(1, "кгхжчшщ", sChar), "и", "ы")
                Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "а", "ов")
                Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "у", "ам")
                Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, "а", ""), IIf(Animate, "ов", "ы")) ' "ы"",""ов"
                Case DeclineCaseTvor: WordEnd = Choose(IsFio + 1, Choose(NewNumb, "ом", "ами"), Choose(NewNumb, "ым", "ыми"), Choose(NewNumb, "ом", "ами"), Choose(NewNumb, "ем", "ами"))
                Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ах")
                End Select
            End Select
        Case "x"
    ' заканчивается на -ь
        i = i - 1
        Select Case VBA.Left$(VBA.Right$(Result, 2), 1)
        Case "ч", "ш", "ж":
            If IsFio = 1 Then GoTo HandleExit
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, "ь", "и")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "и", "ей")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "и", "ам")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "ь", IIf(Animate, "ей", "и")) ' если дочь -> дочерей
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ью", "ами")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "и", "ах")
            End Select
        Case Else:
            If IsFio = 1 And NewGend <> 1 Then GoTo HandleExit
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, "ь", "и")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, IIf(NewGend = 1, "я", "и"), "ей")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, IIf(NewGend = 1, "ю", "и"), "ям")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, "я", "ь"), IIf(Animate, "ей", "и"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, IIf(NewGend = 1, "ём", "ью"), "ями")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, IIf(NewGend = 1, "е", "и"), "ях")
            End Select
        End Select
        End Select
    End Select
HandleExit:
    Result = p_SetSymbCase(Left$(Result, i) & WordEnd, SymbCase, Template)
    DeclineWord = Result
    Exit Function
HandleError: i = iMax: WordEnd = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function DeclineWords( _
    ByVal Words As String, _
    Optional NewCase As DeclineCase = DeclineCaseRod, _
    Optional NewNumb As DeclineNumb = DeclineNumbSingle, _
    Optional ByVal NewGend As DeclineGend = DeclineGendUndef, _
    Optional ByRef Animate As Boolean = False, _
    Optional ByRef IsFio As Boolean = False, _
    Optional SkipWords As String _
    ) As String
' склонение слов. работает весьма условно
'-------------------------
' Word - существительное в единственном числе именительном падеже
' NewCase - падеж ("и","р","д","в","т","п")
' NewNumb - число ("ед","мн")
' NewGend - род ("м","ж")
' Animate - признак одушевлённости / неодушевлённости
' IsFIO - признак ФИО
' SkipWords - список номеров слов, которые необходимо пропустить при склонении словосочетания
'   начиная с 1, по возрастанию через ",", диапазон через "-"
'   например: "2,4-5,7-" - при склонении будут пропущены все слова кроме 1,3 и 6
'-------------------------
' для лучших результатов пользуйтесь Morpher'ом с http://morpher.ru
' или Padej'ом http://www.delphikingdom.com/asp/viewitem.asp?catalogid=412
' или модифицируйте правила и добавляйте мсключения под себя
'-------------------------
' v.1.0.2       : 06.07.2019 - в SkipWords добавлена возможноость задавать диапазоны
' v.1.0.1       : 05.07.2019 - исправлена ошибка при пересборке строки старый способ с Replace приводил к ошибка при наличии повторов в строке
'-------------------------
Const cstrListDelim = ",", cstrDiapDelim = "-"
Dim strWord As String, strTail As String, aWords() As String, aSkip() As String
Dim i As Long, iMax As Long ' позиция в строке
Dim j As Long, jMin As Long ' номер склоняемого слова
Dim n As Long, nMin As Long ' номер элемента в списке исключаемых слов
Dim s As String, sArr() As String ' элемент списка исключаемых слов
Dim sNum As Byte, sMin As Byte ' индексы границ элемента списка
Dim S1 As Long, S2 As Long  ' границы элемента списка
Dim tmpGender As DeclineGend
Dim Result As String
'
    On Error GoTo HandleError
    Result = vbNullString
    strTail = Trim$(Words): iMax = Len(strTail)
    Call Tokenize(Words, aWords): jMin = LBound(aWords): j = UBound(aWords)
    ' получаем список слов которые необходимо пропустить
    n = -1: SkipWords = Trim$(SkipWords)
    If Len(SkipWords) > 0 Then aSkip = Split(SkipWords, cstrListDelim): nMin = LBound(aSkip): n = UBound(aSkip) + 1
    'If Len(SkipWords) > 0 Then Call xSplit(SkipWords, aSkip, cstrListDelim): nMin = LBound(aSkip): n = UBound(aSkip) + 1
    S1 = j - jMin + 1: S2 = 0
    Do While j >= jMin
    ' перебираем слова от конца к началу
    ' проверяем принадлежность слова текущему диапазону
        If n < nMin Then GoTo HandleText
        If S1 <= j - jMin + 1 And S1 <> 0 And S2 <> 0 Then GoTo HandleText
HandleDiap:
        n = n - 1: If n < nMin Then S1 = 0: S2 = 0: GoTo HandleText
    ' получаем элемент списка диапазона пропуска
        ' проверяем в элементе списка пропуска наличие символа диапазона
        s = Trim$(aSkip(n))
        sArr() = Split(s, cstrDiapDelim) 'Call xSplit(S, sArr(), cstrDiapDelim)
        sMin = LBound(sArr): sNum = UBound(sArr) - sMin + 1
        Select Case sNum
        Case 1
        ' "s1" - в диапазоне один числовой элемент
            If IsNumeric(sArr(sMin)) Then
                If S1 <> 0 Then S2 = CLng(Trim$(sArr(sMin)))
                S1 = CLng(Trim$(sArr(sMin)))
            End If
        Case 2 '
            If IsNumeric(sArr(sMin)) And IsNumeric(sArr(sMin + 1)) Then
        ' "s1-s2" -  в диапазоне два числовых элемента
            ' берем численные значения верхней и нижней границы
                If S1 <> 0 Then S2 = CLng(Trim$(sArr(sMin + 1)))
                S1 = CLng(Trim$(sArr(sMin)))
            ElseIf IsNumeric(sArr(sMin)) Then
        ' "s1-" -    в диапазоне один числовой элемент, верхняя граница открыта
            ' новая верхняя граница равна предыдущей нижней
                If S1 <> 0 Then S2 = S1
                S1 = CLng(Trim$(sArr(sMin)))
            ElseIf IsNumeric(sArr(sMin + 1)) Then
        ' "-s2" -    в диапазоне один числовой элемент, нижняя граница открыта
            ' ищем нижнюю границу перебираем диапазоны к началу
            ' пока не найдем диапазон с численным началом или не исчерпаем список
                S2 = CLng(Trim$(sArr(sMin + 1)))
                If n = nMin Then S1 = 1 Else S1 = 0: GoTo HandleDiap
            Else
        ' неведома фигня
                Stop
            End If
        Case Else: Stop ' "-n-","--n" и т.п. в диапазоне нет элементов или больше 2 ???
        End Select
HandleText:
    ' берём склоняемое слово
        strWord = aWords(j)
    ' ищем начало склоняемого слова в строке (с конца строки)
        i = InStrRev(strTail, strWord)
    ' добавляем в начало строки разделители из исходной строки
        Result = Right$(strTail, iMax - i - Len(strWord) + 1) & Result
    ' обрезаем исходную строку по началу склоняемого слова
        iMax = i - 1: strTail = Left$(strTail, iMax)
    ' проверяем номер слова по списку пропуска
        Select Case j - jMin + 1 ' номера текущего слова
        Case S1 To S2:  ' попадает в границы диапазона пропуска - пропускаем
            'newWord = strWord
        Case Else:      ' не попадает - склоняем слово в строке
            tmpGender = NewGend
            'newWord = DeclineWord(strWord, NewCase, NewNumb, tmpGender, Animate, IsFio:=IIf(IsFio, j - jMin + 1, 0))
            strWord = DeclineWord(strWord, NewCase, NewNumb, tmpGender, Animate, IsFio:=IIf(IsFio, j - jMin + 1, 0))
        End Select
    ' добавляем в начало строки слово получившееся после склонения исходного
        Result = strWord & Result 'Result = newWord & Result
    ' переходим к следующему слову
HandleNext:
        j = j - 1
    Loop
    ' добавляем оставшиеся разделители
    Result = strTail & Result
    Erase aWords: Erase aSkip
HandleExit:  DeclineWords = Result: Exit Function
HandleError: i = iMax: Err.Clear: Resume HandleExit ': WordEnd = vbNullString
End Function

Private Function p_NumDecline( _
    ByRef Numb As Variant, Optional Triplet As Byte = 0, _
    Optional ByVal NewCase As DeclineCase = DeclineCaseRod, _
    Optional ByVal NewNumb As DeclineNumb = DeclineNumbSingle, _
    Optional ByRef NewGend As DeclineGend = DeclineGendUndef, _
    Optional ByVal NewNumeralType As NumeralType = NumeralOrdinal, _
    Optional ByRef Animate As Boolean = False, _
    Optional ByRef SymbCase As Integer = 0, Optional ByRef Template As String, _
    Optional ByRef NumbRest As Integer = 0 _
    ) As String
' вспомогательная для склонение слов основных числительных: 0-9,10-19,x0,x00,1000,100000...
'-------------------------
' Numb - строка содержащая целое число (без пробелов и пр.символов), либо одна из текстовых констант чисел (см.p_NumWordsArray)
' NewCase - падеж ("и","р","д","в","т","п")
' NewNumb - число ("ед","мн")
' NewGend - род ("м","ж")
' NewNumeralType - тип числительного ("кол","пор") - применимо только для числительных
' Animate - признак одушевленности, нужен для правильного склонения в вин.п.
' SymbCase, Template - состояние регистра символов исходного слова и шаблон
' Triplet - номер анализируемого триплета. если =0 или >len(Numb)\3-1 - анализируется младший триплет
' NumbRest - если входной параметр был задан числовой строкой,- возвращает неразобранный остаток анализируемого триплета
'-------------------------
' выделяем в отдельную функцию чтобы сократить количество проверок когда нужно только число (например при вызове из NumToWords)
' функция возвращает распознаваемое число в текстовом виде, в соотв склонении
' Результат работы функции:
'   Numb=123, Triplet=0 - разбираем число 123       >> Result="сто",NumbRest=23
'   Numb=123, Triplet=1 - разбираем число 123000    >> Result="тысяча",NumbRest=123
'   Numb=123000 - разбираем число 123000            >> Result="тысяча",Triplet=1,NumbRest=123
'-------------------------
Dim WordWhole As String, WordBeg As String, WordEnd As String
Dim i As Long, iMax As Long
Dim Result As String

    On Error GoTo HandleError
    If NewGend = DeclineGendNeut Then Animate = False
    If NewNumeralType = NumeralUndef Then NewNumeralType = NumeralOrdinal
    If IsNumeric(Numb) Then
' если передано число цифрами - получаем индекс в массиве и соответствующее слово
    On Error Resume Next
        NumbRest = CLng(Numb): i = Err.Number: Err.Clear
    On Error GoTo HandleError
        If i = 0 And NumbRest < 1000 And Triplet = 0 Then
        ' нет ошибки, число в диапазоне 0..999 и номер триплета не задан
'    ' разбор триплета слева направо (сотни>десятки>единицы)
'            Select Case NumbRest
'            Case 0 To 19:       i = NumbRest:               NumbRest = 0
'            Case 20 To 99:      i = 18 + NumbRest \ 10:     NumbRest = NumbRest Mod 10
'            Case 100 To 999:    i = 27 + NumbRest \ 100:    NumbRest = NumbRest Mod 100
'            End Select
    ' разбор триплета справа налево (единицы>десятки>сотни)
            i = NumbRest
            If i > 0 Then ' не пустой триплет
                i = NumbRest Mod 100 ' смотрим хвост
            If i = 0 Then                                   ' x00 - сотни (x=1-9)
                i = 27 + NumbRest \ 100: NumbRest = 0
            ElseIf i >= 20 And (i Mod 10 = 0) Then     ' xy0 - десятки (y=2-9)
                i = 18 + i \ 10:  NumbRest = 100 * (NumbRest \ 100)
            ElseIf i < 20 Then                              ' x1z - второй десяток (z=0-9)
                NumbRest = 100 * (NumbRest \ 100)
            Else                                            ' xyz - первый десяток (z=1-9)
                i = i Mod 10: NumbRest = 10 * (NumbRest \ 10)
            End If
            End If
        Else
        ' число >1000 ( >Long, >1000 и <Long, указан триплет)
            iMax = Len(Numb)
            If Triplet = 0 Then Triplet = (iMax - 1) \ 3    ' номер старшего триплета
            i = iMax - 3 * Triplet: If i < 1 Then i = iMax  '
            NumbRest = Right$(Left$(Numb, i), 3)                 ' численное значение соотв триплета
            i = 36 + Triplet  ' индекс текстового значения (порядка триплета) в массиве текстовых констант
        End If
    ' получаем значение текстовой константы соответствующей части числа
        Result = p_NumWordsArray(i): WordWhole = Result: iMax = Len(Result)
    Else
' если передано слово - получаем индекс в массиве
        Result = Trim$(Numb): WordWhole = LCase$(Result) ': WordWhole = Replace(WordWhole, "ё", "е")
        If SymbCase = 0 Then SymbCase = p_GetSymbCase(Result, Template)
        i = 0: iMax = Len(Result)
        Do Until p_NumWordsArray(i) = WordWhole
            If i <= UBound(p_NumWordsArray) Then i = i + 1 Else Err.Raise 6
        Loop
    End If
' Подготавливаем слово к склонению
    If NewCase > DeclineCasePred Then Err.Raise vbObjectError + 512
    If NewNumb = DeclineNumbUndef Then NewNumb = DeclineNumbSingle
    If NewGend = DeclineGendUndef Then NewGend = DeclineGendMale
    If NewNumeralType = NumeralOrdinal Then
    ' количественные числительные
        Select Case i
        Case 0
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = Choose(NewNumb, "ь", "и")
        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "я", "ей")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "ю", "ям")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ём", "ями")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ях")
        End Select
        i = Len(Result) - 1: GoTo HandleExit
        Case 1
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewNumb, Choose(NewGend, "ин", "на", "но"), "ни")
        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, Choose(NewGend, "ного", "ной", "ного"), "них")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, Choose(NewGend, "ному", "ной", "ному"), "ним")
        Case DeclineCaseVin:  WordEnd = Choose(NewNumb, Choose(NewGend, IIf(Animate, "ного", "ин"), "ну", "но"), "них")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, Choose(NewGend, "ним", "ной", "ним"), "ними")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, Choose(NewGend, "ном", "ной", "ном"), "них")
        End Select
        i = Len(Result) - 2: GoTo HandleExit
        Case 2
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewGend, "а", "е", "а")
        Case DeclineCaseRod, DeclineCasePred: WordEnd = "ух"
        Case DeclineCaseVin:  WordEnd = IIf(Animate, "ух", Choose(NewGend, "а", "е", "а"))
        Case DeclineCaseDat: WordEnd = "ум"
        Case DeclineCaseTvor: WordEnd = "умя"
        End Select
        i = Len(Result) - 1: GoTo HandleExit
        Case 3
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = "и"
        Case DeclineCaseRod, DeclineCasePred: WordEnd = "ёх"
        Case DeclineCaseVin: WordEnd = IIf(Animate, "ёх", "и")
        Case DeclineCaseDat: WordEnd = "ём"
        Case DeclineCaseTvor: WordEnd = "емя"
        End Select
        i = Len(Result) - 1: GoTo HandleExit
        Case 4
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = "е"
        Case DeclineCaseRod, DeclineCasePred: WordEnd = "ёх"
        Case DeclineCaseVin: WordEnd = IIf(Animate, "ёх", "е")
        Case DeclineCaseDat: WordEnd = "ём"
        Case DeclineCaseTvor: WordEnd = "ьмя"
        End Select
        i = Len(Result) - 1: GoTo HandleExit
        Case 5 To 7, 9, 10 To 21 '5-7,9,10-20,30
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = "ь"
        Case DeclineCaseRod, DeclineCaseDat, DeclineCasePred: WordEnd = "и"
        Case DeclineCaseTvor: WordEnd = "ью"
        End Select
        i = Len(Result) - 1: GoTo HandleExit
        Case 8
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = "емь"
        Case DeclineCaseRod, DeclineCaseDat, DeclineCasePred: WordEnd = "ьми"
        Case DeclineCaseTvor: WordEnd = "ьмью"
        End Select
        i = Len(Result) - 3: GoTo HandleExit
        Case 22 '40
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = ""
        Case DeclineCaseRod, DeclineCaseDat, DeclineCaseTvor, DeclineCasePred: WordEnd = "а"
        End Select
        i = Len(Result): GoTo HandleExit
        Case 23 To 26 '50-80 (для 80 - чередование -е-/-ь-)
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = "ьдесят"
        Case DeclineCaseRod, DeclineCaseDat, DeclineCasePred: WordEnd = "идесяти":  If i = 26 Then Mid$(Result, 4, 1) = "ь"
        Case DeclineCaseTvor: WordEnd = "ьюдесятью":      If i = 26 Then Mid$(Result, 4, 1) = "ь"
        End Select
        i = Len(Result) - 6: GoTo HandleExit
        Case 27, 28 '90,100
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = "о"
        Case DeclineCaseRod, DeclineCaseDat, DeclineCaseTvor, DeclineCasePred: WordEnd = "а"
        End Select
        i = Len(Result) - 1: GoTo HandleExit
        Case 29 '200
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = "ести"
        Case DeclineCaseRod:  WordEnd = "ухсот"
        Case DeclineCaseDat:  WordEnd = "умстам"
        Case DeclineCaseTvor: WordEnd = "умястами"
        Case DeclineCasePred: WordEnd = "ухстах"
        End Select
        i = Len(Result) - 4: GoTo HandleExit
        Case 30, 31 '300,400
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = VBA.Right$(Result, 4)
        Case DeclineCaseRod:  WordEnd = "ёхсот"
        Case DeclineCaseDat:  WordEnd = "ёмстам"
        Case DeclineCaseTvor: WordEnd = IIf(i = 30, "е", "ь") & "мястами"
        Case DeclineCasePred: WordEnd = "ёхстах"
        End Select
        i = Len(Result) - 4: GoTo HandleExit
        Case 32 To 36 '500-900 (для 800 - чередование -е-/-ь-)
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = "ьсот"
        Case DeclineCaseRod:  WordEnd = "исот":    If i = 35 Then Mid$(Result, 4, 1) = "ь"
        Case DeclineCaseDat:  WordEnd = "истам":   If i = 35 Then Mid$(Result, 4, 1) = "ь"
        Case DeclineCaseTvor: WordEnd = "ьюстами": If i = 35 Then Mid$(Result, 4, 1) = "ь"
        Case DeclineCasePred: WordEnd = "истах":   If i = 35 Then Mid$(Result, 4, 1) = "ь"
        End Select
        i = Len(Result) - 4: GoTo HandleExit
        ' в принципе 1E3,1E6 и т.д. нормально склоняются DeclineWord
        Case 37 '1000
        NewGend = DeclineGendFem
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewNumb, "а", "и")
        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "и", "")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "е", "ам")
        Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "у", "и")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ей", "ами") ', "ью", "ами")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ах")
        End Select
        i = Len(Result) - 1: GoTo HandleExit
        Case 38 To 47 '1E6,1E9,...
        NewGend = DeclineGendMale
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = Choose(NewNumb, "", "ы")
        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "а", "ов")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "у", "ам")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "ом", "ами")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "е", "ах")
        End Select
        i = Len(Result): GoTo HandleExit
        Case Else: Err.Raise 6
        End Select
        Result = p_SetSymbCase(Left$(Result, i) & WordEnd, SymbCase, Template)
    Else
    ' порядковые числительные
        ' преобразуем в порядковое и склоняем как прилагательное
        Select Case i
        Case 0: WordEnd = "нулевой":    i = 0: Result = vbNullString
        Case 1: WordEnd = "первый":     i = 0: Result = vbNullString
        Case 2: WordEnd = "второй":     i = 0: Result = vbNullString
        Case 3: WordEnd = "етий":       i = 2
        Case 4: WordEnd = "вёртый":     i = 3
        Case 6: WordEnd = "ой":         i = Len(Result) - 1
        Case 7: WordEnd = "дьмой":      i = 2
        Case 8: WordEnd = "ьмой":       i = 3
        Case 5, 9 To 21, 27: WordEnd = "ый": i = Len(Result) - 1        ' 5,9-19,20,30,90
        Case 22: WordEnd = "овой":      i = Len(Result)                 ' 40
        Case 23 To 26: WordEnd = "ый":  i = Len(Result)                 ' 50-80
        Case 28: WordEnd = "отый":      i = 1                           ' 100
        Case 29: WordEnd = "ухсотый":   i = 2                           ' 200
        Case 30, 31: WordEnd = "ёхсотый":   i = Len(Result) - 4         ' 300,400
        Case 32 To 34, 36: WordEnd = "исотый":   i = Len(Result) - 4    ' 500-700,900
        Case 35: WordEnd = "ьмисотый":  i = 3                           ' 800
        Case 37: WordEnd = "ный":       i = Len(Result) - 1             ' 1000
        Case 38 To 47:  WordEnd = "ный": i = Len(Result)                ' 10^6, 10^9 etc
        Case Else: Err.Raise 6
        End Select
        Result = p_SetSymbCase(Left$(Result, i) & WordEnd, SymbCase, Template)
    ' склоняем как прилагательное
        Result = DeclineWord(Result, NewCase, NewNumb, NewGend, Animate)  ', SymbCase:=SymbCase, Template:=Template)
        WordEnd = vbNullString: i = Len(Result)
    End If
HandleExit:  p_NumDecline = Left$(Result, i) & WordEnd: Exit Function
HandleError: i = iMax: WordEnd = vbNullString: Err.Clear: Resume HandleExit
End Function
Private Function p_NumWordsArray()
' массив текстовых констант для числительных
'-------------------------
' i = 00..09    -   единицы      x,     где x=0-9
' i = 10..19    -   1й десяток  1x,     где x=0-9
' i = 20..27    -   десятки     x0,     где x=2-9
' i = 28..36    -   сотни       x00,    где x=1-9
' i = 37..47    -   тысячи и д. 10^(3*x), где x=1-11
'-------------------------
On Error Resume Next
Static arrData(), iMin As Long: iMin = LBound(arrData): If Err Then Err.Clear Else p_NumWordsArray = arrData: Exit Function
    arrData = Array( _
        "ноль", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять", _
        "десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать", _
        "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто", _
        "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот", _
        "тысяча", "миллион", "миллиард", "триллион", "квадриллион", "квинтиллион", "секстиллион", "септиллион", "октиллион", "нониллион", "дециллион")
HandleExit: p_NumWordsArray = arrData
End Function
Private Function p_GetWordTemplate(ByVal Word As String, Optional CheckCase As Boolean = False) As String
' Создает на шаблон слова для последующего анализа
'-------------------------
' CheckCase - определяет будет ли при создании шаблона учитываться регистр символа
'-------------------------
Dim Result As String: Result = vbNullString
    On Error GoTo HandleError
Dim sArr:   sArr = Array(c_strSymbRusSign & c_strSymbEngSign, c_strSymbRusVowel & c_strSymbEngVowel, c_strSymbRusConson & c_strSymbEngConson)
' заменяем символы исходной строки их обозначениями в паттерне (x, g, s)
Dim i As Long, j As Long, m As String, s As String
    s = "xgs" ' символы для подстановки
    For i = 1 To Len(Word)
    ' перебор символов слова
        m = Mid$(Word, i, 1)
        ' если учитываем регистр символа меняем регистр символов подстановки в соответствии с регистром символа
        If CheckCase Then If m = LCase(m) Then s = LCase(s) Else s = UCase(s)
        m = LCase$(m)
        For j = 0 To UBound(sArr)
        ' перебор элементов массива подстановки
            If InStr(sArr(j), m) Then Mid$(Word, i, 1) = Mid$(s, j + 1, 1): Exit For
    Next j, i
    Result = Word
HandleExit:  p_GetWordTemplate = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function

Private Function p_GetWordParts(ByVal Word As String, _
    Optional ByRef WordBeg As String, Optional ByRef WordEnd As String, _
    Optional ByRef Template As String) As Boolean
' выделяет в слове начало и окончание (условно)
'-------------------------
Dim Result As Boolean ' Result = False
    On Error GoTo HandleError
    If Len(Template) = 0 Then Template = p_GetWordTemplate(Word)
    Dim i As Long: i = Len(Template)
' окончанием считаем все что идет справа от первой (кроме
' согласной, стоящей в самом конце слова) согласной с конца
    i = InStrRev(LCase$(Template), "s", i - 1)
    WordBeg = Left$(Word, i): WordEnd = Mid$(Word, i + 1)
    Result = True
HandleExit:  p_GetWordParts = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function p_GetWordParts2(ByVal Word As String, _
    Optional ByRef WordBeg As String, Optional ByRef WordEnd As String _
    ) As Boolean
' выделяет в слове начало и окончание (старый вариант)
'-------------------------
Dim sChar As String * 1
Dim fStop As Boolean, c As Long, v As Long
Dim i As Long, iMax As Long
Dim Result As Boolean

    Result = False
    On Error GoTo HandleError
    Word = LCase$(Word)
    iMax = Len(Word): i = iMax: c = 0: v = 0: fStop = False
    WordBeg = Word: WordEnd = vbNullString
    Do Until fStop
        i = i - 1
        sChar = LCase$(Right$(WordBeg, 1))
        Select Case GetCharType(Right$(WordBeg, 1))
        Case SymbolTypeVowel: WordEnd = sChar & WordEnd: WordBeg = Left$(Word, i): v = v + 1
        Case SymbolTypeCons: If v > 0 Or c = 2 Then fStop = True Else WordEnd = sChar & WordEnd: WordBeg = Left$(Word, i): c = c + 1
        Case SymbolTypeSign: WordBeg = Left$(Word, i): If v = 0 And c = 0 Then WordEnd = sChar & WordEnd: Exit Do
        Case Else: Err.Raise vbObjectError + 512
        End Select
        If i < 1 And Not fStop Then WordBeg = Word: WordEnd = vbNullString: fStop = True
    Loop
    Result = True
HandleExit:  p_GetWordParts2 = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Private Function p_GetWordGender(ByVal Word As String, Optional ByRef WordEnd As String) As DeclineGend
' определяет род по окончанию слова (условно)
'-------------------------
Dim Result As DeclineGend

    Result = DeclineGendUndef
    On Error GoTo HandleError
    Word = LCase$(Word)
    If Len(WordEnd) = 0 Then Call p_GetWordParts(Word, WordEnd:=WordEnd)
    Select Case WordEnd
    'Мужской имеют окончания -а, -я, и нулевое (папа, дядя, нож, стол, ястреб)
    'Женский имеют окончания -а, -я, и нулевое (жена, няня, ночь, слава, пустыня)
    Case "ь"
        Select Case LCase$(Word)
        Case "ноль", "рубль", "конь", _
             "огонь", "уголь", "февраль": Result = 1   'мужской род
        Case "лень", "тень", "сень", _
             "сажень": Result = 2           'женский род
        Case Else:
            Select Case VBA.Left$(VBA.Right$(Word, 2), 1)
            Case "з", "ч", "ш", "ж":    Result = 2  'женский род (бязь,ночь,брешь,рожь...)
            Case "б", "п", "л":         Result = 2  'женский род (рябь,выпь,боль но - ноль)
            Case "н":
                Select Case VBA.Left$(VBA.Right$(Word, 3), 1)
                Case "е":               Result = 1  'мужской род (-ень)
                Case "о":               Result = 2  'женский род (-онь)
                Case Else:              Result = 1  'мужской род
                End Select
            Case Else:                  Result = 1  'мужской род (конь,пень...)
            End Select
        End Select
    Case "а", "я"
        Select Case Word
        Case "папа", "дядя", "дедушка": Result = 1  'мужской род
        Case Else:                      Result = 2  'женский род
        End Select
'    ''Общий род - в зависимости от контекста, могут употребляться и в мужском, и в женском роде
'    ''    (зануда, неженка, плакса, умница, жадина).
'    Case "ий", "ый":                    NewGend = 1 'мужской род
    Case "ая", "яя":                    Result = 2  'женский род
    Case "ова", "ева", "ёва":           Result = 2  'женский род
    Case "о", "е", "ое", "ее", _
         "ё", "оё", "её": Result = 3               'средний род
    Case Else
'    ''мужской род (последняя согласная)
        If GetCharType(Right$(WordEnd, 1)) = SymbolTypeCons Then Result = 1
    End Select
HandleExit:  p_GetWordGender = Result: Exit Function
HandleError: Result = DeclineGendUndef: Err.Clear: Resume HandleExit
End Function

Private Function p_GetWordSpeechPartType(ByVal Word As String) As SpeechPartType
' определяет часть речи по окончанию слова (условно)
'-------------------------
' не знаю зачем я это сделал, - разве так, на будущее...
' может когда и перепишу процедуру склонения с учётом части речи
'-------------------------
Dim Result As SpeechPartType

    On Error GoTo HandleError
    Select Case Word
    ' местоимения
    Case "я", "ты", "он", "она", "оно", "то", "это", "тот", "этот", _
        "вы", "мы", "они", "те", "эти":
            Result = SpeechPartTypePronoun
    ' предлоги
    Case "в", "с", "к", "у", "а", "и", "о", "об", "на", "по", "во", "за", "от", "не", "ни", "ли", _
         "или", "еле", "над", "при", "под", "для", "через", "перед", "ввиду", "наподобие", "вроде", _
         "вблизи", "вглубь", "вдоль", "возле", "около", "среди", "вокруг", "внутри", "впереди", "после", _
         "насчет", "навстречу", "вслед", "вместо", "ввиду", "благодаря", "вследствие"
            Result = SpeechPartTypePreposition
    ' существительные на -ть и ая,ие и т.п.
    Case "мать", "рать", "тать", "зять", "суть", "путь", "муть", "нежить", "пажить", "сыть", "нить", "лапоть", "копоть", _
         "стая" ', "событие", "предложение", "поручение", "последствие", "преследование", "лезвие", _
         "сомнение"
            Result = SpeechPartTypeNoun
    Case Else
'        If Len(WordEnd) = 0 Then Call p_GetWordParts(Word, WordEnd:=WordEnd)
        Select Case Right$(Word, 2)
    ' прилагательные
        Case "ая", "яя", "ую", "юю", "ое", "ее", "оё", "её", "иё", "ий", "ый", "ой" ', "ие" ' - слишком много сущ. на -ие
            Result = SpeechPartTypeAdject: GoTo HandleExit
        End Select
    ' числительные
        Dim tmp: For Each tmp In p_NumWordsArray
            If Word = tmp Then _
            Result = SpeechPartTypeNumeral: GoTo HandleExit
        Next tmp
        Select Case Right$(Word, 3)
    ' глаголы
        Case "ать", "ять", "уть", "оть", "еть", "ить", "ыть"
        ' при этом: мать,рать и т.п. - сущ.,а пять и на -дцать - числительные
            Result = SpeechPartTypeVerb: GoTo HandleExit
        End Select
    ' остальные считаем существительными
            Result = SpeechPartTypeNoun
    End Select
HandleExit:  p_GetWordSpeechPartType = Result: Exit Function
HandleError: Result = SpeechPartTypeUndef: Err.Clear: Resume HandleExit
End Function

Private Function p_GetSymbCase(Word As String, Optional Template As String) As Integer
' возвращает состояние регистра символов слова
'-------------------------
Dim Result As Integer
    Result = False
    On Error GoTo HandleError
' 0 - не определено
    If Word = UCase$(Word) Then
' 1 (vbUpperCase) - все символы в верхнем регистре
        Result = vbUpperCase
    ElseIf Word = LCase$(Word) Then
' 2 (vbLowerCase) - все символы в нижнем регистре
        Result = vbLowerCase
    ElseIf Word = StrConv(Word, vbProperCase) Then
' 3 (vbProperCase) - первый символ в верхнем остальные, - в нижнем регистре
        Result = vbProperCase
    Else
'-1 - регистр символов определяется по шаблону (часть букв в верхнем, часть - в нижнем регистре)
    ' формируем шаблон регистра по слову
        Template = p_GetWordTemplate(Word, True): Result = -1
    End If
HandleExit:  p_GetSymbCase = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Private Function p_SetSymbCase(Word As String, SymbCase As Integer, Optional ByVal Template As String) As String
' устанавливает состояние регистра символов слова
'-------------------------
Dim Result As String
    Result = Word
    On Error GoTo HandleError
    Select Case SymbCase
    Case vbUpperCase:   Result = UCase$(Word)
    Case vbLowerCase:   Result = LCase$(Word)
    Case vbProperCase:  Result = StrConv(Word, vbProperCase)
    Case -1             ' форматируем по шаблону
' нафига это?? - нууу... - просто эксперимент - сам не знаю, а вдруг? )
' впрочем все равно это все пока сложно и неправильно,
' а случай с разнокалиберными регистрами в слове можно просто игнорировать
' в силу сильно упрощенного алгоритма определения окончаний часто работает не корректно
        ' формируем шаблон текущего слова и сравниваем их
        Dim s As String * 1, с As Integer
        Dim i As Long, iMax As Long: i = 1: iMax = Len(Template)
        Dim NewTemp As String: NewTemp = p_GetWordTemplate(Word, True)
        If LCase$(Template) = LCase$(NewTemp) Then
        ' если совпадают - форматируем по шаблону
            Do Until i > iMax
                s = Mid$(Template, i, 1)
                If LCase(s) = s Then
                    Mid$(Result, i, 1) = LCase$(Mid$(Result, i, 1))
                Else
                    Mid$(Result, i, 1) = UCase$(Mid$(Result, i, 1))
                End If
                i = i + 1
            Loop
        Else
        ' если нет приводим исходный шаблон к шаблону текущего слова
        Dim j As Long, jMax As Long: j = 1: jMax = Len(NewTemp)
        ' форматируем по следующему правилу
            ' если тип символа (xgs) приводимого (исходного) шаблона совпадает
            ' с типом шаблона нового слова форматируем как в приводимом шаблоне,
            ' если не совпадает - повторяем регистр предыдущего символа
            ' знаки ьъ форматируем как гласные
            Template = Replace$(Template, "x", "g"): NewTemp = Replace$(NewTemp, "x", "g")
            ' первый символ берем в регистре исходного шаблона
            s = Mid$(Template, i, 1)
            If LCase$(s) = s Then
                Mid$(Result, j, 1) = LCase(Left$(Result, 1))
            Else
                Mid$(Result, j, 1) = UCase(Left$(Result, 1))
            End If
            i = i + 1: j = j + 1
            Do Until j > jMax
                If i > iMax Then i = 2 ' если центральная часть приводимого закончилась - начинаем сначала
                s = Mid$(Template, i, 1)
                ' если тип символа приводимого шаблона совпадает с типом шаблона нового слова
                ' берем регистр символа приводимого шаблона и присваиваем символу нового
                If LCase(Mid$(NewTemp, j, 1)) = LCase(s) Then
                    i = i + 1
                    If LCase(s) = s Then
                        Mid$(Result, j, 1) = LCase$(Mid$(Result, j, 1))
                        Mid$(NewTemp, j, 1) = LCase$(Mid$(NewTemp, j, 1))
                    Else
                        Mid$(Result, j, 1) = UCase$(Mid$(Result, j, 1))
                        Mid$(NewTemp, j, 1) = UCase$(Mid$(NewTemp, j, 1))
                    End If
                End If
                j = j + 1
            Loop
            Template = NewTemp
        End If
        ' собственно форматируем слово
    Case Else ' оставляем как есть
    End Select
HandleExit:  p_SetSymbCase = Result: Exit Function
HandleError: Result = Word: Err.Clear: Resume HandleExit
End Function

'=========================
' Вспомогательные
'=========================
Private Function Pwr2(Index) As Long
' возвращает степени числа 2. нужна для битовых операций
'-------------------------
    On Error GoTo HandleError
    If Index < 0 Or Index > 31 Then Err.Raise vbObjectError
    Select Case Index
    Case 0:     Pwr2 = &H1&
    Case 1:     Pwr2 = &H2&
    Case 2:     Pwr2 = &H4&
    Case 3:     Pwr2 = &H8&
    Case 4:     Pwr2 = &H10&
    Case 5:     Pwr2 = &H20&
    Case 6:     Pwr2 = &H40&
    Case 7:     Pwr2 = &H80&
    Case 8:     Pwr2 = &H100&
    Case 9:     Pwr2 = &H200&
    Case 10:    Pwr2 = &H400&
    Case 11:    Pwr2 = &H800&
    Case 12:    Pwr2 = &H1000&
    Case 13:    Pwr2 = &H2000&
    Case 14:    Pwr2 = &H4000&
    Case 15:    Pwr2 = &H8000&
    Case 16:    Pwr2 = &H10000
    Case 17:    Pwr2 = &H20000
    Case 18:    Pwr2 = &H40000
    Case 19:    Pwr2 = &H80000
    Case 20:    Pwr2 = &H100000
    Case 21:    Pwr2 = &H200000
    Case 22:    Pwr2 = &H400000
    Case 23:    Pwr2 = &H800000
    Case 24:    Pwr2 = &H1000000
    Case 25:    Pwr2 = &H2000000
    Case 26:    Pwr2 = &H4000000
    Case 27:    Pwr2 = &H8000000
    Case 28:    Pwr2 = &H10000000
    Case 29:    Pwr2 = &H20000000
    Case 30:    Pwr2 = &H40000000
    Case 31:    Pwr2 = &H80000000
    Case Else:  Pwr2 = 0
    End Select
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Private Function p_GetLocaleInfo(LCType As Long) As String
' возвращает региональные настройки
'-------------------------
' lcType - константы LOCALE_
Dim lpBuffer As String * 100
    On Error GoTo HandleError
    If GetLocaleInfo(LOCALE_USER_DEFAULT, LCType, lpBuffer, 99) = 0 Then Err.Raise vbObjectError + 512
    p_GetLocaleInfo = Left$(lpBuffer, InStr(lpBuffer, Chr$(0)) - 1)
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
'    Select Case Err.Number
'    Case vbObjectError + 512: MsgBox "Unknown constant", vbOKOnly Or vbExclamation, "Error!"
'    End Select
End Function
Private Function p_GetCollKeys(oColl As Collection, Optional oIdxs) As String()
' возвращает массив ключей коллекции (Base=1)
'-------------------------
' оригинал взят отсюда: https://stackoverflow.com/questions/5702362/vba-collection-list-of-keys
'-------------------------
Dim CollPtr As LongPtr, KeyPtr As LongPtr, ItemPtr As LongPtr, Address As LongPtr
Dim bIdxs As Boolean: bIdxs = Not IsMissing(oIdxs): If bIdxs Then Set oIdxs = New Collection
Dim Result() As String, Length As Long
Dim i As Long, iMax As Long
    CollPtr = VBA.ObjPtr(oColl)                             ' адрес коллекции в памяти
    Address = CollPtr + 3 * PTR_LENGTH + 4                  ' адрес количества элементов коллекции
    If Address <> 0 Then Call CopyMemory(ByVal VarPtr(iMax), ByVal Address, 4)
    If iMax <> oColl.Count Then Stop                        ' не совпадает с количеством возвращаемым объектом - ошибка!
    ReDim Result(1 To iMax)                                 ' объявляем массив для хранения ключей
    Address = CollPtr + 4 * PTR_LENGTH + 8                  ' адрес первого элемента коллекции
    If Address <> 0 Then Call CopyMemory(ByVal VarPtr(ItemPtr), ByVal Address, PTR_LENGTH)
    For i = 1 To iMax
        If ItemPtr = 0 Then Exit For
        Address = ItemPtr + 2 * PTR_LENGTH + 8              ' адрес ключа элемента коллекции
        If Address <> 0 Then Call CopyMemory(ByVal VarPtr(KeyPtr), ByVal Address, PTR_LENGTH)
        If KeyPtr <> 0 Then                                 ' извлекаем ключ элемента коллекции
        Call CopyMemory(ByVal VarPtr(Length), ByVal KeyPtr - 4, PTR_LENGTH)
        Result(i) = Space(Length \ 2)
        Call CopyMemory(ByVal StrPtr(Result(i)), ByVal KeyPtr, ByVal Length)
        End If
        Address = ItemPtr + 4 * PTR_LENGTH + 8              ' адрес следующего элемента коллекции
        If Address <> 0 Then Call CopyMemory(ByVal VarPtr(ItemPtr), ByVal Address, PTR_LENGTH)
        If bIdxs Then oIdxs.Add i, Result(i)                ' если также надо получить коллекцию соответствий тегов индексам
    Next i
    p_GetCollKeys = Result
End Function
Private Function p_HFontByControl(Optional ctl As Variant, Optional FontName, Optional FontSize, _
    Optional FontColor, Optional FontWeight, Optional FontUnderline, Optional FontStrikeOut, Optional FontItalic, Optional hdc As LongPtr = 0) As LongPtr
' создает hFont из параметров контрола
'-------------------------
    'If Not TypeOf ctl Is Access.Control Then Err.Raise vbObjectError + 512
Dim tDC As LongPtr, hFont As LongPtr
    If hdc = 0 Then tDC = GetDC(0) Else tDC = hdc
' создаём шрифт
On Error Resume Next
Dim fName As String:    fName = IIf(IsMissing(FontName), ctl.FontName, FontName): If Err Then fName = "Arial": Err.Clear
Dim fSize As Long:      fSize = IIf(IsMissing(FontSize), ctl.FontSize, FontSize): If Err Then fSize = 10: Err.Clear
'Dim fColor As Long:     fColor = IIf(IsMissing(FontColor), ctl.ForeColor, FontColor): If Err Then fColor = vbBlack: Err.Clear
Dim fWeight As Long:    fWeight = IIf(IsMissing(FontWeight), ctl.FontWeight, FontWeight): If Err Then fWeight = 0: Err.Clear
Dim fItalic As Long:    fItalic = IIf(IsMissing(FontItalic), ctl.FontItalic, FontItalic): If Err Then fItalic = False: Err.Clear
Dim fUnderline As Long: fUnderline = IIf(IsMissing(FontUnderline), ctl.FontUnderline, FontUnderline): If Err Then fUnderline = False: Err.Clear
'Dim fStrikeOut As Long: fStrikeOut = IIf(IsMissing(FontStrikeOut), ctl.FontStrikeOut, FontStrikeOut): If Err Then fStrikeOut = False: Err.Clear
On Error GoTo HandleError
    'FontSize = -(FontSize * PT / TwipsPerPixels)
    'fSize = -Int(fSize * GetDeviceCaps(tDC, LOGPIXELSY) / PointsPerInch)
    fSize = -MulDiv(fSize, GetDeviceCaps(tDC, LOGPIXELSY), PointsPerInch)
    hFont = CreateFont(fSize, 0, 0, 0, _
        fWeight, fItalic, fUnderline, 0, _
        RUSSIAN_CHARSET, 0, 0, ANTIALIASED_QUALITY, 0, fName)  ' PROOF_QUALITY | CLEARTYPE_QUALITY | ANTIALIASED_QUALITY
    If hdc = 0 Then ReleaseDC 0, tDC
HandleExit:  p_HFontByControl = hFont: Exit Function
HandleError: hFont = False: Err.Clear: Resume HandleExit
End Function
Private Function p_IsEvalutable(ByRef Expr As String, Optional ByRef Value) As Boolean
' проверяет необходимость вычисления выражения, в Value возвращает результат вычисления
'-------------------------
    On Error GoTo HandleError
    If IsNumeric(Expr) Then p_IsEvalutable = False: Exit Function ' нет необходимости вычислять числа
#If APPTYPE = 0 Then ' Access
    Value = Application.Eval(Expr)
#ElseIf APPTYPE = 1 Then ' Excel
    Value = Application.Evaluate(Expr)
#Else
    Err.Raise 2438
#End If
' для возможности последующего вычисления выражений с десятичными дробями
' лучше получить региональные настройки
    If IsNumeric(Value) Then
'Dim cPosDelim As String * 1: cPosDelim = p_GetLocaleInfo(LOCALE_STHOUSAND)  ' Chr(160) - разделитель разрядов целой части
'        Value = Replace(Value, cPosDelim, vbNullString)
Dim cDecDelim As String * 1: cDecDelim = "," ' p_GetLocaleInfo(LOCALE_SDECIMAL)   ' Chr(44)  - разделитель целой/дробной части десятичной дроби
        Value = Replace(Value, cDecDelim, ".")
    End If
HandleExit:  p_IsEvalutable = True:  Exit Function
HandleError: p_IsEvalutable = False: Err.Clear
End Function
Private Function p_IsExist(Key As String, Coll As Collection, Optional ByRef Value) As Boolean
' проверяет наличие элемента в коллекции
'-------------------------
    On Error GoTo HandleError
    Value = Coll(Key)
HandleExit:  p_IsExist = True:  Exit Function
HandleError: p_IsExist = False: Err.Clear
End Function
#If APPTYPE = 1 Then ' для Excel нужна замена Nz
Private Function Nz(p1, Optional p2) As Variant
    Select Case True
    Case Not IsNull(p1): Nz = p1
    Case IsMissing(p2):  Nz = Empty
    Case Else: Nz = p2
    End Select
' Nz(Null) return Empty in MS Access, so the following Excel vba matches MS Access perfectly.

' to test it open vba immediate window and type    ?(nz(null) = 0) & " " & (nz(null) = "")
' You will get True True
End Function
#End If


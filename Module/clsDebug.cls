VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'=========================
Private Const c_strModule As String = "clsDebug"
'=========================
' Описание      : Класс для отладки, обработки ошибок и замера производительности
' Версия        : 1.6.2.448264194
' Дата          : 22.09.2022 10:03:56
' Автор         : Кашкин Р.В. (KashRus@gmail.com)
' Примечание    : идея счетчика производительности взята кажется отсюда: _
'               : http://www.xbeat.net/vbspeed/download/CTiming.zip _
'               : http://www.xtremevbtalk.com/general/49211-timing-script.html
' v.1.6.2       : 27.12.2018 - добавлена возможность вывода в дате события милисекунд (экспериментально)
'=========================
' Основные свойства и методы модуля:
'-------------------------
' Error     - выводит информацию об ошибке
' Message   - выводит отладочную информацию
' Counter   - выводит результаты замера производительности
'=========================
#Const DEBUGGING = True
#Const SHOWMSECS = True                 ' выводить в протоколе милисекунды "FFF"
Const c_strLogFileName = "Debug.log"    ' имя файла протокола
Const c_strLogTableName = "SysLog"      ' имя таблицы протокола

#If SHOWMSECS Then
Const c_strDateTimeFormat = "YYYY-MM-DD HH:NN:SS,FFF"   ' шаблон формата вывода времени события
Const c_strReportFormat = "YYYYMMDDHHNNSSFFF"           ' шаблон имени файла отчета об ошибке/сообщения
#Else
Const c_strDateTimeFormat = "YYYY-MM-DD HH:NN:SS"
Const c_strReportFormat = "YYYYMMDDHHNNSS"
#End If

Private m_dbgHandleType As Byte         ' флаг типа обработки событий отладчика
Public Enum dbgHandleType
' описание типа обработки события/ошибки
    dbgDefault = 0
    dbgShowMessage = 1      ' показать сообщение об ошибке
    dbgCreateReport = 2     ' создать файл отчета об ошибке
    dbgWriteToImmediate = 4 ' вывести сообщение в Immediate
    dbgWriteToLogFile = 8   ' записать в лог-файл
    dbgWriteToLogTable = 16 ' записать в лог-таблицу
    dbgSendToSupport = 128  ' отправить сообщение в техподдержку
End Enum
' флаг обработки события по умолчанию
#If DEBUGGING Then
Const c_flgHandleDefault = dbgWriteToImmediate + dbgWriteToLogFile + dbgCreateReport
#Else
Const c_flgHandleDefault = dbgWriteToImmediate
#End If
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Const STANDARD_RIGHTS_READ = &H20000
Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const SYNCHRONIZE = &H100000
Private Const MaxLen = 256
Private Const ERROR_SUCCESS = &H0&
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const REG_SZ = 1
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
' почта
Private Const RET_OK As Long = 0

Private Const RET_FAIL As Long = vbObjectError - 503
' версия ОС
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128   ' Maintenance string for PSS usage.
End Type


Private Const MAX_PATH As Integer = 255

Private m_strAppName As String
Private m_strAppVersion As String
Private m_strAppVersionDate As Date
Private m_varAppAuthor As String
Private m_lngAppUser As Long
Private m_strSupportEmail As String
Private m_strLogPath As String

Private LogFile As Byte ', ErrFile As Byte

Const c_strLogValueDelim = vbTab
Const c_strErrorMessage = "Error:"
Const c_strDebugMessage = "Debug:"

Public Enum eCounterAction
    CounterStart = 1
    CounterStep = 0
    CounterStop = -1
End Enum
Private Type SYSTEMTIME
    wYear          As Integer
    wMonth         As Integer
    wDayOfWeek     As Integer
    wDay           As Integer
    wHour          As Integer
    wMinute        As Integer
    wSecond        As Integer
    wMilliseconds  As Integer
End Type


Private Type LARGE_INTEGER ' структура для хранения 64-разрядного числа
' LARGE_INTEGER is faster than Currency type Currency requires CPU to execute slow floating-point instructions
    LowPart As Long
    HighPart As Long
End Type
Private Type tCounter
    Num As Long         ' номер шага
    Start As Currency   ' время начала
    Step As Currency    ' время окончания
End Type
Private Type tEvent
    EventType As Long
    EventDate As Double 'Date 'String
    Number As Long
    Description As String
    Source As String
    LineNum As Long
    Procedure As String
    Comment As String
    Target As String
End Type
Const c_strCountPrefix = "c"
Private m_colCounters As Collection
Private m_curFreq As Currency
Private m_CounterEnabled As Boolean
 
Private sysLocalTime As SYSTEMTIME
 
'--------------------------------------------------------------------------------
' KERNEL32
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystem As SYSTEMTIME)
    Private Declare PtrSafe Function GetTempDir Lib "kernel32" Alias "GetTmpPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare PtrSafe Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
    
    Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
    Private Declare PtrSafe Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, ByRef lpcbClass As Long, ByVal lpReserved As Long, ByRef lpcSubKeys As Long, ByRef lpcbMaxSubKeyLen As Long, ByRef lpcbMaxClassLen As Long, ByRef lpcValues As Long, ByRef lpcbMaxValueNameLen As Long, ByRef lpcbMaxValueLen As Long, ByRef lpcbSecurityDescriptor As Long, ByRef lpftLastWriteTime As FILETIME) As Long
    Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, lpData As Any, ByRef lpcbData As Long) As Long
    Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
#Else                   '<WIN32>
    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Sub GetLocalTime Lib "kernel32" (lpSystem As SYSTEMTIME)
    Private Declare Function GetTempDir Lib "kernel32" Alias "GetTmpPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
    
    Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
    Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, ByRef lpcbClass As Long, ByVal lpReserved As Long, ByRef lpcSubKeys As Long, ByRef lpcbMaxSubKeyLen As Long, ByRef lpcbMaxClassLen As Long, ByRef lpcValues As Long, ByRef lpcbMaxValueNameLen As Long, ByRef lpcbMaxValueLen As Long, ByRef lpcbSecurityDescriptor As Long, ByRef lpftLastWriteTime As FILETIME) As Long
    Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, lpData As Any, ByRef lpcbData As Long) As Long
    Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
#End If                 '<WIN32>

Private Sub Class_Initialize()
    On Error Resume Next
' запоминаем свойства приложения для отчетов
    With CodeProject.Properties ' все это оч.криво
        m_strAppName = !Application
        m_strAppVersion = !Version
        m_strAppVersionDate = !VersionDate
        m_varAppAuthor = !Author
        m_lngAppUser = !UserCode
        m_strSupportEmail = !Support
        m_strLogPath = CodeProject.path & "\" '& "\" & !LogPath & "\"
    End With
' определяем тип обработки событий отладки по-умолчанию
    Me.HandleType = dbgDefault
' подготавливаем счетчик производительности
    Dim liFrequency As LARGE_INTEGER
    m_CounterEnabled = QueryPerformanceFrequency(liFrequency) <> 0
    If m_CounterEnabled Then m_curFreq = p_LargeIntToCurrency(liFrequency)
    Set m_colCounters = New Collection
End Sub
Private Sub Class_Terminate()
'    ' закрываем все счетчики
'    With m_colCounters: Do Until .Count = 0: Me.Counter , .Count, CounterStop: Loop: End With
    Set m_colCounters = Nothing: Close #LogFile
End Sub
Public Sub Error( _
    ByVal Number As Long, _
    ByVal Description As String, _
    Optional ByVal Source As String, Optional Procedure As String, Optional LineNum As Long, _
    Optional Comment As String, _
    Optional OutputType As dbgHandleType)
' обработчик ошибок/событий приложения
Const c_strProcedure = "Error"
' Number - номер ошибки
' Description - описание ошибки
' Source,Procedure,LineNum - информация о месте вызова
' Comment - комментарий
' OutputType - тип вывода
Dim strMessage As String
Dim newEvent As tEvent
'' формируем текст сообщения
'    'Set objModule = VBE.ActiveCodePane.CodeModule
'    strExtendedErrInfo = p_AppFunc("rGetExtendedErrInfo")
'    strProductVersion = p_AppFunc("rProductVersionInfoGet")
'    p_AppRun "rWordSpellingCheckReset"
    If IsMissing(OutputType) Then
        OutputType = dbgWriteToImmediate + dbgWriteToLogFile + dbgCreateReport + dbgSendToSupport
    ElseIf OutputType = dbgDefault Then
        OutputType = m_dbgHandleType
    End If
    With newEvent
        .EventType = 100: .EventDate = GetDateTime() 'Now()
        .Number = Number
        .Description = Description
        .Source = Source: .Procedure = Procedure: .LineNum = LineNum
        .Comment = Comment
    End With
    p_OutputEvent newEvent, OutputType
HandleExit:
    Exit Sub
End Sub
Public Sub Message( _
    ByVal Description As String, _
    Optional Source As String, Optional Procedure As String, Optional LineNum As Long, _
    Optional Comment As String, _
    Optional OutputType As dbgHandleType)
' обработчик событий отладки
Const c_strProcedure = "Message"
' Description - сообщение отладки
' Source,Procedure,LineNum - информация о месте вызова
' Comment - комментарий
' OutputType - тип вывода информации
Dim newEvent As tEvent
Dim Result As Boolean

    Result = False
    On Error GoTo HandleError
    If IsMissing(OutputType) Then
        OutputType = dbgWriteToImmediate + dbgWriteToLogFile '+ dbgCreateReport
    ElseIf OutputType = dbgDefault Then
        OutputType = m_dbgHandleType
    End If
    With newEvent
        .EventType = 200: .EventDate = GetDateTime() 'Now()
        .Description = Description
        .Source = Source: .Procedure = Procedure: .LineNum = LineNum
        .Comment = Comment
    End With
    p_OutputEvent newEvent, OutputType
    
    Result = True
HandleExit:
'    Message = Result
    Exit Sub
HandleError:
    Result = False: Err.Clear: Resume HandleExit
End Sub
Public Function Counter( _
    Optional CounterAction As eCounterAction = CounterStep, _
    Optional ByVal CounterName As String, _
    Optional Description As String, _
    Optional Source As String, Optional Procedure As String, Optional LineNum As Long, _
    Optional Comment As String, _
    Optional OutputType As dbgHandleType _
    ) As Boolean
' управление счётчиком производительности
Const c_strProcedure = "Counter"
' CounterAction - действие счетчика CounterStart/CounterStep/CounterStop
' CounterName - имя счетчика. нужно если запускается несколько счетчиков.
'   !!! Внимание корректность имени не проверяется. должно подходить для использования в качестве имени элемента коллекции
' Description - сообщение счетчика. перед сообщением выводится информация счетчика
' Source,Procedure,LineNum - информация о месте вызова
' Comment - комментарий
' OutputType - тип вывода
Dim strName As String
Dim curNext As Currency, curTime As Double 'Date
Dim curCount As tCounter, strCount As String
Static lngCount As Long
Dim strText As String, strTime As String
Dim newEvent As tEvent
Dim Result As Boolean

    Result = False
    On Error GoTo HandleError
    If IsMissing(OutputType) Then
        OutputType = dbgWriteToImmediate + dbgWriteToLogFile '+ dbgCreateReport
    ElseIf OutputType = dbgDefault Then
        OutputType = m_dbgHandleType
    End If
    curTime = GetDateTime() ' Now()
    If Len(CounterName) = 0 Then CounterName = "Main"
    If Not m_CounterEnabled Then Err.Raise vbObjectError + 512
    strText = "<" & CounterName & ">"
    strName = c_strCountPrefix & CounterName
    curNext = p_GetPerfCounter
    Select Case CounterAction
    Case CounterStart
'Запуск счётчика производительности
        lngCount = Len(curCount)
        With curCount
            .Num = 0: .Start = curNext: .Step = curNext
            strText = strText & c_strLogValueDelim & "Start"
        End With
        If p_CounterExists(strName) Then m_colCounters.Remove (strName)
    Case CounterStep
'Запрос времени этапа измерения счётчика производительности
        Result = p_CounterExists(strName)
        If Not Result Then Result = Me.Counter(CounterStart, CounterName:=CounterName) ' счетчик не запущен - запускаем
        If Not Result Then Err.Raise vbObjectError + 513 ' не возможно запустить счетчик с именем
        CopyMemory ByVal VarPtr(curCount), ByVal StrPtr(m_colCounters(strName)), ByVal lngCount
        With curCount
            .Num = .Num + 1: strTime = p_GetPerfomance(.Step, curNext)
            .Step = curNext: strText = strText & c_strLogValueDelim & "Step"
            strText = strText & c_strLogValueDelim & "#" & Format$(.Num, String(6, "0"))
            If Len(strTime) > 0 Then strText = strText & c_strLogValueDelim & "Time: " & strTime & " sec."
        End With
        m_colCounters.Remove (strName)
    Case CounterStop
'Останов счётчика производительности
        If Not p_CounterExists(strName) Then Err.Raise vbObjectError + 514 ' не возможно остановить не запущенный счетчик
        CopyMemory ByVal VarPtr(curCount), ByVal StrPtr(m_colCounters(strName)), ByVal lngCount
        With curCount
            strTime = p_GetPerfomance(.Start, curNext)
            strText = strText & c_strLogValueDelim & "Stop"
            If .Num Then strText = strText & c_strLogValueDelim & "Steps passed: " & Format$(.Num, String(6, "0"))
            If Len(strTime) > 0 Then strText = strText & c_strLogValueDelim & "Elapsed time: " & strTime & " sec."
        End With
        m_colCounters.Remove (strName)
        GoTo HandleOutput
    End Select
'Сохраняем результаты
    strCount = Space(lngCount)
    CopyMemory ByVal StrPtr(strCount), ByVal VarPtr(curCount), ByVal lngCount
    m_colCounters.Add strCount, strName
HandleOutput:
    If Len(Description) > 0 Then strText = strText & c_strLogValueDelim & Description
    With newEvent
        .EventType = 300: .EventDate = curTime
        .Source = Source: .Procedure = Procedure: .LineNum = LineNum
        .Description = strText 'Description
        .Comment = Comment
    End With
    p_OutputEvent newEvent, OutputType
    Result = True
HandleExit:
    Counter = Result
    Exit Function
HandleError:
    Result = False
    Select Case Err.Number
    Case vbObjectError + 512: strText = "Can't stop Counter because it wasn't started!": Err.Clear: Resume HandleOutput
    Case vbObjectError + 513: strText = "Can't stop Counter because it wasn't started!": Err.Clear: Resume HandleOutput
    Case vbObjectError + 514: strText = "QueryPerformanceCounter not supported!":  Err.Clear: Resume HandleOutput
    Case Else: Err.Clear: Resume HandleExit
    End Select
End Function
Public Property Get Support() As String: Support = m_strSupportEmail: End Property
Public Function GetDateTime() As Double
Dim Result As Double
    On Error GoTo HandleError
#If Not SHOWMSECS Then
    Result = Now
#Else
    GetLocalTime sysLocalTime
    Result = Date
    With sysLocalTime ' системное время
        Result = Result + .wHour / 24                 ' HH  24
        Result = Result + .wMinute / 1440             ' NN  24*60
        Result = Result + .wSecond / 86400            ' SS  24*60*60
        Result = Result + .wMilliseconds / 86400000   ' FFF 24*60*60*1000
    End With
#End If
HandleExit:
    GetDateTime = Result
    Exit Function
HandleError:
    Err.Clear: Resume HandleExit
End Function
Public Function FormatDateTime(DateTime As Double, FormatString As String) As String
' форматирует дату по шаблону
Dim ad As Date
Const cstrYear = "YYYY" ' YYYY - год
Const cstrMonth = "MM"  ' MM - месяц
Const cstrDay = "DD"    ' DD - день
Const cstrHour = "HH"   ' HH - часы
Const cstrMin = "NN"    ' NN - минуты
Const cstrSec = "SS"    ' SS - секунды
Const cstrMSec = "FFF"  ' FFF - милисекунды
Dim Result As String
    On Error GoTo HandleError
    If InStr(FormatString, cstrMSec) <= 0 Then ad = DateTime: GoTo Handleresult
Dim dt As SYSTEMTIME
Dim tmp As Double
    tmp = DateTime
    ad = DateValue(CDate(tmp))  ' берём отдельно дату
    tmp = tmp - Int(tmp)        ' отдельно время
    With dt
' извлекаем компоненты времени - руками чтобы избежать округления милисекунд до секунд
        .wHour = Int(tmp * 24):         tmp = tmp - .wHour / 24
        .wMinute = Int(tmp * 1440):     tmp = tmp - .wMinute / 1440
        .wSecond = Int(tmp * 86400):    tmp = tmp - .wSecond / 86400
        .wMilliseconds = Int(tmp * 86400000) ':    tmp = tmp - .wSecond / 86400000
' меняем шаблон
        FormatString = VBA.Replace$(FormatString, cstrMSec, VBA.Format$(.wMilliseconds, "000"))
' собираем
        ad = ad + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
Handleresult:
    Result = VBA.Format$(ad, FormatString)
HandleExit:
    FormatDateTime = Result
    Exit Function
HandleError:
    Err.Clear: Resume HandleExit
End Function
Public Property Get HandleType() As dbgHandleType: HandleType = m_dbgHandleType: End Property
Public Property Let HandleType(ByVal NewValue As dbgHandleType)
    Select Case NewValue
    Case dbgDefault:    m_dbgHandleType = c_flgHandleDefault
    Case Else:          m_dbgHandleType = NewValue
    End Select
On Error Resume Next
'' записать в лог-файл
'    If m_dbgHandleType And dbgWriteToLogFile = dbgWriteToLogFile Then p_InitLogFile Else Close #LogFile
' записать в лог-таблицу
'    If m_dbgHandleType And dbgWriteToLogTable = dbgWriteToLogTable Then p_InitLogTable
End Property

Private Sub p_OutputEvent( _
    EventInfo As tEvent, _
    Optional ByVal OutputType As Byte)
Const c_strProcedure = "p_OutputEvent"
Dim strMessage As String, lngMsgBoxResult As Long
Dim strEventType As String ', strEventType2 As String
Dim strDateTime As String
Dim strSource As String
Dim Result As String
    
    On Error GoTo HandleError
    If OutputType = dbgDefault Then OutputType = m_dbgHandleType
    
    With EventInfo
        Select Case .EventType
        Case 100: strEventType = "ERROR" '"Ошибка "
        Case 200: strEventType = "DEBUG" '"Отладка"
        Case 300: strEventType = "COUNT" '"Счетчик"
        End Select
        strMessage = strEventType & ":"
        
        If Len(.Procedure) > 0 Then
            strSource = .Procedure
            If .LineNum <> 0 Then strSource = strSource & " #" & .LineNum
        End If
        If Len(.Source) > 0 And Len(strSource) > 0 Then
            strSource = strSource & " (" & .Source & ")"
        ElseIf Len(.Source) > 0 Then
            strSource = .Source
        End If
        
        If .Number <> 0 Then strMessage = strMessage & " #" & .Number
        If Len(.Description) > 0 Then strMessage = strMessage & " " & .Description
        If Len(strSource) > 0 Then strMessage = strMessage & " in " & strSource
        If Len(.Comment) > 0 Then strMessage = strMessage & " " & .Comment
        strDateTime = FormatDateTime(.EventDate, c_strDateTimeFormat)
    ' выводим в окно Immediate
        If (OutputType And dbgWriteToImmediate) = dbgWriteToImmediate Then Debug.Print strMessage
    ' выводим в Log файл
        If (OutputType And dbgWriteToLogFile) = dbgWriteToLogFile Then
'            If LogFile = 0 Then p_InitLogFile
' проверяем существование пути m_strLogPath если нет - создаем
            If Len(dir(m_strLogPath, vbDirectory)) = 0 Then MkDir m_strLogPath
            LogFile = FreeFile
            Open m_strLogPath & c_strLogFileName For Append Lock Write As #LogFile
            Print #LogFile, strDateTime & c_strLogValueDelim & strMessage
            Close #LogFile
        End If
    ' выводим в Log таблицу
        If (OutputType And dbgWriteToLogTable) = dbgWriteToLogTable Then p_WriteToTable EventInfo
' только для ошибок
    If .EventType <> 100 Then GoTo HandleExit
    Dim strUserName As String, strCompName As String
        strEventType = "Ошибка "
        strUserName = m_lngAppUser  '& " (" & DLookup(c_strCName, "SysUsers", c_strKey & sqlEqual & m_lngAppUser) & ")"
        strCompName = p_GetNetUserComp
        strMessage = strEventType & vbTab & vbTab & ": " & "#" & .Number & " " & .Description & vbCrLf
        strMessage = strMessage & "Источник" & vbTab & ": " & strSource & vbCrLf
        strMessage = strMessage & "Комментарий" & vbTab & ": " & .Comment
    ' выводим сообщение
        If (OutputType And dbgShowMessage) = dbgShowMessage Then MsgBox strMessage, vbCritical
    ' при необходимости создаем отчет и отправляем письмо
        If Not ((OutputType And dbgCreateReport) = dbgCreateReport _
            Or (OutputType And dbgSendToSupport) = dbgSendToSupport) Then GoTo HandleExit
        ' cформироваnm отчет об ошибке
        'If p_CheckForOLEMsg Then
        strMessage = strMessage & vbCrLf & vbCrLf & "Хотите сообщить разработчику?"
        lngMsgBoxResult = MsgBox(strMessage, vbYesNo + vbCritical + vbDefaultButton2, strEventType)
        
        strMessage = vbNullString
        strMessage = strMessage & String(60, "-") & vbCrLf
        strMessage = strMessage & "Дата отчета" & vbTab & ": " & strDateTime & vbCrLf
        strMessage = strMessage & strEventType & vbTab & vbTab & ": " & "#" & .Number & " " & .Description & vbCrLf
        strMessage = strMessage & "Источник" & vbTab & ": " & strSource & vbCrLf
        strMessage = strMessage & "Комментарий" & vbTab & ": " & .Comment & vbCrLf
        strMessage = strMessage & String(60, "-") & vbCrLf
        strMessage = strMessage & "Действия" & vbTab & ": " & vbCrLf & vbCrLf & vbCrLf
        strMessage = strMessage & "Результат" & vbTab & ": " & vbCrLf & vbCrLf & vbCrLf
        strMessage = strMessage & String(60, "-") & vbCrLf
        strMessage = strMessage & "Приложение" & vbTab & ": " & m_strAppName & " v." & m_strAppVersion & " от " & m_strAppVersionDate & vbCrLf
'        strMessage = strMessage & "Пользователь" & vbTab & ": " & "#" & strUserName & " (" & DLookup(c_strCName, "SysUsers", c_strKey & sqlEqual & strUserName) & ")" & vbCrLf
        strMessage = strMessage & "Компьютер" & vbTab & ": " & strCompName & vbCrLf
        strMessage = strMessage & "Система" & vbTab & ": " & p_GetWinVersion & ", " & p_GetAccVersion & vbCrLf
        strMessage = strMessage & p_GetRefVersions & vbCrLf
        strMessage = strMessage & String(60, "-") & vbCrLf
        strMessage = strMessage & "Для ускорения поддержки пожалуйста пришлите заархивированные" & vbCrLf
        strMessage = strMessage & "базу данных, в которой возникла ошибка, и файлы протоколов."

        Select Case lngMsgBoxResult
        Case vbYes  ' послать сообщение
            p_SendMessage m_strSupportEmail, m_strAppName, strMessage ' отправить письмо разработчику
        Case vbNo   ' сформировать отчёт
            p_CreateReport strMessage, m_strLogPath & "~" & FormatDateTime(.EventDate, c_strReportFormat) & ".txt"
        End Select
        'End If
    End With
    
HandleExit:
    Exit Sub
HandleError:
    Result = False
    Select Case Err.Number
    Case 52:    Err.Clear: p_InitLogFile ' Bad file name or number
    Case 55:    Err.Clear: p_InitLogFile ' File already open
    Case Else:  Err.Clear
    End Select
    Resume HandleExit
End Sub
Private Function p_SendMessage( _
    ByVal AddrTo As String, _
    ByVal Subject As String, _
    ByVal MessageText As String, _
    Optional ByVal AddrCC As String = "", _
    Optional ByVal AddrBCC As String = "", _
    Optional ByVal TruncUntilLastCrLf As Boolean = False) As Long
Dim strMsg As String
Dim lngRet As Long
Dim lngIdx As Long
    strMsg = "mailto:" & AddrTo & "?"
    If Len(AddrCC) > 0 Then strMsg = strMsg & "CC=" & AddrCC & "&"
    If Len(AddrBCC) > 0 Then strMsg = strMsg & "BCC=" & AddrBCC & "&"
    Subject = Replace(Subject, vbCrLf, "%0d%0a"): Subject = Replace(Subject, " ", "%20")
    strMsg = strMsg & "Subject=" & Subject & "&"
    MessageText = Replace(MessageText, vbCrLf, "%0d%0a"): MessageText = Replace(MessageText, " ", "%20")
    strMsg = strMsg & "Body=" & MessageText
    If Len(strMsg) > 2000 Then
        strMsg = VBA.Left(strMsg, 2000) & "..."
        If TruncUntilLastCrLf = True Then
            For lngIdx = Len(strMsg) To 1 Step -1
                If VBA.Mid(strMsg, lngIdx, 6) = "%0d%0a" Then
                    strMsg = VBA.Left(strMsg, lngIdx + 5) & "..."
                    Exit For
                End If
            Next lngIdx
        End If
    End If
    lngRet = ShellExecute(0&, vbNullString, strMsg, vbNullString, vbNullString, vbNormalFocus)
    If lngRet >= 42 Then p_SendMessage = RET_OK Else p_SendMessage = RET_FAIL
HandleExit:
    Exit Function
HandleErr:
    p_SendMessage = Err.Number
    Resume HandleExit
End Function
Private Function p_InitLogTable() As Boolean
Const c_strProcedure = "p_InitLogTable"
Dim dbs As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field
Dim Result As Boolean
    
    Result = False
    On Error GoTo HandleError
    Set dbs = CurrentDb
    Set tdf = dbs.CreateTableDef(c_strLogTableName)
    Set fld = tdf.CreateField("EventDate", dbDouble): tdf.Fields.Append fld
    Set fld = tdf.CreateField("EventType", dbLong): tdf.Fields.Append fld
    Set fld = tdf.CreateField("Number", dbLong): tdf.Fields.Append fld
    Set fld = tdf.CreateField("Description", dbText, 255): tdf.Fields.Append fld
    Set fld = tdf.CreateField("Source", dbText, 50): tdf.Fields.Append fld
    'Set fld = tdf.CreateField("Target", dbText, 50): tdf.Fields.Append fld
    Set fld = tdf.CreateField("Comment", dbText, 255): tdf.Fields.Append fld
    Set fld = tdf.CreateField("CompName", dbText, 50): tdf.Fields.Append fld
    Set fld = tdf.CreateField("UserName", dbText, 50): tdf.Fields.Append fld
    dbs.TableDefs.Append tdf
    Result = True
HandleExit:
    p_InitLogTable = Result
    Exit Function
HandleError:
    Result = False
    Err.Clear
    Resume HandleExit
End Function
Private Function p_InitLogFile()
    Close '#LogFile
' проверяем существование пути m_strLogPath если нет - создаем
    If Len(dir(m_strLogPath, vbDirectory)) = 0 Then MkDir m_strLogPath
    LogFile = FreeFile
    Open m_strLogPath & c_strLogFileName For Append Lock Write As #LogFile
End Function
'Public Property Let Support(rData As String)
'    If IsEMail(rData) Then
'        CodeProject.Properties!Support = rData
'        m_strSupportEmail = rData
'    End If
'End Property

Private Function p_GetNetUserComp() As String
' Возвращает сетевое имя компьютера _
  или пустую строку, если комьютер не имеет сетевого имени
Dim Pos As Long
Dim cBuffer$: cBuffer = Space(64)
Dim nSize&: nSize = 64
Dim Result As String
    Call GetComputerName(cBuffer, nSize)
    Pos = InStr(1, cBuffer, vbNullChar)
    Select Case Pos
        Case Is > 1: Result = Left$(cBuffer, Pos - 1)
        Case 0:      Result = Trim$(cBuffer)
        Case 1:      Result = vbNullString
    End Select
    p_GetNetUserComp = Result
End Function
Private Function p_GetNetUserName() As String
' Возвращает имя пользователя, под которым он зарегистрировался в сети _
  или пустую строку, если регистрации не было
Dim Pos As Long
Dim cBuffer$: cBuffer = Space(64)
Dim nSize&: nSize = 64
Dim Result As String
    Call GetUserName(cBuffer, nSize)
    Pos = InStr(1, cBuffer, vbNullChar)
    Select Case Pos
        Case Is > 1: Result = Left$(cBuffer, Pos - 1)
        Case 0:      Result = Trim$(cBuffer)
        Case 1:      Result = vbNullString
    End Select
    p_GetNetUserName = Result
End Function
Private Function p_GetWinUserName() As String
Const c_strProcedure = "GetWinUserName"
Dim sBuffer As String, lSize As Long
Dim Result As String
    On Error GoTo HandleError
    sBuffer = Space$(255): lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    Result = vbNullString
    If lSize > 0 Then Result = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    'If lSize > 0 Then Result = Left$(sBuffer, lSize)
HandleExit:
    p_GetWinUserName = Result
    Exit Function
HandleError:
    Result = vbNullString
    Err.Clear: Resume HandleExit
End Function
Private Function p_AppRun(ByVal vstrProcName As String): Application.Run vstrProcName: End Function
Private Function p_AppFunc(ByVal vstrProcName As String) As String: p_AppFunc = Application.Run(vstrProcName): End Function
Private Function p_CheckForOLEmsg() As Boolean: Dim bOK As Boolean: p_CheckForOLEmsg = (p_GetRegistryValue(&H80000002, "SOFTWARE\Microsoft\Windows Messaging Subsystem\", "OleMessaging", bOK) = "1"): End Function
Private Function p_GetTempDir() As String
'Возвращает путь к Temp папке
    Dim strTemp As String: strTemp = String$(MAX_PATH, 0)
    Dim lngRtn As Long: lngRtn = GetTempDir(MAX_PATH, strTemp)
    If lngRtn <> 0 Then p_GetTempDir = Left$(strTemp, lngRtn) Else p_GetTempDir = ""
End Function
Private Function p_GetRegistryValue(ByVal lngKeyToGet As Long, _
    ByVal strKeyName As String, _
    ByVal strValueName As String, _
    bOK As Boolean) _
    As String

Dim lnghKey As Long
Dim strClassName As String
Dim lngClassLen As Long
Dim lngReserved As Long
Dim lngSubKeys As Long
Dim lngMaxSubKeyLen As Long
Dim lngMaxClassLen As Long
Dim lngValues As Long
Dim lngMaxValueNameLen As Long
Dim lngMaxValueLen As Long
Dim lngSecurity As Long
Dim ftLastWrite As FILETIME
Dim lngType As Long
Dim lngData As Long
Dim lngTmp As Long
Dim strRet As String
Dim varRet As Variant
Dim lngRet As Long
'    On Error GoTo RegistryValueGet_Err
    ' открываем ключ
    lngTmp = RegOpenKeyEx(lngKeyToGet, strKeyName, 0&, KEY_READ, lnghKey)

    If Not (lngTmp = ERROR_SUCCESS) Then Err.Raise lngTmp + vbObjectError

    lngReserved = 0&
    strClassName = String$(MaxLen, 0):  lngClassLen = MaxLen
    ' берем граничные значения
    lngTmp = RegQueryInfoKey(lnghKey, strClassName, _
        lngClassLen, lngReserved, lngSubKeys, lngMaxSubKeyLen, _
        lngMaxClassLen, lngValues, lngMaxValueNameLen, _
        lngMaxValueLen, lngSecurity, ftLastWrite)

    If Not (lngTmp = ERROR_SUCCESS) Then Err.Raise lngTmp + vbObjectError

    ' получаем значение ключа
    strRet = String$(MaxLen - 1, 0)
    lngTmp = RegQueryValueEx(lnghKey, strValueName, _
        lngReserved, lngType, ByVal strRet, lngData)
    Select Case lngType
        Case REG_SZ
            lngTmp = RegQueryValueEx(lnghKey, strValueName, lngReserved, lngType, ByVal strRet, lngData)
            varRet = VBA.Left(strRet, lngData - 1)
        Case REG_DWORD
            lngTmp = RegQueryValueEx(lnghKey, strValueName, lngReserved, lngType, lngRet, lngData)
            varRet = lngRet
        Case REG_BINARY
            lngTmp = RegQueryValueEx(lnghKey, strValueName, lngReserved, lngType, ByVal strRet, lngData)
            varRet = VBA.Left(strRet, lngData)
    End Select

    If Not (lngTmp = ERROR_SUCCESS) Then
'        Err.Raise lngTmp + vbObjectError
    Else
        bOK = True
    End If
    
RegistryValueGet_Exit:
    p_GetRegistryValue = varRet
    lngTmp = RegCloseKey(lnghKey)
    Exit Function
RegistryValueGet_Err:
    varRet = vbNullString
    Resume RegistryValueGet_Exit
End Function
Private Function p_GetWinVersion( _
    Optional lngPlatformId As Long, _
    Optional lngMajorVersion As Long, _
    Optional lngMinorVersion As Long, _
    Optional lngBuildNumber As Long _
    ) As String
' получаем основную системную информация
'                   Win95     Win98     WinNT 3.51     WinNT 4.0   Win2000 Win XP   VISTA
'                  -----------------------------------------------------------------------
'dwPlatFormID         1         1            2            2        2        2       2
'
'dwMajorVersion       4         4            3            4        5        5       6
'
'dwMinorVersion       0        10           51            0        0        1       0
'WIN32s=WIN32s or WinNT <3.51
Dim v As OSVERSIONINFO, RetVal As Long
Dim strWinVersion As String, BuildVersion As String
Const cDiv As String = "/"
Dim Result As String
    v.dwOSVersionInfoSize = Len(v)
    RetVal = GetVersionEx(v)
    
    lngPlatformId = v.dwPlatformId
    lngMajorVersion = v.dwMajorVersion
    lngMinorVersion = v.dwMinorVersion
    lngBuildNumber = v.dwBuildNumber
    
    strWinVersion = v.dwMajorVersion & "." & v.dwMinorVersion
    BuildVersion = v.dwBuildNumber And &HFFFF& 'mu?
    
    Select Case v.dwPlatformId
        Case VER_PLATFORM_WIN32_WINDOWS
            Select Case v.dwMinorVersion
                Case 0:  Result = "Windows 95"
                Case 10: Result = "Windows 98"
            End Select
        Case VER_PLATFORM_WIN32_NT
            Select Case v.dwMajorVersion
                Case 3:  Result = "WinNT 3.51"
                Case 4:  Result = "WinNT 4"
                Case 5
                    If v.dwMinorVersion = 0 Then
                        Result = "Windows 2000"
                    Else
                        Result = "Windows XP"
                    End If
                Case 6: Result = "VISTA"
            End Select
        Case VER_PLATFORM_WIN32s
            Result = "NT < 3.51"
    End Select
HandleExit:
    p_GetWinVersion = Result
    Exit Function
HandleError:
    Result = "SysInfo not available"
    Resume HandleExit
End Function
Private Function p_GetRefVersions() As String
Const c_strProcedure = "p_GetRefVersions"
' список подключенных библиотечных ссылок с версиями
Const c_strDelim = ", " 'vbCrLf
Dim fs As Object
Dim ref As Reference
Dim strFullPath As String
Dim Result As String
'0 - GID
'1 - Имя (не обязательно)
'2 - Версия - Major
'3 - Версия - Minor
'4 - Имя файла или полный путь (не обязательно)
'5 - Полное название (не обязательно)
    Result = vbNullString
    On Error GoTo HandleNext
    Set fs = CreateObject("Scripting.FileSystemObject")
    For Each ref In References
        strFullPath = ref.FullPath
        With fs
            Result = Result & c_strDelim & .GetFileName(strFullPath) & " (v." & .GetFileVersion(strFullPath) & ")"
        End With
HandleNext:
    Next ref
    On Error GoTo HandleError
HandleExit:
    Set fs = Nothing
    p_GetRefVersions = Mid$(Result, Len(c_strDelim) + 1, Len(Result) - Len(c_strDelim) + 1)
    Exit Function
HandleError:
    'RESULT = ""
    Resume HandleExit
End Function
Private Function p_GetAccVersion() As String
Const c_strProcedure = "p_GetAccVersion"
' версия Access
Dim Result As String
    Select Case Version
    Case "2.0":  Result = "2.0"
    Case "7.0":  Result = "95"
    Case "8.0":  Result = "97"
    Case "9.0":  Result = "2000"
    Case "10.0": Result = "XP"
    Case "11.0": Result = "2003"
    Case "12.0": Result = "2007"
    Case "14.0": Result = "2010"
    Case "15.0": Result = "2013"
    Case "16.0": Result = "2016"
    Case "17.0": Result = "2019"
    Case Else:   Result = "?" & Version
    End Select
HandleExit:
    p_GetAccVersion = "Access " & Result
    Exit Function
HandleError:
    'RESULT = ""
    Resume HandleExit
End Function
Public Property Get CounterEnabled() As Boolean: CounterEnabled = m_CounterEnabled: End Property
Private Function p_CounterExists(Optional CounterName As String, Optional ByRef Counter) As Boolean
Const c_strProcedure = "p_CounterExists"
Dim Result As Boolean
    Result = False
    On Error GoTo HandleError
    Counter = m_colCounters(CounterName)
    'With Counter: Result = .Num = .Num: End With
    Result = True
HandleExit:
    p_CounterExists = Result
    Exit Function
HandleError:
    Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_GetPerfCounter() As Currency
Dim liCount As LARGE_INTEGER
Dim Result As Currency
    On Error GoTo HandleError
    If Not m_CounterEnabled Then Err.Raise vbObjectError + 512
    QueryPerformanceCounter liCount: Result = p_LargeIntToCurrency(liCount)
HandleExit:
    p_GetPerfCounter = Result
    Exit Function
HandleError:
    Result = 0: Err.Clear:  Resume HandleExit
End Function
Private Function p_GetPerfomance(curBeg As Currency, curEnd As Currency) As String: p_GetPerfomance = CStr(Round((curEnd - curBeg) / m_curFreq, 10)): End Function
Private Function p_LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
'функция для конвертирования 64-разрядного числа в тип Currency
Dim Result As Currency: CopyMemory Result, liInput, LenB(liInput): p_LargeIntToCurrency = Result * 10000
End Function
Private Function p_LargeIntToDouble(liInput As LARGE_INTEGER) As Double
'функция для конвертирования 64-разрядного числа в тип Double
Const OFFSET_4 As Double = 4294967296#  ' 2^32
Dim Low As Double, High As Double
  If liInput.LowPart < 0 Then Low = liInput.LowPart + OFFSET_4 Else Low = liInput.LowPart
  If liInput.HighPart < 0 Then High = liInput.HighPart + OFFSET_4 Else High = liInput.HighPart
  p_LargeIntToDouble = Low + High * OFFSET_4
End Function

Private Function p_WriteToTable(EventInfo As tEvent) As Boolean
Const c_strProcedure = "p_WriteToTable"
Dim dbs As DAO.Database, rst As DAO.Recordset
Dim Try As Byte
Dim Result As Boolean
    Try = 0
    Result = False
    On Error GoTo HandleError
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset(c_strLogTableName)
    rst.AddNew
    With EventInfo
    rst!EventDate = .EventDate
    rst!EventType = .EventType
    rst!Number = .Number
    If Len(.Description) > 0 Then rst!Description = .Description
    If Len(.Source) > 0 Then rst!Source = .Source
    If Len(.Comment) > 0 Then rst!Comment = .Comment
    End With
    rst!CompName = p_GetNetUserComp
    rst!UserName = m_lngAppUser '& " (" & DLookup(c_strCName, "SysUsers", c_strKey & sqlEqual & m_lngAppUser) & ")"
    rst.Update
    
    Result = True
HandleExit:
    p_WriteToTable = Result
    Exit Function
HandleError:
    Result = False
    Select Case Err.Number
    Case 3078 ' нет таблицы - пытаемся создать
        Err.Clear
        If Try = 0 Then Result = p_InitLogTable: Try = 1: If Result Then Resume 0
    End Select
    Err.Clear
    Resume HandleExit
End Function
Private Function p_CreateReport(ReportText As String, FilePath As String)
Dim intFn As Long
    If Len(FilePath) = 0 Then FilePath = m_strLogPath & "~" & FormatDateTime(Now(), c_strReportFormat) & ".txt"
    intFn = FreeFile
    Open FilePath For Output Access Write As #intFn
    Print #intFn, ReportText
    Close #intFn
    Shell "Notepad.exe """ & FilePath & """"
    Kill FilePath
End Function
Private Function p_ApiErrorText(ByVal ErrNum As Long) As String
Dim Result As String
Dim msg As String: msg = Space$(1024)
Dim nRet As Long: nRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, ErrNum, 0&, msg, Len(msg), ByVal 0&)
   If nRet Then Result = Left$(msg, nRet) Else Result = "Error (" & ErrNum & ") not defined."
   p_ApiErrorText = Result
End Function


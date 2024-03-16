Attribute VB_Name = "~App"
Option Compare Database
Option Base 1
'=========================
Private Const c_strModule As String = "~App"
'=========================
' Описание      :
' Версия        : 1.0.0.0
' Дата          : 01.07.2016 15:00:00
' Автор         : Кашкин Р.В. (KashRus@gmail.com)
' Примечание    :
'=========================
' Параметры отображения
'=========================
'-------------------------
' Цветовая схема
'-------------------------
Public Enum AppColorScheme
    appColorGrey = &H969696
    appColorLightGrey = &H767676
' Цветовая схема
    appColorDark = &H993333 '&H732A0A             ' темный цвет
    appColorBright = &HDFA000           ' яркий цвет
    appColorLight = &HE5D1C0            ' светлый цвет
' аналог 1
    appColorDark2 = &H730A1F
    appColorBright2 = &HDF3000
    appColorLight2 = &HE5C0C1
' аналог 2
    appColorDark3 = &H730A53
    appColorBright3 = &HDF003F
    appColorLight3 = &HE5C0D4
End Enum
'-------------------------
' Шрифт
'-------------------------
Public Const appFontNameDef = "Arial"
Public Const appFontSizeDef = 10
Public Const appFontSize1 = 8
Public Const appFontSize2 = 12
Public Const appFontSize3 = 18
'-------------------------

'=========================
' Параметры отображения
'=========================
'-------------------------
Private Const c_strApplication = "Картинки"
Private Const c_strAuthor As String = "Кашкин Р.В."
Private Const c_strVersion As String = "2.0.0"
Private Const c_strSupport As String = "KashRus@gmail.com"
Private Const c_strShowClock As Boolean = True
'-------------------------
'Public Const cpHash = "SHA256", HashType = eHashSHA256 '
'-------------------------
' Описания основных таблиц
'SysBookmarks, SysFields, SysLog, SysMenu, SysObjData, SysObjTypes, SysOrderTypes, SysUsers, SysVersion
Public Const c_strSysObjects = "MSysObjects"    ' Системная таблица
Public Const c_strTableVers = "SysVersion"      ' Таблиица версий приложения
Public Const c_strTableUser = "SysUsers"        ' Таблиица данных пользователей
Public Const c_strTableMenu = "SysMenu"         ' Таблиица пунктов меню
Public Const c_strTableData = "SysObjData"      ' Таблиица системных объектов
Public Const c_strTableLogs = "SysLog"          ' Таблиица протокола работы
Public Const c_strTableDate = "SysCalendar"     ' Справочник дат
'Public Const c_strTableBook = "SysBookmarks"    ' Справочник закладок отчётов
'Public Const c_strTableFlds = "SysFields"       ' Справочник соответствия полей
'Public Const c_strTableTObj = "SysObjTypes"     ' Справочник типов объектов
'Public Const c_strTableTOrd = "SysOrderTypes"   ' Справочник типов операций

'-------------------------
' имена свойств проекта
'-------------------------
' свойства приложения
Public Const c_strPropAppName = "Application" '= c_strApplication ' ="Собрания"
Public Const c_strPropAuthor = "Author"
Public Const c_strPropSupport = "Support"
Public Const c_strPropVersion = "Version"
Public Const c_strPropVerDate = "VersionDate"
Public Const c_strPropLastDate = "LastDate"
Public Const c_strPropLastUser = "LastUserName"
Public Const c_strPropFirstRun = "FirstRun"
Public Const c_strPropShowClock = "ShowClock"
' свойства путей приложения
Public Const c_strPropSrvPath = "SrvPath"
Public Const c_strPropDatPath = "DatPath"
Public Const c_strPropSecPath = "SecPath"
Public Const c_strPropLogPath = "LogPath"
Public Const c_strPropDllPath = "DllPath"
Public Const c_strPropDocPath = "DocPath"
Public Const c_strPropTmpPath = "TmpPath"
' свойства параметров экрана при разработке приложения
Public Const c_strDesignRes = "DesignRes" ' 1280x1024
Public Const c_strDesignDpi = "DesignDpi" '
Public Const c_strResDelim = "x" '
'-------------------------
' значения свойств проекта
'-------------------------
Public Const c_strLastUserName As String = "Администратор" ' Администратор по SysUsers
Public Const c_strSrvPath As String = "\"
Public Const c_strDatPath As String = "DAT"
Public Const c_strDllPath As String = "LIB"
Public Const c_strLogPath As String = "LOG"
Public Const c_strTmpPath As String = "DOT"
Public Const c_strDocPath As String = "DOC"
Public Const c_strSrcPath As String = "SRC"
Public Const c_strDbfPath As String = "DBF"
Public Const c_bolShowClock As Boolean = True
'-------------------------
' иконки
'-------------------------
Public Const c_strAppIco = "App"
Public Const c_strMenuIco = "ContextMenu"
'-------------------------
' для процедур обновляющих код приложения
'-------------------------
' маркеры начала и окончания области вставки
Private Const strBegLineMarker = "'=== BEGIN INSERT ==="
Private Const strEndLineMarker = "'==== END INSERT ===="
' чтоб не вис проц на циклах с DoEvents по умолчанию = 333
Public Const appDoEventsPause = 100

Public Const c_strTagDelim = "_"
Public Const c_strDelim = ";"
Public Const c_strInDelim = ","
' Описание инструкций SQL
Public Const sqlSelect = "SELECT ", sqlAll = "*"
Public Const sqlUpdate = "UPDATE ", sqlSet = " SET "
Public Const sqlInsert = "INSERT ", sqlInto = " INTO "
Public Const sqlTransform = "TRANSFORM ", sqlPivot = " PIVOT "
Public Const sqlDelete = "DELETE ", sqlUnion = "UNION "
Public Const sqlDrop = "DROP ", sqlTable = " TABLE ", sqlIndex = " INDEX "
Public Const sqlAs = " AS "
Public Const sqlDistinct = "DISTINCT ", sqlDistinctRow = "DISTINCTROW "
Public Const sqlFrom = " FROM ", sqlWhere = " WHERE "
Public Const sqlOrder = " ORDER BY ", sqlGroup = " GROUP BY "
Public Const sqlHaving = " HAVING ", sqlTop = " TOP ", sqlTop1 = "TOP 1 ", sqlPercent = " PERCENT "
Public Const sqlJoin = " JOIN ", sqlInner = " INNER", sqlLeft = " LEFT", sqlRight = " RIGHT", sqlOn = " ON "
Public Const sqlIdentity = "@@Identity"
Public Const sqlSelectAll = sqlSelect & sqlAll & sqlFrom
Public Const sqlSelect1st = sqlSelect & sqlTop1 & sqlAll & sqlFrom
Public Const sqlDeleteAll = sqlDelete & sqlAll & sqlFrom
Public Const sqlDropTable = sqlDrop & sqlTable, sqlDropIndex = sqlDrop & sqlIndex
Public Const sqlOR = " OR ", sqlAnd = " AND ", sqlNot = " NOT "
Public Const sqlEqual = "=", sqlGreater = ">", sqlLess = "<"
Public Const sqlGreaterOrEqual = ">=", sqlLessOrEqual = "<=", sqlNotEqual = "<>"
Public Const sqlIn = " IN ", sqlLike = " LIKE ", sqlBetween = " BETWEEN "
Public Const sqlAsc = " ASC", sqlDesc = " DESC"
Public Const sqlSimilar = "SIMILAR"  ' нестандартная - нечеткий поиск
Public Const sqlIs = " IS ", sqlNull = "NULL", sqlTrue = "True", sqlFalse = "False"
Public Const sqlIsNull = sqlIs & sqlNull, sqlIsNotNull = sqlIs & sqlNot & sqlNull
' Описания основных таблиц
'SysBookmarks, SysFields, SysLog, SysMenu, SysObjData, SysObjTypes, SysOrderTypes, SysUsers, SysVersion
'Private Const c_strSysObjects = "MSysObjects"    ' Системная таблица

' названия обрабатываемых параметров
Public Const c_strParamType = "Type"
Public Const c_strParamMode = "Mode"
Public Const c_strParamKey = "Key"
'-------------------------
' префиксы объектов Access
'-------------------------
' AccessObjectType
Public Const c_strTmpTypePref = "tmp" ' вспомогательный объект
Public Const c_strTmpTablPref = "@&%" ' временная таблица

' значение для подстановки в свойства "On[...]" объекта для перехвата его событий
Public Const c_strCustomProc = "[Event Procedure]"
Public Const c_strCmdMnuProc = "ContextMenu_Click"
' коллекция объектов для групповой обработки событий
Public frmDROP_Date_Controls As Collection ' коллекция контролов формы frmDROP_Date
' FormType
Public Const c_strMenuType = "MENU" ' меню
' еще
Public Const c_strServType = "SERV" ' служебный
Public Const c_strDropType = "DROP" ' выпадающая форма
' FormMode
Public Const c_strMainMode = "MAIN" ' для меню - основное
' для выпадающих форм FormType=c_strDropType
Public Const c_strRealMode = "Real" ' выпадающие доп.сведения
Public Const c_strCalcMode = "Calc" ' выпадающий калькулятор
Public Const c_strDateMode = "Date" ' выпадающий календарь
' для служебных форм FormType=c_strServType
Public Const c_strUserMode = "User"
Public Const c_strUChgMode = "UserChg"
Public Const c_strFloat = "Float"   ' плавающая кнопка
Public Const c_strNavBar = "NavBar" ' панель навигации по записям
Public Const c_strPrtBar = "PrtBar" ' фильтр по родительским записям
' для передачи результатов из формы авторизации

'==============================
Public Enum appErrors
' пользовательские ошибки приложения
    errAuthGrant = vbObjectError + 1000     ' сеанс начат
    errAuthError = vbObjectError + 1001     ' ошибка авторизации
    errAuthFailed = vbObjectError + 1002    ' отказ авторизации
    errAuthEnd = vbObjectError + 1009       ' сеанс завершен
    errAppNoConn = vbObjectError + 1010     ' отсутствует подключение к серверу данных
    errAppNoData = vbObjectError + 1011     ' отсутствуют данные
    errAppClose = vbObjectError + 1109      ' ошибка закрытия приложения
    errAppPathWrong = vbObjectError + 1110  ' заданы неверные пути приложения
End Enum
Public Enum enmPathType
' Типы путей приложения пользователя
    enmPathUndef = 0 ' не определен
    enmPathAll = 255 ' все пути
    enmPathSrv = 1  ' путь к серверу данных приложения
    enmPathDll = 2  ' путь к папке внешних библиотек приложения
    enmPathTmp = 3  ' путь к папке шаблонов отчетов приложения
    enmPathDoc = 4  ' путь к папке отчетов приложения
    enmPathDat = 5  ' путь к файлу данных
    enmPathSec = 6  ' путь к файлу рабочей группы
    enmPathLoc = 7  ' путь к локальной базе (интерфейсу)
    enmPathLog = 8  ' путь к папке протоколов приложения
    enmPathSrc = 9  ' путь к папке копий исходников
    enmPathDbf = 10 ' путь к папке выгрузки
End Enum
Public Enum appUserType
' Типы прав пользователя
    appUserTypeAdmin = 100
    appUserTypeUser = 200
End Enum
Public Enum appModeType
' Переключатель режимов приложения
    appModeDebug = -1                   ' режим отладки
    appModeNormal = 0                   ' рабочий режим
End Enum
Public Enum appRecState
' Признаки состояния записи (SPReal)
    appRecStateTemp = -1 'временная - не актуальная
    appRecStateReal = 0  'активная - актуальная
    appRecStateOld = 10  'старая - была изменена, есть более актуальная
    appRecStateArc = 11  'архивная - была отменена, более актуальной нет
    appRecStateDen = 91  'исключена
    appRecStateDel = 99  'помечена удаленной
End Enum
'======================
Private bolFirstRun As Boolean
'----------------------
' POINTER
'----------------------
#If VBA7 = 0 Then       'LongPtr trick by @Greedo (https://github.com/Greedquest)
Public Enum LongPtr
    [_]
End Enum
#End If
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>  PtrSafe, LongPtr and LongLong
Private Const PTR_LENGTH As Long = 8
Private Const VARIANT_SIZE As Long = 24
#Else                   '<OFFICE97-2010>        Long
Private Const PTR_LENGTH As Long = 4
Private Const VARIANT_SIZE As Long = 16
#End If                 '<WIN32>
'======================
'Public Function App(): Static myApp As New clsApp: Set App = myApp: End Function                ' возвращает ссылку на класс текущего приложения
'Public Function Cmd(): Static myCmd As New clsCommands: Set Cmd = myCmd: End Function           ' возвращает ссылку на класс действий текущего приложения
'Public Function Dbg(): Static myDebug As New clsDebug: Set Dbg = myDebug: End Function          ' возвращает ссылку на класс модуля отладки
'Public Function Crypto(): Static myCrypto As New clsCrypto: Set Crypto = myCrypto: End Function ' возвращает ссылку на класс модуля шифрования
'Public Function Fso(): Static myFso As Object: Set Fso = CreateObject("Scripting.FileSystemObject"): End Function ' возвращает ссылку на класс модуля обработки событий/ошибок
'Public Function Wdr(): Static myWord As New clsWordReport: Set Wdr = myWord: End Function ' возвращает ссылку на класс отчетов Word
'======================
Public Function StartApp(): Call App.AppStart(True): End Function
Public Function StopApp(): App.AppStop (False): End Function
Public Function UpdateAppPath(): Call App.UpdatePath: End Function
Public Function UpdateAppMode(): App.ModeSwitch: End Function
Public Function UpdateLocalRefs(): App.UpdateRefs: End Function
Public Function Setup()
' временное решение для автоматической установки
' установка приложения
'    RestoreRefs     ' восстановление ссылок на библиотеки
    RestoreProp     ' установка свойств приложения
    CloseAll        ' закрываем сохраняем все объекты
    CompileAll      ' компиляция
End Function
Public Sub RestoreProp()
Const c_strProcedure = "RestoreProp"
'Востановление свойств приложения
On Error GoTo HandleError
    SetOption ("Auto Compact"), True            ' сжимать при выходе
    SetOption ("ShowWindowsInTaskbar"), True    ' отключаем окна в панели задач
'    SetOption ("Show Status Bar"), False        ' отключаем строку состояния
    With CurrentProject.Properties
'=== BEGIN INSERT ===
        .Add "Application", c_strApplication ' ="Собрания"
        .Add "Version", c_strVersion ' =""
        .Add "Author", c_strAuthor   ' ="Кашкин Р.В."
        .Add "Support", c_strSupport ' ="KashRus@gmail.com"
        .Add "FirstRun", c_strFirstRun ' ="0"
'==== END INSERT ====
    End With
HandleExit:
    Exit Sub
HandleError:
    Dbg.Error Err.Number, Err.Description, Err.Source, c_strModule & "." & c_strProcedure, Erl()
    Resume HandleExit
End Sub
'======================

Public Sub ContextMenu_Click()
' обработчик событий контекстного меню
    With Application.CommandBars.ActionControl
Stop
Debug.Print .Tag, .Caption
    End With
End Sub



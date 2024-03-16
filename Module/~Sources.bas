Attribute VB_Name = "~Sources"
Option Compare Database
Option Explicit
Option Base 0
#Const APPTYPE = 0          '0|1        '0=ACCESS,1=EXCEL ' not yet
#Const USEZIPCLASS = False  'False|True
'=========================
Private Const c_strModule As String = "~Sources"
'=========================
' Описание      : Модуль для работы с исходным текстом приложения
' Версия        : 1.9.1.451754634
' Дата          : 06.09.2023 11:07:18
' Автор         : Кашкин Р.В. (KashRus@gmail.com)
' Примечание    : при USEZIPCLASS = True, нужен модуль clzZipArchive
' v.1.9.1       : 24.10.2022 - изменения в ZipPack - изменён способ контроля завершения архивирования элемента
' v.1.9.0       : 04.10.2022 - изменены методы сохранения/обновления объектов приложения
' v.1.8.5       : 12.09.2022 - добавлена процедура SourceUpdateAll для массового обновления модулей новыми версиями из локального хранилища
' v.1.8.3       : 24.04.2019 - добавлена возможность чтения версий сохраненных файлов модулей
' v.1.8.1       : 19.04.2019 - обновлен формат сохранения бэкапов - теперь модули сохраняются в стандартные bas и cls.
' v.1.7.12      : 02.04.2019 - добавлена процедура UpdateFunc - изменение версии функции
' v.1.7.6       : 27.09.2018 - в UpdateModule добавлена возможность добавлять комментарии к версии
'=========================
' ToDo: сделать единообразную классификацию для ObjectType (acObjectType|MSysObjectsType|vbext_ComponentType)
'=========================
' Основные свойства и методы модуля:
'-------------------------
' AskAuthor - запрашивает информацию о разработчике и сохраняет в свойствах проекта (используется в комментариях к версиям)
' AskAppData - запрашивает информацию о приложении и сохраняет в свойствах проекта
' ProjBackup/ProjRestore - сохраняют/восстанавливают все объекты проекта в/из файла
' UpdateModule/UpdateFunc - обновляет версию модуля/функции
'=========================
' информация о версии вида A.B.C.D[r][n], где
'   A – главный номер версии (major version number).
'   B – вспомогательный номер версии (minor version number).
'   C – номер сборки, номер логической итерации по работе над функционалом версии A.B (build number).
'   D – номер ревизии, сквозной номер назначаемый автоматически программным обеспечением хранения версий (SVN). Номер ревизии SVN должен синхронизироваться с номером ревизии в AssemblyInfo при каждой сборке релиза (revision number).
'       поскольку в Access номер сборки не отслеживается - здесь будем хранить дату в виде:
'       DDDDDTTTT = CCur(Now)*10^c_bytTimeDig.
'       Восстановление даты билда: CDate(DDDDDTTTT/10^c_bytTimeDig)
'   [r] – условное обозначение релиза, [n] - номер релиза
'       Pre-alpha (pa) – соответствует этапу начала работ над версией. Характеризуется большими изменениями в функционале и большим количеством ошибок. Pre-alpha релизы не покидают отдела разработки ПО.
'       Alpha (a) – соответствует этапу завершения разработки нового функционала. Начиная с alpha версии новый функционал не разрабатывается, а все заявки на новый функционал уходят в план работ по следующей версии. Этап характеризуется высокой активностью по тестированию внутри подразделения разработки ПО и устранению ошибок.
'       Beta (b) – соответствует этапу публичного тестирования. Это первый релиз, который выходит за пределы отдела разработки ПО. На этом этапе принимаются замечания от пользователей по интерфейсу продукта и прочим найденным пользователями ошибкам и неточностям.
'       Release Candidate (rc) – весь функционал реализован и полностью оттестирован, все найденные на предыдущих этапах ошибки исправлены. На этом этапе могут вноситься изменения в документацию и конфигурации продукта.
'       Release to manufacturing или Release to marketing (rtm) – служит для индикации того, что ПО соответствует всем требованиям качества, и готово для массового распространения. RTM не определяет способа доставки релиза (сеть или носитель) и служит лишь для индикации того, что качество достаточно для массового распространения.
'       General availability (ga) – финальный релиз, соответствующий завершению всех работ по коммерциализации продукта, продукт полностью готов к продажам через веб или на физических носителях.
'       End of life (eol) – работы по развитию и поддержке продукта завершены.
'-------------------------

'Private Const c_strLibPath = "%UserProfile%\Documents\VBA Code\" ' путь к библиотеке объектов приложения
Private Const c_strLibPath = "D:\Documents\VBA Code\" ' путь к библиотеке объектов приложения
Private Const c_strPrefModName = "Private Const c_strModule As String = "   ' начало текста строки выше для поиска и замены в модуле

' для процедур обновляющих код приложения
' маркеры начала и окончания области вставки
Private Const strBegLineMarker = "'=== BEGIN INSERT ==="
Private Const strEndLineMarker = "'==== END INSERT ===="
' заголовок модуля
    ' от первого "Attribute VB_"
    ' до первого "[Public | Private | Friend] [Static] [Function | Sub | Property]
Private Const c_strPrefLen As Byte = 17 ' длина префикса до конечного ":"
'Private Const c_strPrefModAttr As String = "Attribute VB_"
Private Const c_strPrefModLine = "'========================="
Private Const c_strPrefModNone = "'               :"
Private Const c_strPrefModDesc = "' Описание      :"
Private Const c_strPrefModAuth = "' Автор         :"
Private Const c_strPrefModVers = "' Версия        :"
Private Const c_strPrefModDate = "' Дата          :"
Private Const c_strPrefModComm = "' Примечание    :"
Private Const c_strPrefModHist = "^\s*'\s*v\.\s*(\d{1,}?\.\d{1,}?\.\d{1,}?)\s*:\s*(\d{1,2}?\.\d{1,2}?\.\d{2,4}?)\s*-\s*(.*)"
Private Const c_strCodeProcBeg = "^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Function|Sub|Property)" '[Public | Private | Friend] [Static] [Function | Sub | Property]" ' заголовок модуля заканчивается с объявлением первой процедуры
Private Const c_strPrefModDebg = "#Const DEBUGGING"

Private Const c_strLibPathLine = "Private Const c_strLibPath = "

Private Const c_strCodeHeadBeg = "CodeBehindForm"       ' текст модуля класса формы/отчета начинается после тэга "CodeBehindForm"
Private Const cEmptyVers = "0.0.0.0", cEmptyDate = #1/1/1980#

Private Const c_strMSysObjects = "MSysObjects"

Private Const c_strHyphen = " _"                ' перенос строки в коде
Private Const c_strSpace = " "                  ' пробел
Private Const c_strBrokenQuotes = """ & """     ' " & " - объединение двух текстовых строк

Private Const c_strDelim As String = ";"
Private Const c_strInDelim As String = ", " ' разделитель элементов текстовых списков

' для сохранения текстов
Private Const c_strSrcPath As String = "SRC"    ' имя подпапки для выгрузки исходников
Private Const c_strBreakProcessMessage = "Прервать процесс?"
' объекты пропускаемые при импорте (чтоб не получилась замена исполняемого кода)
Private Const c_strObjIgnore = "~Sources;clsProgress;frmSERV_Progress" '
' типы и названия папок сохраняемых объектов
Private Const c_strObjTypModule = "Module"
#If APPTYPE = 0 Then        ' APPTYPE=Access
Private Const c_strObjTypAccFrm = "Form", c_strFrmModPref = c_strObjTypAccFrm & "_"
Private Const c_strObjTypAccRep = "Report", c_strRepModPref = c_strObjTypAccRep & "_"   ' префикс модуля отчёта
Private Const c_strObjTypAccMac = "Macro"
Private Const c_strObjTypAccQry = "Query"
Private Const c_strObjTypAccTbl = "Table"
#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
#Else                       ' APPTYPE=Неразбери пойми что
#End If                     ' APPTYPE

' расширение файла выгрузки объектов проекта
Private Const c_strObjExtZip = "zip"
' имена и расширения сохраняемых объектов
Private Const c_strAppNamPrj = "Project"  ' имя файла свойств проекта
Private Const c_strObjExtUndef = "src"  ' старое расширение файла бэкапа объекта
' расширения сохраняемых файлов объектов проекта
Private Const c_strObjExtBas = "bas"    ' стандартный модуль
Private Const c_strObjExtCls = "cls"    ' модуля класса
Private Const c_strObjExtXml = "xml"    ' локальная таблица в XML
Private Const c_strObjExtTxt = "txt"    ' локальная таблица в TXT
Private Const c_strObjExtCsv = "csv"    ' локальная таблица в CSV (текст с разделителями)
#If APPTYPE = 0 Then        ' APPTYPE=Access
Private Const c_strObjExtFrm = "accfrm" ' форма Access (включая модуль)
Private Const c_strObjExtRep = "accrep" ' отчет Access (включая модуль)
Private Const c_strObjExtLnk = "acclnk" ' связанная таблица
Private Const c_strObjExtQry = "accqry" ' запрос Access
Private Const c_strObjExtDoc = "doccls" ' модуль класса документа Access (Form или Report)
Private Const c_strObjExtMac = "accmac" ' макрос Access
Private Const c_strObjExtPrj = "accprj" ' расширение файла свойств проекта Access
#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
#Else                       ' APPTYPE=Неразбери пойми что
#End If                     ' APPTYPE

' имена разделов/ключей в сейвах:
' для свойств
Private Const c_strPrpSecProject = "VBProject"
Private Const c_strPrpKeyPrj = "ProjectName"    ' VBE.ActiveVBProject.Name

Private Const c_strPrpSecCustom = "Custom"      ' CurrentProject.Properties
Private Const c_strPrpKeyApp = "Application"
Private Const c_strPrpKeyVer = "Version"
Private Const c_strPrpKeyAuthor = "Author"
Private Const c_strPrpKeySupport = "Support"

' для свойств проекта
Private Const c_strPrjSecName = "Project"
Private Const c_strPrjKeyName = "Name"
Private Const c_strPrjKeyDesc = "Description"
Private Const c_strPrjKeyHelp = "Help"
' для свойств базы данных
Private Const c_strDbsSecName = "Database Properties"
' для пользовательских свойств
Private Const c_strPrpSecName = "Custom Properties"
' для ссылок (References)
Private Const c_strRefSecName = "References"
' для прилинкованых таблиц
Private Const c_strLnkSecParam = "Params"
Private Const c_strLnkKeyTable = "TableName"
Private Const c_strLnkKeyConnect = "ConnectString"
Private Const c_strLnkKeyLocal = "LocalName"
Private Const c_strLnkKeyAttribute = "Attributes"

' для версий
Private Const c_strPrefVerComm As String = " v."
Private Const c_strVerDelim = "."   ' разделитель компонент версии
Private Const c_bytMajorDig = 2     ' количество разрядов для A
Private Const c_bytMinorDig = 3     ' количество разрядов для B
Private Const c_bytDateDig = 5      ' количество разрядов в выражении ревизии DDDDDTTTT приходящихся на код даты
Private Const c_bytTimeDig = 4      ' количество разрядов в выражении ревизии DDDDDTTTT приходящихся на код времени (не больше 4 иначе в Long переполнение)

Public Enum appRelType      ' тип релиза
    appReleaseNotDefine = 0     'не задан
    appReleasePreAlpha = 1      'Pre-alpha (pa) – соответствует этапу начала работ над версией. Характеризуется большими изменениями в функционале и большим количеством ошибок. Pre-alpha релизы не покидают отдела разработки ПО.
    appReleaseAlpha = 2         'Alpha (a) – соответствует этапу завершения разработки нового функционала. Начиная с alpha версии новый функционал не разрабатывается, а все заявки на новый функционал уходят в план работ по следующей версии. Этап характеризуется высокой активностью по тестированию внутри подразделения разработки ПО и устранению ошибок.
    appReleaseBeta = 3          'Beta (b) – соответствует этапу публичного тестирования. Это первый релиз, который выходит за пределы отдела разработки ПО. На этом этапе принимаются замечания от пользователей по интерфейсу продукта и прочим найденным пользователями ошибкам и неточностям.
    appReleaseCandidate = 4     'Release Candidate (rc) – весь функционал реализован и полностью оттестирован, все найденные на предыдущих этапах ошибки исправлены. На этом этапе могут вноситься изменения в документацию и конфигурации продукта.
    appReleaseToMarketing = 5   'Release to manufacturing или Release to marketing (rtm) – служит для индикации того, что ПО соответствует всем требованиям качества, и готово для массового распространения. RTM не определяет способа доставки релиза (сеть или носитель) и служит лишь для индикации того, что качество достаточно для массового распространения.
    appReleaseGeneral = 6       'General availability (ga) – финальный релиз, соответствующий завершению всех работ по коммерциализации продукта, продукт полностью готов к продажам через веб или на физических носителях.
    appReleaseEOL = 7           'End of life (eol) – работы по развитию и поддержке продукта завершены.
End Enum
Public Enum appVerType      ' тип разряды версии
    appVerMajor = 0
    appVerMinor = 1
    appVerBuild = 2
    appVerRevis = 3
    appRelease = 4
    appRelSubNum = 5
End Enum
Public Type typVersion  ' тип данных версия
    Major As Long           ' A старшая версия
    Minor As Long           ' B младшая версия
    Build As Long           ' C билд
    Revision As Long        ' D ревизия (String?)
    Release As appRelType   ' [r] код типа релиза
    RelSubNum As Integer    ' [n] номер релиза
    RelShort As String      ' краткий текст релиза
    RelFull As String       ' полный текст релиза
    VerDate As Date         ' дата версии CDate(Revision/10^c_bytTimeDig)
    VerCode As Long         ' код версии AABBB
End Type

Private Const c_bolDebugMode = True ' устанавливаемый режим отладки при запуске CompileAll

Private Const c_strSysPath = "%SYSTEMROOT%", c_strPrgPath = "%PROGRAMFILES%"
Private Const c_strSys32Path = c_strSysPath & "\System32\", c_strSys64Path = c_strSysPath & "\SysWoW64\"
Private Const c_strRegKey = "HKEY_CLASSES_ROOT\TypeLib\"

' для преобразования имен
Private Const c_strCodeLen = 2
Private Const c_strCodeSym = "%"
Private Const c_strHexPref = "&H"

Private Const c_strSymbDigits = "0123456789"
Private Const c_strSymbRusAll = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ"
Private Const c_strSymbEngAll = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const c_strOthers = " -~_"

Private Const msoFileDialogViewList = 1
Private Const msoFileDialogFolderPicker = 4

Public Enum eDataType
    dbBoolean = 1
    dbByte = 2
    dbInteger = 3
    dbLong = 4
    dbCurrency = 5
    dbSingle = 6
    dbDouble = 7
    dbDate = 8
    dbBinary = 9
    dbText = 10
    dbLongBinary = 11
    dbMemo = 12
    dbGUID = 15
    dbBigInt = 16
    dbVarBinary = 17
    dbChar = 18
    dbNumeric = 19
    dbDecimal = 20
    dbFloat = 21
    dbTime = 22
    dbTimeStamp = 23
End Enum

' объявления типов и API функций
'--------------------------------------------------------------------------------
' POINTER
'--------------------------------------------------------------------------------
#If VBA7 = 0 Then       'LongPtr trick by @Greedo (https://github.com/Greedquest)
Private Enum LongPtr
    [_]
End Enum
#End If
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>  PtrSafe, LongPtr and LongLong
Private Const PTR_LENGTH As Long = 8
#Else                   '<OFFICE97-2010>        Long
Private Const PTR_LENGTH As Long = 4
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
''#if (_WIN32_WINNT >= 0x0500)
'    pvReserved As LongPtr
'    dwReserved As Long
'    FlagsEx As Long
''#endif // (_WIN32_WINNT >= 0x0500)
End Type
Private Type BROWSEINFO
    hOwner As LongPtr
    pidlRoot As LongPtr
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As LongPtr
    lParam As LongPtr
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1

Private Enum vbext_ProcKind
    vbext_pk_Proc = 0 'A Sub or Function procedure.
    vbext_pk_Let = 1 'A Property Let procedure.
    vbext_pk_Set = 2 'A Property Set procedure.
    vbext_pk_Get = 3 'A Property Get procedure.
End Enum
'!!! сделать единый тип для ObjectType
Private Enum vbext_ComponentType
    vbext_ct_Undef = 0
    vbext_ct_StdModule = 1
    vbext_ct_ClassModule = 2
    vbext_ct_MSForm = 3
    vbext_ct_ActiveXDesigner = 11
    vbext_ct_Document = 100
End Enum
Private Enum msys_ObjectType
    msys_ObjectUndef = 0
    msys_ObjectTable = 1
    msys_ObjectQuery = 5
    msys_ObjectLinked = 6
    msys_ObjectForm = &HFFFF8000
    msys_ObjectReport = &HFFFF8004
    msys_ObjectMacro = &HFFFF8002
    msys_ObjectModule = &HFFFF8007
End Enum

Public Enum appObjectType
' Тип объекта приложения
    appObjTypUndef = 0      ' &HFFFFFFFF
' модули проекта
    appObjTypMod = &H100        ' признак (множитель) типа модуля
    appObjTypBas = appObjTypMod * vbext_ct_StdModule        ' стандартный модуль (vbext_ct_StdModule, acObjectModule; msys_ObjectModule)
    appObjTypCls = appObjTypMod * vbext_ct_ClassModule      ' модуль класса (vbext_ct_ClassModule, acObjectModule; msys_ObjectModule)
    appObjTypMsf = appObjTypMod * vbext_ct_MSForm           ' модуль MSForm (vbext_ct_MSForm)
    appObjTypAxd = appObjTypMod * vbext_ct_ActiveXDesigner  ' модуль ActiveXDesigner (vbext_ct_ActiveXDesigner)
    appObjTypDoc = appObjTypMod * vbext_ct_Document         ' модуль документа (vbext_ct_Document; (acForm; msys_ObjectForm)|(acReport; msys_ObjectReport))
' документы проекта
#If APPTYPE = 0 Then        ' APPTYPE=Access
    appObjTypAccDoc = &H10000   ' признак типа документа
    appObjTypAccTbl = acTable + appObjTypAccDoc             ' таблица (acTable; msys_ObjectTable)
    appObjTypAcclnk = acTable + &H80 + appObjTypAccDoc      ' связанная таблица (acObjectTable & msys_ObjectLinked)
    appObjTypAccQry = acQuery + appObjTypAccDoc             ' запрос Access (acQuery; msys_ObjectQuery)
    appObjTypAccMac = acMacro + appObjTypAccDoc             ' макрос Access (acObjectMacro; msys_ObjectMacro)
    appObjTypAccFrm = acForm + appObjTypDoc + appObjTypAccDoc ' форма Access (acForm; msys_ObjectForm; ModulePrefix="Form_")
    appObjTypAccRep = acReport + appObjTypDoc + appObjTypAccDoc ' отчет Access (acReport & msys_ObjectReport; ModulePrefix="Report_")
    appObjTypAccDap = acDataAccessPage + appObjTypAccDoc    '
    appObjTypAccSrv = acServerView + appObjTypAccDoc        '
    appObjTypAccDia = acDiagram + appObjTypAccDoc           '
    appObjTypAccPrc = acStoredProcedure + appObjTypAccDoc   '
    appObjTypAccFun = acFunction + appObjTypAccDoc          '
#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
    appObjTypXlsDoc = &H20000   ' признак типа документа
'' ...
#Else                       ' APPTYPE=Неразбери пойми что
'' ...
#End If                     ' APPTYPE
End Enum
Public Enum ObjectRwType
' тип операции чтения/записи объектов
    orwUndef = 0
    orwSrcNewer = 1    ' загружаем/сохраняем если источник новее
    orwDestMiss = 2    ' загружаем/сохраняем если назначение отсутствует
    orwSrcOlder = 4    ' загружаем/сохраняем если источник старше
    orwSrcNewerOrDestMissing = orwSrcNewer + orwDestMiss
    orwAlways = orwSrcNewer + orwDestMiss + orwSrcOlder
End Enum

Private Enum m_CodeLineType
    m_CodeNone = -1     ' строка до начала заголовка модуля
    m_CodeHead = 0      ' заголовок модуля. читаемый параметр не найден или чтение параметра завершено
    m_CodeName = 1      ' заголовок модуля: ModName
    m_CodeDesc = 3      ' заголовок модуля: ModDesc
    m_CodeVers = 4      ' заголовок модуля: ModVers
    m_CodeDate = 5      ' заголовок модуля: ModDate
    m_CodeAuth = 6      ' заголовок модуля: ModAuth
    m_CodeComm = 7      ' заголовок модуля: ModComm
    m_CodeHist = 8      ' заголовок модуля: ModHist
    m_CodeProc = 100    ' область процедур модуля (строка после заголовка)
End Enum

Private Enum m_ModErrors
' ошибки при обработке текста модуля
   m_errProcNameWrong = vbObjectError + 511:            ' некорректное имя процедуры
   m_errModuleNameWrong = vbObjectError + 512:          ' некорректное имя модуля
   m_errModuleIsActive = vbObjectError + 513:           ' невозможно изменить активный модуль
   m_errModuleDontFind = vbObjectError + 514:           ' модуль не найден!
' ошибки взаимодействия с объектами
   m_errObjectTypeUnknown = vbObjectError + 520         ' невозможно прочитать объект данного типа
   m_errObjectActionUndef = vbObjectError + 530         ' непонятно что делать с объектом
   m_errObjectCantRemove = vbObjectError + 538          ' не удалось удалить объект из проекта
   m_errCantGetSrcVersion = vbObjectError + 541         ' не удалось прочитать версию источника
   m_errCantGetDestVersion = vbObjectError + 542        ' не удалось прочитать версию назначения
' причины пропуска при опреациях с объектами
   m_errWrongVersion = vbObjectError + 1001             ' несоответствие версии
   m_errDestMissing = vbObjectError + 1002              ' целевой объект отсутствует
   m_errDestExists = vbObjectError + 1003               ' целевой объект уже существует
   m_errSkippedByUser = vbObjectError + 1004            ' решение пользователя
   m_errSkippedByList = vbObjectError + 1005            ' наличие в списке пропуска
' причины пропуска при опреациях с объектами
   m_errZipError = vbObjectError + 2001                 ' ошибка упаковки
   m_errUnZipError = vbObjectError + 2002               ' ошибка распаковки
   m_errExportError = vbObjectError + 2101              ' ошибка экспорта
   m_errImportError = vbObjectError + 2102              ' ошибка импорта
End Enum

#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare PtrSafe Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Boolean
Private Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As LongPtr
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As LongPtr, ByVal pszPath As String) As Boolean

Private Declare PtrSafe Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal strSection As String, ByVal strRegKeyName As String, ByVal strDefault As String, ByVal strReturned As String, ByVal lngSize As Long, ByVal strFilename As String) As Long
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal strSection As String, ByVal strRegKeyNam As String, ByVal strValue As String, ByVal strFilename As String) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else                   '<OFFICE97-2010>
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Boolean
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Boolean

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal strSection As String, ByVal strRegKeyName As String, ByVal strDefault As String, ByVal strReturned As String, ByVal lngSize As Long, ByVal strFilename As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal strSection As String, ByVal strRegKeyNam As String, ByVal strValue As String, ByVal strFilename As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If                 '<WIN32>
'-------------------------
' функции для получения базовой информации о приложении и разработчике
'-------------------------
Public Function UpdateAppVer(Optional strVersion As String)
' спрашиваем и устанавливаем в свойство версию приложения
Const c_strProcedure = "UpdateAppVer"
Dim Result As Boolean ':Result = False
    On Error GoTo HandleError
Dim strTitle As String:     strTitle = "Версия приложения"
Dim strMessage As String:   strMessage = "Введите версию приложения."
    If Len(strVersion) = 0 Then Call PropertyGet(c_strPrpKeyVer, strVersion)
    Dim strVer1 As String, strVer2 As String, datVerDate As Date
    VersionGet strVersion, VerShort:=strVer1, VerDate:=datVerDate   ' получаем короткую версию из имеющейся
    strVersion = VBA.Trim$(InputBox(strMessage, strTitle, strVer1)) ' спрашиваем новую версию
    VersionGet strVersion, VerShort:=strVer2, VerDate:=datVerDate   ' получаем короткую версию из новой
    ' если версии не совпадают - обновляем
    If strVer2 <> strVer1 Then datVerDate = Now: VersionSet strVersion, VerDate:=datVerDate: Call PropertySet(c_strPrpKeyVer, strVersion)
    Result = True
HandleExit:     UpdateAppVer = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function

Public Function AskAppData( _
    Optional strAppName As String, _
    Optional strVersion As String, _
    Optional strCodeName As String, _
    Optional strDescription As String _
    )
' спрашивает имя/версию приложения и обновляет соответствующие свойства проекта
Const c_strProcedure = "AskAppData"
Dim strMessage As String, strTitle As String
Dim Result As Boolean

    Result = False
    On Error GoTo HandleError
    strTitle = "Обновление свойств приложения!"
    strMessage = "Будет произведено обновление свойств приложения," & vbCrLf _
        & "отвечающих за название и текущую версию приложения."
    MsgBox strMessage, vbOKOnly Or vbExclamation, strTitle
' спрашиваем и устанавливаем в свойство название приложения
    strTitle = "Название приложения"
    strMessage = "Введите название приложения."
    If Len(strAppName) = 0 Then Call PropertyGet(c_strPrpKeyApp, strAppName): If Len(strAppName) = 0 Then strVersion = "MyApp"
    strAppName = VBA.Trim$(InputBox(strMessage, strTitle, strAppName)): Call PropertySet(c_strPrpKeyApp, strAppName)
' спрашиваем и устанавливаем в свойство версию приложения
    UpdateAppVer strVersion
' спрашиваем и устанавливаем кодовое имя приложения (имя проекта)
' VBE.ActiveVBProject.Name
    strTitle = "Кодовое имя приложения"
    strMessage = "Введите кодовое имя приложения."
    If Len(strCodeName) = 0 Then strCodeName = VBE.ActiveVBProject.Name
    strCodeName = VBA.Trim$(InputBox(strMessage, strTitle, strCodeName)): VBE.ActiveVBProject.Name = strCodeName
' VBE.ActiveVBProject.Description
    strTitle = "Описание приложения"
    strMessage = "Введите краткое описание приложения."
    If Len(strDescription) = 0 Then strDescription = VBE.ActiveVBProject.Description
    strDescription = VBA.Trim$(InputBox(strMessage, strTitle, strDescription)): VBE.ActiveVBProject.Description = strDescription
    
    Result = True
HandleExit:     AskAppData = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Public Function AskAuthor( _
    Optional strAuthor As String, _
    Optional strSupport As String)
' спрашивает имя/контакты разработчика и обновляет соответствующие свойства проекта
Const c_strProcedure = "AskAuthor"
Dim strMessage As String, strTitle As String
    On Error GoTo HandleError
    strTitle = "Обновление свойств приложения!"
    strMessage = "Будет произведено обновление свойств приложения," & vbCrLf _
        & "отвечающих за имя и контактные данные разработчика." & vbCrLf _
        & "Эти данные используются для указания автора изменений " & vbCrLf _
        & "в комментариях при использовании функций UpdateModule и UpdateFunc."
    MsgBox strMessage, vbOKOnly Or vbExclamation, strTitle
' спрашиваем и устанавливаем в свойство проекта имя разработчика
    strTitle = "Имя автора"
    strMessage = "Введите данные об имени разработчика, " & vbCrLf & _
        "которые будут выводиться в комментариях " & vbCrLf & _
        "к изменениям модулей/функций."
    If Len(strAuthor) = 0 Then Call PropertyGet(c_strPrpKeyAuthor, strAuthor): If Len(strAuthor) = 0 Then strAuthor = "Unknown"
    strAuthor = VBA.Trim$(InputBox(strMessage, strTitle, strAuthor)): Call PropertySet(c_strPrpKeyAuthor, strAuthor)
' спрашиваем и устанавливаем в свойство проекта контакты разработчика
    strTitle = "Контакты автора"
    strMessage = "Введите контактные данные разработчика," & vbCrLf & _
        "которые будут выводиться в комментариях " & vbCrLf & _
        "к изменениям модулей/функций вместе с именем."
    If Len(strSupport) = 0 Then Call PropertyGet(c_strPrpKeySupport, strSupport): If Len(strSupport) = 0 Then strSupport = "dont@mail.me"
    strSupport = VBA.Trim$(InputBox(strMessage, strTitle, strSupport)): Call PropertySet(c_strPrpKeySupport, strSupport)
HandleExit:     Exit Function
HandleError:    Err.Clear: Resume HandleExit
End Function
Public Function Author() As String
Dim Result As String
    On Error GoTo HandleError
    Call PropertyGet(c_strPrpKeyAuthor, Result): If Len(Result) = 0 Then AskAuthor strAuthor:=Result
HandleExit:     Author = Result: Exit Function
HandleError:    Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function Support() As String
Dim Result As String
    On Error GoTo HandleError
    Call PropertyGet(c_strPrpKeySupport, Result): If Len(Result) = 0 Then AskAuthor strSupport:=Result
HandleExit:     Support = Result: Exit Function
HandleError:    Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function AppName() As String
Dim Result As String
    On Error GoTo HandleError
    Call PropertyGet(c_strPrpKeyApp, Result): If Len(Result) = 0 Then AskAppData strAppName:=Result
HandleExit:     AppName = Result: Exit Function
HandleError:    Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function AppVersion() As String
Dim Result As String
    On Error GoTo HandleError
    Call PropertyGet(c_strPrpKeyVer, Result): If Len(Result) = 0 Then AskAppData strVersion:=Result
HandleExit:     AppVersion = Result: Exit Function
HandleError:    Result = vbNullString: Err.Clear: Resume HandleExit
End Function
'-------------------------
' функции сохранения/восстановления проекта
'-------------------------
Public Function Init()
' инициализация проекта
Dim Result As VbMsgBoxResult
    CloseAll
    Dim strMessage As String
    strMessage = "Восстановить базу с архивной копии?"
    Result = MsgBox(strMessage, vbQuestion + vbYesNo)
    If Result = vbYes Then
        strMessage = "Для завершения настройки после восстановления базы " & vbCrLf & _
            "проверьте восстановление из архива всех необходимых таблиц и др.объектов." & vbCrLf & vbCrLf & _
            "Иногда при удалении временной папки после восстановления из архива падает Access." & vbCrLf & _
            "В таком случае надо перезапустить базу и завершить восстановление выполнив макрос ""Setup"" позже."
        Call MsgBox(strMessage, vbExclamation + vbOKOnly)
        SourcesRestore
        'App.Setup ' !!! появится после восстановления
    End If
'    ' запрашиваем данные разработчика
'    strMessage = "Обновить данные о разработчике?"
'    Result = MsgBox(strMessage, vbQuestion + vbYesNo)
'    If Result = vbYes Then AskAuthor
End Function
Public Function SourcesBackup(Optional BackupPath As String, _
    Optional WithoutData As Boolean = True) As Boolean
' сохраняет все объекты проекта в архивную копию для последующего восстановления
Const c_strProcedure = "SourcesBackup"
Dim Result As Boolean
On Error GoTo HandleError

'Dim WithoutData As Boolean:     WithoutData = True      ' при бэкапе пропускает объекты данных (таблицы и т.п.)
Dim WriteType As ObjectRwType:  WriteType = orwAlways   ' делаем бэкап всех объектов приложения независимо от версий
Dim AskBefore As Boolean:       AskBefore = False       ' не спрашивать пользователя перед сохранением объекта
Dim UseTypeFolders As Boolean:  UseTypeFolders = True   ' файлы в бэкапе сортируем по папкам объектов
Dim DelAfterZip As Boolean:     DelAfterZip = True      ' удаляем временную папку после архивирования

#If APPTYPE = 0 Then        ' APPTYPE=Access
    DoCmd.SetWarnings False ' отключаем предупреждения
    Call SysCmd(504, 16484) ' сохраняем все модули
#Else
#End If                     ' APPTYPE

Dim strCaption As String, strMessage As String
Dim ParentPath As String, FilePath As String, FileName As String
' создаем путь сохранения
    If Len(BackupPath) = 0 Then BackupPath = oFso.BuildPath(CurrentProject.path, c_strSrcPath)
    If Not oFso.FolderExists(BackupPath) Then Call oFso.CreateFolder(BackupPath) 'Then Err.Raise 76 ' Path not Found
' создаем имя архива
    FileName = Split(CurrentProject.Name, ".")(0)
    FileName = FileName & "_" & VBA.Format$(Now(), "YYYYMMDD_hhmmss")
    ParentPath = oFso.BuildPath(BackupPath, FileName)
    If Not oFso.FolderExists(ParentPath) Then Call oFso.CreateFolder(ParentPath) 'Then Err.Raise 76 ' Path not Found
' формируем коллекцию всех объектов подлежащих сохранению
Dim colObjects As Collection:   Set colObjects = New Collection
    If Not p_ObjectsCollectionCreate(colObjects, WithoutData) Then Err.Raise m_errExportError
' инициалиизация прогрессбара
Dim prg As clsProgress: Set prg = New clsProgress
    strCaption = "Сохранение объектов проекта"
    strMessage = strCaption & " в: """ & BackupPath & """"
    prg.Init pCount:=1, pMin:=1, pMax:=colObjects.Count + 1, pCaption:=strCaption, pText:=strMessage ': prg.ProgressStep = 1
' сохраняем свойства и библиотечные ссылки проекта
    prg.Detail = "Сохранение свойств и библиотечных ссылок проекта"
    FilePath = oFso.BuildPath(ParentPath, c_strAppNamPrj & "." & c_strObjExtPrj)
    p_PropertiesWrite FilePath
    p_ReferencesWrite FilePath
    prg.Update
' выгружаем поочерёдно все объекты
    Result = p_ObjectsBackup(colObjects, ParentPath, prg, WriteType, AskBefore, UseTypeFolders, strMessage) = 0:   If Not Result Then Err.Raise m_errExportError
    'If colObjects.Count > 0 Then ' не все элементы удалось сохранить
    Set colObjects = Nothing
    
' упаковываем выгрузку
    FilePath = ParentPath & "." & c_strObjExtZip
    prg.Detail = "Завершено сохранение свойств и библиотечных ссылок проекта." & vbCrLf & _
                 "Идёт упаковка сохранённых объектов в архив: " & FilePath
#If USEZIPCLASS Then
    Result = oZip.AddFromFolder(ParentPath & "\*.*", True, , True)
    Result = oZip.CompressArchive(FilePath)
    If Result And DelAfterZip Then Result = oFso.DeleteFolder(ParentPath) = 0
#Else
    Result = ZipPack(FilePath:=ParentPath & "\*.*", ZipName:=FilePath, DelAfterZip:=True)
#End If ' USEZIPCLASS
    If Not Result Then Err.Raise m_errZipError
#If APPTYPE = 0 Then        ' APPTYPE=Access
    DoCmd.SetWarnings True  ' включаем предупреждения
    Call SysCmd(504, 16484) ' сохраняем все модули
#Else
#End If                     ' APPTYPE
' информируем пользователя о результатах
    strCaption = "Экспорт завершен"
    strMessage = strMessage & vbCrLf & "Объекты текущего проекта были сохранены в:" & vbCrLf & FilePath & "." & vbCrLf
    Call MsgBox(strMessage, vbInformation + vbOKOnly, strCaption)
HandleExit:     SourcesBackup = Result: Set prg = Nothing: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 75:    Message = "Ошибка доступа!" ' возможно папка уже создана
    If Len(BackupPath) > 0 Then Message = Message & " BackupPath=""" & BackupPath & """ "
    Case 76:    Message = "Путь не найден!" ' возможно нет родительской папки
    If Len(BackupPath) > 0 Then Message = Message & " BackupPath=""" & BackupPath & """ "
    Case 1004 '??? ' Ошибка: Программный доступ к проекту Visual Basic не является доверенным
                Message = "Программный доступ к проекту Visual Basic не является доверенным. Для возможности программного сохранения/восстановления модулей необходимо установить разрешение: ""Сервис\Макрос\Безопасность\Доверять доступ к Visual Basic Project"""
    Case m_errZipError:     Message = "Ошибка упаковки объектов проекта"
    Case m_errExportError:  Message = "Ошибка при экспорте объектов проекта"
    Case Else:  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    MsgBox Message & vbCrLf & "Создание архивной копии не завершено.", vbExclamation
    Err.Clear: Resume HandleExit
    Err.Clear: Resume 0
End Function
Public Function SourcesRestore(Optional SourcePath As String, _
    Optional WithoutData As Boolean = True) As Boolean
' восстанавливает все объекты проекта из архивной копии
Const c_strProcedure = "SourcesRestore"
Dim iTry As Integer
Dim Result As Boolean
On Error GoTo HandleError

'Dim WithoutData As Boolean:     WithoutData = True      ' при восстановлении пропускает объекты данных (таблицы и т.п.)
Dim ReadType As ObjectRwType:   ReadType = orwAlways    ' восстанавливаем все объекты приложения сохраненные в бэкапе независимо от версий
Dim AskBefore As Boolean:       AskBefore = True        ' спрашивать пользователя перед восстановлением существующего объекта
Dim UseTypeFolders As Boolean:  UseTypeFolders = True   ' файлы в бэкапе отсортированы по папкам объектов

#If APPTYPE = 0 Then        ' APPTYPE=Access
    DoCmd.SetWarnings False ' отключаем предупреждения
    Call SysCmd(504, 16484) ' сохраняем все модули
#Else
#End If                     ' APPTYPE
Dim strCaption As String, strMessage As String
Dim TempPath As String, FilePath As String

    If Len(SourcePath) = 0 Then SourcePath = oFso.BuildPath(CurrentProject.path, c_strSrcPath)
' проверяем/готовим пути
    Result = oFso.FileExists(SourcePath): If Result Then GoTo HandleCreateTempPath
    ' запрашиваем имя файла backup
    strCaption = "Выберите архивную копию проекта для восстановления"
    strMessage = "Архивная копия проекта " & VBA.Chr$(0) & "*.zip"
    SourcePath = p_SelectFile(SourcePath, strMessage, c_strObjExtZip, strCaption)
    Result = Len(SourcePath) > 0: If Not Result Then Err.Raise 76
HandleCreateTempPath:
    ' создаем временную папку
    TempPath = oFso.BuildPath(VBA.Environ$("Temp"), "~" & oFso.GetFileName(SourcePath))
    If Not oFso.FolderExists(TempPath) Then Call oFso.CreateFolder(TempPath) 'Then Err.Raise 76 ' Path not Found
' инициалиизация прогрессбара

Dim prg As clsProgress: Set prg = New clsProgress
    strCaption = "Восстановление объектов проекта"
    strMessage = strCaption & " из: """ & SourcePath & """"
    prg.Init pCount:=1, pMin:=1, pMax:=1, pCaption:=strCaption, pText:=strMessage ': prg.ProgressStep = 1
' распаковываем выгрузку во временную папку архива
    prg.Detail = "Распаковка файлов бэкапа во временную папку"
#If USEZIPCLASS Then
    Result = oZip.OpenArchive(SourcePath): If Not Result Then Err.Raise m_errUnZipError
    Result = oZip.Extract(TempPath): If Not Result Then Err.Raise m_errUnZipError
#Else
    Result = ZipUnPack(SourcePath, TempPath): If Not Result Then Err.Raise m_errUnZipError
#End If ' UseZipArchive

' формируем коллекцию всех объектов подлежащих восстановлению
Dim colObjects As Collection:     Set colObjects = New Collection
    If UseTypeFolders Then  ' читаем содержимое подпапок
Dim oItem As Object
        For Each oItem In oFso.GetFolder(TempPath).SubFolders
        ' проверяем соответствие имени папки допустимым типам объектов
            Select Case oFso.GetFileName(oItem)
            Case c_strObjTypModule
#If APPTYPE = 0 Then        ' APPTYPE=Access
            Case c_strObjTypAccFrm, c_strObjTypAccRep
            Case c_strObjTypAccQry, c_strObjTypAccMac
            Case c_strObjTypAccTbl: If WithoutData Then GoTo HandleNextFolder
'#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
#Else                       '
#End If                     ' APPTYPE
            Case Else: GoTo HandleNextFolder
            End Select
'
            If Not p_ObjectFilesCollectionCreate(oItem.path, colObjects, WithoutData) Then Err.Raise m_errImportError
HandleNextFolder:
        Next oItem
        Set oItem = Nothing
    Else                     ' читаем содержимое только текущей папки
            If Not p_ObjectFilesCollectionCreate(TempPath, colObjects, WithoutData) Then Err.Raise m_errImportError
    End If

' инициалиизация шкалы прогрессбара счётчиком
'    strCaption = "Восстановление объектов проекта"
'    strMessage = strCaption & " из: """ & SourcePath & """"
    prg.ProgressMax = colObjects.Count + 1
' восстанавливаем свойства и библиотечные ссылки проекта
    prg.Detail = "Восстановление свойств и библиотечных ссылок проекта"
    FilePath = oFso.BuildPath(TempPath, c_strAppNamPrj & "." & c_strObjExtPrj)
    p_PropertiesRead FilePath: p_ReferencesRead FilePath
    prg.Update
    'prg.Detail = "Завершено восстановление свойств и библиотечных ссылок проекта": prg.Detail = strMessage
' подгружаем поочерёдно все объекты
    Result = p_ObjectsRestore(colObjects, prg, ReadType, AskBefore, UseTypeFolders, strMessage) = 0:    If Not Result Then Err.Raise m_errExportError
    'If colObjects.Count > 0 Then ' не все элементы удалось восстановить
    Set colObjects = Nothing
' удаляем временную папку ' иногда здесь падает
    Result = oFso.DeleteFolder(TempPath) = 0
#If APPTYPE = 0 Then        ' APPTYPE=Access
    DoCmd.SetWarnings True  ' включаем предупреждения
    Call SysCmd(504, 16484) ' сохраняем все модули
#Else
#End If                     ' APPTYPE
' информируем пользователя о результатах
    strCaption = "Импорт завершен"
    strMessage = strMessage & vbCrLf & "Объекты текущего проекта были восстановлены из:" & vbCrLf & SourcePath & "." & vbCrLf
    Call MsgBox(strMessage, vbInformation + vbOKOnly, strCaption)

HandleExit:     SourcesRestore = Result: Set prg = Nothing: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 58:    Err.Clear: Resume Next      ' возможно папка уже создана
    Case 70:    If iTry < 3 Then iTry = iTry + 1: Err.Clear: Sleep 333: Resume Next
                Message = "Не удалось полностью удалить временную папку."
                If Len(TempPath) > 0 Then Message = Message & " TempPath=""" & TempPath & """ "
    Case 75:    Message = "Ошибка доступа!" ' возможно папка уже создана
    Case 76:    Message = "Путь не найден!" ' возможно нет родительской папки
    If Len(SourcePath) > 0 Then Message = Message & " SourcePath=""" & SourcePath & """ "
    If Len(TempPath) > 0 Then Message = Message & " TempPath=""" & TempPath & """ "
    Case 1004 '??? ' Ошибка: Программный доступ к проекту Visual Basic не является доверенным
                Message = "Программный доступ к проекту Visual Basic не является доверенным. Для возможности программного сохранения/восстановления модулей необходимо установить разрешение: ""Сервис\Макрос\Безопасность\Доверять доступ к Visual Basic Project"""
    Case m_errUnZipError:   Message = "Ошибка распаковки объектов проекта"
    Case m_errImportError:  Message = "Ошибка при импорте объектов проекта"
    Case Else:  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    MsgBox Message & vbCrLf & "Восстановление из архивной копии не завершено.", vbExclamation
    Err.Clear: Resume HandleExit
End Function
Public Function SourcesUpdateFromStorage(Optional SourcePath As String)
' производит массовое обновление модулей расположенными в хранилище имеющими более старшие версии
Const c_strProcedure = "SourcesUpdateFromStorage"

Dim Result As Boolean
On Error GoTo HandleError

'Dim WithoutData As Boolean:     WithoutData = True      ' при восстановлении пропускает объекты данных (таблицы и т.п.)
Dim ReadType As ObjectRwType:   ReadType = orwSrcNewer  ' обновляем только если объект в библиотеке новее
Dim AskBefore As Boolean:       AskBefore = True        ' спрашиваем пользователя перед обновлением объекта
Dim OnlyExisting As Boolean:    OnlyExisting = True     ' добавляем в коллекцию на обновление только существующие в приложении объекты
'Dim UseTypeFolders As Boolean:  UseTypeFolders = False  ' файлы в библиотеке не отсортированы по папкам типов

#If APPTYPE = 0 Then        ' APPTYPE=Access
    DoCmd.SetWarnings False ' отключаем предупреждения
    Call SysCmd(504, 16484) ' сохраняем все модули
#Else
#End If                     ' APPTYPE

Dim strCaption As String, strMessage As String
    
    If Len(SourcePath) = 0 Then SourcePath = c_strLibPath
' проверяем/готовим пути
    Result = oFso.FolderExists(SourcePath): If Result Then GoTo HandleUpdateLibPath
    ' запрашиваем путь к хранилищу модулей
    strCaption = "Укажите путь к хранилищу последних версий используемых модулей."
    SourcePath = p_SelectFolder(SourcePath, DialogTitle:=strCaption)
    Result = Len(SourcePath) > 0: If Not Result Then Err.Raise 75
HandleUpdateLibPath:
    If SourcePath = c_strLibPath Then GoTo HandleUpdateSources
' обновляем в тексте модуля значение константы c_strLibPath на SourcePath
Dim objModule As Object: If Not ModuleExists(c_strModule, objModule) Then Err.Raise m_errModuleDontFind
Dim CodeLine As Long: CodeLine = p_CodeLineFind(objModule, c_strLibPathLine, CodeLine) ' маркер на начало области вставки
    With objModule
        If CodeLine > 0 Then .DeleteLines CodeLine, 1 Else CodeLine = .CountOfDeclarationLines
        .InsertLines CodeLine, c_strLibPathLine & " """ & SourcePath & """"
    End With

HandleUpdateSources:
' инициалиизация прогрессбара
Dim prg As clsProgress: Set prg = New clsProgress
    strCaption = "Обновление объектов проекта"
    strMessage = strCaption & " из библиотеки: """ & SourcePath & """"
    prg.Init pCount:=1, pMin:=1, pMax:=1, pCaption:=strCaption, pText:=strMessage ': prg.ProgressStep = 1
' формируем коллекцию всех объектов подлежащих восстановлению
Dim colObjects As Collection:     Set colObjects = New Collection
' читаем содержимое только текущей папки
    If Not p_ObjectFilesCollectionCreate(SourcePath, colObjects, True, OnlyExisting) Then Err.Raise m_errImportError
' инициалиизация шкалы прогрессбара счётчиком
    prg.ProgressMax = colObjects.Count ': prg.ProgressStep = 1
' подгружаем поочерёдно все объекты
    Result = p_ObjectsRestore(colObjects, prg, ReadType, AskBefore, False, strMessage) = 0:     If Not Result Then Err.Raise m_errExportError
    'If colObjects.Count > 0 Then ' не все элементы удалось восстановить
    Set colObjects = Nothing
' информируем пользователя о результатах
    strCaption = "Обновление завершено"
    strMessage = strMessage & vbCrLf & vbCrLf & "Объекты текущего проекта были обновлены из:" & vbCrLf & SourcePath & "." & vbCrLf
    Call MsgBox(strMessage, vbInformation + vbOKOnly, strCaption)
    Set prg = Nothing
HandleExit:     SourcesUpdateFromStorage = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 75:    Message = "Ошибка доступа!" ' возможно папка уже создана
    If Len(SourcePath) > 0 Then Message = Message & " SourcePath=""" & SourcePath & """ "
    Case 76:    Message = "Путь не найден!" ' возможно нет родительской папки
    If Len(SourcePath) > 0 Then Message = Message & " SourcePath=""" & SourcePath & """ "
    Case 1004 '??? ' Ошибка: Программный доступ к проекту Visual Basic не является доверенным
                Message = "Программный доступ к проекту Visual Basic не является доверенным. Для возможности программного сохранения/восстановления модулей необходимо установить разрешение: ""Сервис\Макрос\Безопасность\Доверять доступ к Visual Basic Project"""
    Case m_errModuleNameWrong:  Message = "Неверно задано имя объекта!"
    Case m_errObjectTypeUnknown: Message = "Неизвестный тип модуля!"
    Case Else:                  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function SourcesUpdateStorage(Optional BackupPath As String)
' производит массовое сохранение модулей проекта имеющих более старшие версии в хранилище файлов объектов
Const c_strProcedure = "SourcesUpdateStorage"

Dim Result As Boolean
On Error GoTo HandleError

'Dim WithoutData As Boolean:     WithoutData = True      ' при восстановлении пропускает объекты данных (таблицы и т.п.)
Dim WriteType As ObjectRwType:  WriteType = orwSrcNewerOrDestMissing  ' сохраняем если объект в проекте новее или отсутствует в библиотеке
Dim AskBefore As Boolean:       AskBefore = True        ' спрашиваем пользователя перед обновлением объекта
Dim OnlyExisting As Boolean:    OnlyExisting = True     ' добавляем в коллекцию на обновление только существующие в приложении объекты
'Dim UseTypeFolders As Boolean:  UseTypeFolders = False  ' файлы в библиотеке не отсортированы по папкам типов

#If APPTYPE = 0 Then        ' APPTYPE=Access
    DoCmd.SetWarnings False ' отключаем предупреждения
    Call SysCmd(504, 16484) ' сохраняем все модули
#Else
#End If                     ' APPTYPE

Dim strCaption As String, strMessage As String
    
    If Len(BackupPath) = 0 Then BackupPath = c_strLibPath
' проверяем/готовим пути
    Result = oFso.FolderExists(BackupPath): If Result Then GoTo HandleUpdateLibPath
    ' запрашиваем путь к хранилищу модулей
    strCaption = "Укажите путь к хранилищу последних версий используемых модулей."
    BackupPath = p_SelectFolder(BackupPath, DialogTitle:=strCaption)
    Result = Len(BackupPath) > 0: If Not Result Then Err.Raise 75
HandleUpdateLibPath:
    If BackupPath = c_strLibPath Then GoTo HandleUpdateSources
' обновляем в тексте модуля значение константы c_strLibPath на SourcePath
Dim objModule As Object: If Not ModuleExists(c_strModule, objModule) Then Err.Raise m_errModuleDontFind
Dim CodeLine As Long: CodeLine = p_CodeLineFind(objModule, c_strLibPathLine, CodeLine) ' маркер на начало области вставки
    With objModule
        If CodeLine > 0 Then .DeleteLines CodeLine, 1 Else CodeLine = .CountOfDeclarationLines
        .InsertLines CodeLine, c_strLibPathLine & " """ & BackupPath & """"
    End With

HandleUpdateSources:
' инициалиизация прогрессбара
Dim prg As clsProgress: Set prg = New clsProgress
    strCaption = "Сохранение объектов проекта"
    strMessage = strCaption & " в библиотеку: """ & BackupPath & """"
    prg.Init pCount:=1, pMin:=1, pMax:=1, pCaption:=strCaption, pText:=strMessage ': prg.ProgressStep = 1
' формируем коллекцию всех объектов подлежащих восстановлению
Dim colObjects As Collection:     Set colObjects = New Collection
' читаем содержимое только текущей папки
    If Not p_ObjectsCollectionCreate(colObjects, True) Then Err.Raise m_errImportError
' инициалиизация шкалы прогрессбара счётчиком
    prg.ProgressMax = colObjects.Count
' сохраняем поочерёдно все объекты
    Result = p_ObjectsBackup(colObjects, BackupPath, prg, WriteType, AskBefore, False, strMessage) = 0:      If Not Result Then Err.Raise m_errExportError
    'If colObjects.Count > 0 Then ' не все элементы удалось сохранить
    Set colObjects = Nothing
' информируем пользователя о результатах
    strCaption = "Сохранение завершено"
    strMessage = strMessage & vbCrLf & vbCrLf & "Объекты текущего проекта были сохранены в:" & vbCrLf & BackupPath & "." & vbCrLf
    Call MsgBox(strMessage, vbInformation + vbOKOnly, strCaption)
    
HandleExit:     SourcesUpdateStorage = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 75:    Message = "Ошибка доступа!" ' возможно папка уже создана
    If Len(BackupPath) > 0 Then Message = Message & " BackupPath=""" & BackupPath & """ "
    Case 76:    Message = "Путь не найден!" ' возможно нет родительской папки
    If Len(BackupPath) > 0 Then Message = Message & " BackupPath=""" & BackupPath & """ "
    Case 1004 '??? ' Ошибка: Программный доступ к проекту Visual Basic не является доверенным
                Message = "Программный доступ к проекту Visual Basic не является доверенным. Для возможности программного сохранения/восстановления модулей необходимо установить разрешение: ""Сервис\Макрос\Безопасность\Доверять доступ к Visual Basic Project"""
    Case m_errModuleNameWrong:  Message = "Неверно задано имя объекта!"
    Case m_errObjectTypeUnknown: Message = "Неизвестный тип модуля!"
    Case Else:                  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_ObjectsCollectionCreate(colObjects As Collection, _
    Optional WithoutData As Boolean = True) As Boolean
' формирует коллекцию объектов приложения
Dim Result As Boolean: Result = False
' WithoutData - определяет будут ли включены в коллекцию объекты данных (таблицы и пр.)
Dim BackupDocs As Boolean
#If APPTYPE = 0 Then        ' APPTYPE=Access
    BackupDocs = False      ' модули форм отчетов итак выгружаются вместе с формой/отчётом
#Else
    BackupDocs = True       ' с Excel и пр. что делать пока не решил
#End If                     ' APPTYPE
    On Error GoTo HandleError
Dim oItems As Object, oItem
' сначала перебираем все модули
    ' VBE.ActiveVBProject.VBComponents существует во всех приложениях Microsoft
    ' но к ним может не быть доступа - тогда возникнет ошибка 1004
    ' надо будет вручную утановить: "Доверять доступ к Visual Basic Project"
    Set oItems = Application.VBE.ActiveVBProject.VBComponents
    For Each oItem In oItems
        Select Case oItem.Type
        Case vbext_ct_StdModule, vbext_ct_ClassModule: colObjects.Add oItem, oItem.Name
        Case vbext_ct_Document: If BackupDocs Then colObjects.Add oItem, oItem.Name
        Case Else ' ???
        End Select
    Next oItem
' потом специфические для приложений объекты
#If APPTYPE = 0 Then        ' APPTYPE=Access
    ' в Access это проще всего сделать запросом по таблице c_strMSysObjects
    '' или можно пройтись по коллекциям объектов
    'Set oItems = CurrentProject.AllModules:         For Each oItem In oItems: colObjects.Add oItem, oItem.NAME: Next oItem
    Set oItems = CurrentProject.AllForms:           For Each oItem In oItems: colObjects.Add oItem, oItem.Name: Next oItem '
    Set oItems = CurrentProject.AllReports:         For Each oItem In oItems: colObjects.Add oItem, oItem.Name: Next oItem
    Set oItems = CurrentProject.AllMacros:          For Each oItem In oItems: colObjects.Add oItem, oItem.Name: Next oItem
    Set oItems = CurrentProject.AllDataAccessPages: For Each oItem In oItems: colObjects.Add oItem, oItem.Name: Next oItem
    Set oItems = CurrentData.AllQueries:            For Each oItem In oItems: colObjects.Add oItem, oItem.Name: Next oItem
    If Not WithoutData Then
    Set oItems = CurrentData.AllTables:             For Each oItem In oItems
                                                    ' пропускаем системные таблицы
                                                        If Left(oItem.Name, 4) <> "MSys" Then colObjects.Add oItem, oItem.Name
                                                    Next oItem
    End If
#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
    ' в Excel прийдётся перебирать коллекции объектов в ThisWorkbook
#Else                       ' APPTYPE=Неразбери пойми что
#End If                     ' APPTYPE
    Set oItems = Nothing: Set oItem = Nothing
    Result = True
HandleExit:  p_ObjectsCollectionCreate = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_ObjectFilesCollectionCreate(FilesPath As String, colObjects As Collection, _
    Optional WithoutData As Boolean = True, Optional OnlyExisting As Boolean = False) As Boolean
' формирует коллекцию файлов объектов приложения расположенных по указанному пути
Dim Result As Boolean ': Result = False
' WithoutData - определяет будут ли включены в коллекцию файлы бэкапов данных (таблицы и пр.)
' OnlyExisting - определяет будут ли включены в коллекцию файлы объектов отсутствующих в приложении
    On Error GoTo HandleError
Dim oItem As Object
    ' перебираем все файлы в подпапках типов объектов в корневой папке
'' SubFolders похоже не имеет числового индекса
    For Each oItem In oFso.GetFolder(FilesPath).Files
        Select Case oFso.GetExtensionName(oItem.path)
        Case c_strObjExtBas, c_strObjExtCls
        'Case c_strObjExtDoc ' модуль класса документа
#If APPTYPE = 0 Then        ' APPTYPE=Access
        Case c_strObjExtFrm, c_strObjExtRep ': GoTo HandleNextFile
        'Case c_strObjExtDoc ' модуль класса документа Access (Form или Report)
        Case c_strObjExtQry, c_strObjExtMac ': GoTo HandleNextFile
        ' пропускаем файлы с данными
        Case c_strObjExtXml:    If WithoutData Then GoTo HandleNextFile  ' локальная таблица в XML
        Case c_strObjExtTxt:    If WithoutData Then GoTo HandleNextFile  ' локальная таблица в TXT
        Case c_strObjExtCsv:    If WithoutData Then GoTo HandleNextFile  ' локальная таблица в CSV (текст с разделителями)
        Case c_strObjExtLnk:    If WithoutData Then GoTo HandleNextFile  ' связанная таблица
'#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
#Else
#End If                     ' APPTYPE
        Case Else:              GoTo HandleNextFile
        End Select
' при необходимости проверяем существует ли соответствующий объект в приложении
Dim ObjectName As String:    ObjectName = p_TextCode2Alpha(oFso.GetBaseName(oItem)) ' имя объекта из имени файла
        If OnlyExisting Then If Not ObjectExists(ObjectName) Then GoTo HandleNextFile
' добавляем пути в коллекцию файлов для восстановления
        colObjects.Add oItem, ObjectName
HandleNextFile:
    Next oItem
    Set oItem = Nothing
    Result = True
HandleExit:  p_ObjectFilesCollectionCreate = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_ObjectsBackup(colObjects As Collection, BackupPath As String, prg As clsProgress, _
    Optional WriteType As ObjectRwType = orwAlways, Optional AskBefore As Boolean = True, Optional UseTypeFolders As Boolean = True, _
    Optional Message As String) As Long
' сохранение объектов указанной коллекции в файлы по указанному пути
Const c_strProcedure = "p_ObjectsBackup"
Dim Result As Long
On Error GoTo HandleError
Dim oItem
Dim ParentPath As String, strFilePath As String, strFileExtn As String
Dim strObjName As String, strTypeName As String, strTypeDesc As String
Dim strSkip As String ', strDone As String
Dim strSkipByUser As String, strSkipByVers As String ', strSkipByList As String
' инициализация прогрессбара
Dim strCaption As String, strMessage As String
'Dim i As Long, iMax As Long: i = 1: iMax = colObjects.Count
'Dim prg As clsProgress: Set prg = New clsProgress
'    strCaption = "Сохранение объектов проекта"
'    prg.Init pCount:=1, pMin:=i, pMax:=iMax, pCaption:=strCaption, pText:=strCaption & " в: """ & BackupPath & """": prg.ProgressStep = 1

Dim iCount As Long
' выгружаем поочерёдно все объекты
    For Each oItem In colObjects
        strObjName = oItem.Name
    ' получаем информацию об объекте
        p_ObjectInfo strObjName, ObjectTypeName:=strTypeName, ObjectTypeDesc:=strTypeDesc
    ' обновляем прогрессбар
        strMessage = "Идёт сохранение объекта: " & strTypeDesc & " """ & strObjName & """":
        prg.Update: prg.Detail = strMessage
    ' формируем путь сохранения объекта
        strFilePath = BackupPath: If UseTypeFolders Then strFilePath = oFso.BuildPath(strFilePath, strTypeName)
    ' сохраняем объект по указанному пути
        If Not oFso.FolderExists(strFilePath) Then Call oFso.CreateFolder(strFilePath) 'Then Err.Raise 76 ' Path not Found
    ' проверяем результат сохранения и формируем текст сообщения пользователю
        Select Case p_ObjectWrite(strObjName, strFilePath, WriteType:=WriteType, AskBefore:=AskBefore)  ', Message:=strMessage)
        Case 0:     'strDone = strDone & c_strInDelim & """" & strObjName & """"  ' объект успешно обработан
                    colObjects.Remove (strObjName): iCount = iCount + 1 ' удаляем из коллекции обработанный объект
'        ' причины пропуска:
        Case m_errSkippedByUser ' Пропущено по решению пользователя
            If Len(strObjName) > 0 Then strSkipByUser = strSkipByUser & c_strInDelim & """" & strObjName & """"   ' объект пропущен
        Case m_errWrongVersion  ' Несоответствие версии условию сохранения
            If Len(strObjName) > 0 Then strSkipByVers = strSkipByVers & c_strInDelim & """" & strObjName & """"   ' объект пропущен
        Case m_errSkippedByList ' Наличие в списке пропуска
            'If Len(strObjName) > 0 Then strSkipByList = strSkipByList & c_strInDelim & """" & strObjName & """"   ' объект пропущен
'        Case m_errDestMissing   ' Отсутствeет обновляемый объект
'        Case m_errDestExists    ' Загружаемый объект уже существует
        Case Else:  If Len(strObjName) > 0 Then strSkip = strSkip & c_strInDelim & """" & strObjName & """"   ' объект пропущен
        End Select
    ' проверяем признак прерывания процесса
        If prg.Canceled Then If MsgBox(c_strBreakProcessMessage, vbYesNo Or vbExclamation Or vbDefaultButton2) = vbYes Then GoTo HandleExit
'    ' выводим в прогрессбар информацию о завершении операции
'        prg.Detail = strMessage
    Next oItem
' итоговое сообщение
    Message = "Сохранение объектов проекта завершено."
'    If Left(strDone, Len(c_strInDelim)) = c_strInDelim Then strDone = Mid(strDone, Len(c_strInDelim) + 1)
'    If Len(strDone) > 0 Then Message = Message & vbCrLf & "Сохранены следующие объекты: " & strDone
    Message = Message & vbCrLf & "Всего сохранено " & iCount & " объектов." & vbCrLf 'в: """ & BackupPath & """" & vbCrLf
    If colObjects.Count = 0 Then GoTo HandleExit
    Message = Message & vbCrLf & "Пропущены следующие объекты: "
    If Left(strSkipByUser, Len(c_strInDelim)) = c_strInDelim Then strSkipByUser = Mid(strSkipByUser, Len(c_strInDelim) + 1)
    If Len(strSkipByUser) > 0 Then Message = Message & vbCrLf & ChrW(&H2022) & " по решению пользователя: " & vbCrLf & strSkipByUser
    If Left(strSkipByVers, Len(c_strInDelim)) = c_strInDelim Then strSkipByVers = Mid(strSkipByVers, Len(c_strInDelim) + 1)
    If Len(strSkipByVers) > 0 Then Message = Message & vbCrLf & ChrW(&H2022) & " в связи с несоответстивем версии условию сохранения: " & vbCrLf & strSkipByVers
'    If Left(strSkipByList, Len(c_strInDelim)) = c_strInDelim Then strSkipByList = Mid(strSkipByList, Len(c_strInDelim) + 1)
'    If Len(strSkipByList) > 0 Then Message = Message & vbCrLf & ChrW(&h2022) & " могут быть сохранены, но были пропущены т.к. используются в процессе обновления: "  & vbCrLf & strSkipByList & vbcrLf & "При необходимости их можно сохранить вручную."
    If Left(strSkip, Len(c_strInDelim)) = c_strInDelim Then strSkip = Mid(strSkip, Len(c_strInDelim) + 1)
    If Len(strSkip) > 0 Then Message = Message & vbCrLf & ChrW(&H2022) & " по иным причинам: " & vbCrLf & strSkip
    Message = Message & vbCrLf & "Всего пропущено " & colObjects.Count & " объектов."
HandleExit:     Set oItem = Nothing
                p_ObjectsBackup = Result: Exit Function
HandleError:
'    Dim Message As String
    Err.Clear: Resume 0
    Select Case Err.Number
    Case 75:    Message = "Ошибка доступа!" ' возможно папка уже создана
    If Len(strFilePath) > 0 Then Message = Message & " strFilePath=""" & strFilePath & """ "
    Case 76:    Message = "Путь не найден!" ' возможно нет strFilePath папки
    If Len(strFilePath) > 0 Then Message = Message & " BackupPath=""" & strFilePath & """ "
    Case 1004 '??? ' Ошибка: Программный доступ к проекту Visual Basic не является доверенным
                Message = "Программный доступ к проекту Visual Basic не является доверенным. Для возможности программного сохранения/восстановления модулей необходимо установить разрешение: ""Сервис\Макрос\Безопасность\Доверять доступ к Visual Basic Project"""
    Case Else:  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Result = Err: Err.Clear: Resume HandleExit
End Function
Private Function p_ObjectsRestore(colObjects As Collection, prg As clsProgress, _
    Optional ReadType As ObjectRwType = orwAlways, Optional AskBefore As Boolean = True, Optional UseTypeFolders As Boolean = True, _
    Optional Message As String) As Long
' восстановление объектов из файлов указанной коллекции в проект
Const c_strProcedure = "p_ObjectsRestore"
Dim Result As Long
On Error GoTo HandleError
Dim oItem
Dim strCaption As String, strMessage As String
Dim ParentPath As String, strFilePath As String, strFileExtn As String
Dim strObjName As String, strTypeName As String, strTypeDesc As String
Dim strSkip As String ', strDone As String
Dim strSkipByUser As String, strSkipByVers As String, strSkipByList As String
'' инициализация прогрессбара
'Dim i As Long, iMax As Long: i = 1: iMax = colObjects.Count
'Dim prg As clsProgress: Set prg = New clsProgress
'    strCaption = "Восстановление объектов проекта"
'    prg.Init pCount:=1, pMin:=i, pMax:=iMax, pCaption:=strCaption, pText:=strCaption & " в: """ & SourcePath & """": prg.ProgressStep = 1

Dim iCount As Long
' загружаем поочерёдно все объекты
    For Each oItem In colObjects
        strFilePath = oItem.path: strObjName = strFilePath
    ' получаем информацию об объекте
        p_ObjectInfo strObjName, ObjectTypeName:=strTypeName, ObjectTypeDesc:=strTypeDesc
    ' обновляем прогрессбар
        strMessage = "Идёт восстановление объекта: " & strTypeDesc & " """ & strObjName & """": prg.Update: prg.Detail = strMessage
    ' проверяем результат восстановления и формируем текст сообщения пользователю
        Select Case p_ObjectRead(strFilePath, ReadType:=ReadType, AskBefore:=AskBefore)   ', Message:=strMessage)
        Case 0:     'strDone = strDone & c_strInDelim & """" & strObjName & """"  ' объект успешно обработан
                    colObjects.Remove (strObjName): iCount = iCount + 1  ' удаляем из коллекции обработанный объект
'        ' причины пропуска:
        Case m_errSkippedByUser ' Пропущено по решению пользователя
            If Len(strObjName) > 0 Then strSkipByUser = strSkipByUser & c_strInDelim & """" & strObjName & """"
        Case m_errWrongVersion  ' Несоответствие версии условию обновления
            If Len(strObjName) > 0 Then strSkipByVers = strSkipByVers & c_strInDelim & """" & strObjName & """"
        Case m_errSkippedByList ' Наличие в списке пропуска
            If Len(strObjName) > 0 Then strSkipByList = strSkipByList & c_strInDelim & """" & strObjName & """"
'        Case m_errDestMissing   ' Отсутствует обновляемый объект
'        Case m_errDestExists    ' Загружаемый объект уже существует
        Case Else:  If Len(strObjName) > 0 Then strSkip = strSkip & c_strInDelim & """" & strObjName & """"   ' объект пропущен
        End Select
    ' проверяем признак прерывания процесса
        If prg.Canceled Then If MsgBox(c_strBreakProcessMessage, vbYesNo Or vbExclamation Or vbDefaultButton2) = vbYes Then GoTo HandleExit
    ' выводим в прогрессбар информацию о завершении операции
        prg.Detail = strMessage
    Next oItem
' итоговое сообщение
    Message = "Обновление объектов проекта завершено."
    Message = Message & vbCrLf & "Всего обновлено " & iCount & " объектов." & vbCrLf ' из: """ & RestorePath & """" & vbCrLf
'    If Left(strDone, Len(c_strInDelim)) = c_strInDelim Then strDone = Mid(strDone, Len(c_strInDelim) + 1)
'    If Len(strDone) > 0 Then Message = Message & vbCrLf & "Обновлены следующие объекты: " & strDone
    If colObjects.Count = 0 Then GoTo HandleExit
    Message = Message & vbCrLf & "Пропущены следующие объекты: "
    If Left(strSkipByUser, Len(c_strInDelim)) = c_strInDelim Then strSkipByUser = Mid(strSkipByUser, Len(c_strInDelim) + 1)
    If Len(strSkipByUser) > 0 Then Message = Message & vbCrLf & ChrW(&H2022) & " по решению пользователя: " & vbCrLf & strSkipByUser
    If Left(strSkipByVers, Len(c_strInDelim)) = c_strInDelim Then strSkipByVers = Mid(strSkipByVers, Len(c_strInDelim) + 1)
    If Len(strSkipByVers) > 0 Then Message = Message & vbCrLf & ChrW(&H2022) & " в связи с несоответстивем версии условию обновления: " & vbCrLf & strSkipByVers
    If Left(strSkipByList, Len(c_strInDelim)) = c_strInDelim Then strSkipByList = Mid(strSkipByList, Len(c_strInDelim) + 1)
    If Len(strSkipByList) > 0 Then Message = Message & vbCrLf & ChrW(&H2022) & " могут быть обновлены, но были пропущены т.к. используются в процессе обновления: " & vbCrLf & strSkipByList & vbCrLf & "При необходимости их можно обновить вручную."
    If Left(strSkip, Len(c_strInDelim)) = c_strInDelim Then strSkip = Mid(strSkip, Len(c_strInDelim) + 1)
    If Len(strSkip) > 0 Then Message = Message & vbCrLf & ChrW(&H2022) & " по иным причинам: " & vbCrLf & strSkip
    Message = Message & vbCrLf & "Всего пропущено " & colObjects.Count & " объектов."
HandleExit:     Set oItem = Nothing
                p_ObjectsRestore = Result: Exit Function
HandleError:
'    Dim Message As String
    Err.Clear: Resume 0
    Select Case Err.Number
    Case 75:    Message = "Ошибка доступа!" ' возможно папка уже создана
    If Len(strFilePath) > 0 Then Message = Message & " strFilePath=""" & strFilePath & """ "
    Case 76:    Message = "Путь не найден!" ' возможно нет родительской папки
    If Len(strFilePath) > 0 Then Message = Message & " strFilePath=""" & strFilePath & """ "
    Case 1004 '??? ' Ошибка: Программный доступ к проекту Visual Basic не является доверенным
                Message = "Программный доступ к проекту Visual Basic не является доверенным. Для возможности программного сохранения/восстановления модулей необходимо установить разрешение: ""Сервис\Макрос\Безопасность\Доверять доступ к Visual Basic Project"""
    Case Else:  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Result = Err: Err.Clear: Resume HandleExit
End Function

Public Function ReferencesRestore(Optional FilePath As String)
' восстановление ссылок на библиотеки
Const c_strProcedure = "ReferencesRestore"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
'    p_ReferencesDropBroken
'    p_ReferencesRestore

    Call SysCmd(504, 16484) ' сохраняем все модули
    ' проверяем наличие указанного пути/файла
    Result = oFso.FileExists(FilePath): If Result Then GoTo HandleUpdate
    FilePath = oFso.BuildPath(CurrentProject.path, c_strSrcPath) '& "\" & strText
    ' запрашиваем имя файла backup
Dim strTitle As String: strTitle = "Выберите файл свойств проекта"
Dim strText As String: strText = "Архивная копия свойств проекта " & VBA.Chr$(0) & strText & "*." & c_strObjExtZip & ";" & "*." & c_strObjExtPrj
    FilePath = p_SelectFile(FilePath, strText, c_strObjExtZip & ";" & c_strObjExtPrj, strTitle)
    Result = Len(FilePath) > 0: If Result Then GoTo HandleUpdate
    MsgBox Prompt:="Не указано имя файла архивной копии." & vbCrLf & _
        "Данные не были восстановлены.", Buttons:=vbOKOnly + vbInformation
    GoTo HandleExit
HandleUpdate:
    ' проверяем расширение файла
    Select Case oFso.GetExtensionName(FilePath)
    Case c_strObjExtPrj ' файл свойств проекта
    ' читаем ссылки из файла свойств проекта
        Result = p_ReferencesRead(FilePath)
    Case c_strObjExtZip ' файл архива проекта
    ' извлекаем только файл свойств проекта
Dim TempPath As String, FileName As String
    ' создаем временную папку для извлечения
        FileName = "~" & oFso.GetFileName(FilePath) 'oFso.GetBaseName(FilePath) & "_" & VBA.Format$(Now(), "YYYYMMDD_hhmmss")
        TempPath = oFso.BuildPath(VBA.Environ$("Temp"), FileName)
        If Not oFso.FolderExists(TempPath) Then Call oFso.CreateFolder(TempPath) 'Then Err.Raise 76 ' Path not Found
    ' извлекаем искомый файл и заменяем путь
        FileName = c_strAppNamPrj & "." & c_strObjExtPrj
        oApp.Namespace((TempPath)).CopyHere ((oFso.BuildPath(FilePath, FileName))) ': DoEvents: DoEvents
    ' читаем ссылки из файла свойств проекта
        Result = p_ReferencesRead(oFso.BuildPath(TempPath, FileName))
    ' удаляем временную папку после извлечения
        oFso.DeleteFolder (TempPath)
    Case Else: Err.Raise 75
    End Select
HandleExit:     ReferencesRestore = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 75:                    Message = "Ошибка доступа! Путь: " & FilePath ' Path/File access error
    Case 76:                    Message = "Путь не найден! Путь: " & FilePath ' Path/File access error
    Case Else:                  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function ReferencesBackup(Optional FilePath As String)
' сохранение набора текущих библиотек в тексте функции восстановления ссылок на библиотеки
Const c_strProcedure = "ReferencesBackup"
Dim Result As Boolean: Result = False
On Error GoTo HandleError
    With oFso
    If Len(FilePath) = 0 Then
        FilePath = .BuildPath(CurrentProject.path, c_strSrcPath)
        FilePath = .BuildPath(FilePath, Split(CurrentProject.Name, ".")(0) & "_" & VBA.Format$(Now(), "YYYYMMDD_hhmmss") & "." & c_strObjExtPrj)
    End If
    If Not .FolderExists(.GetParentFolderName(FilePath)) Then Call .CreateFolder(.GetParentFolderName(FilePath)) 'Then Err.Raise 76 ' Path not Found
    End With
' сохраняем в файл
    'Result = p_PropertiesWrite(FilePath)
    Result = p_ReferencesWrite(FilePath)
HandleExit:     ReferencesBackup = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 75:                    Message = "Ошибка доступа! Путь: " & FilePath ' Path/File access error
    Case 76:                    Message = "Путь не найден! Путь: " & FilePath ' Path/File access error
    Case Else:                  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Sub ReferencesPrint()
' отладочная - выводит в Immediate все библиотечные ссылки текущего проекта
    Dim Itm As Object
    Debug.Print "Project """ & VBE.ActiveVBProject.Name & """ References:"
    For Each Itm In References
        With Itm
            If .BuiltIn Then GoTo HandleNext
            Debug.Print " " & IIf(.IsBroken, "X", " ") & .Name, .GUID & " " & " " & .FullPath ' & .Major & " " & .Minor
        End With
HandleNext:
    Next Itm
End Sub
Public Function PropertiesBackup(Optional FilePath As String)
' сохранение свойств проекта
Const c_strProcedure = "PropertiesBackup"
Dim Result As Boolean
' создаем имя файла
    With oFso
    If Len(FilePath) = 0 Then
        FilePath = .BuildPath(CurrentProject.path, c_strSrcPath)
        FilePath = .BuildPath(FilePath, Split(CurrentProject.Name, ".")(0) & "_" & VBA.Format$(Now(), "YYYYMMDD_hhmmss") & "." & c_strObjExtPrj)
    End If
    If Not .FolderExists(.GetParentFolderName(FilePath)) Then Call .CreateFolder(.GetParentFolderName(FilePath)) 'Then Err.Raise 76 ' Path not Found
    End With
' сохраняем свойства
    Result = p_PropertiesWrite(FilePath)
    'Result = p_ReferencesWrite(FilePath)
HandleExit:     PropertiesBackup = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 75:                    Message = "Ошибка доступа! Путь: " & FilePath ' Path/File access error
    Case 76:                    Message = "Путь не найден! Путь: " & FilePath ' Path/File access error
    Case Else:                  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function PropertiesRestore(Optional FilePath As String)
' восстановление свойств проекта
Const c_strProcedure = "PropertiesRestore"
Dim Result As Boolean ': Result = False
    On Error GoTo HandleError
    
    Call SysCmd(504, 16484) ' сохраняем все модули
    ' проверяем наличие указанного пути/файла
    Result = oFso.FileExists(FilePath): If Result Then GoTo HandleUpdate
    FilePath = oFso.BuildPath(CurrentProject.path, c_strSrcPath) '& "\" & strText
    ' запрашиваем имя файла backup
Dim strTitle As String: strTitle = "Выберите файл свойств проекта"
Dim strText As String: strText = "Архивная копия свойств проекта " & VBA.Chr$(0) & strText & "*." & c_strObjExtZip & ";" & "*." & c_strObjExtPrj
    FilePath = p_SelectFile(FilePath, strText, c_strObjExtZip & ";" & c_strObjExtPrj, strTitle)
    Result = Len(FilePath) > 0: If Result Then GoTo HandleUpdate
    MsgBox Prompt:="Не указано имя файла архивной копии." & vbCrLf & _
        "Данные не были восстановлены.", Buttons:=vbOKOnly + vbInformation
    GoTo HandleExit
HandleUpdate:
    ' проверяем расширение файла
    Select Case oFso.GetExtensionName(FilePath)
    Case c_strObjExtPrj ' файл свойств проекта
    ' читаем свойства из файла свойств проекта
        Result = p_PropertiesRead(FilePath)
    Case c_strObjExtZip ' файл архива проекта
    ' извлекаем только файл свойств проекта
Dim TempPath As String, FileName As String
    ' создаем временную папку для извлечения
        FileName = "~" & oFso.GetFileName(FilePath) 'oFso.GetBaseName(FilePath) & "_" & VBA.Format$(Now(), "YYYYMMDD_hhmmss")
        TempPath = oFso.BuildPath(VBA.Environ$("Temp"), FileName)
        If Not oFso.FolderExists(TempPath) Then Call oFso.CreateFolder(TempPath) 'Then Err.Raise 76 ' Path not Found
    ' извлекаем искомый файл и заменяем путь
        FileName = c_strAppNamPrj & "." & c_strObjExtPrj
        oApp.Namespace((TempPath)).CopyHere ((oFso.BuildPath(FilePath, FileName))) ': DoEvents: DoEvents
    ' читаем свойства из файла свойств проекта
        Result = p_PropertiesRead(oFso.BuildPath(TempPath, FileName))
    ' удаляем временную папку после извлечения
        oFso.DeleteFolder (TempPath)
    Case Else: Err.Raise 75
    End Select
    
HandleExit:     PropertiesRestore = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 75:                    Message = "Ошибка доступа! Путь: " & FilePath ' Path/File access error
    Case 76:                    Message = "Путь не найден! Путь: " & FilePath ' Path/File access error
    Case Else:                  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Sub PropertiesClear()
' удаляет все пользовательские свойства
    With CurrentProject.Properties: Do While .Count > 0: .Remove .Item(0).Name: Loop: End With
End Sub
Public Sub PropertiesPrint()
' отладочная - выводит в Immediate все свойства текущего проекта
    Dim Itm As Object
    Debug.Print "Project """ & VBE.ActiveVBProject.Name & """ Properties:"
    For Each Itm In CurrentProject.Properties
        Debug.Print Itm.Name & "=" & Itm.Value
    Next Itm
End Sub
Public Function PropertyGet(PropName As String, PropValue As Variant, Optional PropObject As Object) As Boolean
' читает свойство произвольного объекта
Const c_strProcedure = "PropertyGet"
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
HandleExit:  PropertyGet = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function PropertySet(PropName As String, PropValue As Variant, Optional PropObject As Object, Optional PropType As eDataType = dbText) As Boolean
' добавляет пользовательское свойство к объекту DAO или AccessObject
Const c_strProcedure = "PropertySet"
' PropName      - имя свойства
' PropValue     - значение свойства
' PropObject    - объект к которому добавляется свойство
' PropType      - тип данных свойства
Dim Result As Boolean
    On Error Resume Next
    If PropObject Is Nothing Then Set PropObject = CurrentProject ' по-умолчанию
    ' пытаемся записать свойство
    PropObject.Properties(PropName) = PropValue
    Select Case Err.Number
    Case 0: Result = True: GoTo HandleExit
    Case 3270, 2455: ' Свойство не найдено
    Case Else: On Error GoTo HandleExit: Err.Raise Err.Number
    End Select
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
        Err.Raise 438 ' Object doesn't support this property or method ' vbObjectError + 512
    End If
    Result = True
HandleExit:     PropertySet = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_IsSkippedObject(ObjectName) As Boolean
' проверяет принадлежит ли объект к пропускаемым при экспорте
Dim Result As Boolean:   ' Result = False
On Error GoTo HandleError
Dim Arr, i As Long ', iMax As Long
    Arr = Split(c_strObjIgnore, ";")
    For i = LBound(Arr) To UBound(Arr)
        Result = (VBA.UCase$(Arr(i)) = VBA.UCase$(ObjectName))
        If Result Then Exit For
    Next i
HandleExit:  p_IsSkippedObject = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
'-------------------------
' функции обслуживания проекта
'-------------------------
Public Function CompileAll()
' Вызов скрытой функции SysCmd для автоматической компиляции/сохранения модулей.
'Enum accSaveVbaCode
'    accSaveVBAUnkn01 = 16481
'    accSaveVBAUnkn02 = 16482
'    accSaveVBAwCode = 16483 'Save VBA with compiled code
'    accSaveVBAwoCode = 16484 'Save VBA without compiled code
'End Enum
'    For i = 0 To Modules.Count - 1
'        UpdateModule Modules(i).Name
'    Next i
    If Not IsCompiled Then Call SysCmd(504, 16483)
End Function
Public Function CloseAll( _
    Optional OnlyInDesignMode As Boolean = False)
' закрывает открытые объекты Access
Const c_strProcedure = "CloseAll"
#If APPTYPE = 0 Then        ' APPTYPE=Access
Dim i As Byte
Dim oColl, oItem, eObjType As AcObjectType
    On Error GoTo HandleError
'    If SysCmd(714) Then ' не всегда правильно работает
    ' проверяем есть ли объекты открытые в режиме Конструктора
    
    ' закрываем все открытые объекты
    For eObjType = acTable To acModule
        Select Case eObjType
        Case acTable:   Set oColl = CurrentData.AllTables
        Case acQuery:   Set oColl = CurrentData.AllQueries
        Case acForm:    Set oColl = CurrentProject.AllForms
        Case acReport:  Set oColl = CurrentProject.AllReports
        Case acMacro:   Set oColl = CurrentProject.AllMacros
        Case acModule:  Set oColl = CurrentProject.AllModules
        Case Else:      GoTo HandleNextType
        End Select
        For Each oItem In oColl
            With oItem
                If .IsLoaded Then
                    If OnlyInDesignMode And .CurrentView <> 0 Then GoTo HandleNextItem
                    DoCmd.Close eObjType, .Name, acSaveYes
'Debug.Print IIf(.IsLoaded, "Can't close object", "Object was closed") & ": """ & oItem.Name & """."
                End If
            End With
HandleNextItem:
        Next oItem
HandleNextType:
    Next eObjType
#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
#Else                       ' APPTYPE=Неразбери пойми что
#End If                     ' APPTYPE
'    End If
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
#If APPTYPE = 0 Then        ' APPTYPE=Access
Public Function TablesReLink(DatabasePath As String)
' обновление путей к связанным таблицам
Const c_strProcedure = "TablesReLink"
Const c_ConnString = ";DATABASE="
Dim tdf As Object 'TableDef
Dim i As Long, iMax As Long
Dim lngBrokenLinks As Long
Dim strCaption As String
Dim prg As clsProgress

    On Error GoTo HandleError
    i = 0:    iMax = CurrentDb.TableDefs.Count
    strCaption = "Обновление связей таблиц"
'    SysCmd acSysCmdInitMeter, "Обновление связей таблиц " & DatabasePath, _
    Set prg = New clsProgress
    prg.Init pCount:=1, pMin:=i, pMax:=iMax, pCaption:=strCaption, pText:=strCaption & " с: """ & DatabasePath & """"
    lngBrokenLinks = 0
    On Error Resume Next
    For Each tdf In CurrentDb.TableDefs
        If Len(tdf.Connect) > 0 Then
            If VBA.Mid$(tdf.Connect, Len(c_ConnString) + 1) <> DatabasePath Then
                With prg
                    If .Canceled Then If MsgBox(c_strBreakProcessMessage, vbYesNo Or vbExclamation Or vbDefaultButton2) = vbYes Then GoTo HandleExit
                    .Canceled = False
                    .Detail = strCaption & "ы: " & tdf.Name
                    .Progress = i
                End With
                tdf.Connect = c_ConnString & DatabasePath
                tdf.RefreshLink
                lngBrokenLinks = lngBrokenLinks + 1
            End If
        End If
        i = i + 1
'        SysCmd acSysCmdUpdateMeter, i
    Next tdf
HandleExit:
    Set prg = Nothing
    If lngBrokenLinks > 0 Then MsgBox strCaption & " c:" & vbCrLf & DatabasePath & vbCrLf & "завершено.", vbOKOnly, "Операция завершена!"
'    SysCmd (acSysCmdClearStatus)
    CurrentDb.Close
    Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
#End If                     ' APPTYPE

'-------------------------
' функции обновления версий модулей/функций
'-------------------------
Public Function UpdateModule( _
    ObjectName As String, _
    Optional Comment As String, _
    Optional VerType As appVerType = appVerBuild, _
    Optional SkipDialog As Boolean = False)
' обновляет информацию о версии указанного модуля
Const c_strProcedure = "UpdateModule"
Dim strBegLineMarker As String
Dim BegLine As Long, EndLine As Long
Dim VerDate As Date
Dim strVersion As String, strVerShort As String
Dim strValue As String, strComment As String
Dim IsLoaded As Boolean
Dim Result As Boolean
    Result = False
    On Error GoTo HandleError
DoCmd.Echo False
    If ObjectName = vbNullString Then ObjectName = InputBox("Введите имя обновляемого объекта:", , ObjectName)
Dim ModuleName As String: ModuleName = ObjectName
Dim ObjectType As AcObjectType
' проверяем введенное имя
    If ModuleName = vbNullString Then
        Err.Raise vbObjectError + 512
'    ElseIf ModuleName = VBE.ActiveCodePane.CodeModule Then
'        Err.Raise vbObjectError + 513
    ElseIf Not IsModuleExists(ModuleName, ObjectName, ObjectType) Then
        Err.Raise vbObjectError + 514
    'ElseIf Not IsFuncExists(ModuleName) Then
    '    Result = Update Func ... : Goto HandleExit
    End If
' открываем объект в режиме дизайна
    Select Case ObjectType
    Case acModule:  'Do Nothing
    Case acForm:    DoCmd.OpenForm ObjectName, acDesign
    Case acReport:  DoCmd.OpenReport ObjectName, acDesign
    Case Else:      Err.Raise vbObjectError + 514
    End Select
    ' если модуль открыт - сохраняем
    'DoCmd.Save acModule, ModuleName ': DoCmd.Close acModule, ModuleName, acSaveYes
    ' получаем текущую версию модуля
    strVersion = ModuleVersGet(ModuleName)
    ' увеличиваем версию модуля
    VersionInc strVersion, VerShort:=strVerShort, VerDate:=VerDate, IncType:=VerType
    If SkipDialog Then
        strComment = Comment
    Else
        strValue = InputBox("Укажите новый номер версии " & vbCrLf & "модуля " & ModuleName & ":", , strVerShort)
        ' если нажата отмена
        If strValue <> vbNullString Then strVersion = strValue: VersionSet strVersion, VerShort:=strVerShort, VerDate:=VerDate
        strComment = InputBox("Добавьте описание изменений в новой версии" & vbCrLf & "модуля " & ModuleName & ":", , Comment)
    End If
    ModuleVersSet ModuleName, strVersion
    ModuleDateSet ModuleName, CStr(VerDate)
    'ModuleAuthSet ModuleName, Author '& " (" & Support & ")"
' добавляем комментарий к версии
    If Len(strComment) > 0 Then
Dim tmpString As String: tmpString = ModuleAuthGet(ModuleName)
Dim strAuthor As String, strSupport As String: strAuthor = Author: strSupport = Support
    ' если имя автора изменений не совпадает с именем автора модуля добавляем автора и дату изменений
        Select Case VBA.UCase$(tmpString)
        Case VBA.UCase$(strAuthor), VBA.UCase$(strAuthor) & " (" & VBA.UCase$(strSupport) & ")": tmpString = vbNullString
        Case Else: tmpString = vbNullString
            If Len(strAuthor) > 0 Then tmpString = tmpString & strAuthor
            If Len(strSupport) > 0 Then tmpString = tmpString & " (" & strSupport & ")"
            If Len(tmpString) > 0 Then tmpString = " {" & tmpString & "}"
        End Select
    ' выводим в комментарий номер версии и краткое описание изменений
        tmpString = c_strPrefVerComm & strVerShort & VBA.String$(IIf(c_strPrefLen - Len(strVerShort) - 5 > 0, c_strPrefLen - Len(strVerShort) - 5, 1), " ") & ": " & VBA.Format$(Now, "dd.mm.yyyy") & " - " & strComment & tmpString
        ModuleCommSet ModuleName, tmpString
        'DoCmd.Save acModule, ModuleName
    End If
    'ModuleDebugSet ModuleName, c_bolDebugMode
' закрываем объект в режиме дизайна
    DoCmd.Close ObjectType, ObjectName, acSaveYes
' обновляем версию приложения
    ' получаем текущую версию приложения
    Call PropertyGet(c_strPrpKeyVer, strVersion)
    ' обновляем ревизию версии приложения
    VersionInc strVersion, VerShort:=strVerShort, VerDate:=VerDate, IncType:=appVerRevis
    ' сохраняем текущую версию приложения в свойство
    'Call PropertySet(c_strPrpKeyVer, strVersion)
    CurrentProject.Properties(c_strPrpKeyVer) = strVersion
    Result = True
HandleExit:  DoCmd.Echo True: UpdateModule = Result: Exit Function
HandleError:
    Result = False
    Select Case Err.Number
    Case vbObjectError + 512: Debug.Print "Не задано имя модуля!"
    Case vbObjectError + 513: Debug.Print "Не возможно изменить активный модуль: """ & ModuleName & """!"
    Case vbObjectError + 514: Debug.Print "Модуль: """ & ModuleName & """ не найден!"
    Case Else: Debug.Print Err.Description
    End Select
    Err.Clear: Resume HandleExit
End Function
Private Function IsModuleExists( _
    ByRef ModuleName As String, _
    Optional ByRef ObjectName As String, _
    Optional ByRef ObjectType As AcObjectType _
    ) As Boolean
' проверяет наличие указанного модуля
Dim Result As Boolean
Dim strObjName As String
' возвращает True, если есть модуль с таким именем.

    Result = False
    On Error Resume Next
' Application.Modules видит только загруженные модули.
'   cоответственно существующий модуль если он Not IsLoaded будет не найден
' CurrentProject.AllModules видит только модули и не видит модулей форм и отчетов
    With CurrentProject
' проверяем коллекцию модулей проекта
        If (.AllModules(ModuleName).Name = ModuleName) Then ObjectName = ModuleName
        Result = (Err = 0): If Result Then ObjectType = acModule: GoTo HandleExit
        Err.Clear
' проверяем коллекцию форм проекта
        If Left$(ObjectName, Len(c_strFrmModPref)) = c_strFrmModPref Then ObjectName = Mid$(ModuleName, Len(c_strFrmModPref) + 1)
        DoCmd.OpenForm ObjectName, acDesign '1 = acDesign
        ModuleName = Forms(ObjectName).Module.Name
        Result = (Err = 0): If Result Then ObjectType = acForm: GoTo HandleExit
        Err.Clear
' проверяем коллекцию отчетов проекта
        If Left$(ObjectName, Len(c_strFrmModPref)) = c_strRepModPref Then ObjectName = Mid$(ModuleName, Len(c_strFrmModPref) + 1)
        DoCmd.OpenReport ObjectName, acDesign '1 = acDesign
        ModuleName = Forms(ObjectName).Module.Name
        Result = (Err = 0): If Result Then ObjectType = acReport: GoTo HandleExit
        Err.Clear
    End With
    DoCmd.Close ObjectType, ObjectName
HandleExit:     IsModuleExists = Result:    Exit Function
HandleError:    Result = False: Err.Clear:  Resume HandleExit
End Function
Public Sub UpdateFunc( _
    funcName As String, _
    Optional ModuleName As String, _
    Optional VerType As appVerType = appVerBuild, _
    Optional SkipDialog As Boolean = False)
' обновляет информацию о версии функции
Const c_strProcedure = "UpdateFunc"
' v.1.0.0       : 02.04.2019 - исходная версия
Dim BegLine As Long, EndLine As Long
Dim i As Long
Dim Result As Boolean

    Result = False
    On Error GoTo HandleError
    If funcName = vbNullString Then Err.Raise m_errProcNameWrong
    If ModuleName = vbNullString Then
' возможно функция задана в виде clsForm.Form(Set), тогда выделяем имя модуля
        ModuleName = Split(funcName, ".")(0)
        funcName = VBA.Mid$(funcName, Len(ModuleName) + 2)
    End If
' проверяем тип функции
Dim ProcKind As vbext_ProcKind: ProcKind = vbext_pk_Proc
    If VBA.Right$(funcName, 1) = ")" Then
        Select Case VBA.Mid$(funcName, Len(funcName) - 3, 3)
        Case "Let": ProcKind = vbext_pk_Let
        Case "Set": ProcKind = vbext_pk_Set
        Case "Get": ProcKind = vbext_pk_Get
        End Select
        If VBA.Left$(VBA.Right$(funcName, 5), 1) = "(" Then
            funcName = VBA.Left$(funcName, Len(funcName) - 5)
        ElseIf VBA.Left$(VBA.Right$(funcName, 14), 10) = "(Property " Then
            funcName = VBA.Left$(funcName, Len(funcName) - 14)
        Else
            Err.Raise m_errProcNameWrong
        End If
    End If
' проверяем доступность функции и ее позицию
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    With objModule
        BegLine = .ProcBodyLine(funcName, ProcKind) ' маркер на начало функции
        EndLine = BegLine + .ProcCountLines(funcName, ProcKind) - 3
        BegLine = CodeLineNext(ModuleName, BegLine)
' пролучаем информацию о версии функции
Dim strVersion As String, strVerShort As String, VerDate As Date
Dim strComment As String
Dim tmpString As String: tmpString = "'" & c_strPrefVerComm & "*.*.*:"
Dim NxtLine As Long
'        tmpString = c_strPrefVerComm & strVerShort & String$(IIf(c_strPrefLen - Len(strVerShort) - 5 > 0, c_strPrefLen - Len(strVerShort) - 5, 1), " ") & ": " & VBA.format$(Now, "dd.mm.yyyy") & " - " & strComment & tmpString
'        Do
            NxtLine = EndLine
            Result = .Find(tmpString, _
                StartLine:=BegLine, EndLine:=NxtLine, _
                StartColumn:=0, EndColumn:=0, _
                PatternSearch:=True)
            If Result Then
    ' если найдено - извлекаем номер последней версии из функции
                strVersion = VBA.Trim(VBA.Mid$(.Lines(BegLine, 1), Len(c_strPrefVerComm) + 2))
                strVersion = VBA.Trim$(VBA.Left$(strVersion, InStr(strVersion, ":") - 1))
'                Exit Do
'                BegLine = NxtLine + 1
'            Else
'                Exit Do
            End If
'        Loop
    ' увеличиваем версию функции
        VersionInc strVersion, VerShort:=strVerShort, VerDate:=VerDate, IncType:=VerType
        If Not SkipDialog Then
            tmpString = InputBox("Укажите новый номер версии " & vbCrLf & "функции " & ModuleName & "." & funcName & ":", , strVerShort)
            ' если нажата отмена
            If tmpString <> vbNullString Then strVersion = tmpString: VersionSet strVersion, VerShort:=strVerShort, VerDate:=VerDate
            strComment = InputBox("Добавьте описание изменений в новой версии" & vbCrLf & "функции " & ModuleName & "." & funcName & ":")
        End If
    ' добавляем комментарий к версии
    tmpString = ModuleAuthGet(ModuleName)
Dim strAuthor As String, strSupport As String: strAuthor = Author: strSupport = Support
    ' если имя автора изменений не совпадает с именем автора модуля добавляем автора и дату изменений
        Select Case VBA.UCase$(tmpString)
        Case VBA.UCase$(strAuthor), VBA.UCase$(strAuthor) & " (" & VBA.UCase$(strSupport) & ")": tmpString = vbNullString
        Case Else:  tmpString = " (" & strAuthor & " (" & strSupport & ")"
        End Select
    ' выводим в комментарий номер версии и краткое описание изменений
        tmpString = "'" & c_strPrefVerComm & strVerShort & VBA.String$(IIf(c_strPrefLen - Len(strVerShort) - 5 > 0, c_strPrefLen - Len(strVerShort) - 5, 1), " ") & ": " & VBA.Format$(Now, "dd.mm.yyyy") & " - " & strComment & tmpString
        .InsertLines BegLine, tmpString
    ' выводим в комментарий номер версии и краткое описание изменений
        tmpString = "изменения в " & funcName & " - " & strComment
        UpdateModule ModuleName, tmpString, SkipDialog:=True
    End With
HandleExit:     Exit Sub
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "Неверно задано имя объекта!"
    Case m_errModuleIsActive: Debug.Print "Невозможно изменить активный модуль: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "Модуль: """ & ModuleName & """ не найден!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Sub
'-------------------------
' функции для работы с версиями
'-------------------------
Public Function VersionSet( _
    ByRef VerText As String, _
    Optional ByRef VerDate As Date, _
    Optional ByRef VerShort As String _
    ) As Long
' формирует из строки версии полную версию, добавляет
Const c_strProcedure = "VersionSet"
' VerText - исходная строка версии, на выходе содержит полную версию
' VerDate - содержит дату версии для сохранения в Revision
' VerShort- на выходе содержит краткую версию
' функция возвращает числовой код версии в виде AABBB
Dim VerType As typVersion
Dim Result As Long

    Result = False
    On Error GoTo HandleError
    Call p_VersionFromString(VerText, VerType)
    With VerType
        ' VerText - исходная строка версии, на выходе содержит полную версию с учетом указанного сдвига
        ' VerShort- на выходе содержит краткую версию с учетом указанного сдвига
    ' собираем информацию о версии
        Result = .VerCode
        .Revision = VersionDate2Rev(VerDate): .VerDate = VerDate
        VerShort = .Major & c_strVerDelim & .Minor & c_strVerDelim & .Build
        VerText = VerShort & c_strVerDelim & .Revision
    End With
HandleExit:     VersionSet = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Public Function VersionGet( _
    ByRef VerText As String, _
    Optional ByRef VerShort As String, _
    Optional ByRef VerDate As Date, _
    Optional RelType As Byte = 1 _
    ) As Long
' получает из строки версии ее компоненты
Const c_strProcedure = "VersionGet"
' VerText - исходная строка версии, на выходе содержит полную версию
' VerShort- на выходе содержит краткую версию
' VerDate - на выходе содержит дату версии полученную из расшифровки Revision
' RelType - тип возвращаемой информации о релизе 0 - код, 1 - краткое описание, 3 - полное описание
' функция возвращает числовой код версии в виде AABBB
Dim VerType As typVersion
Dim Result As Long

    Result = False
    On Error GoTo HandleError
    Call p_VersionFromString(VerText, VerType)
    With VerType
        ' VerText - исходная строка версии, на выходе содержит полную версию с учетом указанного сдвига
        ' VerShort- на выходе содержит краткую версию с учетом указанного сдвига
    ' собираем информацию о версии
        Result = .VerCode: VerDate = .VerDate
        VerShort = .Major & c_strVerDelim & .Minor & c_strVerDelim & .Build
        VerText = VerShort & c_strVerDelim & .Revision
    ' добавляем информацию о релизе
        Select Case RelType
        Case 0: VerText = VerText & c_strVerDelim & .Release: If .RelSubNum > 0 Then VerText = VerText & " " & .RelSubNum
        Case 1: If .Release > 0 Then VerText = VerText & " " & .RelShort: If .RelSubNum > 0 Then VerText = VerText & .RelSubNum
        Case 2: If .Release > 0 Then VerText = VerText & " " & .RelFull: If .RelSubNum > 0 Then VerText = VerText & " " & .RelSubNum
        End Select
    End With
HandleExit:     VersionGet = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Public Function VersionInc( _
    ByRef VerText As String, _
    Optional ByRef VerShort As String, _
    Optional ByRef VerDate As Date, _
    Optional IncStep As Integer = 1, _
    Optional IncType As appVerType = appVerBuild, _
    Optional RelType As Byte = 1 _
    ) As Long
' изменяет номер версии заданной строкой и возвращает ее числовое значение
Const c_strProcedure = "VersionInc"
' VerText - исходная строка версии, на выходе содержит полную версию с учетом указанного сдвига
' VerShort- на выходе содержит краткую версию с учетом указанного сдвига
' VerDate - на выходе содержит дату версии полученную из расшифровки Revision
' IncStep - количество единиц изменения разряда версии (+-)
' IncType - номер изменяемого разряда версии
' RelType - тип возвращаемой информации о релизе 0 - код, 1 - краткое описание, 3 - полное описание
' функция возвращает числовой код версии в виде AABBB
Dim VerType As typVersion
Dim Result As Long

    Result = False
    On Error GoTo HandleError
' Разбираем текущую версию
    Call p_VersionFromString(VerText, VerType): If IncStep = 0 Then GoTo HandleExit
    With VerType
' Обновляем текущую версию
    ' ревизию обновляем при каждом обновлении независимо от заданного IncType
        .Revision = VersionDate2Rev(Now): VerDate = VersionRev2Date(.Revision)
        Select Case IncType
        Case appVerMajor: .Major = .Major + IncStep
        Case appVerMinor: .Minor = .Minor + IncStep
        Case appVerBuild: .Build = .Build + IncStep
        Case appRelease:  If .Release > appReleaseNotDefine Then .Release = .Release + IncStep
        Case appRelSubNum:  If .Release > appReleaseNotDefine Then .RelSubNum = .RelSubNum + IncStep
        End Select
    End With
' Пересобираем компоненты
    p_VersionCheck VerType
    With VerType
        ' VerText - исходная строка версии, на выходе содержит полную версию с учетом указанного сдвига
        ' VerShort- на выходе содержит краткую версию с учетом указанного сдвига
    ' собираем информацию о версии
        Result = .VerCode: VerDate = .VerDate
        VerShort = .Major & c_strVerDelim & .Minor & c_strVerDelim & .Build
        VerText = VerShort & c_strVerDelim & .Revision
    ' добавляем информацию о релизе
        Select Case RelType
        Case 0: VerText = VerText & c_strVerDelim & .Release: If .RelSubNum > 0 Then VerText = VerText & " " & .RelSubNum
        Case 1: If .Release > 0 Then VerText = VerText & " " & .RelShort: If .RelSubNum > 0 Then VerText = VerText & .RelSubNum
        Case 2: If .Release > 0 Then VerText = VerText & " " & .RelFull: If .RelSubNum > 0 Then VerText = VerText & " " & .RelSubNum
        End Select
    End With
HandleExit:     VersionInc = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Public Function VersionCmp(Ver1 As String, Ver2 As String) As Integer
' сравнивает версии
Const c_strProcedure = "VersionCmp"
' возвращает:
' 0 - если Ver1=Ver2
' 1 - если Ver1>Ver2
'-1 - если Ver1<Ver2
'-32768 - если ошибка
Dim Result As Integer
    On Error GoTo HandleError
Dim v1 As typVersion: If Not p_VersionFromString(Ver1, v1) Then Err.Raise vbObjectError + 512
Dim v2 As typVersion: If Not p_VersionFromString(Ver2, v2) Then Err.Raise vbObjectError + 512
    If v1.Revision > 0 And v2.Revision > 0 Then Result = Sgn(v1.Revision - v2.Revision): GoTo HandleExit
    Result = Sgn(v1.Major - v2.Major): If Result <> 0 Then GoTo HandleExit
    Result = Sgn(v1.Minor - v2.Minor): If Result <> 0 Then GoTo HandleExit
    Result = Sgn(v1.Build - v2.Build) ': If Result <> 0 Then GoTo HandleExit
HandleExit:     VersionCmp = Result: Exit Function
HandleError:    Result = -32768: Err.Clear: Resume HandleExit
End Function
Public Function VersionDate2Rev(DateTime As Date) As Long: VersionDate2Rev = CCur(DateTime) * 10 ^ c_bytTimeDig: End Function
Public Function VersionRev2Date(RevString As Long) As String: VersionRev2Date = VBA.Format$(CDate(RevString / 10 ^ c_bytTimeDig), "dd.mm.yyyy hh:nn:ss"): End Function
Private Function p_VersionCheck(ByRef VerType As typVersion) As Boolean
' проверяем данные о текущей версии
Const c_strProcedure = "p_VersionCheck"
Dim Result As Boolean

    Result = False
    On Error GoTo HandleError
    With VerType
    ' проверяем диапазоны
        If .Major < 0 Then .Major = 0 Else If .Major > 10 ^ c_bytMajorDig - 1 Then .Major = 10 ^ c_bytMajorDig - 1: Err.Raise 6 'Overflow
        If .Minor < 0 Then .Minor = 0 Else If .Minor > 10 ^ c_bytMinorDig - 1 Then .Minor = 10 ^ c_bytMinorDig - 1: Err.Raise 6 'Overflow
        If .Build < 0 Then .Build = 0
        If .Release < 0 Then .Release = 0: If .Release > appReleaseEOL Then .Release = appReleaseNotDefine: Err.Raise 6 'Overflow
        If .RelSubNum < 0 Then .RelSubNum = 0
        If .Revision > 0 Then
    ' получаем из ревизии дату билда
        Dim tmp As Long: tmp = 10 ^ (Len(CStr(.Revision)) - c_bytDateDig): If tmp = 0 Then tmp = 1
            .VerDate = CDate(.Revision / tmp)
        End If
    ' формируем числовой код версии (AABBB)
        .VerCode = .Major * 10 ^ c_bytMinorDig + .Minor
    End With
    Result = True
HandleExit:     p_VersionCheck = Result: Exit Function
HandleError:    Result = False
    Select Case Err.Number
    Case 6:     Err.Clear: Resume Next ' Overflow
    Case Else:  Err.Clear: Resume HandleExit
    End Select
End Function
Private Function p_VersionFromString( _
    ByRef VerText As String, _
    ByRef VerType As typVersion _
    ) As Boolean
' разбираем текущую версию из строки
Const c_strProcedure = "p_VersionFromString"
' VerText - строка номера версии вида A.B.C.D[r][n]
' VerType - информация о версии разложенная на компоненты
Dim strRel As String
Dim Result As Boolean

    Result = False
    On Error GoTo HandleError
    VerText = VBA.Trim$(VerText)
Dim arrVer: arrVer = Split(VerText, c_strVerDelim)
Dim i As Long, iMax As Long: i = 0: iMax = UBound(arrVer) - LBound(arrVer)
' Разбираем строку версии
    With VerType
        Do While i <= iMax
            Select Case i
            Case appVerMajor: .Major = arrVer(i) ' A – главный номер версии
            Case appVerMinor: .Minor = arrVer(i) ' B – вспомогательный номер версии
            Case appVerBuild: .Build = arrVer(i) ' C – номер билда версии
            Case appVerRevis                     ' D - номер ревизии (здесь-даты билда)
                .Revision = Fix(Val(arrVer(i)))
                ' возможно это номер ревизии + тип релиза, - проверяем и разбираем
                If Not IsNumeric(arrVer(i)) Then strRel = VBA.Mid$(arrVer(i), Len(.Revision) + 1): Exit Do
            Case Else:  strRel = arrVer(i)       '[r]- тип релиза
            End Select
            i = i + 1
        Loop
        If Len(strRel) > 0 Then .Release = p_GetReleaseType(strRel, .RelSubNum, .RelShort, .RelFull)
    End With
' Пересобираем строку версии
    p_VersionCheck VerType
    Result = True
HandleExit:     p_VersionFromString = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_GetReleaseType( _
    ReleaseData As String, _
    ReleaseSubNum As Integer, _
    Optional ReleaseShortDesc As String, _
    Optional ReleaseFullDesc As String _
    ) As appRelType
' возвращает код и описание типа релиза
Const c_strProcedure = "p_GetReleaseType"
' ReleaseData -      исходные данные
' ReleaseSubNum -    подномер релиза для типа релиза (rc1)
' ReleaseShortDesc - краткое описание типа релиза
' ReleaseFullDesc -  полное описание типа релиза
Dim strData As String
Dim Result As appRelType
    On Error GoTo HandleError
    strData = VBA.Trim$(ReleaseData)
    Result = False: ReleaseShortDesc = vbNullString: ReleaseFullDesc = vbNullString
' от конца к началу ищем цифры
    Dim c As String * 1, strNum As String
    Dim i As Long: i = Len(strData)
    Do While i > 0
        c = VBA.Right$(strData, 1)
        Select Case c
        Case 0 To 9: strNum = c & strNum
        Case Else:  Exit Do
        End Select
        i = i - 1: strData = VBA.Left$(strData, i)
    Loop
    strData = VBA.Trim$(strData)
' получаем дополнительный номер типа релиза
    If Len(strNum) > 0 Then ReleaseSubNum = CByte(strNum)
' получаем описание типа релиза
    Select Case strData
    Case 1, "pa", "Pre-alpha":                  Result = 1: ReleaseShortDesc = "pa":  ReleaseFullDesc = "Pre-alpha"
    Case 2, "a", "Alpha":                       Result = 2: ReleaseShortDesc = "a":   ReleaseFullDesc = "Alpha"
    Case 3, "b", "Beta":                        Result = 3: ReleaseShortDesc = "b":   ReleaseFullDesc = "Beta"
    Case 4, "rc", "Release Candidate":          Result = 4: ReleaseShortDesc = "rc":  ReleaseFullDesc = "Release Candidate"
    Case 5, "rtm", "Release to manufacturing":  Result = 5: ReleaseShortDesc = "rtm": ReleaseFullDesc = "Release to manufacturing"
    Case 6, "ga", "General availability":       Result = 6: ReleaseShortDesc = "ga":  ReleaseFullDesc = "General availability"
    Case 7, "eol", "End of life":               Result = 7: ReleaseShortDesc = "eol": ReleaseFullDesc = "End of life"
    End Select
HandleExit:     p_GetReleaseType = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
'----------------------
' функции работы с текстом модулей
'----------------------
Private Function ModuleExists(ByVal ModuleName As String, Optional objModule As Object) As Boolean
' проверяет наличие указанного модуля
Dim Result As Boolean
' возвращает True, если есть модуль с таким именем.
    On Error Resume Next
    Set objModule = Application.VBE.ActiveVBProject.VBComponents(ModuleName).CodeModule
    Result = (Err = 0): Err.Clear
' Application.Modules видит только загруженные модули.
'   cоответственно существующий модуль если он Not IsLoaded будет не найден
HandleExit:     ModuleExists = Result:    Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Public Function ObjectExists(ObjectName As String) As Boolean
' проверяет наличие указанного объекта в приложении
    ObjectExists = p_ObjectInfo(ObjectName)
End Function
Private Function p_ObjectInfo(ObjectName As String, _
    Optional ObjectType As appObjectType, Optional ObjectTypeName, Optional ObjectTypeDesc, _
    Optional ObjectFileName, Optional ObjectFileExt, Optional ObjectModuleName) As Boolean
' возвращает информацию об объекте приложения
Const c_strProcedure = "p_ObjectInfo"
' ObjectName - имя проверяемого объекта
' ObjectType - возвращаемый тип объекта
' ObjectTypeName - название типа объекта (= имя папки в файле бэкапа)
' ObjectTypeDesc - описание типа объекта
' ObjectFileName - имя файла
' ObjectFileExt - расширение для файла
' ObjectModuleName - имя модуля объекта (при наличии)
' возвращает True, если объект с таким именем существует в приложении.
'-------------------------
Dim Result As Long ': Result = False
    On Error Resume Next
' по-умолчанию
    ObjectType = appObjTypUndef:    ObjectTypeDesc = "Объект"
    ObjectFileExt = vbNullString:   ObjectModuleName = vbNullString
' проверяем модули проекта
    Result = p_GetModuleType(ObjectName)
    If Result Then
    Select Case Result
    Case vbext_ct_StdModule:        ObjectType = appObjTypBas: ObjectTypeName = c_strObjTypModule: ObjectTypeDesc = "Стандартный модуль": ObjectModuleName = ObjectName: ObjectFileExt = c_strObjExtBas
    Case vbext_ct_ClassModule:      ObjectType = appObjTypCls: ObjectTypeName = c_strObjTypModule: ObjectTypeDesc = "Модуль класса": ObjectModuleName = ObjectName: ObjectFileExt = c_strObjExtCls
    'Case vbext_ct_Document:
    'Case vbext_ct_MSForm
    'Case vbext_ct_ActiveXDesigner
    End Select
    End If
    If ObjectType <> appObjTypUndef Then ObjectFileName = p_TextAlpha2Code(ObjectName) & "." & ObjectFileExt: GoTo HandleExit
' проверяем объекты специфичные для конкретных приложений
#If APPTYPE = 0 Then        ' APPTYPE=Access
' проверяем проверяем объекты Access по MSysObjects
    Result = Nz(DLookup("Type", c_strMSysObjects, "Name=""" & ObjectName & """"), msys_ObjectUndef) '(ObjectName)
    If Result Then
    Select Case Result
    ' проверяем наличие объекта с таким именем в базе по таблице
    Case msys_ObjectForm:   ObjectType = appObjTypAccFrm: ObjectTypeName = c_strObjTypAccFrm: ObjectTypeDesc = "Форма Access": ObjectFileExt = c_strObjExtFrm
        If p_GetModuleType(c_strFrmModPref & ObjectName) = vbext_ct_Document Then ObjectModuleName = c_strFrmModPref & ObjectName
    Case msys_ObjectReport: ObjectType = appObjTypAccRep: ObjectTypeName = c_strObjTypAccRep: ObjectTypeDesc = "Отчёт Access": ObjectFileExt = c_strObjExtRep
        If p_GetModuleType(c_strRepModPref & ObjectName) = vbext_ct_Document Then ObjectModuleName = c_strRepModPref & ObjectName
    Case msys_ObjectMacro:  ObjectType = appObjTypAccMac: ObjectTypeName = c_strObjTypAccMac: ObjectTypeDesc = "Макрос Access": ObjectFileExt = c_strObjExtMac
    Case msys_ObjectQuery:  ObjectType = appObjTypAccQry: ObjectTypeName = c_strObjTypAccQry: ObjectTypeDesc = "Запрос": ObjectFileExt = c_strObjExtQry
    Case msys_ObjectTable:  ObjectType = appObjTypAccTbl: ObjectTypeName = c_strObjTypAccTbl: ObjectTypeDesc = "Таблица Access": ObjectFileExt = c_strObjExtXml
    Case msys_ObjectLinked: ObjectType = appObjTypAcclnk: ObjectTypeName = c_strObjTypAccTbl: ObjectTypeDesc = "Связанная таблица": ObjectFileExt = c_strObjExtLnk
    Case Else: Err.Raise m_errModuleNameWrong
    End Select
    End If
'#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
#Else                       '
'Stop    ' ???
#End If                     ' APPTYPE
' если нашли объект в проекте получаем имя файла из имени объекта и выходим
' если не нашли - возможно имя объекта это путь к файлу
    If ObjectType <> appObjTypUndef Then ObjectFileName = p_TextAlpha2Code(ObjectName) & "." & ObjectFileExt: GoTo HandleExit
    Result = oFso.FileExists(ObjectName): If Not Result Then GoTo HandleExit
    ObjectFileExt = oFso.GetExtensionName(ObjectName)           ' расширение
    ObjectName = p_TextCode2Alpha(oFso.GetBaseName(ObjectName)) ' имя объекта из имени файла
' если указан путь к файлу выгрузки - получаем тип по расширению
    Select Case ObjectFileExt
    Case c_strObjExtBas:    ObjectType = appObjTypBas: ObjectTypeName = c_strObjTypModule: ObjectTypeDesc = "Стандартный модуль":   ObjectModuleName = ObjectName
    Case c_strObjExtCls:    ObjectType = appObjTypCls: ObjectTypeName = c_strObjTypModule: ObjectTypeDesc = "Модуль класса": ObjectModuleName = ObjectName
#If APPTYPE = 0 Then        ' APPTYPE=Access
    Case c_strObjExtFrm:    ObjectType = appObjTypAccFrm: ObjectTypeName = c_strObjTypAccFrm: ObjectTypeDesc = "Форма Access":      ObjectModuleName = c_strFrmModPref & ObjectName
    Case c_strObjExtRep:    ObjectType = appObjTypAccRep: ObjectTypeName = c_strObjTypAccRep: ObjectTypeDesc = "Отчёт Access":      ObjectModuleName = c_strRepModPref & ObjectName
    Case c_strObjExtMac:    ObjectType = appObjTypAccMac: ObjectTypeName = c_strObjTypAccMac: ObjectTypeDesc = "Макрос Access"
    Case c_strObjExtQry:    ObjectType = appObjTypAccQry: ObjectTypeName = c_strObjTypAccQry: ObjectTypeDesc = "Запрос"
    Case c_strObjExtTxt:    ObjectType = appObjTypAccTbl: ObjectTypeName = c_strObjTypAccTbl: ObjectTypeDesc = "Таблица Access (TXT)"
    Case c_strObjExtCsv:    ObjectType = appObjTypAccTbl: ObjectTypeName = c_strObjTypAccTbl: ObjectTypeDesc = "Таблица Access (CSV)"
    Case c_strObjExtXml:    ObjectType = appObjTypAccTbl: ObjectTypeName = c_strObjTypAccTbl: ObjectTypeDesc = "Таблица Access (XML)"
    Case c_strObjExtLnk:    ObjectType = appObjTypAcclnk: ObjectTypeName = c_strObjTypAccTbl: ObjectTypeDesc = "Связанная таблица"
    Case Else: Err.Raise m_errModuleNameWrong
'#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
#Else                       '
'Stop    ' ???
#End If                     ' APPTYPE
    End Select
HandleExit:     p_ObjectInfo = ObjectType <> appObjTypUndef:    Exit Function
HandleError:    Dim Message As String
    Select Case Err.Number
    Case m_errModuleNameWrong:  Message = "Неверно задано имя объекта!"
    Case Else:                  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_GetModuleType(ModuleName As String) As vbext_ComponentType
    On Error GoTo HandleError
    p_GetModuleType = Application.VBE.ActiveVBProject.VBComponents(ModuleName).Type
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function ObjectDelete(ObjectName As String, Optional ObjectType As appObjectType) As Boolean
' удаляет объект из проекта
Dim Result As Boolean
' вынесено в отдельную функцию на будущеее, чтобы когда-нибудь сделать поддержку разных приложений
    On Error Resume Next
Dim oTmp As Object
    If ObjectType = appObjTypUndef Then Result = p_ObjectInfo(ObjectName, ObjectType): If Not Result Then Err.Raise m_errObjectTypeUnknown
    Select Case ObjectType
    Case appObjTypBas, appObjTypCls:  With Application.VBE.ActiveVBProject: Set oTmp = .VBComponents(ObjectName): .VBComponents.Remove oTmp: End With
    'Case appObjTypDoc: :  With Application.VBE.ActiveVBProject: Set oTmp = .VBComponents(ObjectName): .VBComponents.Remove oTmp: End With
    'Case appObjTypMsf, appObjTypAxd: :  With Application.VBE.ActiveVBProject: Set oTmp = .VBComponents(ObjectName): .VBComponents.Remove oTmp: End With
#If APPTYPE = 0 Then        ' если Access
    Case appObjTypAccTbl, appObjTypAcclnk: DoCmd.DeleteObject ObjectType And &HFF&, ObjectName
    Case appObjTypAccQry, appObjTypAccMac: DoCmd.DeleteObject ObjectType And &HFF&, ObjectName
    Case appObjTypAccFrm, appObjTypAccRep ', appObjTypAccDoc: DoCmd.DeleteObject ObjectType And &HFF&, ObjectName
    'Case appObjTypAccDap, appObjTypAccSrv: DoCmd.DeleteObject ObjectType And &HFF&, ObjectName
    'Case appObjTypAccDia, appObjTypAccPrc, appObjTypAccFun: DoCmd.DeleteObject ObjectType And &HFF&, ObjectName
'#ElseIf APPTYPE = 1 Then    ' если Excel
    'Case appObjTypXlsDoc
#Else                       '
#End If
    Case Else: Err.Raise vbObjectError + 512
    End Select
    Result = (Err = 0): Err.Clear
HandleExit:     ObjectDelete = Result:    Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Public Function ObjectRead(SourceFile As String) As Boolean
Dim Result As Long
    On Error GoTo HandleError
Dim strMessage As String: strMessage = "Импорт объекта произведён успешно."
    Result = p_ObjectRead(SourceFile, ReadType:=orwAlways, AskBefore:=True, Message:=strMessage)
    MsgBox strMessage
HandleExit:  ObjectRead = Result = 0: Exit Function
HandleError: Result = Err: Err.Clear: Resume HandleExit
End Function
Public Function ObjectWrite(ObjectName As String, _
    Optional ByVal FilePath As String) As Boolean
Dim Result As Long
    On Error GoTo HandleError
Dim strMessage As String: strMessage = "Экспорт объекта произведён успешно."
    Result = p_ObjectWrite(ObjectName, FilePath, WriteType:=orwAlways, AskBefore:=True, Message:=strMessage)
    MsgBox strMessage
HandleExit:  ObjectWrite = Result = 0: Exit Function
HandleError: Result = Err: Err.Clear: Resume HandleExit
End Function

Private Function p_ObjectRead(SourceFile As String, _
    Optional TargetName As String, _
    Optional ObjectType As appObjectType = appObjTypUndef, _
    Optional ReadType As ObjectRwType = orwDestMiss, _
    Optional AskBefore As Boolean = True, _
    Optional Message As String) As Long
' читает объект из сохраненных файлов в текущий проект с проверкой версий (где доступно)
Const c_strProcedure = "p_ObjectRead"
' SourceFile    - путь к файлу читаемого объекта
' TargetName    - имя объекта в проекте который должен быть обновлен/создан
' ObjectType    - тип читаемого объекта
' ReadType      - тип операции чтения
' AskBefore     - перед обновлением модуля спрашивать подтверждение пользователя/обновлять без подтверждения
' Message       - возвращает сообщение по итогам операции (для вывода вызывающей функцией)
' возвращает результат операции:   0 - прочитан, иначе код ошибки
Const cstrMsgYesNo = """Да""  - если хотите чтобы действие было выполнено, " & vbCrLf & _
                     """Нет"" - если хотите пропустить действие"
Const cstrMsgCancel = """Отмена"" - если хотите чтобы данное и все последующие действия выполнялись без предупреждения."
Dim Result As Long: Result = 1
Dim bolClear As Boolean:  bolClear = False ' в случае ошибки чтения не заменяем оригинал "очищенным"
Dim bolRestored As Boolean ' признак того что файл был очищен для корректного импорта
Dim errCount As Integer    ' счётчик ошибок
    On Error GoTo HandleError
'    DoCmd.SetWarnings False
'Dim SourceFile As String, strObjName As String ', strFileExtn As String
Dim strObjName As String, lngObjType As appObjectType
Dim strObjType As String, strObjFile As String, strObjExtn As String, strModName As String
Dim bolProceed As Boolean       ' флаг необходимости чтения объекта из файла
Dim intCmp As Integer
' проверяем наличие файла
    strObjName = SourceFile
Dim bolSrcExist As Boolean:  bolSrcExist = p_ObjectInfo(strObjName, ObjectType, , strObjType, strObjFile, strObjExtn): If Not bolSrcExist Then Err.Raise 76
    'strFileExtn = VBA.LCase$(.GetExtensionName(SourceFile))    ' расширение файла
    If Len(TargetName) = 0 Then TargetName = strObjName
' проверяем наличие объекта в проекте
    strObjName = TargetName
Dim bolDestExist As Boolean: bolDestExist = p_ObjectInfo(strObjName, lngObjType, ObjectModuleName:=strModName)
'!!! если тип прочитанный из файла и приложения не совпадают - происходит какая-то ерунда
    If bolDestExist And (ObjectType <> lngObjType) Then ObjectType = lngObjType ': Stop
' проверяем возможность чтения версий и читаем их
    If (ObjectType And &HFFFF&) > 0 Then
Dim strSrcVer As String, datSrcDate As Date ', strFilDesc As String
Dim strDestVer As String, datDestDate As Date ', strModDesc As String
' читаем текст существующего модуля и извлекаем из него данные о версии
' ToDo: прочитать из текста истинное имя объекта TargetName
        If bolDestExist Then If Len(strModName) > 0 Then Call ModuleInfo(strModName, strDestVer, datDestDate)
' читаем файл как текст и извлекаем из него данные о версии, дате и пр.
        If bolSrcExist Then Call ModuleInfoFromFile(SourceFile, strSrcVer, datSrcDate)
    ' сравниваем версии
        If bolSrcExist And bolDestExist Then intCmp = VersionCmp(strSrcVer, strDestVer)
    End If
' проверяем необходимость выполнения операции иначе вызываем ошибки
    Message = strObjType & " """ & TargetName & """"
    Select Case ReadType
    Case orwAlways                  ' читаем всегда (восстановление)
        Message = Message & " полное восстановление из резервной копии."
    Case orwSrcNewerOrDestMissing   ' читаем если файл новее или объект отсутствует
        Message = Message & " восстановление из резервной копии."
        If (intCmp <> 1) And bolDestExist Then Err.Raise m_errWrongVersion
    Case orwSrcNewer                ' читаем если файл новее и объект существует
        Message = Message & " обновление из файла."
        If Not (intCmp = 1) Then Err.Raise m_errWrongVersion
        If Not bolDestExist Then Err.Raise m_errDestMissing
    Case orwDestMiss                ' читаем если объект отсутствует
        Message = Message & " добавление из файла."
        If bolDestExist Then Err.Raise m_errDestExists
    Case orwSrcOlder                ' читаем если файл старше и объект существует
        Message = Message & " восстановление предыдущей версии из файла."
        If Not (intCmp = -1) Then Err.Raise m_errWrongVersion
        If Not bolDestExist Then Err.Raise m_errDestMissing
    Case Else: Err.Raise m_errObjectActionUndef ' несуществующий вариант - непонятно что делать с файлом
    End Select
' добавляем в вывод информацию о версиях
    If Len(strDestVer) > 0 Or datDestDate > 0 Then
        Message = Message & vbCrLf & "текущая версия"
        If Len(strDestVer) > 0 Then Message = Message & ": " & strDestVer
        If datDestDate > 0 Then Message = Message & " от " & datDestDate
    End If
    If Len(strSrcVer) > 0 Or datSrcDate > 0 Then
        Message = Message & vbCrLf & "новая версия"
        If Len(strSrcVer) > 0 Then Message = Message & ": " & strSrcVer
        If datSrcDate > 0 Then Message = Message & " от " & datSrcDate
    End If
' !!! пропускаем объекты необходимые для работы процесса импорта для избежания обновления исполняемого кода
    If p_IsSkippedObject(TargetName) Then Err.Raise m_errSkippedByList

' спросить разрешение пользователя на обновление модуля
    If AskBefore Then ' intAsk <> vbCancel Then
        Select Case MsgBox(Message & vbCrLf & vbCrLf & cstrMsgYesNo & ", " & vbCrLf & cstrMsgCancel, vbYesNoCancel Or vbInformation, "Обновление объекта")
        Case vbYes                                      ' обновить только текущий
        Case vbCancel:  AskBefore = False               ' обновить текущий и последующие
        Case Else:      Err.Raise m_errSkippedByUser    ' пропустить текущий
        End Select
    End If
' если объект существует удаляем его из проекта
    If bolDestExist Then If Not ObjectDelete(TargetName, ObjectType) Then Err.Raise m_errObjectCantRemove
' в зависимости от типа загружаем объекты
' !!! при загрузке проверка корректности файлов не производится !!!
    Select Case ObjectType
    Case appObjTypMod, appObjTypBas, appObjTypCls
                                            VBE.ActiveVBProject.VBComponents.Import SourceFile
'    Case appObjTypMsf
'    Case appObjTypAxd
'    Case appObjTypDoc
#If APPTYPE = 0 Then        ' APPTYPE=Access
    Case appObjTypAccTbl
        Select Case strObjExtn    ' определяем по расширению файла
        Case c_strObjExtXml:                ImportXML DataSource:=SourceFile, ImportOptions:=acStructureAndData
        Case c_strObjExtTxt:                DoCmd.TransferText TransferType:=acImportDelim, TableName:=TargetName, FileName:=SourceFile, HasFieldNames:=True
        Case c_strObjExtCsv:                DoCmd.TransferText TransferType:=acImportDelim, TableName:=TargetName, FileName:=SourceFile, HasFieldNames:=True
        Case c_strObjExtUndef:              ImportXML DataSource:=SourceFile, ImportOptions:=acStructureAndData
        Case Else:                          Err.Raise m_errObjectTypeUnknown
        End Select
    Case appObjTypAcclnk:                   p_LinkedRead TargetName, SourceFile    ' читаем из файла SourceFile
    Case appObjTypAccMac, appObjTypAccQry:  LoadFromText (ObjectType And &HFF&), TargetName, SourceFile
    Case appObjTypAccFrm, appObjTypAccRep:  LoadFromText (ObjectType And &HFF&), TargetName, SourceFile
'    Case appObjTypAccDap
'    Case appObjTypAccSrv
'    Case appObjTypAccDia
'    Case appObjTypAccPrc
'    Case appObjTypAccFun
#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
'    Case appObjTypXlsDoc
#Else                       ' APPTYPE=Неразбери пойми что
#End If                     ' APPTYPE
    Case Else:                              Err.Raise m_errObjectTypeUnknown
    End Select
' если импортированный файл был изменён для загрузки и замена не производилась
' надо удалить во временной папке его модифицированную версию
    If bolRestored And Not bolClear Then Kill SourceFile
    Result = 0
HandleExit:  p_ObjectRead = Result: Exit Function
HandleError:    'Dim Message As String
    Select Case Err 'Err.Number
    Case 75: Err.Clear: Resume Next     ' Path/File access error
    Case 76: Message = Err.Description ' Path not found
    Case m_errModuleNameWrong: Message = "Неверно задано имя объекта!": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
    Case m_errObjectActionUndef: Message = "Непонятно что делать с файлом": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
    Case m_errObjectCantRemove: Message = "Не удалось удалить объект": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
    Case m_errObjectTypeUnknown: Message = "Невозможно прочитать модуль данного типа": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
' причины пропуска:
    Case m_errWrongVersion: Message = "Ошибка сравнения версий файла и объекта": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errDestMissing: Message = "Загружаемый объект отсутствует в приложении": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errDestExists: Message = "Загружаемый объект уже существует в приложении": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errSkippedByUser: Message = "Пропущено по решению пользователя": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errSkippedByList: Message = "Наличие в списке пропуска": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case 2285 'Приложению "Microsoft Office Access" не удается создать выходной файл.
        Message = Err.Description
        Message = Message & vbCrLf & "Вероятно объект: " & strObjType & " """ & TargetName & """" & vbCrLf & _
                    "сохранён в формате другой версии приложения."
        If Not bolRestored Then
            If MsgBox(Message & vbCrLf & vbCrLf & _
                    "Возможно удастся загрузить файл в проект с частичной потерей информации." & vbCrLf & _
                    "Очистить файл от фрагментов, которые могут мешать загрузке?", _
                    vbYesNo Or vbQuestion, c_strProcedure) _
                    = vbYes Then
            ' попытка очистки файла от вызывающих ошибку фрагментов
                bolRestored = p_ObjectFileClear(SourceFile, TargetName, ObjectType, bolClear)
                If bolRestored Then errCount = errCount + 1: Err.Clear: Resume 0
            End If
        End If: Message = Replace(Message, vbCrLf, " ")
    Case 1004 '??? ' Ошибка: Программный доступ к проекту Visual Basic не является доверенным
            Message = "Программный доступ к проекту Visual Basic не является доверенным. Для возможности программного сохранения/восстановления модулей необходимо установить разрешение: ""Сервис\Макрос\Безопасность\Доверять доступ к Visual Basic Project"""
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Result = Err: Err.Clear: Resume HandleExit
End Function
Private Function p_ObjectWrite(SourceName As String, _
    Optional ByVal TargetPath As String, _
    Optional ObjectType As appObjectType = appObjTypUndef, _
    Optional WriteType As ObjectRwType = orwDestMiss, _
    Optional AskBefore As Boolean = True, _
    Optional Message As String) As Long
' читает объект из сохраненных файлов в текущий проект с проверкой версий (где доступно)
Const c_strProcedure = "p_ObjectWrite"
' SourceName    - имя объекта в проекте который должен быть сохранён
' TargetPath    - путь для сохранения объекта
' ObjectType    - тип сохраняемого объекта
' WriteType     - тип операции сохранения
' AskBefore     - перед обновлением модуля спрашивать подтверждение пользователя/обновлять без подтверждения
' Message       - возвращает сообщение по итогам операции (для вывода вызывающей функцией)
' возвращает результат операции:   0 - прочитан, иначе код ошибки
Const cstrMsgYesNo = """Да""  - если хотите чтобы действие было выполнено, " & vbCrLf & _
                     """Нет"" - если хотите пропустить действие"
Const cstrMsgCancel = """Отмена"" - если хотите чтобы данное и все последующие действия выполнялись без предупреждения."
Dim Result As Long: Result = 1
Dim bolClear As Boolean:  bolClear = False ' флаг необходимости очистки файлов перед сохранением от фрагментов могущих препятствовать их повторной загрузке
Dim errCount As Integer    ' счётчик ошибок
    On Error GoTo HandleError
'    DoCmd.SetWarnings False
'Dim SourceName As String, strObjName As String ', strFileExtn As String
Dim strObjName As String, lngObjType As appObjectType
Dim strObjType As String, strObjFile As String, strModName As String
Dim bolProceed As Boolean       ' флаг необходимости чтения объекта из файла
Dim intCmp As Integer
' проверяем наличие объекта в проекте
Dim bolSrcExist As Boolean:  bolSrcExist = p_ObjectInfo(SourceName, ObjectType, , strObjType, strObjFile, , strModName): If Not bolSrcExist Then Err.Raise 76
' проверить - передан путь к папке или имя файла
' и наличие файла по указанному пути
Dim bolDestExist As Boolean
    With oFso
        If Len(TargetPath) = 0 Then
    ' не задано - берем путь проекта, надо добавить имя файла
            TargetPath = .BuildPath(CurrentProject.path, c_strSrcPath)
            TargetPath = .BuildPath(TargetPath, strObjFile)
            bolDestExist = .FileExists(TargetPath)
        ElseIf .FolderExists(TargetPath) Then
    ' задан путь, надо добавить имя файла
            TargetPath = .BuildPath(TargetPath, strObjFile)
            bolDestExist = .FileExists(TargetPath)
        ElseIf Right(TargetPath, 1) = "\" Then
    ' задан несуществующий путь, надо добавить имя файла
            If Not .FolderExists(TargetPath) Then Call .CreateFolder(TargetPath) 'Then Err.Raise 76 ' Path not Found
            TargetPath = .BuildPath(TargetPath, strObjFile)
        ElseIf .GetExtensionName(TargetPath) > 0 Then
    ' задано имя файла ' !!! криво - проверяем что это файл по наличию расширения в имени
        Else: Err.Raise 76
        End If
    End With
' читаем тип из файла
    strObjName = TargetPath: bolDestExist = p_ObjectInfo(strObjName, lngObjType) ': If Not bolDestExist Then Err.Raise 76
'!!! если тип прочитанный из файла и приложения не совпадают - происходит какая-то ерунда
    If bolDestExist And (ObjectType <> lngObjType) Then Err.Raise m_errObjectActionUndef
' проверяем возможность чтения версий и читаем их
    If (ObjectType And &HFFFF&) > 0 Then
Dim strSrcVer As String, datSrcDate As Date ', strFilDesc As String
Dim strDestVer As String, datDestDate As Date ', strModDesc As String
' читаем файл как текст и извлекаем из него данные о версии, дате и пр.
' ToDo: прочитать из файла истинное имя объекта TargetPath
'Stop
        If bolDestExist Then Call ModuleInfoFromFile(TargetPath, strDestVer, datDestDate)
' читаем текст существующего модуля и извлекаем из него данные о версии
        If bolSrcExist And Len(strModName) > 0 Then Call ModuleInfo(strModName, strSrcVer, datSrcDate)
    ' сравниваем версии
        If bolSrcExist And bolDestExist Then intCmp = VersionCmp(strSrcVer, strDestVer)
    End If
' проверяем необходимость выполнения операции иначе вызываем ошибки
    Message = strObjType & " """ & SourceName & """"
    Select Case WriteType
    Case orwAlways                  ' сохраняем всегда (бэкап)
        Message = Message & " полное сохранение в резервную копию."
    Case orwSrcNewerOrDestMissing   ' сохраняем если объект новее или файл отсутствует
        Message = Message & " сохранение резервной копии."
        If (intCmp <> 1) And bolDestExist Then Err.Raise m_errWrongVersion
    Case orwSrcNewer                ' сохраняем если объект новее и файл существует
        Message = Message & " обновление сохранённой версии из проекта."
        If Not (intCmp = 1) Then Err.Raise m_errWrongVersion
        If Not bolDestExist Then Err.Raise m_errDestMissing
    Case orwDestMiss                ' сохраняем если файл отсутствует
        Message = Message & " добавление в папку сохраннения из проекта."
        If bolDestExist Then Err.Raise m_errDestExists
    Case orwSrcOlder                ' сохраняем если объект старше и файл существует
        Message = Message & " замена предыдущей версией из проекта."
        If Not (intCmp = -1) Then Err.Raise m_errWrongVersion
        If Not bolDestExist Then Err.Raise m_errDestMissing
    Case Else: Err.Raise m_errObjectActionUndef ' несуществующий вариант - непонятно что делать с файлом
    End Select
' добавляем в вывод информацию о версиях
    If Len(strDestVer) > 0 Or datDestDate > 0 Then
        Message = Message & vbCrLf & "сохранённая версия"
        If Len(strDestVer) > 0 Then Message = Message & ": " & strDestVer
        If datDestDate > 0 Then Message = Message & " от " & datDestDate
    End If
    If Len(strSrcVer) > 0 Or datSrcDate > 0 Then
        Message = Message & vbCrLf & "сохраняемая версия"
        If Len(strSrcVer) > 0 Then Message = Message & ": " & strSrcVer
        If datSrcDate > 0 Then Message = Message & " от " & datSrcDate
    End If
'' !!! пропускаем объекты необходимые для работы процесса импорта для избежания обновления исполняемого кода
'    If p_IsSkippedObject(SourceObject) Then Err.Raise m_errSkippedByList
' спросить разрешение пользователя на обновление модуля
    If AskBefore Then ' intAsk <> vbCancel Then
        Select Case MsgBox(Message & vbCrLf & vbCrLf & cstrMsgYesNo & ", " & vbCrLf & cstrMsgCancel, vbYesNoCancel Or vbInformation, "Сохранение объекта")
        Case vbYes                                      ' обновить только текущий
        Case vbCancel:  AskBefore = False               ' обновить текущий и последующие
        Case Else:      Err.Raise m_errSkippedByUser    ' пропустить текущий
        End Select
    End If
' если файл существует удаляем его
    If bolDestExist Then Kill TargetPath
' в зависимости от типа сохраняем объекты
    Select Case ObjectType
    Case appObjTypMod, appObjTypBas, appObjTypCls
                                            VBE.ActiveVBProject.VBComponents(SourceName).Export TargetPath
'    Case appObjTypMsf
'    Case appObjTypAxd
'    Case appObjTypDoc
#If APPTYPE = 0 Then        ' APPTYPE=Access
    Case appObjTypAccTbl:                   ExportXML DataSource:=SourceName, DataTarget:=TargetPath, ObjectType:=acExportTable, OtherFlags:=acEmbedSchema
    Case appObjTypAcclnk:                   p_LinkedWrite SourceName, TargetPath
    Case appObjTypAccMac, appObjTypAccQry:  SaveAsText (ObjectType And &HFF&), SourceName, TargetPath
    Case appObjTypAccFrm, appObjTypAccRep:  SaveAsText (ObjectType And &HFF&), SourceName, TargetPath
                                            If bolClear Then Call p_ObjectFileClear(TargetPath, ObjectType:=ObjectType, ReplaceOriginal:=True)
'    Case appObjTypAccDap
'    Case appObjTypAccSrv
'    Case appObjTypAccDia
'    Case appObjTypAccPrc
'    Case appObjTypAccFun
#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
'    Case appObjTypXlsDoc
#Else                       ' APPTYPE=Неразбери пойми что
#End If                     ' APPTYPE
    Case Else:                              Err.Raise m_errObjectTypeUnknown
    End Select
    Result = 0
HandleExit:  p_ObjectWrite = Result: Exit Function
HandleError:    'Dim Message As String
    Select Case Err 'Err.Number
    Case 75: Err.Clear: Resume Next     ' Path/File access error
    Case 76: Message = Err.Description ' Path not found
    Case m_errModuleNameWrong: Message = "Неверно задано имя объекта!": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
    Case m_errObjectActionUndef: Message = "Непонятно что делать с файлом": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
    Case m_errObjectCantRemove: Message = "Не удалось удалить объект": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
    Case m_errObjectTypeUnknown: Message = "Невозможно прочитать модуль данного типа": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
' причины пропуска:
    Case m_errWrongVersion: Message = "Ошибка сравнения версий файла и объекта": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errDestMissing: Message = "Файл сохраняемого объекта отсутствует": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errDestExists: Message = "Файл сохраняемого объекта уже существует": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errSkippedByUser: Message = "Пропущено по решению пользователя": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errSkippedByList: Message = "Наличие в списке пропуска": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
        Result = Err: Err.Clear: Resume HandleExit
 '
    Case 1004 '??? ' Ошибка: Программный доступ к проекту Visual Basic не является доверенным
                Message = "Программный доступ к проекту Visual Basic не является доверенным. Для возможности программного сохранения/восстановления модулей необходимо установить разрешение: ""Сервис\Макрос\Безопасность\Доверять доступ к Visual Basic Project"""
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Result = Err: Err.Clear: Resume HandleExit
End Function
Private Function p_ObjectFileClear( _
    SourceFile As String, _
    Optional TargetName As String, _
    Optional ObjectType As appObjectType = appObjTypUndef, _
    Optional ReplaceOriginal As Boolean = False) As Boolean
' очистка файла сохранения от фрагментов препятствующих его загрузке
Dim Result As Boolean: Result = False
' based on: https://github.com/bkidwell/msaccess-vcs-integration/issues/12?utm_source=pocket_mylist#issuecomment-29564498
    On Error GoTo HandleError
    'strReadLine = oFso.ReadLine
    If Not oFso.FileExists(SourceFile) Then Err.Raise 76 ' Path not Found
Dim TargetFile As String: TargetFile = VBA.Environ$("Temp") & "\~@#" & Format$(Now, "yyyymmddhhnnss")

Dim aSkipStrings(), aSkipBlocks() 'As String
Dim sBlockBeg As String, sBlockEnd As String
    Select Case ObjectType
    Case appObjTypAccFrm, appObjTypAccRep
' для форм и отчетов при чтении пропускать
'' строки:
    aSkipStrings = Array( _
        "Checksum", _
        "FilterOnLoad", _
        "AllowLayoutView", _
        "NoSaveCTIWhenDisabled", _
        "Overlaps", "BorderLineStyle", _
        "WebImagePaddingLeft", "WebImagePaddingTop", "WebImagePaddingRight", "WebImagePaddingBottom")
'' блоки:
    sBlockBeg = "Begin": sBlockEnd = "End"
    aSkipBlocks = Array( _
        "PrtDevNames", "PrtDevNamesW", _
        "PrtDevMode", "PrtDevModeW", _
        "NameMap", "NameMapW") ', _
        "GUID") ', _
        "PrtMip", "PrtMipW") ', _
        "dbLongBinary ""DOL""", _

    Case Else: Err.Raise vbObjectError + 2048 ' необрабатываемый объект
    End Select
Dim strReadLine As String

'1. читать файл построчно
Dim iFileIn As Integer:     iFileIn = FreeFile:  Open SourceFile For Input As iFileIn
Dim iFileOut As Integer:    iFileOut = FreeFile: Open TargetFile For Output Access Write As #iFileOut

Dim strLine As String, strTemp As String, varTemp '
Dim bolEndOfBlock As Boolean
Dim LineType As m_CodeLineType
    Do Until EOF(iFileIn) ' Or LineType = m_CodeProc
    ' читаем построчно пока не достигнем конца
        Line Input #iFileIn, strLine: If LineType = m_CodeProc Then GoTo HandleOutput
'Stop
' для Access !!!
    ' проверяем строку на конец заголовка объекта (объявление первой процедуры)
        If InStrRegEx(1, strLine, c_strCodeProcBeg) > 0 Then LineType = m_CodeProc: GoTo HandleOutput
'2. пропускать недопустимые фрагменты
    ' проверяем по списку пропуска строк и пропускаем если нашли
        For Each varTemp In aSkipStrings
            strTemp = "^\s*" & varTemp & "\s*=\s*"
            If InStrRegEx(1, strLine, strTemp) > 0 Then GoTo HandleNextLine
        Next varTemp
    ' проверяем по списку пропуска блоков
        For Each varTemp In aSkipBlocks
        ' ищем открывающие блок строки
            strTemp = "^\s*" & varTemp & "\s*=\s*" & sBlockBeg
            If InStrRegEx(1, strLine, strTemp$) > 0 Then
                strTemp = "^\s*" & sBlockEnd
                Do
        ' читаем построчно пока не дойдем до закрывающей блок строки или не достигнем конца
                    Line Input #iFileIn, strLine
                    If InStrRegEx(1, strLine, strTemp) > 0 Then Exit Do ' закрытый блок
                    If EOF(iFileIn) Then Err.Raise vbObjectError + 2049 ' незакрытый блок
                    If InStrRegEx(1, strLine, c_strCodeProcBeg) > 0 Then Err.Raise vbObjectError + 2049 ' незакрытый блок
                Loop
        ' и пропускаем весь блок если нашли
                GoTo HandleNextLine
            End If
        Next varTemp
'3. сохранить результат в Temp
HandleOutput:
        Print #iFileOut, strLine
HandleNextLine:
    Loop
    Close #iFileIn: Close #iFileOut
' финальная обработка
    If ReplaceOriginal Then
' заменяем исходный файл преобразованным и возвращаем путь исходного файла
        Kill PathName:=SourceFile
        FileCopy Source:=TargetFile, Destination:=SourceFile
        Kill PathName:=TargetFile
    Else
' возвращаем путь к преобразованному файлу
        SourceFile = TargetFile
    End If
    Result = True
HandleExit:  p_ObjectFileClear = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Sub ModuleHeaderCreate( _
    ModuleName As String)
Const c_strProcedure = "ModuleHeaderCreate"
Dim Result As Boolean ': Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
' Строка разделителя
Dim CodeLine As Long, PrevLine As Long
    CodeLine = CodeLineFind(ModuleName, c_strPrefModLine)
    If CodeLine = 0 Then objModule.InsertLines 1, c_strPrefModLine: PrevLine = CodeLine + 1
' Имя модуля
    CodeLine = CodeLineFind(ModuleName, c_strPrefModName, PrevLine)
    If CodeLine = 0 Then CodeLine = PrevLine + 1 Else objModule.DeleteLines CodeLine, 1
    objModule.InsertLines CodeLine, c_strPrefModName & VBA.Chr(34) & ModuleName & VBA.Chr(34): PrevLine = CodeLine
' Строка разделителя
    objModule.InsertLines CodeLine, c_strPrefModLine: PrevLine = CodeLine
' Описание
    CodeLine = CodeLineFind(ModuleName, c_strPrefModDesc, PrevLine)
    If CodeLine = 0 Then CodeLine = PrevLine + 1 Else objModule.DeleteLines CodeLine, 1
    objModule.InsertLines CodeLine, c_strPrefModDesc: PrevLine = CodeLine
' Версия
    CodeLine = CodeLineFind(ModuleName, c_strPrefModVers, PrevLine)
    If CodeLine = 0 Then CodeLine = PrevLine + 1 Else objModule.DeleteLines CodeLine, 1
    objModule.InsertLines CodeLine, c_strPrefModVers: PrevLine = CodeLine
' Дата
    CodeLine = CodeLineFind(ModuleName, c_strPrefModDate, PrevLine)
    If CodeLine = 0 Then CodeLine = PrevLine + 1 Else objModule.DeleteLines CodeLine, 1
    objModule.InsertLines CodeLine, c_strPrefModDate: PrevLine = CodeLine
' Автор
    CodeLine = CodeLineFind(ModuleName, c_strPrefModAuth, PrevLine)
    If CodeLine = 0 Then CodeLine = PrevLine + 1 Else objModule.DeleteLines CodeLine, 1
    objModule.InsertLines CodeLine, c_strPrefModAuth: PrevLine = CodeLine
' Примечание
    CodeLine = CodeLineFind(ModuleName, c_strPrefModComm, PrevLine)
    If CodeLine = 0 Then CodeLine = PrevLine + 1 Else objModule.DeleteLines CodeLine, 1
    objModule.InsertLines CodeLine, c_strPrefModComm: PrevLine = CodeLine + 1
' Строка разделителя
    objModule.InsertLines CodeLine, c_strPrefModLine
    
    Result = True
HandleExit:     Exit Sub
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "Неверно задано имя объекта!"
    Case m_errModuleIsActive: Debug.Print "Невозможно изменить активный модуль: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "Модуль: """ & ModuleName & """ не найден!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Sub
Public Function ModuleInfo( _
    ModuleName As String, _
    Optional ModVers, Optional ModDate, _
    Optional ModDesc, Optional ModAuth, _
    Optional ModComm, Optional ModHist) As Boolean
' читает информацию о модуле из текста модуля
Const c_strProcedure = "ModuleInfo"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
Dim CodeLine As Long
Dim LineType As m_CodeLineType, LineBreak As Boolean
' LineType - флаг управляющий циклом проверки
' LineBreak - признак продолжения чтения строки (строка заканчивается символом переноса)
    LineType = m_CodeHead ' модуль сразу начинается с кода

Dim strLine As String, strResult As String
    For CodeLine = 1 To objModule.CountOfDeclarationLines
        'If LineType = m_CodeProc Then Exit For
    ' читаем построчно пока не достигнем конца или не дойдем до объявления первой процедуры
        strLine = objModule.Lines(CodeLine, 1)
    ' проверяем в каком месте модуля находимся
    ' и в зависимости от содержимого строки выполнняем действия
        ' в выгрузках форм и отчетов сначала идет описание расположения элементов формы
        ' поэтому пропускаем все строки до начала текста модуля
        ' в модулях анализ начинаем с первой строки
        If LineType = m_CodeNone Then
            If VBA.Trim$(strLine) = c_strCodeHeadBeg Then LineType = m_CodeHead
            GoTo HandleNextLine
        End If
        ' определяем тип строки
        strLine = p_CodeLineGet(strLine, LineType, LineBreak, vbCrLf)
        ' формируем строку результата
        strResult = strResult & strLine
        ' сжимаем строку результата - удаляем мусор
        strResult = Replace(strResult, c_strBrokenQuotes, vbNullString)   ' объединяем разорванные текстовые строки
HandleRead:
        If LineType <> m_CodeHead Then
        ' сохраниить результат
            If LineBreak Then GoTo HandleNextLine
            If Not IsMissing(ModVers) And LineType = m_CodeVers And Len(strResult) > 0 Then ModVers = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModDate) And LineType = m_CodeDate And Len(strResult) > 0 Then ModDate = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModDesc) And LineType = m_CodeDesc And Len(strResult) > 0 Then ModDesc = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModAuth) And LineType = m_CodeAuth And Len(strResult) > 0 Then ModAuth = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModComm) And LineType = m_CodeComm And Len(strResult) > 0 Then ModComm = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModHist) And LineType = m_CodeHist And Len(strResult) > 0 Then ModHist = ModHist & vbCrLf & strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
        End If
        LineType = m_CodeHead
        strResult = vbNullString
HandleNextLine:
    Next
' финальная обработка параметров
    If Not IsMissing(ModVers) Then If Len(ModVers) = 0 Then ModVers = cEmptyVers
    If Not IsMissing(ModDate) Then If Len(ModDate) = 0 Then ModDate = cEmptyDate
    If Not IsMissing(ModHist) Then If Left(ModHist, Len(vbCrLf)) = vbCrLf Then ModHist = Mid(ModHist, Len(vbCrLf) + 1)
HandleExit:  ModuleInfo = Result: Exit Function
HandleError: Result = False
Dim Message As String
    Select Case Err.Number
    Case m_errModuleNameWrong: Message = "Не задано имя модуля!"
    Case m_errModuleIsActive: Message = "Невозможно изменить активный модуль!"
    Case m_errModuleDontFind: Message = "Модуль не найден!"
    Case Else: Message = Err.Description
    End Select
    If Len(ModuleName) > 0 Then Message = Message & "; ModuleName=""" & ModuleName & """ "
    'If Len(strResult) > 0 Then Message = Message & "; strResult=""" & strResult & """ "
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
End Function
Public Function ModuleVersGet(ModuleName As String) As String: Call ModuleInfo(ModuleName, ModVers:=ModuleVersGet): End Function
Public Function ModuleDateGet(ModuleName As String) As String: Call ModuleInfo(ModuleName, ModDate:=ModuleDateGet): End Function
Public Function ModuleAuthGet(ModuleName As String) As String: Call ModuleInfo(ModuleName, ModAuth:=ModuleAuthGet): End Function
Public Function ModuleDescGet(ModuleName As String) As String: Call ModuleInfo(ModuleName, ModDesc:=ModuleDescGet): End Function
Public Function ModuleCommGet(ModuleName As String) As String: Call ModuleInfo(ModuleName, ModComm:=ModuleCommGet): End Function
Public Function ModuleHistGet(ModuleName As String): Call ModuleInfo(ModuleName, ModHist:=ModuleHistGet): End Function

Public Function ModuleInfoFromFile( _
    FilePath As String, _
    Optional ModVers, Optional ModDate, _
    Optional ModDesc, Optional ModAuth, _
    Optional ModComm, Optional ModHist) As Boolean
' читает информацию о модуле из текста файла модуля
Const c_strProcedure = "ModuleInfoFromFile"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
Dim CodeLine As Long
Dim strPath As String:  strPath = VBA.Trim$(FilePath)
' ToDo: прочитать из файла истинное имя объекта TargetPath
    ' проверяем наличие указанного пути/файла
    If Not oFso.FileExists(strPath) Then Err.Raise 76 ' Path not Found
' проверяем тип файла (по расширению)
Dim oType As msys_ObjectType, cType As vbext_ComponentType
Dim strExtn As String: strExtn = VBA.LCase$(oFso.GetExtensionName(strPath))
    Select Case strExtn
    Case c_strObjExtBas: oType = msys_ObjectModule  ' стандартный модуль
    Case c_strObjExtCls: oType = msys_ObjectModule  ' модуля класса
    Case c_strObjExtDoc: oType = msys_ObjectModule  ' модуля класса документа Access (Form или Report)
    Case c_strObjExtFrm: oType = msys_ObjectForm    ' форма Access (включая модуль)
    Case c_strObjExtRep: oType = msys_ObjectReport  ' отчет Access (включая модуль)
    Case c_strObjExtUndef: oType = msys_ObjectModule  ' старое расширение файла бэкапа объекта Access
    'Case c_strObjExtXml: oType = msys_ObjectTable   ' локальная таблица в XML
    'Case c_strObjExtTxt: oType = msys_ObjectTable   ' локальная таблица в TXT
    'Case c_strObjExtCsv: oType = msys_ObjectTable   ' локальная таблица в CSV (текст с разделителями)
    'Case c_strObjExtLnk: oType = msys_ObjectLinked  ' связанная таблица
    'Case c_strObjExtQry: oType = msys_ObjectQuery   ' запрос Access
    'Case c_strObjExtMac: oType = msys_ObjectMacro   ' макрос Access
    Case Else: Err.Raise vbObjectError + 512
    End Select
Dim LineType As m_CodeLineType, LineBreak As Boolean
' LineType - флаг управляющий циклом проверки
' LineBreak - признак продолжения чтения строки (строка заканчивается символом переноса)
    Select Case oType
    Case msys_ObjectModule:                     LineType = m_CodeHead ' модуль сразу начинается с кода
    Case msys_ObjectForm, msys_ObjectReport:    LineType = m_CodeNone ' форма/отчет сначала содержат информацию о расположении контролов
    Case Else: Err.Raise vbObjectError + 512              ' прочие элементы не обрабатываются т.к. не содержат кода
    End Select

Dim iFile As Integer:   iFile = FreeFile
    Open strPath For Input As #iFile: CodeLine = 0

Dim strLine As String, strResult As String
    Do Until EOF(iFile) Or LineType = m_CodeProc
    ' читаем построчно пока не достигнем конца или не дойдем до объявления первой процедуры
        Line Input #iFile, strLine
        CodeLine = CodeLine + 1
    ' проверяем в каком месте модуля находимся
    ' и в зависимости от содержимого строки выполнняем действия
        ' в выгрузках форм и отчетов сначала идет описание расположения элементов формы
        ' поэтому пропускаем все строки до начала текста модуля
        ' в модулях анализ начинаем с первой строки
        If LineType = m_CodeNone Then
            If VBA.Trim$(strLine) = c_strCodeHeadBeg Then LineType = m_CodeHead
            GoTo HandleNextLine
        End If
        ' Определяем тип строки
        strLine = p_CodeLineGet(strLine, LineType, LineBreak, vbCrLf)
        ' формируем строку результата
        strResult = strResult & strLine
        ' сжимаем строку результата - удаляем мусор
        strResult = Replace(strResult, c_strBrokenQuotes, vbNullString)   ' объединяем разорванные текстовые строки
HandleRead:
        If LineType <> m_CodeHead Then
        ' сохраниить результат
            If LineBreak Then GoTo HandleNextLine
            If Not IsMissing(ModVers) And LineType = m_CodeVers Then ModVers = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModDate) And LineType = m_CodeDate Then ModDate = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModDesc) And LineType = m_CodeDesc Then ModDesc = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModAuth) And LineType = m_CodeAuth Then ModAuth = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModComm) And LineType = m_CodeComm Then ModComm = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModHist) And LineType = m_CodeHist Then ModHist = ModHist & vbCrLf & strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
        End If
        LineType = m_CodeHead
        strResult = vbNullString
HandleNextLine:
    Loop
    Close #iFile
' финальная обработка параметров
    If Not IsMissing(ModVers) Then If Len(ModVers) = 0 Then ModVers = cEmptyVers
    If Not IsMissing(ModDate) Then If Len(ModDate) = 0 Then ModDate = cEmptyDate
    If Not IsMissing(ModHist) Then If Left(ModHist, Len(vbCrLf)) = vbCrLf Then ModHist = Mid(ModHist, Len(vbCrLf) + 1)
HandleExit:   ModuleInfoFromFile = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function ModuleVersGetFromFile(ModulePath As String) As String: Call ModuleInfoFromFile(ModulePath, ModVers:=ModuleVersGetFromFile): End Function
Public Function ModuleDateGetFromFile(ModulePath As String) As String: Call ModuleInfoFromFile(ModulePath, ModDate:=ModuleDateGetFromFile): End Function
Public Function ModuleAuthGetFromFile(ModulePath As String) As String: Call ModuleInfoFromFile(ModulePath, ModAuth:=ModuleAuthGetFromFile): End Function
Public Function ModuleDescGetFromFile(ModulePath As String) As String: Call ModuleInfoFromFile(ModulePath, ModDesc:=ModuleDescGetFromFile): End Function
Public Function ModuleCommGetFromFile(ModulePath As String) As String: Call ModuleInfoFromFile(ModulePath, ModComm:=ModuleCommGetFromFile): End Function
Public Function ModuleHistGetFromFile(ModulePath As String): Call ModuleInfoFromFile(ModulePath, ModHist:=ModuleHistGetFromFile): End Function

Public Function ModuleVersSet( _
    ModuleName As String, _
    VersionString As String, _
    Optional CodeLine As Long)
Const c_strProcedure = "ModuleVersSet"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLine = CodeLineFind(ModuleName, c_strPrefModVers, CodeLine) ' маркер на начало области вставки
    If CodeLine > 0 Then
        objModule.DeleteLines CodeLine, 1:
    Else
        CodeLine = 1 ' нет версии создаем в первой строке
    End If
    objModule.InsertLines CodeLine, c_strPrefModVers & " " & VersionString
    Result = True
HandleExit:     ModuleVersSet = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "Неверно задано имя объекта!"
    Case m_errModuleIsActive: Debug.Print "Невозможно изменить активный модуль: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "Модуль: """ & ModuleName & """ не найден!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function ModuleDateSet( _
    ModuleName As String, _
    DateString As String, _
    Optional CodeLine As Long)
Const c_strProcedure = "ModuleDateSet"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLine = CodeLineFind(ModuleName, c_strPrefModDate, CodeLine) ' маркер на начало области вставки
    If CodeLine > 0 Then
        objModule.DeleteLines CodeLine, 1:
    Else
        CodeLine = 1 ' нет даты создаем в первой строке
    End If
    objModule.InsertLines CodeLine, c_strPrefModDate & " " & DateString
    Result = True
HandleExit:     ModuleDateSet = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "Неверно задано имя объекта!"
    Case m_errModuleIsActive: Debug.Print "Невозможно изменить активный модуль: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "Модуль: """ & ModuleName & """ не найден!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function ModuleDescSet( _
    ModuleName As String, _
    DescString As String, _
    Optional CodeLine As Long)
Const c_strProcedure = "ModuleDescSet"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLine = CodeLineFind(ModuleName, c_strPrefModDesc, CodeLine) ' маркер на начало области вставки
    If CodeLine > 0 Then
        objModule.DeleteLines CodeLine, 1:
    Else
        CodeLine = 1 ' нет описания создаем в первой строке
    End If
    objModule.InsertLines CodeLine, c_strPrefModDesc & " " & DescString
    Result = True
HandleExit:     ModuleDescSet = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "Неверно задано имя объекта!"
    Case m_errModuleIsActive: Debug.Print "Невозможно изменить активный модуль: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "Модуль: """ & ModuleName & """ не найден!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function ModuleCommSet( _
    ModuleName As String, _
    CommString As String, _
    Optional Replace As Boolean = False, _
    Optional CodeLine As Long)
' записывает в заголовок модуля комментарий
Const c_strProcedure = "ModuleCommSet"
' если Replace = True - заменяет текущий, иначе добавляет под заголовком новый
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
Dim NumLines As Long: NumLines = 1
    CodeLine = CodeLineFind(ModuleName, c_strPrefModComm, CodeLine) ' маркер на начало области вставки
    If CodeLine > 0 Then
        NumLines = CodeLineNext(ModuleName, CodeLine) - CodeLine
        If Replace Then objModule.DeleteLines CodeLine, NumLines: NumLines = 1
    Else
        CodeLine = 1: Replace = True ' нет комментария создаем в первой строке
    End If
    With objModule
        If Replace Then .InsertLines CodeLine, c_strPrefModComm
        CodeLine = CodeLine + NumLines
        .InsertLines CodeLine, "'" & CommString
    End With
    Result = True
HandleExit:     ModuleCommSet = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "Неверно задано имя объекта!"
    Case m_errModuleIsActive: Debug.Print "Невозможно изменить активный модуль: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "Модуль: """ & ModuleName & """ не найден!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function ModuleAuthSet(ModuleName As String, AuthString As String, Optional CodeLine As Long)
' устанавливает данные автора модуля
Const c_strProcedure = "ModuleAuthSet"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLine = CodeLineFind(ModuleName, c_strPrefModAuth, CodeLine) ' маркер на начало области вставки
    If CodeLine > 0 Then
        objModule.DeleteLines CodeLine, 1:
    Else
        CodeLine = 1 ' нет комментария создаем в первой строке
    End If
    objModule.InsertLines CodeLine, c_strPrefModAuth & " " & AuthString
    Result = True
HandleExit:     ModuleAuthSet = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "Неверно задано имя объекта!"
    Case m_errModuleIsActive: Debug.Print "Невозможно изменить активный модуль: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "Модуль: """ & ModuleName & """ не найден!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function ModuleDebugSet( _
    ModuleName As String, _
    DEBUGGING As Boolean, _
    Optional CodeLine As Long)
'#Const DEBUGGING = False
Const c_strProcedure = "ModuleDebugSet"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLine = CodeLineFind(ModuleName, c_strPrefModDebg, CodeLine) ' маркер на начало области вставки
    If CodeLine > 0 Then
        objModule.DeleteLines CodeLine, 1:
    Else
        CodeLine = 1 ' нет комментария создаем в первой строке
    End If
    objModule.InsertLines CodeLine, c_strPrefModDebg & "=" & DEBUGGING
    Result = True
HandleExit:     ModuleDebugSet = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "Неверно задано имя объекта!"
    Case m_errModuleIsActive: Debug.Print "Невозможно изменить активный модуль: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "Модуль: """ & ModuleName & """ не найден!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function

'======================
Public Function CodeLineNext(ModuleName As String, Optional BegLine As Long = 1) As Long
' возвращает номер следующей строки в тексте с учетом знаков переноса
Const c_strProcedure = "CodeLineNext"
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLineNext = p_CodeLineNext(objModule, BegLine)
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "Неверно задано имя объекта!"
    Case m_errModuleIsActive: Debug.Print "Невозможно изменить активный модуль: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "Модуль: """ & ModuleName & """ не найден!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function CodeLineGet(ModuleName As String, Optional BegLine As Long = 1, Optional NumOfLines As Long) As String
' возвращает полную строку в тексте за вычетом знаков переноса
Const c_strProcedure = "CodeLineGet"
' через NumOfLines передает количество строк которые были изначально
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLineGet = p_CodeLineFull(objModule, BegLine, NumOfLines)
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "Неверно задано имя объекта!"
    Case m_errModuleIsActive: Debug.Print "Невозможно изменить активный модуль: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "Модуль: """ & ModuleName & """ не найден!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function CodeLineFind(ModuleName As String, FindString As String, Optional BegLine As Long = 1, Optional FindNum As Integer = 0) As Long
' Ищет строку кода в модуле ModuleName начинающуюся с FindString возвращает номер строки
Const c_strProcedure = "CodeLineFind"
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLineFind = p_CodeLineFind(objModule, FindString, BegLine, FindNum)
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "Неверно задано имя объекта!"
    Case m_errModuleIsActive: Debug.Print "Невозможно изменить активный модуль: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "Модуль: """ & ModuleName & """ не найден!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_CodeLineNext(objModule As Object, Optional BegLine As Long = 1) As Long
' возвращает номер следующей строки в тексте с учетом знаков переноса
Const c_strProcedure = "p_CodeLineNext"
Dim Result As Long: Result = False
    On Error GoTo HandleError
Dim i As Long: i = BegLine
    With objModule
        Do While VBA.Right$(.Lines(i, 1), Len(c_strHyphen)) = c_strHyphen
            i = i + 1
        Loop
    End With
    Result = i + 1
HandleExit:     p_CodeLineNext = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_CodeLineFull(objModule As Object, _
    Optional BegLine As Long = 1, Optional NumOfLines As Long, _
    Optional TrimSpaces As Boolean, _
    Optional ReplaceHyphensWith As String = c_strSpace, _
    Optional LinePrefToWrap As String) As String
' возвращает полную строку в тексте, объединяя части строки разделенные символом переноса
Const c_strProcedure = "p_CodeLineFull"
' через NumOfLines передает количество строк которые были изначально
Dim Result As String
    On Error GoTo HandleError
    NumOfLines = 0
Dim tmpString As String
Dim LineBreak As Boolean, LineType As m_CodeLineType
    With objModule
        Do
            tmpString = VBA.Trim$(.Lines(BegLine + NumOfLines, 1))
            Result = Result & p_CodeLineGet(tmpString, LineType, LineBreak, ReplaceHyphensWith)
            NumOfLines = NumOfLines + 1
        Loop While LineBreak
    End With
' перед окончанием сжимаем строку - удаляем мусор
    Result = Replace(Result, c_strBrokenQuotes, vbNullString)   ' объединяем разорванные текстовые строки
HandleExit:  p_CodeLineFull = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_CodeLineGet(CodeLine As String, _
    Optional LineType As m_CodeLineType, _
    Optional LineBreak As Boolean, _
    Optional ReplaceHyphensWith As String = c_strSpace, _
    Optional TrimPrefix As Boolean = True, _
    Optional TrimSpaces As Boolean) As String
' возвращает отдельную строку текста кода, удаляет префиксы и знаки переноса, возвращает флаги типа строки
Const c_strProcedure = "p_CodeLineGet"
' CodeLine - обрабатываемая строка
' LineType - тип строки кода (опредееляется по содержимому)
' LineBreak - признак продолжения чтения строки (строка заканчивается символом переноса)
' ReplaceHyphensWith - символ для замены символов переноса строки
' TrimPrefix - признак необходимости удаления префиксов строк
' TrimSpaces - признак необходимости сжимать пробелы вначале/конце строки
Dim Result As String: Result = CodeLine
    On Error GoTo HandleError
Static arrPrefs(): arrPrefs = Array(m_CodeVers, c_strPrefModVers, _
                                    m_CodeDate, c_strPrefModDate, _
                                    m_CodeDesc, c_strPrefModDesc, _
                                    m_CodeAuth, c_strPrefModAuth, _
                                    m_CodeComm, c_strPrefModComm, _
                                    m_CodeHist, c_strPrefModHist, _
                                    m_CodeProc, c_strCodeProcBeg)
Dim i As Long
Dim strPref As String, lngLineType As Long
' если это продолжение строки - удаляем пустой префикс вначале строки
    If LineBreak Then
        Select Case LineType
        Case m_CodeDesc, m_CodeComm, m_CodeHist: strPref = c_strPrefModNone
        End Select
        GoTo HandleProceed
    End If
' иначе пытаемся понять что это за строка
    Select Case LineType
    Case m_CodeHead
' ищем полезную информацию в заголовке модуля
        For i = LBound(arrPrefs) To UBound(arrPrefs) Step 2
            lngLineType = arrPrefs(i): strPref = arrPrefs(i + 1)
            Select Case lngLineType
            Case m_CodeProc: If InStrRegEx(1, Result, strPref) > 0 Then LineType = lngLineType: strPref = vbNullString: Exit For
            Case m_CodeHist: If InStrRegEx(1, Result, strPref) > 0 Then LineType = lngLineType: strPref = "'": Exit For
            Case Else:       If VBA.Left$(Result, Len(strPref)) = strPref Then LineType = lngLineType: Exit For
            End Select
            strPref = vbNullString ' если не нашли обнуляем префикс
        Next i
    Case m_CodeProc
' ищем полезную информацию в процедуре
'Stop
    Case Else
    End Select
'Stop
HandleProceed:
' проверяем признак переноса строки
    LineBreak = (VBA.Right$(Result, Len(c_strHyphen)) = c_strHyphen)
' если удаляем префиксы и строка начинается с указанного префикса - удалить его
    If TrimPrefix And Len(strPref) > 0 Then Result = VBA.Trim$(VBA.Mid$(Result, Len(strPref) + 1))
' если удаляем пробелы и строка начинается/заканчивается на пробелы - обрезаем
    If TrimSpaces Then Result = Trim$(Result)
' если строка заканчивается на  _ заменить на символ объединения и продолжить чтение
    If LineBreak Then Result = VBA.Left$(Result, Len(Result) - Len(c_strHyphen)) & ReplaceHyphensWith
HandleExit:     p_CodeLineGet = Result: Exit Function
HandleError:    Result = vbNullString
    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_CodeLineFind(objModule As Object, FindString As String, Optional BegLine As Long = 1, Optional FindNum As Integer = 0) As Long
' Ищет строку кода в модуле ModuleName начинающуюся с FindString возвращает номер строки
Const c_strProcedure = "p_CodeLineFind"
Dim Result As Long ':Result = 0
    On Error GoTo HandleError
Dim lngBegLine As Long, lngEndLine As Long: lngBegLine = BegLine
Dim lngBegCol As Long, lngEndCol As Long
Dim FindCount As Long ': FindCount = 0
Dim bolIsLoaded As Boolean
    With objModule
        Do ' маркер на начало
            If .Find(FindString, lngBegLine, lngBegCol, 0, -1) Then
                If lngBegCol = 1 Then   ' ищем с начала строки
                    Result = lngBegLine
                    FindCount = FindCount + 1
                    Select Case FindNum
                    Case 0:  Exit Do    ' первое найденное
                    Case -1:            ' последнее найденное
                    Case Else: If FindCount = FindNum Then Exit Do ' n-ное найденное
                    End Select
                End If
                lngBegLine = lngBegLine + 1 ' если не вышли из цикла - переходим на следующую строку и продолжаем поиск
            Else
                Exit Do
            End If
        Loop
    End With
HandleExit:     p_CodeLineFind = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function

'-------------------------
' функции для работы с zip архивами
'-------------------------
Public Function ZipCreate( _
    ZipName As String _
    ) As Boolean
' создает ZIP файл
' возвращает True если создан
Dim strZipFileHeader As String
Dim Result As Boolean
    On Error GoTo HandleError
    ' Проверка наличия расширения zip в полном пути-имени файла
    If VBA.UCase$(oFso.GetExtensionName(ZipName)) <> c_strObjExtZip Then Exit Function
    ' Создание пустого zip архива
    strZipFileHeader = "PK" & VBA.Chr(5) & VBA.Chr(6) & VBA.String$(18, 0)
    oFso.OpenTextFile(ZipName, 2, True).Write strZipFileHeader
    Dim oArch As Object: Set oArch = oApp.Namespace((ZipName))
    ' проверка создания архива
    Result = Not (oArch Is Nothing)
HandleExit:  ZipCreate = Result: Exit Function
HandleError: Result = False: Resume HandleExit
End Function
Public Function ZipPack( _
    FilePath As String, _
    Optional ZipName As String, _
    Optional DelAfterZip As Boolean = False _
    ) As Boolean
' добавляет файлы к архиву
Const c_strProcedure = "ZipPack"
' FileNames - имена архивируемых файлов/папок через точку с запятой
' FilePath - путь к папке где находятся архивируемые файлы
' ZipName - имя файла архива (полный путь к архиву)
' v.1.0.1       : 24.10.2022 - изменён способ контроля завершения архивирования элемента
Const DelayAfterZip = 333
Const iTryMax = 3
Dim iTry As Integer
Dim i As Long, iMax As Long
Dim strFilePath As String, strFilename As String
Dim strZipPath As String, strZipName As String ', strFileName As String
Dim Result As Boolean
    On Error GoTo HandleError
' подготовка путей и создание архива
    If Len(FilePath) = 0 Then Err.Raise vbObjectError + 512
    
    With oFso
        strFilePath = .GetParentFolderName(FilePath)
        strFilename = .GetFileName(FilePath)
        If Len(ZipName) > 0 Then
            strZipPath = .GetParentFolderName(ZipName)
            strZipName = .GetFileName(ZipName)
        Else
    ' если путь к файлу архива не задан берем имя родительской папки пути FilePath
            strZipPath = .GetParentFolderName(strFilePath)
            strZipName = .GetBaseName(strFilePath) & "." & c_strObjExtZip
            ZipName = .BuildPath(strZipPath, strZipName)
        End If
    ' если указанный файл архива существует переходим к упаковке
        Result = .FileExists(ZipName): If Result Then GoTo HandlePack
    ' указанный файл архива не существует
        ' создаем указанный путь
        If Not oFso.FolderExists(strZipPath) Then Call oFso.CreateFolder(strZipPath)  ' Then Err.Raise 76 ' Path not Found
        ZipName = .BuildPath(strZipPath, strZipName)
        ' создаем файл архива
        Result = ZipCreate(ZipName): If Not Result Then GoTo HandleExit ' если не удалось создать - выходим по ошибке
    End With
HandlePack:
' собственно архивирование
    Dim oItm As Object, oZip As Object, lItm As Long, sItm As String
    
    Set oZip = oApp.Namespace((ZipName))
    For Each oItm In oApp.Namespace((strFilePath)).Items
        If oItm.IsFolder Then
        ' если это папка
            ' получаем количество файлов в папке
            sItm = oItm.Name: lItm = oItm.GetFolder.Items.Count
            ' если папка пуста - переходим к следущему объекту
            If lItm = 0 Then GoTo HandleNext
        End If
        ' перемещаем файлы и папки
        oZip.MoveHere (oItm.path), 4 + 8 + 16 + 1024
        ' ожидаем окончание сжатия файлов
        Do
            Sleep DelayAfterZip: DoEvents: DoEvents
            If oItm.IsFolder Then Result = oApp.Namespace((ZipName & "\" & sItm)).Items.Count = lItm
        Loop Until Result
HandleNext:
    Next oItm
    Sleep DelayAfterZip: DoEvents: DoEvents
HandleDelAfterZip:
    If DelAfterZip Then Result = oFso.DeleteFolder(strFilePath) = 0
    Set oItm = Nothing: Set oZip = Nothing
HandleExit:     ZipPack = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case 91:    If iTry < iTryMax Then iTry = iTry + 1: Err.Clear: Sleep 333: Resume 0
                Message = "Ошибка при упакоке объектов проекта"
                If Len(sItm) > 0 Then Message = Message & " sItm=""" & sItm & """ "
                'Stop: Err.Clear: Resume 0
    Case 70, 76: If iTry < iTryMax Then iTry = iTry + 1: Err.Clear: Sleep 333: Resume Next
                Message = "Не удалось полностью удалить папку."
                If Len(strFilePath) > 0 Then Message = Message & " FilePath=""" & strFilePath & """ "
                'Stop: Err.Clear: Resume 0
    Case 10094: Err.Clear: Resume 0: 'Несмотря на то что файлы были успешно добавлены в архив, компоненту “Сжатые ZIP-папки" не удалось полностью удалить оригиналы (убедитесь, что файлы не имеют защиты пс записи, и что вы имеете разрешение на их удаление)
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
    Err.Clear: Resume 0
End Function
Public Function ZipUnPack( _
    ZipFile As String, _
    Optional FilePath As String = vbNullString, _
    Optional Overwrite As Boolean = True _
    ) As Boolean
' извлекает архив в указаную папку
Const c_strProcedure = "ZipUnPack"
Dim strPath As String, strZipName As String
Dim Result As Boolean
On Error GoTo HandleError
' Определяем наличие архива для извлечения
    If oFso.FileExists(ZipFile) = False Then
    ' если файл отсутствует - выходим
        MsgBox "System could not find " & ZipFile & " Unpack cancelled.", vbInformation, "Error Unziping File"
        GoTo HandleExit 'Exit Function
    End If
    strZipName = oFso.GetFileName(ZipFile)
' Определяем путь извлечения
    If Len(FilePath) = 0 Then
    ' если задан пустой путь - извлекаем во временную папку
        strPath = CurrentProject.path & "\" & c_strSrcPath & "\"
    Else
        If oFso.FolderExists(FilePath) Then
            strPath = oFso.GetFolder(FilePath).path 'oFso.BuildPath(FilePath)
        Else
            Err.Raise 53, , "Folder not found"
        End If
    End If
' Извлекаем архив в созданую директорию
    Dim oZip As Object:  Set oZip = oApp.Namespace((ZipFile & "\"))
    Dim oItm As Object
  
    With oZip
        If Overwrite Then
            For Each oItm In .Items
            ' если файл или папка уже существуют - удаляем перед извлечением
                If oFso.FileExists(strPath & oItm.Name) Then
                    Kill strPath & oItm.Name
                ElseIf oFso.FolderExists(strPath & oItm.Name) Then
                    Kill strPath & oItm.Name & "\*.*"
                    RmDir strPath & oItm.Name
                End If
            Next
        End If
        oApp.Namespace(CVar(strPath)).CopyHere .Items, 4 + 16
    End With
'    'если указано удалять архив после извлечения - удаляем
'    If DelAfterUnZip Then Kill ZipFile
    Result = True
HandleExit:     On Error Resume Next
    ZipUnPack = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit: Resume
End Function
Public Function ZipItemsCount(ZipName As String) As Integer
' возвращает количество файлов в архиве
' ZipName - полный путь к архиву
Dim oItms As Object
Dim Result As Integer
    On Error GoTo HandleError
Dim oFld As Object: Set oFld = oApp.Namespace((ZipName))
    Result = oFld.Items().Count
HandleExit:  ZipItemsCount = Result: Exit Function
HandleError: Result = -1: Resume HandleExit
End Function
Public Function ZipItemName( _
    ZipName As String, _
    Optional i As Integer = 0, _
    Optional fExt As Boolean = True) As String
' Возвращает имя i-того файла в архиве
' ZipName - имя архива
' i - номер файла в архиве (начало с 0), по умолчанию - 0
' fExt - включать расширение в имя файла, по умолчанию - true
Dim Result As String
    On Error GoTo HandleError
Dim oFld As Object: Set oFld = oApp.Namespace((ZipName))
    With oFld.Items().Item((i)):  Result = IIf(fExt, .path, .Name): End With
HandleExit:  ZipItemName = Result: Exit Function
HandleError: Result = vbNullString: Resume HandleExit
End Function
'-------------------------
' функции сохранения отдельных параметров проекта
'-------------------------
Private Function p_LinkedRead(LocalName As String, FilePath As String) ', Optional TableName)
' читает параметры и создает связанную таблицу из файла FilePath
Const c_strProcedure = "p_LinkedRead"
Dim strName As String, Connect As String, TableName As String, Attributes  As Long
    On Error GoTo HandleError
    LocalName = VBA.Trim$(LocalName): If Len(LocalName) = 0 Then LocalName = VBA.Trim$(p_SettingKeyRead(FilePath, c_strLnkSecParam, c_strLnkKeyLocal))
    TableName = VBA.Trim$(p_SettingKeyRead(FilePath, c_strLnkSecParam, c_strLnkKeyTable))
    Connect = VBA.Trim$(p_SettingKeyRead(FilePath, c_strLnkSecParam, c_strLnkKeyConnect))
    Attributes = VBA.Trim$(p_SettingKeyRead(FilePath, c_strLnkSecParam, c_strLnkKeyAttribute))
' создает прилинкованую таблицу с указаными параметрами
    If Len(TableName) = 0 Or Len(Connect) = 0 Then GoTo HandleExit
    If Len(LocalName) = 0 Then LocalName = TableName
    Dim tdf As Object 'dao.TableDef
    Set tdf = CurrentDb.CreateTableDef(LocalName) 'это псевдоним таблицы
    With tdf
        .Connect = Connect ' строка подключения к базе на сервере,будет не лишним понять что внутри
        .SourceTableName = TableName ' имя таблицы источника
        '.Attributes = Attributes
    End With
' добавляем линк на таблицу
    With CurrentDb.TableDefs: .Append tdf: .Refresh: End With: Set tdf = Nothing
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_LinkedWrite(TableName As String, FilePath As String)
' сохранет в файл FilePath параметры присоединенной таблицы TableName
Const c_strProcedure = "p_LinkedWrite"
    On Error GoTo HandleError
    p_SettingKeyWrite FilePath, c_strLnkSecParam, c_strLnkKeyLocal, CurrentDb.TableDefs(TableName).Name
    p_SettingKeyWrite FilePath, c_strLnkSecParam, c_strLnkKeyConnect, CurrentDb.TableDefs(TableName).Connect
    p_SettingKeyWrite FilePath, c_strLnkSecParam, c_strLnkKeyTable, CurrentDb.TableDefs(TableName).SourceTableName
    p_SettingKeyWrite FilePath, c_strLnkSecParam, c_strLnkKeyAttribute, CurrentDb.TableDefs(TableName).Attributes
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function

Private Function p_ReferencesDropBroken() As Boolean
' тестовая функция
On Error Resume Next
Dim refs As Access.References, ref As Access.Reference, i As Integer, fBroken As Boolean
Dim wshShell As Object, strGUID As String, strKey As String
  
    Set refs = Access.References
    'Проход по ссылкам в обратном порядке
    For i = refs.Count To 0 Step -1
        Set ref = refs(i)
        fBroken = ref.IsBroken: If Err.Number Then fBroken = True: Err.Clear ' Err.Number=48
        If Not fBroken Then GoTo HandleNext
        If ref.BuiltIn Then fBroken = True: Exit For
        'Первая попытка удаления битой ссылки
        refs.Remove ref
        If Err.Number = 0 Then GoTo HandleNext
        Err.Clear
        If ref.Kind = 1 Then fBroken = True: GoTo HandleNext
        'Если возникла ошибка при удалении ссылки на библиотеку, методами WSH пытается
        'добавить в реестр ветку с GUID и версией из битой ссылки, удалить ссылку,
        'а затем удалить ветку.
        If wshShell Is Nothing Then Set wshShell = CreateObject("WScript.Shell")
        If Err.Number <> 0 Then fBroken = True: GoTo HandleNext
        strGUID = ref.GUID
        strKey = "HKCR\TypeLib\" & strGUID & "\" & ref.Major & "." & ref.Minor & "\"
        wshShell.RegWrite strKey & "0\win32\", ""
        'Вторая попытка удаления ссылки (она, типа, зарегистрирована)
        refs.Remove ref
        If Err.Number <> 0 Then Err.Clear: fBroken = True
        
        wshShell.RegDelete strKey & "0\win32\"
        wshShell.RegDelete strKey & "0\"
        wshShell.RegDelete strKey
        
        'Пытается создать ссылку на самую свежую зарегистрированную библиотеку.
        refs.AddFromGuid strGUID, 0, 0
        Err.Clear
HandleNext:
    Next i
    '  Next ref
    p_ReferencesDropBroken = fBroken
HandleExit:
    Set ref = Nothing
    Set refs = Nothing
    Set wshShell = Nothing
    Exit Function
End Function
Private Function p_ReferencesRead(FilePath As String) ', Optional TableName)
' читает и восстанавливает ссылки (References) из файла FilePath
Const c_strProcedure = "p_ReferencesRead"
    On Error GoTo HandleError
    Dim strText As String: strText = p_SettingSecRead(FilePath, c_strRefSecName)
    If Len(strText) = 0 Then Err.Raise 76 ' Path not Found
    Dim aRefs() As String: aRefs = Split(strText, ";")
    Dim ref As Reference
' проверяем все имеющиеся ссылки и удаляем если они нарушены
    On Error Resume Next
    Dim Broken As Boolean: Broken = False
    For Each ref In References
' не может удалить ошибочную ссылку на MSComctlLib после запуска Win64
        Err.Clear
        If ref.BuiltIn Then GoTo HandleRemoveNext
        If Not ref.IsBroken Then GoTo HandleRemoveNext
        strText = ref.GUID      ' запоминаем GUID
        References.Remove ref
    ' удаляем ссылку
        DoEvents
    ' если удалось удалить - пытаемся восстановить из GUID
        If Err.Number = 0 Then References.AddFromGuid strText, 0, 0: If Err.Number = 0 Then GoTo HandleRemoveNext
''        'Если возникла ошибка при удалении ссылки на библиотеку, методами WSH пытается
''        'добавить в реестр ветку с GUID и версией из битой ссылки, удалить ссылку,
''        'а затем удалить ветку.
''        Err.Clear
''        If oWsh Is Nothing Then Set oWsh = CreateObject("WScript.Shell")
''        If Err.Number <> 0 Then fBroken = True: GoTo HandleNext
''        strRegKey = c_strRegKey & strGUID & "\" & lngMajor & "." & lngMinor & "\0\win32\"
''        oWsh.RegWrite strRegKey, ""
''        'Вторая попытка удаления ссылки (она, типа, зарегистрирована)
''        refs.Remove ref
''        If Err.Number <> 0 Then Err.Clear: fBroken = True
''
''        oWsh.RegDelete c_strRegKey 'c_strRegKey & strGUID & "\" & iMajor & "." & iMinor & "\0\win32\"
''        oWsh.RegDelete c_strRegKey & strGUID & "\" & lngMajor & "." & lngMinor & "\0\"
''        oWsh.RegDelete c_strRegKey & strGUID & "\" & lngMajor & "." & lngMinor & "\"
'' если все равно не удалось восстановить - придется делать это руками
        Broken = Broken Or Err.Number: Err.Clear
HandleRemoveNext:
    Next ref
' восстанавливаем ссылки из файла
    Dim Itm, aRef() As String ', strName As String, strDesc As String
    On Error Resume Next
    For Each Itm In aRefs
        Err.Clear
    ' получаем кодовое имя библиотечной ссылки
        aRef = Split(Itm, "=")
    ' проверяем ее наличие
        Set ref = References(aRef(0)): If Err.Number = 0 Then GoTo HandleNext
        Err.Clear
    ' получаем ее параметры из значения прочитанной переменной
        aRef = Split(aRef(1), "|") ' GUID|Major|Minor|FullPath
        If Err.Number = 0 Then
            Set ref = References.AddFromGuid(aRef(0), aRef(1), aRef(2))  ' GUID|Major|Minor
'            Set ref = References.AddFromFile(aRef(3))' FullPath
        End If
HandleNext:
    Next Itm
    If Not Broken Then GoTo HandleExit
' не все ссылки удалось исправить - потребуется участие пользователя
    Dim strTitle As String, strMessage As String
    strTitle = "Внимание!"
    strMessage = "При восстановлении сылок некоторые ссылки автоматически восстановить не удалось." & vbCrLf & _
        "Для восстановления ссылок вручную необходимо войти в редактор VBA (Alt+F11)," & vbCrLf & _
        "открыть меню Tools\References, снять галочку со ссылок отмеченных MISSING," & vbCrLf & _
        "и перезапустить данный макрос."
    MsgBox Title:=strTitle, Prompt:=strMessage, Buttons:=vbOKOnly + vbInformation
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_ReferencesWrite(FilePath As String)
' сохранет в файл FilePath ссылки (References)
Const c_strProcedure = "p_ReferencesWrite"
' Сохраняем так:
' [References]
'   Name=GUID|Major|Minor|FullPath
    On Error GoTo HandleError
    Dim Itm As Object
    For Each Itm In References
' пропускаем встроенные и ошибочные ссылки
        If Itm.BuiltIn Then GoTo HandleNext
        'If Itm.IsBroken Or Err.Number <> 0 Then Else GoTo HandleNext
' сохраняем в виде: Name=GUID|Major|Minor|FullPath
        With Itm: p_SettingKeyWrite FilePath, c_strRefSecName, .Name, Join(Array(.GUID, .Major, .Minor, .FullPath), "|"): End With
HandleNext:
    Next Itm
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_PropertiesRead(FilePath As String)
' читает и восстанавливает свойства проекта из файла FilePath
Const c_strProcedure = "p_PropertiesRead"
    On Error GoTo HandleError
' восстанавливаем свойства проекта
    With VBE.ActiveVBProject
        .Name = p_SettingKeyRead(FilePath, c_strPrjSecName, c_strPrjKeyName)
        .Description = p_SettingKeyRead(FilePath, c_strPrjSecName, c_strPrjKeyDesc)
        .HelpFile = p_SettingKeyRead(FilePath, c_strPrjSecName, c_strPrjKeyHelp)
    End With
Dim strText As String
Dim aPrps() As String, Itm
Dim strName As String, strValue As String, intType As eDataType
' восстанавливаем пользовательские свойства
    strText = p_SettingSecRead(FilePath, c_strPrpSecName): If Len(strText) = 0 Then Err.Raise 76 ' Path not Found
    ' удаляем все старые свойства
    With CurrentProject.Properties
        Do While .Count > 0: .Remove .Item(0).Name: Loop
    ' заполняем прочитанные свойства
    aPrps = Split(strText, ";")
    For Each Itm In aPrps
        Itm = VBA.Trim$(Itm)
        If Len(Itm) > 0 Then
            strValue = p_PropertyStringRead(CStr(Itm), PropName:=strName)
            If Len(strName) > 0 Then .Add strName, strValue
        End If
    Next Itm
    End With
' восстанавливаем свойства базы данных
    strText = p_SettingSecRead(FilePath, c_strDbsSecName): If Len(strText) = 0 Then Err.Raise 76 ' Path not Found
    ' заполняем прочитанные свойства
    aPrps = Split(strText, ";")
On Error Resume Next
Dim prp As DAO.Property
    With CurrentDb
    For Each Itm In aPrps
    ' читаем строку параметров
        Itm = VBA.Trim$(Itm)
        If Len(Itm) > 0 Then
            strValue = p_PropertyStringRead(CStr(Itm), PropName:=strName, PropType:=intType)
            If Len(strName) > 0 Then PropertySet strName, strValue, CurrentDb, intType
        End If
    Next Itm
    End With
Err.Clear
On Error GoTo HandleError
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_PropertiesWrite(FilePath As String)
' сохранет в файл FilePath свойства проекта
Const c_strProcedure = "p_PropertiesWrite"
    On Error GoTo HandleError
' сохраняем свойства проекта
    With VBE.ActiveVBProject
        p_SettingKeyWrite FilePath, c_strPrjSecName, c_strPrjKeyName, p_PropertyStringCreate(.Name, PropType:=dbText)
        p_SettingKeyWrite FilePath, c_strPrjSecName, c_strPrjKeyDesc, p_PropertyStringCreate(.Description, PropType:=dbText)
        p_SettingKeyWrite FilePath, c_strPrjSecName, c_strPrjKeyHelp, p_PropertyStringCreate(.HelpFile, PropType:=dbText)
    End With
Dim Itm As Object, strName As String, varValue, intType As eDataType: intType = dbText
' сохраняем пользовательские свойства
    With CurrentProject
    For Each Itm In .Properties
        With Itm:  strName = .Name: varValue = .Value: End With
        varValue = p_PropertyStringCreate(varValue, PropType:=intType)
        p_SettingKeyWrite FilePath, c_strPrpSecName, strName, CStr(varValue)
    Next Itm
    End With
' сохраняем свойства базы данных
On Error Resume Next
    Dim i As Long
    With CurrentDb
    For Each Itm In .Properties
        Err.Clear
        With Itm
            strName = .Name
            Select Case strName
            ' пропускаем ненужные свойства по имени
            Case "DesignMasterID", "Name", "Transactions", "Updatable", "CollatingOrder", _
                 "Version", "RecordsAffected", "ReplicaID", "Connection", "AccessVersion", _
                 "Build", "ProjVer"
                GoTo HandleNext
            End Select
            varValue = .Value: intType = .Type
        ' для проверки доступности поля для записи пытаемся записать в него значение и ловим ошибку
            .Value = varValue
        End With
        ' если ошибка - свойство невозможно будет записать - нет необходимости его сохранять
        If Err.Number Then Debug.Print strName & " - " & Err.Number & " " & Err.Description: Err.Clear: GoTo HandleNext
        ' если нет ошибки - сохраняем свойство для последующего восстановления
        varValue = p_PropertyStringCreate(varValue, PropType:=intType)
        p_SettingKeyWrite FilePath, c_strDbsSecName, strName, CStr(varValue)
HandleNext:
    Next Itm
    End With
    Err.Clear
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_PropertyStringRead(PropString As String, _
    Optional PropName, Optional PropType As eDataType _
    ) As Variant
' читает значение параметра свойства из строки
Const c_strProcedure = "p_PropertyStringRead"
Dim Result

    Result = PropString: If Len(Result) = 0 Then GoTo HandleExit
    On Error GoTo HandleError
    If Not IsMissing(PropName) Then
        PropName = Left$(PropString, InStr(1, PropString, "=") - 1)
        Result = Mid$(PropString, InStr(1, PropString, "=") + 1): If Len(Result) = 0 Then GoTo HandleExit
    End If
    If (Len(Result) > 1) And ((Left$(Result, 1) = """") And (Right$(Result, 1) = """")) Then
        PropType = dbText: Result = Mid$(Result, 2, Len(Result) - 2)
    ElseIf IsNumeric(Result) Then
        PropType = dbNumeric: Result = Val(Result)
    ElseIf (Len(Result) > 1) And ((Left$(Result, 1) = "#") And (Right$(Result, 1) = "#")) Then
        PropType = dbDate: Result = CDate(Mid$(Result, 2, Len(Result) - 2))
    ElseIf UCase(Result) = "TRUE" Then
        PropType = dbBoolean: Result = CBool(True)
    ElseIf UCase(Result) = "FALSE" Then
        PropType = dbBoolean: Result = CBool(False)
    Else
    Dim strType As String: strType = UCase(Left$(Result, InStr(1, Result, ":") - 1))
    If Len(strType) > 0 Then Result = Mid$(Result, InStr(1, Result, ":") + 1)
        Select Case strType
        Case "DECIMAL": PropType = dbDecimal:   Result = CDec(Result)
        Case "BYTE":    PropType = dbByte:      Result = CByte(Result)
        Case "SINGLE":  PropType = dbSingle:    Result = CSng(Result)
        Case "DOUBLE":  PropType = dbDouble:    Result = CDbl(Result)
        Case "INTEGER": PropType = dbInteger:   Result = CInt(Result)
        Case "LONG":    PropType = dbLong:      Result = CLng(Result)
        Case "CURRENCY": PropType = dbCurrency: Result = CCur(Result)
        Case "FLOAT":   PropType = dbFloat:     Result = Val(Result)
        Case Else:      Err.Raise vbObjectError + 512
        End Select
    End If
HandleExit:  p_PropertyStringRead = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Private Function p_PropertyStringCreate(PropValue As Variant, _
    Optional PropName As String, Optional PropType As eDataType = dbText _
    ) As String
' формирует строку значения для свойства
Const c_strProcedure = "p_PropertyStringCreate"
Dim Result 'As String

    On Error GoTo HandleError
    Select Case PropType
    Case dbText:    Result = """" & CStr(PropValue) & """"
'    Case dbChar
'    Case dbMemo
    Case dbBoolean: Result = IIf(CBool(PropValue), "True", "False")
    Case dbNumeric: Result = Val(PropValue)
    Case dbDecimal: Result = "Decimal:" & CDec(PropValue)
    Case dbByte:    Result = "Byte:" & CByte(PropValue)
    Case dbSingle:  Result = "Single:" & CSng(PropValue)
    Case dbDouble:  Result = "Double:" & CDbl(PropValue)
    Case dbInteger: Result = "Integer:" & CInt(PropValue)
    Case dbLong:    Result = "Long:" & CLng(PropValue)
    Case dbCurrency: Result = "Currency:" & CCur(PropValue)
    Case dbFloat:   Result = "Float:" & CDec(PropValue)
'    Case dbBigInt
'    Case dbBinary
'    Case dbLongBinary
'    Case dbVarBinary
'    Case dbGUID
    Case dbDate:    Result = "#" & CDate(PropValue) & "#"
'    Case dbTime:    Result = "#" & CDate(PropValue) & "#"
'    Case dbTimeStamp
    Case Else: Err.Raise vbObjectError + 512
    End Select
    If Len(PropName) > 0 Then Result = PropName & "=" & Result
HandleExit:  p_PropertyStringCreate = Result:  Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
'-------------------------
' функции для работы с ini
'-------------------------
Private Function p_SettingKeyRead(strFile As String, strSection As String, strRegKeyName As String) As String
' возвращает строковое значение из INI файла
Const c_strProcedure = "p_SettingKeyRead"
' strSection - имя секции
' strRegKeyName - имя искомого параметра
' strFile - путь к ini файлу
Dim strBuffer As String * 4096
Dim intSize As Integer

    On Error GoTo HandleError
    intSize = GetPrivateProfileString(strSection, strRegKeyName, "", strBuffer, 4096, strFile)
    p_SettingKeyRead = VBA.Left$(strBuffer, intSize)
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_SettingKeyWrite(strFile As String, strSection As String, strRegKeyName As String, strValue As String) As Integer
' сохраняет значение параметра в файл
Const c_strProcedure = "p_SettingKeyWrite"
' strSection - имя секции
' strRegKeyName - имя искомого параметра
' strValue - значение параметра
' strFile - путь к ini файлу
' возвращает True если успешно, иначе - False
Dim intStatus As Integer

    On Error GoTo HandleError
    intStatus = WritePrivateProfileString(strSection, strRegKeyName, strValue, strFile)
    p_SettingKeyWrite = (intStatus <> 0)

HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_SettingSecRead(strFile As String, strSection As String) As String
' читает имена и значения ключей в заданной секции из файла .INI
Const c_strProcedure = "p_SettingSecRead"
'возвращает: Param1=Val1;Param2=Val2...
' strSection - имя секции
' strFile - путь к ini файлу
Dim strBuffer As String * 4096
Dim intSize As Integer
    intSize = GetPrivateProfileSection(strSection, strBuffer, 4096, strFile)
    p_SettingSecRead = Replace(VBA.Left$(strBuffer, intSize), VBA.Chr$(0), ";")
End Function
'-------------------------
' вспомогательные функции
'-------------------------
Private Function p_SelectFile(Optional InitPath As String, _
    Optional FileMask As String = "*.*", Optional Extention As String = "*", _
    Optional DialogTitle As String) As Variant
' диалог выбора файла
Const c_strProcedure = "p_SelectFile"
Dim of As OPENFILENAME
Dim Result: Result = vbNullString
    On Error GoTo HandleError
    With of
        .lStructSize = Len(of)
        .hwndOwner = hWndAccessApp
        .lpstrInitialDir = InitPath
        .lpstrFilter = FileMask
        .nFilterIndex = 1
        .lpstrFile = VBA.String$(512, 0)
        .nMaxFile = 511
        .lpstrDefExt = Extention
        .lpstrTitle = DialogTitle
    End With
    If GetOpenFileName(of) Then Result = StrZ(of.lpstrFile)
HandleExit:  p_SelectFile = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Private Function p_SelectFolder( _
    Optional InitPath As String = vbNullString, _
    Optional InitView As Integer = msoFileDialogViewList, _
    Optional DialogTitle As String _
    ) As Variant
' диалог выбора папки
Const c_strProcedure = "p_SelectFolder"
Dim Result As Variant: Result = vbNullString
    On Error GoTo HandleError
'
'Dim fRet As Long, bi As BROWSEINFO, dwIList As Long
'Dim szPath As String, wPos As Integer
'
'    With bi
'        .hOwner = hWndAccessApp
'        .lpszTitle = DialogTitle
'        .ulFlags = BIF_RETURNONLYFSDIRS
'    End With
'
'    dwIList = SHBrowseForFolder(bi)
'    szPath = Space$(512)
'    fRet = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
'
'    If fRet Then Result = StrZ(szPath)
Dim InitFolder As String
' SHBrowseForFolder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = DialogTitle
        .InitialView = InitView
        If Len(InitPath) > 0 Then
            If dir(InitPath, vbDirectory) <> vbNullString Then
                InitFolder = InitPath
                If Right(InitFolder, 1) <> "\" Then
                    InitFolder = InitFolder & "\"
                End If
                .InitialFileName = InitFolder
            End If
        End If
        .Show
        On Error Resume Next
        Err.Clear
        Result = .SelectedItems(1)
        If Err.Number <> 0 Then Result = vbNullString
    End With
HandleExit:  p_SelectFolder = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Private Function p_TextAlpha2Code(Source As String)
' Заменяет символы не входящие в список допустимых их шеснадцатиричным Asc кодом вида %XX
Dim c As Long, cMax As Long
Dim PermissedSymb As String
Dim Char As String
Dim Result As String

    Result = vbNullString
' задаем разрешенные символы
    PermissedSymb = VBA.UCase$(c_strSymbDigits & c_strSymbRusAll & c_strSymbEngAll & c_strOthers)
    c = 1: cMax = Len(Source)
    Do Until c > cMax
        Char = VBA.Mid$(Source, c, 1)
        If InStr(1, PermissedSymb, VBA.UCase$(Char)) = 0 Then
            Char = VBA.Hex$(Asc(Char))
            If Len(Char) < c_strCodeLen Then Char = VBA.String$(c_strCodeLen - Len(Char), "0") & Char
            Char = c_strCodeSym & Char
        End If
        Result = Result & Char
        c = c + 1
    Loop
HandleExit: p_TextAlpha2Code = Result
End Function
Private Function p_TextCode2Alpha(Source As String)
' Заменяет шеснадцатиричный Asc код вида %XX, символов не входящих в список допустимых, их значением
Dim c As Long, cMax As Long
Dim i As Byte, tmpChar As String
Dim Char As String
Dim Result As String
    Result = vbNullString
    c = 1: cMax = Len(Source)
    Do Until c > cMax
        Char = VBA.Mid$(Source, c, 1)
        If Char = c_strCodeSym Then
            tmpChar = VBA.UCase$(VBA.Mid$(Source, c + 1, c_strCodeLen))
            i = 1
            Do Until i > c_strCodeLen
                Select Case VBA.Mid$(tmpChar, i, 1)
                Case "0" To "9", "A" To "F"
                Case Else: GoTo HandleSkip
                End Select
                i = i + 1
            Loop
            c = c + c_strCodeLen
            Char = VBA.Chr$(Val(c_strHexPref & tmpChar))
HandleSkip:
        End If
        Result = Result & Char
        c = c + 1
    Loop
HandleExit: p_TextCode2Alpha = Result
End Function
Private Function InStrRegEx( _
    Start As Long, _
    String1 As String, String2 As String, _
    Optional Found As String, _
    Optional Compare As VbCompareMethod = vbTextCompare) As Long
' InStr позволяющий искать совпадения по маске выражений RegEx
Const c_strProcedure = "p_InstrRegEx"
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
'Static oRegEx As Object: If oRegEx Is Nothing Then Set oRegEx = CreateObject("VBScript.RegExp")
    ' вызов RegExp и передача ему маски
    With oRegEx: .IgnoreCase = (Compare = vbTextCompare): .Pattern = String2: Set oMatches = .Execute(S1): End With
    If oMatches.Count = 0 Then GoTo HandleError
    For Each oMatch In oMatches
        Result = oMatch.FirstIndex + Start: Found = oMatch
        Exit For
    Next
HandleExit:  InStrRegEx = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
'-------------------------
' функции вызова статических объектов
'-------------------------
Private Function oFso() As Object
Static soFSO As Object: If soFSO Is Nothing Then Set soFSO = CreateObject("Scripting.FileSystemObject")
    Set oFso = soFSO
End Function
Private Function oApp() As Object
Static soApp As Object: If soApp Is Nothing Then Set soApp = CreateObject("Shell.Application")
    Set oApp = soApp
End Function
Private Function oRegEx() As Object
Static soRegEx As Object: If soRegEx Is Nothing Then Set soRegEx = CreateObject("VBScript.RegExp")
    Set oRegEx = soRegEx
End Function
'-------------------------
' вспомогательные функции
'-------------------------
Private Function StrZ(Par As String) As String
Dim nSize As Long, i As Long ', Rez As String
   nSize = Len(Par)
   i = InStr(1, Par, VBA.Chr(0)) - 1
   If i > nSize Then i = nSize
   If i < 0 Then i = nSize
   StrZ = VBA.Mid$(Par, 1, i)
End Function
#If USEZIPCLASS Then
Private Function oZip() As clzZipArchive
Static soZip As clzZipArchive: If soZip Is Nothing Then Set soZip = New clzZipArchive
    Set oZip = soZip
End Function
#End If
'-------------------------
' неиспользуемые функции
'-------------------------
Private Function p_WinUserName() As String
' имя пользователя Windows
Const c_strProcedure = "p_WinUserName"
Dim sBuffer As String, lSize As Long
Dim Result As String
    On Error GoTo HandleError
    sBuffer = VBA.Space$(255): lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    Result = vbNullString
    If lSize > 0 Then Result = VBA.Left$(sBuffer, InStr(sBuffer, VBA.Chr(0)) - 1)
    'If lSize > 0 Then Result = VBA.Left$(sBuffer, lSize)
HandleExit:     p_WinUserName = Result: Exit Function
HandleError:    Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Private Function p_GetSysDiskSerial() As String
' простая проверка серийного номера диска
Dim Result As String
Dim VolLabel As String, VolSize As Long, Serial As Long, MaxLen As Long
Dim strName As String, NameSize As Long, Flags As Long
Const DiskName = "C:\"

    Result = VBA.String(8, "0")
    On Error GoTo HandleError
    If GetVolumeInformation(DiskName, VolLabel, VolSize, Serial, MaxLen, Flags, strName, NameSize) _
        Then Result = VBA.Format$(VBA.Hex$(Serial), VBA.String$(8, "0"))
HandleExit:     p_GetSysDiskSerial = Result: Exit Function
HandleError:    Result = VBA.String$(8, "0"): Err.Clear: Resume HandleExit
End Function




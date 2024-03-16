Attribute VB_Name = "~App"
Option Compare Database
Option Base 1
'=========================
Private Const c_strModule As String = "~App"
'=========================
' ��������      :
' ������        : 1.0.0.0
' ����          : 01.07.2016 15:00:00
' �����         : ������ �.�. (KashRus@gmail.com)
' ����������    :
'=========================
' ��������� �����������
'=========================
'-------------------------
' �������� �����
'-------------------------
Public Enum AppColorScheme
    appColorGrey = &H969696
    appColorLightGrey = &H767676
' �������� �����
    appColorDark = &H993333 '&H732A0A             ' ������ ����
    appColorBright = &HDFA000           ' ����� ����
    appColorLight = &HE5D1C0            ' ������� ����
' ������ 1
    appColorDark2 = &H730A1F
    appColorBright2 = &HDF3000
    appColorLight2 = &HE5C0C1
' ������ 2
    appColorDark3 = &H730A53
    appColorBright3 = &HDF003F
    appColorLight3 = &HE5C0D4
End Enum
'-------------------------
' �����
'-------------------------
Public Const appFontNameDef = "Arial"
Public Const appFontSizeDef = 10
Public Const appFontSize1 = 8
Public Const appFontSize2 = 12
Public Const appFontSize3 = 18
'-------------------------

'=========================
' ��������� �����������
'=========================
'-------------------------
Private Const c_strApplication = "��������"
Private Const c_strAuthor As String = "������ �.�."
Private Const c_strVersion As String = "2.0.0"
Private Const c_strSupport As String = "KashRus@gmail.com"
Private Const c_strShowClock As Boolean = True
'-------------------------
'Public Const cpHash = "SHA256", HashType = eHashSHA256 '
'-------------------------
' �������� �������� ������
'SysBookmarks, SysFields, SysLog, SysMenu, SysObjData, SysObjTypes, SysOrderTypes, SysUsers, SysVersion
Public Const c_strSysObjects = "MSysObjects"    ' ��������� �������
Public Const c_strTableVers = "SysVersion"      ' �������� ������ ����������
Public Const c_strTableUser = "SysUsers"        ' �������� ������ �������������
Public Const c_strTableMenu = "SysMenu"         ' �������� ������� ����
Public Const c_strTableData = "SysObjData"      ' �������� ��������� ��������
Public Const c_strTableLogs = "SysLog"          ' �������� ��������� ������
Public Const c_strTableDate = "SysCalendar"     ' ���������� ���
'Public Const c_strTableBook = "SysBookmarks"    ' ���������� �������� �������
'Public Const c_strTableFlds = "SysFields"       ' ���������� ������������ �����
'Public Const c_strTableTObj = "SysObjTypes"     ' ���������� ����� ��������
'Public Const c_strTableTOrd = "SysOrderTypes"   ' ���������� ����� ��������

'-------------------------
' ����� ������� �������
'-------------------------
' �������� ����������
Public Const c_strPropAppName = "Application" '= c_strApplication ' ="��������"
Public Const c_strPropAuthor = "Author"
Public Const c_strPropSupport = "Support"
Public Const c_strPropVersion = "Version"
Public Const c_strPropVerDate = "VersionDate"
Public Const c_strPropLastDate = "LastDate"
Public Const c_strPropLastUser = "LastUserName"
Public Const c_strPropFirstRun = "FirstRun"
Public Const c_strPropShowClock = "ShowClock"
' �������� ����� ����������
Public Const c_strPropSrvPath = "SrvPath"
Public Const c_strPropDatPath = "DatPath"
Public Const c_strPropSecPath = "SecPath"
Public Const c_strPropLogPath = "LogPath"
Public Const c_strPropDllPath = "DllPath"
Public Const c_strPropDocPath = "DocPath"
Public Const c_strPropTmpPath = "TmpPath"
' �������� ���������� ������ ��� ���������� ����������
Public Const c_strDesignRes = "DesignRes" ' 1280x1024
Public Const c_strDesignDpi = "DesignDpi" '
Public Const c_strResDelim = "x" '
'-------------------------
' �������� ������� �������
'-------------------------
Public Const c_strLastUserName As String = "�������������" ' ������������� �� SysUsers
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
' ������
'-------------------------
Public Const c_strAppIco = "App"
Public Const c_strMenuIco = "ContextMenu"
'-------------------------
' ��� �������� ����������� ��� ����������
'-------------------------
' ������� ������ � ��������� ������� �������
Private Const strBegLineMarker = "'=== BEGIN INSERT ==="
Private Const strEndLineMarker = "'==== END INSERT ===="
' ���� �� ��� ���� �� ������ � DoEvents �� ��������� = 333
Public Const appDoEventsPause = 100

Public Const c_strTagDelim = "_"
Public Const c_strDelim = ";"
Public Const c_strInDelim = ","
' �������� ���������� SQL
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
Public Const sqlSimilar = "SIMILAR"  ' ������������� - �������� �����
Public Const sqlIs = " IS ", sqlNull = "NULL", sqlTrue = "True", sqlFalse = "False"
Public Const sqlIsNull = sqlIs & sqlNull, sqlIsNotNull = sqlIs & sqlNot & sqlNull
' �������� �������� ������
'SysBookmarks, SysFields, SysLog, SysMenu, SysObjData, SysObjTypes, SysOrderTypes, SysUsers, SysVersion
'Private Const c_strSysObjects = "MSysObjects"    ' ��������� �������

' �������� �������������� ����������
Public Const c_strParamType = "Type"
Public Const c_strParamMode = "Mode"
Public Const c_strParamKey = "Key"
'-------------------------
' �������� �������� Access
'-------------------------
' AccessObjectType
Public Const c_strTmpTypePref = "tmp" ' ��������������� ������
Public Const c_strTmpTablPref = "@&%" ' ��������� �������

' �������� ��� ����������� � �������� "On[...]" ������� ��� ��������� ��� �������
Public Const c_strCustomProc = "[Event Procedure]"
Public Const c_strCmdMnuProc = "ContextMenu_Click"
' ��������� �������� ��� ��������� ��������� �������
Public frmDROP_Date_Controls As Collection ' ��������� ��������� ����� frmDROP_Date
' FormType
Public Const c_strMenuType = "MENU" ' ����
' ���
Public Const c_strServType = "SERV" ' ���������
Public Const c_strDropType = "DROP" ' ���������� �����
' FormMode
Public Const c_strMainMode = "MAIN" ' ��� ���� - ��������
' ��� ���������� ���� FormType=c_strDropType
Public Const c_strRealMode = "Real" ' ���������� ���.��������
Public Const c_strCalcMode = "Calc" ' ���������� �����������
Public Const c_strDateMode = "Date" ' ���������� ���������
' ��� ��������� ���� FormType=c_strServType
Public Const c_strUserMode = "User"
Public Const c_strUChgMode = "UserChg"
Public Const c_strFloat = "Float"   ' ��������� ������
Public Const c_strNavBar = "NavBar" ' ������ ��������� �� �������
Public Const c_strPrtBar = "PrtBar" ' ������ �� ������������ �������
' ��� �������� ����������� �� ����� �����������

'==============================
Public Enum appErrors
' ���������������� ������ ����������
    errAuthGrant = vbObjectError + 1000     ' ����� �����
    errAuthError = vbObjectError + 1001     ' ������ �����������
    errAuthFailed = vbObjectError + 1002    ' ����� �����������
    errAuthEnd = vbObjectError + 1009       ' ����� ��������
    errAppNoConn = vbObjectError + 1010     ' ����������� ����������� � ������� ������
    errAppNoData = vbObjectError + 1011     ' ����������� ������
    errAppClose = vbObjectError + 1109      ' ������ �������� ����������
    errAppPathWrong = vbObjectError + 1110  ' ������ �������� ���� ����������
End Enum
Public Enum enmPathType
' ���� ����� ���������� ������������
    enmPathUndef = 0 ' �� ���������
    enmPathAll = 255 ' ��� ����
    enmPathSrv = 1  ' ���� � ������� ������ ����������
    enmPathDll = 2  ' ���� � ����� ������� ��������� ����������
    enmPathTmp = 3  ' ���� � ����� �������� ������� ����������
    enmPathDoc = 4  ' ���� � ����� ������� ����������
    enmPathDat = 5  ' ���� � ����� ������
    enmPathSec = 6  ' ���� � ����� ������� ������
    enmPathLoc = 7  ' ���� � ��������� ���� (����������)
    enmPathLog = 8  ' ���� � ����� ���������� ����������
    enmPathSrc = 9  ' ���� � ����� ����� ����������
    enmPathDbf = 10 ' ���� � ����� ��������
End Enum
Public Enum appUserType
' ���� ���� ������������
    appUserTypeAdmin = 100
    appUserTypeUser = 200
End Enum
Public Enum appModeType
' ������������� ������� ����������
    appModeDebug = -1                   ' ����� �������
    appModeNormal = 0                   ' ������� �����
End Enum
Public Enum appRecState
' �������� ��������� ������ (SPReal)
    appRecStateTemp = -1 '��������� - �� ����������
    appRecStateReal = 0  '�������� - ����������
    appRecStateOld = 10  '������ - ���� ��������, ���� ����� ����������
    appRecStateArc = 11  '�������� - ���� ��������, ����� ���������� ���
    appRecStateDen = 91  '���������
    appRecStateDel = 99  '�������� ���������
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
'Public Function App(): Static myApp As New clsApp: Set App = myApp: End Function                ' ���������� ������ �� ����� �������� ����������
'Public Function Cmd(): Static myCmd As New clsCommands: Set Cmd = myCmd: End Function           ' ���������� ������ �� ����� �������� �������� ����������
'Public Function Dbg(): Static myDebug As New clsDebug: Set Dbg = myDebug: End Function          ' ���������� ������ �� ����� ������ �������
'Public Function Crypto(): Static myCrypto As New clsCrypto: Set Crypto = myCrypto: End Function ' ���������� ������ �� ����� ������ ����������
'Public Function Fso(): Static myFso As Object: Set Fso = CreateObject("Scripting.FileSystemObject"): End Function ' ���������� ������ �� ����� ������ ��������� �������/������
'Public Function Wdr(): Static myWord As New clsWordReport: Set Wdr = myWord: End Function ' ���������� ������ �� ����� ������� Word
'======================
Public Function StartApp(): Call App.AppStart(True): End Function
Public Function StopApp(): App.AppStop (False): End Function
Public Function UpdateAppPath(): Call App.UpdatePath: End Function
Public Function UpdateAppMode(): App.ModeSwitch: End Function
Public Function UpdateLocalRefs(): App.UpdateRefs: End Function
Public Function Setup()
' ��������� ������� ��� �������������� ���������
' ��������� ����������
'    RestoreRefs     ' �������������� ������ �� ����������
    RestoreProp     ' ��������� ������� ����������
    CloseAll        ' ��������� ��������� ��� �������
    CompileAll      ' ����������
End Function
Public Sub RestoreProp()
Const c_strProcedure = "RestoreProp"
'������������� ������� ����������
On Error GoTo HandleError
    SetOption ("Auto Compact"), True            ' ������� ��� ������
    SetOption ("ShowWindowsInTaskbar"), True    ' ��������� ���� � ������ �����
'    SetOption ("Show Status Bar"), False        ' ��������� ������ ���������
    With CurrentProject.Properties
'=== BEGIN INSERT ===
        .Add "Application", c_strApplication ' ="��������"
        .Add "Version", c_strVersion ' =""
        .Add "Author", c_strAuthor   ' ="������ �.�."
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
' ���������� ������� ������������ ����
    With Application.CommandBars.ActionControl
Stop
Debug.Print .Tag, .Caption
    End With
End Sub



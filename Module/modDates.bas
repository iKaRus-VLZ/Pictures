Attribute VB_Name = "modDates"
Option Compare Database
Option Explicit
'=========================
Private Const c_strModule As String = "modDates"
'=========================
' ��������      : ������� ��� ������ � ������
' ������        : 1.0.7.453565659
' ����          : 05.03.2024 13:34:54
' �����         : ������ �.�. (KashRus@gmail.com)
' ����������    : ��� �������� ��� ���������� ������� SysCalendar, ����� ���������� ���������������� ��������� � buh.ru
' v.1.0.7       : 20.02.2024 - ��������� ��������� �������������, ������������� � �������� ������� �� ������ ��������� �������
' v.1.0.5       : 07.02.2024 - ���������� ������� ������ � �������� ����� ��� �������� �������� � ������������ � Excel 2010+
' v.1.0.4       : 27.12.2023 - ��������� ���������� ����������������� ��������� ����� �������� � buh.ru
' v.1.0.2       : 22.03.2019 - ��������� ������� ������� ���������� �������/�������� ���� ����� ������
'=========================
' ������ ������ ��������� ����������� �� ����, ����� ���� ������ ���������� ������ � �������� �����,
' ������� ��� ���������� ���� ���������/�������������� ������� (�������� ����������)
' ���������� ���� ������� ������ ��� ������� ��������������� � �� �������,- � ����� �����������?))
' ��������� ������� (��-��������): �������, ���������������, �����������, ���������, ����������������
' �.�. ���� ��� ������ ��� ������� ������� ������� ���� �� ����� ��������� �������
'=========================
' ToDo: ������� ���������� �������� ��������� ������� ����� � �������������
'       ������� ����� �������������� ������� � �������������� ������ ������� (��. p_DatesTableEventEdit, p_DatesTableEventsList)
'=========================
Private Const NOERROR As Long = 0
' ������ ��� �������� ����������������� ��������� (��. p_HolidaysFromWeb)
Private Const �_strLink = "https://buh.ru/calendar/" ' "https://www.consultant.ru/law/ref/calendar/proizvodstvennye/"
' �������� ������� ����������� � �������� ����
Private Const c_strDatesTable = "SysCalendar"                                           ' �������� ������� ��������� ���
Private Const c_strKey = "ID", c_strParent = "PARENT"                                   ' ��������/������������ ����
Private Const c_strDateType = "DATETYPE", c_strDateDesc = "DATEDESC"                    ' ���/�������� ����
Private Const c_strDateBeg = "DATEBEG", c_strDateEnd = "DATEEND"                        ' ��� ������� � �������� (���� ������/�����)
Private Const c_strOffsetType = "OFFSET", c_strOffsetValue = c_strOffsetType & "VAL"    ' ��� ������������� (��� ��������/�������� ��������)
Private Const c_strPeriodType = "PERIOD", c_strPeriodValue = c_strPeriodType & "VAL"    ' ��� ������������� (��� �������/�������� �������)
Private Const c_strActBegDate = "ACTBEG", c_strActEndDate = "ACTEND"                    ' ���� ������/��������� ������������ ������
Private Const c_strEditDate = "EDITDATE", c_strComment = "COMMENT"                      ' ���� ��������� ������/�����������

Private Const c_strTmpTablePref = "@&%"                                                  ' ������� ��������� �������
Private m_datTempBeg As Date, m_datTempEnd As Date                                      ' ������ �� ������� ������������ ��������� ������� ���������� ����� �������
Private m_strTempName As String                                                         ' ��� ��������� ������� ���������� ����� ������� �� ��������� �������

' �������������� �������� Interval ��� ������� DateDiffEx/DateAddEx
Public Const c_strWorkdayLiteral = "wd", c_strWorkdayLiteral2 = "workday"          ' ������� ���
Public Const c_strNonWorkdayLiteral = "hd", c_strNonWorkdayLiteral2 = "holiday"    ' ��������� ��� ??
    ' ��� ������
Public Const c_strMondayLiteral = "mon", c_strMondayLiteral2 = "monday", _
      c_strMondayLiteral1 = "��", c_strMondayLiteral3 = "�����������"       ' ������������
Public Const c_strTuesdayLiteral = "tues", c_strTuesdayLiteral2 = "tuesday", _
      c_strTuesdayLiteral1 = "��", c_strTuesdayLiteral3 = "�������"         ' ��������
Public Const c_strWednesdayLiteral = "wed", c_strWednesdayLiteral2 = "wednesday", _
      c_strWednesdayLiteral1 = "��", c_strWednesdayLiteral3 = "�����"       ' �����
Public Const c_strThursdayLiteral = "thur", c_strThursdayLiteral2 = "thursday", _
      c_strThursdayLiteral1 = "��", c_strThursdayLiteral3 = "�������"       ' ��������
Public Const c_strFridayLiteral = "fri", c_strFridayLiteral2 = "friday", _
      c_strFridayLiteral1 = "��", c_strFridayLiteral3 = "�������"           ' �������
Public Const c_strSaturdayLiteral = "sat", c_strSaturdayLiteral2 = "saturday", _
      c_strSaturdayLiteral1 = "��", c_strSaturdayLiteral3 = "�������"       ' �������
Public Const c_strSundayLiteral = "sun", c_strSundayLiteral2 = "sunday", _
      c_strSundayLiteral1 = "��", c_strSundayLiteral3 = "�����������"       ' �����������
' ��� ��� �������� ��������� �� �����������  (��.DateTypes)
Public Enum eDateType
    eDateTypeUndef = 0              ' 0   <�� ������>
    ' ����������� ���
    eDateTypeWeekday = 1            ' 1   ������� (������) ����
    eDateTypeSatday = 6             ' 6   �������� (�������) ����
    eDateTypeSunday = 7             ' 7   �������� (�����������) ����
    ' �������
    eDateTypeWorkday = 10           ' 10  ������� ����
    eDateTypeHolidayPre = 20        ' 20  ������� (�����������) ����
    eDateTypeHoliday = 70           ' 70  �������� (�����������) ����
    eDateTypeNonWorkday = 80        ' 80  ��������� ����
    eDateTypeUser = 99              ' 99  ���� ����������� �������������
    ' ������ �������
    eDateServOfficial = 500         ' 500 (��������� ������) ��������������� ���������, �������� ���� � ������ ���� �������� ������� � ���������������
    eDateServWorkCalendar = 900     ' 900 (��������� ������) ���������� � �������� ������ ����������������� ���������
End Enum

Private Type POINT
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
#If Win64 Then
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare PtrSafe Function ScreenToClient Lib "user32.dll" (ByVal hwnd As LongPtr, ByRef lpPoint As POINT) As Long
Private Declare PtrSafe Function ClientToScreen Lib "user32.dll" (ByVal hwnd As LongPtr, ByRef lpPoint As POINT) As Long
#Else
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINT) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINT) As Long
#End If
'=========================
' �������� ������� ����� Excel � ����������� ������� ��� ������ � ������
'=========================
Public Function AGE(BirthDate As Date, Optional TestDate) As Long
' ���������� ������� (���������� ������ ���) �������� �� ��������� ����
'-------------------------
' BirthDate - ���� ��������
' TestDate  - ���� �� ������� ������������ �������
'-------------------------
Dim Result As Long ': Result = False
    On Error GoTo HandleError
    If Not IsDate(TestDate) Then TestDate = Date
    If TestDate >= BirthDate Then Result = DateDiff("yyyy", BirthDate, TestDate) + (DateSerial(Year(TestDate), Month(BirthDate), Day(BirthDate)) > TestDate)
HandleExit:  AGE = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function MONTHEND(Date1 As Date) As Date
' ���������� ���� ����� ������ ��� �������� ����
'-------------------------
    MONTHEND = DateSerial(Year(Date1), Month(Date1) + 1, 1) - 1
End Function

Public Function WEEKDAYDATE(Date1 As Date, Number As Double, Weekday1 As VbDayOfWeek) As Date
' ���������� ����, ��������� �� �������� ���������� ��������� ���� ������ ������ ��� ����� �� ��������� ����.
'-------------------------
' Date1     - ��������� ���� �� ������� ����������� ������
' Weekday1  - ������� ���� ������
' Number    - ���������� ��������� ���� ������, �� ������� ������ �������� ������� ���� �� ���������
'   1,2..   - ���� ������ ����� ��������� ���� (������� �)
'   0,-1..  - ���� ������ ����� ��������� ����� (�� ������� �)
'-------------------------
' ��������: ���� �������������� ����������� �������� ������: WEEKDAYDATE (MONTHEND(Now)+1,-1,vbSunday)
'-------------------------
'' ���� ���� - ������� ��������� ByVal � �����������������:
'    If Number < 0 Then Number = Number + 1 ' ��� ������� ����� �������� � -1
'    If Number <= 0 Then Date1 = Date1 + 1  ' ��� ������� ����� ��������� ��������� ����
    WEEKDAYDATE = DateAdd("ww", Number, Date1 - 1) - (Date1 - Weekday1 - 1) Mod 7
End Function

Public Function WEEKDAYCOUNT(ByVal Date1 As Date, ByVal Date2 As Date, Weekday1 As VbDayOfWeek) As Long
' ���������� ���������� ��������� ���� ������ ����� ��������� ����� � �������� �����.
'-------------------------
' Date1     - ��������� ���� �������
' Date2     - �������� ���� �������
' Weekday1  - ������� ���� ������
'-------------------------
Dim Sign As Boolean, Temp As Date: If Date1 > Date2 Then Sign = True: Temp = Date2: Date2 = Date1: Date1 = Temp ' ��������� ������������������ ���������� ��� � ������ ����� ������� �������, � ������ �������
    Date2 = Date2 - Weekday1 + 1: If Weekday1 >= Weekday(Date1) Then Date2 = Date2 + 7
    WEEKDAYCOUNT = DateDiff("ww", Date1, Date2): If Sign Then WEEKDAYCOUNT = WEEKDAYCOUNT * -1
'' ��� ���?
'    If Date1 <= Date2 Then
'        WEEKDAYCOUNT = DateDiff("ww", Date1, Date2 - Weekday1 + IIf(Weekday1 < Weekday(Date1), 1, 8))
'    Else
'        WEEKDAYCOUNT = DateDiff("ww", Date1 - Weekday1 + IIf(Weekday1 < Weekday(Date2), 1, 8), Date2)
'    End If
End Function

Public Function ISWORKDAY(Date1 As Date, _
        Optional Holidays = False, Optional Weekends, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Boolean
' ���������� �������� ���� �������. ��������� ���� �� �������� �������� �������� � ����������
'-------------------------
' Date1     - ����������� ����
' Holidays  - ������������� ������ ����������� ��� (����������� � ��������� ����, � ����� ������� �������� ����), ����:
'             0 - ���� ����� �������� �� �������, 1 - �� ���������, ����� ������ - ��������� �� ����� �����������
' Weekends  - ����� ��� ������ �������� ��������� � �� ��������� �������� (��. p_Weekends). ��-���������: �������, �����������
'-------------------------
Dim Result As Boolean: On Error GoTo HandleError
    If IsMissing(Holidays) Then
' ��������� ���� �� ������ ��������
        Result = Not ISWEEKEND(Date1, Weekends)
    Else
' ��������� �� Holidays
        Select Case p_CheckForHolidays(Date1, Date1, Holidays, Weekends, dbs, wks)
        Case 1:     Result = True                           ' +1 - �������
        Case -1:    Result = False                          ' -1 - ��������
        Case Else:  Result = Not ISWEEKEND(Date1, Weekends) ' �� ��������� (� ������ ������) - ������ ������ Holidays
        End Select
    End If
HandleExit:  ISWORKDAY = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Function ISWEEKEND(Date1, Optional Weekends) As Boolean
' ���������� �������� ���� ��������. ��������� ���� �� ��������� ������� �������� (�� ��������� ��������� � �� ��������� ���)
'-------------------------
' Date1     - ����������� ����
' Weekends  - ����� ��� ������ �������� ��������� � �� ��������� �������� (��. p_Weekends). ��-���������: �������, �����������
'-------------------------
Dim Result As Boolean: On Error GoTo HandleError
Dim uw As Long: uw = UBound(p_Weekends(Weekends)) ' ��� ������������� �������������� ������ �������� ���� � ���� ������� �������
Dim i As Long, d As VbDayOfWeek: d = DatePart("w", Date1): Do: Result = (p_Weekends()(i) = d): i = i + 1: Loop While i <= uw And Not Result   ' ��������� ���� �� ������ ��������
HandleExit:  ISWEEKEND = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Sub SetWeekends(Optional Weekends)
' ������ ����� ���� ������, ������� �������� ��������� � �� ��������� ��������.
'-------------------------
' Weekends  - ����� ��� ������ �������� ��������� � �� ��������� ��������. ��-���������: �������, �����������
'   �������� �������� 1-7, 11-17
'   ��� ������ ����: 0000011, ��� 0- ������� ����, 1-�������� (������� � ������������)
'-------------------------
    If IsMissing(Weekends) Then Weekends = 1 Else If Len(Weekends) = 0 Then Weekends = 1
    Call p_Weekends(Weekends)
End Sub

Public Function NETWORKDAYS(Date1 As Date, Date2 As Date, _
        Optional Holidays = False, Optional Weekends, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Long
' ���������� ���������� ������� ���� ����� ��������� ����� � �������� �����.
'-------------------------
' Date1     - ��������� ���� �������
' Date2     - �������� ���� �������
' Holidays  - ������������� ������ ����������� ��� (����������� � ��������� ����, � ����� ������� �������� ����), ����:
'             0 - ���� ����� �������� �� �������, 1 - �� ���������, ����� ������ - ��������� �� ����� �����������
' Weekends  - ����� ��� ������ �������� ��������� � �� ��������� �������� (��. p_Weekends). ��-���������: �������, �����������
'-------------------------
' �������� ����� �� ��������� �������� ��� (����� ��������� � Holidays ��� ������� ��������) � ��� ������������ � Holidays ��� �����������
'-------------------------
Dim Result As Long: On Error GoTo HandleError
Dim Sign As Boolean, Temp As Date: If Date1 > Date2 Then Sign = True: Temp = Date2: Date2 = Date1: Date1 = Temp ' ��������� ������������������ ���������� ��� � ������ ����� ������� �������, � ������ �������
Dim uw As Long: uw = UBound(p_Weekends(Weekends))             ' UBound(p_Weekends)+1 = ���������� ��������� ���� � ������
' �������� ���������� �������/�������� ���� � ������ ������� (��� ����� ����������)
Dim ww As Long: ww = DateDiff("w", Date1, Date2)    ' ������ ������ � �������
    Result = (6 - uw) * ww  ': hd = (uw+1) * ww
' ����������� ������ ������� ��� ������ ������ � ��������� �������
    Temp = DateAdd("ww", ww, Date1)
    Do While Temp <= Date2 ' <=
        If Not ISWEEKEND(Temp) Then Result = Result + 1 ' ��������� � �������
        Temp = DateAdd("d", 1, Temp)                    ' ��������� ����
    Loop
' ���������� ����������� � ���������
    If IsMissing(Holidays) Then GoTo HandleExit
    Result = Result + p_CheckForHolidays(Date1, Date2, Holidays, Weekends, dbs, wks)
HandleExit:  If Sign Then Result = -Result              ' ���� ���� - ������ ����
             NETWORKDAYS = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Function WORKDAY(Date1 As Date, Number As Double, _
        Optional Holidays = False, Optional Weekends, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Date
' ���������� ����, ��������� �� �������� ���������� ������� ���� ������ ��� ����� �� ��������� ����.
'-------------------------
' Date1     - ��������� ���� �� ������� ����������� ������
' Number    - ���������� ������� ����, �� ������� ������ �������� ������� ���� �� ���������
' Holidays  - ������������� ������ ����������� ��� (����������� � ��������� ����, � ����� ������� �������� ����), ����:
'             0 - ���� ����� �������� �� �������, 1 - �� ���������, ����� ������ - ��������� �� ����� �����������
' Weekends  - ����� ��� ������ �������� ��������� � �� ��������� �������� (��. p_Weekends). ��-���������: �������, �����������
'-------------------------
' �������� ����� �� ��������� �������� ��� (����� ��������� � Holidays ��� ������� ��������) � ��� ������������ � Holidays ��� �����������
'-------------------------
' ������� ����������, �� ��������������
'-------------------------
Dim Result As Date: Result = Date1
    On Error GoTo HandleError
Dim d As Integer:   d = Sgn(Number)     ' ����������� ��������
Dim n As Double:    n = Abs(Number)     ' �������� �������� �� �������� ����
Dim ww As Long                          ' ���������� ������ ������ � �������
Dim wd As Long                          ' ���������� ������� ����
Dim WorkDaysPerWeek As Long:    WorkDaysPerWeek = 6 - UBound(p_Weekends(Weekends))  ' ���������� ������� ���� � ������ (��� ����� ����������)
    ww = n \ 7
    Do While n  '<> 0
    ' ������� ��� ����� ���������� � ��������� ���� (��������� ������ Weekends)
        If ww = 0 Then
    ' ������� ���� �� ���� ���� � ��������� �������� �� �� �������
            Result = DateAdd("d", d, Result):       wd = Abs(ISWORKDAY(Result, Holidays, Weekends, dbs, wks))
        Else
    ' ������� ���� �� ������ ���������� ������� ������
Dim db As Date: db = Result
            Result = DateAdd("ww", d * ww, Result): wd = ww * WorkDaysPerWeek
    ' ����������� ��������� � ��������� ���
            If Not IsMissing(Holidays) Then wd = wd + p_CheckForHolidays(db + d, Result, Holidays, Weekends, dbs, wks)
        End If
    ' ��������� ������� ������� ���� �������� �����
        n = n - wd: ww = n \ 7
    Loop
HandleExit:  WORKDAY = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Function DateAddEx(Interval As String, ByVal Number As Double, Date1 As Date, _
        Optional FirstDayOfWeek As VbDayOfWeek, Optional FirstWeekOfYear As VbFirstWeekOfYear, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Date
' ����������� DateAdd ����������� ������������ �������������� ���������
'-------------------------
' Interval..- ���������� ������������ DateDiff
' dbs,wks   - ������ �� ���� � ������� ���������� ������������  ���������
'-------------------------
Dim Result As Long
    On Error GoTo HandleError
    Select Case Interval
' ���� ��������� �� �������� ���������� �������/��������� ���� �� ��������� ����
    Case c_strWorkdayLiteral, c_strWorkdayLiteral2:                                                     Result = WORKDAY(Date1, Number)
    Case c_strNonWorkdayLiteral, c_strNonWorkdayLiteral2:                                               Result = DateAdd("d", Number, Date1) - WORKDAY(Date1, Number)
' ���� ��������� �� �������� ���������� ��������� ���� ������ �� ��������� ����
    Case c_strSundayLiteral, c_strSundayLiteral1, c_strSundayLiteral2, c_strSundayLiteral3:             Result = WEEKDAYDATE(Date1, Number, vbSunday)
    Case c_strMondayLiteral, c_strMondayLiteral1, c_strMondayLiteral2, c_strMondayLiteral3:             Result = WEEKDAYDATE(Date1, Number, vbMonday)
    Case c_strTuesdayLiteral, c_strTuesdayLiteral1, c_strTuesdayLiteral2, c_strTuesdayLiteral3:         Result = WEEKDAYDATE(Date1, Number, vbTuesday)
    Case c_strWednesdayLiteral, c_strWednesdayLiteral1, c_strWednesdayLiteral2, c_strWednesdayLiteral3: Result = WEEKDAYDATE(Date1, Number, vbWednesday)
    Case c_strThursdayLiteral, c_strThursdayLiteral1, c_strThursdayLiteral2, c_strThursdayLiteral3:     Result = WEEKDAYDATE(Date1, Number, vbThursday)
    Case c_strFridayLiteral, c_strFridayLiteral1, c_strFridayLiteral2, c_strFridayLiteral3:             Result = WEEKDAYDATE(Date1, Number, vbFriday)
    Case c_strSaturdayLiteral, c_strSaturdayLiteral1, c_strSaturdayLiteral2, c_strSaturdayLiteral3:     Result = WEEKDAYDATE(Date1, Number, vbSaturday)
' ����������� �����
    Case Else:                                                                                          Result = DateAdd(Interval, Number, Date1)
    End Select
HandleExit:  DateAddEx = Result: Exit Function
HandleError: Resume HandleExit
End Function

Public Function DateDiffEx(Interval As String, ByVal Date1 As Date, ByVal Date2 As Date, _
        Optional FirstDayOfWeek As VbDayOfWeek, Optional FirstWeekOfYear As VbFirstWeekOfYear, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Long
' ����������� DateDiff ����������� ������������ �������������� ���������
'-------------------------
' Interval..- ���������� ������������ DateDiff
' dbs,wks   - ������ �� ���� � ������� ���������� ������������  ���������
'-------------------------
' !!! �� ������ ������� ����� ��� ������ ������� �� ��������� ������������ DateDiff, �������,- ������� ��������:
' �������� ���� ������� �� ����������� (������� ���������� ��������� ����� ����������� �� ��������� ���������)
' ������� ��� �������� ���������� ���� ��� � Between Date1 And Date2 (������� �������) ���� ������ Date2 ���������� ������� ��������� DateAdd(Interval, 1, Date2)
' �.� ����� 01/01/2015 � 31/12/2015 ����� 364 ����������� ���,  260 ������� � 104 �������� (246/118 � ������ ����������), - �������� 31/12/2015 (������� ����)
' � �����   01/01/2016 � 31/12/2016 ����� 365 ����������� ����, 260 ������� � 105 �������� (247/118 � ������ ����������), - �������� 31/12/2016 (�������� ����)
'-------------------------
Dim Result As Long
    On Error GoTo HandleError
Dim ww As Long, wd As Long, hd As Long, cd As Date, sg As Integer ', dd As Long
    Select Case Interval
' ���������� �������/��������� ����  ����� ����� ������
    Case c_strWorkdayLiteral, c_strWorkdayLiteral2:                                                     Result = NETWORKDAYS(Date1, Date2)
    Case c_strNonWorkdayLiteral, c_strNonWorkdayLiteral2:                                               Result = DateDiff("d", Date1, Date2) - NETWORKDAYS(Date1, Date2)
' ���������� �������� ���� ������ ����� ����� ������
    Case c_strSundayLiteral, c_strSundayLiteral1, c_strSundayLiteral2, c_strSundayLiteral3:             Result = WEEKDAYCOUNT(Date1, Date2, vbSunday)
    Case c_strMondayLiteral, c_strMondayLiteral1, c_strMondayLiteral2, c_strMondayLiteral3:             Result = WEEKDAYCOUNT(Date1, Date2, vbMonday)
    Case c_strTuesdayLiteral, c_strTuesdayLiteral1, c_strTuesdayLiteral2, c_strTuesdayLiteral3:         Result = WEEKDAYCOUNT(Date1, Date2, vbTuesday)
    Case c_strWednesdayLiteral, c_strWednesdayLiteral1, c_strWednesdayLiteral2, c_strWednesdayLiteral3: Result = WEEKDAYCOUNT(Date1, Date2, vbWednesday)
    Case c_strThursdayLiteral, c_strThursdayLiteral1, c_strThursdayLiteral2, c_strThursdayLiteral3:     Result = WEEKDAYCOUNT(Date1, Date2, vbThursday)
    Case c_strFridayLiteral, c_strFridayLiteral1, c_strFridayLiteral2, c_strFridayLiteral3:             Result = WEEKDAYCOUNT(Date1, Date2, vbFriday)
    Case c_strSaturdayLiteral, c_strSaturdayLiteral1, c_strSaturdayLiteral2, c_strSaturdayLiteral3:     Result = WEEKDAYCOUNT(Date1, Date2, vbSaturday)
' ����������� �����
    Case Else:                                                                                          Result = DateDiff(Interval, Date1, Date2, FirstDayOfWeek, FirstWeekOfYear)
    End Select
HandleExit:  DateDiffEx = Result: Exit Function
HandleError: Resume HandleExit
End Function

'=========================
' ������� ��� ������ � �������������� ������
'=========================
Public Function DateTypes(Optional ID, Optional Col, Optional Row, Optional iStep As Long)
' ���������� ������/�������� �� ����������� ����� ���
'-------------------------
' Id        - ��� �������� �������� arrrData(0)
' Col/Row   - �������/������ �������� �������� (������� � 0)
' iStep     - (������������) ���������� ��������� � ������
'-------------------------
On Error Resume Next
Static aData(): iStep = 6 '[i+0]=ID(eDateType); [i+1]=CNAME;[i+2]=NAME; [i+3]=DESC; [i+54]=DateColor; [i+5]=FaceId
Dim i As Long: i = LBound(aData)
    If Err Then
        Err.Clear
        aData = Array(eDateTypeUndef, "undef", "<�� ������>", "<�� ������>", "Black", "", _
                eDateTypeWeekday, "weekday", "������� (������)", "������� (������) ����", "Navy", "DaysWork", _
                eDateTypeSatday, "dayoff0", "����������", "�������� ��� (�������)", "HotPink", "DaysDayOff", _
                eDateTypeSunday, "dayoff", "��������", "�������� ���", "DeepPink", "DaysDayOff", _
                eDateTypeWorkday, "work", "�������", "������� ����", "Navy", "DaysWork", _
                eDateTypeHolidayPre, "holiday_pre", "���������������", "��������������� (�����������) ���", "MediumPurple", "DaysHolidayPre", _
                eDateTypeHoliday, "holiday", "�����������", "����������� ���", "Red", "DaysHoliday", _
                eDateTypeNonWorkday, "nonworking", "���������", "��������� ���", "PaleVioletRed", "DaysNonWorking", _
                eDateTypeUser, "userday", "�������������", "���� ����������� �������������", "Teal", "DaysUser", _
                eDateServOfficial, "official", "[official]", "����������� ����, ����������� �������", "", "", _
                eDateServWorkCalendar, "workdays", "[workdays]", "���������������� ���������", "")
    End If
' �� ��������� ���������� ���� ������
    If IsMissing(ID) And IsMissing(Row) And IsMissing(Col) Then DateTypes = aData: Exit Function ':Result = aData: GoTo HandleExit
On Error GoTo HandleError
Dim Result
' ������ ������ �������� - ���������� ������ ������� ���� ������
    If Not IsMissing(ID) Then
        Row = 0
        For i = LBound(aData) To UBound(aData) Step iStep
            If aData(i) = ID Then Row = i \ iStep:  Exit For
        Next i
    End If
    If IsMissing(Row) Then
' ������� ������ �������                        - ���������� ������ ��������� �������
        i = (UBound(aData) - LBound(aData) + 1)
        i = i \ iStep + Abs((i Mod iStep) > 0)
        ReDim Result(0 To i - 1)
        For i = LBound(Result) To UBound(Result)
            Result(i) = aData(i * iStep + Col)
        Next i
    ElseIf IsMissing(Col) Then
' ������� ������ ������                         - ���������� ������ ��������� ������
        ReDim Result(0 To iStep - 1)
        For i = LBound(Result) To UBound(Result)
            Result(i) = aData(Row * iStep + i)
        Next i
    Else
' ������� ������ � �������                      - ���������� ������� �� ��������� ������� ��������� ������
        Result = aData(Row * iStep + Col)
    End If
HandleExit:  DateTypes = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Function DropTempCalendar() As Boolean
' ������� ��������� �������
'-------------------------
Dim Result As Boolean ': Result = False
    On Error GoTo HandleError
    If Len(m_strTempName) = 0 Then m_strTempName = c_strDatesTable: Mid(m_strTempName, 1, 3) = c_strTmpTablePref
    DoCmd.DeleteObject acTable, m_strTempName 'DropTempCalendar = p_TableDrop(m_strTempName):
    m_strTempName = vbNullString: m_datTempBeg = 0: m_datTempEnd = 0: Result = True
HandleExit:  DropTempCalendar = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function DateInfoGet(Date1 As Date, _
            Optional DateDesc, Optional EventId As Long, Optional EventIds, _
            Optional AskTable As Boolean = False, _
            Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As eDateType
' ���������� ��� ������� �������� ���� � �������� �������
'-------------------------
' Date1     - ���� ��� ��������
' DateDesc  - �������������� ���������� � ���
' EventId   - (������������) ���� ��������� ������� ���������������� ����
' EventIds  - (������������) ������ ������ ���� ������� ��������������� ����
' AskTable  - ���� True, ��� ���������� ����� ���������� ������� ������� ���
'-------------------------
' ��������� �� ��������� � ������� c_strDatesTable � ������� ���������, ��������:
' ���� ���� ���������� �� �����������, � � ������� ���� ������� ��� ������� - ������������ ������� ����
' �� ���� ���� ����� ����������� ��������� ������ �������, ����� �������� ��������, �������� ��������� ����� �������
'-------------------------
' ������ �� ������ ��������� ��������� ������� � ������ ������� �� ��� � ������ �� �� - ����� ����, �� ����� ���� ������
'-------------------------
Const cstrDelim = ","
Const cstrDescDelim = "; "
Dim Result As eDateType: Result = eDateTypeUndef
    On Error GoTo HandleError
    Result = eDateTypeUndef: EventId = 0 ': EventIds = vbNullString
'Dim datDateBeg As Date:  datDateBeg = Date1
' ���� ������� ��� ��������� � ��������� ���������� �� ��� ������
    If Not p_DatesTableExists(AskTable:=AskTable, dbs:=dbs, wks:=wks) Then GoTo HandleResult
'' ���������� ��� �� ������� ������������� ���
'    ' �������� ���� ��� (�� ������/�� ��������� ����������)
'Dim strTypes As String:     strTypes = Join(Array(eDateTypeWeekday, eDateTypeHolidayPre, eDateTypeHoliday, eDateTypeNonWorkday, eDateTypeUser), ",")
' ��������� ����� ������� ��� �������� ���� �� ��������� ������� - ����� ���������� ������ - �������
Dim rst As DAO.Recordset
    Select Case Date1
    Case m_datTempBeg To m_datTempEnd       ' ������ �� ������� ������������ ��������� ������� ���������� ����� �������
        ' ��������� ������� ������� - ���� ��� - ������� � ���������
    Case Else                               ' ��� ����� ������� ��� ����������q ���� - ������ ����� �� ������ ���
    ' ������ ����� ��������� �������, ������ - ������ ���
        Call p_TempTableOpen(DateSerial(Year(Date1), 1, 1), DateSerial(Year(Date1), 12, 31), TempName:=m_strTempName, dbs:=dbs, wks:=wks)
    End Select
' ��������� ������ ��� ���������� ����
Dim strSQL As String: strSQL = sqlSelectAll & "[" & m_strTempName & "]" & _
                sqlWhere & c_strDateBeg & sqlEqual & p_DateToSQL(Int(Date1)) & sqlOrder & c_strDateType & ";"
        Set rst = CurrentDb.OpenRecordset(strSQL)
    With rst
        EventIds = vbNullString
        If Not (.BOF And .EOF) Then .MoveFirst
Dim tmpType As eDateType, strDesc As String, tmpDesc As String
        Do Until .EOF
' ���������� ��� ������ ����������� � ���� �:
    ' �������� �������� ��� ����
    ' �������� �������� ���� ����������� � ���� �������
    ' � Result ������� �������� ��� ����
    ' � tmpType ������� ��� ���� �� �����������
            'tmpType = .Fields(c_strDateType)
            EventIds = EventIds & cstrDelim & .Fields(c_strKey)
            tmpDesc = Trim(Nz(.Fields(c_strDateDesc).Value, vbNullString)): If Len(tmpDesc) > 0 Then strDesc = strDesc & cstrDescDelim & tmpDesc
            If Result = eDateTypeUndef Then
                Result = .Fields(c_strDateType): If EventId = 0 Then EventId = .Fields(c_strKey)
            Else
                Select Case .Fields(c_strDateType)
                Case eDateTypeUndef:    ' ��� ������� �� ����� - ���������
                Case Is < Result:       ' ���� ������ ��� ����� ����� ������� ��������� - ��������� ���
                ' ��������� ��������:
                    'eDateTypeWorkday = 10           ' 10  ������� ����
                    'eDateTypeHolidayPre = 20        ' 20  ������� (�����������) ����
                    'eDateTypeHoliday = 70           ' 70  �������� (�����������) ����
                    'eDateTypeNonWorkday = 80        ' 80  ��������� ����
                    'eDateTypeUser = 99              ' 99  ���� ����������� �������������
                    Result = .Fields(c_strDateType): EventId = .Fields(c_strKey)
                End Select
            End If
HandleNext: .MoveNext
        Loop
    End With
HandleResult:
' ���� ��� �� ��� �� ����� - ���������� ��� �� ��� ������
    If Result = eDateTypeUndef Then
        If ISWEEKEND(Date1) Then
            Result = eDateTypeSunday
        Else
            Result = eDateTypeWeekday
        End If
        'Select Case DatePart("w", Date1)
        'Case 7:     Result = eDateTypeSatday
        'Case 1:     Result = eDateTypeSunday
        'Case Else:  Result = eDateTypeWeekday
        'End Select
    End If
    If Not IsMissing(DateDesc) Then
' ���������� �������� ���
        DateDesc = VBA.Format$(Date1, "dddd") '
        Select Case Result
        Case eDateTypeHolidayPre, eDateTypeHoliday, eDateTypeNonWorkday:
                                                 DateDesc = DateDesc & " " & "(" & LCase(DateTypes(Result, 2)) & ")"
        Case eDateTypeSatday, eDateTypeSunday:   DateDesc = DateDesc & " " & "(" & LCase(DateTypes(eDateTypeSunday, 2)) & ")"
        Case Else:                               DateDesc = DateDesc & " " & "(" & LCase(DateTypes(eDateTypeWeekday, 2)) & ")"
        End Select
        If Len(strDesc) Then
            If Left(strDesc, Len(cstrDescDelim)) = cstrDescDelim Then strDesc = Mid$(strDesc, Len(cstrDescDelim) + 1)
            DateDesc = DateDesc & ": " & strDesc
        End If
    End If
    If Not IsMissing(EventIds) Then If Len(EventIds) Then If Left(EventIds, Len(cstrDelim)) = cstrDelim Then EventIds = Mid$(EventIds, Len(cstrDelim) + 1)
HandleExit:  DateInfoGet = Result: Exit Function
HandleError: Result = eDateTypeUndef: Err.Clear: Resume HandleExit
End Function

Public Function DateInfoSet(Date1 As Date, DateType As eDateType, _
            Optional ExtInfo As Boolean = False, _
            Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As eDateType
' ������������� ���������� � �������� ���. ���������� ��� ���������� �������
'-------------------------
' Date1     - ���� ��� ������� ������������� ������ ����
' DateType  - ��������������� ��� ����
' ExtInfo   - ������� ������������� ��������� ����������� ���������� �������
'-------------------------
' ToDo: ��� � �����-�� ������� �����-��, �� � ��� ������ � �� ����� ���� ��� ���� �� ��,- ������ ���
'-------------------------
Dim Result As eDateType: Result = eDateTypeUndef
    On Error GoTo HandleError
    If Not p_DatesTableExists Then GoTo HandleExit
' ��������� ��� ���������������� �������
    Select Case DateType
    Case Is < eDateTypeWorkday: DateType = eDateTypeUndef   ' ���� ������ ������������  - ������ ������������� ���
    Case Is > eDateTypeUser:    DateType = eDateTypeUser    ' ���� ������ ������������� - ������ ���������������� ���
    End Select
Dim EventIds As String:         Result = DateInfoGet(Date1, EventIds:=EventIds, dbs:=dbs, wks:=wks)
' ��������� ������ �� ������� ������� ����������� � ����
Dim rst As DAO.Recordset, strWhere As String:
Dim bolNew As Boolean: bolNew = Len(EventIds) = 0:  If bolNew Then strWhere = False Else strWhere = c_strKey & sqlIn & "(" & EventIds & ")" ' <<< ����� ������ �� ����� ������
Dim bolRst As Boolean: bolRst = rst Is Nothing
    Set rst = p_DatesTableOpen(strWhere, dbs:=dbs, wks:=wks)
    Select Case Result                              ' Result -> DateType
    Case Is < eDateTypeWorkday                      ' ������� ������ ���� ���������� -> ������ ����� ������� ���������� ����
    Case DateType: Err.Raise vbObjectError + 512    ' ������� ������ ���� ��������� � ��������������� (������ �� ��������) -> ������
    'Case eDateTypeUser                              ' ������� ������ ���� ���������������� ������� -> ������ ����� ������� ���������� ����
    'Case Is > DateType:                             ' ������� ������ ���� ����� ������� ���������, ��� ��������������� -> ������ ����� ������� ���������� ����
    Case Is < DateType:                             ' ������� ������ ���� ����� ������� ���������, ��� ��������������� ->
    ' ����� ���������� ������ � ������� ����������� ���� ������� ������� ��� ������� ����������� � ���� ������� ����� ������� ���������, ��� ������ ���� �������� ������������
    ' ������ ����� �������� ������� � ����� ������� ����������� ����� ��������� ����� ������� � �������� �������� ����� ������ ���� �� ���������
Dim strTitle As String, strMessage As String
Dim msgRet As VbMsgBoxResult, bolAll As Boolean
Const cSpaces = 7
        strTitle = "�������� �������"
        strMessage = "����� ���������� ���: " & Format(Date1, "dd.mm.yyyy") & " ��� ����: " & DateTypes(DateType, 2) & vbCrLf & _
                "���������� ������� ��� �������, ����������� � ����, ������� ����� ������� ���������," & vbCrLf & _
                "����� ������� ����� �������, �� ������ ���� �� ���������." & vbCrLf & _
                "" & vbCrLf & _
                "��" & vbTab & " - ������� ��� c ����� ������� �����������;" & vbCrLf & _
                "���" & vbTab & " - �������� ������������ �� �������� �������;" & vbCrLf & _
                "������" & vbTab & " - �� �������"
        msgRet = MsgBox(strMessage, vbYesNoCancel Or vbExclamation, strTitle)
        Select Case msgRet
        Case vbYes:                     ' ������� ��� c ����� ������� ����������� ��� ��� ��������
        Case vbNo:                      ' �������� ������������ �� �������� �������
        Case Else: msgRet = vbCancel    ' �� �������, ������� ����� �������, �� ������ ���� �� ���������
        End Select
    bolAll = msgRet = vbYes
    With rst
    'If Not .EOF Then .MoveFirst
    Do Until .EOF Or msgRet = vbCancel
        If .Fields(c_strDateType) < DateType Then
    ' ���������� ���� ���� ��������� ���� ���������������� - ���������� �������
    ' �������� ����� ��� �������� �������� ������� ��� ����������
        If Not bolAll Then
        ' ���� ���������� ��� ������
        strMessage = "������� �������: " & Format(Date1, "dd.mm.yyyy") & " ��� ����: " & DateTypes(.Fields(c_strDateType), 2) & vbCrLf & _
                "��������: " & Nz(.Fields(c_strDateDesc), vbNullString) & vbCrLf & _
                "" & vbCrLf & _
                "��" & vbTab & " - ������� �������;" & vbCrLf & _
                "���" & vbTab & " - �� ������� �������;" & vbCrLf & _
                "������" & vbTab & " - ��������"
        msgRet = MsgBox(strMessage, vbYesNoCancel Or vbExclamation, strTitle)
        End If
        ' �������� ��������
        Select Case msgRet
        Case vbYes:  .Delete            ' ������� �������
        Case vbNo                       ' �� ������� �������
        Case Else:  msgRet = vbCancel   ' �������,- ������� ����� �������, �� ������ ���� ����� ���� �� ������������ �� ����������
        End Select
        End If
        .MoveNext
    Loop
    End With
    End Select
' ������ ����� ������� ���������� ����
HandleNew:   Call p_DatesTableEventEdit(, Date1, DateType, ExtInfo, rst, dbs, wks): Result = DateType
HandleExit:  If bolRst Then rst.Close: Set rst = Nothing
             DateInfoSet = Result: Exit Function
HandleError: Result = eDateTypeUndef: Err.Clear: Resume HandleExit
End Function

Public Function DateInfoAsk(Date1 As Date, _
        Optional ParentControl As Access.Control, _
        Optional ByVal x As Long, Optional ByVal y As Long, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As eDateType
' ������� ����������� ���� � �������� ���� ����������� ���
'-------------------------
' Date1 - ���� ����������� �������
' ParentControl - ������ �� ������� ��� ������� �.�. �������� ����������� ����
' X,Y - ������� ������ ���� (� �������� �����������)
'-------------------------
Dim Result As eDateType: Result = eDateTypeUndef
    On Error GoTo HandleError
' ������� �� ������ ������� ���� ��������
Dim tmpPoint As POINT
Dim w As Long, h As Long ', varChild
    If IsMissing(ParentControl) Then
        GetCursorPos tmpPoint: x = tmpPoint.x: y = tmpPoint.y
    ElseIf ParentControl Is Nothing Then
        GetCursorPos tmpPoint: x = tmpPoint.x: y = tmpPoint.y
    Else
        Call AccControlLocation(ParentControl, x, y, , h): y = y + h
    End If
' ��������� � ������� ����������� ���� (��� ����� �������� � ��������)
    ' ��� ������������ ������������ ����  ����������:
    '  ������������� ��� ����, ��������� ����, � ����� ����������� ������� � �������� ��� - ��� ������������ ����������� ����� ��������� �� ������� ��� �������������
Dim strSkip As String: strSkip = Join(Array(eDateTypeUndef, eDateTypeWeekday, eDateTypeSatday, eDateTypeSunday, eDateServWorkCalendar, eDateServOfficial), ";")
Dim EventId As Long, EventIds 'As String
    Result = DateInfoGet(Date1, , EventId, EventIds, dbs:=dbs, wks:=wks)
    Select Case Result
    Case eDateTypeWeekday:   strSkip = eDateTypeWorkday & ";" & strSkip ' ��� ������ ����������� ������� ���� ������ ������� ��� ������ �������
    Case eDateTypeWorkday, eDateTypeHolidayPre, eDateTypeHoliday, _
        eDateTypeNonWorkday: strSkip = Result & ";" & strSkip           ' ��������� � ���������� ������� ��� ���� (����� ���������������� - �� ����� ��������� ����� ����������)
    End Select
' ��������� ������ ��� ������������ ����
    ' <��� �������� 1>,<������������ �������� 1>,<��� ������ 1>
Dim strList As String: strList = "3;1;6"    ' ������ ������� ������� DataTypes, ����������� ��� ������������ ������������ ����
    strList = p_DateTypesList(strList, strSkip)     ' �������� ����������� ��� ������������ ���� ������ ��� ��������� �����
    strList = strList & ";"                         ' � ����� ��������� ����������� (����� ������)
    strList = strList & ";�������������...;-1;#534" ' � ���.�������� - ������������� ������� ����
' ��������� ����������� ���� (����� ��� ��� ������ ����?)
    Result = eDateTypeUndef
    FormOpenContext strList, ContextVal:=Result, x:=x, y:=y:
' �� ��������� ������ ������������ ����������� � ��������� ���������� � ����
    Select Case Result
' ������ ������ ������� ���� (��������� ���� ����).
    Case eDateTypeHoliday, eDateTypeNonWorkday  ' ���������
    Case eDateTypeWorkday, eDateTypeHolidayPre  ' �������
    Case eDateTypeUser                          ' ����������������
    Case -1:
' ������������� ������� ����
        If Len(EventIds) = 0 Then Result = eDateTypeUndef: GoTo HandleNew  ' ���� ��� ���� ������� �� ������ - ������ ����� ������� �������������� ���� (��� ������� �.�. �������� � �����)
        If EventId <> EventIds Then EventId = DateInfoList(Date1, dbs, wks) ' ���� ��� ���� ������ ��������� ������� - ��������� ������ ������� ���� � ���������� ������������ ����� ������, ������� �������� �������������
        Result = DateInfoEdit(EventId, , dbs, wks): GoTo HandleExit          ' ������ ��������� ������� � �������
    Case Else: Err.Raise vbObjectError + 512 'HandleError
    End Select
' ������ ����� ������� ���������� ����
HandleNew:   Result = DateInfoSet(Date1, Result)
HandleExit:  DateInfoAsk = Result:  Exit Function
HandleError: Result = eDateTypeUndef: Err.Clear: Resume HandleExit
End Function

Public Function DateInfoList(Date1 As Date, _
            Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Long
' ��������� ������ ������� ��� ����. ���������� ��� ���������� �������
'-------------------------
' Date1   - ���� ��� ������� ������� ������ �������
'-------------------------
Dim Result As eDateType
    On Error GoTo HandleError
    If Not p_DatesTableExists Then GoTo HandleExit
    Result = p_DatesTableEventsList(Date1, , dbs, wks) ' ��������� ������ ��� ��������� ����
HandleExit:  DateInfoList = Result: Exit Function
HandleError: Result = eDateTypeUndef: Err.Clear: Resume HandleExit
End Function

Public Function DateInfoEdit(DateId As Long, _
            Optional ExtInfo As Boolean = False, _
            Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As eDateType
' ����������� ���������� � ��������� �������. ���������� ��� ���������� �������
'-------------------------
' DateId   - ��� ������� ������� ���������� ���������������
' ExtInfo   - ������� ������������� ��������� ����������� ���������� �������
'-------------------------
Dim Result As eDateType: Result = eDateTypeUndef
    On Error GoTo HandleError
    If Not p_DatesTableExists Then GoTo HandleExit
' ����������� ��������� �������
HandleNew:   Call p_DatesTableEventEdit(DateId, , Result, ExtInfo, dbs:=dbs, wks:=wks)
HandleExit:  DateInfoEdit = Result: Exit Function
HandleError: Result = eDateTypeUndef: Err.Clear: Resume HandleExit
End Function

Public Function DateInfoUpdate(Optional DateBeg, Optional DateEnd, _
        Optional AskUpdate As Long = True, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Boolean
' ��������� ������ � ����������� � �������� ���� � ������� c_strDatesTable ������� ����������������� ��������� �� ��������
'-------------------------
' DateBeg   - ������ ������� �� ������� ����������� ������ ����������������� ���������
' DateEnd   - ����� �������
' AskTable  - ���� True ��� ���������� ����� ���������� ������� ������� ���
    ' 0 - �� ���������� � �� ��������� ������,
    '-1 - ��������� ��������� ������ ��� �������
    ' 1 - ���������� � � ����������� �� ������ ������������ ��������� ������ ��� ���
' dbs,wks   - ������ �� ���� � ������� ���������� ������������  ���������
'-------------------------
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
Dim datDateBeg As Date, datDateEnd As Date
Dim bolUpdate As Boolean
' ��������� ���� �������
If IsMissing(DateBeg) Then datDateBeg = Now Else If Not IsDate(DateBeg) Then datDateBeg = Now Else datDateBeg = DateBeg
If IsMissing(DateEnd) Then datDateEnd = Now Else If Not IsDate(DateEnd) Then datDateEnd = Now Else datDateEnd = DateEnd
' ������������ ���� ������� ����� �������� ������ ����� �� ������ ���
    datDateBeg = DateSerial(DatePart("yyyy", datDateBeg), 1, 1): datDateEnd = DateSerial(DatePart("yyyy", datDateEnd), 12, 31)

' ��������� ������� ������� ������������� ��� � ��������� �
Dim rst As DAO.Recordset
    Set rst = p_DatesTableOpen(dbs:=dbs, wks:=wks): Result = Not (rst Is Nothing)
    If Not Result Then Result = p_DatesTableCreate(dbs:=dbs, wks:=wks): If Result Then Set rst = p_DatesTableOpen(dbs:=dbs, wks:=wks): Result = Not (rst Is Nothing)
    If Not Result Then Err.Raise vbObjectError + 512
'Dim lYears As Long: lYears = DateDiff("yyyy", datDateBeg, datDateEnd) + 1 ' �������� ���������� ��� � �������

    'Application.EnableEvents = True
Dim i As Single, iMax As Single
Dim j As Long, jMax As Long
Dim strTitle As String, strText As String, strMessage As String
    strTitle = "��������� ������ �� ��������"
    strText = "��� �������� ������"
    strMessage = "��������� ������ ����������������� ��������� " & vbCrLf & _
        "�� ������ � " & DateBeg & " �� " & DateEnd & " " & vbCrLf & _
        "c �����: " & �_strLink & ""
    i = CSng(datDateBeg): iMax = CSng(datDateEnd)
    j = 1: jMax = 12
''------------------------------------------
Dim prg As clsProgress: Set prg = New clsProgress
Dim aData(), r As Long, c As Long
    With prg
        .Init pMin:=i, pMax:=iMax, pCaption:=strTitle
        .Detail = strMessage
        .Show
Dim datDateCur As Date ' ������� ���� ����� �������
Dim lngId As Long, datDateEdit As Date
        Do Until .Progress = .ProgressMax 'And Not .Canceled
            DoEvents
            .Text = strText & String(j, ".")
    ' �������� ������ �� �������� ������ � ���� ������ �������� - ������� � �������
        ' ������ �������� � ��������� ������ ����� ����������� � �������
        ' ������� p_HolidaysFromWeb �� ���������� �� �������, - ����� ���� ���� ����������� ������������ ���� ��� Access. ����� ����� �����������)
            datDateCur = datDateEnd
        ' ��������� ������� ������ �� ��������� ������ � �������
            ' ���� ���� - ���������� ������������ - �������� ��� ����������
            If p_IsUpdateExists(datDateBeg, datDateCur, lngId, datDateEdit, dbs:=dbs, wks:=wks) Then
            ' ���� ���� ������ �� ������
                If AskUpdate = 1 Then
            ' ������������� ���������� ���������� ������������
                    strTitle = "���������� ������ �������"
                    strMessage = "�������� ������ ����������������� ��������� " & vbCrLf & _
                              "�� ������ � " & datDateBeg & " �� " & datDateCur & ", " & vbCrLf & _
                              "������� c " & �_strLink & " ?"
                              '"(���� ���������� ���������� " & datDateEdit & ") "
                    bolUpdate = (MsgBox(strMessage, vbYesNo Or vbExclamation, strTitle) = vbYes)
                Else
            ' ������������� ���������� ������������ ���������� �����
                    bolUpdate = AskUpdate
                End If
        ' ��������� ���� ����������
                'bolUpdate = bolUpdate and (DateDiff("d", datDateEdit, Now()) > cUpdPeriod)
        ' ���� ������ �� ������, �� ������ ���������� ��������� ������ - ��������� � ���������� �������
                If Not bolUpdate Then GoTo HandleNext
            End If
        ' �������� ������������ ������ �� �������� ������� �� �������� � ������
            aData = p_HolidaysFromWeb(datDateBeg, datDateCur, prg:=prg)
        ' ���� ������ �� ������ �� �������� ��� �������� ������������ - ��������� � ���������� �������
            If Not IsArray(aData) Then GoTo HandleNext
            If UBound(aData, 2) - LBound(aData, 2) < 0 Then GoTo HandleNext
            If bolUpdate Then
        ' ���� ������ ��������� ��������� ������ - ������� ������
Dim strWhere As String: strWhere = c_strKey & sqlEqual & lngId & sqlOR & c_strParent & sqlEqual & lngId
                Call dbs.Execute(sqlDeleteAll & c_strDatesTable & sqlWhere & strWhere)
            End If
            With rst
        ' ��������� � ������� ��������� ������ � ������� �� ������� ���� ��������� ����������
                    .AddNew
                    .Fields(c_strDateType) = eDateServWorkCalendar:
                    .Fields(c_strDateBeg) = datDateBeg: .Fields(c_strDateEnd) = datDateCur:
                    .Fields(c_strDateDesc) = "���������������� ��������� (� " & datDateBeg & " �� " & datDateCur & ")"
                    .Fields(c_strComment) = "��������� �: " & �_strLink
                    .Update
                lngId = dbs.OpenRecordset(sqlSelect & sqlIdentity)(0).Value ' ���� ��������� ������ � ������� �� ������� �������� ������
        ' ��������� ���������� ������ �� ������� � ������� � ��������� ������ �� ���� ������ � ������� �� ������� ��� ��������
                For r = LBound(aData, 2) To UBound(aData, 2)
                    .AddNew:
                    .Fields(c_strParent) = lngId            ' ������ �� ������������ ������ ()
                    .Fields(c_strDateType) = aData(1, r)    ' ��� �������
                    .Fields(c_strDateBeg) = aData(0, r)     ' ���� �������
                    .Fields(c_strActBegDate) = datDateBeg   ' ���� ������ ������������ ������ �������
                    '.Fields(c_strActEndDate) = datDateCur   ' ���� ����� ������������ ������ �������
                    .Update
                Next r
            End With
        ' ��������� � ���������� ��������� �������
HandleNext: datDateBeg = datDateCur + 1: If datDateBeg >= datDateEnd Then Exit Do
            If .Canceled Then
        ' ���������� ��������
                MsgBox " ������� �������� ������ �������." & vbCrLf & _
                "����: " & Format$(datDateCur, "dd.mm.yyyy")
                Exit Do
            End If
            j = j + 1: If j > jMax Then j = 1
        Loop
    End With
    Set prg = Nothing
    rst.Close
HandleExit:  Set rst = Nothing: DateInfoUpdate = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
'=========================
' ���������� ������� ������
'=========================

Private Function p_DatesTableEventEdit(Optional EventId, Optional Date1, _
    Optional DateType As eDateType, Optional ExtInfo As Boolean = False, _
    Optional rst As DAO.Recordset, Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Long
' ����������� / ������ ������� � �������
'-------------------------
' EventId   - ��� �������  (���� ����� - ������������� ��������,����� �������� ����� �������)
' Date1     - ���� ������� (���� �� ������ - �.�. ������ EventId)
' DateType  - ��� �������
' ExtInfo   - ������� ������������� ��������� ����������� ���������� �������
' rst/dbs/wks - ������ �� �������� ������
'-------------------------
Dim Result As Long ': Result = False
    On Error GoTo HandleError
Dim strWhere As String:
Dim bolRst As Boolean: bolRst = rst Is Nothing
Dim bolNew As Boolean: bolNew = IsMissing(EventId): If bolNew Then strWhere = False Else strWhere = c_strKey & sqlEqual & EventId
    If bolNew Then If IsMissing(Date1) Then Err.Raise vbObjectError + 512
' ��������� ��� ���������������� �������
'    Select Case DateType
'    Case Is < eDateTypeWorkday: DateType = eDateTypeUndef   ' ���� ������ ������������  - ������ ������������� ���
'    Case Is > eDateTypeUser:    DateType = eDateTypeUser    ' ���� ������ ������������� - ������ ���������������� ���
'    End Select
    ' ��������� ������ �� ������� �������
    If bolRst Then Set rst = p_DatesTableOpen(strWhere, dbs:=dbs, wks:=wks)
Dim strDesc As String:
    With rst
        If bolNew Then
            .AddNew
        Else
            .MoveFirst: .FindFirst (strWhere): If .NoMatch Then Err.Raise vbObjectError + 512
            DateType = .Fields(c_strDateType): Date1 = .Fields(c_strDateBeg)
            .Edit
        End If
        strDesc = Nz(.Fields(c_strDateDesc), Format(Date1, "dd.mm.yyyy"))
        '.Fields(c_strParent) = ???            ' ������ �� ������������ ������ ()
        If Not ExtInfo Then
    ' ������� ������� ���������� ���� + �������� (Desc)
        strDesc = InputBox("������� ������� �������� �������:", "��������", strDesc)
        .Fields(c_strDateType) = DateType       ' ��� �������
        .Fields(c_strDateBeg) = Date1           ' ���� �������
        If Len(strDesc) > 0 Then .Fields(c_strDateDesc) = strDesc    ' �������� �������
    Else
    ' ������ ������������� ������� - ��������� ����� ��������� �������, ���� DateType=Undef � ����� ����� ������������� ��� �������
MsgBox "��� �������� ��� �������! " & vbCrLf & "����� ����� ��������� ����� " & vbCrLf & "�������������� ���������� �������.", _
        vbOKOnly Or vbExclamation, "��������!"
Stop
'Dim tmpForm As Form, FormName As String: FormName = ""
'        FormOpenDrop FormName, NewForm:=tmpForm, FormVal:=Result, X:=X, Y:=Y, Icon:=Icon ', Arrange:= eAlignRightBottom , Visible:= True, FormParent:=ParentControl
'        With tmpForm
'        '.DateType = DateType:.AskType = (DateType = eDateTypeUndef) ' ���� ��� ������� ������������� - � ����� ��������� ������� ���� ��������� ����� � ���
'        '.DateBeg = Date1
'        Do While .Visible: DoEvents: Loop
'        If .ModalResult = vbOK Then
'        'DateType = .DateType: Date1 = .DateBeg
'        'rst.Fields(c_strDateType) = .DateType      ' ��� �������
'        'rst.Fields(c_strDateBeg) = .DateBeg        ' ���� �������
'        'rst.Fields(c_strDateDesc) = .DateDesc      ' �������� �������
'        'If Not IsNull(.DateEnd) Then rst.Fields(c_strDateEnd) = .DateEnd       ' ���� ��������� ��������� �������
'        'If Not IsNull(.OffsetType) Then rst.Fields(c_strOffsetType) = .OffsetType: rst.Fields(c_strOffsetValue) = .OffsetValue   ' ��� ������������� (��� ��������/�������� ��������)
'        'If Not IsNull(.PeriodType) Then rst.Fields(c_strPeriodType) = .PeriodType: rst.Fields(c_strPeriodValue) = .PeriodValue   ' ��� ������������� (��� �������/�������� �������)
'        '.Fields(c_strDateDesc) = strDesc        ' ����������� � �������
'        End If
'        End With
    End If
        .Fields(c_strActBegDate) = Date1        ' ���� ������ ������������ ������ �������
        '.Fields(c_strActEndDate) = Date1       ' ���� ����� ������������ ������ �������
        .Fields(c_strEditDate) = Now()          ' ��������� ������
        .Update
    End With
    ' ���������� ��������� �������
    Call p_TempTableOpen(m_datTempBeg, m_datTempEnd, TempName:=m_strTempName, Requery:=True, dbs:=dbs, wks:=wks)
HandleExit:  If bolRst Then rst.Close: Set rst = Nothing
             p_DatesTableEventEdit = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_DatesTableEventsList(Date1 As Date, _
    Optional AllowEdit As Boolean = False, _
    Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Long
' �������� ������� ����������� � ����
'-------------------------
' Date1     - ���� ������ ������� ������� ���������� �������
' AllowEdit - ������� ����������� ������������� ������ ������� ���� (���������/�������/�������� ��� ���������� ������������)
' dbs/wks - ������ �� �������� ������
'-------------------------
Dim Result As Long ': Result = False
    On Error GoTo HandleError
    MsgBox "��� �������� ��� �������! " & vbCrLf & "����� ����� ��������� ����� " & vbCrLf & "������ ������� ��� ����.", _
            vbOKOnly Or vbExclamation, "��������!"
Stop

'Dim EventIds As String:     Result = DateInfoGet(Date1, EventIds:=EventIds, dbs:=dbs, wks:=wks)
'' ���������� ������ �� ������� ������� ����������� � ����
'Dim strWhere As String:     strWhere = c_strKey & sqlIn & "(" & EventIds & ")"
'Dim Rst As DAO.Recordset:   Set Rst = p_DatesTableOpen(strWhere, dbs:=dbs, wks:=wks)
'Dim tmpForm As Form, FormName As String: FormName = ""
'        FormOpenDrop FormName, NewForm:=tmpForm, FormVal:=Result, X:=X, Y:=Y, Icon:=Icon ', Arrange:= eAlignRightBottom , Visible:= True, FormParent:=ParentControl
'        With tmpForm
'        'Set .Recordset = rst
'        Do While .Visible: DoEvents: Loop
'        If .ModalResult = vbOK Then
'        '...
'        End If
'        End With
'    ' ���������� ��������� �������
'    Call p_TempTableOpen(m_datTempBeg, m_datTempEnd, TempName:=m_strTempName, dbs:=dbs, wks:=wks)
HandleExit:  'Rst.Close: Set Rst = Nothing
             p_DatesTableEventsList = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_CheckForHolidays(ByVal Date1 As Date, ByVal Date2 As Date, _
        Optional Holidays = False, Optional Weekends, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Long
' ��������� ���� �� ������� ���������� � ���������� ����� �� ������� �.�. ��������������� ���������� ������� ���� � �������� ������� ��� ������������ Holidays
'-------------------------
' Date1     - ���� ������ �������
' Date2     - ���� ��������� �������
' Holidays  - ������������� ������ ����������� ��� (����������� � ��������� ����, � ����� ������� �������� ����), ����:
'             0 - ���� ����� �������� �� �������, 1 - �� ���������, ����� ������ - ��������� �� ����� �����������
'!!! �������� !!! ���� ������������� ���� ������ �������� ��������� ��-�� �������� ��������� ���������
' ��������� ������������� Holidays � ��������� �������� � ������ �� ������� ������� ������� � �������
' Weekends  - ����� ��� ������ �������� ��������� � �� ��������� �������� (��. p_Weekends). ��-���������: �������, �����������
'-------------------------
' ������ �.�.:
'   ���������:  ����/��� (��� ���������� p_HolidaysFromTable � p_HolidaysFromWeb)
'   ����������: ���� (��� ���� �������� �������� ���� �� ��������� ������ ��������)
'-------------------------
Dim Result As Long
    On Error GoTo HandleError
Dim Temp
Dim aHolidays()
' ��������� ���������� ���� ������ � ����� �������
    If Date2 < Date1 Then Temp = Date2: Date2 = Date1: Date1 = Temp
' ��������� ���������� Holidays
    If Not IsArray(Holidays) Then
        Select Case Holidays
        Case 0:     aHolidays = p_HolidaysFromTable(Date1, Date2, dbs:=dbs, wks:=wks)
        Case 1:     aHolidays = p_HolidaysFromWeb(Date1, Date2)
        Case Else:  Err.Raise vbObjectError + 512
        End Select
        If Not IsArray(aHolidays) Then GoTo HandleExit
    Else
        aHolidays = Holidays
    End If
' ��������� ��� ������� �� Holidays
Dim rMax As Long, cMax As Long
Dim d As Integer
    rMax = UBound(aHolidays, 2): If Err = 0 Then cMax = UBound(aHolidays) Else Err.Clear: rMax = UBound(aHolidays)
    If cMax > 0 Then
    ' Holidays ��������� ������  (����/���) - ��������� �� ���� ���
' ! ����������� ��� ���������� ���������� �������� ����������� � ����
Dim r As Long: For r = 0 To rMax
            Temp = aHolidays(0, r)          ' ������ ������� - ����
            Select Case Temp
            Case Date1 To Date2             ' ������ ���� ����������� ������� - ���������
                Select Case aHolidays(1, r) ' ������ ������� - ���
                Case eDateTypeWorkday, eDateTypeHolidayPre: If ISWEEKEND(Temp, Weekends) Then Result = Result + 1     ' ������� �������� ���� ��� ��������������� (�����������) ����
                Case eDateTypeHoliday, eDateTypeNonWorkday: If Not ISWEEKEND(Temp, Weekends) Then Result = Result - 1 ' �������� (�����������) ���  ��������� ����
                End Select
            Case Is > Date2: Exit For       ' ������ ���� ������ ���� ������ ������� - �������
            Case Else                       ' ������ ���� ������ ���� ������ ������� - ����������
            End Select
        Next
    Else
    ' Holidays ���������� ������ (������ ����) - ��������� �� �������������� � �������� ����
        ' ������� ���� ��������� � ������ ������� ��������
        ' �������� ���� ��������� � ������� ������� �������
        For Each Temp In aHolidays
            Select Case Temp
            Case Date1 To Date2: Result = Result + IIf(ISWEEKEND(Temp, Weekends), 1, -1) ' ������ ���� ����������� ������� - ���������
            Case Is > Date2: Exit For       ' ������ ���� ������ ���� ������ ������� - �������
            Case Else                       ' ������ ���� ������ ���� ������ ������� - ����������
            End Select
        Next
    End If
HandleExit:  p_CheckForHolidays = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Private Function p_HolidaysFromTable(Date1 As Date, Optional Date2, _
    Optional dbs As DAO.Database, Optional wks As DAO.Workspace)
' ���������� ������� ��������� � ���� ���������� �������: ����/���
'-------------------------
' !!! �������� ������ ������ ���� ������������ �� ����
'-------------------------
Dim Result
On Error GoTo HandleError

' ������ �� ���� ������.    �������� ������: ������� ��������, ���������������, ����������� � ���������
Dim strTypes As String:     strTypes = Join(Array(eDateTypeWorkday, eDateTypeHolidayPre, eDateTypeHoliday, eDateTypeNonWorkday), ",")
' ��������� ����� ������� ��� ��������� ������� �� ��������� ������� - ����� ���������� ������ - �������
' ������������ ����.        �������� ������: ���� ���� ������� � ��� �������
Dim strFields As String:    strFields = Join(Array(c_strDateBeg, c_strDateType), ",")
Dim strTable As String
Dim rst As DAO.Recordset:   Set rst = p_TempTableOpen(Date1, Date2, strTypes, strFields, strTable, Unique:=True, dbs:=dbs, wks:=wks)
' ������ � ������
    With rst: .MoveLast: .MoveFirst: Result = .GetRows(.RecordCount): End With
'' ������� ��������� �������
'    DropTempCalendar 'rst.Close: Set rst = Nothing: Call p_TableDrop(strTable)
HandleExit:  p_HolidaysFromTable = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Private Function p_HolidaysFromWeb(ByRef DateBeg As Date, Optional DateEnd, _
        Optional Weekends, Optional prg)  ' prg As clsProgress)
' ����������� �� ��������� ������, ������ �� � ���������� ������ c �����������
'-------------------------
' !!! ������ ������ �������� ���������. �� ������ ��� ��������� ��� ������������� �������, ��������� � ������, �� �� ���������� � ����� ���������
'-------------------------
Const cstrDivTagName = "div"
'Const cstrDataClassName = "calendar_full"
Const c_strWorkDayClassName = "calendar_day  "
Const cstrHolyPreClassName = c_strWorkDayClassName & "calendar_day__holiday_pre"
Const c_strHolidayClassName = c_strWorkDayClassName & "calendar_day__holiday"
Const cstrDayOffClassName = c_strWorkDayClassName & "calendar_day__dayoff"
'Const cstrHolidays = "��������� ����������� ��� "
'Const cstrTransfers = "������� �������� ���� �"
'Const cstrTransfers = "�������� � ������� ��� ������� ��������� ��������� �� ����� ���� ����������� <b>�������������� ��������� ����������� ���</b> (��. 6 �� ��).&nbsp;"

Const cstrDateAttrName = "data-day"
Const cMaxErr = 10
Const cstrDelim = " " 'Chr(32)

Dim Result As Boolean ':Result = False
On Error GoTo HandleError
Dim bolProgress As Boolean: bolProgress = Not IsMissing(prg) 'TypeOf prg Is clsProgress
    If IsMissing(DateEnd) Then DateEnd = DateBeg
Dim lErrCount As Long
' ����������� �������� ������ �� ��������
Dim Temp As Date ': Temp = DateBeg
Dim strYear As String: strYear = Format(DateBeg, "yyyy")
Dim strURL As String:  strURL = �_strLink & strYear & "/"
Static HTML As Object: Set HTML = CreateObject("htmlFile")          'Dim HTML As New MSHTML.HTMLDocument
Static HXML As Object: Set HXML = CreateObject("MSXML2.XMLHTTP")    '("Msxml2.ServerXMLHTTP")
' ��������� ����� �������
    With HXML
        ''Call .setOption(2, 13056) ' ignore all SSL Cert issues
        .Open "GET", strURL, False: .sEnd
        '        .Open "POST", strURL, False
        '        .setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
        '        .setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        '        .sEnd 'argumentString
        If .ReadyState = 4 Then
            Select Case .Status
            Case 200:  HTML.body.innerHTML = .responseText
            Case 404:  DateBeg = DateSerial(DatePart("yyyy", DateBeg), 12, 31): Err.Raise vbObjectError + 513
            Case Else: Err.Raise vbObjectError + 512
            End Select
        Else
            Err.Raise vbObjectError + 512
        End If
    End With

' ������ ����������� ��������� � ��������� ���� �� ��������� ������ � ���
Dim aData(), i As Long
Dim enmDayType As eDateType
Dim Itm As Object
'Stop
    For Each Itm In HTML.getElementsByTagName(cstrDivTagName)
    On Error GoTo HandleError
    DoEvents
        If bolProgress Then If prg.Canceled Then Err.Raise vbObjectError + 512
        With Itm
            Select Case .ClassName      ' ��������� ��� ������
            Case c_strWorkDayClassName: enmDayType = eDateTypeWorkday           ' ������� ����
            Case cstrDayOffClassName:   enmDayType = eDateTypeSunday            ' ��������
            Case cstrHolyPreClassName:  enmDayType = eDateTypeHolidayPre        ' ���������������
            Case c_strHolidayClassName: enmDayType = eDateTypeHoliday           ' �����������
            Case Else:                  GoTo HandleNext                         ' ������
            End Select
            Temp = CDate(Replace(.Attributes(cstrDateAttrName).Value, "_", ".")): If Temp < DateBeg Then GoTo HandleNext
        End With
' ��������� ���������: ���� �������� ���� � ���������������� ��������� ������� ��� �������� - �� ���� ������ ���������
        If ISWEEKEND(Temp, Weekends) Then
            If enmDayType = eDateTypeSunday Then GoTo HandleNext            ' �������� ���������� ��� ��������
        Else
            If enmDayType = eDateTypeWorkday Then GoTo HandleNext           ' ������� ���������� ��� �������
        End If
' ������� ������ � �������
        ReDim Preserve aData(0 To 1, 0 To i)
        aData(0, i) = Temp: aData(1, i) = enmDayType
        i = i + 1
        If bolProgress Then With prg: .Update: .Progress = CSng(Temp): End With
        If Temp >= DateEnd Then Exit For
HandleNext:
    Next Itm
    DateEnd = Temp: Result = True
HandleExit:  p_HolidaysFromWeb = aData: Exit Function
HandleError: Select Case Err.Number
    Case -2146697208 '
    MsgBox "���������� � IE ��������� ������������ TLS1.2" & vbCrLf _
    & "", vbOKOnly Or vbCritical, "������ 0x800C0008 (INET_E_DOWNLOAD_FAILURE)"
    'Case -2147012721 ' A security error occurred
    'Case -2147220991 ' Automation error ������� �� ������ ������� �� ������ �� ���������
    '                 ' � HTML ����������� ��������� ������ ��� ���������
    Case vbObjectError + 513 ' Error 404
Debug.Print "Can't get data for year " & strYear
    End Select
    Result = False: Err.Clear: Resume HandleExit
End Function

Private Function p_DatesTableExists(Optional bTest As Boolean = False, Optional AskTable As Boolean = True, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Boolean
' ��������� ������� ������� ������ ��� ������ �������
'-------------------------
' bTest     = True ��� ��������� ��������
' AskTable  - ���� True ��� ���������� ����� ���������� ������� ������� ���
' dbs,wks   - ������ �� ���� � ������� ���������� ������������  ���������
'-------------------------
Static bolExists As Boolean ' ������� ������� �������
Static bolInit As Boolean   ' ������� ���� ��� �������� ��� �����������
    If bTest Then GoTo HandleTest
    If bolInit Then p_DatesTableExists = bolExists: Exit Function
HandleTest:
    On Error Resume Next
Dim Result As Boolean
' ��������� ������� �������
    Result = Not p_DatesTableOpen(bTest:=True, dbs:=dbs, wks:=wks) Is Nothing
' ����� ��������� ��� ������������ (������� �������� �����), �� - �� �����
    If Not Result Then
        If AskTable Then
' ���� ����������� - ���������� � ������ �����
Dim strText As String, strTitle As String
            strTitle = "����������� �������"
            strText = "����������� ������� ��� �������� ������������� ���." & vbCrLf & "������� ������� """ & c_strDatesTable & """?"
            If (MsgBox(strText, vbYesNo Or vbExclamation, strTitle) = vbYes) Then Result = p_DatesTableCreate(dbs:=dbs, wks:=wks)
        End If
    End If
    bolExists = Result
    bolInit = True: p_DatesTableExists = bolExists
End Function
Private Function p_IsUpdateExists(Date1 As Date, Date2 As Date, _
    Optional ID As Long, Optional RecDate As Date, _
    Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Long
' ��������� ������� � ������� ����������� ������ �� ��������� ������ ���������� ��� ������ ������ ���������� ��������������� �������
'-------------------------
' Date1, Date2 - ���� ������/����� �������
' ID, RecDate - ���/���� �������� ������
' dbs,wks   - ������ �� ���� � ������� ���������� ������������  ���������
'-------------------------
' ��������� ������ ��������� ������ ���������� ����������������� ���������,
' ����� ��� �������� ������������ ��������� � ������� ������ ��� ���������� �� ��������
'-------------------------
Dim Result As Boolean:
    If Not p_DatesTableExists(AskTable:=True, dbs:=dbs, wks:=wks) Then Err.Raise vbObjectError + 512
'Dim strFields As String:    strFields = Join(Array(c_strKey, c_strEditDate), ",")
Dim strOrder As String:     strOrder = c_strDateBeg
Dim strTypes As String:     strTypes = eDateServWorkCalendar                                                ' �������� ������ ���������� ����������������� ���������
Dim strWhere As String:     strWhere = c_strDateType & sqlIn & "(" & strTypes & ")"                         ' ������ �� ���� ������
    strWhere = strWhere & sqlAnd & p_DateToSQL(Date1) & sqlBetween & c_strDateBeg & sqlAnd & c_strDateEnd   ' ��������� ������ �� ����
Dim rst As DAO.Recordset:   Set rst = p_DatesTableOpen(strWhere, strOrder, dbs:=dbs, wks:=wks)
' ��������� ������� ����������� ������ �� ������
    With rst
        Result = Not (.EOF And .BOF)
        If Result Then Date1 = .Fields(c_strDateBeg): Date2 = .Fields(c_strDateEnd): ID = .Fields(c_strKey): RecDate = .Fields(c_strEditDate)
    End With
HandleExit:  p_IsUpdateExists = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Private Function p_DatesTableOpen(Optional strWhere As String, Optional strOrder As String, Optional strFields As String, _
        Optional bTest As Boolean, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As DAO.Recordset
' C�������� ������� ������������� ���
'-------------------------
    On Error GoTo HandleError
    If wks Is Nothing Then Set wks = DBEngine.Workspaces(0): Set dbs = CurrentDb                                         ' ��-��������� ������� ���� � ������� ������������
    If dbs Is Nothing Then If wks.Databases.Count > 0 Then Set dbs = wks.Databases(0) Else Err.Raise vbObjectError + 512 ' ���� �� ������ ����, �� ������ ������� ������������ - ���� ������ � ������� ������������, ����� - ������
Dim strSQL As String: strSQL = sqlSelect
    If bTest Then strSQL = strSQL & sqlTop1
    If Len(strFields) = 0 Then strSQL = strSQL & sqlAll Else strSQL = strSQL & strFields
    strSQL = strSQL & sqlFrom & c_strDatesTable
    If Len(strWhere) > 0 Then strSQL = strSQL & sqlWhere & strWhere
    If Len(strOrder) > 0 Then strSQL = strSQL & sqlOrder & strOrder
    strSQL = strSQL & ";"
    Set p_DatesTableOpen = dbs.OpenRecordset(strSQL)
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Private Function p_DatesTableCreate(Optional bTemp As Boolean = False, Optional TableName As String, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Boolean
' ������ ������� ������������� ���
'-------------------------
' bTemp     = False ��� �������� �������� ������� ���������, (Id �������� �������� �����)
'           = True ��� �������� ��������� ������� ���������
' TableName - ��� ����������� �������. �� ������ ���������� ����� ��������� ��� ��������� �������
' dbs,wks   - ������ �� ���� � ������� ���������� ������������  ���������
'-------------------------
Dim Result As Boolean
Dim bolTransOpen As Boolean
    On Error GoTo HandleError
    If wks Is Nothing Then Set wks = DBEngine.Workspaces(0): Set dbs = CurrentDb                                         ' ��-��������� ������� ���� � ������� ������������
    If dbs Is Nothing Then If wks.Databases.Count > 0 Then Set dbs = wks.Databases(0) Else Err.Raise vbObjectError + 512 ' ���� �� ������ ����, �� ������ ������� ������������ - ���� ������ � ������� ������������, ����� - ������
Dim tdf  As DAO.TableDef, fld As DAO.Field, idx As DAO.Index
' �������� ��� ����������� �������
    If Len(TableName) = 0 Then TableName = c_strDatesTable: If bTemp Then Mid(TableName, 1, 3) = c_strTmpTablePref
HandleCreate:
    If bTemp Then
    ' ��� ��������� - ��������� ������� ������� TableName � ��� ������� ������� �
        With wks
            If p_IsTableExists(TableName, dbs:=CurrentDb) Then .BeginTrans: bolTransOpen = True: dbs.Execute sqlDropTable & "[" & TableName & "]": .CommitTrans: bolTransOpen = False
        End With
    End If
' ������ �������
    Set tdf = dbs.CreateTableDef(TableName)
    With tdf
' ������ ����
        Set fld = .CreateField(c_strKey, dbLong):           .Fields.Append fld  ' ����
        If Not bTemp Then fld.Attributes = dbAutoIncrField                      ' ������������� ��������� ���� �������� �������
        Set fld = .CreateField(c_strDateType, dbLong):      .Fields.Append fld  ' ��� ����
        Set fld = .CreateField(c_strDateBeg, dbDate):       .Fields.Append fld  ' ���� ������ �������
        Set fld = .CreateField(c_strDateDesc, dbText, 100): .Fields.Append fld  ' �������� ����
        If Not bTemp Then
    ' ���������� ���� ������� ����� ������ ��� �������� �������:
        Set fld = .CreateField(c_strDateEnd, dbDate):       .Fields.Append fld  ' ���� ��������� �������
        Set fld = .CreateField(c_strOffsetType, dbText, 4): .Fields.Append fld  ' ��� ��������
        Set fld = .CreateField(c_strOffsetValue, dbLong):   .Fields.Append fld  ' �������� ��������
        Set fld = .CreateField(c_strPeriodType, dbText, 4): .Fields.Append fld  ' ��� �������
        Set fld = .CreateField(c_strPeriodValue, dbLong):   .Fields.Append fld  ' �������� �������
        Set fld = .CreateField(c_strComment, dbMemo):       .Fields.Append fld  ' ����������� � ����
        Set fld = .CreateField(c_strParent, dbLong):        .Fields.Append fld  ' ��� ��������
        Set fld = .CreateField(c_strActBegDate, dbDate):    .Fields.Append fld  ' ���� ������ ������������ ������
        Set fld = .CreateField(c_strActEndDate, dbDate):    .Fields.Append fld  ' ���� ��������� ������������ ������
        Set fld = .CreateField(c_strEditDate, dbDate):      .Fields.Append fld  ' ���� ��������� ������
        End If
' ������ �������
        Set idx = .CreateIndex("PrimaryKey"): With idx                  ' �������� ������ �� �����
            .Fields.Append .CreateField(c_strKey)
        If Not bTemp Then .Primary = True: .Unique = True               ' ������ ���������� �������� ������ �������� �������
        End With: .Indexes.Append idx
        Set idx = .CreateIndex("DayTypeKey"): With idx                  ' �������������� �� ����/����
            .Fields.Append .CreateField(c_strDateBeg)
            If Not bTemp Then .Fields.Append .CreateField(c_strDateEnd)
            .Fields.Append .CreateField(c_strDateType)
            If Not bTemp Then .Fields.Append .CreateField(c_strOffsetType)
            If Not bTemp Then .Fields.Append .CreateField(c_strPeriodType)
            '.IgnoreNulls = True
        End With: .Indexes.Append idx
    End With
    dbs.TableDefs.Append tdf
    If bTemp Then GoTo HandleTest ' ��� ��������� ������� �� ����� ���������� ���������� � ������ ��������
' ����������� �������������� �������� �����
Dim strList As String
    Set fld = tdf.Fields(c_strEditDate)                                 ' ���� �������� ������ (�������� �� ���������)
        Call PropertySet("DefaultValue", "=Date()", fld)
    Set fld = tdf.Fields(c_strDateBeg)                                  ' ���� ������� (�������� �� ���������)
        Call PropertySet("DefaultValue", "=Date()", fld)
    Set fld = tdf.Fields(c_strActBegDate)                               ' ���� ������ ���������������� �������
        Call PropertySet("DefaultValue", "=Date()", fld)
    Set fld = tdf.Fields(c_strDateType)                                 ' ��� ���� (��������� ������); �������� �� ���������
        strList = p_DateTypesList("1;3", Join(Array(eDateTypeWeekday, eDateTypeSatday, eDateTypeSunday), ";"))
        Call PropertySet("DisplayControl", acComboBox, fld, dbInteger)
        Call PropertySet("RowSourceType", "Value List", fld)
        Call PropertySet("RowSource", strList, fld)
        Call PropertySet("ColumnCount", 2, fld, dbInteger)
        Call PropertySet("ColumnWidths", 0, fld)
        Call PropertySet("DefaultValue", eDateTypeUser, fld)            ' ��� ���� (�������� ��-���������)
        Call PropertySet("TextAlign", 1, fld, dbByte)
    Set fld = tdf.Fields(c_strOffsetType)                               ' ��� �������� (��������� ������)
        strList = p_DateIntervalList
        Call PropertySet("DisplayControl", acComboBox, fld, dbInteger)
        Call PropertySet("RowSourceType", "Value List", fld)
        Call PropertySet("RowSource", strList, fld)
        Call PropertySet("ColumnCount", 2, fld, dbInteger)
        Call PropertySet("ColumnWidths", 0, fld)
    Set fld = tdf.Fields(c_strPeriodType)                               ' ��� ������� (��������� ������)
        strList = p_DateIntervalList
        Call PropertySet("DisplayControl", acComboBox, fld, dbInteger)
        Call PropertySet("RowSourceType", "Value List", fld)
        Call PropertySet("RowSource", strList, fld)
        Call PropertySet("ColumnCount", 2, fld, dbInteger)
        Call PropertySet("ColumnWidths", 0, fld)
' ����������� �������������� �������� �������
        Call PropertySet("SubdatasheetName", "Table." & c_strDatesTable, tdf)                                       ' ���������� � ������������ �� ��������
        Call PropertySet("LinkMasterFields", c_strKey, tdf): Call PropertySet("LinkChildFields", c_strParent, tdf)  ' ����� �� PARENT=ID
        Call PropertySet("Filter", c_strParent & sqlIsNull, tdf): Call PropertySet("FilterOnLoad", True, tdf)       ' �������� ������ ������������ ������
        'Call PropertySet("OrderByOnLoad", False, tdf): Call PropertySet("OrderByOn", False, tdf)
' ��������� ������ ������
    dbs.TableDefs.Refresh: Application.RefreshDatabaseWindow
HandleTest:  Result = p_DatesTableExists(True, dbs:=dbs, wks:=wks)
HandleExit:  p_DatesTableCreate = Result: Exit Function
HandleError: If bolTransOpen Then wks.Rollback: bolTransOpen = False     ' ���������� ����������
    Select Case Err.Number
    Case 3211: 'Stop ' �� ���� ������� ��������� ������� - �������������
    If bTemp Then TableName = c_strTmpTablePref & GenPassword(8): Err.Clear: Resume HandleCreate    ' ���� ������ ��������� - ������� ������� � �� ������
    Case 3734: Stop ' ���� ������ ���� ��������� ������������� � ���������, �������������� �� �������� ��� ����������
    End Select
    Result = False: Err.Clear: Resume HandleExit
End Function

Private Function p_Weekends(Optional Weekends)
' ���������� ������ ���� ������ ������� �������� ��������� � �� ��������� ��������.
'-------------------------
' Weekends  - ����� ��� ������ �������� ��������� � �� ��������� ��������. ��-���������: �������, �����������
'   �������� �������� 1-7, 11-17
'   ��� ������ ����: 0000011, ��� 0- ������� ����, 1-�������� (������� � ������������)
'-------------------------
' ��� ��������� �������� ������� ��������: Weekends()(i), ����� ����� ������������� ��� �������� ������������� �������
'-------------------------
On Error Resume Next
Static aData()
Dim i As Long: i = LBound(aData)
    If Err = 0 Then If IsMissing(Weekends) Then p_Weekends = aData: Exit Function
    Err.Clear: If IsMissing(Weekends) Then Weekends = 1
On Error GoTo HandleError
    If Len(Weekends) = 7 Then
' Weekends - ��������� �������� ���� ������.
    ' �������� ���� ������, ������ �� ������� ���������� ���� ������ (������� � ������������).
    ' �������� 1 ������������ ��������� ���, � 0 � ������� ���. � ������ ��������� ������������ ������ ����� 1 � 0. ������ 1111111 �����������.
    ' ��������, 0000011 ��������, ��� ��������� ����� �������� ������� � �����������
Dim j As Long: For j = 1 To 7
        Select Case Mid(Weekends, j, 1)
        Case 1: ReDim Preserve aData(0 To i)
            Select Case j
            Case 1: aData(i) = vbMonday             ' �����������
            Case 2: aData(i) = vbTuesday            ' �������
            Case 3: aData(i) = vbWednesday          ' �����
            Case 4: aData(i) = vbThursday           ' �������
            Case 5: aData(i) = vbFriday             ' �������
            Case 6: aData(i) = vbSaturday           ' �������
            Case 7: aData(i) = vbSunday             ' �����������
            Case Else: Err.Raise vbObjectError + 512
            End Select
            i = i + 1: If i >= 7 Then Err.Raise vbObjectError + 512
        Case 0:
        Case Else: Err.Raise vbObjectError + 512
        End Select
        Next j: GoTo HandleExit
    End If
HandleSelect: Select Case Weekends
' Weekends - �������� ��������
        Case 1:  aData = Array(vbSaturday, vbSunday)    ' �������, �����������
        Case 2:  aData = Array(vbSunday, vbMonday)      ' �����������, �����������
        Case 3:  aData = Array(vbMonday, vbTuesday)     ' �����������, �������
        Case 4:  aData = Array(vbTuesday, vbWednesday)  ' �������, �����
        Case 5:  aData = Array(vbWednesday, vbThursday) ' �����, �������
        Case 6:  aData = Array(vbThursday, vbFriday)    ' ������� , �������
        Case 7:  aData = Array(vbFriday, vbSaturday)    ' ������� , �������
        Case 11: aData = Array(vbSunday)                ' ������ �����������
        Case 12: aData = Array(vbMonday)                ' ������ �����������
        Case 13: aData = Array(vbTuesday)               ' ������ �������
        Case 14: aData = Array(vbWednesday)             ' ������ �����
        Case 15: aData = Array(vbThursday)              ' ������ �������
        Case 16: aData = Array(vbFriday)                ' ������ �������
        Case 17: aData = Array(vbSaturday)              ' ������ �������
        Case Else:
        End Select
HandleExit:  p_Weekends = aData: Exit Function
HandleError: Err.Clear: Weekends = 1: Resume HandleSelect
End Function

Private Function p_WeekendsList(Optional Weekends, Optional Delim As String = ",") As String
' ���������� �������� ��� ������ ����� ������ (��� SQL: IN (..))
'-------------------------
    p_WeekendsList = Join(p_Weekends(Weekends), Delim)
End Function

Private Function p_DateTypesList(Optional Columns As String, Optional Skip As String, Optional Delim = ";") As String
' ���������� ������ ������ ����������� ����� ���
'-------------------------
On Error Resume Next
Dim i As Long, j As Long, iStep As Long
Dim aSkip() As String:      aSkip = Split(Skip, Delim)
Dim aColumns() As String:   aColumns = Split(Columns, Delim)
Dim aDateTypes():           aDateTypes = DateTypes(iStep:=iStep)
Dim strList As String ': strList = Join(DateTypes, ";")
Dim Value
    For i = LBound(aDateTypes) To UBound(aDateTypes) Step iStep
    ' ��������� �� ������� �����������, � ������ iStep �������
        For j = LBound(aSkip) To UBound(aSkip)
        ' ��������� �� ������ ��������
            If aDateTypes(i) = aSkip(j) Then GoTo HandleNext
        Next j
        For j = LBound(aColumns) To UBound(aColumns)
        ' ��������� ������ ������
            'Value=aDateTypes(i + aColumns(j) - 1)
            strList = strList & Delim & aDateTypes(i + aColumns(j) - 1) 'Value
        Next j
HandleNext:  Err.Clear: Next i
    If Left$(strList, Len(Delim)) = Delim Then strList = Mid(strList, Len(Delim) + 1)
HandleExit: p_DateTypesList = strList
End Function

Private Function p_DateIntervalList(Optional Columns As String, Optional Skip As String, Optional Delim = ";") As String
' ���������� ������ ��������� ����������
'-------------------------
Static strData As String
    If Len(strData) = 0 Then
Dim arrData(): arrData = Array( _
            "", "<none>", _
            "yyyy", "���", _
            "q", "�������", _
            "m", "�����", _
            "w", "������", _
            "d", "����", _
            c_strWorkdayLiteral, "���� (�������)", _
            c_strMondayLiteral, "����������� (���� ������)", _
            c_strTuesdayLiteral, "������� (���� ������)", _
            c_strWednesdayLiteral, "����� (���� ������)", _
            c_strThursdayLiteral, "������� (���� ������)", _
            c_strFridayLiteral, "������� (���� ������)", _
            c_strSaturdayLiteral, "������� (���� ������)", _
            c_strSundayLiteral, "����������� (���� ������)")
'
        strData = Join(arrData, Delim)
    End If
    p_DateIntervalList = strData
End Function

Private Function p_TempTableOpen( _
    Date1 As Date, Optional Date2, _
    Optional DateTypes As String, Optional Fields As String, Optional TempName As String, _
    Optional Unique As Boolean, Optional Requery As Boolean, _
    Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As DAO.Recordset
' ��������� ��������� ����� ������� �� ��������� ������, �������� �� ��������� ��������� �������
'-------------------------
' Date1     - ��������� ���� �������
' Date2     - �������� ���� ������� (�.�.>=Date1)
' DateTypes - ������ ���������� ����� ������
' Fields    - ������ ������������ �������� �����
' TempName  - ��� ��������� ������� ���������� ����������� ������
' Unique    - ���� True ����� ��������� ������ ���������� ������ ���������� ������ �� ���� � ������������ ����������� (�� ��� ����� ����������� ��� ���������)
' Requery   - ���� True ��������� ������� ����� ����������� ���������� �� � �������
' dbs,wks   - ������ �� ���� ������ �� ������� �������������� ������ � ������� ������������
'-------------------------
' ToDo: ������� ��������� � ���� ������ ����� Date1<=m_datTempBeg OR Date2>=m_datTempEnd
' � ��������� ������� ��� � ��� ���-�� ���� - ��������� ���������� ������ ��� �������� ��������� �� �����
' ����� ��������� ������� �� �������� ������� ���������� ������� ������� (
' �� � ����� ��������� ������ ���� �� �������
Dim Result As Boolean ': Result = False
    On Error GoTo HandleError
    If IsMissing(Date2) Then Date2 = Date1
Const cTmp = "[" & c_strTmpTablPref & "]"               ' ��� ����������
Dim sqlDate1 As String: sqlDate1 = p_DateToSQL(Date1)
Dim sqlDate2 As String: sqlDate2 = p_DateToSQL(Date2)
Dim strWhereDate As String: strWhereDate = "(" & c_strDateBeg & sqlBetween & sqlDate1 & sqlAnd & sqlDate2 & ")" ' ������ �� ����
    If Not Requery Then If Date1 >= m_datTempBeg And Date2 <= m_datTempEnd Then GoTo HandleResult
Dim bolTransOpen As Boolean
    If wks Is Nothing Then Set wks = DBEngine.Workspaces(0): Set dbs = CurrentDb                                         ' ��-��������� ������� ���� � ������� ������������
    If dbs Is Nothing Then If wks.Databases.Count > 0 Then Set dbs = wks.Databases(0) Else Err.Raise vbObjectError + 512 ' ���� �� ������ ����, �� ������ ������� ������������ - ���� ������ � ������� ������������, ����� - ������
Dim strSQL As String, strWhere As String, strOrder As String, strFields As String
' ������ ����� ��� ��������� �������
    '!!! ��������� ������� �������� ������ 4 �������� ���� ��. p_DatesTableCreate
Dim arrFields: arrFields = Array(c_strDateBeg, c_strDateType, c_strKey, c_strDateDesc) ', c_strComment)
    strFields = Join(arrFields, ",")
' ������ ��������� ������� (��������)
    Result = p_DatesTableCreate(True, TempName): If Not Result Then Err.Raise vbObjectError + 512
    ' �������� ������
    ' ������ �� ���� �������
    If Len(DateTypes) = 0 Then strWhere = "<100" Else strWhere = sqlIn & "(" & DateTypes & ")"
    strWhere = "(" & c_strDateType & strWhere & ")"
    ' ������ �� ������������ ������
    strWhere = strWhere & _
        sqlAnd & "(IIf([" & c_strActBegDate & "]" & sqlIsNull & ",[" & c_strDateBeg & "],[" & c_strActBegDate & "])<=" & sqlDate2 & ")" & _
        sqlAnd & "(IIf([" & c_strActEndDate & "]" & sqlIsNull & "," & sqlDate2 & ",[" & c_strActEndDate & "])>=" & sqlDate1 & ")"
Dim strWhereTemp As String  ' ������ �� ��������� �������
    strWhereTemp = "((" & c_strDateEnd & sqlIsNull & ")" & _
        sqlAnd & "(" & c_strOffsetType & sqlIsNull & ")" & _
        sqlAnd & "(" & c_strPeriodType & sqlIsNull & "))"
' ��������� ������ ���� ���������� �� ������� "�������" ������� (��������, ������������� ��� �������������)
    ' ��������� ������ ���� � ������ ������������, �������� ������ ����
    ' �������������� ������ �� ��������� �������
    ' ��������� ������ �� �������
    strSQL = sqlSelectAll & c_strDatesTable & sqlWhere & strWhere & sqlAnd & sqlNot & strWhereTemp & ";"
Dim rstSrc As DAO.Recordset, rst As DAO.Recordset
    ' ��������� ������ �������� �� ������� ������� �������
    Set rstSrc = dbs.OpenRecordset(strSQL)
    With rstSrc
        .MoveFirst: If .BOF And .EOF Then GoTo HandleExit
    ' ��������� ������ �� ��������� ��������� ������� ���� ����� ����������� ������
    Set rst = CurrentDb.OpenRecordset(TempName, dbOpenDynaset)
Dim bDate As Date   ' ��������� ���� ��������������� �������
Dim eDate As Date   ' ��������� ���� ��������������� �������
Dim lCount As Long  ' ���������� �������� ���� � �������
Dim lLen As Double  ' ����������� �������
Dim iDate As Date   ' ������� ���� ������ �������
Dim i As Long       ' ������� ������������� ������� �� �������
Dim fld
        Do
'    ' ��������� ���������� ������� � ��� ������� ��������� � ����� ����������� ���������� ��������
            i = 0: lLen = 0: lCount = 0
            bDate = .Fields(c_strDateBeg)
'Debug.Print bDate, .Fields("ID"), .Fields("DateDesc"), .Fields("Comment")
        ' ������������� �������
            If Not IsNull(.Fields(c_strPeriodType)) Then
            ' �������� ���������� ������ �������� �� ��������� ���� �������� �������
                lCount = DateDiffEx(.Fields(c_strPeriodType), bDate, Date1, wks:=wks, dbs:=dbs)
            ' �������� ��������� ���� �������������� �������
                bDate = DateAddEx(.Fields(c_strPeriodType), lCount, bDate, wks:=wks, dbs:=dbs)
            ' �������� ���������� �������� ������� ����� ��������� � �������� ����� �������� �������
                lCount = DateDiffEx(.Fields(c_strPeriodType), bDate, Date2, wks:=wks, dbs:=dbs)
            End If
        ' �������� ������� - �������� ������������ �������
            If Not IsNull(.Fields(c_strDateEnd)) Then lLen = .Fields(c_strDateEnd) - .Fields(c_strDateBeg) ': Stop
            iDate = bDate
        ' ������������� ������� - ��������� �������� � ��������� ����
            If Not IsNull(.Fields(c_strOffsetType)) Then iDate = DateAddEx(.Fields(c_strOffsetType), .Fields(c_strOffsetValue), bDate)
            eDate = iDate + lLen
            Do While iDate <= Date2
            ' ��������� ������
                ' ���� �������� � ������ - ��������� �������, ����� - ��������� ����
'Debug.Print iDate, .Fields("ID"), .Fields("DateDesc")
'Stop
                If eDate < Date1 Then GoTo HandleNext
                If iDate < Date1 Then GoTo HandleNext
            ' ��������� ����
                rst.AddNew
                For Each fld In arrFields
                    rst.Fields(fld) = IIf(fld = c_strDateBeg, iDate, .Fields(fld))
                Next fld
                rst.Update
HandleNext: ' ��������� � ���������� �������� ����� ��� ������� �������
                ' �������� �������
                If (eDate - iDate) >= 1 Then
                    iDate = Int(iDate + 1)  ' ����������� ����/������ ��� ���� ����� ������� � ���������� � �������� �������
                ElseIf (eDate - iDate) > 0 Then
                    iDate = eDate           ' �������� ���� ��������� ������� �� �������� ����������
                ElseIf (iDate < Date2) And (lCount > 0) Then
                ' ������������� �������
                    i = i + 1: iDate = bDate
                    If Not IsNull(.Fields(c_strPeriodType)) Then iDate = DateAddEx(.Fields(c_strPeriodType), i * .Fields(c_strPeriodValue), bDate)
                    If Not IsNull(.Fields(c_strOffsetType)) Then iDate = DateAddEx(.Fields(c_strOffsetType), .Fields(c_strOffsetValue), iDate)
                    eDate = iDate + lLen
                Else: Exit Do
                End If
            Loop
            .MoveNext
        Loop Until .EOF
    End With
' ��������� � ������� ������� ������� (����� ��������, ������������� ��� �������������), ���������� � ������
    ' ������ �� ���� ������� (����������� �������)
    ' � ��������� ������ ������� (������� ������� ��� �������� ��� ���� ����� ���� �������� ��� ��� �����)
    strWhere = strWhere & sqlAnd & "(" & c_strKey & sqlNot & sqlIn & "(" & _
        sqlSelect & c_strDatesTable & "." & c_strKey & sqlFrom & "[" & CurrentDb.Name & "].[" & TempName & "]" & sqlAs & cTmp & _
        sqlInner & sqlJoin & c_strDatesTable & sqlOn & _
        "(" & cTmp & "." & c_strDateBeg & sqlEqual & "[" & c_strDatesTable & "]." & c_strDateBeg & ")" & _
        sqlAnd & "(" & cTmp & "." & c_strDateType & sqlEqual & "[" & c_strDatesTable & "]." & c_strDateType & ")" & _
        sqlWhere & "[" & c_strDatesTable & "]![" & c_strDateDesc & "]" & sqlIsNull & "))"
' ��������� ������ �� ����������
    strSQL = sqlInsert & sqlInto & "[" & CurrentDb.Name & "].[" & TempName & "] (" & strFields & ") " & _
             sqlSelect & strFields & sqlFrom & c_strDatesTable & _
             sqlWhere & strWhere & sqlAnd & strWhereTemp & sqlAnd & strWhereDate & ";"
    ' ������ ��������� ������� � ��������� ���� � ��������� ������
    With wks: .BeginTrans: bolTransOpen = True: dbs.Execute strSQL: .CommitTrans: bolTransOpen = False: End With
    
    rstSrc.Close: Set rstSrc = Nothing
    rst.Close
'' ��������� ������ ������
'    ''CurrentDb.TableDefs.Refresh
'    'Application.RefreshDatabaseWindow
    m_datTempBeg = Date1: m_datTempEnd = Date2: m_strTempName = TempName

HandleResult:
' ��������� ��������� �� ���� � ����������
    If Len(Fields) = 0 Then strFields = sqlAll Else strFields = Fields
    strOrder = Join(Array(c_strDateBeg, c_strDateType), ",")
    ' ������ ��� ���� ������� ������� (��� ���������)
    strWhere = strWhereDate
    ' ��������� ������ ��� ������ ���������� ������� (��� ��������)
    If Unique Then strWhere = strWhere & sqlAnd & "(" & c_strKey & sqlEqual & "(" & _
            sqlSelect & sqlTop1 & cTmp & ".ID" & sqlFrom & "[" & m_strTempName & "]" & sqlAs & cTmp & _
            sqlWhere & cTmp & ".[" & c_strDateBeg & "]=[" & m_strTempName & "].[" & c_strDateBeg & "]" & _
            sqlOrder & cTmp & ".[" & c_strDateType & "]," & cTmp & ".[" & c_strKey & "])" & ")"
    strSQL = sqlSelect & strFields & sqlFrom & "[" & m_strTempName & "]" & sqlWhere & strWhere & sqlOrder & strOrder & ";"
    Set p_TempTableOpen = CurrentDb.OpenRecordset(strSQL) ': rst.MoveFirst
HandleExit:  Exit Function
HandleError: If bolTransOpen Then wks.Rollback: bolTransOpen = False     ' ���������� ����������
    Select Case Err.Number
    Case 3211: Stop ' �� ���� ������� ��������� ������� - �������������
    Case 3734: Stop ' ���� ������ ���� ��������� ������������� � ���������, �������������� �� �������� ��� ����������
    End Select
    Result = False: Err.Clear: Resume HandleExit
End Function

Private Function p_TableDrop(Source, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Boolean
' �������  �������
'-------------------------
' Source    - ��� ��������� ������� ��� ������ DAO.Recordset, �������� �� ���
' dbs,wks   - ������ �� ���� ������ �� ������� �������������� ������ � ������� ������������
'-------------------------
Dim strSQL As String, strSource As String, strTarget As String, strFields As String
Dim bolTransOpen As Boolean
    If wks Is Nothing Then Set wks = DBEngine.Workspaces(0): Set dbs = CurrentDb                                         ' ��-��������� ������� ���� � ������� ������������
    If dbs Is Nothing Then If wks.Databases.Count > 0 Then Set dbs = wks.Databases(0) Else Err.Raise vbObjectError + 512 ' ���� �� ������ ����, �� ������ ������� ������������ - ���� ������ � ������� ������������, ����� - ������
Dim Result As Boolean ': Result = False
    On Error GoTo HandleError
    If VarType(Source) = vbString Then
    ' �������� - SQL ���������� ��� ��� �������
        Source = Trim$(Source): If Len(Source) = 0 Then Err.Raise vbObjectError + 512
        strSource = Source
    ElseIf TypeOf Source Is DAO.Recordset Then
    ' �������� - DAO.Recordset - ����� ��� ���������
        strSource = Source.Name
        Source.Close: Set Source = Nothing
    Else
    ' �������� �� ��������
        Err.Raise vbObjectError + 512
    End If
    With wks
        If p_IsTableExists(strSource, dbs, wks) Then .BeginTrans: dbs.Execute sqlDropTable & "[" & strSource & "]": .CommitTrans ': bolTransOpen = False
    End With
    Result = True
HandleExit:  p_TableDrop = Result: Exit Function
HandleError: If bolTransOpen Then wks.Rollback: bolTransOpen = False     ' ���������� ����������
    Select Case Err.Number
    Case 3211: Stop ' �� ���� ������� ��������� ������� - �������������
    Case 3734: Stop ' ���� ������ ���� ��������� ������������� � ���������, �������������� �� �������� ��� ����������
    End Select
    Result = False: Err.Clear: Resume HandleExit
End Function

Private Function p_IsTableExists(ByVal TableName As String, _
    Optional dbs As DAO.Database, Optional wks As DAO.Workspace _
    ) As Boolean
' ���������� �������� True, ���� ���� ������� � ����� ������.
'-------------------------
' dbs,wks   - ������ �� ���� ������ �� ������� �������������� ������ � ������� ������������
'-------------------------
Dim Result As Boolean ': Result = False
    On Error GoTo HandleError
Dim strSQL As String
' ---
    If wks Is Nothing Then Set wks = DBEngine.Workspaces(0): Set dbs = CurrentDb                                         ' ��-��������� ������� ���� � ������� ������������
    If dbs Is Nothing Then If wks.Databases.Count > 0 Then Set dbs = wks.Databases(0) Else Err.Raise vbObjectError + 512 ' ���� �� ������ ����, �� ������ ������� ������������ - ���� ������ � ������� ������������, ����� - ������
' ---
'' ������ ������� �������� ������ TableDef
'    Result = dbs.TableDefs(TableName).Name = TableName
'' ������ ������� �������� AllTables
'    With CurrentData.AllTables(TableName)
'        IsLoaded = .IsLoaded
'        Result = .Name = TableName
'    End With
' ������ ������� �������� - ������� ������� ���������
    strSQL = TableName
    'strSQL = sqlSelect1st & "[" & TableName & "]"
    'If IsMissing(Hash) Then strSQL = sqlSelect1st & "(" & strSQL & ")"
Dim rst As DAO.Recordset: Set rst = dbs.OpenRecordset(strSQL, dbOpenTable)
    rst.Close: Set rst = Nothing
    Result = True
HandleExit:  p_IsTableExists = Result: Exit Function
HandleError: Select Case Err
    Case 3008: Result = True ' ������� ������ ������������� ��� ������������ �������������
    Case Else: Result = False
    End Select
    Err.Clear: Resume HandleExit
End Function

Private Function p_DateToSQL(FormatDate) As String
' ����������� ����/����� ��� ������������� � SQL ��������
Dim strTemp As String: strTemp = "m\/d\/yyyy": If (FormatDate - Int(FormatDate)) Then strTemp = strTemp & " h\:n\:s"
    p_DateToSQL = Format$(FormatDate, "\#" & strTemp & "\#")
End Function


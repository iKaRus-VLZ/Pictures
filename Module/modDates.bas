Attribute VB_Name = "modDates"
Option Compare Database
Option Explicit
'=========================
Private Const c_strModule As String = "modDates"
'=========================
' Описание      : Функции для работы с датами
' Версия        : 1.0.7.453565659
' Дата          : 05.03.2024 13:34:54
' Автор         : Кашкин Р.В. (KashRus@gmail.com)
' Примечание    : для хранения дат использует таблицу SysCalendar, может подгружать производственный календарь с buh.ru
' v.1.0.7       : 20.02.2024 - добавлена обработка периодических, относительных и длящихся событий на основе временной таблицы
' v.1.0.5       : 07.02.2024 - переписаны функции работы с рабочими днями для большего сходства с аналогичными в Excel 2010+
' v.1.0.4       : 27.12.2023 - добавлено обновление производственного календаря через интернет с buh.ru
' v.1.0.2       : 22.03.2019 - добавлена функция расчета количества рабочих/выходных дней между датами
'=========================
' Задачи делать очередной планировщик не было, нужна была только корректная работа с рабочими днями,
' поэтому нет нормальных форм настройки/редактирования событий (заменены заглушками)
' дальнейшее было сделано просто для большей универсальности и на будущее,- а вдруг понадобится?))
' Приоритет событий (по-убыванию): рабочий, предпраздничный, праздничный, нерабочий, пользовательский
' т.е. если для любого дня создать событие рабочий день он будет считаться рабочим
'=========================
' ToDo: сделать нормальное удаление временной таблицы после её использования
'       сделать формы редактирования событий и редактирования списка событий (см. p_DatesTableEventEdit, p_DatesTableEventsList)
'=========================
Private Const NOERROR As Long = 0
' ссылка для загрузки производственного календаря (см. p_HolidaysFromWeb)
Private Const с_strLink = "https://buh.ru/calendar/" ' "https://www.consultant.ru/law/ref/calendar/proizvodstvennye/"
' описание таблицы праздничных и выходных дней
Private Const c_strDatesTable = "SysCalendar"                                           ' название таблицы календаря дат
Private Const c_strKey = "ID", c_strParent = "PARENT"                                   ' ключевое/родительское поле
Private Const c_strDateType = "DATETYPE", c_strDateDesc = "DATEDESC"                    ' тип/описание даты
Private Const c_strDateBeg = "DATEBEG", c_strDateEnd = "DATEEND"                        ' для обычных и длящихся (дата начала/конца)
Private Const c_strOffsetType = "OFFSET", c_strOffsetValue = c_strOffsetType & "VAL"    ' для относительных (тип смещения/величина смещения)
Private Const c_strPeriodType = "PERIOD", c_strPeriodValue = c_strPeriodType & "VAL"    ' для периодических (тип периода/величина периода)
Private Const c_strActBegDate = "ACTBEG", c_strActEndDate = "ACTEND"                    ' дата начала/окончания актуальности записи
Private Const c_strEditDate = "EDITDATE", c_strComment = "COMMENT"                      ' дата изменения записи/комментарий

Private Const c_strTmpTablePref = "@&%"                                                  ' префикс временной таблицы
Private m_datTempBeg As Date, m_datTempEnd As Date                                      ' период за который сформирована временная таблица содержащая ленту событий
Private m_strTempName As String                                                         ' имя временной таблицы содержащей ленту событий за указанный перриод

' дополнительные значения Interval для функции DateDiffEx/DateAddEx
Public Const c_strWorkdayLiteral = "wd", c_strWorkdayLiteral2 = "workday"          ' рабочие дни
Public Const c_strNonWorkdayLiteral = "hd", c_strNonWorkdayLiteral2 = "holiday"    ' нерабочие дни ??
    ' дни недели
Public Const c_strMondayLiteral = "mon", c_strMondayLiteral2 = "monday", _
      c_strMondayLiteral1 = "пн", c_strMondayLiteral3 = "понедельник"       ' понедельники
Public Const c_strTuesdayLiteral = "tues", c_strTuesdayLiteral2 = "tuesday", _
      c_strTuesdayLiteral1 = "вт", c_strTuesdayLiteral3 = "вторник"         ' вторники
Public Const c_strWednesdayLiteral = "wed", c_strWednesdayLiteral2 = "wednesday", _
      c_strWednesdayLiteral1 = "ср", c_strWednesdayLiteral3 = "среда"       ' среды
Public Const c_strThursdayLiteral = "thur", c_strThursdayLiteral2 = "thursday", _
      c_strThursdayLiteral1 = "чт", c_strThursdayLiteral3 = "четверг"       ' четверги
Public Const c_strFridayLiteral = "fri", c_strFridayLiteral2 = "friday", _
      c_strFridayLiteral1 = "пт", c_strFridayLiteral3 = "пятница"           ' пятницы
Public Const c_strSaturdayLiteral = "sat", c_strSaturdayLiteral2 = "saturday", _
      c_strSaturdayLiteral1 = "сб", c_strSaturdayLiteral3 = "суббота"       ' субботы
Public Const c_strSundayLiteral = "sun", c_strSundayLiteral2 = "sunday", _
      c_strSundayLiteral1 = "вс", c_strSundayLiteral3 = "воскресенье"       ' воскресения
' тип дня рабочего календаря по справочнику  (см.DateTypes)
Public Enum eDateType
    eDateTypeUndef = 0              ' 0   <не задано>
    ' календарные дни
    eDateTypeWeekday = 1            ' 1   рабочий (будний) день
    eDateTypeSatday = 6             ' 6   выходной (суббота) день
    eDateTypeSunday = 7             ' 7   выходной (воскресенье) день
    ' события
    eDateTypeWorkday = 10           ' 10  рабочий день
    eDateTypeHolidayPre = 20        ' 20  рабочий (сокращенный) день
    eDateTypeHoliday = 70           ' 70  выходной (праздничный) день
    eDateTypeNonWorkday = 80        ' 80  нерабочий день
    eDateTypeUser = 99              ' 99  дата определённая пользователем
    ' группы событий
    eDateServOfficial = 500         ' 500 (служебная группа) государственные праздники, памятные даты и прочие даты введённые указами и постановлениями
    eDateServWorkCalendar = 900     ' 900 (служебная группа) информация о загрузке данных производственного календаря
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
' Имитации функций листа Excel и расширенные функции для работы с датами
'=========================
Public Function AGE(BirthDate As Date, Optional TestDate) As Long
' Возвращает возраст (количество полных лет) человека на указанную дату
'-------------------------
' BirthDate - дата рождения
' TestDate  - дата на которую определяется возраст
'-------------------------
Dim Result As Long ': Result = False
    On Error GoTo HandleError
    If Not IsDate(TestDate) Then TestDate = Date
    If TestDate >= BirthDate Then Result = DateDiff("yyyy", BirthDate, TestDate) + (DateSerial(Year(TestDate), Month(BirthDate), Day(BirthDate)) > TestDate)
HandleExit:  AGE = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function MONTHEND(Date1 As Date) As Date
' Возвращает дату конца месяца для заданной даты
'-------------------------
    MONTHEND = DateSerial(Year(Date1), Month(Date1) + 1, 1) - 1
End Function

Public Function WEEKDAYDATE(Date1 As Date, Number As Double, Weekday1 As VbDayOfWeek) As Date
' Возвращает дату, отстоящую на заданное количество указанных дней недели вперед или назад от начальной даты.
'-------------------------
' Date1     - начальная дата от которой отсчитываем период
' Weekday1  - искомый день недели
' Number    - количество указанных дней недели, на которое должна отстоять искомая дата от начальной
'   1,2..   - день недели после начальной даты (включая её)
'   0,-1..  - день недели перед начальной датой (не включая её)
'-------------------------
' Например: дата предпоследнего воскресенья текущего месяца: WEEKDAYDATE (MONTHEND(Now)+1,-1,vbSunday)
'-------------------------
'' если надо - сделать параметры ByVal и раскомментировать:
'    If Number < 0 Then Number = Number + 1 ' при отсчете назад начинать с -1
'    If Number <= 0 Then Date1 = Date1 + 1  ' при отсчете назад учитывать начальную дату
    WEEKDAYDATE = DateAdd("ww", Number, Date1 - 1) - (Date1 - Weekday1 - 1) Mod 7
End Function

Public Function WEEKDAYCOUNT(ByVal Date1 As Date, ByVal Date2 As Date, Weekday1 As VbDayOfWeek) As Long
' Возвращает количество указанных дней недели между начальной датой и конечной датой.
'-------------------------
' Date1     - начальная дата периода
' Date2     - конечная дата периода
' Weekday1  - искомый день недели
'-------------------------
Dim Sign As Boolean, Temp As Date: If Date1 > Date2 Then Sign = True: Temp = Date2: Date2 = Date1: Date1 = Temp ' проверяем последовательность переданных дат и делаем левую границу меньшей, а правую большей
    Date2 = Date2 - Weekday1 + 1: If Weekday1 >= Weekday(Date1) Then Date2 = Date2 + 7
    WEEKDAYCOUNT = DateDiff("ww", Date1, Date2): If Sign Then WEEKDAYCOUNT = WEEKDAYCOUNT * -1
'' или так?
'    If Date1 <= Date2 Then
'        WEEKDAYCOUNT = DateDiff("ww", Date1, Date2 - Weekday1 + IIf(Weekday1 < Weekday(Date1), 1, 8))
'    Else
'        WEEKDAYCOUNT = DateDiff("ww", Date1 - Weekday1 + IIf(Weekday1 < Weekday(Date2), 1, 8), Date2)
'    End If
End Function

Public Function ISWORKDAY(Date1 As Date, _
        Optional Holidays = False, Optional Weekends, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Boolean
' Возвращает является день рабочим. Проверяет дату по заданным массивам выходных и праздников
'-------------------------
' Date1     - проверяемая дата
' Holidays  - упорядоченный массив назначенных дат (праздничных и нерабочих дней, а также рабочих выходных дней), либо:
'             0 - даты будут получены из таблицы, 1 - из интернета, любое другое - праздники не будут учитываться
' Weekends  - какие дни недели являются выходными и не считаются рабочими (см. p_Weekends). по-умолчанию: суббота, воскресенье
'-------------------------
Dim Result As Boolean: On Error GoTo HandleError
    If IsMissing(Holidays) Then
' проверяем день по списку выходных
        Result = Not ISWEEKEND(Date1, Weekends)
    Else
' проверяем по Holidays
        Select Case p_CheckForHolidays(Date1, Date1, Holidays, Weekends, dbs, wks)
        Case 1:     Result = True                           ' +1 - рабочий
        Case -1:    Result = False                          ' -1 - выходной
        Case Else:  Result = Not ISWEEKEND(Date1, Weekends) ' всё остальное (в данном случае) - ошибка чтения Holidays
        End Select
    End If
HandleExit:  ISWORKDAY = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Function ISWEEKEND(Date1, Optional Weekends) As Boolean
' Возвращает является день выходным. Проверяет дату по заданному массиву выходных (не учитывает праздники и др нерабочие дни)
'-------------------------
' Date1     - проверяемая дата
' Weekends  - какие дни недели являются выходными и не считаются рабочими (см. p_Weekends). по-умолчанию: суббота, воскресенье
'-------------------------
Dim Result As Boolean: On Error GoTo HandleError
Dim uw As Long: uw = UBound(p_Weekends(Weekends)) ' при необходимости инициализируем массив выходных дней и берём верхнюю границу
Dim i As Long, d As VbDayOfWeek: d = DatePart("w", Date1): Do: Result = (p_Weekends()(i) = d): i = i + 1: Loop While i <= uw And Not Result   ' проверяем день по списку выходных
HandleExit:  ISWEEKEND = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Sub SetWeekends(Optional Weekends)
' Меняет набор дней недели, которые являются выходными и не считаются рабочими.
'-------------------------
' Weekends  - какие дни недели являются выходными и не считаются рабочими. по-умолчанию: суббота, воскресенье
'   числовое значение 1-7, 11-17
'   или строка вида: 0000011, где 0- рабочий день, 1-выходной (начиная с понедельника)
'-------------------------
    If IsMissing(Weekends) Then Weekends = 1 Else If Len(Weekends) = 0 Then Weekends = 1
    Call p_Weekends(Weekends)
End Sub

Public Function NETWORKDAYS(Date1 As Date, Date2 As Date, _
        Optional Holidays = False, Optional Weekends, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Long
' Возвращает количество рабочих дней между начальной датой и конечной датой.
'-------------------------
' Date1     - начальная дата периода
' Date2     - конечная дата периода
' Holidays  - упорядоченный массив назначенных дат (праздничных и нерабочих дней, а также рабочих выходных дней), либо:
'             0 - даты будут получены из таблицы, 1 - из интернета, любое другое - праздники не будут учитываться
' Weekends  - какие дни недели являются выходными и не считаются рабочими (см. p_Weekends). по-умолчанию: суббота, воскресенье
'-------------------------
' Рабочими днями не считаются выходные дни (кроме указанных в Holidays как рабочие выходные) и дни определенные в Holidays как праздничные
'-------------------------
Dim Result As Long: On Error GoTo HandleError
Dim Sign As Boolean, Temp As Date: If Date1 > Date2 Then Sign = True: Temp = Date2: Date2 = Date1: Date1 = Temp ' проверяем последовательность переданных дат и делаем левую границу меньшей, а правую большей
Dim uw As Long: uw = UBound(p_Weekends(Weekends))             ' UBound(p_Weekends)+1 = количество нерабочих дней в неделе
' получаем количество рабочих/выходных дней в полных неделях (без учёта праздников)
Dim ww As Long: ww = DateDiff("w", Date1, Date2)    ' полных недель в периоде
    Result = (6 - uw) * ww  ': hd = (uw+1) * ww
' отбрасываем ранеее учтённые дни полных недель и проверяем остаток
    Temp = DateAdd("ww", ww, Date1)
    Do While Temp <= Date2 ' <=
        If Not ISWEEKEND(Temp) Then Result = Result + 1 ' добавляем к рабочим
        Temp = DateAdd("d", 1, Temp)                    ' следующий день
    Loop
' довычитаем праздничные и нерабочие
    If IsMissing(Holidays) Then GoTo HandleExit
    Result = Result + p_CheckForHolidays(Date1, Date2, Holidays, Weekends, dbs, wks)
HandleExit:  If Sign Then Result = -Result              ' если надо - меняем знак
             NETWORKDAYS = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Function WORKDAY(Date1 As Date, Number As Double, _
        Optional Holidays = False, Optional Weekends, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Date
' Возвращает дату, отстоящую на заданное количество рабочих дней вперед или назад от начальной даты.
'-------------------------
' Date1     - начальная дата от которой отсчитываем период
' Number    - количество рабочих дней, на которое должна отстоять искомая дата от начальной
' Holidays  - упорядоченный массив назначенных дат (праздничных и нерабочих дней, а также рабочих выходных дней), либо:
'             0 - даты будут получены из таблицы, 1 - из интернета, любое другое - праздники не будут учитываться
' Weekends  - какие дни недели являются выходными и не считаются рабочими (см. p_Weekends). по-умолчанию: суббота, воскресенье
'-------------------------
' Рабочими днями не считаются выходные дни (кроме указанных в Holidays как рабочие выходные) и дни определенные в Holidays как праздничные
'-------------------------
' Сделано корявенько, но работоспособно
'-------------------------
Dim Result As Date: Result = Date1
    On Error GoTo HandleError
Dim d As Integer:   d = Sgn(Number)     ' направление смещения
Dim n As Double:    n = Abs(Number)     ' величина смещения от заданной даты
Dim ww As Long                          ' количество полных недель в периоде
Dim wd As Long                          ' количество рабочих дней
Dim WorkDaysPerWeek As Long:    WorkDaysPerWeek = 6 - UBound(p_Weekends(Weekends))  ' количество рабочих дней в неделе (без учёта праздников)
    ww = n \ 7
    Do While n  '<> 0
    ' считаем без учета праздников и нерабочих дней (учитываем только Weekends)
        If ww = 0 Then
    ' смещаем дату на один день и проверяем является ли он рабочим
            Result = DateAdd("d", d, Result):       wd = Abs(ISWORKDAY(Result, Holidays, Weekends, dbs, wks))
        Else
    ' смещаем дату на полное количество рабочих недель
Dim db As Date: db = Result
            Result = DateAdd("ww", d * ww, Result): wd = ww * WorkDaysPerWeek
    ' доучитываем праздники и нерабочие дни
            If Not IsMissing(Holidays) Then wd = wd + p_CheckForHolidays(db + d, Result, Holidays, Weekends, dbs, wks)
        End If
    ' проверяем сколько рабочих дней осталось найти
        n = n - wd: ww = n \ 7
    Loop
HandleExit:  WORKDAY = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Function DateAddEx(Interval As String, ByVal Number As Double, Date1 As Date, _
        Optional FirstDayOfWeek As VbDayOfWeek, Optional FirstWeekOfYear As VbFirstWeekOfYear, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Date
' Расширенный DateAdd позволяющий обрабатывать дополнительные интервалы
'-------------------------
' Interval..- аналогично стандартному DateDiff
' dbs,wks   - ссылка на базу в которой расположен используемый  календарь
'-------------------------
Dim Result As Long
    On Error GoTo HandleError
    Select Case Interval
' дата отстоящая на заданное количество рабочих/нерабочих дней от указанной даты
    Case c_strWorkdayLiteral, c_strWorkdayLiteral2:                                                     Result = WORKDAY(Date1, Number)
    Case c_strNonWorkdayLiteral, c_strNonWorkdayLiteral2:                                               Result = DateAdd("d", Number, Date1) - WORKDAY(Date1, Number)
' дата отстоящая на заданное количество указанных дней недели от указанной даты
    Case c_strSundayLiteral, c_strSundayLiteral1, c_strSundayLiteral2, c_strSundayLiteral3:             Result = WEEKDAYDATE(Date1, Number, vbSunday)
    Case c_strMondayLiteral, c_strMondayLiteral1, c_strMondayLiteral2, c_strMondayLiteral3:             Result = WEEKDAYDATE(Date1, Number, vbMonday)
    Case c_strTuesdayLiteral, c_strTuesdayLiteral1, c_strTuesdayLiteral2, c_strTuesdayLiteral3:         Result = WEEKDAYDATE(Date1, Number, vbTuesday)
    Case c_strWednesdayLiteral, c_strWednesdayLiteral1, c_strWednesdayLiteral2, c_strWednesdayLiteral3: Result = WEEKDAYDATE(Date1, Number, vbWednesday)
    Case c_strThursdayLiteral, c_strThursdayLiteral1, c_strThursdayLiteral2, c_strThursdayLiteral3:     Result = WEEKDAYDATE(Date1, Number, vbThursday)
    Case c_strFridayLiteral, c_strFridayLiteral1, c_strFridayLiteral2, c_strFridayLiteral3:             Result = WEEKDAYDATE(Date1, Number, vbFriday)
    Case c_strSaturdayLiteral, c_strSaturdayLiteral1, c_strSaturdayLiteral2, c_strSaturdayLiteral3:     Result = WEEKDAYDATE(Date1, Number, vbSaturday)
' стандартный вывод
    Case Else:                                                                                          Result = DateAdd(Interval, Number, Date1)
    End Select
HandleExit:  DateAddEx = Result: Exit Function
HandleError: Resume HandleExit
End Function

Public Function DateDiffEx(Interval As String, ByVal Date1 As Date, ByVal Date2 As Date, _
        Optional FirstDayOfWeek As VbDayOfWeek, Optional FirstWeekOfYear As VbFirstWeekOfYear, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Long
' Расширенный DateDiff позволяющий обрабатывать дополнительные интервалы
'-------------------------
' Interval..- аналогично стандартному DateDiff
' dbs,wks   - ссылка на базу в которой расположен используемый  календарь
'-------------------------
' !!! не совсем понятно конец или начало периода не учитывает оригинальный DateDiff, поэтому,- волевым решением:
' конечная дата периода не учитывается (считаем количество переходов между интервалами на указанном диапазоне)
' поэтому для подсчёта количества дней как в Between Date1 And Date2 (включая границы) надо вместо Date2 параметром функции указывать DateAdd(Interval, 1, Date2)
' т.е между 01/01/2015 и 31/12/2015 будет 364 календарных дня,  260 рабочих и 104 выходных (246/118 с учётом праздников), - выпадает 31/12/2015 (рабочий день)
' а между   01/01/2016 и 31/12/2016 будет 365 календарных дней, 260 рабочих и 105 выходных (247/118 с учётом праздников), - выпадает 31/12/2016 (выходной день)
'-------------------------
Dim Result As Long
    On Error GoTo HandleError
Dim ww As Long, wd As Long, hd As Long, cd As Date, sg As Integer ', dd As Long
    Select Case Interval
' количество рабочих/нерабочих дней  между двумя датами
    Case c_strWorkdayLiteral, c_strWorkdayLiteral2:                                                     Result = NETWORKDAYS(Date1, Date2)
    Case c_strNonWorkdayLiteral, c_strNonWorkdayLiteral2:                                               Result = DateDiff("d", Date1, Date2) - NETWORKDAYS(Date1, Date2)
' количество заданных дней недели между двумя датами
    Case c_strSundayLiteral, c_strSundayLiteral1, c_strSundayLiteral2, c_strSundayLiteral3:             Result = WEEKDAYCOUNT(Date1, Date2, vbSunday)
    Case c_strMondayLiteral, c_strMondayLiteral1, c_strMondayLiteral2, c_strMondayLiteral3:             Result = WEEKDAYCOUNT(Date1, Date2, vbMonday)
    Case c_strTuesdayLiteral, c_strTuesdayLiteral1, c_strTuesdayLiteral2, c_strTuesdayLiteral3:         Result = WEEKDAYCOUNT(Date1, Date2, vbTuesday)
    Case c_strWednesdayLiteral, c_strWednesdayLiteral1, c_strWednesdayLiteral2, c_strWednesdayLiteral3: Result = WEEKDAYCOUNT(Date1, Date2, vbWednesday)
    Case c_strThursdayLiteral, c_strThursdayLiteral1, c_strThursdayLiteral2, c_strThursdayLiteral3:     Result = WEEKDAYCOUNT(Date1, Date2, vbThursday)
    Case c_strFridayLiteral, c_strFridayLiteral1, c_strFridayLiteral2, c_strFridayLiteral3:             Result = WEEKDAYCOUNT(Date1, Date2, vbFriday)
    Case c_strSaturdayLiteral, c_strSaturdayLiteral1, c_strSaturdayLiteral2, c_strSaturdayLiteral3:     Result = WEEKDAYCOUNT(Date1, Date2, vbSaturday)
' стандартный вывод
    Case Else:                                                                                          Result = DateDiff(Interval, Date1, Date2, FirstDayOfWeek, FirstWeekOfYear)
    End Select
HandleExit:  DateDiffEx = Result: Exit Function
HandleError: Resume HandleExit
End Function

'=========================
' Функции для работы с настраиваемыми датами
'=========================
Public Function DateTypes(Optional ID, Optional Col, Optional Row, Optional iStep As Long)
' Возвращает массив/значение из справочника типов дат
'-------------------------
' Id        - код искомого элемента arrrData(0)
' Col/Row   - колонка/строка искомого элемента (начиная с 0)
' iStep     - (возвращаемое) количество элементов в строке
'-------------------------
On Error Resume Next
Static aData(): iStep = 6 '[i+0]=ID(eDateType); [i+1]=CNAME;[i+2]=NAME; [i+3]=DESC; [i+54]=DateColor; [i+5]=FaceId
Dim i As Long: i = LBound(aData)
    If Err Then
        Err.Clear
        aData = Array(eDateTypeUndef, "undef", "<не задано>", "<не задано>", "Black", "", _
                eDateTypeWeekday, "weekday", "Рабочий (будний)", "рабочий (будний) день", "Navy", "DaysWork", _
                eDateTypeSatday, "dayoff0", "Субботоний", "выходные дни (Суббота)", "HotPink", "DaysDayOff", _
                eDateTypeSunday, "dayoff", "Выходной", "выходные дни", "DeepPink", "DaysDayOff", _
                eDateTypeWorkday, "work", "Рабочий", "рабочий день", "Navy", "DaysWork", _
                eDateTypeHolidayPre, "holiday_pre", "Предпраздничный", "предпраздничные (сокращённые) дни", "MediumPurple", "DaysHolidayPre", _
                eDateTypeHoliday, "holiday", "Праздничный", "праздничные дни", "Red", "DaysHoliday", _
                eDateTypeNonWorkday, "nonworking", "Нерабочий", "нерабочие дни", "PaleVioletRed", "DaysNonWorking", _
                eDateTypeUser, "userday", "Настраиваемый", "даты определённые пользователем", "Teal", "DaysUser", _
                eDateServOfficial, "official", "[official]", "официальные даты, определённые указами", "", "", _
                eDateServWorkCalendar, "workdays", "[workdays]", "производственный календарь", "")
    End If
' по умолчанию возвращаем весь массив
    If IsMissing(ID) And IsMissing(Row) And IsMissing(Col) Then DateTypes = aData: Exit Function ':Result = aData: GoTo HandleExit
On Error GoTo HandleError
Dim Result
' указан индекс элемента - перебираем строки массива ищем индекс
    If Not IsMissing(ID) Then
        Row = 0
        For i = LBound(aData) To UBound(aData) Step iStep
            If aData(i) = ID Then Row = i \ iStep:  Exit For
        Next i
    End If
    If IsMissing(Row) Then
' указана только колонка                        - возвращаем массив элементов колонки
        i = (UBound(aData) - LBound(aData) + 1)
        i = i \ iStep + Abs((i Mod iStep) > 0)
        ReDim Result(0 To i - 1)
        For i = LBound(Result) To UBound(Result)
            Result(i) = aData(i * iStep + Col)
        Next i
    ElseIf IsMissing(Col) Then
' указана только строка                         - возвращаем массив элементов строки
        ReDim Result(0 To iStep - 1)
        For i = LBound(Result) To UBound(Result)
            Result(i) = aData(Row * iStep + i)
        Next i
    Else
' указана строка и колонка                      - возвращаем элемент из указанной колонки указанной строки
        Result = aData(Row * iStep + Col)
    End If
HandleExit:  DateTypes = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Function DropTempCalendar() As Boolean
' удаляет временную таблицу
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
' Возвращает код события основной даты и описание события
'-------------------------
' Date1     - дата для проверки
' DateDesc  - дополнительная информация о дне
' EventId   - (возвращаемое) ключ основного события соответствующего дате
' EventIds  - (возвращаемое) строка ключей ВСЕХ событий соответствующих дате
' AskTable  - если True, при отсутствии будет предлагать создать таблицу дат
'-------------------------
' проверяет по календарю и таблице c_strDatesTable у таблицы приоритет, например:
' если дата приходится на воскресенье, а в таблице день помечен как рабочий - возвращается рабочий день
' на одну дату может приходиться несколько разных событий, тогда выбираем основное, описания остальных через запятую
'-------------------------
' создаёт на основе календаря временную таблицу с лентой событий на год и читает из неё - очень тупо, но лучше идей небыло
'-------------------------
Const cstrDelim = ","
Const cstrDescDelim = "; "
Dim Result As eDateType: Result = eDateTypeUndef
    On Error GoTo HandleError
    Result = eDateTypeUndef: EventId = 0 ': EventIds = vbNullString
'Dim datDateBeg As Date:  datDateBeg = Date1
' если таблицы нет переходим к получению информации по дню недели
    If Not p_DatesTableExists(AskTable:=AskTable, dbs:=dbs, wks:=wks) Then GoTo HandleResult
'' определяем тип по таблице настраиваимых дат
'    ' выбираем типы дат (не группа/не служебная информация)
'Dim strTypes As String:     strTypes = Join(Array(eDateTypeWeekday, eDateTypeHolidayPre, eDateTypeHoliday, eDateTypeNonWorkday, eDateTypeUser), ",")
' открываем ленту событий для заданной даты на временной таблице - после завершениы работы - удалить
Dim rst As DAO.Recordset
    Select Case Date1
    Case m_datTempBeg To m_datTempEnd       ' период за который сформирована временная таблица содержащая ленту событий
        ' проверить наличие таблицы - если нет - создать и заполнить
    Case Else                               ' нет ленты событий для запрошенноq даты - создаём сразу на полный год
    ' создаём новую временную таблицу, период - полный год
        Call p_TempTableOpen(DateSerial(Year(Date1), 1, 1), DateSerial(Year(Date1), 12, 31), TempName:=m_strTempName, dbs:=dbs, wks:=wks)
    End Select
' открываем запрос для конкретной даты
Dim strSQL As String: strSQL = sqlSelectAll & "[" & m_strTempName & "]" & _
                sqlWhere & c_strDateBeg & sqlEqual & p_DateToSQL(Int(Date1)) & sqlOrder & c_strDateType & ";"
        Set rst = CurrentDb.OpenRecordset(strSQL)
    With rst
        EventIds = vbNullString
        If Not (.BOF And .EOF) Then .MoveFirst
Dim tmpType As eDateType, strDesc As String, tmpDesc As String
        Do Until .EOF
' перебираем все записи относящиеся к дате и:
    ' выбираем основной тип даты
    ' собираем описания всех относящихся к дате событий
    ' в Result текущий основной тип даты
    ' в tmpType текущий тип даты по справочнику
            'tmpType = .Fields(c_strDateType)
            EventIds = EventIds & cstrDelim & .Fields(c_strKey)
            tmpDesc = Trim(Nz(.Fields(c_strDateDesc).Value, vbNullString)): If Len(tmpDesc) > 0 Then strDesc = strDesc & cstrDescDelim & tmpDesc
            If Result = eDateTypeUndef Then
                Result = .Fields(c_strDateType): If EventId = 0 Then EventId = .Fields(c_strKey)
            Else
                Select Case .Fields(c_strDateType)
                Case eDateTypeUndef:    ' тип события не задан - оставляем
                Case Is < Result:       ' если данный тип имеет более высокий приоритет - сохраняем его
                ' возможные варианты:
                    'eDateTypeWorkday = 10           ' 10  рабочий день
                    'eDateTypeHolidayPre = 20        ' 20  рабочий (сокращенный) день
                    'eDateTypeHoliday = 70           ' 70  выходной (праздничный) день
                    'eDateTypeNonWorkday = 80        ' 80  нерабочий день
                    'eDateTypeUser = 99              ' 99  дата определённая пользователем
                    Result = .Fields(c_strDateType): EventId = .Fields(c_strKey)
                End Select
            End If
HandleNext: .MoveNext
        Loop
    End With
HandleResult:
' если тип всё ещё не задан - определяем его по дню недели
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
' определяем описание дня
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
' Устанавливает информацию о заданном дне. Возвращает тип созданного события
'-------------------------
' Date1     - дата для которой устанавливаем меняем даты
' DateType  - устанавливаемый тип даты
' ExtInfo   - признак необходимости настройки расширенных параметров события
'-------------------------
' ToDo: тут в общем-то кривуля какая-то, но я так толком и не понял чего сам хочу от неё,- потому так
'-------------------------
Dim Result As eDateType: Result = eDateTypeUndef
    On Error GoTo HandleError
    If Not p_DatesTableExists Then GoTo HandleExit
' проверяем тип устанавливаемого события
    Select Case DateType
    Case Is < eDateTypeWorkday: DateType = eDateTypeUndef   ' если меньше минимального  - ставим неопределённый тип
    Case Is > eDateTypeUser:    DateType = eDateTypeUser    ' если больше максимального - ставим пользовательский тип
    End Select
Dim EventIds As String:         Result = DateInfoGet(Date1, EventIds:=EventIds, dbs:=dbs, wks:=wks)
' открываем запрос со списком событий относящихся к дате
Dim rst As DAO.Recordset, strWhere As String:
Dim bolNew As Boolean: bolNew = Len(EventIds) = 0:  If bolNew Then strWhere = False Else strWhere = c_strKey & sqlIn & "(" & EventIds & ")" ' <<< здесь ошибка на новой записи
Dim bolRst As Boolean: bolRst = rst Is Nothing
    Set rst = p_DatesTableOpen(strWhere, dbs:=dbs, wks:=wks)
    Select Case Result                              ' Result -> DateType
    Case Is < eDateTypeWorkday                      ' текущий статус даты неопределён -> создаём новое событие указанного типа
    Case DateType: Err.Raise vbObjectError + 512    ' текущий статус даты совпадает с устанавливаемым (статус не меняется) -> ошибка
    'Case eDateTypeUser                              ' текущий статус даты пользовательское событие -> создаём новое событие указанного типа
    'Case Is > DateType:                             ' текущий статус даты имеет меньший приоритет, чем устанавливаемый -> создаём новое событие указанного типа
    Case Is < DateType:                             ' текущий статус даты имеет больший приоритет, чем устанавливаемый ->
    ' чтобы установить статус с меньшим приоритетом надо сначала удалить все события относящиеся к дате имеющие более высокий приоритет, про каждое надо спросить пользователя
    ' только после удаления событий с более высоким приоритетом можно добавлять новое событие с заданным статусом иначе статус даты не изменится
Dim strTitle As String, strMessage As String
Dim msgRet As VbMsgBoxResult, bolAll As Boolean
Const cSpaces = 7
        strTitle = "Удаление событий"
        strMessage = "Чтобы установить для: " & Format(Date1, "dd.mm.yyyy") & " тип даты: " & DateTypes(DateType, 2) & vbCrLf & _
                "необходимо удалить все события, относящиеся к дате, имеющие более высокий приоритет," & vbCrLf & _
                "иначе событие будет создано, но статус даты не изменится." & vbCrLf & _
                "" & vbCrLf & _
                "Да" & vbTab & " - удалить все c более высоким приоритетом;" & vbCrLf & _
                "Нет" & vbTab & " - спросить пользователя об удалении каждого;" & vbCrLf & _
                "Отмена" & vbTab & " - не удалять"
        msgRet = MsgBox(strMessage, vbYesNoCancel Or vbExclamation, strTitle)
        Select Case msgRet
        Case vbYes:                     ' удалить все c более высоким приоритетом без доп вопросов
        Case vbNo:                      ' спросить пользователя об удалении каждого
        Case Else: msgRet = vbCancel    ' не удалять, событие будет создано, но статус даты не изменится
        End Select
    bolAll = msgRet = vbYes
    With rst
    'If Not .EOF Then .MoveFirst
    Do Until .EOF Or msgRet = vbCancel
        If .Fields(c_strDateType) < DateType Then
    ' сравниваем типы если приоритет выше устанавливаемого - предлагаем удалить
    ' вероятно нужны доп критерии например пропуск гос праздников
        If Not bolAll Then
        ' если спрашивать про каждый
        strMessage = "Удалить событие: " & Format(Date1, "dd.mm.yyyy") & " тип даты: " & DateTypes(.Fields(c_strDateType), 2) & vbCrLf & _
                "Описание: " & Nz(.Fields(c_strDateDesc), vbNullString) & vbCrLf & _
                "" & vbCrLf & _
                "Да" & vbTab & " - удалить событие;" & vbCrLf & _
                "Нет" & vbTab & " - не удалять событие;" & vbCrLf & _
                "Отмена" & vbTab & " - прервать"
        msgRet = MsgBox(strMessage, vbYesNoCancel Or vbExclamation, strTitle)
        End If
        ' выбираем действие
        Select Case msgRet
        Case vbYes:  .Delete            ' удаляем событие
        Case vbNo                       ' не удаляем событие
        Case Else:  msgRet = vbCancel   ' выходим,- событие будет создано, но статус даты будет взят по минимальному из оставшихся
        End Select
        End If
        .MoveNext
    Loop
    End With
    End Select
' создаём новое событие указанного типа
HandleNew:   Call p_DatesTableEventEdit(, Date1, DateType, ExtInfo, rst, dbs, wks): Result = DateType
HandleExit:  If bolRst Then rst.Close: Set rst = Nothing
             DateInfoSet = Result: Exit Function
HandleError: Result = eDateTypeUndef: Err.Clear: Resume HandleExit
End Function

Public Function DateInfoAsk(Date1 As Date, _
        Optional ParentControl As Access.Control, _
        Optional ByVal x As Long, Optional ByVal y As Long, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As eDateType
' Выводит контекстное меню с запросом типа укаазанного дня
'-------------------------
' Date1 - дата изменяемого события
' ParentControl - ссылка на контрол под которым д.б. выведено контекстное меню
' X,Y - позиция вывода меню (в экранных координатах)
'-------------------------
Dim Result As eDateType: Result = eDateTypeUndef
    On Error GoTo HandleError
' выводим по левому нижнему углу контрола
Dim tmpPoint As POINT
Dim w As Long, h As Long ', varChild
    If IsMissing(ParentControl) Then
        GetCursorPos tmpPoint: x = tmpPoint.x: y = tmpPoint.y
    ElseIf ParentControl Is Nothing Then
        GetCursorPos tmpPoint: x = tmpPoint.x: y = tmpPoint.y
    Else
        Call AccControlLocation(ParentControl, x, y, , h): y = y + h
    End If
' формируем и выводим контекстное меню (все кроме текущего и выходных)
    ' при формировании контекстного меню  пропускаем:
    '  неопределённый тип даты, служебные типы, а также календарные рабочие и выходные дни - они определяются календарной датой назначать их вручную нет необходимости
Dim strSkip As String: strSkip = Join(Array(eDateTypeUndef, eDateTypeWeekday, eDateTypeSatday, eDateTypeSunday, eDateServWorkCalendar, eDateServOfficial), ";")
Dim EventId As Long, EventIds 'As String
    Result = DateInfoGet(Date1, , EventId, EventIds, dbs:=dbs, wks:=wks)
    Select Case Result
    Case eDateTypeWeekday:   strSkip = eDateTypeWorkday & ";" & strSkip ' нет смысла календарный рабочий день делать рабочим при помощи события
    Case eDateTypeWorkday, eDateTypeHolidayPre, eDateTypeHoliday, _
        eDateTypeNonWorkday: strSkip = Result & ";" & strSkip           ' добавляем в исключения текущий тип даты (кроме пользовательской - их можно добавлять любое количество)
    End Select
' формируем строку для контекстного меню
    ' <имя элемента 1>,<возвращаемое значение 1>,<имя иконки 1>
Dim strList As String: strList = "3;1;6"    ' номера колонок массива DataTypes, необходимых для формирования контекстного меню
    strList = p_DateTypesList(strList, strSkip)     ' получаем необходимые для контекстного меню данные для выбранных типов
    strList = strList & ";"                         ' в конце добавляем разделитель (новую группу)
    strList = strList & ";Редактировать...;-1;#534" ' и доп.действие - редактировать события даты
' открываем контекстное меню (какой тип дня задать дате?)
    Result = eDateTypeUndef
    FormOpenContext strList, ContextVal:=Result, x:=x, y:=y:
' на основании выбора пользователя запрашиваем и заполняем информацию о дате
    Select Case Result
' меняем статус текущей даты (непонятно чего хочу).
    Case eDateTypeHoliday, eDateTypeNonWorkday  ' нерабочий
    Case eDateTypeWorkday, eDateTypeHolidayPre  ' рабочий
    Case eDateTypeUser                          ' пользовательский
    Case -1:
' редактировать события даты
        If Len(EventIds) = 0 Then Result = eDateTypeUndef: GoTo HandleNew  ' если для даты события не заданы - создаём новое событие неопределённого типа (тип события д.б. запрошен в форме)
        If EventId <> EventIds Then EventId = DateInfoList(Date1, dbs, wks) ' если для даты задано несколько событий - открываем список событий даты и предлагаем пользователю выбор записи, которую надлежит редактировать
        Result = DateInfoEdit(EventId, , dbs, wks): GoTo HandleExit          ' меняем выбранное событие и выходим
    Case Else: Err.Raise vbObjectError + 512 'HandleError
    End Select
' создаём новое событие указанного типа
HandleNew:   Result = DateInfoSet(Date1, Result)
HandleExit:  DateInfoAsk = Result:  Exit Function
HandleError: Result = eDateTypeUndef: Err.Clear: Resume HandleExit
End Function

Public Function DateInfoList(Date1 As Date, _
            Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Long
' Открывает список событий для даты. Возвращает код выбранного события
'-------------------------
' Date1   - дата для которой выводим список событий
'-------------------------
Dim Result As eDateType
    On Error GoTo HandleError
    If Not p_DatesTableExists Then GoTo HandleExit
    Result = p_DatesTableEventsList(Date1, , dbs, wks) ' открываем список для указанной даты
HandleExit:  DateInfoList = Result: Exit Function
HandleError: Result = eDateTypeUndef: Err.Clear: Resume HandleExit
End Function

Public Function DateInfoEdit(DateId As Long, _
            Optional ExtInfo As Boolean = False, _
            Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As eDateType
' Редактирует информацию о указанном событии. Возвращает тип созданного события
'-------------------------
' DateId   - код события которое необходимо отредактировать
' ExtInfo   - признак необходимости настройки расширенных параметров события
'-------------------------
Dim Result As eDateType: Result = eDateTypeUndef
    On Error GoTo HandleError
    If Not p_DatesTableExists Then GoTo HandleExit
' редактируем указанное событие
HandleNew:   Call p_DatesTableEventEdit(DateId, , Result, ExtInfo, dbs:=dbs, wks:=wks)
HandleExit:  DateInfoEdit = Result: Exit Function
HandleError: Result = eDateTypeUndef: Err.Clear: Resume HandleExit
End Function

Public Function DateInfoUpdate(Optional DateBeg, Optional DateEnd, _
        Optional AskUpdate As Long = True, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Boolean
' Обновляет данные о праздничных и выходных днях в таблице c_strDatesTable данными производственного календаря из интернет
'-------------------------
' DateBeg   - начало периода за который запрашиваем данные производственного календаря
' DateEnd   - конец периода
' AskTable  - если True при отсутствии будет предлагать создать таблицу дат
    ' 0 - не спрашивать и не обновлять данные,
    '-1 - обновлять имеющиеся данные без запроса
    ' 1 - спрашивать и в зависимости от ответа пользователя обновлять данные или нет
' dbs,wks   - ссылка на базу в которой расположен используемый  календарь
'-------------------------
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
Dim datDateBeg As Date, datDateEnd As Date
Dim bolUpdate As Boolean
' проверяем даты периода
If IsMissing(DateBeg) Then datDateBeg = Now Else If Not IsDate(DateBeg) Then datDateBeg = Now Else datDateBeg = DateBeg
If IsMissing(DateEnd) Then datDateEnd = Now Else If Not IsDate(DateEnd) Then datDateEnd = Now Else datDateEnd = DateEnd
' корректируем даты периода чтобы получить данные сразу за полный год
    datDateBeg = DateSerial(DatePart("yyyy", datDateBeg), 1, 1): datDateEnd = DateSerial(DatePart("yyyy", datDateEnd), 12, 31)

' проверяем наличие таблицы настраиваемых дат и открываем её
Dim rst As DAO.Recordset
    Set rst = p_DatesTableOpen(dbs:=dbs, wks:=wks): Result = Not (rst Is Nothing)
    If Not Result Then Result = p_DatesTableCreate(dbs:=dbs, wks:=wks): If Result Then Set rst = p_DatesTableOpen(dbs:=dbs, wks:=wks): Result = Not (rst Is Nothing)
    If Not Result Then Err.Raise vbObjectError + 512
'Dim lYears As Long: lYears = DateDiff("yyyy", datDateBeg, datDateEnd) + 1 ' получаем количество лет в периоде

    'Application.EnableEvents = True
Dim i As Single, iMax As Single
Dim j As Long, jMax As Long
Dim strTitle As String, strText As String, strMessage As String
    strTitle = "Получение данных из интернет"
    strText = "Идёт загрузка данных"
    strMessage = "Получение данных производственного календаря " & vbCrLf & _
        "за период с " & DateBeg & " по " & DateEnd & " " & vbCrLf & _
        "c сайта: " & с_strLink & ""
    i = CSng(datDateBeg): iMax = CSng(datDateEnd)
    j = 1: jMax = 12
''------------------------------------------
Dim prg As clsProgress: Set prg = New clsProgress
Dim aData(), r As Long, c As Long
    With prg
        .Init pMin:=i, pMax:=iMax, pCaption:=strTitle
        .Detail = strMessage
        .Show
Dim datDateCur As Date ' текущая дата конца периода
Dim lngId As Long, datDateEdit As Date
        Do Until .Progress = .ProgressMax 'And Not .Canceled
            DoEvents
            .Text = strText & String(j, ".")
    ' получаем данные из интернет парсим и если данные подходят - заносим в таблицу
        ' данные читаются в двумерный массив потом переносятся в таблицу
        ' отвязал p_HolidaysFromWeb от рекордсета на будущее, - чтобы дать себе возможность пользоваться этим без Access. Вдруг когда понадобится)
            datDateCur = datDateEnd
        ' проверяем наличие данных за указанный период в таблице
            ' если есть - спрашиваем пользователя - заменить или пропустить
            If p_IsUpdateExists(datDateBeg, datDateCur, lngId, datDateEdit, dbs:=dbs, wks:=wks) Then
            ' если есть данные за период
                If AskUpdate = 1 Then
            ' необходимость обновления определяет пользователь
                    strTitle = "Обновление данных периода"
                    strMessage = "Обновить данные производственного календаря " & vbCrLf & _
                              "за период с " & datDateBeg & " по " & datDateCur & ", " & vbCrLf & _
                              "данными c " & с_strLink & " ?"
                              '"(дата последнего обновления " & datDateEdit & ") "
                    bolUpdate = (MsgBox(strMessage, vbYesNo Or vbExclamation, strTitle) = vbYes)
                Else
            ' необходимость обновления определяется состоянием флага
                    bolUpdate = AskUpdate
                End If
        ' проверяем дату обновления
                'bolUpdate = bolUpdate and (DateDiff("d", datDateEdit, Now()) > cUpdPeriod)
        ' есть данные за период, но решено пропускать имеющиеся данные - переходим к следующему периоду
                If Not bolUpdate Then GoTo HandleNext
            End If
        ' получаем отсутсвующие данные за фрагмент периода из интернет в массив
            aData = p_HolidaysFromWeb(datDateBeg, datDateCur, prg:=prg)
        ' если данные за период не получены или получены некорректные - переходим к следующему периоду
            If Not IsArray(aData) Then GoTo HandleNext
            If UBound(aData, 2) - LBound(aData, 2) < 0 Then GoTo HandleNext
            If bolUpdate Then
        ' если решено обновлять имеющиеся данные - удаляем старые
Dim strWhere As String: strWhere = c_strKey & sqlEqual & lngId & sqlOR & c_strParent & sqlEqual & lngId
                Call dbs.Execute(sqlDeleteAll & c_strDatesTable & sqlWhere & strWhere)
            End If
            With rst
        ' заполняем в таблице служебные данные о периоде за который была запрошена информация
                    .AddNew
                    .Fields(c_strDateType) = eDateServWorkCalendar:
                    .Fields(c_strDateBeg) = datDateBeg: .Fields(c_strDateEnd) = datDateCur:
                    .Fields(c_strDateDesc) = "Производственный календарь (с " & datDateBeg & " по " & datDateCur & ")"
                    .Fields(c_strComment) = "Загружено с: " & с_strLink
                    .Update
                lngId = dbs.OpenRecordset(sqlSelect & sqlIdentity)(0).Value ' ключ служебной записи о периоде за который получены данные
        ' переносим полученные данные из массива в таблицу и добавляем ссылку на ключ записи о периоде за который они получены
                For r = LBound(aData, 2) To UBound(aData, 2)
                    .AddNew:
                    .Fields(c_strParent) = lngId            ' ссылка на родительскую запись ()
                    .Fields(c_strDateType) = aData(1, r)    ' тип события
                    .Fields(c_strDateBeg) = aData(0, r)     ' дата события
                    .Fields(c_strActBegDate) = datDateBeg   ' дата начала актуальности записи события
                    '.Fields(c_strActEndDate) = datDateCur   ' дата конца актуальности записи события
                    .Update
                Next r
            End With
        ' переходим к следующему фрагменту периода
HandleNext: datDateBeg = datDateCur + 1: If datDateBeg >= datDateEnd Then Exit Do
            If .Canceled Then
        ' прерывание процесса
                MsgBox " Процесс загрузки данных прерван." & vbCrLf & _
                "Дата: " & Format$(datDateCur, "dd.mm.yyyy")
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
' Внутренние функции модуля
'=========================

Private Function p_DatesTableEventEdit(Optional EventId, Optional Date1, _
    Optional DateType As eDateType, Optional ExtInfo As Boolean = False, _
    Optional rst As DAO.Recordset, Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Long
' Редактирует / создаёт событие в таблице
'-------------------------
' EventId   - код события  (если задан - редактируется заданный,иначе создаётся новое событие)
' Date1     - дата события (если не задана - д.б. задана EventId)
' DateType  - тип события
' ExtInfo   - признак необходимости настройки расширенных параметров события
' rst/dbs/wks - ссылка на источник данных
'-------------------------
Dim Result As Long ': Result = False
    On Error GoTo HandleError
Dim strWhere As String:
Dim bolRst As Boolean: bolRst = rst Is Nothing
Dim bolNew As Boolean: bolNew = IsMissing(EventId): If bolNew Then strWhere = False Else strWhere = c_strKey & sqlEqual & EventId
    If bolNew Then If IsMissing(Date1) Then Err.Raise vbObjectError + 512
' проверяем тип устанавливаемого события
'    Select Case DateType
'    Case Is < eDateTypeWorkday: DateType = eDateTypeUndef   ' если меньше минимального  - ставим неопределённый тип
'    Case Is > eDateTypeUser:    DateType = eDateTypeUser    ' если больше максимального - ставим пользовательский тип
'    End Select
    ' открываем запрос на таблице событий
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
        '.Fields(c_strParent) = ???            ' ссылка на родительскую запись ()
        If Not ExtInfo Then
    ' простое событие указанного типа + описание (Desc)
        strDesc = InputBox("Введите краткое описание события:", "Описание", strDesc)
        .Fields(c_strDateType) = DateType       ' тип события
        .Fields(c_strDateBeg) = Date1           ' дата события
        If Len(strDesc) > 0 Then .Fields(c_strDateDesc) = strDesc    ' описание события
    Else
    ' создаём настраиваемое событие - открывает форму настройки события, если DateType=Undef в форме также запрашивается тип события
MsgBox "Это заглушка под функцию! " & vbCrLf & "Здесь нужна отдельная форма " & vbCrLf & "редактирования параметров события.", _
        vbOKOnly Or vbExclamation, "Внимание!"
Stop
'Dim tmpForm As Form, FormName As String: FormName = ""
'        FormOpenDrop FormName, NewForm:=tmpForm, FormVal:=Result, X:=X, Y:=Y, Icon:=Icon ', Arrange:= eAlignRightBottom , Visible:= True, FormParent:=ParentControl
'        With tmpForm
'        '.DateType = DateType:.AskType = (DateType = eDateTypeUndef) ' если тип события неопределённый - в форме настройки события надо запросить также и тип
'        '.DateBeg = Date1
'        Do While .Visible: DoEvents: Loop
'        If .ModalResult = vbOK Then
'        'DateType = .DateType: Date1 = .DateBeg
'        'rst.Fields(c_strDateType) = .DateType      ' тип события
'        'rst.Fields(c_strDateBeg) = .DateBeg        ' дата события
'        'rst.Fields(c_strDateDesc) = .DateDesc      ' описание события
'        'If Not IsNull(.DateEnd) Then rst.Fields(c_strDateEnd) = .DateEnd       ' дата окончания длящегося события
'        'If Not IsNull(.OffsetType) Then rst.Fields(c_strOffsetType) = .OffsetType: rst.Fields(c_strOffsetValue) = .OffsetValue   ' для относительных (тип смещения/величина смещения)
'        'If Not IsNull(.PeriodType) Then rst.Fields(c_strPeriodType) = .PeriodType: rst.Fields(c_strPeriodValue) = .PeriodValue   ' для периодических (тип периода/величина периода)
'        '.Fields(c_strDateDesc) = strDesc        ' комментарий к событию
'        End If
'        End With
    End If
        .Fields(c_strActBegDate) = Date1        ' дата начала актуальности записи события
        '.Fields(c_strActEndDate) = Date1       ' дата конца актуальности записи события
        .Fields(c_strEditDate) = Now()          ' изменения записи
        .Update
    End With
    ' пересоздаём временную таблицу
    Call p_TempTableOpen(m_datTempBeg, m_datTempEnd, TempName:=m_strTempName, Requery:=True, dbs:=dbs, wks:=wks)
HandleExit:  If bolRst Then rst.Close: Set rst = Nothing
             p_DatesTableEventEdit = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_DatesTableEventsList(Date1 As Date, _
    Optional AllowEdit As Boolean = False, _
    Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Long
' Выбирает событие относящееся к дате
'-------------------------
' Date1     - дата список событий которой необходимо вывести
' AllowEdit - признак возможности редактировать состав событий даты (добавлять/удалять/помечать как утратившие актуальность)
' dbs/wks - ссылка на источник данных
'-------------------------
Dim Result As Long ': Result = False
    On Error GoTo HandleError
    MsgBox "Это заглушка под функцию! " & vbCrLf & "Здесь нужна отдельная форма " & vbCrLf & "списка событий для даты.", _
            vbOKOnly Or vbExclamation, "Внимание!"
Stop

'Dim EventIds As String:     Result = DateInfoGet(Date1, EventIds:=EventIds, dbs:=dbs, wks:=wks)
'' отктрываем запрос со списком событий относящихся к дате
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
'    ' пересоздаём временную таблицу
'    Call p_TempTableOpen(m_datTempBeg, m_datTempEnd, TempName:=m_strTempName, dbs:=dbs, wks:=wks)
HandleExit:  'Rst.Close: Set Rst = Nothing
             p_DatesTableEventsList = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_CheckForHolidays(ByVal Date1 As Date, ByVal Date2 As Date, _
        Optional Holidays = False, Optional Weekends, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Long
' Проверяет дату по массиву праздников и возвращает число на которое д.б. скорректировано количество рабочих дней в заданном периоде для соответствия Holidays
'-------------------------
' Date1     - дата начала периода
' Date2     - дата окончания периода
' Holidays  - упорядоченный массив назначенных дат (праздничных и нерабочих дней, а также рабочих выходных дней), либо:
'             0 - даты будут получены из таблицы, 1 - из интернета, любое другое - праздники не будут учитываться
'!!! ВНИМАНИЕ !!! если первоначально была задана числовым значением из-за кривости обработки параметра
' повторное использование Holidays в процедуре приводит к вызову со старыми данными массива и ошибкам
' Weekends  - какие дни недели являются выходными и не считаются рабочими (см. p_Weekends). по-умолчанию: суббота, воскресенье
'-------------------------
' массив м.б.:
'   двумерным:  дата/тип (как возвращают p_HolidaysFromTable и p_HolidaysFromWeb)
'   одномерным: дата (тип даты получаем проверяя дату по заданному списку выходных)
'-------------------------
Dim Result As Long
    On Error GoTo HandleError
Dim Temp
Dim aHolidays()
' проверяем переданные даты начала и конца периода
    If Date2 < Date1 Then Temp = Date2: Date2 = Date1: Date1 = Temp
' проверяем переданный Holidays
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
' проверяем дни периода по Holidays
Dim rMax As Long, cMax As Long
Dim d As Integer
    rMax = UBound(aHolidays, 2): If Err = 0 Then cMax = UBound(aHolidays) Else Err.Clear: rMax = UBound(aHolidays)
    If cMax > 0 Then
    ' Holidays двумерный массив  (дата/тип) - проверяем по типу дня
' ! некорректно при нескольких однотипных событиях относящихся к дате
Dim r As Long: For r = 0 To rMax
            Temp = aHolidays(0, r)          ' первая колонка - дата
            Select Case Temp
            Case Date1 To Date2             ' особая дата принадлежит периоду - проверяем
                Select Case aHolidays(1, r) ' вторая колонка - тип
                Case eDateTypeWorkday, eDateTypeHolidayPre: If ISWEEKEND(Temp, Weekends) Then Result = Result + 1     ' рабочий выходной день или предпраздничный (сокращённый) день
                Case eDateTypeHoliday, eDateTypeNonWorkday: If Not ISWEEKEND(Temp, Weekends) Then Result = Result - 1 ' выходной (праздничный) или  нерабочий день
                End Select
            Case Is > Date2: Exit For       ' особая дата больше даты начала периода - выходим
            Case Else                       ' особая дата меньше даты начала периода - пропускаем
            End Select
        Next
    Else
    ' Holidays одномерный массив (только дата) - проверяем по принадлежности к выходным дням
        ' рабочий день указанный в списке считаем выходным
        ' выходной день указанный в массиве считаем рабочим
        For Each Temp In aHolidays
            Select Case Temp
            Case Date1 To Date2: Result = Result + IIf(ISWEEKEND(Temp, Weekends), 1, -1) ' особая дата принадлежит периоду - проверяем
            Case Is > Date2: Exit For       ' особая дата больше даты начала периода - выходим
            Case Else                       ' особая дата меньше даты начала периода - пропускаем
            End Select
        Next
    End If
HandleExit:  p_CheckForHolidays = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Private Function p_HolidaysFromTable(Date1 As Date, Optional Date2, _
    Optional dbs As DAO.Database, Optional wks As DAO.Workspace)
' Возвращает рабочий календарь в виде двумерного массива: дата/тип
'-------------------------
' !!! выходной массив должен быть отсортирован по дате
'-------------------------
Dim Result
On Error GoTo HandleError

' Фильтр по типу записи.    Отбираем только: рабочие выходные, предпраздничные, праздничные и нерабочие
Dim strTypes As String:     strTypes = Join(Array(eDateTypeWorkday, eDateTypeHolidayPre, eDateTypeHoliday, eDateTypeNonWorkday), ",")
' открываем ленту событий для заданного периода на временной таблице - после завершениы работы - удалить
' Возвращаемые поля.        Отбираем только: поле даты события и тип события
Dim strFields As String:    strFields = Join(Array(c_strDateBeg, c_strDateType), ",")
Dim strTable As String
Dim rst As DAO.Recordset:   Set rst = p_TempTableOpen(Date1, Date2, strTypes, strFields, strTable, Unique:=True, dbs:=dbs, wks:=wks)
' читаем в массив
    With rst: .MoveLast: .MoveFirst: Result = .GetRows(.RecordCount): End With
'' удаляем временную таблицу
'    DropTempCalendar 'rst.Close: Set rst = Nothing: Call p_TableDrop(strTable)
HandleExit:  p_HolidaysFromTable = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Private Function p_HolidaysFromWeb(ByRef DateBeg As Date, Optional DateEnd, _
        Optional Weekends, Optional prg)  ' prg As clsProgress)
' Запрашивает из интернета данные, парсит их и возвращает массив c результатом
'-------------------------
' !!! парсит только основной календарь. не читает доп нерабочие дни установленные указами, описанные в тексте, но не отмеченные в самом календаре
'-------------------------
Const cstrDivTagName = "div"
'Const cstrDataClassName = "calendar_full"
Const c_strWorkDayClassName = "calendar_day  "
Const cstrHolyPreClassName = c_strWorkDayClassName & "calendar_day__holiday_pre"
Const c_strHolidayClassName = c_strWorkDayClassName & "calendar_day__holiday"
Const cstrDayOffClassName = c_strWorkDayClassName & "calendar_day__dayoff"
'Const cstrHolidays = "Нерабочие праздничные дни "
'Const cstrTransfers = "Перенос выходных дней в"
'Const cstrTransfers = "Законами и другими НПА органов госвласти субъектов РФ могут быть установлены <b>дополнительные нерабочие праздничные дни</b> (ст. 6 ТК РФ).&nbsp;"

Const cstrDateAttrName = "data-day"
Const cMaxErr = 10
Const cstrDelim = " " 'Chr(32)

Dim Result As Boolean ':Result = False
On Error GoTo HandleError
Dim bolProgress As Boolean: bolProgress = Not IsMissing(prg) 'TypeOf prg Is clsProgress
    If IsMissing(DateEnd) Then DateEnd = DateBeg
Dim lErrCount As Long
' запрашиваем исходные данные из интернет
Dim Temp As Date ': Temp = DateBeg
Dim strYear As String: strYear = Format(DateBeg, "yyyy")
Dim strURL As String:  strURL = с_strLink & strYear & "/"
Static HTML As Object: Set HTML = CreateObject("htmlFile")          'Dim HTML As New MSHTML.HTMLDocument
Static HXML As Object: Set HXML = CreateObject("MSXML2.XMLHTTP")    '("Msxml2.ServerXMLHTTP")
' формируем текст запроса
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

' читаем загруженный календарь и переносим даты за указанный период в наш
Dim aData(), i As Long
Dim enmDayType As eDateType
Dim Itm As Object
'Stop
    For Each Itm In HTML.getElementsByTagName(cstrDivTagName)
    On Error GoTo HandleError
    DoEvents
        If bolProgress Then If prg.Canceled Then Err.Raise vbObjectError + 512
        With Itm
            Select Case .ClassName      ' проверяем имя класса
            Case c_strWorkDayClassName: enmDayType = eDateTypeWorkday           ' рабочий день
            Case cstrDayOffClassName:   enmDayType = eDateTypeSunday            ' выходной
            Case cstrHolyPreClassName:  enmDayType = eDateTypeHolidayPre        ' предпраздничный
            Case c_strHolidayClassName: enmDayType = eDateTypeHoliday           ' праздничный
            Case Else:                  GoTo HandleNext                         ' другое
            End Select
            Temp = CDate(Replace(.Attributes(cstrDateAttrName).Value, "_", ".")): If Temp < DateBeg Then GoTo HandleNext
        End With
' проверяем календарь: если выходной день в производственном календаре отмечен как выходной - не надо ничего добавлять
        If ISWEEKEND(Temp, Weekends) Then
            If enmDayType = eDateTypeSunday Then GoTo HandleNext            ' выходной помеченный как выходной
        Else
            If enmDayType = eDateTypeWorkday Then GoTo HandleNext           ' рабочий помеченный как рабочий
        End If
' заносим данные в таблицу
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
    MsgBox "Необходимо в IE разрешить использовать TLS1.2" & vbCrLf _
    & "", vbOKOnly Or vbCritical, "Ошибка 0x800C0008 (INET_E_DOWNLOAD_FAILURE)"
    'Case -2147012721 ' A security error occurred
    'Case -2147220991 ' Automation error Событие не смогло вызвать ни одного из абонентов
    '                 ' в HTML закончились фрагменты данных для обработки
    Case vbObjectError + 513 ' Error 404
Debug.Print "Can't get data for year " & strYear
    End Select
    Result = False: Err.Clear: Resume HandleExit
End Function

Private Function p_DatesTableExists(Optional bTest As Boolean = False, Optional AskTable As Boolean = True, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Boolean
' Проверяет наличие таблицы только при первом запуске
'-------------------------
' bTest     = True для повторной проверки
' AskTable  - если True при отсутствии будет предлагать создать таблицу дат
' dbs,wks   - ссылка на базу в которой расположен используемый  календарь
'-------------------------
Static bolExists As Boolean ' признак наличия таблицы
Static bolInit As Boolean   ' признак того что проверка уже проводилась
    If bTest Then GoTo HandleTest
    If bolInit Then p_DatesTableExists = bolExists: Exit Function
HandleTest:
    On Error Resume Next
Dim Result As Boolean
' проверяем наличие таблицы
    Result = Not p_DatesTableOpen(bTest:=True, dbs:=dbs, wks:=wks) Is Nothing
' можно проверить ещё корректность (наличие основных полей), но - не будем
    If Not Result Then
        If AskTable Then
' если отсутствует - спрашиваем и создаём новую
Dim strText As String, strTitle As String
            strTitle = "Отсутствует таблица"
            strText = "Отсутствует таблица для хранения настраиваемых дат." & vbCrLf & "Создать таблицу """ & c_strDatesTable & """?"
            If (MsgBox(strText, vbYesNo Or vbExclamation, strTitle) = vbYes) Then Result = p_DatesTableCreate(dbs:=dbs, wks:=wks)
        End If
    End If
    bolExists = Result
    bolInit = True: p_DatesTableExists = bolExists
End Function
Private Function p_IsUpdateExists(Date1 As Date, Date2 As Date, _
    Optional ID As Long, Optional RecDate As Date, _
    Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Long
' Проверяет наличие в таблице загруженных данных за указанный период возвращает код первой записи обновления соответствующих периоду
'-------------------------
' Date1, Date2 - дата начала/конца периода
' ID, RecDate - код/дата найденой записи
' dbs,wks   - ссылка на базу в которой расположен используемый  календарь
'-------------------------
' проверяет только служебные записи обновлений производственного календаря,
' нужна для проверки актуальности имеющихся в таблице данных при обновлении из интернет
'-------------------------
Dim Result As Boolean:
    If Not p_DatesTableExists(AskTable:=True, dbs:=dbs, wks:=wks) Then Err.Raise vbObjectError + 512
'Dim strFields As String:    strFields = Join(Array(c_strKey, c_strEditDate), ",")
Dim strOrder As String:     strOrder = c_strDateBeg
Dim strTypes As String:     strTypes = eDateServWorkCalendar                                                ' отбираем записи обновлений производственного календаря
Dim strWhere As String:     strWhere = c_strDateType & sqlIn & "(" & strTypes & ")"                         ' фильтр по типу записи
    strWhere = strWhere & sqlAnd & p_DateToSQL(Date1) & sqlBetween & c_strDateBeg & sqlAnd & c_strDateEnd   ' добавляем фильтр по дате
Dim rst As DAO.Recordset:   Set rst = p_DatesTableOpen(strWhere, strOrder, dbs:=dbs, wks:=wks)
' проверяем наличие загруженных данных за период
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
' Cткрывает таблицу настраиваемых дат
'-------------------------
    On Error GoTo HandleError
    If wks Is Nothing Then Set wks = DBEngine.Workspaces(0): Set dbs = CurrentDb                                         ' по-умолчанию текущая база и рабочее пространство
    If dbs Is Nothing Then If wks.Databases.Count > 0 Then Set dbs = wks.Databases(0) Else Err.Raise vbObjectError + 512 ' если не задана база, но задано рабочее пространство - берём первую в рабочем пространстве, иначе - ошибка
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
' Создаёт таблицу настраиваемых дат
'-------------------------
' bTemp     = False для создания основной таблицы календаря, (Id делается ключевым полем)
'           = True для создания временной таблицы календаря
' TableName - имя создаваемой таблицы. на выходе переменная будет содержать имя созданной таблицы
' dbs,wks   - ссылка на базу в которой расположен используемый  календарь
'-------------------------
Dim Result As Boolean
Dim bolTransOpen As Boolean
    On Error GoTo HandleError
    If wks Is Nothing Then Set wks = DBEngine.Workspaces(0): Set dbs = CurrentDb                                         ' по-умолчанию текущая база и рабочее пространство
    If dbs Is Nothing Then If wks.Databases.Count > 0 Then Set dbs = wks.Databases(0) Else Err.Raise vbObjectError + 512 ' если не задана база, но задано рабочее пространство - берём первую в рабочем пространстве, иначе - ошибка
Dim tdf  As DAO.TableDef, fld As DAO.Field, idx As DAO.Index
' получаем имя создаваемой таблицы
    If Len(TableName) = 0 Then TableName = c_strDatesTable: If bTemp Then Mid(TableName, 1, 3) = c_strTmpTablePref
HandleCreate:
    If bTemp Then
    ' для временной - проверяем наличие таблицы TableName и при наличии удаляем её
        With wks
            If p_IsTableExists(TableName, dbs:=CurrentDb) Then .BeginTrans: bolTransOpen = True: dbs.Execute sqlDropTable & "[" & TableName & "]": .CommitTrans: bolTransOpen = False
        End With
    End If
' создаём таблицу
    Set tdf = dbs.CreateTableDef(TableName)
    With tdf
' создаём поля
        Set fld = .CreateField(c_strKey, dbLong):           .Fields.Append fld  ' ключ
        If Not bTemp Then fld.Attributes = dbAutoIncrField                      ' автонумерация ключевого поля основной таблицы
        Set fld = .CreateField(c_strDateType, dbLong):      .Fields.Append fld  ' тип даты
        Set fld = .CreateField(c_strDateBeg, dbDate):       .Fields.Append fld  ' дата начала события
        Set fld = .CreateField(c_strDateDesc, dbText, 100): .Fields.Append fld  ' описание даты
        If Not bTemp Then
    ' пропускаем поля которые нужны только для основной таблицы:
        Set fld = .CreateField(c_strDateEnd, dbDate):       .Fields.Append fld  ' дата окончания события
        Set fld = .CreateField(c_strOffsetType, dbText, 4): .Fields.Append fld  ' тип смещения
        Set fld = .CreateField(c_strOffsetValue, dbLong):   .Fields.Append fld  ' величина смещения
        Set fld = .CreateField(c_strPeriodType, dbText, 4): .Fields.Append fld  ' тип периода
        Set fld = .CreateField(c_strPeriodValue, dbLong):   .Fields.Append fld  ' величина периода
        Set fld = .CreateField(c_strComment, dbMemo):       .Fields.Append fld  ' комментарий к дате
        Set fld = .CreateField(c_strParent, dbLong):        .Fields.Append fld  ' код родителя
        Set fld = .CreateField(c_strActBegDate, dbDate):    .Fields.Append fld  ' дата начала актуальности записи
        Set fld = .CreateField(c_strActEndDate, dbDate):    .Fields.Append fld  ' дата окончания актуальности записи
        Set fld = .CreateField(c_strEditDate, dbDate):      .Fields.Append fld  ' дата изменения записи
        End If
' создаём индексы
        Set idx = .CreateIndex("PrimaryKey"): With idx                  ' основной индекс по ключу
            .Fields.Append .CreateField(c_strKey)
        If Not bTemp Then .Primary = True: .Unique = True               ' делаем уникальным основной индекс основной таблицы
        End With: .Indexes.Append idx
        Set idx = .CreateIndex("DayTypeKey"): With idx                  ' дополнительный по дате/типу
            .Fields.Append .CreateField(c_strDateBeg)
            If Not bTemp Then .Fields.Append .CreateField(c_strDateEnd)
            .Fields.Append .CreateField(c_strDateType)
            If Not bTemp Then .Fields.Append .CreateField(c_strOffsetType)
            If Not bTemp Then .Fields.Append .CreateField(c_strPeriodType)
            '.IgnoreNulls = True
        End With: .Indexes.Append idx
    End With
    dbs.TableDefs.Append tdf
    If bTemp Then GoTo HandleTest ' для временной таблицы не будем созадавать комбобоксы и прочие фитюльки
' настраиваем дополнительные свойства полей
Dim strList As String
    Set fld = tdf.Fields(c_strEditDate)                                 ' дата создания записи (значение по умолчанию)
        Call PropertySet("DefaultValue", "=Date()", fld)
    Set fld = tdf.Fields(c_strDateBeg)                                  ' дата события (значение по умолчанию)
        Call PropertySet("DefaultValue", "=Date()", fld)
    Set fld = tdf.Fields(c_strActBegDate)                               ' дата начала действительности события
        Call PropertySet("DefaultValue", "=Date()", fld)
    Set fld = tdf.Fields(c_strDateType)                                 ' тип даты (настройка списка); значение по умолчанию
        strList = p_DateTypesList("1;3", Join(Array(eDateTypeWeekday, eDateTypeSatday, eDateTypeSunday), ";"))
        Call PropertySet("DisplayControl", acComboBox, fld, dbInteger)
        Call PropertySet("RowSourceType", "Value List", fld)
        Call PropertySet("RowSource", strList, fld)
        Call PropertySet("ColumnCount", 2, fld, dbInteger)
        Call PropertySet("ColumnWidths", 0, fld)
        Call PropertySet("DefaultValue", eDateTypeUser, fld)            ' тип даты (значение по-умолчанию)
        Call PropertySet("TextAlign", 1, fld, dbByte)
    Set fld = tdf.Fields(c_strOffsetType)                               ' тип смещения (настройка списка)
        strList = p_DateIntervalList
        Call PropertySet("DisplayControl", acComboBox, fld, dbInteger)
        Call PropertySet("RowSourceType", "Value List", fld)
        Call PropertySet("RowSource", strList, fld)
        Call PropertySet("ColumnCount", 2, fld, dbInteger)
        Call PropertySet("ColumnWidths", 0, fld)
    Set fld = tdf.Fields(c_strPeriodType)                               ' тип периода (настройка списка)
        strList = p_DateIntervalList
        Call PropertySet("DisplayControl", acComboBox, fld, dbInteger)
        Call PropertySet("RowSourceType", "Value List", fld)
        Call PropertySet("RowSource", strList, fld)
        Call PropertySet("ColumnCount", 2, fld, dbInteger)
        Call PropertySet("ColumnWidths", 0, fld)
' настраиваем дополнительные свойства таблицы
        Call PropertySet("SubdatasheetName", "Table." & c_strDatesTable, tdf)                                       ' подтаблица с группировкой по родителю
        Call PropertySet("LinkMasterFields", c_strKey, tdf): Call PropertySet("LinkChildFields", c_strParent, tdf)  ' связь по PARENT=ID
        Call PropertySet("Filter", c_strParent & sqlIsNull, tdf): Call PropertySet("FilterOnLoad", True, tdf)       ' выводить только родительские записи
        'Call PropertySet("OrderByOnLoad", False, tdf): Call PropertySet("OrderByOn", False, tdf)
' обновляем список таблиц
    dbs.TableDefs.Refresh: Application.RefreshDatabaseWindow
HandleTest:  Result = p_DatesTableExists(True, dbs:=dbs, wks:=wks)
HandleExit:  p_DatesTableCreate = Result: Exit Function
HandleError: If bolTransOpen Then wks.Rollback: bolTransOpen = False     ' откатываем транзакцию
    Select Case Err.Number
    Case 3211: 'Stop ' не смог удалить локальную таблицу - заблокирована
    If bTemp Then TableName = c_strTmpTablePref & GenPassword(8): Err.Clear: Resume HandleCreate    ' если создаём временную - попытка создать с др именем
    Case 3734: Stop ' база данных была приведена пользователем в состояние, препятствующее ее открытию или блокировке
    End Select
    Result = False: Err.Clear: Resume HandleExit
End Function

Private Function p_Weekends(Optional Weekends)
' Возвращает массив дней недели которые являются выходными и не считаются рабочими.
'-------------------------
' Weekends  - какие дни недели являются выходными и не считаются рабочими. по-умолчанию: суббота, воскресенье
'   числовое значение 1-7, 11-17
'   или строка вида: 0000011, где 0- рабочий день, 1-выходной (начиная с понедельника)
'-------------------------
' для получения значения массива вызывать: Weekends()(i), иначе будет рассматривать как параметр инициализации массива
'-------------------------
On Error Resume Next
Static aData()
Dim i As Long: i = LBound(aData)
    If Err = 0 Then If IsMissing(Weekends) Then p_Weekends = aData: Exit Function
    Err.Clear: If IsMissing(Weekends) Then Weekends = 1
On Error GoTo HandleError
    If Len(Weekends) = 7 Then
' Weekends - строковое значение дней недели.
    ' Включают семь знаков, каждый из которых обозначает день недели (начиная с понедельника).
    ' Значение 1 представляет нерабочие дни, а 0 — рабочие дни. В строке допустимо использовать только знаки 1 и 0. Строка 1111111 недопустима.
    ' Например, 0000011 означает, что выходными днями являются суббота и воскресенье
Dim j As Long: For j = 1 To 7
        Select Case Mid(Weekends, j, 1)
        Case 1: ReDim Preserve aData(0 To i)
            Select Case j
            Case 1: aData(i) = vbMonday             ' понедельник
            Case 2: aData(i) = vbTuesday            ' вторник
            Case 3: aData(i) = vbWednesday          ' среда
            Case 4: aData(i) = vbThursday           ' четверг
            Case 5: aData(i) = vbFriday             ' пятница
            Case 6: aData(i) = vbSaturday           ' суббота
            Case 7: aData(i) = vbSunday             ' воскресенье
            Case Else: Err.Raise vbObjectError + 512
            End Select
            i = i + 1: If i >= 7 Then Err.Raise vbObjectError + 512
        Case 0:
        Case Else: Err.Raise vbObjectError + 512
        End Select
        Next j: GoTo HandleExit
    End If
HandleSelect: Select Case Weekends
' Weekends - числовое значение
        Case 1:  aData = Array(vbSaturday, vbSunday)    ' суббота, воскресенье
        Case 2:  aData = Array(vbSunday, vbMonday)      ' воскресенье, понедельник
        Case 3:  aData = Array(vbMonday, vbTuesday)     ' понедельник, вторник
        Case 4:  aData = Array(vbTuesday, vbWednesday)  ' вторник, среда
        Case 5:  aData = Array(vbWednesday, vbThursday) ' среда, четверг
        Case 6:  aData = Array(vbThursday, vbFriday)    ' четверг , пятница
        Case 7:  aData = Array(vbFriday, vbSaturday)    ' пятница , суббота
        Case 11: aData = Array(vbSunday)                ' только воскресенье
        Case 12: aData = Array(vbMonday)                ' только понедельник
        Case 13: aData = Array(vbTuesday)               ' только вторник
        Case 14: aData = Array(vbWednesday)             ' только среда
        Case 15: aData = Array(vbThursday)              ' только четверг
        Case 16: aData = Array(vbFriday)                ' только пятница
        Case 17: aData = Array(vbSaturday)              ' только суббота
        Case Else:
        End Select
HandleExit:  p_Weekends = aData: Exit Function
HandleError: Err.Clear: Weekends = 1: Resume HandleSelect
End Function

Private Function p_WeekendsList(Optional Weekends, Optional Delim As String = ",") As String
' Возвращает выходные дни недели ввиде списка (для SQL: IN (..))
'-------------------------
    p_WeekendsList = Join(p_Weekends(Weekends), Delim)
End Function

Private Function p_DateTypesList(Optional Columns As String, Optional Skip As String, Optional Delim = ";") As String
' Возвращает список данных справочника типов дат
'-------------------------
On Error Resume Next
Dim i As Long, j As Long, iStep As Long
Dim aSkip() As String:      aSkip = Split(Skip, Delim)
Dim aColumns() As String:   aColumns = Split(Columns, Delim)
Dim aDateTypes():           aDateTypes = DateTypes(iStep:=iStep)
Dim strList As String ': strList = Join(DateTypes, ";")
Dim Value
    For i = LBound(aDateTypes) To UBound(aDateTypes) Step iStep
    ' пробегаем по строкам справочника, в каждой iStep колонок
        For j = LBound(aSkip) To UBound(aSkip)
        ' проверяем по списку пропуска
            If aDateTypes(i) = aSkip(j) Then GoTo HandleNext
        Next j
        For j = LBound(aColumns) To UBound(aColumns)
        ' формируем строку списка
            'Value=aDateTypes(i + aColumns(j) - 1)
            strList = strList & Delim & aDateTypes(i + aColumns(j) - 1) 'Value
        Next j
HandleNext:  Err.Clear: Next i
    If Left$(strList, Len(Delim)) = Delim Then strList = Mid(strList, Len(Delim) + 1)
HandleExit: p_DateTypesList = strList
End Function

Private Function p_DateIntervalList(Optional Columns As String, Optional Skip As String, Optional Delim = ";") As String
' Возвращает список верменных интервалов
'-------------------------
Static strData As String
    If Len(strData) = 0 Then
Dim arrData(): arrData = Array( _
            "", "<none>", _
            "yyyy", "год", _
            "q", "квартал", _
            "m", "месяц", _
            "w", "неделя", _
            "d", "день", _
            c_strWorkdayLiteral, "день (рабочий)", _
            c_strMondayLiteral, "понедельник (день недели)", _
            c_strTuesdayLiteral, "вторник (день недели)", _
            c_strWednesdayLiteral, "среда (день недели)", _
            c_strThursdayLiteral, "четверг (день недели)", _
            c_strFridayLiteral, "пятница (день недели)", _
            c_strSaturdayLiteral, "суббота (день недели)", _
            c_strSundayLiteral, "воскресенье (день недели)")
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
' возвращет рекордсет ленты событий за указанный период, открытый на созданной временной таблице
'-------------------------
' Date1     - начальная дата периода
' Date2     - конечная дата периода (д.б.>=Date1)
' DateTypes - список допустимых типов записи
' Fields    - список возвращаемых запросом полей
' TempName  - имя временной таблицы содержащей запрошенные данные
' Unique    - если True будет возвращён запрос содержащий только уникальные записи по дате с максимальным приоритетом (те что будут учитываться при подсчётах)
' Requery   - если True временная таблица будет пересоздана независимо от её наличия
' dbs,wks   - ссылки на базу данных из которой осуществляется импорт и рабочее пространство
'-------------------------
' ToDo: Реально создавать её надо только когда Date1<=m_datTempBeg OR Date2>=m_datTempEnd
' в остальных случаях она у нас как-бы есть - насколько корректные данные она содержит проверять не будем
' лучше временной таблицы не придумал запросы получались слишком сложные (
' да и здесь выхоодной запрос надо бы попроще
Dim Result As Boolean ': Result = False
    On Error GoTo HandleError
    If IsMissing(Date2) Then Date2 = Date1
Const cTmp = "[" & c_strTmpTablPref & "]"               ' имя подзапроса
Dim sqlDate1 As String: sqlDate1 = p_DateToSQL(Date1)
Dim sqlDate2 As String: sqlDate2 = p_DateToSQL(Date2)
Dim strWhereDate As String: strWhereDate = "(" & c_strDateBeg & sqlBetween & sqlDate1 & sqlAnd & sqlDate2 & ")" ' фильтр по дате
    If Not Requery Then If Date1 >= m_datTempBeg And Date2 <= m_datTempEnd Then GoTo HandleResult
Dim bolTransOpen As Boolean
    If wks Is Nothing Then Set wks = DBEngine.Workspaces(0): Set dbs = CurrentDb                                         ' по-умолчанию текущая база и рабочее пространство
    If dbs Is Nothing Then If wks.Databases.Count > 0 Then Set dbs = wks.Databases(0) Else Err.Raise vbObjectError + 512 ' если не задана база, но задано рабочее пространство - берём первую в рабочем пространстве, иначе - ошибка
Dim strSQL As String, strWhere As String, strOrder As String, strFields As String
' список полей для временной таблицы
    '!!! временная таблица содержит только 4 основных поля см. p_DatesTableCreate
Dim arrFields: arrFields = Array(c_strDateBeg, c_strDateType, c_strKey, c_strDateDesc) ', c_strComment)
    strFields = Join(arrFields, ",")
' создаём временную таблицу (локально)
    Result = p_DatesTableCreate(True, TempName): If Not Result Then Err.Raise vbObjectError + 512
    ' собираем фильтр
    ' фильтр по типу события
    If Len(DateTypes) = 0 Then strWhere = "<100" Else strWhere = sqlIn & "(" & DateTypes & ")"
    strWhere = "(" & c_strDateType & strWhere & ")"
    ' фильтр по актуальности записи
    strWhere = strWhere & _
        sqlAnd & "(IIf([" & c_strActBegDate & "]" & sqlIsNull & ",[" & c_strDateBeg & "],[" & c_strActBegDate & "])<=" & sqlDate2 & ")" & _
        sqlAnd & "(IIf([" & c_strActEndDate & "]" & sqlIsNull & "," & sqlDate2 & ",[" & c_strActEndDate & "])>=" & sqlDate1 & ")"
Dim strWhereTemp As String  ' фильтр по сложности события
    strWhereTemp = "((" & c_strDateEnd & sqlIsNull & ")" & _
        sqlAnd & "(" & c_strOffsetType & sqlIsNull & ")" & _
        sqlAnd & "(" & c_strPeriodType & sqlIsNull & "))"
' открываем список всех актуальных на периоде "сложных" событий (длящиеся, относительные или периодические)
    ' оставляем фильтр типа и фильтр актуальности, опускаем фильтр даты
    ' переворачиваем фильтр по сложности события
    ' формируем запрос на выборку
    strSQL = sqlSelectAll & c_strDatesTable & sqlWhere & strWhere & sqlAnd & sqlNot & strWhereTemp & ";"
Dim rstSrc As DAO.Recordset, rst As DAO.Recordset
    ' открываем запрос источник со списком сложных событий
    Set rstSrc = dbs.OpenRecordset(strSQL)
    With rstSrc
        .MoveFirst: If .BOF And .EOF Then GoTo HandleExit
    ' открываем запрос на созданной временной таблице куда будут добавляться записи
    Set rst = CurrentDb.OpenRecordset(TempName, dbOpenDynaset)
Dim bDate As Date   ' начальная дата соответствующая периоду
Dim eDate As Date   ' конечноая дата соответствующая периоду
Dim lCount As Long  ' количество повторов даты в периоде
Dim lLen As Double  ' дительность события
Dim iDate As Date   ' текущая дата начала события
Dim i As Long       ' повторы периодических событий на периоде
Dim fld
        Do
'    ' пробегаем отобранные события и для каждого добавляем в ленту необходимое количество повторов
            i = 0: lLen = 0: lCount = 0
            bDate = .Fields(c_strDateBeg)
'Debug.Print bDate, .Fields("ID"), .Fields("DateDesc"), .Fields("Comment")
        ' периодическое событие
            If Not IsNull(.Fields(c_strPeriodType)) Then
            ' получаем количество полных периодов до начальной даты искомого периода
                lCount = DateDiffEx(.Fields(c_strPeriodType), bDate, Date1, wks:=wks, dbs:=dbs)
            ' получаем начальную дату периодического события
                bDate = DateAddEx(.Fields(c_strPeriodType), lCount, bDate, wks:=wks, dbs:=dbs)
            ' получаем количество повторов события между начальной и конечной датой искомого периода
                lCount = DateDiffEx(.Fields(c_strPeriodType), bDate, Date2, wks:=wks, dbs:=dbs)
            End If
        ' длящееся событие - получаем длительность события
            If Not IsNull(.Fields(c_strDateEnd)) Then lLen = .Fields(c_strDateEnd) - .Fields(c_strDateBeg) ': Stop
            iDate = bDate
        ' относительное событие - добавляем смещение к начальной дате
            If Not IsNull(.Fields(c_strOffsetType)) Then iDate = DateAddEx(.Fields(c_strOffsetType), .Fields(c_strOffsetValue), bDate)
            eDate = iDate + lLen
            Do While iDate <= Date2
            ' проверяем период
                ' если попадаем в период - добавляем событие, иначе - следующая дата
'Debug.Print iDate, .Fields("ID"), .Fields("DateDesc")
'Stop
                If eDate < Date1 Then GoTo HandleNext
                If iDate < Date1 Then GoTo HandleNext
            ' заполняем поля
                rst.AddNew
                For Each fld In arrFields
                    rst.Fields(fld) = IIf(fld = c_strDateBeg, iDate, .Fields(fld))
                Next fld
                rst.Update
HandleNext: ' переходим к следующему элементу ленты для данного события
                ' длящееся событие
                If (eDate - iDate) >= 1 Then
                    iDate = Int(iDate + 1)  ' отбрасываем часы/минуты для дней кроме первого и последнего в длящемся событии
                ElseIf (eDate - iDate) > 0 Then
                    iDate = eDate           ' конечная дата длящегося события со временем завершения
                ElseIf (iDate < Date2) And (lCount > 0) Then
                ' повторяющееся событие
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
' добавляем в таблицу простые события (кроме длящихся, относительных или периодических), попадающие в период
    ' фильтр по дате события (принадлежит периоду)
    ' и исключаем пустые повторы (простые события уже заданных для даты типов если описание для них пусто)
    strWhere = strWhere & sqlAnd & "(" & c_strKey & sqlNot & sqlIn & "(" & _
        sqlSelect & c_strDatesTable & "." & c_strKey & sqlFrom & "[" & CurrentDb.Name & "].[" & TempName & "]" & sqlAs & cTmp & _
        sqlInner & sqlJoin & c_strDatesTable & sqlOn & _
        "(" & cTmp & "." & c_strDateBeg & sqlEqual & "[" & c_strDatesTable & "]." & c_strDateBeg & ")" & _
        sqlAnd & "(" & cTmp & "." & c_strDateType & sqlEqual & "[" & c_strDatesTable & "]." & c_strDateType & ")" & _
        sqlWhere & "[" & c_strDatesTable & "]![" & c_strDateDesc & "]" & sqlIsNull & "))"
' формируем запрос на добавление
    strSQL = sqlInsert & sqlInto & "[" & CurrentDb.Name & "].[" & TempName & "] (" & strFields & ") " & _
             sqlSelect & strFields & sqlFrom & c_strDatesTable & _
             sqlWhere & strWhere & sqlAnd & strWhereTemp & sqlAnd & strWhereDate & ";"
    ' создаём временную таблицу в локальной базе и заполняем данные
    With wks: .BeginTrans: bolTransOpen = True: dbs.Execute strSQL: .CommitTrans: bolTransOpen = False: End With
    
    rstSrc.Close: Set rstSrc = Nothing
    rst.Close
'' обновляем список таблиц
'    ''CurrentDb.TableDefs.Refresh
'    'Application.RefreshDatabaseWindow
    m_datTempBeg = Date1: m_datTempEnd = Date2: m_strTempName = TempName

HandleResult:
' сортируем рекордсет по дате и возвращаем
    If Len(Fields) = 0 Then strFields = sqlAll Else strFields = Fields
    strOrder = Join(Array(c_strDateBeg, c_strDateType), ",")
    ' фильтр для всех записей периода (для календаря)
    strWhere = strWhereDate
    ' добавляем фильтр для отбора уникальных записей (для расчётов)
    If Unique Then strWhere = strWhere & sqlAnd & "(" & c_strKey & sqlEqual & "(" & _
            sqlSelect & sqlTop1 & cTmp & ".ID" & sqlFrom & "[" & m_strTempName & "]" & sqlAs & cTmp & _
            sqlWhere & cTmp & ".[" & c_strDateBeg & "]=[" & m_strTempName & "].[" & c_strDateBeg & "]" & _
            sqlOrder & cTmp & ".[" & c_strDateType & "]," & cTmp & ".[" & c_strKey & "])" & ")"
    strSQL = sqlSelect & strFields & sqlFrom & "[" & m_strTempName & "]" & sqlWhere & strWhere & sqlOrder & strOrder & ";"
    Set p_TempTableOpen = CurrentDb.OpenRecordset(strSQL) ': rst.MoveFirst
HandleExit:  Exit Function
HandleError: If bolTransOpen Then wks.Rollback: bolTransOpen = False     ' откатываем транзакцию
    Select Case Err.Number
    Case 3211: Stop ' не смог удалить локальную таблицу - заблокирована
    Case 3734: Stop ' база данных была приведена пользователем в состояние, препятствующее ее открытию или блокировке
    End Select
    Result = False: Err.Clear: Resume HandleExit
End Function

Private Function p_TableDrop(Source, _
        Optional dbs As DAO.Database, Optional wks As DAO.Workspace) As Boolean
' удаляет  таблицу
'-------------------------
' Source    - имя временной таблицы или объект DAO.Recordset, открытый на ней
' dbs,wks   - ссылки на базу данных из которой осуществляется импорт и рабочее пространство
'-------------------------
Dim strSQL As String, strSource As String, strTarget As String, strFields As String
Dim bolTransOpen As Boolean
    If wks Is Nothing Then Set wks = DBEngine.Workspaces(0): Set dbs = CurrentDb                                         ' по-умолчанию текущая база и рабочее пространство
    If dbs Is Nothing Then If wks.Databases.Count > 0 Then Set dbs = wks.Databases(0) Else Err.Raise vbObjectError + 512 ' если не задана база, но задано рабочее пространство - берём первую в рабочем пространстве, иначе - ошибка
Dim Result As Boolean ': Result = False
    On Error GoTo HandleError
    If VarType(Source) = vbString Then
    ' источник - SQL инструкция или имя таблицы
        Source = Trim$(Source): If Len(Source) = 0 Then Err.Raise vbObjectError + 512
        strSource = Source
    ElseIf TypeOf Source Is DAO.Recordset Then
    ' источник - DAO.Recordset - берем имя источника
        strSource = Source.Name
        Source.Close: Set Source = Nothing
    Else
    ' источник не известен
        Err.Raise vbObjectError + 512
    End If
    With wks
        If p_IsTableExists(strSource, dbs, wks) Then .BeginTrans: dbs.Execute sqlDropTable & "[" & strSource & "]": .CommitTrans ': bolTransOpen = False
    End With
    Result = True
HandleExit:  p_TableDrop = Result: Exit Function
HandleError: If bolTransOpen Then wks.Rollback: bolTransOpen = False     ' откатываем транзакцию
    Select Case Err.Number
    Case 3211: Stop ' не смог удалить локальную таблицу - заблокирована
    Case 3734: Stop ' база данных была приведена пользователем в состояние, препятствующее ее открытию или блокировке
    End Select
    Result = False: Err.Clear: Resume HandleExit
End Function

Private Function p_IsTableExists(ByVal TableName As String, _
    Optional dbs As DAO.Database, Optional wks As DAO.Workspace _
    ) As Boolean
' Возвращает значения True, если Есть таблица с таким именем.
'-------------------------
' dbs,wks   - ссылки на базу данных из которой осуществляется импорт и рабочее пространство
'-------------------------
Dim Result As Boolean ': Result = False
    On Error GoTo HandleError
Dim strSQL As String
' ---
    If wks Is Nothing Then Set wks = DBEngine.Workspaces(0): Set dbs = CurrentDb                                         ' по-умолчанию текущая база и рабочее пространство
    If dbs Is Nothing Then If wks.Databases.Count > 0 Then Set dbs = wks.Databases(0) Else Err.Raise vbObjectError + 512 ' если не задана база, но задано рабочее пространство - берём первую в рабочем пространстве, иначе - ошибка
' ---
'' первый вариант проверки объект TableDef
'    Result = dbs.TableDefs(TableName).Name = TableName
'' второй вариант проверки AllTables
'    With CurrentData.AllTables(TableName)
'        IsLoaded = .IsLoaded
'        Result = .Name = TableName
'    End With
' третий вариант проверки - попытка открыть рекордсет
    strSQL = TableName
    'strSQL = sqlSelect1st & "[" & TableName & "]"
    'If IsMissing(Hash) Then strSQL = sqlSelect1st & "(" & strSQL & ")"
Dim rst As DAO.Recordset: Set rst = dbs.OpenRecordset(strSQL, dbOpenTable)
    rst.Close: Set rst = Nothing
    Result = True
HandleExit:  p_IsTableExists = Result: Exit Function
HandleError: Select Case Err
    Case 3008: Result = True ' Открыта другим пользователем для монопольного использования
    Case Else: Result = False
    End Select
    Err.Clear: Resume HandleExit
End Function

Private Function p_DateToSQL(FormatDate) As String
' форматирует дату/время для использования в SQL запросах
Dim strTemp As String: strTemp = "m\/d\/yyyy": If (FormatDate - Int(FormatDate)) Then strTemp = strTemp & " h\:n\:s"
    p_DateToSQL = Format$(FormatDate, "\#" & strTemp & "\#")
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFloatButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'=========================
Private Const c_strModule As String = "clsFloatButton"
'=========================
' Описание      : Класс работы с плавающей кнопкой
' Версия        : 1.0.2.453565708
' Дата          : 05.03.2024 13:41:57
' Автор         : Александр Куклюк (aka kluck)
' Примечание    : использует класс clsRelay. Автор идеи Александр Куклюк (aka kluck). _
'               : http://hiprog.com/index.php?option=com_content&task=view&id=687 _
' v.1.0.2       : 25.03.2019 - изменения позиционирования для ленточной формы (Кашкин Р.В.)
' v.1.0.0       : 06.04.2003 - исходная версия
'=========================
Private Const clngMinButtonSize = 330  ' Минимальный размер кнопки  ~0.6 cm
Private Const clngMaxButtonSize = 567  ' Максимальный размер кнопки ~1.0 cm
Private Const clngTimerInterval = 5    ' Период опроса формы по таймеру в мс
'=========================
' Пример использования класса:
'=========================
'Private Const cstrFloats = "txt1;txt2"  ' имена контролов с плавающими кнопками
'Private WithEvents mRelay As clsRelay   ' для ретрансляции событий группы контролов в форму
'Private colFloats As Collection         ' коллекция для хранения элементов clsFloatButton
'' инициализация кнопок
'Private Sub Form_Open(Cancel As Integer)
'    Set colFloats = New Collection: Set mRelay = New clsRelay: p_InitFloats
'End Sub
'' возвращает ссылку на вызвавшую форму
'Private Sub mRelay_GetParent(PARENT As Object): Set PARENT = Me: End Sub
'' обрабатывает нажатие
'Public Sub FloatButton_Click( _
'    EventControl As Access.Control) 'As Integer
'' обрабатываем нажатие
'    Select Case strControlName
'    Case "txt1": ' действие для кнопки txt1
'    Case "txt2": ' действие для кнопки txt2
'    End Select
'End Sub
'Private Sub p_InitFloats()
'' собственно инициализация обрабатываемых контролов и сохранение ссылок на них в массиве
'Dim i As Long, iMax As Long
'Dim arrCtls() As String
'Dim Ctl As Access.Control, cTransf As clsFloatButton
'
'    arrCtls = Split(cstrFloats, ";")
'    i = LBound(arrCtls): iMax = UBound(arrCtls)
'    Do While i <= iMax
'        Set Ctl = Me.Controls(Trim$(arrCtls(i)))
'    ' создаем новый экземпляр класса
'        Set cTransf = New clsFloatButton
''        cTransf.Action = "MsgBox('Поле " & i & "')"
'    ' собственно инициализация
'        cTransf.Init Index:=i, EventControl:=Ctl, Relay:=mRelay
'    ' сохраняем ссылку на объект класса,
'    ' чтобы обрабатывались события его встроенного объекта EventControl.
'        colFloats.Add cTransf
'        i = i + 1
'    Loop
'End Sub
'=========================
' !!! ОШИБКИ !!!:
'-------------------------
' 1. при прокрутке формы кнопка висит в прежней позиции -
' мораль: или отслеживать прокрутку и вносить поправки,
' можно попробовать фиксировать позицию и проверять по таймеру
' корректируя на смещение скролла
' или - не использовать на формах с прокруткой и ленточных(табличных) формах вовсе
' 2. при переходе от подчиненной к родительской форме, скроллбару и т.п.
' Exit и Deactivate не происходят - кнопка "зависает"
' помогает OnExit для контрола субформы, но не могу получить на него ссылку
'=========================
'--------------------------------------------------------------------------------
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINT
    x As Long
    y As Long
End Type
Private Type SCROLLBARINFO
    cbSize As Long
    rcScrollBar As RECT
    dxyLineButton As Long
    xyThumbTop As Long
    xyThumbBottom As Long
    reserved As Long
    rgstate(0 To 5) As Long
End Type
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Const SB_CTL = 2
 
'--------------------------------------------------------------------------------
' USER32
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>  PtrSafe, LongPtr and LongLong
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As Long
Private Declare PtrSafe Function ClientToScreen Lib "user32.dll" (ByVal hwnd As LongPtr, ByRef lpPoint As POINT) As Long
Private Declare PtrSafe Function ScreenToClient Lib "user32.dll" (ByVal hwnd As LongPtr, ByRef lpPoint As POINT) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Boolean) As Long
'Функции управления скроллбаром
Private Declare PtrSafe Function GetScrollPos Lib "user32" (ByVal hwnd As LongPtr, ByVal nBar As Long) As Long
Private Declare PtrSafe Function SetScrollPos Lib "user32" (ByVal hwnd As LongPtr, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function GetScrollBarInfo Lib "user32" (ByVal hwnd As LongPtr, ByVal idObject As Long, psbi As SCROLLBARINFO) As Long
Private Declare PtrSafe Function GetScrollInfo Lib "user32" (ByVal hwnd As LongPtr, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Private Declare PtrSafe Function SetScrollInfo Lib "user32" (ByVal hwnd As LongPtr, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
#Else                   '<OFFICE97-2010>        Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINT) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Boolean) As Long
'Функции управления скроллбаром
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetScrollBarInfo Lib "user32" (ByVal hwnd As Long, ByVal idObject As Long, psbi As SCROLLBARINFO) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
#End If                 '<WIN32>

Private mFloatForm As Form                              ' форма плавающей кнопки
Private WithEvents mFloatButton As Access.CommandButton ' плавающая кнопка
Attribute mFloatButton.VB_VarHelpID = -1
Private WithEvents mTextBox As Access.TextBox           ' вызывающие кнопку контролы
Attribute mTextBox.VB_VarHelpID = -1
Private WithEvents mComboBox As Access.ComboBox
Attribute mComboBox.VB_VarHelpID = -1
Private WithEvents mCheckBox As Access.CheckBox
Attribute mCheckBox.VB_VarHelpID = -1
Private WithEvents mListBox As Access.ListBox
Attribute mListBox.VB_VarHelpID = -1
Private mForm As Access.Form    ' форма, на которой появляется плавающая кнопка
Private mParent As Access.Form  ' верхняя родительская форма для контрола вызвавшего появление кнопки (если mForm не является подчиненной, указывает на mForm)
'Private WithEvents mSubForm As Access.SubForm ' контрол субформы формы mForm для отлова события Exit
Private mRelay As clsRelay
Private mIndex As Long
Private mControl As Control
Private mControlType As AcControlType
Private mAction As String
Private mIcon As String
Private mTooltip As String
Private mSize As Long
Private mVisible As Boolean
Private mFloatWidth As Long, mFloatHeight As Long

'Private mTargetHeight As Long   ' высота формы (строки для ленточной формы)
Private mSection As Long  ' номер секции контрола формы
Private mHwnd As LongPtr     ' hWnd формы
Private mVertSbHwnd As LongPtr, mVertSbSize As Long ', mVertSInfo As SCROLLINFO, mVertSbInfo As SCROLLBARINFO  ' информация о вертикальном сколбаре формы
Private mHorzSbHwnd As LongPtr, mHorzSbSize As Long ', mHorzSInfo As SCROLLINFO, mHorzSbInfo As SCROLLBARINFO ' информация о горизонтальном сколбаре формы

Private Const c_strCustomProc = "[Event Procedure]"

Private Sub Class_Initialize()
    Set mFloatForm = New Form_frmSERV_FloatButton
    Set mFloatButton = mFloatForm.btnButton
    mFloatButton.OnClick = c_strCustomProc
End Sub
Private Sub Class_Terminate()
    Set mFloatForm = Nothing
    Set mFloatButton = Nothing
    
    Set mTextBox = Nothing
    Set mComboBox = Nothing
    Set mListBox = Nothing
    Set mCheckBox = Nothing
    
    Set mForm = Nothing
    Set mParent = Nothing
    Set mRelay = Nothing
End Sub
Friend Sub Init( _
    Index As Long, _
    EventControl As Access.Control, _
    Relay As clsRelay _
    )
' инициализирует контрол для которого будет обработка событий
' вызов д.б. при открытии формы
Dim rc As RECT
Dim Result As Long
    mIndex = Index
    mSection = EventControl.Section
    Set mControl = EventControl
    Set mForm = mControl.PARENT.Form
    Set mParent = GetTopParent(mControl)
    Set mRelay = Relay
' получаем информацию о форме контролов для которых возникает кнопка
    mHwnd = mForm.hwnd
' обрабатываем скроллбары
    mHorzSbHwnd = ScrollbarGetHwnd(mHwnd, DIRECTION_HORIZONTAL)
    If mHorzSbHwnd <> 0 Then: Call GetWindowRect(mHorzSbHwnd, rc): mHorzSbSize = rc.Bottom - rc.Top ' ширина горизонтального скрола
    mVertSbHwnd = ScrollbarGetHwnd(mHwnd, DIRECTION_VERTICAL)
    If mVertSbHwnd <> 0 Then: Call GetWindowRect(mVertSbHwnd, rc): mVertSbSize = rc.Right - rc.Left ' ширина вертикального скрола
' цепляемся к событиям формы контролов для которых возникает кнопка
    mForm.OnTimer = c_strCustomProc
' цепляемся к событиям контролов для которых возникает кнопка
    If TypeOf EventControl Is Access.TextBox Then
        mControlType = acTextBox: Set mTextBox = EventControl
    ElseIf TypeOf EventControl Is Access.ComboBox Then
        mControlType = acComboBox: Set mComboBox = EventControl
    ElseIf TypeOf EventControl Is Access.ListBox Then
        mControlType = acListBox: Set mListBox = EventControl
    ElseIf TypeOf EventControl Is Access.CheckBox Then
        mControlType = acCheckBox: Set mCheckBox = EventControl
    Else
        GoTo HandleExit
    End If
    With EventControl
        .OnEnter = c_strCustomProc
        .OnKeyDown = c_strCustomProc
        .OnExit = c_strCustomProc
    End With
HandleExit:  Exit Sub
HandleError: Err.Clear: Resume HandleExit
End Sub
Public Property Let Index(rData As Long):       mIndex = rData: End Property
Public Property Get Index() As Long:            Index = mIndex: End Property
Public Property Let Action(rData As String):    mAction = rData: End Property
Public Property Let Icon(rData As String):      mIcon = rData: End Property
Public Property Let Tooltip(rData As String):   mTooltip = rData: End Property
Public Property Let Size(rData As Long):        mSize = rData: End Property
Public Property Get Control() As Access.Control
    Select Case mControlType
    Case acTextBox:    Set Control = mTextBox
    Case acComboBox:   Set Control = mComboBox
    Case acListBox:    Set Control = mListBox
    Case acCheckBox:   Set Control = mCheckBox
    Case Else:         Set Control = Nothing
    End Select
End Property
Public Sub Show()
'установка координат формы с кнопкой
On Error Resume Next
    mVisible = True
    p_InitButton mSize
    ' устанавливаем интервал таймера для фиксации скроллов
    If mHorzSbHwnd <> 0 Or mVertSbHwnd <> 0 Then mForm.TimerInterval = clngTimerInterval
    'передача фокуса основной форме - иначе фокус останется на кнопке
    If mParent.Visible Then mParent.SetFocus
End Sub
Public Sub Hide()
    mVisible = False: p_SetButton
    mForm.TimerInterval = 0 ' снимаем интервал таймера
End Sub

'-----------------------------
' обработчики событий контрола передают события и контрол их вызвавший на обработку в родительский объект
'-----------------------------
Private Sub mFloatButton_Click()
On Error Resume Next
    If Len(mAction) > 0 Then Eval mAction: If Err Then Err.Clear
    mParent.SetFocus
    'RaiseEvent Click(Me.Control) ', Err)
    Call mRelay.PARENT.FloatButton_Click(mControl, mAction)
End Sub
' при входе - выводим
Private Sub mTextBox_Enter(): Show: End Sub ':DoEvents
Private Sub mComboBox_Enter(): Show: End Sub ':DoEvents
Private Sub mListBox_Enter(): Show: End Sub ':DoEvents
Private Sub mCheckBox_Enter(): Show: End Sub ':DoEvents
' при нажатии кнопки - обрабатываем
Private Sub mTextbox_KeyDown(KeyCode As Integer, Shift As Integer): Call p_KeyHandler(KeyCode, Shift): End Sub
Private Sub mCombobox_KeyDown(KeyCode As Integer, Shift As Integer): Call p_KeyHandler(KeyCode, Shift): End Sub
Private Sub mListBox_KeyDown(KeyCode As Integer, Shift As Integer): Call p_KeyHandler(KeyCode, Shift): End Sub
Private Sub mCheckBox_KeyDown(KeyCode As Integer, Shift As Integer): Call p_KeyHandler(KeyCode, Shift): End Sub
' при потере фокуса - прячем
Private Sub mTextBox_Exit(Cancel As Integer): Hide: End Sub ':DoEvents
Private Sub mComboBox_Exit(Cancel As Integer): Hide: End Sub ':DoEvents
Private Sub mListBox_Exit(Cancel As Integer): Hide: End Sub ':DoEvents
Private Sub mSubForm_Exit(Cancel As Integer): Hide: DoEvents: End Sub '
Private Sub mCheckBox_Exit(Cancel As Integer): Hide: DoEvents: End Sub '

Private Sub mTarget_Timer()
Dim Cancel As Boolean
' плохо - событие возникает для каждого инициализированого контрола
' соответственно все действия повторяются соответствующее число раз

' сравниваем Screen.ActiveControl с mControl
' и скрываем если не совпадает
' mVisible = IsActiveControl

' сравниваем позиции скроллбаров с сохраненными
' и меняем координаты кнопки при необходимости
Static hPos As Long, vPos As Long  ' позиции скроллбаров
Dim hCur As Long, vCur As Long
    If mHorzSbHwnd <> 0 Then hCur = GetScrollPos(mHorzSbHwnd, SB_CTL): If hCur <> hPos Then hPos = hCur: p_SetButton
    If mVertSbHwnd <> 0 Then vCur = GetScrollPos(mVertSbHwnd, SB_CTL): If vCur <> vPos Then vPos = vCur: p_SetButton
    'DoEvents
End Sub
'-----------------------------
' внутренние процедуры
'-----------------------------
Private Sub p_KeyHandler(KeyCode As Integer, Shift As Integer)
' обрабатывает нажатия клавиш
' нужна для перехода по записям ленточной формы
    On Error GoTo HandleError
    Select Case KeyCode
    Case vbKeyDown: If mForm.NewRecord = False Then DoCmd.RunCommand acCmdRecordsGoToNext
    Case vbKeyUp:   If mForm.CurrentRecord > 1 Then DoCmd.RunCommand acCmdRecordsGoToPrevious
    End Select
HandleExit:  Exit Sub
HandleError: KeyCode = 0
End Sub
Private Function p_InitButton( _
    Optional ButtonSize As Long)
' ButtonSize - размер кнопки (Высота=Ширина)
Dim Result As Long
' не обрабатываем если не отображается (Hide)
    If Not mVisible Then mFloatForm.Visible = mVisible: GoTo HandleExit
' определяем размер кнопки
    If ButtonSize <= 0 Then ButtonSize = mControl.Height
    If ButtonSize < clngMinButtonSize Then
        ButtonSize = clngMinButtonSize
    ElseIf ButtonSize > clngMaxButtonSize Then
        ButtonSize = clngMaxButtonSize
    End If
' задаем размер кнопки
    With mFloatButton: .Height = ButtonSize: .Width = ButtonSize: End With
    mFloatWidth = TwipsToPixels(ButtonSize, DIRECTION_HORIZONTAL)
    mFloatHeight = TwipsToPixels(ButtonSize, DIRECTION_VERTICAL)
' подчинение формы с кнопкой целевой форме
    SetParent mFloatForm.hwnd, mHwnd
' настраиваем оформление плавающей кнопки
    If Len(mIcon) > 0 Then PictureData_SetToControl mFloatButton, mIcon
    mFloatButton.ControlTipText = mTooltip
' устанавливаем кнопку в нужную позицию и задаём размеры
    Call p_SetButton
HandleExit:  p_InitButton = Result: Exit Function
HandleError: Result = Err.Number: Err.Clear: Resume HandleExit
End Function
Private Function p_SetButton()
Dim Result As Long
' не обрабатываем если не отображается (Hide)
    If Not mVisible Then: mFloatForm.Visible = mVisible: GoTo HandleExit
' экранные координаты верхнего правого угла контрола
Dim xClient As Long, rtClient As RECT: xClient = GetClientRect(mHwnd, rtClient)
Dim lpPoint As POINT, w As Long
    With lpPoint:
        Call AccControlLocation(mControl, .x, .y, w)
        .x = .x + w: .y = .y - 2
    End With
' экранные координаты в клиентские
    ScreenToClient mHwnd, lpPoint
' прячем если вылазит за границы
    If lpPoint.x < rtClient.Left Or lpPoint.y < rtClient.Top Then
    ' если кнопка за левым или верхним краем - прячем
            mFloatForm.Visible = False
    ElseIf lpPoint.y + mFloatHeight > rtClient.Bottom - mHorzSbSize Then
    ' если кнопка за нижним краем - прячем
            mFloatForm.Visible = False
    ElseIf lpPoint.x + mFloatWidth > rtClient.Right - mVertSbSize Then
    ' если кнопка за правым краем - проверяем
        If lpPoint.x - mFloatWidth < rtClient.Right - mVertSbSize Then
        ' если кнопка умещается в пределах поля,
        ' при выравнивании по правой границе, -
        ' выравниваем по правой границе поля
            lpPoint.x = rtClient.Right - mFloatWidth - mVertSbSize
            mFloatForm.Visible = True
        Else
        ' иначе - скрываем
            mFloatForm.Visible = False
        End If
    Else
    ' если находится в видимой области формы - выводим
            mFloatForm.Visible = True
    End If
' установка координат формы с кнопкой
    If mFloatForm.Visible Then Result = MoveWindow(mFloatForm.hwnd, lpPoint.x, lpPoint.y, mFloatWidth, mFloatHeight, -1)
HandleExit:  p_SetButton = Result: Exit Function
HandleError: Result = Err.Number: Err.Clear: Resume HandleExit
End Function


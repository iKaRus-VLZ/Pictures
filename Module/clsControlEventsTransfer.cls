VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsControlEventsTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const DEBUGGING = True
Option Explicit
Option Compare Database
'=========================
Private Const c_strModule As String = "clsControlEventsTransfer"
'=========================
' Описание      : Класс трансляции событий контрола
' Версия        : 1.0.1.436484819
' Дата          : 02.07.2019 11:33:56
' Автор         : Кашкин Р.В. (KashRus@gmail.com)
' Примечание    : необходим модуль clsRelay, форма должна иметь свойство HasModule=True
' v.1.0.1       : добавлены события клавиатуры, слегка изменен дизайн
'=========================
' неиспользуемые в приложении типы контролов и событий лучше закомментить
'' Пример использования класса для групповой обработки событий контролов:
'    Private WithEvents mRelay As clsRelay   ' для ретрансляции событий группы контролов в форму
'    Private colCtls As Collection           ' коллекция для хранения элементов clsControlEventsTransfer
'
'    Private Sub Form_Open(Cancel As Integer)
'        Set colCtls = New Collection
'        Set mRelay = New clsRelay
'    End Sub
'    Private Sub Form_Close()
'        Set colCtls = Nothing
'        Set mRelay = Nothing
'    End Sub
'    Private Sub mRelay_GetParent(PARENT As Object)
'        Set PARENT = Me
'    End Sub

'    Private Sub p_InitControl(ctl As Access.Control)
'    'p_InitControl необходимо вызвать для каждого контрола который должен получить групповой обработчик событий
'    Dim cTransf As New clsControlEventsTransfer
'        cTransf.Init Index:=colCtls.Count, EventControl:=ctl, Relay:=mRelay
'        colCtls.Add cTransf
'    End Sub
'    Public Sub EventControl_Click(EventControl As Access.Control) 'As Integer
'    ' собственно пример группового обработчика событий
'    ' EventControl - содержит ссылку на контрол вызвавший событие
'    End Sub
'=========================
' !!! неиспользуемые закомментировать !!!
Private WithEvents mLabel As Access.Label
Attribute mLabel.VB_VarHelpID = -1
Private WithEvents mTextBox As Access.TextBox
Attribute mTextBox.VB_VarHelpID = -1
Private WithEvents mComboBox As Access.ComboBox
Attribute mComboBox.VB_VarHelpID = -1
Private WithEvents mCheckBox As Access.CheckBox
Attribute mCheckBox.VB_VarHelpID = -1
Private WithEvents mCommandButton As Access.CommandButton
Attribute mCommandButton.VB_VarHelpID = -1
Private WithEvents mToggleButton As Access.ToggleButton
Attribute mToggleButton.VB_VarHelpID = -1

Private mRelay As clsRelay
Private mIndex As Long
Private mControlType As AcControlType

Friend Sub Init( _
    Index As Long, _
    EventControl As Access.Control, _
    Relay As clsRelay _
    )
' инициализирует контрол для которого будет обработка событий
' вызов д.б. при открытии формы
Const c_strProcedure = "Init"
    
    On Error GoTo HandleError
    mIndex = Index
    Set mRelay = Relay
    If TypeOf EventControl Is Access.Label Then
        mControlType = acLabel: Set mLabel = EventControl
    ElseIf TypeOf EventControl Is Access.TextBox Then
        mControlType = acTextBox: Set mTextBox = EventControl
    ElseIf TypeOf EventControl Is Access.ComboBox Then
        mControlType = acComboBox: Set mComboBox = EventControl
    ElseIf TypeOf EventControl Is Access.CheckBox Then
        mControlType = acCheckBox: Set mCheckBox = EventControl
    ElseIf TypeOf EventControl Is Access.CommandButton Then
        mControlType = acCommandButton: Set mCommandButton = EventControl
    ElseIf TypeOf EventControl Is Access.ToggleButton Then
        mControlType = acToggleButton: Set mToggleButton = EventControl
    Else
        GoTo HandleExit
    End If
    With EventControl
' !!! неиспользуемые закомментировать !!!
        .OnClick = c_strCustomProc
        .OnDblClick = c_strCustomProc
        .OnMouseMove = c_strCustomProc
        .OnMouseDown = c_strCustomProc
        .OnMouseUp = c_strCustomProc
        .OnAfterUpdate = c_strCustomProc
        '.OnKeyDown = c_strCustomProc
        .OnKeyUp = c_strCustomProc
        '.ShortcutMenuBar = Space(1) ' чтобы не открывалось меню по-умолчанию
    End With
HandleExit:
    Exit Sub
HandleError:
'    Dbg.Error Err.Number, Err.Description, _
'        COMMENT:="", _
'        Source:=c_strModule, Procedure:=c_strProcedure ', LineNum :=Erl
    Err.Clear
    Resume HandleExit
End Sub
Private Sub Class_Terminate()
' !!! неиспользуемые закомментировать !!!
    Set mLabel = Nothing
    Set mTextBox = Nothing
    Set mComboBox = Nothing
    Set mCheckBox = Nothing
    Set mToggleButton = Nothing
    Set mCommandButton = Nothing
    Set mRelay = Nothing
End Sub
'-----------------------------
' обработчики событий контрола
'-----------------------------
Public Property Let Index(rData As Long): mIndex = rData: End Property
Public Property Get Index() As Long: Index = mIndex: End Property

Public Property Get Control() As Access.Control
    Select Case mControlType
' !!! неиспользуемые закомментировать !!!
     Case acLabel:          Set Control = mLabel
     Case acTextBox:        Set Control = mTextBox
     Case acComboBox:       Set Control = mComboBox
     Case acCheckBox:       Set Control = mCheckBox
     Case acCommandButton:  Set Control = mCommandButton
     Case acToggleButton:   Set Control = mToggleButton
     Case Else:             Set Control = Nothing
    End Select
End Property

'-----------------------------
' обработчики событий контрола
' передают события и контрол их вызвавший на обработку в родительский объект
' !!! неиспользуемые закомментировать !!!
'-----------------------------
' Click
'-----------------------------
Private Sub mLabel_Click(): On Error Resume Next: Call mRelay.PARENT.EventControl_Click(Me.Control): End Sub
Private Sub mTextbox_Click(): On Error Resume Next: Call mRelay.PARENT.EventControl_Click(Me.Control): End Sub
Private Sub mCombobox_Click(): On Error Resume Next: Call mRelay.PARENT.EventControl_Click(Me.Control): End Sub
Private Sub mCommandButton_Click(): On Error Resume Next: Call mRelay.PARENT.EventControl_Click(Me.Control): End Sub
Private Sub mCheckBox_Click(): On Error Resume Next: Call mRelay.PARENT.EventControl_Click(Me.Control): End Sub
Private Sub mToggleButton_Click(): On Error Resume Next: Call mRelay.PARENT.EventControl_Click(Me.Control): End Sub
'-----------------------------
' DblClick
'-----------------------------
Private Sub mLabel_DblClick(Cancel As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_DblClick(Me.Control, Cancel): End Sub
Private Sub mTextbox_DblClick(Cancel As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_DblClick(Me.Control, Cancel): End Sub
Private Sub mCombobox_DblClick(Cancel As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_DblClick(Me.Control, Cancel): End Sub
Private Sub mCommandButton_DblClick(Cancel As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_DblClick(Me.Control, Cancel): End Sub
Private Sub mCheckBox_DblClick(Cancel As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_DblClick(Me.Control, Cancel): End Sub
Private Sub mToggleButton_DblClick(Cancel As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_DblClick(Me.Control, Cancel): End Sub
'-----------------------------
' MouseMove
'-----------------------------
Private Sub mLabel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): On Error Resume Next: Call mRelay.PARENT.EventControl_MouseMove(Me.Control, Button, Shift, x, y): End Sub
Private Sub mTextbox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): On Error Resume Next: Call mRelay.PARENT.EventControl_MouseMove(Me.Control, Button, Shift, x, y): End Sub
Private Sub mCombobox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): On Error Resume Next: Call mRelay.PARENT.EventControl_MouseMove(Me.Control, Button, Shift, x, y): End Sub
Private Sub mCommandButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): On Error Resume Next: Call mRelay.PARENT.EventControl_MouseMove(Me.Control, Button, Shift, x, y): End Sub
Private Sub mCheckBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): On Error Resume Next: Call mRelay.PARENT.EventControl_MouseMove(Me.Control, Button, Shift, x, y): End Sub
Private Sub mToggleButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single): On Error Resume Next: Call mRelay.PARENT.EventControl_MouseMove(Me.Control, Button, Shift, x, y): End Sub
'-----------------------------
' MouseDown
'-----------------------------
Private Sub mLabel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single): On Error Resume Next: Call mRelay.PARENT.EventControl_MouseDown(Me.Control, Button, Shift, x, y): End Sub
Private Sub mTextbox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single): On Error Resume Next: Call mRelay.PARENT.EventControl_MouseDown(Me.Control, Button, Shift, x, y): End Sub
Private Sub mCombobox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single): On Error Resume Next: Call mRelay.PARENT.EventControl_MouseDown(Me.Control, Button, Shift, x, y): End Sub
Private Sub mCommandButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single): On Error Resume Next: Call mRelay.PARENT.EventControl_MouseDown(Me.Control, Button, Shift, x, y): End Sub
Private Sub mCheckBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single): On Error Resume Next: Call mRelay.PARENT.EventControl_MouseDown(Me.Control, Button, Shift, x, y): End Sub
Private Sub mToggleButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single): On Error Resume Next: Call mRelay.PARENT.EventControl_MouseDown(Me.Control, Button, Shift, x, y): End Sub
''-----------------------------
'' KeyDown
''-----------------------------
'Private Sub mTextbox_KeyDown(KeyCode As Integer, Shift As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_KeyDown(Me.Control, KeyCode, Shift): End Sub
'Private Sub mCombobox_KeyDown(KeyCode As Integer, Shift As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_KeyDown(Me.Control, KeyCode, Shift): End Sub
'Private Sub mCheckBox_KeyDown(KeyCode As Integer, Shift As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_KeyDown(Me.Control, KeyCode, Shift): End Sub
'Private Sub mCommandButton_KeyDown(KeyCode As Integer, Shift As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_KeyDown(Me.Control, KeyCode, Shift): End Sub
'Private Sub mToggleButton_KeyDown(KeyCode As Integer, Shift As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_KeyDown(Me.Control, KeyCode, Shift): End Sub
'-----------------------------
' KeyUp
'-----------------------------
Private Sub mTextbox_KeyUp(KeyCode As Integer, Shift As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_KeyUp(Me.Control, KeyCode, Shift): End Sub
Private Sub mCombobox_KeyUp(KeyCode As Integer, Shift As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_KeyUp(Me.Control, KeyCode, Shift): End Sub
Private Sub mCheckBox_KeyUp(KeyCode As Integer, Shift As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_KeyUp(Me.Control, KeyCode, Shift): End Sub
Private Sub mCommandButton_KeyUp(KeyCode As Integer, Shift As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_KeyUp(Me.Control, KeyCode, Shift): End Sub
Private Sub mToggleButton_KeyUp(KeyCode As Integer, Shift As Integer): On Error Resume Next: Call mRelay.PARENT.EventControl_KeyUp(Me.Control, KeyCode, Shift): End Sub
'-----------------------------
' AfterUpdate
'-----------------------------
Private Sub mTextbox_AfterUpdate(): On Error Resume Next: Call mRelay.PARENT.EventControl_AfterUpdate(Me.Control): End Sub
Private Sub mCombobox_AfterUpdate(): On Error Resume Next: Call mRelay.PARENT.EventControl_AfterUpdate(Me.Control): End Sub
Private Sub mCheckBox_AfterUpdate(): On Error Resume Next: Call mRelay.PARENT.EventControl_AfterUpdate(Me.Control): End Sub
Private Sub mCommandButton_AfterUpdate(): On Error Resume Next: Call mRelay.PARENT.EventControl_AfterUpdate(Me.Control): End Sub
Private Sub mToggleButton_AfterUpdate(): On Error Resume Next: Call mRelay.PARENT.EventControl_AfterUpdate(Me.Control): End Sub


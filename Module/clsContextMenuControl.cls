VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsContextMenuControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const DEBUGGING = True
Option Compare Database
Option Explicit
'=========================
Private Const c_strModule As String = "clsContextMenuControl"
'=========================
' Описание      : Класс-обёртка для перехвата кликов по пунктам меню
' Версия        : 1.0.0.0
' Дата          : 08.08.2016 10:59:49
' Автор         : Кашкин Р.В.
' Примечание    : использует clsContextMenu для передачи в него событий
'=========================

Private mManager As clsContextMenu
Private mControl As Office.CommandBarControl
Private mValue

Private WithEvents mCommandButton As Office.CommandBarButton
Attribute mCommandButton.VB_VarHelpID = -1
Private WithEvents mCommandComboBox As Office.CommandBarComboBox
Attribute mCommandComboBox.VB_VarHelpID = -1

'Public Event ItemSelect(Control As Office.CommandBarControl, Value)

Public Sub Init( _
    ByRef ContextMenuControl As Office.CommandBarControl, _
    ByRef ContextMenu As clsContextMenu)
    
    Set mControl = ContextMenuControl
    Set mManager = ContextMenu
    If (mControl Is Nothing) Or (mManager Is Nothing) Then Exit Sub
'    With Office.MsoControlType
    Select Case mControl.Type
     Case msoControlButton:   Set mCommandButton = mControl
     Case msoControlComboBox: Set mCommandComboBox = mControl
    End Select
'    End With
End Sub
Public Property Get Control() As Office.CommandBarControl
    Set Control = mControl
End Property

'Public Property Let Value(rData)
'    mValue = rData
'End Property
'Public Property Set Value(rData)
'    Set mValue = rData
'End Property
'Public Property Get Value()
'    If IsObject(mValue) Then Set Value = mValue Else Value = mValue
'End Property

Private Sub mCommandButton_Click(ByVal MenuCtrl As Office.CommandBarButton, CancelDefault As Boolean)
    mManager.MenuClick MenuCtrl
    CancelDefault = False
End Sub

Private Sub mCommandComboBox_Change(ByVal MenuCtrl As Office.CommandBarComboBox)
    mManager.MenuClick MenuCtrl
'    CancelDefault = False
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsContextMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'=========================
Private Const c_strModule As String = "clsContextMenu"
'=========================
' ��������      : ��������� ����������� ����
' ������        : 1.0.0.452896330
' ����          : 29.12.2023 15:11:31
' �����         : ������ �.�.
' ����������    : ���� OnAction ����� - mContextMenu.ActionControl ���������� ������� _
'               : �� ������� ���� ������� OnAction ��������� _
'               : ���� OnAction �� ����� - ���������� clsContextMenuControl ��� ��������� ������� ��������� _
'               : ���������� Microsof Office Object Library
'=========================
' ������ ������������� ������:
'=========================
'=========================
Private Const cstrMenuName = "~tmpMenu"
Private Const NOERROR As Long = 0
Private mValue
Private mstrActionProc As String '"MenuButtonClick"
Private mstrContextMenuName As String '"~tmpMenu"
Private mContextMenu As Office.CommandBar
Private mContextMenuControls As Collection

Private Sub Class_Initialize()
    mValue = Null
    mstrActionProc = vbNullString
    mstrContextMenuName = cstrMenuName
    Set mContextMenuControls = New Collection
End Sub
Private Sub Class_Terminate(): Set mContextMenuControls = Nothing: End Sub
Public Property Let Value(rData): mValue = rData: End Property
Public Property Get Value(): Value = mValue: End Property
Public Function Visible() As Boolean: Visible = mContextMenu.Visible: End Function
Public Property Get ContextMenu() As Office.CommandBar: Set ContextMenu = mContextMenu: End Property

Public Property Let OnActionProc(rData As String)
' ������ ���������� �������� OnAction ��� ��������
Dim tmp As Object
    'If mContextMenuControls.Count > 0 Then
    For Each tmp In mContextMenuControls
        tmp.Control.OnAction = rData
    Next tmp
    'End If
    mstrActionProc = rData
End Property
Public Sub CreateContextMenu(Optional ByVal MenuName As String)
' ������� ������������ ���� � ������� ������
    MenuName = Trim$(MenuName)
    If Len(MenuName) = 0 Then MenuName = mstrContextMenuName
    p_CreateMenu MenuName
    mstrContextMenuName = MenuName
End Sub
Public Sub RemoveContextMenu(Optional ByVal MenuName As String)
' ������� ������������ ���� � ������� ������
    MenuName = Trim$(MenuName)
    If Len(MenuName) = 0 Then MenuName = mstrContextMenuName
    p_RemoveMenu MenuName
    mstrContextMenuName = MenuName
End Sub

Public Function CreateItemsFromString( _
    ByVal Source As String, _
    Optional Delim As String = ";") As Boolean
' ��������� �������� ���� �� ������
Const c_strProcedure = "CreateItemsFromString"
' <��� �������� 1>;<������������ �������� 1><��� ������ 1>;...<��� �������� n>;<������������ �������� n>;<��� ������ n>
'Const iStep = 3 ' � ������ 3 ��������
' �������� ������ ���� ������� �� ��� ���������:
' Caption - ������� (��� ��������), ���� Caption = vbNullString ���� ������� ������� - ������� ������ ���� (������ ����� ������)
' Tag - ���.������ (������������ ��� ������ ��������)
' Icon - ������
Dim Result As Boolean ':Result = False
    On Error GoTo HandleError
Dim Items() As String, i As Long, iMax As Long
Dim strText As String, strTag As String, strIcon As String, bNew As Boolean
    Items = Split(Source, Delim)
    i = LBound(Items): iMax = UBound(Items)
    Do Until i > iMax
        strText = Trim(Items(i))
        If Len(strText) Then                ' ���� ������� ������ ���� ��������
            i = i + 1: strTag = Trim(Items(i))  ' ������ ������������ ��������
            i = i + 1: strIcon = Trim(Items(i)) ' � ������
            ' ��������� ���� � ���������� ���� ����� ������
            p_CreateMenuItem strText, ItemTag:=strTag, ItemIcon:=strIcon, BeginGroup:=bNew: bNew = False
        Else                                ' ���� ������� ������ ���� ������
            bNew = True                         ' ������������� ���� ����� ������
        End If
        i = i + 1
    Loop
    Result = True
HandleExit:  CreateItemsFromString = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function CreateItemsFromSQL( _
    ByVal Source As String, _
    Optional ByVal Captions As String = c_strName, _
    Optional ByVal Tags As String = c_strKey, _
    Optional ByVal Actions As String, _
    Optional ByVal Descs As String, _
    Optional ByVal IconName As String = c_strFaceKey, _
    Optional ByVal Params As String, _
    Optional ByVal WhereCond As String) As Boolean
' ��������� �������� ���� �� ��������� SQL
Const c_strProcedure = "CreateItemsFromSQL"

' ����:
'   ID - � Tag
'   NAME ��� CNAME - � Caption
'   PARENT - ��� ����������� ����������� ���������
'       ��� ��������  .Add(Type:=msoControlPopup)
'       ��� ������������  .Add(Type:=msoControlButton)
Dim Result As Boolean ': Result = False
    On Error GoTo HandleError
    Captions = Trim$(Captions): Tags = Trim$(Tags)
    Actions = Trim$(Actions):   Params = Trim$(Params)
    Descs = Trim$(Descs)
    WhereCond = Trim$(WhereCond)
    
    If Len(Captions) = 0 Then Captions = c_strName  ' ��� ���� ��-���������
    If Len(Tags) = 0 Then Tags = c_strKey           ' ��� ���� ��-���������
    
Dim strSQL As String
Dim strTag As String, strCap As String, strIco As String
    strSQL = sqlSelect & Captions
    If Len(Actions) > 0 Then strSQL = strSQL & c_strInDelim & Actions
    If Len(Descs) > 0 Then strSQL = strSQL & c_strInDelim & Descs
    If Len(Params) > 0 Then strSQL = strSQL & c_strInDelim & Params
    If Len(Tags) > 0 Then strSQL = strSQL & c_strInDelim & Tags
    If Len(IconName) > 0 Then strSQL = strSQL & c_strInDelim & IconName
    strSQL = strSQL & sqlFrom & "(" & Source & ")"
    If Len(WhereCond) > 0 Then strSQL = strSQL & sqlWhere & WhereCond
    If Len(strSQL) > 0 Then strSQL = strSQL & ";"
    
    On Error GoTo HandleError
    
Dim dbs As DAO.Database: Set dbs = CurrentDb ' App.AppData
Dim rst As DAO.Recordset: Set rst = dbs.OpenRecordset(strSQL)
    With rst
        
        Do Until .EOF
    On Error Resume Next
            strCap = Nz(.Fields(Captions).Value, vbNullString)
            strTag = Nz(.Fields(Tags).Value, vbNullString)
            strIco = Nz(.Fields(IconName).Value, vbNullString)
'            strAct = Nz(.Fields(Actions).Value, vbNullString)
'            strDsc = Nz(.Fields(Descs).Value, vbNullString)
'            strPar = Nz(.Fields(Params).Value, vbNullString)
    On Error GoTo HandleError
            p_CreateMenuItem ItemCaption:=strCap, ItemTag:=strTag, ItemIcon:=strIco ',strOnAction
            .MoveNext
        Loop
    End With
    Result = True
HandleExit:  CreateItemsFromSQL = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Sub ShowMenu(Optional x, Optional y)
    'p_CreateMenu
    Application.CommandBars(mstrContextMenuName).ShowPopup x, y
    'p_RemoveMenu
End Sub
Public Sub MenuClick(MenuControl As Object)
' ������������ ����� ������ ������������ ����
On Error GoTo HandleError
    With MenuControl
        Select Case .Type
        Case msoControlButton:   mValue = .Tag
        Case msoControlComboBox: mValue = .Tag
        Case Else
        End Select
    End With
HandleExit:  Exit Sub
HandleError: Err.Clear: Resume HandleExit
End Sub


Private Sub p_CreateMenu(MenuName As String)
    p_RemoveMenu MenuName ' ������� ������� ����
    Set mContextMenu = Application.CommandBars.Add(MenuName, Office.msoBarPopup)
End Sub
Private Sub p_RemoveMenu(ContextMenuName As String)
Dim mContextMenu As Office.CommandBar
    
On Error GoTo HandleError
    For Each mContextMenu In Application.CommandBars
'Debug.Print mContextMenu.Name
        If (mContextMenu.Name = ContextMenuName) Then mContextMenu.Delete
    Next
    Set mContextMenuControls = Nothing
    Set mContextMenuControls = New Collection  '���� ������� ���������
HandleExit:  Exit Sub
HandleError:
Stop
End Sub
    
Private Sub p_CreateMenuItem( _
    ItemCaption As String, _
    Optional ItemOnAction As String, _
    Optional ItemType As MsoControlType = msoControlButton, _
    Optional ItemTag As String, _
    Optional ItemParam As String, _
    Optional ItemDesc As String, _
    Optional ItemTooltip As String, _
    Optional ItemIcon As String, _
    Optional ItemRowSource As String, _
    Optional BeginGroup As Boolean = False, _
    Optional ParentObject)
' ������� ������� ����
Dim mControl As Office.CommandBarControl
Dim mCommand As clsContextMenuControl
    
    If IsMissing(ParentObject) Then
        Set ParentObject = mContextMenu
    ElseIf TypeOf ParentObject Is Office.CommandBarPopup Then
        ' ���� �������� - ���������� ���� ��������� � ����
    Else
        Set ParentObject = mContextMenu
    End If
    Set mControl = ParentObject.Controls.Add(ItemType)
    On Error Resume Next
    With mControl
        '.ID ' ��� ���������
        .OnAction = ItemOnAction
        .Tag = ItemTag
        .BeginGroup = BeginGroup
        '.Parameter =ItemParam
        '.DescriptionText =ItemDesc
        '.TooltipText = ItemTooltip
        Select Case .Type
        Case msoControlButton      ' ������
        .Caption = ItemCaption: .Style = msoButtonCaption
        If Len(ItemIcon) = 0 Then GoTo HandleAdd
        If PictureData_SetToControl(mControl, ItemIcon) = NOERROR Then .Style = msoButtonIconAndCaption
'#If ObjectDataType = 0 Then     'FI
'Dim fiPict As LongPtr:      fiPict = PictureData_LoadFromEx(ItemIcon, 16):  If fiPict = 0 Then GoTo HandleAdd
'        .Picture = FreeImage_GetOlePictureIcon(fiPict, UnloadSource:=True)
'#ElseIf ObjectDataType = 1 Then 'LV
'        If PictureData_SetToControl(mControl, ItemIcon) = NOERROR Then .Style = msoButtonIconAndCaption
'#End If                         'ObjectDataType
        Case msoControlComboBox    ' ���� �� �������
        ' ���� ItemRowSource
        ' Do
        '   .AddItem ItemRowText
        ' Loop
        'Case msoCommandBarPopup       ' ���������� ����
        End Select
    End With
HandleAdd:
    If Err Then Err.Clear ': Stop
    Set mCommand = New clsContextMenuControl
    mCommand.Init mControl, Me
    mContextMenuControls.Add mCommand
End Sub


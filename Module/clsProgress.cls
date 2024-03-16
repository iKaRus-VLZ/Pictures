VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'=========================
Private Const c_strModule As String = "clsProgress"
'=========================
' Описание      : Класс для работы с прогрессбаром
' Версия        : 1.0.0.451174994
' Дата          : 10.07.2023 11:59:08
' Автор         :
' Примечание    : необходимо наличие форм frmSERV_Progress
'=========================
' Примеры использования:
'===========================================
''*** пример одноуровневого прогрессбара ***
''------------------------------------------
'Dim prg As clsProgress: Set prg = New clsProgress
'    With prg
'        .Init pMin:=-100.12, pMax:=1750.56, pStep:=0.5, _
'            pCaption:="Демонстрация прогресса", _
'            pText:="Выполнение..."
'        .Show
'        Do Until .Progress = .ProgressMax 'And Not .Canceled
'            .Detail = "проверяем однопроцессный прогрессбар" & vbCrLf & _
'                "Текущее значение " & Format$(.Progress, "Standard") & ", в диапазоне" & " от " & .ProgressMin & " до " & .ProgressMax
'            'Sleep 50 ' <здесь выполняются какие-то действия
'            If .Canceled Then MsgBox "Процесс прерван на значении: " & .Progress: Exit Do
'            .Update
'        Loop
'    End With
'    Set prg = Nothing
''------------------------------------------
''*** пример двухуровневого прогрессбара ***
''------------------------------------------
'Dim prg As clsProgress: Set prg = New clsProgress
'    With prg
'        .Init pCount:=2
'        .Caption = "Демонстрация двойного прогресса"
'        .FloodColor(eMajor) = vbRed:    .BarTextColor(eMajor) = vbBlue
'        .FloodColor(eMinor) = vbCyan:   .BarTextColor(eMinor) = vbWhite
'        .ProgressMax(eMajor) = 10:      .ProgressStep(eMajor) = 1
'        Do Until .Progress(eMajor) = .ProgressMax(eMajor) 'And Not .Canceled
'            .ProgressMax(eMinor) = 100 * (.Progress(eMajor) + 1)
'            Do Until .Progress(eMinor) = .ProgressMax(eMinor) 'And Not .Canceled
'            ' прерывание процесса
'                If .Canceled Then
'                    Select Case MsgBox("Прервать выполнение?", vbYesNo Or vbExclamation Or vbDefaultButton2)
'                     Case vbYes: Exit Do
'                     Case Else: .Canceled = False
'                    End Select
'                End If
'                'Sleep 50 ' <здесь выполняются какие-то действия второго уровня
'                .Update (eMinor)
'            Loop
'            If .Canceled Then Exit Do
'            'Sleep 50 ' <здесь выполняются какие-то действия первого уровня
'            .Update (eMajor)
'        Loop
'    End With
'    Set prg = Nothing
'=========================
Const cstrHostForm = "frmSERV_Progress"     ' имя основной формы прогреса
Const cstrCaptionLabel = "lblCaptionText"   ' имя контрола текстовой метки подписи
Const cstrMajorTextLabel = "lblMajorText"   ' имя контрола первичной текстовой метки
Const cstrMinorTextLabel = "lblMinorText"   ' имя контрола вторичной текстовой метки
Const cstrDetailLabel = "lblDetailText"     ' имя контрола текстовой метки деталей

Private Enum eType
    eCaption = 0
    eMajor = 1
    eMinor = 2
    eDetail = 3
End Enum

Private mProgressCount As Byte
Private mCaptionText As String, mDetailText As String

Private mMajorValue As Double, mMajorStep As Single
Private mMajorMin As Double, mMajorMax As Double
Private mMajorText As String, mMajorBarText As String
Private mMajorFloodColor As Long, mMajorBarTextColor As Long

Private mMinorValue As Double, mMinorStep As Single
Private mMinorMin As Double, mMinorMax As Double
Private mMinorText As String, mMinorBarText As String
Private mMinorFloodColor As Long, mMinorBarTextColor As Long

Private mProgressForm As Access.Form  ' ссылка на форму прогресса

Private Const c_strCustomProc = "[Event Procedure]"

Private mCancel As Boolean
'Public Event Cancel(ExitProgress As Long)
'========================
Public Sub Init( _
    Optional pCount As Byte = 1, _
    Optional pCaption As String = "", _
    Optional pMin = 0, Optional pMax = 100, Optional pStep = 1, Optional pText As String = "", _
    Optional pMinorMin = 0, Optional pMinorMax = 100, Optional pMinorStep = 1, Optional pMinorText As String = "", _
    Optional pFloodColor As Long = 8402176, Optional pBarTextColor As Long = vbWhite _
    )
' инициализация объектов
' pCount - количество отображаемых прогрессбаров (1 или 2)
' pCaption - заголовок формы
' pMin/pMinorMin - миниммальное значение для верхнего/нижнего прогрессбара
' pMax/pMinorMax - максимальное значение для верхнего/нижнего прогрессбара
' pText/pMinorText - подпись для верхнего/нижнего прогрессбара
' pFloodColor - цвет заполнения прогрессбаров
' pBarTextColor As Long - цвет текста прогрессбаров

    If Not IsNumeric(pMin) Or Not IsNumeric(pMax) Then GoTo HandleExit
'    Set mProgressForm = New Form_frmSERV_Progress
    DoCmd.OpenForm cstrHostForm, acNormal, , , , acHidden
    Set mProgressForm = Application.Forms(cstrHostForm)
' инициализация свойств
    mCancel = False
    Caption = pCaption
' 1 прогрессбар
    If mMajorMax < mMajorMin Then mMajorMax = mMajorMin
    mMajorMin = pMin: mMajorMax = pMax: If pStep <> 0 Then mMajorStep = pStep Else mMajorStep = (mMajorMax - mMajorMin) / 100
    If mMajorStep <= 0 Then mMajorStep = 1
    FloodColor = pFloodColor: BarTextColor = pBarTextColor
    Progress = pMin: Text = pText
    Select Case pCount
     Case 1
        mProgressCount = 1
     Case 2
' 2 прогрессбар
        mProgressCount = 2
    If Not IsNumeric(pMinorMin) Or Not IsNumeric(pMinorMax) Then GoTo HandleExit
        If mMinorMax < mMajorMin Then mMinorMax = mMajorMin
        mMinorMin = pMinorMin: mMinorMax = pMinorMax:: If pMinorStep <> 0 Then mMinorStep = pStep Else mMinorStep = (mMinorMax - mMinorMin) / 100
        If mMinorStep <= 0 Then mMinorStep = 1
        FloodColor(eMinor) = pFloodColor: BarTextColor(eMinor) = pBarTextColor
        Progress(eMinor) = pMinorMin: Text(eMinor) = pMinorText
     Case Else
        mCancel = True
        GoTo HandleExit
    End Select
HandleInit:
' инициализация формы
    mProgressForm.KeyPreview = True
    mProgressForm.OnKeyDown = c_strCustomProc
    mProgressForm.Init mProgressCount
HandleExit:
End Sub
Public Sub Show()
    mProgressForm.Visible = True
    mProgressForm.Repaint
End Sub
Public Sub Clear(Optional ProgressNum As eProgressType = eMajor)
' очищает состояние прогрессбара
    mCancel = False
' верхнюю сбрасываем если задано
    If ProgressNum = eMajor Then Progress = mMajorMin
' нижнюю сбрасываем всегда
    Progress(eMinor) = mMajorMin
End Sub
Public Sub Update( _
    Optional ProgressNum As eProgressType = eMajor, _
    Optional ProgressStep _
    )
' обновляет состояние процесса
Dim pStep As Single, pValue As Double
    DoEvents
    mCancel = mProgressForm.Canceled
    Select Case ProgressNum
    Case eMajor
        If IsMissing(ProgressStep) Then
            pStep = mMajorStep
        ElseIf Not IsNumeric(ProgressStep) Then
            pStep = mMajorStep
        Else
            pStep = ProgressStep
        End If
        ' если новый уровень верхнего счетчика - обнуляем нижний (м.б. и не надо)
        If mProgressCount > 1 Then Progress(eMinor) = mMinorMin
        ' устанавливаем значение верхнего прогресса
        Progress(eMajor) = mMajorValue + pStep
    Case eMinor
        If IsMissing(ProgressStep) Then
            pStep = mMinorStep
        ElseIf Not IsNumeric(ProgressStep) Then
            pStep = mMinorStep
        Else
            pStep = ProgressStep
        End If
        ' устанавливаем значение нижнего прогресса
        Me.Progress(eMinor) = mMinorValue + pStep
    End Select
End Sub
'=============================
Private Sub Class_Initialize()
    mProgressCount = 1
    If p_IsFormOpen(cstrHostForm) Then DoCmd.Close acForm, cstrHostForm, acSaveNo
End Sub
Private Sub Class_Terminate()
    Set mProgressForm = Nothing
    If p_IsFormOpen(cstrHostForm) Then DoCmd.Close acForm, cstrHostForm, acSaveNo
End Sub
'=============================
' количество прогрессбаров
Public Property Let ProgressCount(rData As Byte): mProgressForm.Init rData: mProgressCount = rData: End Property
Public Property Get ProgressCount() As Byte: ProgressCount = mProgressCount: End Property
' заголовок формы процессов
Public Property Let Caption(rData As String):  mProgressForm.lblCaptionText.Caption = rData: End Property
Public Property Get Caption() As String: Caption = mProgressForm.lblCaptionText.Caption: End Property
' детальная информация об этапе процесса
Public Property Let Detail(rData As String):  mProgressForm.lblDetailText.Caption = rData: End Property
Public Property Get Detail() As String: Detail = mProgressForm.lblDetailText.Caption: End Property
' информация в заголовке прогрессбара
Public Property Let Text(Optional ProgressNum As eProgressType = eMajor, rData As String)
    Select Case ProgressNum
    Case eMajor: mMajorText = rData: mProgressForm.lblMajorText.Caption = rData
    Case eMinor: mMinorText = rData: mProgressForm.lblMinorText.Caption = rData
    End Select
    mProgressForm.Repaint
End Property
' текст выводимый в прогрессбар
Public Property Let BarText(Optional ProgressNum As eProgressType = eMajor, rData As String)
    Select Case ProgressNum
    Case eMajor: mMajorBarText = rData: mProgressForm.lblMajorPercent.Caption = rData
    Case eMinor: mMinorBarText = rData: mProgressForm.lblMinorPercent.Caption = rData
    End Select
    mProgressForm.Repaint
End Property
' цвет текста прогрессбара
Public Property Let BarTextColor(Optional ProgressNum As eProgressType = eMajor, rData As Long)
    Select Case ProgressNum
    Case eMajor: mProgressForm.lblMajorPercent.ForeColor = rData
    Case eMinor: mProgressForm.lblMinorPercent.ForeColor = rData
    End Select
    mProgressForm.Repaint
End Property
' цвет прогрессбара
Public Property Let FloodColor(Optional ProgressNum As eProgressType = eMajor, rData As Long)
    Select Case ProgressNum
    Case eMajor: mProgressForm.lblMajorFlood.BackColor = rData
    Case eMinor: mProgressForm.lblMinorFlood.BackColor = rData
    End Select
    mProgressForm.Repaint
End Property
' текущее значение прогресса
Public Property Get Progress(Optional ProgressNum As eProgressType = eMajor) As Double
    If ProgressNum = eMajor Then
        Progress = mMajorValue
    ElseIf ProgressNum = eMinor Then
        Progress = mMinorValue
    End If
End Property
Public Property Let Progress( _
    Optional ProgressNum As eProgressType = eMajor, _
    rData As Double)
Dim pPercent As Single
Dim tmp, msk As String
' прямо задаем текущее значение счетчика итераций
    DoEvents
    mCancel = mProgressForm.Canceled
    Select Case ProgressNum
    Case eMajor
        mMajorValue = rData
        If mMajorValue > mMajorMax Then mMajorValue = mMajorMax
        If mMajorValue < mMajorMin Then mMajorValue = mMajorMin
        If mMajorMax <> mMajorMin Then pPercent = (mMajorValue - mMajorMin) / (mMajorMax - mMajorMin) Else pPercent = 1
        With mProgressForm
            .lblMajorFlood.Width = pPercent * .FloodWidth
            If Len(mMajorText) = 0 Then _
                .lblMajorText.Caption = "Текущее значение: " & _
                p_FormatedNum(mMajorValue) & _
                ", в диапазоне от " & mMajorMin & " до " & mMajorMax
            If Len(mMajorBarText) = 0 Then _
                .lblMajorPercent.Caption = Format$(pPercent, "Percent")
        End With
    Case eMinor
        mMinorValue = rData
        If mMinorValue > mMinorMax Then mMinorValue = mMinorMax
        If mMinorValue < mMinorMin Then mMinorValue = mMinorMin
        If mMinorMax <> mMinorMin Then pPercent = (mMinorValue - mMinorMin) / (mMinorMax - mMinorMin) Else pPercent = 1
        With mProgressForm
        ' обрабатываем старший прогресс
            If Len(mMajorText) = 0 Or Len(mMajorBarText) = 0 Then _
                Me.Progress(eMajor) = mMajorValue
        ' обрабатываем младший прогресс
            .lblMinorFlood.Width = pPercent * .FloodWidth
            If Len(mMinorText) = 0 Then _
                .lblMinorText.Caption = "Текущее значение: " & _
                p_FormatedNum(mMinorValue) & _
                ", в диапазоне от " & mMinorMin & " до " & mMinorMax
            If Len(mMinorBarText) = 0 Then _
                .lblMinorPercent.Caption = Format$(pPercent, "Percent")
        End With
    End Select
    If Not mProgressForm.Visible Then Show
    mProgressForm.Repaint
End Property
Public Property Let ProgressMin(Optional ProgressNum As eProgressType = eMajor, rData As Double)
' минимальное значение прогресса
' если изменилось минимальное - пересчитываем шаг
    Select Case ProgressNum
    Case eMajor: mMajorMin = rData ': mMajorStep = (mMajorMax - mMajorMin)/100
    Case eMinor: mMinorMin = rData ': mMinorStep = (mMinorMax - mMinorMin)/100
    End Select
End Property
Public Property Get ProgressMin(Optional ProgressNum As eProgressType = eMajor) As Double
    Select Case ProgressNum
    Case eMajor: ProgressMin = mMajorMin
    Case eMinor: ProgressMin = mMinorMin
    End Select
End Property
Public Property Let ProgressMax(Optional ProgressNum As eProgressType = eMajor, rData As Double)
' максимальное значение прогресса
' если изменилось максимальное - пересчитываем шаг
    Select Case ProgressNum
    Case eMajor: mMajorMax = rData ': mMajorStep = (mMajorMax - mMajorMin) / 100
    Case eMinor: mMinorMax = rData ': mMinorStep = (mMinorMax - mMinorMin) / 100
    End Select
End Property
Public Property Get ProgressMax(Optional ProgressNum As eProgressType = eMajor) As Double
    Select Case ProgressNum
    Case eMajor: ProgressMax = mMajorMax
    Case eMinor: ProgressMax = mMinorMax
    End Select
End Property
Public Property Get ProgressStep(Optional ProgressNum As eProgressType = eMajor) As Single
' максимальное значение прогресса
    Select Case ProgressNum
    Case eMajor: ProgressStep = mMajorStep
    Case eMinor: ProgressStep = mMinorStep
    End Select
End Property
Public Property Let ProgressStep(Optional ProgressNum As eProgressType = eMajor, rData As Single)
    If rData <= 0 Then rData = 1
    Select Case ProgressNum
    Case eMajor: mMajorStep = rData 's = (mMajorMax-mMajorMin)/100
    Case eMinor: mMinorStep = rData 's = (mMinorMax-mMinorMin)/100
    End Select
End Property
' флаг прерывания процесса
Public Property Let Canceled(rData As Boolean): mProgressForm.Canceled = rData: mCancel = rData: End Property
Public Property Get Canceled() As Boolean: Canceled = mCancel: End Property
'=============================
Private Function p_IsFormOpen(sFormName As String) As Boolean
Dim tmp As String
    On Error Resume Next
    tmp = Application.Forms(sFormName).Caption
    p_IsFormOpen = Not CBool(Err.Number)
End Function
Private Function p_FormatedNum(rNumber As Double) As String
Dim fPart As Double, dig As Byte, tmp As Byte
Dim msk As String
' для целых чисел
    msk = "#,##0"
    fPart = Abs(rNumber - Fix(rNumber))
    If fPart > 0 Then
' есть дробная часть
        dig = 1 'dig = 2 ' два разряда после запятой
        Do While fPart * 10 ^ dig Mod 10 ^ dig = 0: dig = dig + 1: Loop
        msk = msk & "." & String(dig - 1, "#") & "0"
    End If
    p_FormatedNum = Format$(rNumber, msk)
End Function




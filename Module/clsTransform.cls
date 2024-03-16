VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================
Private Const c_strModule As String = "clsTransform"
'=========================
' Описание      : Класс для расчёта простейших трансформаций (поворотов) прямоугольной области
' Версия        : 0.1.2.451984669
' Дата          : 29.09.2023 11:12:20
' Автор         : Кашкин Р.В, (KashRus@gmail.com)
' Примечание    : используется в modPictureData для расчёта позиций на контроле
' v.0.1.2       : 29.09.2023 - основной функционал работает, остальное сильно экспериментально
'=========================
Private Const Pi As Double = 3.14159265358979 '3.14159265358979
Private mAngle As Single             ' угол наклона (в градусах)
Private Type tTrans                  ' объявление типа для матрицы трансформации (на плоскости)
    a As Single: b As Single: p As Single  ' p=0
    c As Single: d As Single: q As Single  ' q=0
    m As Single: n As Single: s As Single  ' s=1
End Type
Private mxTrans As tTrans            ' матрица трансформации
' a,b,c,d  - Поворот
' a,d      - Масштабирование
' a,d      - Отражение (отн-но начала координат, OY, OX)
' c,b      - Отражение (отн-но y = x, y = -x)
' a,b,c,d  - Сдвиг
' m,n      - Перенос
' p,q      - Проекции (не используется)
' s        - Масштабирование глобальное (не используется)
Private CosA As Single, SinA As Single, Cos2A As Single
Private Sub Class_Initialize()
    mOutType = 0
    With mxTrans ' матрица трансформации
        .a = 1: .b = 0: .p = 0
        .c = 0: .d = 1: .q = 0
        .m = 0: .n = 0: .s = 1
    End With
End Sub
Public Property Let Angle(rData As Single)
' угол наклона области в градусах (для экранных координат)
    mAngle = rData - Fix(rData / 360) * 360: If mAngle < 0 Then mAngle = 360 + mAngle ' нормализуем угол
' задаём матрицу трансформации для поворота
    Select Case mAngle
    Case 0:     CosA = 1: SinA = 0: Cos2A = 1
    Case 90:    CosA = 0: SinA = 1: Cos2A = -1
    Case 180:   CosA = -1: SinA = 0: Cos2A = 1
    Case 270:   CosA = 0: SinA = -1: Cos2A = -1
    Case Else:  Dim radA As Single: radA = Pi * mAngle / 180: SinA = Sin(radA): CosA = Cos(radA): Cos2A = Cos(2 * radA)
    End Select
    With mxTrans
        .a = CosA: .c = SinA:  '.m =
        .b = -SinA: .d = CosA: '.n =
    End With
End Property
Public Property Get Angle() As Single: Angle = mAngle: End Property
Public Property Let ShiftX(rData As Single): mxTrans.m = rData: End Property ' Перенос (смещение) по оси 0X
Public Property Get ShiftX() As Single: ShiftX = mxTrans.m: End Property
Public Property Let ShiftY(rData As Single): mxTrans.n = rData: End Property ' Перенос (смещение) по оси 0Y
Public Property Get ShiftY() As Single: ShiftY = mxTrans.n: End Property
'--------------
' Функции модуля
'--------------
Public Function Transform( _
    ByVal x0 As Single, ByVal y0 As Single, _
    ByRef X1 As Single, ByRef Y1 As Single, _
    Optional cX0 As Single = 1, Optional cY0 As Single = 1, _
    Optional cX1 As Single = 1, Optional cY1 As Single = 1)
' преобразование координат по матрице
' -------------
' X0, Y0    - координаты до трансформации (входные)
' X1, Y1    - координаты после трансформации (выходные)
' сX0, сY0  - база для расчёта в случае если заданы пропорциональные координат (до трансформации)
' сX1, сY1  - база для расчёта в случае если нужен результат в пропорциональных координатах (после трансформации)
' -------------
' если входные координаты (X0,Y0) заданы в пропорциональных величинах относительно размеров области до поворота (mX0,mY0)
    x0 = x0 * cX0: y0 = y0 * cY0
    With mxTrans '
' прямое преобразование координат по матрице
        X1 = .a * x0 + .c * y0 + .m
        Y1 = .b * x0 + .d * y0 + .n
    End With
'' если нужен результат относительно левого верхнего угла области, занимаемой повёрнутым объектом нужна поправка на dx,dy
'    If mOutType = 1 Then X1 = X1 + mdX: Y1 = Y1 + mdY
' если выходные координаты (X1,Y1) нужны в пропорциональных величинах относительно размеров области после поворота (mX1,mY1)
    X1 = X1 / cX1: Y1 = Y1 / cY1
End Function
Public Function Inverse( _
    ByVal X1 As Single, ByVal Y1 As Single, _
    ByRef x0 As Single, ByRef y0 As Single, _
    Optional cX1 As Single = 1, Optional cY1 As Single = 1, _
    Optional cX0 As Single = 1, Optional cY0 As Single = 1)
' обратное преобразование координат по матрице
' -------------
' X1, Y1    - координаты после трансформации (входные)
' X0, Y0    - координаты до трансформации (выходные)
' сX1, сY1  - база для расчёта в случае если заданы пропорциональные координат (после трансформации)
' сX0, сY0  - база для расчёта в случае если нужен результат в пропорциональных координатах (до трансформации)
' -------------
' если входные координаты (X1,Y1) заданы в пропорциональных величинах относительно размеров области после поворота (mX1,mY1)
    X1 = X1 * cX1: Y1 = Y1 * cY1
    With mxTrans '
' обратное преобразование координат по матрице
        x0 = .a * X1 + -.c * Y1 + -.m
        y0 = -.b * X1 + .d * Y1 + -.n
    End With
' если выходные координаты (X0,Y0) нужны в пропорциональных величинах относительно размеров области до поворота (mX0,mY0)
    If cX0 <> 0 Then x0 = x0 / cX0 Else x0 = 0
    If cY0 <> 0 Then y0 = y0 / cY0 Else y0 = 0
End Function
Public Function TransformSize( _
    W0 As Single, H0 As Single, _
    W1 As Single, H1 As Single)
' по размерам области до трансформации определяем размеры области в которую она вписана после трансформации
    With mxTrans '
' прямое преобразование размеров области (после поворота)
        W1 = Abs(.a) * W0 + Abs(.c) * H0 '+ Abs(.m)
        H1 = Abs(.b) * W0 + Abs(.d) * H0 '+ Abs(.n)
    End With
End Function
Public Function InverseSize( _
    W1 As Single, H1 As Single, _
    W0 As Single, H0 As Single)
' по размерам области в которую вписана область после трансформации определяем размеры области до трансформации
    If Cos2A = 0 Then Exit Function
' если угол кратен 45 градусам (Cos2A=0) => w1=h1 т.к. область после поворота на такой угол д.б. квадратной,
' но в квадрат под заданным углом можно вписать бесконечное множество прямоугольников со сторонами: w0-h0 = w1/CosA
' можно вернуть нули (результат неопределён) или максимальную возможную область при таких условия - квадрат: w0 = h0 = 0,5*w1/CosA
    With mxTrans '
' обратное преобразование размеров области (до поворота)
        W0 = W1 * Abs(.a) * 1 / Cos2A + H1 * Abs(.c) * -1 / Cos2A '+ -1 * Abs(.m)
        H0 = W1 * Abs(.b) * -1 / Cos2A + H1 * Abs(.d) * 1 / Cos2A '+ -1 * Abs(.n)
    End With
End Function
Public Function GetDelta( _
    cX0 As Single, cY0 As Single, _
    dX0 As Single, dY0 As Single)
' получаем поправки для перехода от координат в системе исходного объекта к координатам заданным относительно области занимаемой объектом после поворота

' рассчитываем по углам - т.к. мы знаем куда какой угол будет смотреть после поворота
    With mxTrans
    Select Case mAngle
    Case 0 To 90:       dX0 = -(.m):                        dY0 = -(.b * cX0 + .n)              ' dX0 ->  Ax; dY0 -> -By
    Case 90 To 180:     dX0 = -(.a * cX0 + .m):             dY0 = -(.b * cX0 + .d * cY0 + .n)   ' dX0 -> -Bx; dY0 -> -Cy
    Case 180 To 270:    dX0 = -(.a * cX0 + .c * cY0 + .m):  dY0 = -(.d * cY0 + .n)              ' dX0 -> -Cx; dY0 -> -Dy
    Case 270 To 360:    dX0 = -(.c * cY0 + .m):             dY0 = -(.n)                         ' dX0 -> -Dx; dY0 ->  Ay
    End Select
    End With  'mxTrans
End Function
' -------------
' в принципе не нужны - сводятся к Transform/Inverse
' -------------
Public Function TransformPoints(aPoints, _
    Optional cX0 As Single = 1, Optional cY0 As Single = 1, _
    Optional cX1 As Single = 1, Optional cY1 As Single = 1 _
    ) As Variant
' позиция точек на области после поворота
' -------------
' aPoints двумерный массив позиций точек относительно левого-верхнего угла области
' -------------
    On Error GoTo HandleError
Dim j As Long, jMax As Long: j = LBound(aPoints, 2): jMax = UBound(aPoints, 2): If (jMax - j + 1) <> 2 Then Err.Raise vbObjectError + 512
Dim i As Long, iMax As Long: i = LBound(aPoints, 1): iMax = UBound(aPoints, 1)
Dim aResult() As Single: ReDim aResult(i To iMax, j To jMax)
    For i = i To iMax
        Call Transform(aPoints(i, j), aPoints(i, j + 1), aResult(i, j), aResult(i, j + 1), cX0, cY0, cX1, cY1)
    Next
HandleExit:  TransformPoints = aResult: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function InversePoints(aPoints, _
    Optional cX1 As Single = 1, Optional cY1 As Single = 1, _
    Optional cX0 As Single = 1, Optional cY0 As Single = 1 _
    ) As Variant
' позиция точек на области до поворота
' -------------
' aPoints двумерный массив позиций точек относительно левого-верхнего угла области
' -------------
    On Error GoTo HandleError
Dim j As Long, jMax As Long: j = LBound(aPoints, 2): jMax = UBound(aPoints, 2): If (jMax - j + 1) <> 2 Then Err.Raise vbObjectError + 512
Dim i As Long, iMax As Long: i = LBound(aPoints, 1): iMax = UBound(aPoints, 1)
Dim aResult() As Single: ReDim aResult(i To iMax, j To jMax)
    For i = i To iMax
        Call Inverse(aPoints(i, j), aPoints(i, j + 1), aResult(i, j), aResult(i, j + 1), cX1, cY1, cX0, cY0)
    Next
HandleExit:  InversePoints = aResult: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function TransformRect( _
    cX0 As Single, cY0 As Single, _
    Optional cX1 As Single = 1, Optional cY1 As Single = 1 _
    ) As Variant

' позиция углов исходной области на области после поворота
' -------------
' (возвращаемое) - массив позиций углов исходной области на области проекции (после поворота)
' -------------
Dim aPoints(1 To 4, 1 To 2) As Single
    i = LBound(aPoints)
    aPoints(i, 1) = 0: aPoints(i, 2) = 0: i = i + 1     ' (A) Left-Top corner
    aPoints(i, 1) = 1: aPoints(i, 2) = 0: i = i + 1     ' (B) Right-Top corner
    aPoints(i, 1) = 1: aPoints(i, 2) = 1: i = i + 1     ' (C) Right-Bottom corner
    aPoints(i, 1) = 0: aPoints(i, 2) = 1 ': i = i + 1   ' (D) Left-Bottom corner
    TransformRect = TransformPoints(aPoints, cX0, cY0, cX1, cY1)
End Function


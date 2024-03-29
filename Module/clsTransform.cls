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
' ��������      : ����� ��� ������� ���������� ������������� (���������) ������������� �������
' ������        : 0.1.2.451984669
' ����          : 29.09.2023 11:12:20
' �����         : ������ �.�, (KashRus@gmail.com)
' ����������    : ������������ � modPictureData ��� ������� ������� �� ��������
' v.0.1.2       : 29.09.2023 - �������� ���������� ��������, ��������� ������ ����������������
'=========================
Private Const Pi As Double = 3.14159265358979 '3.14159265358979
Private mAngle As Single             ' ���� ������� (� ��������)
Private Type tTrans                  ' ���������� ���� ��� ������� ������������� (�� ���������)
    a As Single: b As Single: p As Single  ' p=0
    c As Single: d As Single: q As Single  ' q=0
    m As Single: n As Single: s As Single  ' s=1
End Type
Private mxTrans As tTrans            ' ������� �������������
' a,b,c,d  - �������
' a,d      - ���������������
' a,d      - ��������� (���-�� ������ ���������, OY, OX)
' c,b      - ��������� (���-�� y = x, y = -x)
' a,b,c,d  - �����
' m,n      - �������
' p,q      - �������� (�� ������������)
' s        - ��������������� ���������� (�� ������������)
Private CosA As Single, SinA As Single, Cos2A As Single
Private Sub Class_Initialize()
    mOutType = 0
    With mxTrans ' ������� �������������
        .a = 1: .b = 0: .p = 0
        .c = 0: .d = 1: .q = 0
        .m = 0: .n = 0: .s = 1
    End With
End Sub
Public Property Let Angle(rData As Single)
' ���� ������� ������� � �������� (��� �������� ���������)
    mAngle = rData - Fix(rData / 360) * 360: If mAngle < 0 Then mAngle = 360 + mAngle ' ����������� ����
' ����� ������� ������������� ��� ��������
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
Public Property Let ShiftX(rData As Single): mxTrans.m = rData: End Property ' ������� (��������) �� ��� 0X
Public Property Get ShiftX() As Single: ShiftX = mxTrans.m: End Property
Public Property Let ShiftY(rData As Single): mxTrans.n = rData: End Property ' ������� (��������) �� ��� 0Y
Public Property Get ShiftY() As Single: ShiftY = mxTrans.n: End Property
'--------------
' ������� ������
'--------------
Public Function Transform( _
    ByVal x0 As Single, ByVal y0 As Single, _
    ByRef X1 As Single, ByRef Y1 As Single, _
    Optional cX0 As Single = 1, Optional cY0 As Single = 1, _
    Optional cX1 As Single = 1, Optional cY1 As Single = 1)
' �������������� ��������� �� �������
' -------------
' X0, Y0    - ���������� �� ������������� (�������)
' X1, Y1    - ���������� ����� ������������� (��������)
' �X0, �Y0  - ���� ��� ������� � ������ ���� ������ ���������������� ��������� (�� �������������)
' �X1, �Y1  - ���� ��� ������� � ������ ���� ����� ��������� � ���������������� ����������� (����� �������������)
' -------------
' ���� ������� ���������� (X0,Y0) ������ � ���������������� ��������� ������������ �������� ������� �� �������� (mX0,mY0)
    x0 = x0 * cX0: y0 = y0 * cY0
    With mxTrans '
' ������ �������������� ��������� �� �������
        X1 = .a * x0 + .c * y0 + .m
        Y1 = .b * x0 + .d * y0 + .n
    End With
'' ���� ����� ��������� ������������ ������ �������� ���� �������, ���������� ��������� �������� ����� �������� �� dx,dy
'    If mOutType = 1 Then X1 = X1 + mdX: Y1 = Y1 + mdY
' ���� �������� ���������� (X1,Y1) ����� � ���������������� ��������� ������������ �������� ������� ����� �������� (mX1,mY1)
    X1 = X1 / cX1: Y1 = Y1 / cY1
End Function
Public Function Inverse( _
    ByVal X1 As Single, ByVal Y1 As Single, _
    ByRef x0 As Single, ByRef y0 As Single, _
    Optional cX1 As Single = 1, Optional cY1 As Single = 1, _
    Optional cX0 As Single = 1, Optional cY0 As Single = 1)
' �������� �������������� ��������� �� �������
' -------------
' X1, Y1    - ���������� ����� ������������� (�������)
' X0, Y0    - ���������� �� ������������� (��������)
' �X1, �Y1  - ���� ��� ������� � ������ ���� ������ ���������������� ��������� (����� �������������)
' �X0, �Y0  - ���� ��� ������� � ������ ���� ����� ��������� � ���������������� ����������� (�� �������������)
' -------------
' ���� ������� ���������� (X1,Y1) ������ � ���������������� ��������� ������������ �������� ������� ����� �������� (mX1,mY1)
    X1 = X1 * cX1: Y1 = Y1 * cY1
    With mxTrans '
' �������� �������������� ��������� �� �������
        x0 = .a * X1 + -.c * Y1 + -.m
        y0 = -.b * X1 + .d * Y1 + -.n
    End With
' ���� �������� ���������� (X0,Y0) ����� � ���������������� ��������� ������������ �������� ������� �� �������� (mX0,mY0)
    If cX0 <> 0 Then x0 = x0 / cX0 Else x0 = 0
    If cY0 <> 0 Then y0 = y0 / cY0 Else y0 = 0
End Function
Public Function TransformSize( _
    W0 As Single, H0 As Single, _
    W1 As Single, H1 As Single)
' �� �������� ������� �� ������������� ���������� ������� ������� � ������� ��� ������� ����� �������������
    With mxTrans '
' ������ �������������� �������� ������� (����� ��������)
        W1 = Abs(.a) * W0 + Abs(.c) * H0 '+ Abs(.m)
        H1 = Abs(.b) * W0 + Abs(.d) * H0 '+ Abs(.n)
    End With
End Function
Public Function InverseSize( _
    W1 As Single, H1 As Single, _
    W0 As Single, H0 As Single)
' �� �������� ������� � ������� ������� ������� ����� ������������� ���������� ������� ������� �� �������������
    If Cos2A = 0 Then Exit Function
' ���� ���� ������ 45 �������� (Cos2A=0) => w1=h1 �.�. ������� ����� �������� �� ����� ���� �.�. ����������,
' �� � ������� ��� �������� ����� ����� ������� ����������� ��������� ��������������� �� ���������: w0-h0 = w1/CosA
' ����� ������� ���� (��������� ����������) ��� ������������ ��������� ������� ��� ����� ������� - �������: w0 = h0 = 0,5*w1/CosA
    With mxTrans '
' �������� �������������� �������� ������� (�� ��������)
        W0 = W1 * Abs(.a) * 1 / Cos2A + H1 * Abs(.c) * -1 / Cos2A '+ -1 * Abs(.m)
        H0 = W1 * Abs(.b) * -1 / Cos2A + H1 * Abs(.d) * 1 / Cos2A '+ -1 * Abs(.n)
    End With
End Function
Public Function GetDelta( _
    cX0 As Single, cY0 As Single, _
    dX0 As Single, dY0 As Single)
' �������� �������� ��� �������� �� ��������� � ������� ��������� ������� � ����������� �������� ������������ ������� ���������� �������� ����� ��������

' ������������ �� ����� - �.�. �� ����� ���� ����� ���� ����� �������� ����� ��������
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
' � �������� �� ����� - �������� � Transform/Inverse
' -------------
Public Function TransformPoints(aPoints, _
    Optional cX0 As Single = 1, Optional cY0 As Single = 1, _
    Optional cX1 As Single = 1, Optional cY1 As Single = 1 _
    ) As Variant
' ������� ����� �� ������� ����� ��������
' -------------
' aPoints ��������� ������ ������� ����� ������������ ������-�������� ���� �������
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
' ������� ����� �� ������� �� ��������
' -------------
' aPoints ��������� ������ ������� ����� ������������ ������-�������� ���� �������
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

' ������� ����� �������� ������� �� ������� ����� ��������
' -------------
' (������������) - ������ ������� ����� �������� ������� �� ������� �������� (����� ��������)
' -------------
Dim aPoints(1 To 4, 1 To 2) As Single
    i = LBound(aPoints)
    aPoints(i, 1) = 0: aPoints(i, 2) = 0: i = i + 1     ' (A) Left-Top corner
    aPoints(i, 1) = 1: aPoints(i, 2) = 0: i = i + 1     ' (B) Right-Top corner
    aPoints(i, 1) = 1: aPoints(i, 2) = 1: i = i + 1     ' (C) Right-Bottom corner
    aPoints(i, 1) = 0: aPoints(i, 2) = 1 ': i = i + 1   ' (D) Left-Bottom corner
    TransformRect = TransformPoints(aPoints, cX0, cY0, cX1, cY1)
End Function


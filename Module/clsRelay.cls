VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'������������ �������������� ������ - �������� �������� ��������� ������
'�� ������������ ���������, ����� ��� �����. ����� �������������� ���������
'����� ������� ������������� ���������� ��������� ������ ���������� �������
'(�� ��������� � ���������� ���������) ������. ����� �������, ��� �����
'(Set ������������������ = Nothing) ��� ������� (����� ����������
'������������������ �� ������� ���������) ���������� �������� ������ �� ����
'����������� ���������� ������� Class_Terminate ���������� ���������, �������
'���������� ������������ ��������.
Public Event GetParent(PARENT As Object)
Public Property Get PARENT() As Object
Dim oParent As Object
    RaiseEvent GetParent(oParent)
    Set PARENT = oParent
End Property

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
'≈динственное предназначение класса - передать элементу коллекции ссылку
'на родительскую коллекцию, когда это нужно. “акое посредничество позвол€ет
'иметь счЄтчик использовани€ экземпл€ра коллекции равным количеству внешних
'(по отношении к экземпл€ру коллекции) ссылок. “аким образом, при €вном
'(Set Ёкземпл€р оллекции = Nothing) или не€вном (выход переменной
'Ёкземпл€р оллекции за пределы видимости) уменьшении счЄтчика ссылок до нул€
'срабатывает обработчик событи€ Class_Terminate экземпл€ра коллекции, который
'производит освобождение ресурсов.
Public Event GetParent(PARENT As Object)
Public Property Get PARENT() As Object
Dim oParent As Object
    RaiseEvent GetParent(oParent)
    Set PARENT = oParent
End Property

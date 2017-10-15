VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFTypesArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'PROPERTIES

Public Property Get Data() As Long
    '
End Property

Public Property Get ElementSize() As Long
    '
End Property

Public Property Get Pointer() As Long
    '
End Property

'METHODS

Public Function Add(ByRef Item As Variant) As IFTypesArray
    '
End Function

Public Function Parse(ByRef Value As Variant) As IFTypesArray
    '
End Function

Public Function Resize(ByVal NewUpperBound As Long) As IFTypesArray
    '
End Function

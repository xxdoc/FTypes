VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'VARIABLES

Private p_Array As TSAFEARRAY
Private p_Value As Variant

'PROPERTIES

Public Property Get Dimensions() As Long
    
    Dimensions = p_Array.iDims

End Property

Public Property Get ElementSize() As Long
    
    ElementSize = p_Array.lElementSize

End Property

Public Property Get Pointer() As Long
    
    Pointer = p_Array.lData

End Property

Public Property Get Type_() As VbVarType
    
    Type_ = p_Array.lVarType

End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    If p_Array.lVarType Then Value = p_Value

End Property

Public Property Let Value(ByRef NewValue As Variant)
    
    If VarType(NewValue) = vbArray Then
        
        p_Value = NewValue
        
        RArrPt p_Value, p_Array
    
    End If

End Property

'ROUTINES

Friend Sub Initialize(ByVal ArrayType As VbVarType, ByVal Size As Long)
    
    Select Case ArrayType
        Case vbBoolean: ReDim p_Value(Size) As Boolean
        Case vbByte: ReDim p_Value(Size) As Byte
        Case vbCurrency: ReDim p_Value(Size) As Currency
        Case vbDouble, vbDecimal: ReDim p_Value(Size) As Double
        Case vbInteger, vbError: ReDim p_Value(Size) As Integer
        Case vbLong: ReDim p_Value(Size) As Long
        Case vbSingle: ReDim p_Value(Size) As Single
        Case vbString, vbDate: ReDim p_Value(Size) As String
        Case Else: Exit Sub
    End Select
    
    RArrPt p_Value, p_Array

End Sub

Public Function Clone() As ArrayEx
    
    Set Clone = New ArrayEx
    
    If p_Array.lVarType Then Clone.Value = p_Value

End Function

Public Function Parse(ByRef Value As Variant) As ArrayEx
    
    Set Parse = Me
    
    Initialize VarType(Value), 0&
    
    p_Value(0&) = Value

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ByteEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'VARIABLES

Private p_Value As Byte

'PROPERTIES

Public Property Get Max() As Byte
    
    Max = B_MX

End Property

Public Property Get Min() As Byte
    
    Min = 0

End Property

Public Property Get Value() As Byte
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    Value = p_Value

End Property

Public Property Let Value(ByVal NewValue As Byte)
    
    p_Value = NewValue

End Property

'ROUTINES

Public Function Clone() As ByteEx
    
    Set Clone = New ByteEx
    
    Clone.Value = p_Value

End Function

Public Function Parse(ByRef Value As Variant) As ByteEx
    
    Dim d As Double
    Dim x As Long
    
    Set Parse = Me
    
    x = VarType(Value)
    
    Select Case x
        
        Case 11&, 17& 'BOOLEAN, BYTE
            
            p_Value = Value
        
        Case 2& To 8&, 14& 'INTEGER, LONG, SINGLE, DOUBLE, CURRENCY, DATE, STRING, DECIMAL
            
            If x = 8& Then d = ToNumber(Value)
            
            If d > 0 Then
                If d < B_MX Then p_Value = d Else p_Value = B_MX
            End If
        
        Case Else
            
            p_Value = 0
    
    End Select

End Function

Public Function ToBoolean() As Boolean
    
    ToBoolean = p_Value

End Function

Public Function ToDouble() As DoubleEx
    
    Set ToDouble = New DoubleEx
    
    ToDouble.Value = p_Value

End Function

Public Function ToInteger() As IntegerEx
    
    Set ToInteger = New IntegerEx
    
    ToInteger.Value = p_Value

End Function

Public Function ToLong() As LongEx
    
    Set ToLong = New LongEx
    
    ToLong.Value = p_Value

End Function

Public Function ToString() As StringEx
    
    Set ToString = New StringEx
    
    ToString.Value = p_Value

End Function

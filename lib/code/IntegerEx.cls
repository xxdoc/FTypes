VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntegerEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'VARIABLES

Private p_Value As Integer

'PROPERTIES

Public Property Get Max() As Integer
    
    Max = I_MX

End Property

Public Property Get Min() As Integer
    
    Min = I_MN

End Property

Public Property Get Value() As Integer
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    Value = p_Value

End Property

Public Property Let Value(ByVal NewValue As Integer)
    
    p_Value = NewValue

End Property

Public Property Get WordHigh() As ByteEx
    
    Set WordHigh = New ByteEx
    
    WordHigh.Value = (p_Value And &HFF00&) \ 256

End Property

Public Property Get WordLow() As ByteEx
    
    Set WordLow = New ByteEx
    
    WordLow.Value = p_Value And B_MX

End Property

'ROUTINES

Public Function Clone() As IntegerEx
    
    Set Clone = New IntegerEx
    
    Clone.Value = p_Value

End Function

Public Function Parse(ByRef Value As Variant) As IntegerEx
    
    Dim d As Double
    Dim x As Long
    
    Set Parse = Me
    
    x = VarType(Value)
    
    Select Case x
        
        Case vbInteger, vbBoolean, vbByte
            
            p_Value = Value
        
        Case vbLong To vbString, vbDecimal
            
            If x = 8& Then d = ToNumber(Value)
            
            If d > I_MN Then
                
                If d < I_MX Then p_Value = d Else p_Value = I_MX
            
            Else
                p_Value = I_MN
            End If
        
        Case Else
            
            p_Value = 0
    
    End Select

End Function

Public Function ToBoolean() As Boolean
    
    ToBoolean = p_Value

End Function

Public Function ToByte() As ByteEx
    
    Set ToByte = New ByteEx
    
    If p_Value > 0 Then
        
        If p_Value < B_MX Then ToByte.Value = p_Value Else ToByte.Value = B_MX
    
    End If

End Function

Public Function ToDouble() As DoubleEx
    
    Set ToDouble = New DoubleEx
    
    ToDouble.Value = p_Value

End Function

Public Function ToLong() As LongEx
    
    Set ToLong = New LongEx
    
    ToLong.Value = p_Value

End Function

Public Function ToString() As StringEx
    
    Set ToString = New StringEx
    
    ToString.Value = p_Value

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LongEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'VARIABLES

Private p_Value As Long

'PROPERTIES

Public Property Get Max() As Long
    
    Max = L_MX

End Property

Public Property Get Min() As Long
    
    Min = L_MN

End Property

Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    Value = p_Value

End Property

Public Property Let Value(ByVal NewValue As Long)
    
    p_Value = NewValue

End Property

Public Property Get WordHigh() As IntegerEx
    
    Set WordHigh = New IntegerEx
    
    WordHigh.Value = ((p_Value And &H7FFF0000) \ &H10000) Or (&H8000 And (p_Value < 0))

End Property

Public Property Get WordLow() As IntegerEx
    
    Set WordLow = New IntegerEx
    
    WordLow.Value = (p_Value And &H7FFF) Or (&H8000 And ((p_Value And &H8000&) = &H8000&))

End Property

'ROUTINES

Public Function Clone() As LongEx
    
    Set Clone = New LongEx
    
    Clone.Value = p_Value

End Function

Public Function Parse(ByRef Value As Variant) As LongEx
    
    Dim d As Double
    Dim x As Long
    
    Set Parse = Me
    
    x = VarType(Value)
    
    Select Case x
        
        Case vbInteger, vbLong, vbDate, vbBoolean, vbByte
            
            p_Value = Value
        
        Case vbSingle To vbCurrency, vbString, vbDecimal
            
            If x = 8& Then d = ToNumber(Value)
            
            If d > L_MN Then
                
                If d < L_MX Then p_Value = d Else p_Value = L_MX
            
            Else
                p_Value = L_MN
            End If
        
        Case Else
            
            p_Value = 0&
    
    End Select

End Function

Public Function ToBoolean() As Boolean
    
    ToBoolean = p_Value

End Function

Public Function ToByte() As ByteEx
    
    Set ToByte = New ByteEx
    
    If p_Value > 0& Then
        
        If p_Value < B_MX Then ToByte.Value = p_Value Else ToByte.Value = B_MX
    
    End If

End Function

Public Function ToDouble() As DoubleEx
    
    Set ToDouble = New DoubleEx
    
    ToDouble.Value = p_Value

End Function

Public Function ToInteger() As IntegerEx
    
    Set ToInteger = New IntegerEx
    
    If p_Value > I_MN Then
        
        If p_Value < I_MX Then ToInteger.Value = p_Value Else ToInteger.Value = I_MX
    
    Else
        ToInteger.Value = I_MN
    End If

End Function

Public Function ToString() As StringEx
    
    Set ToString = New StringEx
    
    ToString.Value = p_Value

End Function

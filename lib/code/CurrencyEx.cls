VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CurrencyEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'VARIABLES

Private p_Value As Currency

'PROPERTIES

Public Property Get Value() As Currency
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    Value = p_Value

End Property

Public Property Let Value(ByVal NewValue As Currency)
    
    p_Value = NewValue

End Property

'ROUTINES

Public Function Clone() As CurrencyEx
    
    Set Clone = New CurrencyEx
    
    Clone.Value = p_Value

End Function

Public Function Parse(ByRef Value As Variant) As CurrencyEx
    
    Set Parse = Me
    
    Select Case VarType(Value)
        
        Case vbBoolean, vbByte, vbCurrency, vbDate, vbInteger, vbLong
            
            p_Value = Value
        
        Case vbDecimal, vbDouble, vbSingle, vbString
        
        Case Else
            
            p_Value = 0@
    
    End Select

End Function

Public Function ToBoolean() As Boolean
    
    ToBoolean = p_Value

End Function

Public Function ToByte() As ByteEx
    
    Set ToByte = New ByteEx
    
    If p_Value > 0@ Then
        
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

Public Function ToLong() As LongEx
    
    Set ToLong = New LongEx
    
    If p_Value > L_MN Then
        
        If p_Value < L_MX Then ToLong.Value = p_Value Else ToLong.Value = L_MX
    
    Else
        ToLong.Value = L_MN
    End If

End Function

Public Function ToString() As StringEx
    
    Set ToString = New StringEx
    
    ToString.Value = p_Value

End Function

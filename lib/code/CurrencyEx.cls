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

'CONSTANTS

Private Const C_MN As Currency = -922337203685477.5807@
Private Const C_MX As Currency = 922337203685477.5807@

'VARIABLES

Private p_Value As Currency

'PROPERTIES

Public Property Get Max() As Currency
    
    Max = C_MX

End Property

Public Property Get Min() As Currency
    
    Min = C_MN

End Property

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
    
    Dim d As Double
    
    Set Parse = Me
    
    Select Case VarType(Value)
        
        Case vbBoolean, vbByte, vbCurrency, vbDate, vbError, vbInteger, vbLong
            
            p_Value = Value
        
        Case vbDecimal, vbDouble, vbSingle
            
            If Value > C_MN Then
                If Value < C_MX Then p_Value = Value Else p_Value = C_MX
            Else
                p_Value = C_MN
            End If
        
        Case vbString
            
            d = ToNumber(Value)
            
            If d > C_MN Then
                If d < C_MX Then p_Value = d Else p_Value = C_MX
            Else
                p_Value = C_MN
            End If
        
        Case Else
            
            p_Value = 0@
    
    End Select

End Function

Public Function ToArray() As ArrayEx
    
    Set ToArray = New ArrayEx
    
    ToArray.Parse p_Value

End Function

Public Function ToBoolean() As Boolean
    
    ToBoolean = p_Value

End Function

Public Function ToByte() As ByteEx
    
    Set ToByte = New ByteEx
    
    If p_Value > 0@ Then
        
        If p_Value < ToByte.Max Then ToByte.Value = p_Value Else ToByte.Value = ToByte.Max
    
    End If

End Function

Public Function ToDecimal() As DecimalEx
    
    Set ToDecimal = New DecimalEx
    
    ToDecimal.Value = p_Value

End Function

Public Function ToDouble() As DoubleEx
    
    Set ToDouble = New DoubleEx
    
    ToDouble.Value = p_Value

End Function

Public Function ToInteger() As IntegerEx
    
    Set ToInteger = New IntegerEx
    
    With ToInteger
        
        If p_Value > .Min Then
            
            If p_Value < .Max Then .Value = p_Value Else .Value = .Max
        
        Else
            .Value = .Min
        End If
    
    End With

End Function

Public Function ToLong() As LongEx
    
    Set ToLong = New LongEx
    
    With ToLong
        
        If p_Value > .Min Then
            
            If p_Value < .Max Then .Value = p_Value Else .Value = .Max
        
        Else
            .Value = .Min
        End If
    
    End With

End Function

Public Function ToString() As StringEx
    
    Set ToString = New StringEx
    
    ToString.Value = p_Value

End Function

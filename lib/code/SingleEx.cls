VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SingleEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CONSTANTS

Private Const S_MN As Double = -3.4028235E+38
Private Const S_MX As Double = 3.4028235E+38

'VARIABLES

Private p_Value As Single

'PROPERTIES

Public Property Get Length() As Long
    
    Length = 4&

End Property

Public Property Get Max() As Single
    
    Max = S_MX

End Property

Public Property Get Min() As Single
    
    Min = S_MN

End Property

Public Property Get Pointer() As Long
    
    Pointer = VarPtr(p_Value)

End Property

Public Property Get Value() As Single
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    Value = p_Value

End Property

Public Property Let Value(ByVal NewValue As Single)
    
    p_Value = NewValue

End Property

'ROUTINES

Public Function Clone() As SingleEx
    
    Set Clone = New SingleEx
    
    Clone.Value = p_Value

End Function

Public Function Parse(ByRef Value As Variant) As SingleEx
    
    Set Parse = Me
    
    Select Case VarType(Value)
        
        Case vbInteger To vbSingle, vbCurrency, vbDate, vbBoolean, vbByte, vbError
            
            p_Value = Value
        
        Case vbString
            
            p_Value = ToNumber(Value)
        
        Case vbDouble, vbDecimal
            
            If Value > S_MN Then
                
                If Value < S_MX Then p_Value = Value Else p_Value = S_MX
            
            Else
                p_Value = S_MN
            End If
         
         Case Else
            
            p_Value = 0#
    
    End Select

End Function

Public Function Round(ByVal DecimalDigits As Long) As SingleEx
    
    On Error Resume Next
    
    Set Round = Me
    
    p_Value = Math.Round(p_Value, DecimalDigits)

End Function

Public Function ToArray() As ArrayEx
    
    Set ToArray = New ArrayEx
    
    ToArray.Parse p_Value

End Function

Public Function ToBoolean() As BooleanEx
    
    Set ToBoolean = New BooleanEx
    
    ToBoolean.Value = p_Value

End Function

Public Function ToByte() As ByteEx
    
    Set ToByte = New ByteEx
    
    If p_Value > 0# Then
        
        If p_Value < ToByte.Max Then ToByte.Value = p_Value Else ToByte.Value = ToByte.Max
    
    End If

End Function

Public Function ToCurrency() As CurrencyEx
    
    Set ToCurrency = New CurrencyEx
    
    With ToCurrency
        
        If p_Value > .Min Then
            
            If p_Value < .Max Then .Value = p_Value Else .Value = .Max
        
        Else
            .Value = .Min
        End If
    
    End With

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

Public Function ToString(Optional ByRef Format As String) As StringEx
    
    Set ToString = New StringEx
    
    If Len(Format) Then ToString.Value = Strings.Format$(p_Value, Format) Else ToString.Value = p_Value

End Function

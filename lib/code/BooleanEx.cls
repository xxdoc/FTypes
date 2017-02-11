VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BooleanEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'VARIABLES

Private p_Value As Boolean

'PROPERTIES

Public Property Get Length() As Long
    
    Length = 2&

End Property

Public Property Get Max() As Boolean
    
    Max = True

End Property

Public Property Get Min() As Boolean
    
    Min = False

End Property

Public Property Get Pointer() As Long
    
    Pointer = VarPtr(p_Value)

End Property

Public Property Get Value() As Boolean
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    Value = p_Value

End Property

Public Property Let Value(ByVal NewValue As Boolean)
    
    p_Value = NewValue

End Property

'ROUTINES

Public Function Clone() As BooleanEx
    
    Set Clone = New BooleanEx
    
    Clone.Value = p_Value

End Function

Public Function Parse(ByRef Value As Variant) As BooleanEx
    
    Set Parse = Me
    
    Select Case VarType(Value)
        
        Case vbInteger To vbDate, vbBoolean, vbByte, vbDecimal, vbError
            
            p_Value = Value
        
        Case vbString
            
            p_Value = ToNumber(Value)
        
        Case Else
            
            p_Value = False
    
    End Select

End Function

Public Function Reverse() As BooleanEx
    
    Set Reverse = Me
    
    p_Value = Not p_Value

End Function

Public Function ToArray() As ArrayEx
    
    Set ToArray = New ArrayEx
    
    ToArray.Parse p_Value

End Function

Public Function ToByte() As ByteEx
    
    Set ToByte = New ByteEx
    
    ToByte.Value = p_Value * L_NG

End Function

Public Function ToCurrency() As CurrencyEx
    
    Set ToCurrency = New CurrencyEx
    
    ToCurrency.Value = p_Value

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

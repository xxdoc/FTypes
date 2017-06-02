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

'CONSTANTS

Private Const I_MN As Integer = -32768
Private Const I_MX As Integer = 32767

'VARIABLES

Private p_Value As Integer

'PROPERTIES

Public Property Get Length() As Long
    
    Length = 2&

End Property

Public Property Get Max() As Integer
    
    Max = I_MX

End Property

Public Property Get Min() As Integer
    
    Min = I_MN

End Property

Public Property Get Pointer() As Long
    
    Pointer = VarPtr(p_Value)

End Property

Public Property Get Value() As Integer
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    Value = p_Value

End Property

Public Property Let Value(ByVal NewValue As Integer)
    
    p_Value = NewValue

End Property

Public Property Get WordHigh() As Byte
    
    WordHigh = (p_Value And &HFF00&) \ 256

End Property

Public Property Get WordLow() As Byte
    
    WordLow = p_Value And 255

End Property

'ROUTINES

Public Function Clone() As IntegerEx
    
    Set Clone = New IntegerEx
    
    Clone.Value = p_Value

End Function

Public Function Parse(ByRef Value As Variant) As IntegerEx
    
    Dim d As Double
    
    Set Parse = Me
    
    Select Case VarType(Value)
        
        Case vbInteger, vbBoolean, vbByte
            
            p_Value = Value
        
        Case vbLong To vbDate, vbDecimal, vbError
            
            If Value > I_MN Then
                If Value < I_MX Then p_Value = Value Else p_Value = I_MX
            Else
                p_Value = I_MN
            End If
        
        Case vbString
            
            d = ToNumber(StrPtr(Value), Len(Value))
            
            If d > I_MN Then
                If d < I_MX Then p_Value = d Else p_Value = I_MX
            Else
                p_Value = I_MN
            End If
        
        Case Else
            
            p_Value = 0
    
    End Select

End Function

Public Function Range(ByVal MinValue As Integer, ByVal MaxValue As Integer) As IntegerEx
    
    Set Range = Me
    
    If p_Value > MaxValue Then
        
        p_Value = MaxValue
    
    Else
        
        If p_Value < MinValue Then p_Value = MinValue
    
    End If

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
    
    If p_Value > 0 Then
        
        If p_Value < ToByte.Max Then ToByte.Value = p_Value Else ToByte.Value = ToByte.Max
    
    End If

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

Public Function ToLong() As LongEx
    
    Set ToLong = New LongEx
    
    ToLong.Value = p_Value

End Function

Public Function ToSingle() As SingleEx
    
    Set ToSingle = New SingleEx
    
    ToSingle.Value = p_Value

End Function

Public Function ToString(Optional ByRef Format As String) As StringEx
    
    Set ToString = New StringEx
    
    If Len(Format) Then ToString.Value = Strings.Format$(p_Value, Format) Else ToString.Value = p_Value

End Function

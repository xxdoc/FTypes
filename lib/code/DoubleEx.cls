VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoubleEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CONSTANTS

Private Const D_MN As Double = -1.79769313486231E+308
Private Const D_MX As Double = 1.79769313486231E+308

'VARIABLES

Private p_Value As Double

'PROPERTIES

Public Property Get Length() As Long
    
    Length = 8&

End Property

Public Property Get Max() As Double
    
    Max = D_MX

End Property

Public Property Get Min() As Double
    
    Min = D_MN

End Property

Public Property Get Pointer() As Long
    
    Pointer = VarPtr(p_Value)

End Property

Public Property Get Value() As Double
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    Value = p_Value

End Property

Public Property Let Value(ByVal NewValue As Double)
    
    p_Value = NewValue

End Property

'ROUTINES

Public Function Clone() As DoubleEx
    
    Set Clone = New DoubleEx
    
    Clone.Value = p_Value

End Function

Public Function Parse(ByRef Value As Variant) As DoubleEx
    
    Set Parse = Me
    
    Select Case VarType(Value)
        
        Case vbInteger To vbDate, vbBoolean, vbByte, vbError
            
            p_Value = Value
        
        Case vbString
            
            p_Value = ToNumber(StrPtr(Value), Len(Value))
        
        Case vbDecimal
            
            If Value > D_MN Then
                If Value < D_MX Then p_Value = Value Else p_Value = D_MX
            Else
                p_Value = D_MN
            End If
         
         Case Else
            
            p_Value = 0#
    
    End Select

End Function

Public Function Range(ByVal MinValue As Double, ByVal MaxValue As Double) As DoubleEx
    
    Set Range = Me
    
    If p_Value > MaxValue Then
        
        p_Value = MaxValue
    
    Else
        
        If p_Value < MinValue Then p_Value = MinValue
    
    End If

End Function

Public Function Round(ByVal DecimalDigits As Long) As DoubleEx
    
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
    
    With ToDecimal
        
        If p_Value > .Min Then
            
            If p_Value < .Max Then .Value = p_Value Else .Value = .Max
        
        Else
            .Value = .Min
        End If
    
    End With

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

Public Function ToSingle() As SingleEx
    
    Set ToSingle = New SingleEx
    
    With ToSingle
        
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

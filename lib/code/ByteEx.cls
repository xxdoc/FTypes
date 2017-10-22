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

'INTERFACES

Implements TypesInterface

'CONSTANTS

Private Const B_MX As Byte = 255

'VARIABLES

Private p_Value As Byte

'PROPERTIES

Private Property Get TypesInterface_Pointer() As Long
    
    TypesInterface_Pointer = VarPtr(p_Value)

End Property

Public Property Get Length() As Long
    
    Length = 1&

End Property

Public Property Get Max() As Byte
    
    Max = B_MX

End Property

Public Property Get Min() As Byte
    
    Min = 0

End Property

Public Property Get Pointer() As Long
    
    Pointer = TypesInterface_Pointer

End Property

Public Property Get Value() As Byte
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    Value = p_Value

End Property

Public Property Let Value(ByVal NewValue As Byte)
    
    p_Value = NewValue

End Property

'METHODS

Private Function TypesInterface_Clone() As TypesInterface
    
    Set TypesInterface_Clone = New ByteEx
    
    TypesInterface_Clone.Parse p_Value

End Function

Private Function TypesInterface_Parse(Value As Variant) As TypesInterface
    
    Dim d As Double
    
    Set TypesInterface_Parse = Me
    
    Select Case VarType(Value)
        
        Case vbBoolean, vbByte
            
            p_Value = Value
        
        Case vbInteger To vbDate, vbDecimal, vbError
            
            If Value > 0 Then
                If Value < B_MX Then p_Value = Value Else p_Value = B_MX
            Else
                p_Value = 0
            End If
        
        Case vbString
            
            d = ToNumber(StrPtr(Value), Len(Value))
            
            If d > 0 Then
                If d < B_MX Then p_Value = d Else p_Value = B_MX
            Else
                p_Value = 0
            End If
        
        Case Else
            
            p_Value = 0
    
    End Select

End Function

Public Function Clone() As ByteEx
    
    Set Clone = TypesInterface_Clone

End Function

Public Function Parse(ByRef Value As Variant) As ByteEx
    
    Set Parse = TypesInterface_Parse(Value)

End Function

Public Function Range(ByVal MinValue As Byte, ByVal MaxValue As Byte) As ByteEx
    
    Set Range = Me
    
    If p_Value > MaxValue Then
        
        p_Value = MaxValue
    
    Else
        
        If p_Value < MinValue Then p_Value = MinValue
    
    End If

End Function

Public Function ToArray() As ListEx
    
    Set ToArray = ToList

End Function

Public Function ToBoolean() As BooleanEx
    
    Set ToBoolean = New BooleanEx
    
    ToBoolean.Value = p_Value

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

Public Function ToList() As ListEx
    
    Set ToList = New ListEx
    
    ToList.Parse p_Value

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

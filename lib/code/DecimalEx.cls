VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DecimalEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'INTERFACES

Implements IFTypesBasic

'VARIABLES

Private p_Value As Variant

'EVENTS

Private Sub Class_Initialize()
    
    p_Value = CDec(0&)

End Sub

'PROPERTIES

Private Property Get IFTypesBasic_Length() As Long
    
    IFTypesBasic_Length = 14&

End Property

Private Property Get IFTypesBasic_Pointer() As Long
    
    IFTypesBasic_Pointer = VarPtr(p_Value)

End Property

Public Property Get Length() As Long
    
    Length = IFTypesBasic_Length

End Property

Public Property Get Max() As Variant
    
    Max = CDec("79228162514264337593543950335")

End Property

Public Property Get Min() As Variant
    
    Min = CDec("-79228162514264337593543950335")

End Property

Public Property Get Pointer() As Long
    
    Pointer = IFTypesBasic_Pointer

End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    Value = p_Value

End Property

Public Property Let Value(ByRef NewValue As Variant)
    
    On Error Resume Next
    
    p_Value = CDec(NewValue)

End Property

'METHODS

Private Function IFTypesBasic_Parse(Value As Variant) As IFTypesBasic
    
    On Error GoTo ErrHandler
    
    Dim s As String
    
    Set IFTypesBasic_Parse = Me
    
    Select Case VarType(Value)
        
        Case vbDecimal
            
            p_Value = Value
        
        Case vbInteger To vbDate, vbBoolean, vbByte, vbError
            
            p_Value = CDec(Value)
        
        Case vbString
            
            s = ToNumber(StrPtr(Value), Len(Value))
            
            p_Value = CDec(Left$(s, 29& + (InStr(1&, s, String$(1&, m_Comma), vbBinaryCompare) > 0&) * L_NG))
        
        Case Else
            
            p_Value = CDec(0&)
    
    End Select
    
    Exit Function

ErrHandler:
    
    p_Value = CDec(0&)

End Function

Private Function IFTypesBasic_ToList() As IFTypesArray
    
    Set IFTypesBasic_ToList = New ListEx
    
    IFTypesBasic_ToList.Parse p_Value

End Function

Public Function Clone() As DecimalEx
    
    Set Clone = New DecimalEx
    
    Clone.Value = p_Value

End Function

Public Function Parse(ByRef Value As Variant) As DecimalEx
    
    Set Parse = IFTypesBasic_Parse(Value)

End Function

Public Function Range(ByRef MinValue As Variant, ByRef MaxValue As Variant) As DecimalEx
    
    Set Range = Me
    
    If p_Value > CDec(MaxValue) Then
        
        p_Value = CDec(MaxValue)
    
    Else
        
        If p_Value < CDec(MinValue) Then p_Value = CDec(MinValue)
    
    End If

End Function

Public Function Round(ByVal DecimalDigits As Long) As DecimalEx
    
    On Error Resume Next
    
    Set Round = Me
    
    p_Value = Math.Round(p_Value, DecimalDigits)

End Function

Public Function ToArray() As ListEx
    
    Set ToArray = New ListEx
    
    ToArray.Create vbByte, Length
    
    RtlMoveMemory ByVal ToArray.Data, ByVal Pointer, Length

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
    
    With ToCurrency
        
        If p_Value > .Min Then
            
            If p_Value < .Max Then .Value = p_Value Else .Value = .Max
        
        Else
            .Value = .Min
        End If
    
    End With

End Function

Public Function ToDouble() As DoubleEx
    
    Set ToDouble = New DoubleEx
    
    With ToDouble
        
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

Public Function ToList() As ListEx
    
    Set ToList = IFTypesBasic_ToList

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

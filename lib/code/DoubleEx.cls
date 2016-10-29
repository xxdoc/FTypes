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

'VARIABLES

Private p_Value As Double

'PROPERTIES

Public Property Get Max() As Double
    
    Max = D_MX

End Property

Public Property Get Min() As Double
    
    Min = D_MN

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
        
        Case 2& To 5&, 7&, 11&, 17& 'INTEGER, LONG, SINGLE, DOUBLE, DATE, BOOLEAN, BYTE
            
            p_Value = Value
        
        Case 8& 'STRING
            
            p_Value = ToNumber(Value)
        
        Case 6&, 14& 'CURRENCY, DECIMAL
            
            If Value > D_MN Then
                
                If Value < D_MX Then p_Value = Value Else p_Value = D_MX
            
            Else
                p_Value = D_MN
            End If
         
         Case Else
            
            p_Value = 0#
    
    End Select

End Function

Public Function ToBoolean() As Boolean
    
    ToBoolean = p_Value

End Function

Public Function ToByte() As ByteEx
    
    Set ToByte = New ByteEx
    
    If p_Value > 0# Then
        
        If p_Value < B_MX Then ToByte.Value = p_Value Else ToByte.Value = B_MX
    
    End If

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

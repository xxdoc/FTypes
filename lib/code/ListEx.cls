VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'INTERFACES

Implements IFTypesArray

'VARIABLES

Private p_Array As TSAFEARRAY

'EVENTS

Private Sub Class_Terminate()
    
    ArrayDestroy p_Array

End Sub

'PROPERTIES

Private Property Get IFTypesArray_Data() As Long
    
    IFTypesArray_Data = p_Array.lData

End Property

Private Property Get IFTypesArray_ElementSize() As Long
    
    IFTypesArray_ElementSize = p_Array.lElementSize

End Property

Private Property Get IFTypesArray_Pointer() As Long
    
    IFTypesArray_Pointer = p_Array.lPointer

End Property

Public Property Get Count() As Long
    
    If p_Array.lPointer Then Count = p_Array.uBounds(0).lElements

End Property

Public Property Get Data() As Long
    
    Data = IFTypesArray_Data

End Property

Public Property Get ElementSize() As Long
    
    ElementSize = IFTypesArray_ElementSize

End Property

Public Property Get Item(ByVal Index As Long) As Variant
    
    If Index >= 0& And Index <= Count Then
        
        VariantSetType Item, p_Array.lVarType
        
        ArrayElementGet p_Array, VarPtr(Index), Item
    
    End If

End Property

Public Property Let Item(ByVal Index As Long, ByRef NewValue As Variant)
    
    If Index >= 0& And Index <= Count Then ArrayElementSet p_Array, VarPtr(Index), NewValue

End Property

Public Property Get Pointer() As Long
    
    Pointer = IFTypesArray_Pointer

End Property

Public Property Get Type_() As VbVarType
    
    Type_ = p_Array.lVarType

End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    If p_Array.lPointer Then ArrayCopy p_Array, Value

End Property

Public Property Let Value(ByRef NewValue As Variant)
    
    Dim a As TSAFEARRAY
    Dim b(0) As TSAFEARRAYBOUND
    
    ArrayPtr a, VarPtr(NewValue) * (((VarType(NewValue) And vbArray) = vbArray) * L_NG), True
    
    If a.lPointer > 0& And a.iDims = 1 Then
        
        b(0).lElements = a.uBounds(0).lElements
        
        If ArrayCreate(p_Array, a.lVarType, a.iDims, b) Then ArrayCopy p_Array, a.lPointer
    
    End If

End Property

'METHODS

Private Function IFTypesArray_Add(Item As Variant) As IFTypesArray
    
    Dim b(0) As TSAFEARRAYBOUND
    
    Set IFTypesArray_Add = Me
    
    If p_Array.lPointer Then
        
        ArrayRedim p_Array, p_Array.uBounds(0).lElements + 1&
    
    Else
        
        b(0).lElements = 1&
        
        ArrayCreate p_Array, VarType(Item), 1&, b
    
    End If
    
    ArrayElementSet p_Array, VarPtr(Count + L_NG), Item

End Function

Private Function IFTypesArray_Parse(Value As Variant) As IFTypesArray
    
    Dim b(0) As TSAFEARRAYBOUND
    
    Set IFTypesArray_Parse = Me
    
    b(0).lElements = 1&
    
    If ArrayCreate(p_Array, VarType(Value), 1&, b) Then ArrayElementSet p_Array, VarPtr(b(0).lLowest), Value

End Function

Private Function IFTypesArray_Resize(ByVal NewUpperBound As Long) As IFTypesArray
    
    Set IFTypesArray_Resize = Me
    
    If p_Array.lPointer Then
        
        If NewUpperBound >= 0& And NewUpperBound < p_Array.uBounds(0&).lElements Then ArrayRedim p_Array, NewUpperBound + 1&
    
    End If

End Function

Public Function Add(ByRef Item As Variant) As ListEx
    
    Set Add = IFTypesArray_Add(Item)

End Function

Public Function AddRange(ByRef Range As Variant) As ListEx
    
    Dim a As TSAFEARRAY
    Dim i As Long
    Dim x As Long
    
    Set AddRange = Me
    
    If p_Array.lPointer Then
        
        ArrayPtr a, VarPtr(Range) * (((VarType(Range) And vbArray) = vbArray) * L_NG), True
        
        If a.lPointer > 0& And a.iDims = 1& Then
            
            x = p_Array.uBounds(0).lElements
            
            ArrayRedim p_Array, x + a.uBounds(0).lElements
            
            For i = a.uBounds(0).lLowest To a.uBounds(0).lLowest + a.uBounds(0).lElements + L_NG
                
                x = x + 1&
                
                ArrayElementSet p_Array, VarPtr(x + L_NG), Range(i)
            
            Next i
        
        End If
    
    Else
        
        Me.Value = Range
    
    End If

End Function

Public Function Clone() As ListEx
    
    Set Clone = New ListEx
    
    If p_Array.lPointer Then Clone.Value = Value

End Function

Public Function Create(ByVal ArrayType As VbVarType, ByVal Elements As Long) As ListEx
    
    On Error GoTo ErrHandler
    
    Dim b(0) As TSAFEARRAYBOUND
    
    Set Create = Me
    
    If Elements > 0& Then
        
        b(0).lElements = Elements
        
        ArrayCreate p_Array, ArrayType, 1&, b
    
    End If

ErrHandler:

End Function

Public Function Distinct() As ListEx
    
    Dim i As Long
    Dim j As Long
    Dim v1 As Variant
    Dim v2 As Variant
    Dim u As Long
    Dim x As Long
    
    Set Distinct = Me
    
    If p_Array.lPointer Then
        
        Sort
        
        u = Count + L_NG
        
        VariantSetType v1, p_Array.lVarType
        VariantSetType v2, p_Array.lVarType
        
        Do
            
            ArrayElementGet p_Array, VarPtr(i), v1
            
            For j = i + 1& To u
                
                i = j
                
                ArrayElementGet p_Array, VarPtr(j), v2
                
                If v1 <> v2 Then
                    
                    x = x + 1&
                    
                    ArrayElementSet p_Array, VarPtr(x), v2
                    
                    Exit For
                
                End If
            
            Next j
        
        Loop While i < u
        
        VariantZero v1
        VariantZero v2
        
        Resize x
    
    End If

End Function

Public Function Parse(ByRef Value As Variant) As ListEx
    
    Set Parse = IFTypesArray_Parse(Value)

End Function

Public Function Resize(ByVal NewUpperBound As Long) As ListEx
    
    Set Resize = IFTypesArray_Resize(NewUpperBound)

End Function

Public Function Sort(Optional ByVal Descending As Boolean, Optional ByVal StartIndex As Long = L_NG, Optional ByVal EndIndex As Long = L_NG) As ListEx
    
    On Error GoTo ErrHandler
    
    Dim e As Long
    Dim i As Long
    Dim j As Long
    Dim p As Variant
    Dim s As Long
    Dim x As Long
    Dim v1 As Variant
    Dim v2 As Variant
    Dim u As Long
    
    Set Sort = Me
    
    If p_Array.lPointer Then
        
        e = p_Array.uBounds(0).lElements + L_NG
        
        If StartIndex > 0& And StartIndex <= e Then s = StartIndex
        If EndIndex >= s And EndIndex <= e Then e = EndIndex
        
        u = (s + e) \ 2&
        x = VarPtr(u)
        
        VariantSetType p, p_Array.lVarType
        
        If ArrayElementGet(p_Array, VarPtr(u), p) = 0& Then
            
            VariantSetType v1, p_Array.lVarType
            VariantSetType v2, p_Array.lVarType
            
            i = s
            j = e
            
            While i <= j
                
                Do
                    
                    u = i
                    
                    ArrayElementGet p_Array, x, v1
                    
                    If ((Descending = False And v1 < p) Or (Descending = True And v1 > p)) And i < e Then i = i + 1& Else Exit Do
                
                Loop
                
                Do
                    
                    u = j
                    
                    ArrayElementGet p_Array, x, v2
                    
                    If ((Descending = False And p < v2) Or (Descending = True And p > v2)) And j > s Then j = j + L_NG Else Exit Do
                
                Loop
                
                If i <= j Then
                    
                    ArrayElementSet p_Array, x, v1
                    u = i
                    ArrayElementSet p_Array, x, v2
                    
                    i = i + 1&
                    j = j + L_NG
                
                End If
            
            Wend
            
            VariantZero v1
            VariantZero v2
            
            If s < j Then
                u = s
                Sort Descending, u, j
            End If
            
            If i < e Then
                u = i
                Sort Descending, u, e
            End If
        
        End If
        
        VariantZero p
    
    End If

ErrHandler:

End Function

Public Function ToString(Optional ByRef Delimeter As String, Optional ByVal StartIndex As Long = L_NG, Optional ByVal EndIndex As Long = L_NG) As StringEx
    
    Dim i As Long
    Dim j As Long
    Dim v As Variant
    
    Set ToString = New StringEx
    
    If p_Array.lPointer Then
        
        j = p_Array.uBounds(0).lElements + L_NG
        
        If StartIndex > 0& And StartIndex <= j Then i = StartIndex
        If EndIndex >= i And EndIndex <= j Then j = EndIndex
        
        VariantSetType v, p_Array.lVarType
        
        For i = i To j
            
            If ArrayElementGet(p_Array, VarPtr(i), v) = 0& Then ToString.Concat((v)).Concat Delimeter
        
        Next i
        
        VariantZero v
    
    End If

End Function
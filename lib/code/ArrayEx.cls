VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'INTERFACES

Implements TypesInterface

'VARIABLES

Private p_Array As TSAFEARRAY

'EVENTS

Private Sub Class_Terminate()
    
    ArrayDestroy p_Array

End Sub

'PROPERTIES

Private Property Get TypesInterface_Pointer() As Long
    
    TypesInterface_Pointer = p_Array.lPointer

End Property

Public Property Get BoundL(Optional ByVal Dimension As Long) As Long
    
    If p_Array.lPointer Then
        
        If Dimension > 0& And Dimension <= p_Array.iDims Then BoundL = p_Array.uBounds(p_Array.iDims - Dimension).lLowest Else BoundL = p_Array.uBounds(p_Array.iDims + L_NG).lLowest
    
    End If

End Property

Public Property Get BoundU(Optional ByVal Dimension As Long) As Long
    
    Dim x As Long
    
    If p_Array.lPointer Then
        
        If Dimension > 0& And Dimension <= p_Array.iDims Then x = p_Array.iDims - Dimension Else x = p_Array.iDims + L_NG
        
        BoundU = p_Array.uBounds(x).lLowest + p_Array.uBounds(x).lElements + L_NG
    
    End If

End Property

Public Property Get Data() As Long
    
    Data = p_Array.lData

End Property

Public Property Get Dimensions() As Long
    
    Dimensions = p_Array.iDims

End Property

Public Property Get Elements(Optional ByVal Dimension As Long) As Long
    
    If p_Array.lPointer Then
        
        If Dimension > 1& And Dimension <= p_Array.iDims Then Elements = p_Array.uBounds(p_Array.iDims - Dimension).lElements Else Elements = p_Array.uBounds(p_Array.iDims + L_NG).lElements
    
    End If

End Property

Public Property Get ElementSize() As Long
    
    ElementSize = p_Array.lElementSize

End Property

Public Property Get Item(ParamArray Indexes()) As Variant
    
    On Error GoTo ErrHandler
    
    Dim b() As Long
    
    If p_Array.lPointer Then
        
        VariantSetType Item, p_Array.lVarType
        
        If p_Array.iDims > 1 Then
            
            ArrayElementGet p_Array, ArrVariantToIndexes(Indexes, b), Item
            
            Erase b
        
        Else
            ArrayElementGet p_Array, VarPtr(CLng(Indexes(0))), Item
        End If
    
    End If

ErrHandler:

End Property

Public Property Let Item(ParamArray Indexes(), ByRef NewValue As Variant)
    
    On Error GoTo ErrHandler
    
    Dim b() As Long
    
    If p_Array.lPointer Then
        
        If p_Array.iDims > 1 Then
            
            ArrayElementSet p_Array, ArrVariantToIndexes(Indexes, b), NewValue
            
            Erase b
        
        Else
            ArrayElementSet p_Array, VarPtr(CLng(Indexes(0))), NewValue
        End If
    
    End If

ErrHandler:

End Property

Public Property Get Pointer() As Long
    
    Pointer = TypesInterface_Pointer

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
    Dim b() As TSAFEARRAYBOUND
    Dim i As Long
    
    ArrayPtr a, VarPtr(NewValue) * (((VarType(NewValue) And vbArray) = vbArray) * L_NG), True
    
    If a.lPointer Then
        
        ReDim b(a.iDims + L_NG)
        
        For i = 0& To a.iDims + L_NG: b(a.iDims - i + L_NG) = a.uBounds(i): Next i
        
        If ArrayCreate(p_Array, a.lVarType, a.iDims, b) Then ArrayCopy p_Array, a.lPointer
        
        Erase b
    
    End If

End Property

'METHODS

Private Function TypesInterface_Clone() As TypesInterface
    
    Set TypesInterface_Clone = New ArrayEx
    
    If p_Array.lPointer Then TypesInterface_Clone.Parse Me.Value

End Function

Private Function TypesInterface_Parse(Value As Variant) As TypesInterface
    
    Dim b(0) As TSAFEARRAYBOUND
    
    Set TypesInterface_Parse = Me
    
    If (VarType(Value) And vbArray) = vbArray Then
        
        Me.Value = Value
    
    Else
        
        b(0).lElements = 1&
        
        If ArrayCreate(p_Array, VarType(Value), 1&, b) Then ArrayElementSet p_Array, VarPtr(b(0).lLowest), Value
    
    End If

End Function

Public Function Add(ByRef Item As Variant) As ArrayEx
    
    Dim a(0) As TSAFEARRAYBOUND
    Dim b() As Long
    Dim x As Long
    
    Set Add = Me
    
    If p_Array.lPointer Then
        
        ArrayRedim p_Array, p_Array.uBounds(0).lElements + 1&
        
        If p_Array.iDims > 1& Then
            
            x = p_Array.iDims + L_NG
            
            ReDim b(x)
            
            For x = 0& To x: b(x) = BoundU(x + 1&): Next x
            
            ArrayElementSet p_Array, VarPtr(b(0)), Item
            
            Erase b
        
        Else
            ArrayElementSet p_Array, VarPtr(BoundU), Item
        End If
    
    Else
        
        a(0).lElements = 1&
        
        If ArrayCreate(p_Array, VarType(Item), 1&, a) Then ArrayElementSet p_Array, VarPtr(0&), Item
    
    End If

End Function

Public Function AddRange(ByRef Range As Variant) As ArrayEx
    
    Dim a As TSAFEARRAY
    Dim b() As Long
    Dim i As Long
    Dim x As Long
    
    Set AddRange = Me
    
    If p_Array.lPointer Then
        
        ArrayPtr a, VarPtr(Range) * (((VarType(Range) And vbArray) = vbArray) * L_NG), True
        
        If a.lPointer <> 0& And a.iDims = 1& Then
            
            x = p_Array.iDims + L_NG
            
            ReDim b(x)
            
            For i = 0& To x: b(i) = BoundU(i + 1&): Next i
            
            ArrayRedim p_Array, p_Array.uBounds(0).lElements + a.uBounds(0).lElements
            
            For i = a.uBounds(0).lLowest To a.uBounds(0).lLowest + a.uBounds(0).lElements + L_NG
                b(x) = b(x) + 1&
                ArrayElementSet p_Array, VarPtr(b(0)), Range(i)
            Next i
            
            Erase b
        
        End If
    
    Else
        
        Me.Value = Range
    
    End If

End Function

Public Function Clone() As ArrayEx
    
    Set Clone = TypesInterface_Clone

End Function

Public Function Create(ByVal ArrayType As VbVarType, ParamArray Bounds()) As ArrayEx
    
    On Error GoTo ErrHandler
    
    Dim b() As TSAFEARRAYBOUND
    Dim i As Long
    Dim x As Long
    
    Set Create = Me
    
    x = UBound(Bounds)
    
    ReDim b(x)
    
    For i = 0& To x
        b(i).lLowest = Bounds(i)(0)
        b(i).lElements = Bounds(i)(1) - b(i).lLowest + 1&
    Next i
    
    ArrayCreate p_Array, ArrayType, x + 1&, b
    
    Erase b

ErrHandler:

End Function

Public Function Destroy() As ArrayEx
    
    Set Destroy = Me
    
    ArrayDestroy p_Array

End Function

Public Function Distinct() As ArrayEx
    
    Dim i As Long
    Dim j As Long
    Dim v1 As Variant
    Dim v2 As Variant
    Dim u As Long
    Dim x As Long
    
    Set Distinct = Me
    
    If p_Array.lPointer > 0& And p_Array.iDims = 1& Then
        
        i = BoundL
        x = i
        
        Sort Array(i)
        
        u = BoundU
        
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
        
        p_Array.uBounds(0&).lElements = x - p_Array.uBounds(0&).lLowest + 1&
        
        Resize x
    
    End If

End Function

Public Function Parse(ByRef Value As Variant) As ArrayEx
    
    Set Parse = TypesInterface_Parse(Value)

End Function

Public Function Resize(ByVal NewUpperBound As Long) As ArrayEx
    
    Set Resize = Me
    
    If p_Array.lPointer Then
        
        If NewUpperBound >= p_Array.uBounds(0&).lLowest Then ArrayRedim p_Array, NewUpperBound - p_Array.uBounds(0&).lLowest + 1&
    
    End If

End Function

Public Function Sort(ByRef StartIndexes As Variant, Optional ByVal Descending As Boolean, Optional ByVal EndColumn As Long = L_MX) As ArrayEx
    
    On Error GoTo ErrHandler
    
    Dim b() As Long
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
        
        x = ArrVariantToIndexes(StartIndexes, b)
        
        u = UBound(b)
        
        s = b(u)
        e = p_Array.uBounds(0).lLowest + p_Array.uBounds(0).lElements + L_NG
        If EndColumn > s And EndColumn < e Then e = EndColumn
        
        b(u) = (s + e) \ 2&
        
        VariantSetType p, p_Array.lVarType
        
        If ArrayElementGet(p_Array, x, p) = 0& Then
            
            VariantSetType v1, p_Array.lVarType
            VariantSetType v2, p_Array.lVarType
            
            i = s
            j = e
            
            While i <= j
                
                Do
                    
                    b(u) = i
                    
                    ArrayElementGet p_Array, x, v1
                    
                    If ((Descending = False And v1 < p) Or (Descending = True And v1 > p)) And i < e Then i = i + 1& Else Exit Do
                
                Loop
                
                Do
                    
                    b(u) = j
                    
                    ArrayElementGet p_Array, x, v2
                    
                    If ((Descending = False And p < v2) Or (Descending = True And p > v2)) And j > s Then j = j + L_NG Else Exit Do
                
                Loop
                
                If i <= j Then
                    
                    ArrayElementSet p_Array, x, v1
                    b(u) = i
                    ArrayElementSet p_Array, x, v2
                    
                    i = i + 1&
                    j = j + L_NG
                
                End If
            
            Wend
            
            VariantZero v1
            VariantZero v2
            
            If s < j Then
                b(u) = s
                Sort b, Descending, j
            End If
            
            If i < e Then
                b(u) = i
                Sort b, Descending, e
            End If
        
        End If
        
        VariantZero p
        
        Erase b
    
    End If

ErrHandler:

End Function

Public Function ToString(ByRef StartIndexes As Variant, Optional ByRef Delimeter As String, Optional ByVal EndColumn As Long = L_MX) As StringEx
    
    Dim b() As Long
    Dim i As Long
    Dim x As Long
    Dim u As Long
    Dim v As Variant
    
    Set ToString = New StringEx
    
    If p_Array.lPointer Then
        
        x = ArrVariantToIndexes(StartIndexes, b)
        u = UBound(b)
        
        i = p_Array.uBounds(0).lLowest + p_Array.uBounds(0).lElements + L_NG
        If EndColumn > b(u) And EndColumn < i Then i = EndColumn
        
        VariantSetType v, p_Array.lVarType
        
        For i = b(u) To i
            
            b(u) = i
            
            If ArrayElementGet(p_Array, x, v) = 0& Then ToString.Concat((v)).Concat Delimeter
        
        Next i
        
        VariantZero v
        
        Erase b
    
    End If

End Function

Private Function ArrVariantToIndexes(ByVal SourceArray As Variant, ByRef OutputArray() As Long) As Long
    
    Dim i As Long
    Dim u As Long
    
    u = UBound(SourceArray)
    
    ReDim OutputArray(u)
    
    For i = 0& To u
        OutputArray(i) = SourceArray(i)
    Next i
    
    ArrVariantToIndexes = VarPtr(OutputArray(0))

End Function

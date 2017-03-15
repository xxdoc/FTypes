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

'VARIABLES

Private p_Array As TSAFEARRAY

'WINAPI

Private Declare Function SafeArrayCopy Lib "oleaut32" (ByVal lSource As Long, ByVal lTarget As Any) As Long
Private Declare Function SafeArrayCopyData Lib "oleaut32" (ByVal lSource As Long, ByVal lTarget As Any) As Long
Private Declare Function SafeArrayCreate Lib "oleaut32" (ByVal lType As Integer, ByVal lDims As Long, uBounds As Any) As Long
Private Declare Function SafeArrayDestroy Lib "oleaut32" (ByVal lArray As Long) As Long
Private Declare Function SafeArrayGetElement Lib "oleaut32" (ByVal lArray As Long, ByVal lIndices As Long, uValue As Any) As Long
Private Declare Function SafeArrayPutElement Lib "oleaut32" (ByVal lArray As Long, ByVal lIndices As Long, uValue As Any) As Long
Private Declare Function SafeArrayRedim Lib "oleaut32" (ByVal lArray As Long, uLastBound As TSAFEARRAYBOUND) As Long
Private Declare Sub RtlZeroMemory Lib "kernel32" (uDestination As Any, ByVal lLen As Long)

'EVENTS

Private Sub Class_Terminate()
    
    ArrDestroy

End Sub

'PROPERTIES

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
        
        RtlMoveMemory Item, p_Array.lVarType, 2&
        
        If p_Array.iDims > 1 Then
            
            ArrElementGet ArrVariantToIndexes(Indexes, b), Item
            
            Erase b
        
        Else
            ArrElementGet VarPtr(CLng(Indexes(0))), Item
        End If
    
    End If

ErrHandler:

End Property

Public Property Let Item(ParamArray Indexes(), ByRef NewValue As Variant)
    
    On Error GoTo ErrHandler
    
    Dim b() As Long
    
    If p_Array.lPointer Then
        
        If p_Array.iDims > 1 Then
            
            ArrElementSet ArrVariantToIndexes(Indexes, b), NewValue
            
            Erase b
        
        Else
            ArrElementSet VarPtr(CLng(Indexes(0))), NewValue
        End If
    
    End If

ErrHandler:

End Property

Public Property Get Pointer() As Long
    
    Pointer = p_Array.lPointer

End Property

Public Property Get Type_() As VbVarType
    
    Type_ = p_Array.lVarType

End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    If p_Array.lPointer Then
        
        RtlMoveMemory Value, vbArray + p_Array.lVarType, 2&
        
        SafeArrayCopy p_Array.lPointer, VarPtr(Value) + 8&
    
    End If

End Property

Public Property Let Value(ByRef NewValue As Variant)
    
    Dim a As TSAFEARRAY
    Dim b() As TSAFEARRAYBOUND
    Dim i As Long
    
    ArrPtr VarPtr(NewValue) * (((VarType(NewValue) And vbArray) = vbArray) * L_NG), a, True
    
    If a.lPointer Then
        
        ReDim b(a.iDims + L_NG)
        
        For i = 0& To a.iDims + L_NG: b(a.iDims - i + L_NG) = a.uBounds(i): Next i
        
        If ArrCreate(a.lVarType, a.iDims, b) Then SafeArrayCopyData a.lPointer, p_Array.lPointer
        
        Erase b
    
    End If

End Property

'ROUTINES

Public Function Add(ByRef Item As Variant) As ArrayEx
    
    Dim a(0) As TSAFEARRAYBOUND
    Dim b() As Long
    Dim x As Long
    
    Set Add = Me
    
    x = VarType(Item)
    
    If p_Array.lPointer = 0& Then
        
        a(0).lElements = 1&
        
        ArrCreate x, 1&, a
        
        ArrElementSet VarPtr(0&), Item
    
    ElseIf x = p_Array.lVarType Then
        
        p_Array.uBounds(0).lElements = p_Array.uBounds(0).lElements + 1&
        
        SafeArrayRedim p_Array.lPointer, p_Array.uBounds(0)
        
        If p_Array.iDims > 1& Then
            
            x = p_Array.iDims + L_NG
            
            ReDim b(x)
            
            For x = 0& To x: b(x) = BoundU(x + 1&): Next x
            
            ArrElementSet VarPtr(b(0)), Item
            
            Erase b
        
        Else
            ArrElementSet VarPtr(BoundU), Item
        End If
    
    End If

End Function

Public Function AddRange(ByRef Range As Variant) As ArrayEx
    
    Dim a As TSAFEARRAY
    Dim b() As Long
    Dim i As Long
    Dim x As Long
    
    Set AddRange = Me
    
    ArrPtr VarPtr(Range) * (((VarType(Range) And vbArray) = vbArray) * L_NG), a, True
    
    If a.lPointer <> 0& And a.iDims = 1& Then
        
        If p_Array.lPointer = 0& Then
            
            ArrCreate a.lVarType, 1&, a.uBounds
            
            For i = BoundL To BoundU: ArrElementSet VarPtr(i), Range(i): Next i
        
        ElseIf a.lVarType = p_Array.lVarType Then
            
            x = p_Array.iDims + L_NG
            
            ReDim b(x)
            
            For i = 0& To x: b(i) = BoundU(i + 1&): Next i
            
            p_Array.uBounds(0).lElements = p_Array.uBounds(0).lElements + a.uBounds(0).lElements
            
            SafeArrayRedim p_Array.lPointer, p_Array.uBounds(0)
            
            For i = a.uBounds(0).lLowest To a.uBounds(0).lLowest + a.uBounds(0).lElements + L_NG
                b(x) = b(x) + 1&
                ArrElementSet VarPtr(b(0)), Range(i)
            Next i
            
            Erase b
        
        End If
    
    End If

End Function

Public Function Clone() As ArrayEx
    
    Set Clone = New ArrayEx
    
    If p_Array.lPointer Then Clone.Value = Value

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
    
    ArrCreate ArrayType, x + 1&, b
    
    Erase b

ErrHandler:

End Function

Public Function Destroy() As ArrayEx
    
    Set Destroy = Me
    
    ArrDestroy

End Function

Public Function Parse(ByRef Value As Variant) As ArrayEx
    
    Dim b(0) As TSAFEARRAYBOUND
    
    Set Parse = Me
    
    b(0).lElements = 1&
    
    If ArrCreate(VarType(Value), 1&, b) Then ArrElementSet VarPtr(b(0).lLowest), Value

End Function

Public Function Resize(ByVal NewUpperBound As Long) As ArrayEx
    
    Set Resize = Me
    
    If p_Array.lPointer Then
        
        If NewUpperBound >= p_Array.uBounds(0&).lLowest Then
            
            p_Array.uBounds(0&).lElements = NewUpperBound - p_Array.uBounds(0&).lLowest + 1&
            
            SafeArrayRedim p_Array.lPointer, p_Array.uBounds(0&)
        
        End If
    
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
        
        RtlMoveMemory p, p_Array.lVarType, 2&
        
        If ArrElementGet(x, p) = 0& Then
            
            RtlMoveMemory v1, p_Array.lVarType, 2&
            RtlMoveMemory v2, p_Array.lVarType, 2&
            
            i = s
            j = e
            
            While i <= j
                
                Do
                    
                    b(u) = i
                    
                    ArrElementGet x, v1
                    
                    If ((Descending = False And v1 < p) Or (Descending = True And v1 > p)) And i < e Then i = i + 1& Else Exit Do
                
                Loop
                
                Do
                    
                    b(u) = j
                    
                    ArrElementGet x, v2
                    
                    If ((Descending = False And p < v2) Or (Descending = True And p > v2)) And j > s Then j = j + L_NG Else Exit Do
                
                Loop
                
                If i <= j Then
                    
                    ArrElementSet x, v1
                    b(u) = i
                    ArrElementSet x, v2
                    
                    i = i + 1&
                    j = j + L_NG
                
                End If
            
            Wend
            
            RtlMoveMemory v1, 0&, 4&
            RtlMoveMemory v2, 0&, 4&
            
            x = p_Array.uBounds(0).lLowest
            
            If s < j Then
                b(u) = s + x
                Sort b, Descending, j + x
            End If
            
            If i < e Then
                b(u) = i + x
                Sort b, Descending, e + x
            End If
        
        End If
        
        RtlMoveMemory p, 0&, 4&
        
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
        
        RtlMoveMemory v, p_Array.lVarType, 2&
        
        For i = b(u) To i
            
            b(u) = i
            
            If ArrElementGet(x, v) = 0& Then ToString.Concat((v)).Concat Delimeter
        
        Next i
        
        RtlMoveMemory v, 0&, 4&
        
        Erase b
    
    End If

End Function

Private Function ArrCreate(ByVal lType As VbVarType, ByVal lDims As Long, ByRef uBounds() As TSAFEARRAYBOUND) As Boolean
    
    Dim x As Long
    
    ArrDestroy
    
    x = SafeArrayCreate(lType Xor (vbArray * (((lType And vbArray) = vbArray) * L_NG)), lDims, uBounds(0))
    
    If x Then
        
        ArrPtr x, p_Array
        
        ArrCreate = p_Array.lPointer
    
    End If

End Function

Private Sub ArrDestroy()
    
    SafeArrayDestroy p_Array.lPointer
    
    RtlZeroMemory p_Array, Len(p_Array)

End Sub

Private Function ArrElementGet(ByVal Indexes As Long, ByRef Value As Variant) As Long
    
    Select Case p_Array.lVarType
        Case vbDecimal: ArrElementGet = SafeArrayGetElement(p_Array.lPointer, Indexes, ByVal VarPtr(Value))
        Case vbVariant: ArrElementGet = SafeArrayGetElement(p_Array.lPointer, Indexes, Value)
        Case Else: ArrElementGet = SafeArrayGetElement(p_Array.lPointer, Indexes, ByVal VarPtr(Value) + 8&)
    End Select

End Function

Private Sub ArrElementSet(ByVal Indexes As Long, ByRef Value As Variant)
    
    Select Case p_Array.lVarType
        Case vbBoolean: SafeArrayPutElement p_Array.lPointer, Indexes, CBool(Value)
        Case vbByte: SafeArrayPutElement p_Array.lPointer, Indexes, CByte(Value)
        Case vbCurrency: SafeArrayPutElement p_Array.lPointer, Indexes, CCur(Value)
        Case vbDate: SafeArrayPutElement p_Array.lPointer, Indexes, CDate(Value)
        Case vbError, vbLong: SafeArrayPutElement p_Array.lPointer, Indexes, CLng(Value)
        Case vbDecimal: SafeArrayPutElement p_Array.lPointer, Indexes, ByVal VarPtr(CDec(Value))
        Case vbDouble: SafeArrayPutElement p_Array.lPointer, Indexes, CDbl(Value)
        Case vbInteger: SafeArrayPutElement p_Array.lPointer, Indexes, CInt(Value)
        Case vbSingle: SafeArrayPutElement p_Array.lPointer, Indexes, CSng(Value)
        Case vbString: SafeArrayPutElement p_Array.lPointer, Indexes, ByVal StrPtr(Value)
        Case Else: SafeArrayPutElement p_Array.lPointer, Indexes, Value
    End Select

End Sub

Private Sub ArrPtr(ByVal lArray As Long, ByRef uPtr As TSAFEARRAY, Optional ByVal IsExternal As Boolean)
    
    If lArray Then
        
        RtlMoveMemory uPtr.lVarType, ByVal lArray + (-4& * ((Not IsExternal) * L_NG)), 2& 'get array type
        
        If IsExternal Then
            
            RtlMoveMemory uPtr.lPointer, ByVal lArray + 8&, 4&
            
            If uPtr.lVarType And 16384& Then 'if passed by reference (pointer to pointer)
                RtlMoveMemory uPtr.lPointer, ByVal uPtr.lPointer, 4&
                uPtr.lVarType = uPtr.lVarType Xor 16384&
            End If
            
            uPtr.lVarType = uPtr.lVarType Xor vbArray
        
        Else
            uPtr.lPointer = lArray
        End If
        
        RtlMoveMemory uPtr, ByVal uPtr.lPointer, 16& 'fill first fixed 16 bytes from pointer
        
        ReDim uPtr.uBounds(uPtr.iDims + L_NG) 'allocate bounds member
        
        RtlMoveMemory uPtr.uBounds(0), ByVal uPtr.lPointer + 16&, uPtr.iDims * 8& 'get array dimensions info bytes (in descending order) starting from array pointer adress + 16 bytes offset
    
    End If

End Sub

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

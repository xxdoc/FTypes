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
Private Declare Function SafeArrayCreate Lib "oleaut32" (ByVal lType As Integer, ByVal lDims As Long, uBounds As Any) As Long
Private Declare Function SafeArrayDestroy Lib "oleaut32" (ByVal lArray As Long) As Long
Private Declare Function SafeArrayPutElement Lib "oleaut32" (ByVal lArray As Long, uIndices As Any, uValue As Any) As Long
Private Declare Function SafeArrayRedim Lib "oleaut32" (ByVal lArray As Long, uBounds As TSAFEARRAYBOUND) As Long

'EVENTS

Private Sub Class_Terminate()
    
    SafeArrayDestroy p_Array.lPointer

End Sub

'PROPERTIES

Public Property Get Bounds(Optional ByVal Lower As Boolean, Optional ByVal Dimension As Long) As Long
    
    Dim x As Long
    
    If p_Array.lPointer Then
        
        If Dimension > 1& And Dimension <= p_Array.iDims Then x = p_Array.iDims - Dimension
        
        Bounds = p_Array.uBounds(x).lLowest + ((p_Array.uBounds(x).lElements + L_NG) * ((Not Lower) * L_NG))
    
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
        
        If Dimension > 1& And Dimension <= p_Array.iDims Then Elements = p_Array.uBounds(p_Array.iDims - Dimension).lElements Else Elements = p_Array.uBounds(0&).lElements
    
    End If

End Property

Public Property Get ElementSize() As Long
    
    ElementSize = p_Array.lElementSize

End Property

Public Property Let Item(ByVal Index As Long, Optional ByVal Dimension As Long, ByRef NewValue As Variant)
    
    Dim b() As Long
    
    If p_Array.lPointer Then
        
        ReDim b(p_Array.iDims + L_NG)
        
        b(Dimension) = Index
        
        If p_Array.lVarType = vbString Then SafeArrayPutElement p_Array.lPointer, b(0), ByVal StrPtr(NewValue) Else SafeArrayPutElement p_Array.lPointer, b(0), ByVal VarPtr(NewValue) + 8&
    
    End If

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
    
    ArrPtr NewValue, p_Array

End Property

'ROUTINES

Public Function Add(ByRef Range As Variant) As ArrayEx
    
    Set Add = Me

End Function

Public Function Clone() As ArrayEx
    
    Set Clone = New ArrayEx
    
    If p_Array.lPointer Then Clone.Value = Value

End Function

Public Function Create(ByVal ArrayType As VbVarType, ByVal Dimensions As Long, ParamArray Bounds()) As ArrayEx
    
    Dim b() As TSAFEARRAYBOUND
    Dim i As Long
    Dim x As Long
    
    Set Create = Me
    
    If Dimensions > 0& Then
        
        ReDim b(UBound(Bounds))
        
        For i = 0& To UBound(b)
            b(i).lLowest = Bounds(i)(0)
            b(i).lElements = Bounds(i)(1) - b(i).lLowest + 1&
        Next i
        
        i = ArrayType Xor (vbArray * (((ArrayType And vbArray) = vbArray) * L_NG))
        
        x = SafeArrayCreate(i, Dimensions, b(0))
        
        ArrPtr x, p_Array
        
        p_Array.lVarType = i
    
    End If

End Function

Public Function Parse(ByRef Value As Variant) As ArrayEx
    
    Set Parse = Me
    
    Create VarType(Value), 1&, Array(0&, 0&)
    
    RtlMoveMemory ByVal p_Array.lData, ByVal VarPtr(Value) + 8&, LenB(Value)

End Function

Public Function Resize(ByVal NewUpperBound As Long) As ArrayEx
    
    Dim b As TSAFEARRAYBOUND
    Dim x As Long
    
    Set Resize = Me
    
    If p_Array.lPointer Then
        
        x = p_Array.uBounds(0).lLowest
        
        If NewUpperBound >= x Then
            
            b.lLowest = x
            b.lElements = NewUpperBound - x + 1&
            
            SafeArrayRedim p_Array.lPointer, b
            
            p_Array.uBounds(0) = b
        
        End If
    
    End If

End Function

Private Sub ArrPtr(ByRef uArr As Variant, ByRef uPtr As TSAFEARRAY)
    
    Dim j As Long
    Dim x As Long
    
    RtlMoveMemory x, ByVal VarPtr(uArr) + 8&, 4& 'get pointer to array variable
    
    If x Then
        
        RtlMoveMemory j, uArr, 2& 'get variable type
        
        If j = 0& Then 'do not proceed empty variable
            x = 0&
        ElseIf j And 16384& Then 'if variable is passed by reference (pointer to pointer)
            RtlMoveMemory x, ByVal x, 4& 'get real variable pointer
            j = j - 16384& 'remove VT_BYREF flag
        End If
        
        If x Then 'if pointer data is array and is not empty array then
            
            RtlMoveMemory uPtr, ByVal x, 16& 'fill first fixed 16 bytes of structure from pointer
            
            With uPtr
                .lPointer = x
                .lVarType = j - vbArray 'remove VT_ARRAY flag for convenient use
                ReDim .uBounds(.iDims + L_NG) 'allocate structure member
                RtlMoveMemory .uBounds(0), ByVal x + 16&, .iDims * 8& 'fill structure member with array dimensions info (in descending order) bytes starting from member pointer adress + 16 bytes offset
            End With
            
            Exit Sub
        
        End If
    
    End If
    
    RtlMoveMemory uPtr, 0&, 24&

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'VARIABLES

Private p_Value As String
Private p_ValueLen As Long
Private p_ValueLenReal As Long
Private p_ValuePtr As Long

'WINAPI

Private Declare Function CharLowerBuffW Lib "user32" (ByVal lStr As Long, ByVal lLength As Long) As Long
Private Declare Function CharUpperBuffW Lib "user32" (ByVal lStr As Long, ByVal lLength As Long) As Long
Private Declare Function IsTextUnicode Lib "advapi32" (ByRef uBuffer As Any, ByVal lBufferLen As Long, ByRef lResult As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lString As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal lCodePage As Long, ByVal lFlags As Long, uMultiByte As Any, ByVal lMultiByteLen As Long, uWideCharStr As Any, ByVal lWideCharStrLen As Long) As Long

'EVENTS

Private Sub Class_Initialize()
    
    p_ValueLenReal = 256&

End Sub

'PROPERTIES

Public Property Get Capacity() As Long
    
    Capacity = p_ValueLenReal

End Property

Public Property Get Length() As Long
    
    Length = p_ValueLen

End Property

Public Property Get Pointer() As Long
    
    Pointer = p_ValuePtr

End Property

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    If p_ValueLen Then Value = VBA.Left$(p_Value, p_ValueLen)

End Property

Public Property Let Value(ByRef NewValue As String)
    
    Dim l As Long
    
    l = Len(NewValue)
    
    If l Then
        
        BufferRedim l, True
        
        RtlMoveMemory ByVal p_ValuePtr, ByVal StrPtr(NewValue), p_ValueLen * 2&
    
    Else
        p_ValueLen = 0&
    End If

End Property

'ROUTINES

Public Function Asc(Optional ByVal Index As Long) As IntegerEx
    
    Set Asc = New IntegerEx
    
    If p_ValueLen Then
        
        If Index > 0& And Index <= p_ValueLen Then Asc.Value = AscW(Mid$(p_Value, Index, 1&)) Else Asc.Value = AscW(p_Value)
    
    End If

End Function

Public Function Clone() As StringEx
    
    Set Clone = New StringEx
    
    If p_ValueLen Then Clone.Value = VBA.Left$(p_Value, p_ValueLen)

End Function

Public Function Compare(ByRef CompareString As String, Optional ByVal CompareMethod As VbCompareMethod) As IntegerEx
    
    Set Compare = New IntegerEx
    
    If p_ValueLen Then Compare.Value = StrComp(VBA.Left$(p_Value, p_ValueLen), CompareString, CompareMethod)

End Function

Public Function Concat(ByRef ConcatString As String) As StringEx
    
    Dim l As Long
    
    Set Concat = Me
    
    l = Len(ConcatString)
    
    If l Then
        
        BufferRedim l, False
        
        RtlMoveMemory ByVal p_ValuePtr + ((p_ValueLen - l) * 2&), ByVal StrPtr(ConcatString), l * 2&
    
    End If

End Function

Public Function ConcatPointer(ByVal ConcatStringPointer As Long) As StringEx
    
    Dim l As Long
    
    Set ConcatPointer = Me
    
    If ConcatStringPointer Then
        
        l = lstrlenW(ConcatStringPointer)
        
        If l Then
            
            BufferRedim l, False
            
            RtlMoveMemory ByVal p_ValuePtr + ((p_ValueLen - l) * 2&), ByVal ConcatStringPointer, l * 2&
        
        End If
    
    End If

End Function

Public Function Duplicate(ByVal Count As Long) As StringEx
    
    Dim i As Long
    Dim l As Long
    
    Set Duplicate = Me
    
    If p_ValueLen > 0& And Count > 0& Then
        
        l = p_ValueLen
        
        BufferRedim l * Count, False
        
        For i = 1& To Count
            RtlMoveMemory ByVal p_ValuePtr + ((l * i) * 2&), ByVal p_ValuePtr, l * 2&
        Next i
    
    End If

End Function

Public Function Find(ByRef SearchString As String, Optional ByVal Start As Long, Optional ByVal Reverse As Boolean, Optional ByVal CompareMethod As VbCompareMethod) As LongEx
    
    Dim x As Long
    
    Set Find = New LongEx
    
    If p_ValueLen > 0& And Len(SearchString) > 0& Then
        
        If Reverse Then
            
            If Start > 0& And Start < p_ValueLen Then x = Start Else x = p_ValueLen
            
            Find.Value = InStrRev(p_Value, SearchString, x, CompareMethod)
        
        Else
            
            If Start > 0& Then x = Start Else x = 1&
            
            x = InStr(x, p_Value, SearchString, CompareMethod)
            
            If x <= p_ValueLen Then Find.Value = x
        
        End If
    
    End If

End Function

Public Function Insert(ByVal Index As Long, ByRef InsertString As String) As StringEx
    
    Dim l As Long
    
    Set Insert = Me
    
    l = Len(InsertString)
    
    If l > 0& And Index >= 0& And Index <= p_ValueLen Then
        
        BufferRedim l, False
        
        RtlMoveMemory ByVal p_ValuePtr + (Index * 2&) + l + l, ByVal p_ValuePtr + (Index * 2&), (p_ValueLen - (Index + l)) * 2&
        RtlMoveMemory ByVal p_ValuePtr + (Index * 2&), ByVal StrPtr(InsertString), l + l
    
    End If

End Function

Public Function Left(ByVal Length As Long) As StringEx
    
    Set Left = Me
    
    If Length >= 0& And Length < p_ValueLen Then p_ValueLen = Length

End Function

Public Function Lower() As StringEx
    
    Set Lower = Me
    
    If p_ValueLen Then CharLowerBuffW p_ValuePtr, p_ValueLen

End Function

Public Function Numeric() As StringEx
    
    Set Numeric = Me
    
    If p_ValueLen Then Value = ToNumber(VBA.Left$(p_Value, p_ValueLen))

End Function

Public Function PadLeft(ByVal Width As Long) As StringEx
    
    Dim x As Long
    
    Set PadLeft = Me
    
    If p_ValueLen > 0& And Width > p_ValueLen Then
        
        x = Width - p_ValueLen
        
        BufferRedim x, False
        
        RtlMoveMemory ByVal p_ValuePtr + x + x, ByVal p_ValuePtr, (p_ValueLen - x) * 2&
        RtlMoveMemory ByVal p_ValuePtr, ByVal StrPtr(Space$(x)), x + x
    
    End If

End Function

Public Function PadRight(ByVal Width As Long) As StringEx
    
    Dim x As Long
    
    Set PadRight = Me
    
    If p_ValueLen > 0& And Width > p_ValueLen Then
        
        x = Width - p_ValueLen
        
        BufferRedim x, False
        
        RtlMoveMemory ByVal p_ValuePtr + ((p_ValueLen - x) * 2&), ByVal StrPtr(Space$(x)), x + x
    
    End If

End Function

Public Function Parse(ByRef Value As Variant) As StringEx
    
    Set Parse = Me
    
    Dim a As ArrayEx
    Dim b As Integer
    Dim i As Long
    Dim j As Long
    Dim x As Long
    
    x = VarType(Value)
    
    Select Case x
        
        Case vbInteger To vbString, vbError, vbBoolean, vbDecimal, vbByte
            
            Me.Value = Value
        
        Case vbArray + vbByte
            
            Set a = New ArrayEx
            
            a.Value = Value
            
            If a.Dimensions = 1& Then
                
                x = a.Elements 'get items count
                
                If x Mod 2& Then 'odd bytes count (ANSI, UTF-8)
                    
                    If x >= 3& Then
                        
                        If Value(0) = 239 And Value(1) = 187 And Value(2) = 191 Then j = 3& 'check for UTF-8 byte order mark
                        
                        i = MultiByteToWideChar(65001, 8&, ByVal a.Data, x, ByVal 0&, 0&) 'estimate buffer size (with MB_ERR_INVALID_CHARS flag)
                        
                        If i Then '7-bit ANSI or UTF-8
                            
                            BufferRedim i - ((j > 0&) * L_NG), True
                            MultiByteToWideChar 65001, 0&, ByVal a.Data + j, x, ByVal p_ValuePtr, p_ValueLen 'convert multi-bytes
                            
                            Exit Function
                        
                        End If
                    
                    End If
                    
                    BufferRedim x, True
                    RtlMoveMemory ByVal p_ValuePtr, ByVal a.Data, x 'copy bytes as 8-bit ANSI
                
                ElseIf x > 0& Then 'even bytes count (ANSI, UTF-8, UTF-16 BE/LE)
                    
                    Select Case Value(0) 'check if first byte BOM-like
                        
                        Case 255
                            
                            If Value(1) = 254 Then i = 2& 'UTF-16 Little Endian
                        
                        Case 254
                            
                            If Value(1) = 255 Then 'UTF-16 Big Endian
                                
                                For j = 2& To x + L_NG Step 2& 'convert to Little Endian order
                                    RtlMoveMemory b, ByVal a.Data + j, 2&
                                    b = (((b And &HFF00) \ &H100) And &HFF) Or (b And &H7F) * &H100 Or (b And &H80) * &HFF00 'swap integer low & high words
                                    RtlMoveMemory ByVal a.Data + j, b, 2&
                                Next j
                                
                                i = 2&
                            
                            End If
                        
                        Case 239
                            
                            If x >= 3& Then
                                
                                If Value(1) = 187 And Value(2) = 191 Then 'UTF-8
                                    
                                    i = MultiByteToWideChar(65001, 8&, ByVal a.Data, x, ByVal 0&, 0&) 'estimate buffer size (with MB_ERR_INVALID_CHARS flag)
                                    
                                    If i Then
                                        
                                        BufferRedim i + L_NG, True
                                        MultiByteToWideChar 65001, 0&, ByVal a.Data + 3&, x, ByVal p_ValuePtr, p_ValueLen 'convert multi-bytes
                                        
                                        Exit Function
                                    
                                    End If
                                
                                End If
                            
                            End If
                    
                    End Select
                    
                    If i = 0& Then 'no BOM
                        
                        If IsTextUnicode(ByVal a.Data, x, 7&) = 0& Then 'if IS_TEXT_UNICODE_ASCII16, IS_TEXT_UNICODE_STATISTICS and IS_TEXT_UNICODE_CONTROLS tests fail
                            
                            i = MultiByteToWideChar(65001, 8&, ByVal a.Data, x, ByVal 0&, 0&) 'estimate buffer size (with MB_ERR_INVALID_CHARS flag)
                            
                            If i Then '7-bit ANSI or UTF-8
                                BufferRedim i, True
                                MultiByteToWideChar 65001, 0&, ByVal a.Data, x, ByVal p_ValuePtr, p_ValueLen 'convert multi-bytes
                            Else '8-bit ANSI or UTF-16 case (currently difference can not be determined correctly, f.e. Chinese double-byte encoded text is recognized as ANSI)
                                BufferRedim x, True
                                RtlMoveMemory ByVal p_Value, ByVal a.Data, x 'copy bytes as 8-bit ANSI
                            End If
                            
                            Exit Function
                        
                        End If
                    
                    End If
                    
                    BufferRedim (x - i) \ 2&, True
                    RtlMoveMemory ByVal p_ValuePtr, ByVal a.Data + i, x - i 'copy bytes as Unicode
                
                End If
            
            End If
            
            Set a = Nothing
        
        Case Else
            
            p_ValueLen = 0&
    
    End Select

End Function

Public Function Remove(ByVal Index As Long, ByVal Length As Long) As StringEx
    
    Dim x As Long
    
    Set Remove = Me
    
    If p_ValueLen > 0& And Index > 0& And Index <= p_ValueLen And Length > 0& Then
        
        x = Length - (Length - (p_ValueLen - Index + 1&))
        
        If Length <= x Then x = Length
        
        If x < p_ValueLen Then
            
            RtlMoveMemory ByVal p_ValuePtr + ((Index + L_NG) * 2&), ByVal p_ValuePtr + ((Index + L_NG) * 2&) + (x * 2&), (p_ValueLen - (Index + x) + 1&) * 2&
            
            p_ValueLen = p_ValueLen - x
        
        Else
            p_ValueLen = 0&
        End If
    
    End If

End Function

Public Function Replace(ByRef SearchString As String, ByRef ReplaceString As String, Optional ByVal Start As Long = 1&, Optional ByVal Count As Long = L_NG, Optional ByVal CompareMethod As VbCompareMethod) As StringEx
    
    Set Replace = Me
    
    If p_ValueLen Then Value = VBA.Replace(VBA.Left$(p_Value, p_ValueLen), SearchString, ReplaceString, Start, Count, CompareMethod)

End Function

Public Function Right(ByVal Length As Long) As StringEx
    
    Set Right = Me
    
    If Length >= 0& And Length < p_ValueLen Then
        
        RtlMoveMemory ByVal p_ValuePtr, ByVal p_ValuePtr + ((p_ValueLen - Length) * 2&), Length + Length
        
        p_ValueLen = Length
    
    End If

End Function

Public Function Split(Optional ByRef Delimeter As String, Optional ByVal Limit As Long = L_NG, Optional ByVal CompareMethod As VbCompareMethod, Optional ByVal ArrayType As VbVarType = vbString) As ArrayEx
    
    On Error Resume Next
    
    Dim i As Long
    Dim l As Long
    Dim p As Long
    Dim s As String
    Dim x As Long
    Dim y As Long
    
    Set Split = New ArrayEx
    
    If p_ValueLen > 0& And Len(Delimeter) > 0& Then
        
        Split.Create ArrayType, 1&, Array(0&, p_ValueLen + L_NG)
        
        s = Space$(p_ValueLen)
        p = StrPtr(s)
        
        Do
            
            x = InStr(x + 1&, p_Value, Delimeter, CompareMethod)
            
            If x = 0& Or x > p_ValueLen Then x = p_ValueLen + 1&
            
            l = x - y + L_NG
            
            RtlMoveMemory ByVal p, ByVal p_ValuePtr + y + y, l + l
            
            Split.Item(i) = VBA.Left$(s, l)
            
            i = i + 1&
            y = x
        
        Loop While y < p_ValueLen
        
        If Limit > L_NG And Limit < i Then i = Limit Else i = i + L_NG
        
        Split.Resize i
    
    Else
        
        Split.Parse VBA.Left$(p_Value, p_ValueLen)
    
    End If

End Function

Public Function ToArray(Optional ByVal AsANSI As Boolean) As ArrayEx
    
    Dim x As Long
    
    Set ToArray = New ArrayEx
    
    If p_ValueLen Then
        
        x = p_ValueLen * (((Not AsANSI) * L_NG) + 1&)
        
        ToArray.Create vbByte, 1&, Array(0&, x + L_NG)
        
        If AsANSI Then RtlMoveMemory ByVal ToArray.Data, ByVal p_Value, x Else RtlMoveMemory ByVal ToArray.Data, ByVal p_ValuePtr, x
    
    End If

End Function

Public Function ToBoolean() As BooleanEx
    
    Set ToBoolean = New BooleanEx
    
    If p_ValueLen Then ToBoolean.Parse VBA.Left$(p_Value, p_ValueLen)

End Function

Public Function ToByte() As ByteEx
    
    Set ToByte = New ByteEx
    
    If p_ValueLen Then ToByte.Parse VBA.Left$(p_Value, p_ValueLen)

End Function

Public Function ToCurrency() As CurrencyEx
    
    Set ToCurrency = New CurrencyEx
    
    If p_ValueLen Then ToCurrency.Parse VBA.Left$(p_Value, p_ValueLen)

End Function

Public Function ToDecimal() As DecimalEx
    
    Set ToDecimal = New DecimalEx
    
    If p_ValueLen Then ToDecimal.Parse VBA.Left$(p_Value, p_ValueLen)

End Function

Public Function ToDouble() As DoubleEx
    
    Set ToDouble = New DoubleEx
    
    If p_ValueLen Then ToDouble.Parse VBA.Left$(p_Value, p_ValueLen)

End Function

Public Function ToInteger() As IntegerEx
    
    Set ToInteger = New IntegerEx
    
    If p_ValueLen Then ToInteger.Parse VBA.Left$(p_Value, p_ValueLen)

End Function

Public Function ToLong() As LongEx
    
    Set ToLong = New LongEx
    
    If p_ValueLen Then ToLong.Parse VBA.Left$(p_Value, p_ValueLen)

End Function

Public Function TrimL() As StringEx
    
    Dim x As Long
    
    Set TrimL = Me
    
    If p_ValueLen Then
        
        If AscW(p_Value) = 32 Then
            
            Do
                x = x + 1&
            Loop While InStr(x + 1&, p_Value, ChrW$(32&), vbBinaryCompare) = (x + 1&)
            
            If x Then
                
                p_ValueLen = p_ValueLen - x
                
                RtlMoveMemory ByVal p_ValuePtr, ByVal p_ValuePtr + (x * 2&), p_ValueLen * 2&
            
            End If
        
        End If
    
    End If

End Function

Public Function TrimR() As StringEx
    
    Set TrimR = Me
    
    If p_ValueLen Then
        
        If AscW(Mid$(p_Value, p_ValueLen, 1&)) = 32 Then
            
            Do
                p_ValueLen = p_ValueLen + L_NG
            Loop While InStrRev(p_Value, ChrW$(32&), p_ValueLen, vbBinaryCompare) = p_ValueLen
        
        End If
    
    End If

End Function

Public Function Upper() As StringEx
    
    Set Upper = Me
    
    If p_ValueLen Then CharUpperBuffW p_ValuePtr, p_ValueLen

End Function

Private Sub BufferRedim(ByVal NewStringLen As Long, ByVal FromScratch As Boolean)
    
    Dim b As String
    Dim r As Boolean
    Dim x As Long
    
    p_ValueLen = (p_ValueLen * ((Not FromScratch) * L_NG)) + NewStringLen
    
    If p_ValueLen > (p_ValueLenReal \ 2&) Then p_ValueLenReal = p_ValueLenReal + (p_ValueLen * 2&) Else r = p_ValuePtr
    
    If Not r Then
        
        If FromScratch Then
            
            p_Value = Space$(p_ValueLenReal)
            p_ValuePtr = StrPtr(p_Value)
        
        Else
            
            x = p_ValueLen - NewStringLen
            
            If x Then b = VBA.Left$(p_Value, x)
            
            p_Value = Space$(p_ValueLenReal)
            p_ValuePtr = StrPtr(p_Value)
            
            If x Then RtlMoveMemory ByVal p_ValuePtr, ByVal StrPtr(b), x * 2&
        
        End If
    
    End If

End Sub

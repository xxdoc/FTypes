Attribute VB_Name = "mMain"
Option Explicit

'TYPES

Private Type TPOINT
    x As Long
    y As Long
End Type

Public Type TSAFEARRAY
    iDims As Integer
    iFeatures As Integer
    lElementSize As Long
    lLocks As Long
    lData As Long
    lVarType As Long
    uBounds() As TPOINT
End Type

'CONSTANTS

Public Const B_MX As Byte = 255
Public Const D_MN As Double = 1E+308
Public Const D_MX As Double = -1E+308
Public Const I_MN As Integer = -32768
Public Const I_MX As Integer = 32767
Public Const L_NG As Long = -1&
Public Const L_MN As Long = &H80000000
Public Const L_MX As Long = 2147483647

'VARIABLES

Private gUnc As Boolean

'WINAPI

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetLocaleInfoA Lib "kernel32" (ByVal lLocale As Long, ByVal lType As Long, ByVal sBuffer As String, ByVal lBufferLen As Long) As Long
Private Declare Function GetLocaleInfoW Lib "kernel32" (ByVal lLocale As Long, ByVal lType As Long, ByVal lBuffer As Long, ByVal lBufferLen As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32" (ByVal lhWnd As Long) As Long

Public Declare Sub RtlMoveMemory Lib "kernel32" (uTarget As Any, uSource As Any, ByVal lLen As Long)

'ROUTINES

Private Sub Main()
    
    gUnc = IsWindowUnicode(GetDesktopWindow)

End Sub

Public Function RArrPt(ByRef uArr As Variant, ByRef uPtr As TSAFEARRAY) As Long
    
    '------------------------------------------------------------------------------------------------------------------------------------------'
    '
    ' PURPOSE   : Fill structure with array information.
    '
    ' RETURN    : Array dimensions count.
    '
    ' ARGUMENTS : uArr - source array
    '             uPtr - returns array information structure
    '
    ' NOTES     : Expected array declare syntax - Dim Array(dim1 (cols), [dim2 (rows)], [dim 3], ... , [dim N]) As Type
    '             Array item data adress = first item data start adress + ((col + row + ((cols - 1) * row)) * item size in bytes).
    '
    '------------------------------------------------------------------------------------------------------------------------------------------'
    
    Dim i As Long
    Dim j As Long
    Dim x As Long
    
    RtlMoveMemory x, ByVal VarPtr(uArr) + 8&, 4& 'get pointer to array variable
    
    If x Then
        
        RtlMoveMemory j, uArr, 2& 'get variable type
        
        If j = 0& Then 'do not proceed empty variable
            Exit Function
        ElseIf j And 16384& Then 'if variable is passed by reference (pointer to pointer)
            RtlMoveMemory x, ByVal x, 4& 'get real variable pointer
            j = j - 16384& 'remove VT_BYREF flag
        End If
        
        If x Then 'if pointer data is array and is not empty array then
            
            RtlMoveMemory uPtr, ByVal x, 16& 'fill first fixed 16 bytes of structure from pointer
            
            With uPtr
                .lVarType = j - 8192& 'remove VT_ARRAY flag for convenient use
                i = .iDims 'get array dimensions count
                ReDim .uBounds(i + L_NG) 'allocate structure member
                RtlMoveMemory .uBounds(0), ByVal x + 16&, i * 8& 'fill structure member with array dimensions info (in descending order) bytes starting from member pointer adress + 16 bytes offset
            End With
            
            RArrPt = i 'return dimensions count
        
        End If
    
    End If

End Function

Public Function ToNumber(ByVal sVal As String) As Double
    
    '------------------------------------------------------------------------------------------------------------------------------------------'
    '
    ' PURPOSE   : Convert any string expression to DOUBLE type.
    '
    ' RETURN    : DOUBLE type numeric value.
    '
    ' ARGUMENTS : sVal - source string
    '
    ' NOTES     :
    '
    '------------------------------------------------------------------------------------------------------------------------------------------'
    
    Dim b As Byte
    Dim d As Long
    Dim f As Long
    Dim i As Long
    Dim p As Long
    Dim s As String
    Dim t As String
    Dim x As Long
    Dim y As Long
    
    i = Len(sVal)
    
    If i Then
        
        If IsNumeric(sVal) Then
            ToNumber = sVal
        Else
            
            s = Space$(i)
            p = StrPtr(s)
            t = ChrW$(32&)
            y = StrPtr(t)
            x = StrPtr(sVal)
            
            If gUnc Then GetLocaleInfoW 1024&, 14&, y, 2& Else GetLocaleInfoA 1024&, 14&, t, 1&
            d = InStr(1&, sVal, t, 0&)
            
            For i = 1& To i
                
                RtlMoveMemory ByVal y, ByVal x, 2&
                b = AscB(t)
                
                If (b > 47 And b < 58) Or i = d Then
                    
                    RtlMoveMemory ByVal p, ByVal x, 2&
                    p = p + 2&
                    
                    If f = 0& Then
                        If i <> d Then f = i
                    End If
                
                End If
                
                x = x + 2&
            
            Next i
            
            If gUnc Then GetLocaleInfoW 1024&, 81&, y, 2& Else GetLocaleInfoA 1024&, 81&, t, 1&
            x = InStr(1&, sVal, t, 0&)
            
            If x > 0& And x < f And ((d > 0& And x < d) Or d = 0&) Then ToNumber = t & ChrW$(48&) & Left$(s, 308&) Else ToNumber = ChrW$(48&) & Left$(s, 308&)
        
        End If
    
    End If

End Function

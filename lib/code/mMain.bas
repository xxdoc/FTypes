Attribute VB_Name = "mMain"
Option Explicit

'TYPES

Public Type TSAFEARRAYBOUND
    lElements As Long
    lLowest As Long
End Type

Public Type TSAFEARRAY
    iDims As Integer
    iFeatures As Integer
    lElementSize As Long
    lLocks As Long
    lData As Long
    lPointer As Long
    lVarType As Long
    uBounds() As TSAFEARRAYBOUND
End Type

'CONSTANTS

Public Const L_NG As Long = -1&

'WINAPI

Private Declare Function GetLocaleInfoW Lib "kernel32" (ByVal lLocale As Long, ByVal lType As Long, ByVal lBuffer As Long, ByVal lBufferLen As Long) As Long

Public Declare Sub RtlMoveMemory Lib "kernel32" (uTarget As Any, uSource As Any, ByVal lLen As Long)

'ROUTINES

Public Function ToNumber(ByVal sVal As String) As Double
    
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
            
            GetLocaleInfoW 1024&, 14&, y, 2&
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
            
            GetLocaleInfoW 1024&, 81&, y, 2&
            x = InStr(1&, sVal, t, 0&)
            
            If x > 0& And x < f And ((d > 0& And x < d) Or d = 0&) Then ToNumber = t & ChrW$(48&) & Left$(s, 308&) Else ToNumber = ChrW$(48&) & Left$(s, 308&)
        
        End If
    
    End If

End Function

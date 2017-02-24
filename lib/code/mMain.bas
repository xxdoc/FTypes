Attribute VB_Name = "mMain"
Option Explicit

'LICENSE AGREEMENTS

'This software Is provided 'as-is', without any express or implied warranty. In no event will the author be held liable for any damages arising from the use of this software.
'Permission is granted to anyone to use this software for any purpose, including commercial applications, and to alter it and redistribute it freely, subject to the following restrictions:
'1. The origin of this software must not be misrepresented; you must not claim that you wrote the original software. If you use this software in a product, an acknowledgment in the product documentation would be appreciated but is not required.
'2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original software.

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

'VARIABLES

Public PUB_UNICODE As Boolean
Public m_Comma As Byte
Private m_Minus As Byte

'WINAPI

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetLocaleInfoA Lib "kernel32" (ByVal lLocale As Long, ByVal lType As Long, ByVal sBuffer As String, ByVal lBufferLen As Long) As Long
Private Declare Function GetLocaleInfoW Lib "kernel32" (ByVal lLocale As Long, ByVal lType As Long, ByVal lBuffer As Long, ByVal lBufferLen As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32" (ByVal lhWnd As Long) As Long
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal lSource As Long, bTarget As Byte)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal lTarget As Long, ByVal bSource As Byte)

Public Declare Sub RtlMoveMemory Lib "kernel32" (uTarget As Any, uSource As Any, ByVal lLen As Long)

'ROUTINES

Private Sub Main()
    
    Dim s As String
    Dim p As Long
    
    s = Space$(1&)
    p = StrPtr(s)
    
    PUB_UNICODE = IsWindowUnicode(GetDesktopWindow)
    
    If PUB_UNICODE Then
        
        If GetLocaleInfoW(1024&, 14&, p, 2&) Then m_Comma = AscB(s)
        If GetLocaleInfoW(1024&, 81&, p, 2&) Then m_Minus = AscB(s)
    
    Else
        
        If GetLocaleInfoA(1024&, 14&, s, 2&) Then m_Comma = AscB(s)
        If GetLocaleInfoA(1024&, 81&, s, 2&) Then m_Minus = AscB(s)
    
    End If

End Sub

Public Function ToNumber(ByVal sVal As String) As Double
    
    Dim b As Byte
    Dim c As Long
    Dim i As Long
    Dim f As Boolean
    Dim m As Boolean
    Dim p1 As Long
    Dim p2 As Long
    Dim s As String
    
    i = Len(sVal)
    
    If i Then
        
        If IsNumeric(sVal) Then
            
            ToNumber = sVal
        
        Else
            
            s = Space$(i + 1&)
            p1 = StrPtr(sVal)
            p2 = StrPtr(s)
            
            For i = 1& To i
                
                GetMem1 p1, b
                
                If b > 47 And b < 58 Then
                    
                    PutMem1 p2, b
                    
                    c = c + 1&
                    p2 = p2 + 2&
                
                ElseIf Not f And b = m_Comma Then
                    
                    PutMem1 p2, b
                    
                    p2 = p2 + 2&
                    
                    f = True
                
                ElseIf c = 0& And Not f And Not m And b = m_Minus Then
                    
                    PutMem1 p2, b
                    PutMem1 p2 + 2&, 48
                    
                    p2 = p2 + 4&
                    
                    m = True
                
                End If
                
                p1 = p1 + 2&
            
            Next i
            
            If f Or m Or c Then ToNumber = Left$(s, 308&)
        
        End If
    
    End If

End Function

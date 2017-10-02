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

Public Const L_MX As Long = 2147483647
Public Const L_NG As Long = -1&

'VARIABLES

Public PUB_UNICODE As Boolean
Public m_Comma As Integer
Private m_Minus As Integer

'WINAPI

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetLocaleInfoA Lib "kernel32" (ByVal lLocale As Long, ByVal lType As Long, ByVal sBuffer As String, ByVal lBufferLen As Long) As Long
Private Declare Function GetLocaleInfoW Lib "kernel32" (ByVal lLocale As Long, ByVal lType As Long, ByVal lBuffer As Long, ByVal lBufferLen As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32" (ByVal lhWnd As Long) As Long
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal lSource As Long, iTarget As Integer)
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal lTarget As Long, ByVal iSource As Integer)

Public Declare Sub RtlMoveMemory Lib "kernel32" (uTarget As Any, uSource As Any, ByVal lLen As Long)

'ROUTINES

Private Sub Main()
    
    Dim s As String
    Dim p As Long
    
    s = Space$(1&)
    p = StrPtr(s)
    
    PUB_UNICODE = IsWindowUnicode(GetDesktopWindow)
    
    If PUB_UNICODE Then
        
        If GetLocaleInfoW(1024&, 14&, p, 2&) Then m_Comma = AscW(s)
        If GetLocaleInfoW(1024&, 81&, p, 2&) Then m_Minus = AscW(s)
    
    Else
        
        If GetLocaleInfoA(1024&, 14&, s, 2&) Then m_Comma = AscW(s)
        If GetLocaleInfoA(1024&, 81&, s, 2&) Then m_Minus = AscW(s)
    
    End If

End Sub

Public Function ToNumber(ByVal lPointer As Long, ByVal lLength As Long) As String
    
    Dim b As Integer
    Dim c As Long
    Dim i As Long
    Dim f As Boolean
    Dim m As Boolean
    Dim p As Long
    
    If lLength Then
        
        ToNumber = Space$(lLength)
        p = StrPtr(ToNumber)
        
        For i = 1& To lLength
            
            GetMem2 lPointer, b
            
            If b > 47 And b < 58 Then
                
                PutMem2 p, b
                
                c = c + 1&
                p = p + 2&
            
            ElseIf Not f And b = m_Comma Then
                
                PutMem2 p, b
                
                c = c + 1&
                p = p + 2&
                
                f = True
            
            ElseIf c = 0& And Not f And Not m And b = m_Minus Then
                
                PutMem2 p, b
                
                c = c + 1&
                p = p + 2&
                
                m = True
            
            End If
            
            lPointer = lPointer + 2&
        
        Next i
        
        If c > 308& Then
            
            ToNumber = Left$(ToNumber, 308&)
        
        ElseIf c = 1& Then
            
            If AscW(ToNumber) = m_Comma Or AscW(ToNumber) = m_Minus Then ToNumber = ChrW$(48&)
        
        ElseIf c > 0& Then
            
            ToNumber = Left$(ToNumber, c)
        
        Else
            
            ToNumber = ChrW$(48&)
        
        End If
    
    Else
        ToNumber = ChrW$(48&)
    End If

End Function

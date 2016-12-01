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

'VARIABLES

Public m_Comma As Byte
Private m_Minus As Byte

'WINAPI

Private Declare Function GetLocaleInfoW Lib "kernel32" (ByVal lLocale As Long, ByVal lType As Long, ByVal lBuffer As Long, ByVal lBufferLen As Long) As Long
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal lSource As Long, bTarget As Byte)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal lTarget As Long, ByVal bSource As Byte)

Public Declare Sub RtlMoveMemory Lib "kernel32" (uTarget As Any, uSource As Any, ByVal lLen As Long)

'ROUTINES

Private Sub Main()
    
    Dim s As String
    Dim p As Long
    
    s = Space$(1&)
    p = StrPtr(s)
    
    If GetLocaleInfoW(1024&, 14&, p, 2&) Then m_Comma = AscB(s)
    If GetLocaleInfoW(1024&, 81&, p, 2&) Then m_Minus = AscB(s)

End Sub

Public Function ToNumber(ByVal sVal As String) As String
    
    Dim b As Byte
    Dim i As Long
    Dim f As Boolean
    Dim p1 As Long
    Dim p2 As Long
    
    i = Len(sVal)
    
    If i Then
        
        If IsNumeric(sVal) Then
            
            ToNumber = sVal
        
        Else
            
            ToNumber = Space$(i)
            p1 = StrPtr(sVal)
            p2 = StrPtr(ToNumber)
            
            For i = 1& To i
                
                GetMem1 p1, b
                
                If (b > 47 And b < 58) Or (Not f And b = m_Comma) Or (i = 1& And b = m_Minus) Then
                    
                    PutMem1 p2, b
                    
                    p2 = p2 + 2&
                    
                    If b = m_Comma Then f = True
                
                End If
                
                p1 = p1 + 2&
            
            Next i
            
            ToNumber = Left$(ToNumber, 308&)
        
        End If
    
    Else
        
        ToNumber = ChrW$(48&)
    
    End If

End Function

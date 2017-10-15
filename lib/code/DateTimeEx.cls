VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DateTimeEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'INTERFACES

Implements IFTypesBasic

'TYPES

Private Type TSYSTEMTIME
    iYear As Integer
    iMonth As Integer
    iDayOfWeek As Integer
    iDay As Integer
    iHour As Integer
    iMinute As Integer
    iSecond As Integer
    iMilliseconds As Integer
End Type

'VARIABLES

Private p_Value As TSYSTEMTIME

'WINAPI

Private Declare Sub GetLocalTime Lib "kernel32" (uSystemTime As TSYSTEMTIME)
Private Declare Sub GetSystemTime Lib "kernel32" (uSystemTime As TSYSTEMTIME)

'PROPERTIES

Private Property Get IFTypesBasic_Length() As Long
    
    IFTypesBasic_Length = 16&

End Property

Private Property Get IFTypesBasic_Pointer() As Long
    
    IFTypesBasic_Pointer = VarPtr(p_Value)

End Property

Public Property Get Day() As Integer
    
    Day = p_Value.iDay

End Property

Public Property Get Hour() As Integer
    
    Hour = p_Value.iHour

End Property

Public Property Get IsLeap() As Boolean
    
    IsLeap = (DateTime.Month(DateTime.DateSerial(p_Value.iYear, 2, 29)) = 2)

End Property

Public Property Get Length() As Long
    
    Length = IFTypesBasic_Length

End Property

Public Property Get Max() As Date
    
    Max = DateTime.DateSerial(9999, 12, 31)

End Property

Public Property Get Millisecond() As Integer
    
    Millisecond = p_Value.iMilliseconds

End Property

Public Property Get Min() As Date
    
    Min = DateTime.DateSerial(100, 1, 1)

End Property

Public Property Get Minute() As Integer
    
    Minute = p_Value.iMinute

End Property

Public Property Get Month() As Integer
    
    Month = p_Value.iMonth

End Property

Public Property Get MonthDays() As Long
    
    MonthDays = DateTime.Day(DateTime.DateSerial(p_Value.iYear, p_Value.iMonth + 1, 1) + L_NG)

End Property

Public Property Get Pointer() As Long
    
    Pointer = IFTypesBasic_Pointer

End Property

Public Property Get Second() As Integer
    
    Second = p_Value.iSecond

End Property

Public Property Get Value() As Date
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    
    On Error Resume Next
    
    With p_Value
        
        If .iYear > 0 And .iMonth > 0 And .iDay > 0 Then
            
            If .iHour > 0 Or .iMinute > 0 Or .iSecond > 0 Then Value = DateTime.DateSerial(.iYear, .iMonth, .iDay) & ChrW$(32&) & DateTime.TimeSerial(.iHour, .iMinute, .iSecond) Else Value = DateTime.DateSerial(.iYear, .iMonth, .iDay)
        
        Else
            
            If .iHour > 0 Or .iMinute > 0 Or .iSecond > 0 Then Value = DateTime.TimeSerial(.iHour, .iMinute, .iSecond)
        
        End If
    
    End With

End Property

Public Property Let Value(ByVal NewValue As Date)
    
    With p_Value
        .iDay = DateTime.Day(NewValue)
        .iDayOfWeek = DateTime.Weekday(NewValue, vbUseSystemDayOfWeek)
        .iHour = DateTime.Hour(NewValue)
        .iMinute = DateTime.Minute(NewValue)
        .iMonth = DateTime.Month(NewValue)
        .iSecond = DateTime.Second(NewValue)
        .iYear = DateTime.Year(NewValue)
    End With

End Property

Public Property Get Weekday() As Integer
    
    Weekday = p_Value.iDayOfWeek

End Property

Public Property Get Year() As Integer
    
    Year = p_Value.iYear

End Property

Public Property Get YearDayIndex() As Long
    
    YearDayIndex = Abs(DateTime.DateDiff(ChrW$(100&), DateTime.DateSerial(p_Value.iYear, p_Value.iMonth, p_Value.iDay), DateSerial(p_Value.iYear + L_NG, 12, 31), vbUseSystemDayOfWeek, vbUseSystem))

End Property

'METHODS

Private Function IFTypesBasic_Parse(Value As Variant) As IFTypesBasic
    
    Set IFTypesBasic_Parse = Me
    
    Select Case VarType(Value)
        
        Case vbInteger To vbDate, vbBoolean, vbByte, vbDecimal
            
            Me.Value = Value
        
        Case vbString
            
            If IsDate(Value) Then Me.Value = Value
        
        Case Else
            
            With p_Value
                .iDay = 0
                .iDayOfWeek = 0
                .iHour = 0
                .iMilliseconds = 0
                .iMinute = 0
                .iMonth = 0
                .iSecond = 0
                .iYear = 0
            End With
    
    End Select

End Function

Private Function IFTypesBasic_ToList() As IFTypesArray
    
    Set IFTypesBasic_ToList = New ListEx
    
    IFTypesBasic_ToList.Parse Me.Value

End Function

Public Function AddDays(ByVal Count As Integer) As DateTimeEx
    
    Set AddDays = Me
    
    p_Value.iDay = p_Value.iDay + Count

End Function

Public Function AddHours(ByVal Count As Integer) As DateTimeEx
    
    Set AddHours = Me
    
    p_Value.iHour = p_Value.iHour + Count

End Function

Public Function AddMilliseconds(ByVal Count As Integer) As DateTimeEx
    
    Set AddMilliseconds = Me
    
    p_Value.iMilliseconds = p_Value.iMilliseconds + Count

End Function

Public Function AddMinutes(ByVal Count As Integer) As DateTimeEx
    
    Set AddMinutes = Me
    
    p_Value.iMinute = p_Value.iMinute + Count

End Function

Public Function AddMonths(ByVal Count As Integer) As DateTimeEx
    
    Set AddMonths = Me
    
    p_Value.iMonth = p_Value.iMonth + Count

End Function

Public Function AddSeconds(ByVal Count As Integer) As DateTimeEx
    
    Set AddSeconds = Me
    
    p_Value.iSecond = p_Value.iSecond + Count

End Function

Public Function AddYears(ByVal Count As Integer) As DateTimeEx
    
    Set AddYears = Me
    
    p_Value.iYear = p_Value.iYear + Count

End Function

Public Function Clone() As DateTimeEx
    
    Set Clone = New DateTimeEx
    
    Clone.Value = Me.Value

End Function

Public Function DateSerial(ByVal Year As Integer, ByVal Month As Integer, ByVal Day As Integer) As DateTimeEx
    
    Set DateSerial = Me
    
    With p_Value
        .iYear = Year
        .iMonth = Month
        .iDay = Day
    End With

End Function

Public Function Now(Optional ByVal UTC As Boolean) As DateTimeEx
    
    Set Now = Me
    
    If UTC Then GetSystemTime p_Value Else GetLocalTime p_Value

End Function

Public Function Parse(ByRef Value As Variant) As DateTimeEx
    
    Set Parse = IFTypesBasic_Parse(Value)

End Function

Public Function TimeSerial(ByVal Hour As Integer, ByVal Minute As Integer, ByVal Second As Integer, Optional ByVal Millisecond As Integer) As DateTimeEx
    
    Set TimeSerial = Me
    
    With p_Value
        .iHour = Hour
        .iMinute = Minute
        .iSecond = Second
        .iMilliseconds = Millisecond
    End With

End Function

Public Function ToList() As ListEx
    
    Set ToList = IFTypesBasic_ToList

End Function

Public Function ToLong() As LongEx
    
    Set ToLong = New LongEx
    
    ToLong.Value = Me.Value

End Function

Public Function ToString(Optional ByRef Format As String) As StringEx
    
    Set ToString = New StringEx
    
    If Len(Format) Then ToString.Parse(Strings.Format$(Me.Value, Format, vbUseSystemDayOfWeek, vbUseSystem)).Replace ChrW$(102&), (p_Value.iMilliseconds), , 1& Else ToString.Value = Me.Value

End Function

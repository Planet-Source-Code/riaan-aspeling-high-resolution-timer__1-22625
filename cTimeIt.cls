VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTimeIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'API's to use
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Any) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Any) As Long

'Only three variables to use in this class
Dim Freq As Currency, TStart As Currency, TEnd As Currency

Public Property Get Frequency() As Currency
    'This property only returns the clock frequency
    Frequency = Freq
End Property

Private Sub Class_Initialize()
    'Reset the Start and End timer variables
    TStart = 0
    TEnd = 0
    'Return the clock frequency
    QueryPerformanceFrequency Freq
    If Freq <= 0 Then
        'If for some reason the frequency is less than zero we are dealing with
        'a high resolution clock. This is not supported with this class.
        MsgBox "This timer class doesn't support high frequency clock's!"
        Freq = 0
    End If
End Sub

Public Sub Start()
    'Return the current clock count
    QueryPerformanceCounter TStart
    'Reset the end result
    TEnd = 0
End Sub

Public Sub StopNow()
    'Return the current clock count
    QueryPerformanceCounter TEnd
End Sub

Public Function Result() As Currency
    'Do not return anything if the "StopNow" method is not called.
    If TEnd = 0 Then
        Result = -1
        Exit Function
    End If
    'Double check to make sure the Frequency is not zero
    If Freq <> 0 Then
        'The time it took will equal the end time recorded, subtracting the start time
        'and deviding by the frequency of the clock.
        Result = (TEnd - TStart) / Freq
    Else
        'Return -1 if the clock frequency is to high
        Result = -1
    End If
End Function

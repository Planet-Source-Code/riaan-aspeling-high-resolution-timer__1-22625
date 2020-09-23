VERSION 5.00
Begin VB.Form frmTimerDemo 
   Caption         =   "Test Timer Class"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      Height          =   2835
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   450
      Width           =   3855
   End
   Begin VB.CommandButton btnTest 
      Caption         =   "Run Timer"
      Height          =   825
      Left            =   4380
      TabIndex        =   0
      Top             =   540
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "Results:"
      Height          =   285
      Left            =   300
      TabIndex        =   2
      Top             =   180
      Width           =   2655
   End
End
Attribute VB_Name = "frmTimerDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TTime As New cTimeIt

Private Sub btnTest_Click()
    Dim iCount As Long, sFunnyResult As String
    
    'Reset some variables
    sFunnyResult = "TEST"
    iFunnyResult = 0
    iCount = 0
    
    'Start the timer
    TTime.Start
    For iCount = 1 To 1000000
        'This should take more time than the Len(x) <> 0
        'This is actualy a good example for optimizing code. Always rather compare
        'the lenght of a variable than to compare it with "" (string nothing). The
        'len function normaly executes 20%-25% faster.
        If sFunnyResult <> "" Then
            iFunnyResult = iFunnyResult + 1
        End If
    Next
    'Stop the timer
    TTime.StopNow
    
    'Return the results
    txtResult = txtResult & "For Next Loop with (x <> "") took : " & TTime.Result & " sec to execute." & vbCrLf & vbCrLf
    
    'Reset some variables
    sFunnyResult = "TEST"
    iFunnyResult = 0
    iCount = 0
    
    'Start the timer
    TTime.Start
    For iCount = 1 To 1000000
        'Now compare the len of string to 0
        If Len(sFunnyResult) <> 0 Then
            iFunnyResult = iFunnyResult + 1
        End If
    Next
    'Stop the timer
    TTime.StopNow
    
    'Return the results
    txtResult = txtResult & "For Next Loop with (Len(x) <> 0) took : " & TTime.Result & " sec to execute." & vbCrLf & vbCrLf
    
End Sub

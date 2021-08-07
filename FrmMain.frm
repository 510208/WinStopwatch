VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ShitTimer"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   480
      Top             =   2760
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   2760
   End
   Begin VB.CommandButton CmdStartStop 
      Caption         =   "Start"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      ToolTipText     =   "Ctrl + Shift + F12"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label LblDisplay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00.000"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Type SYSTEMTIME
wYear As Integer
wMonth As Integer
wDayOfWeek As Integer
wDay As Integer
wHour As Integer
wMinute As Integer
wSecond As Integer
wMilliseconds As Integer
End Type

Private StartTime As SYSTEMTIME
Private IsRunning As Boolean
Private HotKeyPressed As Boolean

Private Sub CmdStartStop_Click()
    StartStop
End Sub

Private Sub Timer1_Timer()
    Update
End Sub

Private Sub Timer2_Timer()
    Dim HotKeyPressedNew
    HotKeyPressedNew = GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyShift) And GetAsyncKeyState(vbKeyF12)
    If HotKeyPressedNew And Not HotKeyPressed Then
        StartStop
        Me.Show
    End If
    HotKeyPressed = HotKeyPressedNew
End Sub

Private Sub StartStop()
    IsRunning = Not IsRunning
    If IsRunning Then
        GetSystemTime StartTime
        Timer1.Enabled = True
        CmdStartStop.Caption = "Stop"
    Else
        Update
        Timer1.Enabled = False
        CmdStartStop.Caption = "Restart"
    End If
End Sub

Private Sub Update()
    Dim CurrentTime As SYSTEMTIME
    GetSystemTime CurrentTime
    Dim Minutes As Integer, Seconds As Integer, Milliseconds As Integer

    Milliseconds = CurrentTime.wMilliseconds - StartTime.wMilliseconds
    Seconds = CurrentTime.wSecond - StartTime.wSecond
    Minutes = CurrentTime.wMinute - StartTime.wMinute
    If Milliseconds < 0 Then
        Milliseconds = Milliseconds + 1000
        Seconds = Seconds - 1
    End If
    If Seconds < 0 Then
        Seconds = Seconds + 60
        Minutes = Minutes - 1
    End If
    If Minutes < 0 Then
        Minutes = Minutes + 60
    End If
    LblDisplay.Caption = Format(Minutes, String(2, "0")) & ":" & Format(Seconds, String(2, "0")) & "." & Format(Milliseconds, String(3, "0"))
End Sub

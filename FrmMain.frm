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
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.ComboBox CmbHotKey 
      Height          =   300
      ItemData        =   "FrmMain.frx":0000
      Left            =   2520
      List            =   "FrmMain.frx":015A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.ComboBox CmbHotKeyModifiers 
      Height          =   300
      ItemData        =   "FrmMain.frx":02EC
      Left            =   360
      List            =   "FrmMain.frx":030E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2400
      Width           =   1905
   End
   Begin VB.Timer TmrHotKey 
      Interval        =   10
      Left            =   480
      Top             =   2760
   End
   Begin VB.Timer TmrUpdate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   2760
   End
   Begin VB.CommandButton CmdStartStop 
      Caption         =   "Start"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1560
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
      TabIndex        =   0
      Top             =   480
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

Private HotKey As Integer
Private HotKeyModifiers As New Collection
Private HotKeyPressed As Boolean

Private Sub CmbHotKey_Click()
    HotKey = CmbHotKey.ItemData(CmbHotKey.ListIndex)
End Sub

Private Sub CmbHotKeyModifiers_Click()
    While HotKeyModifiers.Count > 0
        HotKeyModifiers.Remove 1
    Wend

    Dim Chosen As Integer
    Chosen = CmbHotKeyModifiers.ItemData(CmbHotKeyModifiers.ListIndex)

    If Chosen And 1 Then HotKeyModifiers.Add vbKeyShift
    If Chosen And 2 Then HotKeyModifiers.Add vbKeyControl
    If Chosen And 4 Then HotKeyModifiers.Add vbKeyMenu
    If Chosen And 8 Then HotKeyModifiers.Add vbKeyCapital
End Sub

Private Sub CmdStartStop_Click()
    StartStop
End Sub

Private Sub Form_Click()
    CmdStartStop.SetFocus
End Sub

Private Sub Form_Load()
    CmbHotKeyModifiers.ListIndex = 3
    CmbHotKeyModifiers_Click
    CmbHotKey.ListIndex = 12
    CmbHotKey_Click
End Sub

Private Sub TmrUpdate_Timer()
    Update
End Sub

Private Sub TmrHotKey_Timer()
    If HotKey = 0 And HotKeyModifiers.Count = 0 Then Exit Sub

    Dim HotKeyPressedNew As Boolean
    HotKeyPressedNew = True
    If HotKey <> 0 And GetAsyncKeyState(HotKey) = 0 Then
        HotKeyPressedNew = False
    Else
        Dim I As Integer, Key As Integer
        For I = 1 To HotKeyModifiers.Count
            Key = HotKeyModifiers.Item(I)
            If GetAsyncKeyState(Key) = 0 Then
                HotKeyPressedNew = False
                Exit For
            End If
        Next
    End If
    
    If HotKeyPressedNew And Not HotKeyPressed Then
        StartStop
    End If
    HotKeyPressed = HotKeyPressedNew
End Sub

Private Sub StartStop()
    IsRunning = Not IsRunning
    If IsRunning Then
        GetSystemTime StartTime
        TmrUpdate.Enabled = True
        CmdStartStop.Caption = "Stop"
    Else
        Update
        TmrUpdate.Enabled = False
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


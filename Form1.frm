VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Panel 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":058A
   ScaleHeight     =   6390
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog ComD1 
      Left            =   120
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton OpenS 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Set Alarm Sound"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bounce Around"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   15
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TEST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   1440
      Top             =   2280
   End
   Begin ComctlLib.Slider Slider2 
      Height          =   630
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1111
      _Version        =   327682
      Max             =   59
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1085
      _Version        =   327682
      Max             =   23
   End
   Begin VB.CommandButton ARMALARM 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ARMED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton SetTime 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Set Alarm Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton StpAlarm 
      BackColor       =   &H000000FF&
      Caption         =   "TERMINATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   3735
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1080
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   720
      Top             =   2280
   End
   Begin VB.Timer CoolTim 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   2640
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   720
      Top             =   2640
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   17
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "24 Hour Clock"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   2910
      Width           =   7095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Adjust Alarm Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2640
      Width           =   7095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "MINUTE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   10
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "HOUR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   7
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Alarm Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sleep Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   13
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "Panel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'                           Alarm Clock - By Ken Slater - 0x34 - 2007
'                                    Free, Open Source Code
'                                             Enjoy
'
'                                      lpd52@sbcglobal.net

Dim cTime As String
Dim aHR As String
Dim aMN As String
Dim aHRi As Integer
Dim aMNi As Integer
Dim aTime As String
Dim aSET As Boolean
Dim RedLev As Integer
Dim RedD As Boolean
Dim RedLevL As Integer
Dim Triggered As Boolean
Dim CLIK As Long
Dim COOL As Integer

Private Sub StpAlarm_Click() ' Stop Alarm Button
    Timer2.Enabled = False
    StpAlarm.Visible = False
    StpAlarm.FontSize = 18
    Label2.BackColor = &H8000000B
    Label1.ForeColor = vbGreen
    ARMALARM_Click
    Timer4.Enabled = False
End Sub

Private Sub SetTime_Click()
If aSET Then
    ARMALARM_Click
    Exit Sub
End If
If Panel.Height = 5700 Then '5700
    Panel.Height = 2745
    ARMALARM.Enabled = True
    Label8.ForeColor = &H8000000B
    CoolTim.Enabled = False
    SetTime.Caption = "Set Alarm Time"
    SetTime.BackColor = &HC0FFC0
Else
    Panel.Height = 5700 '5700
    SetTime.Caption = ""
    SetTime.BackColor = &HC0FFFF
    CoolTim.Enabled = True
    COOL = 0
    ARMALARM.Enabled = False
    Label8.ForeColor = vbYellow
End If
End Sub

Private Sub ARMALARM_Click()  '  ARM Button
Dim E As String
If AlarmSound = "NA" Then
    MsgBox "No Alarm Sound Selected!", vbExclamation, "Can't Set"
    Exit Sub
End If
E = Format(Now, "hh:mm")
    If aTime = E And Triggered = False Then
        Beep
        Exit Sub
    End If
    If aSET = True Then
        Timer2.Enabled = False
        StpAlarm.Visible = False
        StpAlarm.FontSize = 18
        Label2.BackColor = &HE0E0E0
        Label1.ForeColor = vbGreen
        Timer4.Enabled = False
        aSET = False
        ARMALARM.BackColor = &HC0FFC0
        ARMALARM.Caption = "ARM"
        SetTime.Enabled = True
        Timer3.Enabled = False
        Timer4.Enabled = False
        Triggered = False
        Label8.ForeColor = &H8000000B
        SetTime.BackColor = &HC0FFC0
        SetTime.FontSize = 10
        SetTime.FontBold = True
        SetTime.Caption = "Set Alarm Time"
    Else
        aSET = True
        Label8.ForeColor = vbYellow
        SetTime.BackColor = vbRed
        ARMALARM.BackColor = vbRed
        ARMALARM.Caption = "ARMED"
        SetTime.Caption = "ARMED"
        SetTime.FontSize = 14
        SetTime.FontBold = False
        RedLevL = 100
        SetTime.BackColor = RGB(RedLevL, 0, 0)
        Triggered = False
        RedLev = 255
        Timer3.Enabled = True
    End If
    If Timer2.Enabled = True Then
        StpAlarm_Click
    End If
End Sub

Private Sub Command4_Click()    '   Test Button, to make sure your PC's audio is set correctly.
If AlarmSound = "NA" Then Exit Sub
    CLIK = sndPlaySound(AlarmSound, 1)
End Sub

Private Sub OpenS_Click() ' Open a New Alarm Sound (WAV)
On Error GoTo Error
    With ComD1
        .InitDir = "C:\Windows\Media"   '<---- Change to whatever you like, just a starting place.
        .Flags = &H1
        .Flags = &H2
        .DefaultExt = "wav"
        .Filter = "Wav Files (*.wav)|*.wav"
        .DialogTitle = "Open Alarm Sound"
        .Flags = &H4
        .Flags = &H1000
        .ShowOpen
    End With
        MousePointer = vbHourglass
        If ComD1.FileName = "" Then
            MousePointer = vbDefault
            Exit Sub
        End If
        AlarmSound = ComD1.FileName
        SavePrefs (App.Path & "\APref.ini")
        MousePointer = vbDefault
        Debug.Print AlarmSound
    Exit Sub
Error:
    MousePointer = vbDefault
    If Err.Number = 32755 Then Exit Sub
    MsgBox "Error loading Sound File. " & vbNewLine & _
    "ERROR #" & Err.Number & " - " & Error$(Err.Number), vbCritical, "File Load Error"
    Close #1
    Exit Sub
End Sub

Private Sub CoolTim_Timer() 'Cool animated Caption
    If COOL < 9 Then
        SetTime.Caption = Left$("Close Me", COOL)
    End If
    COOL = COOL + 1
    If COOL > 13 Then COOL = 0
Exit Sub
End Sub

Private Sub Form_Load()
    Panel.Height = 2745
    Label8.ForeColor = &H8000000B
    Label10 = Format(Now, "dddd, mm/dd/yyyy")
    aHR = "00"
    aMN = "00"
    aHRi = 0
    aMNi = 0
    aTime = "00:00"
    Label5 = aTime & ":00"
    cTime = Format(Now, "hh:mm:ss")
    Panel.Caption = Rex
    Label1 = cTime
    StpAlarm.FontSize = 18
    StpAlarm.Visible = False
    ARMALARM.Caption = "ARM"
    Timer2.Enabled = False
    Timer4.Enabled = False
    StpAlarm.Left = 1680
    StpAlarm.Top = 1150
    AlarmSound = "NA"
    OpenPrefs (App.Path & "\APref.ini")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Unload About
End Sub

Private Sub Label1_Click()
If aSET And Timer2.Enabled = True Then
    ARMALARM_Click
    Exit Sub
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        About.Show
    End If
End Sub

Private Sub Slider1_Scroll()
    aHR = Slider1
    aHRi = Slider1
    UpDateAT
End Sub

Private Sub Slider2_Scroll()
    aMN = Slider2
    aMNi = Slider2
    UpDateAT
End Sub

Private Sub Timer1_Timer()  '   Update Main Clock / Compare Alarm Time
Dim E As String
    cTime = Format(Now, "hh:mm:ss")
    Label1 = cTime
    Label10 = Format(Now, "dddd, mm/dd/yyyy")
    If aSET Then
        E = Format(Now, "hh:mm")
        If aTime = E And Triggered = False Then
            Trigger
            Triggered = True
        End If
    End If
    DetTR
End Sub

Private Sub Timer2_Timer()  '   Sound Alarm 1
    If Label2.BackColor = vbRed Then
        Label2.BackColor = &HE0E0E0
        StpAlarm.FontSize = 22
        Label1.ForeColor = vbRed
    Else
        Label2.BackColor = vbRed
        Label1.ForeColor = vbGreen
        StpAlarm.FontSize = 28
    End If
    If Check1.Value = 1 Then
        Bounce
    End If
End Sub

Public Sub UpDateAT()
Dim TotTIME As String
    If Len(aHR) < 2 Then
        aHR = "0" & aHR
    End If
    If Len(aMN) < 2 Then
        aMN = "0" & aMN
    End If
    TotTIME = aHR & ":" & aMN
    Label5.Caption = TotTIME & ":00"
    aTime = TotTIME
    DetTR
End Sub

Private Sub Timer3_Timer()  ' Fade red on buttons when armed
    If RedD Then
        RedLev = RedLev - 10
        RedLevL = RedLevL + 10
        If RedLev < 100 Then
            RedLev = 100
            RedLevL = 255
            RedD = False
        End If
    Else
        RedLev = RedLev + 10
        RedLevL = RedLevL - 10
        If RedLev > 255 Then
            RedLev = 255
            RedLevL = 100
            RedD = True
        End If
    End If
    SetTime.BackColor = RGB(RedLevL, 0, 0)
    ARMALARM.BackColor = RGB(RedLev, 0, 0)
End Sub

Private Sub Trigger()
    StpAlarm.Visible = True
    Timer2.Enabled = True
    CLIK = sndPlaySound(AlarmSound, 1)
    Timer4.Enabled = True
End Sub

Private Sub Timer4_Timer()  ' Sound Alarm 2
    CLIK = sndPlaySound(AlarmSound, 1)
End Sub

Private Sub DetTR() '   Calculate the Remaining amount of Time to Sleep
Dim qMM As Variant
Dim qHH As Variant
Dim qSS As Variant
Dim tTmp As String
Dim ctimeHH As Integer
Dim ctimeMM As Integer
Dim cTimeSS As Integer
Dim sHH As String
Dim sMM As String
Dim sSS As String
    If Timer4.Enabled = True Then   ' If the Alarm is sounding, exit.
        If Label8.ForeColor <> vbRed Then
            Label8.Caption = "00:00:00"
            Label8.ForeColor = vbRed
        End If
        Exit Sub
    End If
    qHH = Left$((Format(Now, "hh,mm")), 2)     ' Get Hour
    qMM = Right$((Format(Now, "hh,mm")), 2)    ' Get Minute
    qSS = Right$((Format(Now, "hh,mm,ss")), 2) ' Get Second
    If aHRi < qHH Then
        ctimeHH = (24 - qHH) + aHRi
    Else
        ctimeHH = aHRi - qHH
    End If
    If qMM > aMNi Then
        ctimeHH = ctimeHH - 1
        If ctimeHH = -1 Then ctimeHH = 23
        ctimeMM = (60 + aMNi) - qMM
    Else
        ctimeMM = aMNi - qMM
    End If
    
    If qSS > 0 Then
        ctimeMM = ctimeMM - 1
        If ctimeMM = -1 Then
            If aHRi <> qHH Then
                ctimeMM = 59
                ctimeHH = ctimeHH - 1
            Else
                ctimeMM = 59
                ctimeHH = 23
            End If
        End If
        cTimeSS = 60 - qSS
    Else
        cTimeSS = 0
    End If
    
    sHH = ctimeHH
    sMM = ctimeMM
    sSS = cTimeSS
    If Len(sSS) < 2 Then    '   This part just cleans it up by adding a 0 if needed.
        sSS = "0" & sSS
    End If
    If Len(sHH) < 2 Then
        sHH = "0" & sHH
    End If
    If Len(sMM) < 2 Then
        sMM = "0" & sMM
    End If
    tTmp = sHH & ":" & sMM & ":" & sSS
    Label8.Caption = tTmp
End Sub

Private Sub Bounce() ' Jump Around, Jump Around, Get up, get up and get down!
Dim qMM As Byte
    qMM = Int((4 * Rnd) + 1)
    Select Case qMM
        Case 1
            If (Me.Left - 100) > 0 Then ' Move Left
                Me.Left = Me.Left - 100
            End If
        Case 2
            If ((Me.Left + Me.Width) + 100) < Screen.Width Then ' Move Right
                Me.Left = Me.Left + 100
            End If
        Case 3
            If ((Me.Top + Me.Height) + 100) < Screen.Height Then ' Move Down
                Me.Top = Me.Top + 100
            End If
        Case 4
            If (Me.Top - 100) > 0 Then ' Move Up
                Me.Top = Me.Top - 100
            End If
    End Select
End Sub

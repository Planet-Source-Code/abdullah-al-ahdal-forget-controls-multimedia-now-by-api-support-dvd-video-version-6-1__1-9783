VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Call Update (II)"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   Icon            =   "Call.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   476
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerAtEndFile 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4980
      Top             =   3570
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   360
      TabIndex        =   14
      Top             =   720
      Width           =   3645
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3240
      Top             =   3120
   End
   Begin VB.Frame Frame1 
      Height          =   7065
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.Frame Frame4 
         Caption         =   "Misc"
         Height          =   1395
         Left            =   4800
         TabIndex        =   30
         Top             =   5580
         Width           =   2415
         Begin VB.Label LbFramesPerSecond 
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1560
            TabIndex        =   42
            Top             =   960
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Frames per second:"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   41
            Top             =   960
            Width           =   1395
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Status: "
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   40
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Current postion:"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   39
            Top             =   420
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total frames:"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   38
            Top             =   600
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total Time:"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   37
            Top             =   780
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Progress (Percent):"
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   36
            Top             =   1140
            Width           =   1350
         End
         Begin VB.Label LbStatus 
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1560
            TabIndex        =   35
            Top             =   240
            Width           =   720
         End
         Begin VB.Label LbCurrPos 
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1560
            TabIndex        =   34
            Top             =   420
            Width           =   720
         End
         Begin VB.Label LbTotalFrames 
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1560
            TabIndex        =   33
            Top             =   600
            Width           =   720
         End
         Begin VB.Label LbTotalTime 
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1560
            TabIndex        =   32
            Top             =   780
            Width           =   720
         End
         Begin VB.Label LbProgress 
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1560
            TabIndex        =   31
            Top             =   1140
            Width           =   720
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Move To"
         Height          =   1395
         Left            =   3150
         TabIndex        =   27
         Top             =   5580
         Width           =   1515
         Begin VB.CommandButton CmdMove 
            Caption         =   "MoveTo"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   810
            Width           =   1275
         End
         Begin VB.TextBox TxtMove 
            Height          =   285
            Left            =   120
            TabIndex        =   28
            Top             =   390
            Width           =   1275
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Resize"
         Height          =   1365
         Left            =   210
         TabIndex        =   17
         Top             =   5580
         Width           =   2835
         Begin VB.CommandButton CmdResize 
            Caption         =   "Resize "
            Height          =   735
            Left            =   2040
            TabIndex        =   22
            Top             =   330
            Width           =   735
         End
         Begin VB.TextBox txtLeft 
            Height          =   315
            Left            =   660
            TabIndex        =   21
            Top             =   330
            Width           =   375
         End
         Begin VB.TextBox TxtTop 
            Height          =   315
            Left            =   1620
            TabIndex        =   20
            Top             =   330
            Width           =   375
         End
         Begin VB.TextBox txtWidth 
            Height          =   315
            Left            =   660
            TabIndex        =   19
            Top             =   750
            Width           =   375
         End
         Begin VB.TextBox TxtHeight 
            Height          =   315
            Left            =   1620
            TabIndex        =   18
            Top             =   750
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Left:"
            ForeColor       =   &H00000040&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   390
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Top:"
            ForeColor       =   &H00000040&
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   25
            Top             =   390
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Width:"
            ForeColor       =   &H00000040&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   810
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Height:"
            ForeColor       =   &H00000040&
            Height          =   195
            Index           =   3
            Left            =   1080
            TabIndex        =   23
            Top             =   810
            Width           =   510
         End
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   4020
         Pattern         =   $"Call.frx":0442
         System          =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   3285
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   300
         TabIndex        =   15
         Top             =   270
         Width           =   6975
      End
      Begin VB.TextBox txtFrom 
         Height          =   315
         Left            =   1860
         TabIndex        =   11
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox TxtTo 
         Height          =   315
         Left            =   2880
         TabIndex        =   10
         Top             =   3240
         Width           =   615
      End
      Begin VB.Frame FrameVideo 
         Caption         =   "Video View"
         Height          =   2175
         Left            =   3900
         TabIndex        =   8
         Top             =   2820
         Width           =   3315
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Auto Repeat"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton CmdOpen 
         Caption         =   "Open"
         Height          =   315
         Left            =   300
         TabIndex        =   6
         Top             =   2880
         Width           =   3495
      End
      Begin VB.CommandButton CmdPlay 
         Caption         =   "Play"
         Height          =   315
         Left            =   300
         TabIndex        =   5
         Top             =   3240
         Width           =   915
      End
      Begin VB.CommandButton CmdPause 
         Caption         =   "Pause"
         Height          =   315
         Left            =   300
         TabIndex        =   4
         Top             =   3600
         Width           =   3495
      End
      Begin VB.CommandButton CmdResume 
         Caption         =   "Resume"
         Height          =   315
         Left            =   300
         TabIndex        =   3
         Top             =   3960
         Width           =   3495
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "Stop"
         Height          =   315
         Left            =   300
         TabIndex        =   2
         Top             =   4320
         Width           =   3495
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
         Height          =   315
         Left            =   300
         TabIndex        =   1
         Top             =   4680
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   5
         Left            =   1380
         TabIndex        =   13
         Top             =   3300
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "To:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   4
         Left            =   2580
         TabIndex        =   12
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label LbResult 
         Caption         =   "Result calling Function is : "
         ForeColor       =   &H00C00000&
         Height          =   555
         Left            =   1680
         TabIndex        =   9
         Top             =   5040
         Width           =   5535
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OpenMPEG Lib "Engine.dll" (ByVal hwnsd As Long, ByVal lpFileName As String, ByVal TypeMpegOrAVI As String) As String
Private Declare Function PlayMPEG Lib "Engine.dll" (ByVal play_from_where As Long, ByVal play_to_where As Long) As String
Private Declare Function GetPercent Lib "Engine.dll" () As Long
Private Declare Function GetTotalframes Lib "Engine.dll" () As Long
Private Declare Function GetTotalTimeByms Lib "Engine.dll" () As Long
Private Declare Function GetCurrentMPEGPos Lib "Engine.dll" () As Long
Private Declare Function GetStatusMPEG Lib "Engine.dll" () As Long
Private Declare Sub PauseMPEG Lib "Engine.dll" ()     'As String
Private Declare Sub ResumeMPEG Lib "Engine.dll" () 'As String
Private Declare Sub StopMPEG Lib "Engine.dll" () 'As String
Private Declare Sub CloseMPEG Lib "Engine.dll" () ' As String
Private Declare Function MoveMPEG Lib "Engine.dll" (ByVal Seek_to As Long) As String
Private Declare Function PutMPEG Lib "Engine.dll" (ByVal left As Long, ByVal top As Long, ByVal Width As Long, ByVal Height As Long) As String
Private Declare Sub SetAutoRepeat Lib "Engine.dll" (ByVal AutoRepeatOrNo As Long)

Private Declare Function GetDefaultDevice Lib "Engine.dll" (ByVal typeDevice As String) As String
Private Declare Sub SetDefaultDevice Lib "Engine.dll" (ByVal typeDevice As String, ByVal drvDefaultDevice As String)

Private Declare Function GetFramesPerSecond Lib "Engine.dll" () As Long
Private Declare Function AreMPEGAtEnd Lib "Engine.dll" () As Boolean


Const PLAYING = 3
Const PAUSED = 2
Const STOPPED = 1
Const NOT_OPENED = -1
Private Sub Check1_Click()
If Check1.Value = 1 Then
SetAutoRepeat Val(1) 'Set Auto repeat true
Else
SetAutoRepeat Val(-1) 'Set Auto repeat false
End If

End Sub

Private Sub CmdClose_Click()
CloseMPEG
End Sub

Private Sub CmdMove_Click()
Dim Result As String
Result = MoveMPEG(Val(TxtMove))
LbResult = "Result calling Function is : " & Result
End Sub

Private Sub CmdOpen_Click()
Dim FileName As String
FileName = GetFile
Dim typeDevice As String
Dim Result As String

If FileName = "error" Then Exit Sub

If LCase(Right(File, 4)) = ".avi" Then 'if the movie is avi then select type
    typeDevice = "AviVideo"
'ElseIf Right(File, 4) = ".rmi" Or Right(File, 4) = ".mid" Then
'typeDevice = "sequencer" ' select this type for midi and rmi files

Else 'else this mean it mpg,mp3,mp2,mp1,wav,,,etc then we will choose "MpegVideo" type
    typeDevice = "MPEGVideo"
End If

Result = OpenMPEG(FrameVideo.hWnd, FileName, typeDevice) 'call now function openMPEG
LbResult = "Result calling Function is : " & Result
If LbResult = "Result calling Function is : Success" Then 'this mean openMPEG success
LbTotalFrames = GetTotalframes 'Get total frames
LbTotalTime = GetTotalTimeByms / 1000 'Get Total Time

'this Function Will return amount frames per second if it
'Success or if not will return value -1
LbFramesPerSecond = GetFramesPerSecond


Timer1.Enabled = True 'Enable timer1 goto Sub Timer1 to See Functions
End If


End Sub

Private Sub CmdPause_Click()
PauseMPEG
End Sub

Private Sub CmdPlay_Click()
CmdResize_Click
Dim Result As String
Result = PlayMPEG(Val(txtFrom), Val(TxtTo))
LbResult = "Result calling Function is : " & Result

If LbResult = "Result calling Function is : Success" Then 'this mean the Function Play Success
'then do the folloing
TimerAtEndFile.Enabled = True 'Enable this to know if the Multimedia
'File at the end now you must Enable the Timer
'If the Function PlayMPEG Success not else.
'Go to Sub TimerAtEnd File And Read it carefully
End If

End Sub

Private Sub CmdResize_Click()
Dim Result As String
Result = PutMPEG(Val(txtLeft), Val(TxtTop), Val(txtWidth), Val(TxtHeight))
LbResult = "Result calling Function is : " & Result

End Sub

Private Sub CmdResume_Click()
ResumeMPEG
End Sub

Private Sub CmdStop_Click()
StopMPEG
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dim StoreDirive As Variant
StoreDirive = Dir1.Path
On Error Resume Next
Dir1.Path = Drive1.Drive
If Err = 68 Then Drive1.Drive = StoreDirive

End Sub
Public Function GetFile() As String
Dim File As String
If File1.FileName = "" Then GetFile = "error": Exit Function
        'CHECK IF THE FILE IN ROOT DIR
        If Len(File1.Path) > 3 Then
            File = File1.Path & "\" & File1.FileName
        Else
            File = File1.Path & File1.FileName
        End If
        
        
        
GetFile = File

End Function

Private Sub Form_Load()
MsgBox "Hello, Select your mutimedia file and then click open and after that click play to play it", vbInformation

'this Function help you if you want to the default device
'the parameter must be the device type like:
'MPEGVideo
'sequencer
'avivideo
'waveaudio
'videodisc
If Not GetDefaultDevice("MPEGVideo") = "mciqtz.drv" Then
'if Driver"mciqtz.drv" not the default device for type
'"MpegVideo" then set mciqtz.drv as a default device


SetDefaultDevice "MPEGVideo", "mciqtz.drv"
'SetDefaultDevice "MPEGVideo", "mciqtz.drv"' this the most
'improtant device and it will receives calls mci
'Some programs change this device like xing mpeg
'and if this occur you can not play all mutimedia files
'and you will unexpected errors
End If

If Not GetDefaultDevice("sequencer") = "mciseq.drv" Then
'if Driver"mciseq.drv" not the default device for type
'"sequencer" then set mciqtz.drv as a default device
SetDefaultDevice "sequencer", "mciseq.drv"
End If

If Not GetDefaultDevice("avivideo") = "mciavi.drv" Then
'if Driver"mciavi.drv" not the default device for type
'"avivideo" then set avivideo as a default device
SetDefaultDevice "avivideo", "mciavi.drv"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim resultMsg As Long
resultMsg = MsgBox("Are you wanna vote me ? click yes to open vote.", vbInformation Or vbYesNo)
If resultMsg = vbYes Then Shell "start http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=9783"
End Sub

Private Sub Timer1_Timer()
Status = GetStatusMPEG
If Status = PLAYING Then LbStatus = "Playing"
If Status = STOPPED Then LbStatus = "Stopped"
If Status = PAUSED Then LbStatus = "Paused"
If Status = NOT_OPENED Then LbStatus = "no file"
LbCurrPos = GetCurrentMPEGPos
LbProgress = GetPercent & " %"
End Sub

Private Sub TimerAtEndFile_Timer()
'The Following Function will tell if multimedia file now at end
'to use this Function put it in a timer and set Interval
'for a timer = 100 and make the timer false and after Play
'Multimedia files Successfully set the timer true.


If AreMPEGAtEnd = 1 Then
''this mean  file multimedia at the end now then
''write your commnad here or call you favourit Fucntion
''or even you can play the file again or paly the next file
''if you had a list of multimedia files.
'.....
'...
'..
'if you wanna know if the multimedia file
'at the end now don't use option Auto Repeat
'you must do auto repeat by yourself by the following command
'in this place after make the previous compare (I mean here)

'Result = PlayMPEG(txtFrom, TxtTo)
'or you have choice to close this File and open
'another file and play it( this if had a list of files)
'like this command after make the previous compare(I mean here)
'Dim Result As String
'Result = CloseMPEG
'Result = OpenMPEG(FrameVideo.hwnd, filename, typeDevice) 'call now function openMPEG
'Result = PlayMPEG(txtFrom, TxtTo)

TimerAtEndFile.Enabled = False ' and remeber don't forget
'write this line becuase you got what you want then
'Close the timer.Okay.

'MsgBox "at the end now"
Else
'this mean result calling function false and this mean the
'multimedia file not at the end now
'....
'...
'..

End If

End Sub

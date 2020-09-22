VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Controlling with Multimedia BY pure API made for Planet Source Code  (II)"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   Icon            =   "MPEG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerAtEndFile 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4320
      Top             =   4290
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3480
      Top             =   3120
   End
   Begin VB.Frame Frame1 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
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
         Left            =   3930
         Pattern         =   $"MPEG.frx":0442
         System          =   -1  'True
         TabIndex        =   40
         Top             =   660
         Width           =   3285
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   270
         TabIndex        =   39
         Top             =   660
         Width           =   3615
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   6975
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
         Height          =   315
         Left            =   300
         TabIndex        =   34
         Top             =   4680
         Width           =   3495
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "Stop"
         Height          =   315
         Left            =   300
         TabIndex        =   33
         Top             =   4320
         Width           =   3495
      End
      Begin VB.CommandButton CmdResume 
         Caption         =   "Resume"
         Height          =   315
         Left            =   300
         TabIndex        =   32
         Top             =   3960
         Width           =   3495
      End
      Begin VB.CommandButton CmdPause 
         Caption         =   "Pause"
         Height          =   315
         Left            =   300
         TabIndex        =   31
         Top             =   3600
         Width           =   3495
      End
      Begin VB.CommandButton CmdPlay 
         Caption         =   "Play"
         Height          =   315
         Left            =   300
         TabIndex        =   30
         Top             =   3240
         Width           =   915
      End
      Begin VB.CommandButton CmdOpen 
         Caption         =   "Open"
         Height          =   315
         Left            =   300
         TabIndex        =   29
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Frame Frame2 
         Caption         =   "Resize"
         Height          =   1365
         Left            =   300
         TabIndex        =   19
         Top             =   5520
         Width           =   2835
         Begin VB.CommandButton CmdResize 
            Caption         =   "Resize "
            Height          =   735
            Left            =   2040
            TabIndex        =   24
            Top             =   330
            Width           =   735
         End
         Begin VB.TextBox txtLeft 
            Height          =   315
            Left            =   660
            TabIndex        =   23
            Top             =   330
            Width           =   375
         End
         Begin VB.TextBox TxtTop 
            Height          =   315
            Left            =   1620
            TabIndex        =   22
            Top             =   330
            Width           =   375
         End
         Begin VB.TextBox txtWidth 
            Height          =   315
            Left            =   660
            TabIndex        =   21
            Top             =   750
            Width           =   375
         End
         Begin VB.TextBox TxtHeight 
            Height          =   315
            Left            =   1620
            TabIndex        =   20
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
            TabIndex        =   28
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
            TabIndex        =   27
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
            TabIndex        =   26
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
            TabIndex        =   25
            Top             =   810
            Width           =   510
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Move To"
         Height          =   1395
         Left            =   3240
         TabIndex        =   16
         Top             =   5520
         Width           =   1515
         Begin VB.CommandButton CmdMove 
            Caption         =   "MoveTo"
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   810
            Width           =   1275
         End
         Begin VB.TextBox TxtMove 
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   390
            Width           =   1275
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Misc"
         Height          =   1395
         Left            =   4860
         TabIndex        =   5
         Top             =   5520
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            TabIndex        =   11
            Top             =   1140
            Width           =   1350
         End
         Begin VB.Label LbStatus 
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1560
            TabIndex        =   10
            Top             =   240
            Width           =   720
         End
         Begin VB.Label LbCurrPos 
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1560
            TabIndex        =   9
            Top             =   420
            Width           =   720
         End
         Begin VB.Label LbTotalFrames 
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1560
            TabIndex        =   8
            Top             =   600
            Width           =   720
         End
         Begin VB.Label LbTotalTime 
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1560
            TabIndex        =   7
            Top             =   780
            Width           =   720
         End
         Begin VB.Label LbProgress 
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1560
            TabIndex        =   6
            Top             =   1140
            Width           =   720
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Auto Repeat"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Frame FrameVideo 
         Caption         =   "Video View"
         Height          =   2175
         Left            =   3900
         TabIndex        =   3
         Top             =   2820
         Width           =   3315
      End
      Begin VB.TextBox TxtTo 
         Height          =   315
         Left            =   2880
         TabIndex        =   2
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtFrom 
         Height          =   315
         Left            =   1860
         TabIndex        =   1
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label LbResult 
         Caption         =   "Result calling Function is : "
         ForeColor       =   &H00C00000&
         Height          =   555
         Left            =   1680
         TabIndex        =   37
         Top             =   5040
         Width           =   5535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "To:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   4
         Left            =   2580
         TabIndex        =   36
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   5
         Left            =   1380
         TabIndex        =   35
         Top             =   3300
         Width           =   390
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 1 Then 'if checked then Set Move or Audio Auto Repeat
SetAutoRepeat True 'Set Auto repeat true
Else 'else don't set it auto repeat
SetAutoRepeat False 'Set Auto repeat false
End If
'returned Value for this function = without because this it sub not Function :).
End Sub

Private Sub CmdClose_Click()
'calling CloseMPEG will close the multimedia file

'Note any function return a string value if the function "Success"
'Will Return Value is "Success" or not will return a string
'for the error for e.g
'Dim Result As String
'Result = CloseMPEG()

'If Result = "Success" Then
' this mean success write your commands below
''....
'Else
'MsgBox Result 'show in message box the error
'End If

Dim Result As String
Result = CloseMPEG()
LbResult = "Result calling Function is : " & Result
Timer1.Enabled = False 'Disable Timer1
End Sub

Private Sub CmdMove_Click()
Dim Result As String
'calling MoveMPEG will seek (change the position)for
'the multimedia file
'it need just one parameter for to where want to change
'position (must be number of frame you want to go to)


'Note any function return a string value if the function "Success"
'Will Return Value is "Success" or not will return a string
'for the error for e.g
'Dim Result As String
'Result = MoveMPEG(1000)

'If Result = "Success" Then
' this mean success write your commands below
''....
'Else
'MsgBox Result 'show in message box the error
'End If

Result = MoveMPEG(Val(TxtMove))
LbResult = "Result calling Function is : " & Result
End Sub

Private Sub CmdOpen_Click()
On Error Resume Next
Dim filename As String
Dim TypeDevice As String
filename = GetFile 'call function to get the file and check if it gave any space
If filename = "error" Then Exit Sub

Dim Result As String

'Callig OpenMPEG will open the multimedia file
'Parameters
'hWnd
'[in]handle of the window
'which you want to play in. you can put handle for
'your desktop if you want to playing movie in your desktop.

'filename
'[in]Specifies file name and the path it can contain any space
'which you want to play.

'typeDevice
'[in] Specifies a type of MCI device and it could be from the following:
'Type MCI       description                     driver file
'sequencer      dealing with mid                mciseq.drv
'               files
'MPEGVideo      dealing with most multimedia    mciqtz.drv
'               like mpg,mp3,mp2..
'               au,aiff,..etc also support
'               avi,vob(for DVD),midi,mid
'               and rmi files.because of this
'               my advice to you to use
'               type "MPEGVideo" to playing
'               MOST FILES even avi!!
'               I got this info from my
'               experiment when I opened
'               System.ini in section MCI
'               Then I must share others.
'avivideo       deling with avi movie           mciavi.drv

'the following types if you had ATI RAGE II or Later
'(This VGA Card to Support DVD Video)

'DvdVideo       This support DVD's Video        MciCinem.drv DVD
'ATIMPEGVIDEO   to playing MPEG Video           mciatim1.drv

'But my advice to you to not use type "ATIMPEGVIDEO" & "DvdVideo" because
'Type MPEGVideo can support most Multimedia files and also support DVD's
'Video if you had ATI RAGE II or LATER.
'last note for DVD Video: you must have a fast computer

'note : Type "MpegVideo" support these extensions:
'qt , mov, dat,snd, mpg, mpa, mpv, enc, m1v, mp2,mp3, mpe, mpeg, mpm
'au , snd, aif, aiff, aifc,wav,wmv,wma,avi,midi,mid,rmi,avi,etc.

'Note if there are any new type in (system.ini in windows 98 or in registry in windows 2000)
'it will supported by Type "MPEGVideo" because of this use type "MPEGVideo" to playing
'Most Files and remember you can use sequencer for mid and avivideo for avi,,etc.

'Now you must note using Type "MPEGVideo" can playing all Multimedia files

'Note : if this Function success will return value string "Success"
'or if not will return value string description the error which occur

'Okay make sure if you used this function don't forget to use function
'CloseMPEG When you will end your program or you
'will got error message

'for the error for e.g
'Dim Result As String
'Result = OpenMPEG(FrameVideo.hwnd, "c:\mymove.dat", "MPEGVIDEO")

'If Result = "Success" Then
' this mean success write your commands below
''....
'Else
'MsgBox Result 'show in message box the error
'End If

If LCase(Right(File, 4)) = ".avi" Then 'if the movie is avi then select type
    TypeDevice = "AviVideo"
'ElseIf Right(File, 4) = ".rmi" Or Right(File, 4) = ".mid" Then
'TypeDevice = "sequencer" ' select this type for midi and rmi files
Else 'else this mean it mpg,mp3,mp2,mp1,wav,,,etc then we will choose "MpegVideo" type
    TypeDevice = "MPEGVideo"
End If

Result = OpenMPEG(FrameVideo.hwnd, filename, TypeDevice) 'call now function openMPEG
LbResult = "Result calling Function is : " & Result

If Result = "Success" Then 'this mean openMPEG success
'this Function Will return amount frames per second if it
'Success or if not will return value -1
LbFramesPerSecond = GetFramesPerSecond

LbTotalFrames = GetTotalframes 'Get total frames
LbTotalTime = GetTotalTimeByMS / 1000 'Get Total Time



Timer1.Enabled = True 'Enable timer1 goto Sub Timer1 to See Functions
End If

'If GetStatusMPEG = "stopped" Then MsgBox "stopped"
End Sub

Private Sub CmdPause_Click()
'calling PauseMPEG will pause the multimedia file

'Note any function return a string value if the function "Success"
'Will Return Value is "Success" or not will return a string
'for the error for e.g
'Dim Result As String
'Result = PauseMPEG()

'If Result = "Success" Then
' this mean success write your commands below
''....
'Else
'MsgBox Result 'show in message box the error
'End If

Dim Result As String
Result = PauseMPEG()
LbResult = "Result calling Function is : " & Result
End Sub

Private Sub CmdPlay_Click()
CmdResize_Click
Dim Result As String
'calling PlayMPEG will playing the multimedia file
'the first parameter for from where playing file
'the second parameter for to where playing file
'if the first parameter is vbNullString and the second parameter is vbNullString the Function Will:
'playing from the beginning to end.
'if the first parameter is 10 and the second parameter is 100 the Function Will:
'playing from 10 to 100 and stop.
'if the first parameter is vbNullString and the second parameter is 100 the Function Will:
'playing from the beginning to 100 and stop.
'if the first parameter is 104 and the second parameter is vbNullString the Function Will:
'playing from 104 to end.
'Note :the numbers 10,100,104 is an example for from playing to where end playing


'Note any function return a string value if the function "Success"
'Will Return Value is "Success" or not will return a string
'for the error for e.g
'Dim Result As String
'Result = PlayMPEG(1, 1000)

'If Result = "Success" Then
' this mean success write your commands below
''....
'Else
'MsgBox Result 'show in message box the error
'End If

Result = PlayMPEG(txtFrom, TxtTo)

If Result = "Success" Then 'this mean the Function Play Success
'then do the folloing
TimerAtEndFile.Enabled = True 'Enable this to know if the Multimedia
'File at the end now you must Enable the Timer
'If the Function PlayMPEG Success not else.
'Go to Sub TimerAtEnd File And Read it carefully
End If

LbResult = "Result calling Function is : " & Result
End Sub

Private Sub CmdResize_Click()
Dim Result As String
'Calling PutMPEG will resize the move and :
'if you are set parameter width or Height zero
'the function will get the actual size of the window which
'want to play in and resize the movie to fit the window

'Note any function return a string value if the function "Success"
'Will Return Value is "Success" or not will return a string
'for the error for e.g
'Dim Result As String
'Result = PutMPEG(0, 0,0, 0)' set the normal size for the move as size window

'If Result = "Success" Then
' this mean success write your commands below
''....
'Else
'MsgBox Result 'show in message box the error
'End If

Result = PutMPEG(Val(txtLeft), Val(TxtTop), Val(txtWidth), Val(TxtHeight))
LbResult = "Result calling Function is : " & Result
End Sub

Private Sub CmdResume_Click()
'calling ResumeMPEG will Resume the multimedia file

'Note any function return a string value if the function "Success"
'Will Return Value is "Success" or not will return a string
'for the error for e.g
'Dim Result As String
'Result = ResumeMPEG()

'If Result = "Success" Then
' this mean success write your commands below
''....
'Else
'MsgBox Result 'show in message box the error
'End If

Dim Result As String
Result = ResumeMPEG()
LbResult = "Result calling Function is : " & Result
End Sub

Private Sub CmdStop_Click()
'calling StopMPEG will Stop the multimedia file

'Note any function return a string value if the function "Success"
'Will Return Value is "Success" or not will return a string
'for the error for e.g
'Dim Result As String
'Result = StopMPEG()

'If Result = "Success" Then
' this mean success write your commands below
''....
'Else
'MsgBox Result 'show in message box the error
'End If

Dim Result As String
Result = StopMPEG()
LbResult = "Result calling Function is : " & Result
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
If File1.filename = "" Then MsgBox "Please select a file": GetFile = "error": Exit Function
        'CHECK IF THE FILE IN ROOT DIR
        If Len(File1.Path) > 3 Then
            File = File1.Path & "\" & File1.filename
        Else
            File = File1.Path & File1.filename
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CloseMPEG ' when we will end the program we must close or we will got error message
End Sub


Private Sub Timer1_Timer()
'Calling Function GetPercent
'the returned value from this function is Percent "Progress"
'and if the function failed will return value -1
LbProgress = GetPercent & " %"

'Calling Function GetCurrentMPEGPos
'the returned value from this function is number of current frame
'and if the function failed will return value -1
LbCurrPos = GetCurrentMPEGPos

'Calling Function GetStatusMPEG
'the returned value from this function is status of file "audio or video" if it "playing" or "stopped" or "paused"
'and if the function failed will return value "ERROR"

'also you can exame the status like this: you can copy it
'Dim Result As String
'Result = GetStatusMPEG
'If Result = "ERROR" Then 'this mean failed write you commands here
''.....
''....
''..
'ElseIf Result = "playing" Then 'this mean it now playing .ok write you commands here
''....
''...
''..
'ElseIf Result = "stopped" Then 'this mean it now stopped .ok write you commands here
''....
''...
''..
'ElseIf Result = "paused" Then 'this mean it now paused .ok write you commands here
''....
''...
''..

'End If
LbStatus = GetStatusMPEG


'Improtant Note:
'Don't Put this Function in any Timers or the program will
'be very slow
'1-GetTotalframes
'2-GetTotalTimeByMS
'3-GetFramesPerSecond
End Sub

'Okay I hope you Enjoyed
'You can use this module in your own projects if you wanna
'the easist deal with multimedia.
'Using API is more stronger than using controls and not take a space
'for any request, suggestions,Devlopment or bugs   e-mail at
'a_ahdl@yahoo.com
'Thank you

Private Sub TimerAtEndFile_Timer()
'The Following Function will tell if multimedia file now at end
'to use this Function put it in a timer and set Interval
'for a timer = 100 and make the timer false and after Play
'Multimedia files Successfully set the timer true.


If AreMPEGAtEnd = True Then
''this mean  file multimedia at the end now then
''write your commnad here or call you favourit Fucntion
''or even you can play the file again or play the next file
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

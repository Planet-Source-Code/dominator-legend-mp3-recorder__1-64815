VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Frm_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3 Recorder"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   FillColor       =   &H00404040&
   ForeColor       =   &H00808080&
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fram_Interface 
      Caption         =   " Capture Options "
      Height          =   3795
      Index           =   1
      Left            =   150
      TabIndex        =   12
      Top             =   4080
      Width           =   5685
      Begin ComctlLib.ListView Lst_Data 
         Height          =   1275
         Left            =   90
         TabIndex        =   24
         Top             =   1350
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   2249
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483636
         BackColor       =   -2147483633
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Description"
            Object.Width           =   5999
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Value"
            Object.Width           =   1942
         EndProperty
      End
      Begin VB.CheckBox Chk_Show 
         Caption         =   "Draw captured wav data "" This may cause voice cutting """
         Height          =   225
         Left            =   180
         TabIndex        =   22
         Top             =   2850
         Value           =   1  'Checked
         Width           =   5295
      End
      Begin VB.PictureBox Pic_Interface 
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   1
         Left            =   180
         ScaleHeight     =   435
         ScaleWidth      =   5295
         TabIndex        =   19
         Top             =   3210
         Width           =   5295
         Begin VB.CommandButton Cmd_Intervals 
            Caption         =   "Set Recording Interval"
            Height          =   420
            Left            =   1515
            TabIndex        =   26
            Top             =   0
            Width           =   2055
         End
         Begin VB.CommandButton Cmd_Help 
            Caption         =   "&Help"
            Height          =   420
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   1365
         End
         Begin VB.CommandButton Cmd_Execute 
            Caption         =   "Start Capturing"
            Height          =   420
            Left            =   3720
            TabIndex        =   21
            Top             =   0
            Width           =   1575
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   0
            ScaleHeight     =   225
            ScaleWidth      =   1875
            TabIndex        =   20
            Top             =   2040
            Width           =   1875
         End
      End
      Begin VB.ComboBox Cmb_Channels 
         Height          =   315
         ItemData        =   "Frm_Main.frx":6852
         Left            =   3750
         List            =   "Frm_Main.frx":685C
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   540
         Width           =   1755
      End
      Begin VB.ComboBox Cmb_SampleRate 
         Height          =   315
         ItemData        =   "Frm_Main.frx":686E
         Left            =   180
         List            =   "Frm_Main.frx":687B
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   540
         Width           =   1665
      End
      Begin VB.Timer Recorder 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5220
         Top             =   120
      End
      Begin VB.ComboBox Cmb_Bitrate 
         Height          =   315
         ItemData        =   "Frm_Main.frx":6894
         Left            =   1965
         List            =   "Frm_Main.frx":68C2
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   540
         Width           =   1665
      End
      Begin VB.Line Lin_Break 
         BorderColor     =   &H00C0C0C0&
         Index           =   1
         X1              =   180
         X2              =   5490
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Line Lin_Break 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   180
         X2              =   5490
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Label Lbl_Interface 
         AutoSize        =   -1  'True
         Caption         =   "Abstract Recording Informations:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   25
         Top             =   990
         Width           =   2400
      End
      Begin VB.Label Lbl_Interface 
         AutoSize        =   -1  'True
         Caption         =   "Sample Rate:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   18
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Lbl_Interface 
         AutoSize        =   -1  'True
         Caption         =   "Channels:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   195
         Index           =   4
         Left            =   3750
         TabIndex        =   17
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Lbl_Interface 
         AutoSize        =   -1  'True
         Caption         =   "Bit Rate:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   195
         Index           =   6
         Left            =   1965
         TabIndex        =   16
         Top             =   330
         Width           =   630
      End
   End
   Begin VB.Frame Fram_Interface 
      Caption         =   " Captured Stream "
      Height          =   2535
      Index           =   0
      Left            =   150
      TabIndex        =   5
      Top             =   1380
      Width           =   5685
      Begin VB.PictureBox Pic_WavDisplay 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         DrawStyle       =   1  'Dash
         FillColor       =   &H00008000&
         FillStyle       =   4  'Upward Diagonal
         ForeColor       =   &H00008000&
         Height          =   1875
         Left            =   180
         Picture         =   "Frm_Main.frx":6905
         ScaleHeight     =   1815
         ScaleWidth      =   5265
         TabIndex        =   6
         Top             =   510
         Width           =   5325
         Begin VB.PictureBox RightPic 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            FillColor       =   &H00008000&
            ForeColor       =   &H00008000&
            Height          =   1650
            Left            =   5115
            ScaleHeight     =   1650
            ScaleWidth      =   150
            TabIndex        =   10
            Top             =   0
            Width           =   150
         End
         Begin VB.PictureBox LeftPic 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            FillColor       =   &H00008000&
            Height          =   1650
            Left            =   0
            ScaleHeight     =   1650
            ScaleWidth      =   150
            TabIndex        =   9
            Top             =   0
            Width           =   150
         End
         Begin VB.PictureBox Pic_Volume 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            FillColor       =   &H00008000&
            ForeColor       =   &H00004000&
            Height          =   135
            Left            =   0
            ScaleHeight     =   135
            ScaleWidth      =   5325
            TabIndex        =   7
            Top             =   1670
            Width           =   5325
            Begin VB.Label Lbl_VolPercent 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "100%"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000C000&
               Height          =   180
               Left            =   2490
               TabIndex        =   8
               Top             =   -30
               Width           =   450
            End
         End
      End
      Begin VB.Label Lbl_Interface 
         AutoSize        =   -1  'True
         Caption         =   "Visualization of captured WAV stream."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   11
         Top             =   300
         Width           =   2745
      End
   End
   Begin VB.PictureBox Pic_Interface 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   0
      Picture         =   "Frm_Main.frx":531CB
      ScaleHeight     =   315
      ScaleWidth      =   5985
      TabIndex        =   3
      Top             =   8055
      Width           =   5985
      Begin VB.Label Lbl_Interface 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "© 3.2006 Dominator Legend."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00592D00&
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   4
         Top             =   45
         Width           =   2115
      End
   End
   Begin VB.Label Lbl_Interface 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dominator Legend MP3 - Recorder"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D09764&
      Height          =   510
      Index           =   0
      Left            =   165
      TabIndex        =   1
      Top             =   240
      Width           =   5640
   End
   Begin VB.Label Lbl_Interface 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "© 2006 Dominator Legend, MP3 - Recorder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D09764&
      Height          =   195
      Index           =   2
      Left            =   165
      TabIndex        =   0
      Top             =   720
      Width           =   3120
   End
   Begin VB.Label Lbl_Interface 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dominator Legend MP3 - Recorder"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006B2401&
      Height          =   510
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   255
      Width           =   5640
   End
   Begin VB.Image Img_Interface 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2400
      Left            =   -3360
      Picture         =   "Frm_Main.frx":AFB4D
      Top             =   -1170
      Width           =   12030
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Rem -> Create refrence from BladeEncoder class which will convert Wav chunk into MP3
Public WithEvents BE                As CEncoder
Attribute BE.VB_VarHelpID = -1
Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Rem -> Create refrence from Frm_DXRecorder which will capture the sound for us
Public WithEvents DX                As CDirectX
Attribute DX.VB_VarHelpID = -1
Dim DataElements(6) As String
Private Sub Form_Initialize()
    If App.PrevInstance Then Beep: End
End Sub
Private Sub Form_Load()
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Reset interface components values
    DisplayType = 3                     'The type into which WAV is Displayed
    DigialVolume = 1                    'Make the volume is the loudest
    Cmb_Bitrate.ListIndex = 8           'Reset BitRate to 128
    Cmb_Channels.ListIndex = 1          'Reset Channel to Stereo
    Cmb_SampleRate.ListIndex = 1        'Reset Sample Rate to 44100
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Reset the volum bar with the half of its size - 100% from 200%
    Pic_Volume.Scale (0, 0)-(Val(Lbl_VolPercent.Caption) * 2, 1)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Start drawing the volum bar, with the value determined above
    Call DrawVolume(Pic_Volume)
    Lst_Data.ListItems.Add(, , "Recording Media Duration").SubItems(1) = "00:00:00"
    Lst_Data.ListItems.Add(, , "Recording File Size").SubItems(1) = "0.0 KByte"
    Lst_Data.ListItems.Add(, , "Recording Packets Size").SubItems(1) = "0.0 KByte"
    Lst_Data.ListItems.Add(, , "Recording Packets Count").SubItems(1) = "0"
    Lst_Data.ListItems.Add(, , "Recording Pre-Define Duration").SubItems(1) = "00:00:00"
    Lst_Data.ListItems.Add(, , "Recording File Path").SubItems(1) = "Not specified."
End Sub
Rem -> ********************************************************************************
Rem -> We have to unload every thing and must deinitialize the BladeEncoder, so we can
Rem -> Use it next time, if it's don't deinitialize we got error when we initalize
Rem -> It again.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Rem -> ~~~~~~~~~~~~~~~~~~~
    Rem -> Destroy every thing
    Call UnloadAll
    End
End Sub
Rem -> *******************************************************************************************************************************************************************************
Rem -> Events required to change the value of the volume bar
Private Sub Pic_Volume_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pic_Volume_MouseMove Button, Shift, X, Y
End Sub
Rem -> *******************************************************************************************************************************************************************************
Rem -> Events required to change the value of the volume bar
Private Sub Pic_Volume_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        DigialVolume = X / 100#
        If DigialVolume < 0 Then DigialVolume = 0
        If X > Pic_Volume.ScaleWidth Then DigialVolume = Pic_Volume.ScaleWidth / 100#
        Call DrawVolume(Pic_Volume)
    End If
End Sub
Rem -> *******************************************************************************************************************************************************************************
Rem -> Events required to change display type of wav data, try to click on the big
Rem -> Picture box
Public Sub Pic_WavDisplay_Click()
    DisplayType = DisplayType + 1
    If DisplayType > 3 Then DisplayType = 0
End Sub
Private Sub Cmd_Execute_Click()
    Select Case (Cmd_Execute.Caption)
        Case ("Start Capturing"):
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> Initialize The DirectX and Blade Encoder library
            Call Initializer(IIf(Cmb_Channels.ListIndex = 0, 1, 2), Val(Cmb_SampleRate.List(Cmb_SampleRate.ListIndex)), Val(Cmb_Bitrate.List(Cmb_Bitrate.ListIndex)))
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> The user request from us to save captured data to file
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> Declaring some variables used to create file
            Dim DirName                      As String
            Dim FileTitle                    As String
            Dim FilePath                     As String
            Dim FileID                       As Integer
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> The file will be created in application path
            DirName = App.Path & "\"
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> This loop is used to generate a file name, (More easyer)
            Do Until Dir(FileTitle, vbArchive + vbHidden) = ""
                FileID = FileID + 1
                FileTitle = "GID-" & Right("0000" & FileID, 4) & ".Mp3"
            Loop
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> Now compose the file path wich is DirName + FileTitle
            FilePath = DirName & FileTitle
            DataElements(5) = FileTitle
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> Returns the next file number available for use by Open statment
            Rem -> This number is used instead of its name, just for making it easy.
            FileNum = FreeFile
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> Open the file for binary writing
            Open FilePath For Binary Access Write As FileNum
            DataElements(3) = ""
            DataElements(1) = ""
            Recorder.Enabled = True
            Duration = 0
            RecordSize = 0
            DataElements(0) = "00:00:00"
            Lst_Data.ListItems.Item(1).SubItems(1) = "00:00:00"
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> Adjust Interface Control
            Call AdjustComponents
        Case ("Stop Capturing"):
            Rem -> ~~~~~~~~~~~~~~~~~~~
            Rem -> Destroy every thing
            Call UnloadAll
            Recorder.Enabled = False
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> Close the file we open above
            If FileNum <> 0 Then Close FileNum
            FileNum = 0
            DataElements(5) = "File Has Been Closed"
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> Adjust Interface Control
            Pic_WavDisplay.Cls
            LeftPic.Cls
            RightPic.Cls
            Lst_Data.ListItems.Item(5).SubItems(1) = "00:00:00"
            Call AdjustComponents
    End Select
End Sub
Private Sub Cmd_Intervals_Click()
    Dim Checker As String: Checker = InputBox("Please enter pre-define interval, Formated hh:mm:ss", "Enter Time", "00:00:00")
    Dim SplitedTime() As String: SplitedTime = Split(Checker, ":")
    If Checker = "" Then Exit Sub
    If UBound(SplitedTime) <> 2 Then MsgBox "Invalide format, Correct format ""hh:mm:ss""", vbCritical: Exit Sub
    If Not IsDate(Checker) Then MsgBox "Invalide format, Correct format ""hh:mm:ss""", vbCritical: Exit Sub
    Lst_Data.ListItems.Item(5).SubItems(1) = Format(Checker, "hh:mm:ss")
End Sub
Private Sub Cmd_Help_Click()
    Frm_Help.Show 1
End Sub
Private Sub Chk_Show_Click()
    Pic_WavDisplay.Cls
    LeftPic.Cls
    RightPic.Cls
End Sub
Private Sub Initializer(ByVal Channels As Integer, ByVal SamplesPerSec As Long, ByVal MP3BitRate As Integer)
    Dim RetStr                  As String
    Dim BuffLen                 As Long
    Dim BitsPerSample           As Integer
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> The user request an encoded stream, so we initialize the encoder
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Creating the actual object from encoder
    Set BE = New CEncoder
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> ALWAYS 16 !! (The blade encoder encodes only 16 bit samples)
    BitsPerSample = 16
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Initialize the blade encoder, because we need the BladeEnc.Samples value
    Rem -> wich we get only after the initialization
    If Not BE.InitStream(SamplesPerSec, IIf(Channels = 2, MP3_MODE_STEREO, MP3_MODE_MONO), MP3BitRate) Then
        MsgBox "Unable to initialize Blade Encoder, you may restart application." & vbCrLf & vbCrLf & "Try to change the quality of recording media to default values!!!", vbCritical, "Initializing Error"
        Call ExitProcess(0)
    End If
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Creating the actual object from Directx
    Set DX = New CDirectX
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem ->  BladeEnc.Samples * 6, change the 6 to a higher number if your computer is slow
    Rem ->  on P4, 1.4GHz 6 works ok, if your comptuer is slower, put 8 or 10
    RetStr = DX.Initialize(Me, SamplesPerSec, BitsPerSample * 1, Channels, BE.SamplesPerBuffer * 4)
    DX.SoundVolume = DigialVolume
    DX.SoundPlay
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Check for any errors may occure
    If Len(RetStr) > 0 Then
        MsgBox RetStr, vbExclamation
        Exit Sub
    End If
End Sub
Private Sub UnloadAll()
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Destroy DirectX Object
    If Not DX Is Nothing Then
        DX.SoundStop
        DX.UninitializeSound
        Set DX = Nothing
    End If
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Destroy Blade Encoder Object
    If Not BE Is Nothing Then
        BE.CloseStream
        Set BE = Nothing
    End If
End Sub
Public Sub AdjustComponents()
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Enable or disable the interface controls according to the App activity
    Cmd_Execute.Caption = IIf(Cmd_Execute.Caption = "Start Capturing", "Stop Capturing", "Start Capturing")
    Cmb_Channels.Enabled = IIf(Cmd_Execute.Caption = "Start Capturing", True, False)
    Cmb_SampleRate.Enabled = Cmb_Channels.Enabled
    Cmb_Bitrate.Enabled = Cmb_Channels.Enabled
    Cmd_Intervals.Enabled = Cmb_Channels.Enabled
End Sub
Private Sub DX_GotWavData(Buffer() As Byte)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~
    Rem -> Display the wave data
    If Chk_Show.Value Then ShowWAVData16Bit LeftPic, RightPic, Pic_WavDisplay, Buffer, DisplayType, IIf(Cmb_Channels.ListIndex = 0, 1, 2)
    Dim IntBuff()           As Integer
    Dim RecordedSamples     As Long
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> First make sure the RecordedSamples value is aligned -> ???
    RecordedSamples = DX.M_BufferPos / 2
    RecordedSamples = RecordedSamples - (RecordedSamples Mod 4)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If RecordedSamples >= BE.SamplesPerBuffer Then
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Blade encoder can encode only a buffer with a length of multiple
        Rem -> Of BE.Samples -> ???
        RecordedSamples = RecordedSamples - (RecordedSamples Mod BE.SamplesPerBuffer)
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Fill IntBuff array with the wav data, after we aligned it -> ???
        ReDim IntBuff(RecordedSamples - 1)
        CopyMemory IntBuff(0), ByVal DX.Ptr_WavBuffer, RecordedSamples * 2
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Encode the WAVE data to MP3 data
        If Not BE.EncodeChunk(IntBuff) Then
            Rem -> This should not happen, because the initializing of the BE
            Rem -> Shoud've returned an error. So if here it does not work,
            Rem -> Have no idea why it would not
            Debug.Print "Error in Encode Chunk"
        End If
        Erase IntBuff
        Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Rem -> Decrement the buffer position by how much data was converted
        DX.M_BufferPos = DX.M_BufferPos - (RecordedSamples * 2)
        If DX.M_BufferPos < 0 Then
            DX.M_BufferPos = 0
        Else
            Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            Rem -> Move the WAVE data that was not converted to MP3 format to
            Rem -> The beginning of the buffer
            CopyMemoryB DX.Ptr_WavBuffer, DX.Ptr_WavBuffer + (RecordedSamples * 2), DX.M_BufferPos
        End If
    End If
End Sub
Private Sub BE_GotMP3Data(Buffer() As Byte)
    DataElements(2) = UBound(Buffer) - LBound(Buffer) + 1 & " Byte"
    DataElements(3) = Val(DataElements(3)) + 1
    RecordSize = RecordSize + Val(DataElements(2))
    DataElements(1) = GetFileSize(RecordSize)
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> We just add the captured chunk to the file, using put function
    Put FileNum, , Buffer
End Sub
Private Sub Recorder_Timer()
    Duration = Duration + 1
    DataElements(0) = SecToTime(Duration)
    Lst_Data.ListItems.Item(1).SubItems(1) = DataElements(0)
    Lst_Data.ListItems.Item(2).SubItems(1) = DataElements(1)
    Lst_Data.ListItems.Item(3).SubItems(1) = DataElements(2)
    Lst_Data.ListItems.Item(4).SubItems(1) = DataElements(3)
    Lst_Data.ListItems.Item(6).SubItems(1) = DataElements(5)
    If (Lst_Data.ListItems.Item(1).SubItems(1) = Lst_Data.ListItems.Item(5).SubItems(1)) Then Call Cmd_Execute_Click
End Sub

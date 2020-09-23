VERSION 5.00
Begin VB.Form sndPlayer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "sndPlayer"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   1590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   495
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   0
      Pattern         =   "*.mid;*.mp3;*.mpe;*.wav"
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   495
   End
End
Attribute VB_Name = "sndPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub sndPlayW(Filename As String)
    Call sndPlaySound(Filename, SND_ASYNC Or SND_NODEFAULT Or SND_NOSTOP)
End Sub
Public Sub sndPlayM(Filename As String)
    Call mciSendString("Open " & Filename & " Alias MM", 0, 0, 0)
    Call mciSendString("Play MM", 0, 0, 0)
End Sub

Public Sub sndPauseM()
    Call mciSendString("Stop MM", 0, 0, 0)
End Sub
Public Sub sndStartM()
    If Mid(File1.Filename, (Len(File1.Filename) - 3), 4) = ".mid" Then
        sndPlayM (File1.Path & "\" & File1.Filename)
    ElseIf Mid(File1.Filename, (Len(File1.Filename) - 3), 4) = ".mp3" Then
        sndPlayM (File1.Path & "\" & File1.Filename)
    ElseIf Mid(File1.Filename, (Len(File1.Filename) - 4), 5) = ".mpeg" Then
        sndPlayM (File1.Path & "\" & File1.Filename)
    ElseIf Mid(File1.Filename, (Len(File1.Filename) - 3), 4) = ".wav" Then
        sndPlayW (File1.Path & "\" & File1.Filename)
    End If
End Sub

Public Sub sndStopM()
    Call mciSendString("Stop MM", 0, 0, 0)
    Call mciSendString("Close MM", 0, 0, 0)
End Sub

Private Sub cmdPause_Click()
    sndPauseM
End Sub

Private Sub cmdPlay_Click()
    sndStartM
End Sub

Private Sub cmdStop_Click()
    sndStopM
End Sub

Private Sub Form_Load()
    Call MsgBox("sndPlayer - By: Argonaut" & Chr(13) & Chr(13) & "This is a simple little app that" & Chr(13) & "will show you have to Start," & Chr(13) & "Pause and Stop different Media" & Chr(13) & "Files such as Midi's, Mpeg's," & Chr(13) & "Mp3's, and Wav's." & Chr(13) & "Everything is done by calls," & Chr(13) & "No more controls needed." & Chr(13) & "More stuff to come later." & Chr(13) & Chr(13) & "www.vbstuff.cjb.net", , "sndPlayer About")
    File1.Path = "C:\windows\desktop\"
End Sub

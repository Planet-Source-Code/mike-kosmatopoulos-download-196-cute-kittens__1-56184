VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "196 kitten pictures to your computer"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Cancel          =   -1  'True
      Caption         =   "Pause"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdDownload 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Download Pictures!"
      Height          =   315
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   2415
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   196
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblDoing 
      BackStyle       =   0  'Transparent
      Caption         =   "Ready"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Image imgKitten 
      Height          =   1575
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Api Declarations
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long 'Api to download a file

Dim strFileName  As String  'The filename that the user selected
Dim strURLOfFile As String  'The URL of the picture that will be downloaded next
Dim strFilepath  As String  'The path of the common dialog control
Dim intStoped    As Integer 'The number where the user canceled
Dim bCancel      As Boolean 'If the user wants to cancel the downloading

Private Sub cmdCancel_Click()
    If cmdCancel.Caption = "Pause" Then
        bCancel = True 'The user wants to cancel
        cmdCancel.Caption = "Resume" 'Set the button's caption to "Resume"
        lblDoing.Caption = "Paused" 'Update the label's caption
    ElseIf cmdCancel.Caption = "Resume" Then
        bCancel = False 'The user wants to resume
        cmdCancel.Caption = "Pause" 'Set the button's caption to "Cancel"
        Call DownloadKittens 'Call the sub to download the kittens
        
        If bCancel = False Then 'If the user didn't cancel
            MsgBox "Successfully downloaded the pictures!", vbInformation, "Done!" 'Show a message that says that we saved the pictures successfully
            Unload Me
            End 'End the program
        End If

    End If
End Sub

Private Sub cmdDownload_Click()
    On Error GoTo err_handler 'If there is an error, go to our error handler

    With Dialog
        .CancelError = True 'Display an error if the user clicks "Cancel"
        .DialogTitle = "Location to save the kitten pictures" 'The title of our common dialog control
        .Filter = "*.jpg" 'Set the filter
        .ShowSave 'Show the save dialog
    End With
    
    strFileName = Dialog.FileTitle 'Copy the dialog's filename to our filename string
    
    
    Call DownloadKittens 'Call the sub to download our kittens
    
    If bCancel = False Then 'If the user didn't cancel
        MsgBox "Successfully downloaded the pictures!", vbInformation, "Done!" 'Show a message that says that we saved the pictures successfully
        Unload Me
        End 'End the program
    End If
    
    Exit Sub 'We have to exit the sub or else the error handler will be activated
    
err_handler: 'Our Error Handler
    If Err.Number = 32755 Then Exit Sub 'The error number 32755 is generated from the common dialog if the user clicks cancel.

    MsgBox "Could not download the kitten pictures to your computer!" & vbCrLf & "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description

End Sub
Private Sub DownloadKittens()
    On Error GoTo err_handler 'If there is an error, go to our error handler

    cmdDownload.Visible = False 'Don't show the "Download" button
    PBar.Visible = True 'Show the Progress Bar
    cmdCancel.Visible = True 'Show the cancel button
    
    strFilepath = InStrRev(Dialog.FileName, "\") 'Look for "\" backforwards
    strFilepath = Left$(Dialog.FileName, strFilepath)

    Dim i As Integer
    
    For i = intStoped To 196 'Initialize a loop from the number the user paused to 196 (the number of pictures)
        DoEvents 'Let the system do what it needs to do
        strURLOfFile = "http://kittens.sytes.org/kitten" & CStr(i) & ".jpg" 'Set the url of the file we need to download
        DownloadFile strURLOfFile, strFilepath & strFileName & CStr(i) & ".jpg" 'Download the picture
        imgKitten.Picture = LoadPicture(strFilepath & strFileName & CStr(i) & ".jpg") 'Set the image box to the picture we just downloaded
        PBar.Value = PBar.Value + 1 'Update the progress bar
        lblDoing.Caption = "Picture " & CStr(i) & " of 196" 'Update the label
        DoEvents 'Let the system do what it needs to do
        If bCancel = True Then 'If the user canceled
            PBar.Value = PBar.Value - 1 'Remove one from the progressbar because we already added one
            intStoped = i 'Update the number used for resuming
            Exit For 'Exit the loop
        End If
    Next i
    
    Exit Sub

err_handler:
    MsgBox "Could not download the kitten pictures to your computer!" & vbCrLf & "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description

End Sub

Private Function DownloadFile(URL As String, LocalFilename As String) As Boolean
'Function from AllApi.Net
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
End Function


Private Sub Form_Load()
intStoped = 1 'Set the number used for resuming at 1 because we will have problems if we don't set it
End Sub

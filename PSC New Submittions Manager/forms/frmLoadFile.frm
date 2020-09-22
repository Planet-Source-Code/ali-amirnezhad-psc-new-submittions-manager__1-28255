VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLoadFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load File"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6165
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   4320
      TabIndex        =   6
      Top             =   2970
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   5250
      TabIndex        =   5
      Top             =   2970
      Width           =   885
   End
   Begin SHDocVwCtl.WebBrowser webNewList 
      Height          =   1965
      Left            =   30
      TabIndex        =   4
      Top             =   960
      Width           =   6105
      ExtentX         =   10769
      ExtentY         =   3466
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComDlg.CommonDialog CDlgLoad 
      Left            =   2100
      Top             =   390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   285
      Left            =   5520
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   600
      Width           =   5385
   End
   Begin VB.Label lblFileName 
      Caption         =   "File Name:"
      Height          =   195
      Left            =   30
      TabIndex        =   1
      Top             =   390
      Width           =   765
   End
   Begin VB.Label lblNotice 
      Caption         =   "Notice: You can load only html source of mails sended from PSC (Planet Source Code)"
      Height          =   225
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6105
   End
End
Attribute VB_Name = "frmLoadFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  blnLoadFile = False
  Unload Me
End Sub

Private Sub cmdLoad_Click()
  With CDlgLoad
    .DialogTitle = "Open New Submittions Lise File"
    .FileName = ""
    .Filter = "Text File (*.txt)|*.txt|HTML Document (*.html,*.htm)|*.html;*.htm"
    .FilterIndex = 1
    .InitDir = App.Path
    
    .ShowOpen
    
    If (.FileName = "") Then Exit Sub
    
    txtFileName.Text = .FileName
    txtFileName.SelStart = Len(txtFileName.Text)
    Call ClearHTMLTags(.FileName, strTempTEXTFileName)
    Call GenerateNewListCollection(strTempTEXTFileName)
    Call CreateTempHTMLFile
    Call webNewList.Navigate(strTempHTMLFileName)
  End With
End Sub

Private Sub cmdOK_Click()
  blnLoadFile = True
  Unload Me
End Sub
    
Private Sub Form_Load()
  Me.Icon = frmMain.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If (isFileExists(strTempHTMLFileName)) Then Call Kill(strTempHTMLFileName)
  If (isFileExists(strTempTEXTFileName)) Then Call Kill(strTempTEXTFileName)
End Sub

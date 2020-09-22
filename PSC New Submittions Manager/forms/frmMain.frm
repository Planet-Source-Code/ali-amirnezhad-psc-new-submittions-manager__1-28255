VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PSC New Submits Manager"
   ClientHeight    =   5610
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9150
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraNew 
      Height          =   4965
      Left            =   150
      TabIndex        =   1
      Top             =   450
      Width           =   8775
      Begin SHDocVwCtl.WebBrowser webNew 
         Height          =   2355
         Left            =   90
         TabIndex        =   4
         Top             =   2490
         Width           =   8595
         ExtentX         =   15161
         ExtentY         =   4154
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
      Begin MSDBGrid.DBGrid dbgNew 
         Bindings        =   "frmMain.frx":0CCA
         Height          =   2295
         Left            =   90
         OleObjectBlob   =   "frmMain.frx":0CDF
         TabIndex        =   3
         Top             =   180
         Width           =   8595
      End
   End
   Begin VB.Data datDownloaded 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Uesr Directories\ali\Programming\Visual Basic\PSC New Submittions Manager\db\main.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Codes"
      Top             =   0
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Frame fraDownloads 
      Height          =   4965
      Left            =   150
      TabIndex        =   2
      Top             =   450
      Width           =   8775
      Begin SHDocVwCtl.WebBrowser webDownloaded 
         Height          =   4695
         Left            =   90
         TabIndex        =   5
         Top             =   210
         Width           =   8565
         ExtentX         =   15108
         ExtentY         =   8281
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
   End
   Begin VB.Data datNew 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Uesr Directories\ali\Programming\Visual Basic\PSC New Submittions Manager\db\main.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7890
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Codes"
      Top             =   0
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MSComctlLib.TabStrip tbsData 
      Height          =   5415
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9551
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&New"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Downloaded"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileLoad 
         Caption         =   "&Load New"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFileBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub CreateNewHTMLFile()
  Dim lngNumber As Long
  
  lngNumber = 0
  Open strNewHTMLFile For Output As #10
  Print #10, "<html>"
  Print #10, ""
  Print #10, "<head>"
  Print #10, "<style type=""text/css"">"
  Print #10, "body {"
  Print #10, "  margin: 0pt;"
  Print #10, "  font-weight: bold;"
  Print #10, "  font-family: arial;"
  Print #10, "  font-size: 10pt;"
  Print #10, "  color: black;"
  Print #10, "  bgcolor: white;"
  Print #10, "}"
  Print #10, "</style>"
  Print #10, "</head>"
  Print #10, ""
  Print #10, "<body>"
  If (dbgNew.Row <> -1) Then
    Print #10, "<div align=""center"">" + datNew.Recordset.Fields("DateAdd").Value + "</div><br><br>"
    Print #10, datNew.Recordset.Fields("Name").Value + "<br>"
    Print #10, "<a target=""_blank"" href=""" + datNew.Recordset.Fields("URL").Value + """>" + datNew.Recordset.Fields("URL").Value + "</a><br><br>"
    Print #10, datNew.Recordset.Fields("Description").Value
  End If
  Print #10, "</body>"
  Print #10, ""
  Print #10, "</html>"
  Close #10
End Sub

Private Sub setWindow()
  dbgNew.Refresh
  Call CreateNewHTMLFile
  Call CreateDownloadedHTMLFile
  Call webNew.Navigate(strNewHTMLFile)
  Call webDownloaded.Navigate(strDownloadedHTMLFile)
End Sub

Private Sub dbgNew_DblClick()
  If (MsgBox("Do you download selected file?" + Chr(13) + "Name: " + datNew.Recordset.Fields("Name").Value + Chr(13) + "URL: " + datNew.Recordset.Fields("URL").Value, vbQuestion + vbYesNo) = vbYes) Then
    datNew.Recordset.Edit
    datNew.Recordset.Fields("DateDownload").Value = CStr(Now)
    datNew.Recordset.Fields("Downloaded").Value = "1"
    datNew.Recordset.Update
    Call setWindow
  End If
End Sub

Private Sub dbgNew_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  setWindow
End Sub

Private Sub Form_Activate()
  Set datNew.Recordset = datNew.Database.OpenRecordset("Select * From Codes Where Downloaded = ""0""")
  Set datDownloaded.Recordset = datNew.Database.OpenRecordset("Select * From Codes Where Downloaded = ""1""")
  setWindow
End Sub

Private Sub Form_Load()
  strTempHTMLFileName = App.Path
  If (Mid(strTempHTMLFileName, Len(strTempHTMLFileName), 1) <> "\") Then strTempHTMLFileName = strTempHTMLFileName + "\"
  strTempTEXTFileName = strTempHTMLFileName + "temp.txt"
  strProgramDBFileName = strTempHTMLFileName + "db\main.mdb"
  strNewHTMLFile = strTempHTMLFileName + "new.html"
  strDownloadedHTMLFile = strTempHTMLFileName + "downloaded.html"
  strTempHTMLFileName = strTempHTMLFileName + "temp.html"
  datNew.DatabaseName = strProgramDBFileName
  fraNew.Visible = True
  fraDownloads.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If (isFileExists(strNewHTMLFile)) Then Kill (strNewHTMLFile)
  If (isFileExists(strDownloadedHTMLFile)) Then Kill (strDownloadedHTMLFile)
  If (isFileExists(strNewHTMLFile)) Then Kill (strNewHTMLFile)
End Sub

Private Sub mnuFileExit_Click()
  Unload Me
End Sub

Private Sub mnuFileLoad_Click()
  frmLoadFile.Show 1
  If (Not (blnLoadFile)) Then Exit Sub
  Call AddNewDataIntoDatabase
  Set datNew.Recordset = datNew.Database.OpenRecordset("Select * From Codes Where Downloaded = ""0""")
  Set datDownloaded.Recordset = datNew.Database.OpenRecordset("Select * From Codes Where Downloaded = ""1""")
  Call setWindow
End Sub

Private Sub tbsData_Click()
  fraNew.Visible = False
  fraDownloads.Visible = False
  Select Case tbsData.SelectedItem.Index
    Case 1: fraNew.Visible = True
    Case 2: fraDownloads.Visible = True
  End Select
End Sub

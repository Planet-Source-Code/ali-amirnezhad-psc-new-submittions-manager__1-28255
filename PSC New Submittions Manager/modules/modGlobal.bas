Attribute VB_Name = "modGlobal"
Option Explicit

Public strTempHTMLFileName As String
Public strTempTEXTFileName As String
Public strProgramDBFileName As String
Public strNewHTMLFile As String
Public strDownloadedHTMLFile As String
Public colNames As New Collection
Public colURLs As New Collection
Public colDescriptions As New Collection
Public strDate As String
Public blnLoadFile As Boolean

Public Function ClearCollection(ByVal colTemp As Collection) As Collection
  Dim intCounter As Integer
  
  For intCounter = 1 To colTemp.Count
    Call colTemp.Remove(1)
  Next intCounter
  Set ClearCollection = colTemp
End Function

Public Sub ClearHTMLTags(ByVal strSourceFileName As String, ByVal strTargetFileName As String)
  Dim strFile, strLine As String
  
  Open strSourceFileName For Input As #5
  While (Not (EOF(5)))
    Line Input #5, strLine
    If (strLine <> "") Then strFile = strFile + strLine + vbCrLf
  Wend
  Close #5
  While (InStr(1, strFile, "<") <> 0)
    strFile = Mid(strFile, 1, InStr(1, strFile, "<") - 1) + Mid(strFile, InStr(1, strFile, ">") + 1)
  Wend
  Open strTargetFileName For Output As #6
  Print #6, strFile
  Close #6
End Sub

Public Sub GenerateNewListCollection(ByVal strFileName As String)
  Dim strLine As String
  Dim strDescription As String
  
  Set colNames = ClearCollection(colNames)
  Set colURLs = ClearCollection(colURLs)
  Set colDescriptions = ClearCollection(colDescriptions)
  Open strFileName For Input As #5
  While (Not (EOF(5)))
    Line Input #5, strLine
    If (InStr(1, strLine, "Date: ") = 1) Then strDate = Mid(strLine, 7)
    If (InStr(1, strLine, "Code of the Day:") = 1) Then
      While (strLine <> "================================================")
        If (InStr(1, strLine, ")") <> 0) Then
          colNames.Add Item:=strLine
        End If
        Line Input #5, strLine
      Wend
      While (Not (EOF(5)))
        If (InStr(1, strLine, "Description: ") = 1) Then
          strDescription = ""
          While (strLine <> "Complete source code is at:")
            strDescription = strDescription + " " + strLine
            Line Input #5, strLine
          Wend
          colDescriptions.Add Item:=Mid(strDescription, 15)
          Line Input #5, strLine
          colURLs.Add Item:=strLine
        End If
        Line Input #5, strLine
      Wend
    End If
  Wend
  Close #5
End Sub

Public Sub CreateTempHTMLFile()
  Dim intCounter As Integer
  Dim strName, strURL As String
  
  Open strTempHTMLFileName For Output As #10
  Print #10, "<html>"
  Print #10, ""
  Print #10, "<head>"
  Print #10, "<style type=""text/css"">"
  Print #10, "body {"
  Print #10, "  font-family: arial;"
  Print #10, "  font-size: 8pt;"
  Print #10, "  color: black;"
  Print #10, "  bgcolor: white;"
  Print #10, "}"
  Print #10, "</style>"
  Print #10, "</head>"
  Print #10, ""
  Print #10, "<body>"
  For intCounter = 1 To colNames.Count
    strName = colNames.Item(intCounter)
    strURL = colURLs.Item(intCounter)
    Print #10, "<div title=""URL:" + strURL + """><b>" + strName + "</b></div><br>"
  Next intCounter
  Print #10, "</body>"
  Print #10, ""
  Print #10, "</html>"
  Close #10
End Sub

Public Sub AddNewDataIntoDatabase()
  Dim datTemp As Database
  Dim recTemp As Recordset
  Dim intCounter As Integer
  Dim strURL, strName, strDescription As String
  
  Set datTemp = OpenDatabase(strProgramDBFileName)
  Set recTemp = datTemp.OpenRecordset("Select * From Codes", dbOpenDynaset)
  For intCounter = 1 To colNames.Count
    If (intCounter < 10) Then
      strName = Mid(colNames.Item(intCounter), 4)
    Else
      strName = Mid(colNames.Item(intCounter), 5)
    End If
    strURL = colURLs.Item(intCounter)
    strDescription = colDescriptions.Item(intCounter)
    recTemp.FindFirst "[URL]=""" + strURL + """"
    If (recTemp.NoMatch) Then
      recTemp.AddNew
      recTemp.Fields("Name").Value = strName
      recTemp.Fields("URL").Value = strURL
      recTemp.Fields("Description").Value = strDescription
      recTemp.Fields("DateAdd").Value = strDate
      recTemp.Fields("Downloaded").Value = "0"
      recTemp.Update
    Else
      If (MsgBox("You have this file in database" + Chr(13) + Chr(13) + _
                 "Database: " + Chr(13) + "Name: " + recTemp.Fields("Name").Value + Chr(13) + "URL: " + recTemp.Fields("URL").Value + Chr(13) + Chr(13) + _
                 "File: " + Chr(13) + "Name: " + strName + Chr(13) + "URL: " + strURL + Chr(13) + Chr(13) + _
                 "Do you want to add as new?", vbQuestion + vbYesNo) = vbYes) Then
        recTemp.AddNew
        recTemp.Fields("Name").Value = strName
        recTemp.Fields("URL").Value = strURL
        recTemp.Fields("Description").Value = strDescription
        recTemp.Fields("DateAdd").Value = strDate
        recTemp.Fields("Downloaded").Value = "0"
        recTemp.Update
      End If
    End If
  Next intCounter
  recTemp.Close
  datTemp.Close
End Sub

Public Sub CreateDownloadedHTMLFile()
  Dim datTemp As Database
  Dim recTemp As Recordset
  Dim lngNumber As Long
  
  lngNumber = 0
  Open strDownloadedHTMLFile For Output As #10
  Print #10, "<html>"
  Print #10, ""
  Print #10, "<head>"
  Print #10, "<style type=""text/css"">"
  Print #10, "body {"
  Print #10, "  margin: 0pt;"
  Print #10, "  font-family: arial;"
  Print #10, "  font-weight: bold;"
  Print #10, "  font-size: 10pt;"
  Print #10, "  color: black;"
  Print #10, "  bgcolor: white;"
  Print #10, "}"
  Print #10, "</style>"
  Print #10, "</head>"
  Print #10, ""
  Print #10, "<body>"
  Set datTemp = OpenDatabase(strProgramDBFileName)
  Set recTemp = datTemp.OpenRecordset("Select * From Codes Where Downloaded = ""1""")
  If (recTemp.RecordCount <> 0) Then
    recTemp.MoveFirst
    While (Not (recTemp.EOF))
      lngNumber = lngNumber + 1
      Print #10, Format(lngNumber, "0###") + " -> <br><a target=""_blank"" href=""" + recTemp.Fields("URL").Value + """>" + recTemp.Fields("URL").Value + "</a><br>" + recTemp.Fields("Name").Value + "<br>Downloaded at: " + recTemp.Fields("DateDownload").Value + "<br><br>"
      recTemp.MoveNext
    Wend
  End If
  recTemp.Close
  datTemp.Close
  Print #10, "</body>"
  Print #10, ""
  Print #10, "</html>"
  Close #10
End Sub

Public Function isFileExists(ByVal strFileName As String) As Boolean
  On Error GoTo HaveError
  isFileExists = True
  Open strFileName For Input As #10
  Close #10
  Exit Function
  
HaveError:
  isFileExists = False
End Function


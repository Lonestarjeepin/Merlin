'-------------------------------------------------------------------------
' Module  : clsUpdate
' Company : JKP Application Development Services (c)
' Author  : Jan Karel Pieterse
' Created : 19-2-2007
' Purpose : Class to check for program updates
'-------------------------------------------------------------------------
Option Explicit


Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
                  Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, _
                                            ByVal szURL As String, ByVal szFileName As String, _
                                            ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As LongPtr

Private mdtLastUpdate As Date

Private msAppName As String
Private msBuild As String
Private msCheckURL As String
Private msCurrentAddinName As String
Private msDownloadName As String
Private msTempAddInName As String
Private mbManual As Boolean
Private msNewBuild As String

Private Sub DownloadFile(strWebFilename As String, strSaveFileName As String)
  ' Download the file.
  URLDownloadToFile 0, strWebFilename, strSaveFileName, 0, 0
End Sub

Public Function IsThereAnUpdate(sError As String, Optional bShowMsg As Boolean = False) As Boolean
'-------------------------------------------------------------------------
' Procedure : IsThereAnUpdate
' Company  : JKP Application Development Services (c)
' Author  : Jan Karel Pieterse (www.jkp-ads.com)
' Created  : 2-6-2009
' Purpose  : Fetches version information from a webpage
'-------------------------------------------------------------------------
  Dim oHTTP As Object
  Dim sURL As String
  Dim sTextResponse As String
  Dim varStart As Integer
  Dim varEnd As Integer
  Dim tagName As String
  Dim lError As LongPtr
   
  On Error GoTo LocErr
  sURL = CheckURL
  Set oHTTP = CreateObject("MSXML2.XMLHTTP")
  NewBuild = 0
  lError = 0
  Do
    With oHTTP
      On Error Resume Next
      .Open "GET", sURL, False
      .setRequestHeader "User-Agent", "VBA"
      .send
       
      If .Status = 200 Then
        sTextResponse = .responseText
        ' Find the "tag_name" key in the JSON response
        varStart = InStr(sTextResponse, """tag_name"":")
        If varStart > 0 Then
          varStart = varStart + Len("""tag_name"":") + 1 ' Move past "tag_name": and the starting quote
          varEnd = InStr(varStart, sTextResponse, """") ' Find the ending quote
          tagName = Mid(sTextResponse, varStart, varEnd - varStart)
          sError = ""
          NewBuild = tagName
          IsThereAnUpdate = True
          Exit Do
        Else
          sError = "Error: 'tag_name' not found in the response."
          lError = lError + 1
          IsThereAnUpdate = False
        End If
      Else
        sError = "Error fetching version information from GitHub API. Please try downloading directly from https://github.com/Lonestarjeepin/Merlin"
        lError = lError + 1
        IsThereAnUpdate = False
      End If

      On Error GoTo 0
      If lError > 2 Then GoTo TidyUp
    End With
  Loop

TidyUp:
  On Error Resume Next
  Set oHTTP = Nothing
  Exit Function
LocErr:
  Select Case ReportError(Err.Description, Err.Number, "IsThereAnUpdate", "Class Module clsUpdate")
  Case vbRetry
    Resume
  Case vbIgnore
    Resume Next
  Case vbAbort
    Resume TidyUp
  End Select
End Function
Public Sub DoUpdate()
                           
  Dim sError As String
  On Error GoTo LocErr
  If IsThereAnUpdate(sError, Manual) Then
    If CLng(Val(NewBuild)) > CLng(Build) Then
      If MsgBox("There is an update for " & AppName & "." & _
               vbNewLine & "Do you wish to download now?", vbQuestion + vbYesNo, AppName) = vbYes Then
        If GetUpdate Then
          Application.Cursor = xlDefault
          MsgBox "Successfully updated the " & AppName & " Add-In, " & vbNewLine & "please restart Excel to start using the new version!", vbOKOnly + vbInformation, AppName
        Else
          Application.Cursor = xlDefault
          MsgBox "Updating " & AppName & " has failed, please try again later.", vbInformation + vbOKOnly, AppName
        End If
      End If
    ElseIf Manual Then
      Application.Cursor = xlDefault
      MsgBox AppName & " is up to date.", vbInformation + vbOKOnly, AppName
    End If
  Else
    MsgBox "Error fetching version information: " & sError, vbExclamation + vbOKOnly, AppName
  End If
TidyUp:
  On Error GoTo 0
  Exit Sub
LocErr:
  Select Case ReportError(Err.Description, Err.Number, "DoUpdate", "Class Module clsUpdate")
  Case vbRetry
    Resume
  Case vbIgnore
    Resume Next
  Case vbAbort
    Resume TidyUp
  End Select
End Sub

Public Property Get Build() As String
  Build = msBuild
End Property

Public Property Let Build(ByVal sBuild As String)
  msBuild = sBuild
End Property

Public Sub RemoveOldCopy()
  On Error GoTo LocErr
  CurrentAddinName = ThisWorkbook.FullName
  TempAddInName = CurrentAddinName & "(OldVersion)"
  On Error Resume Next
  Kill TempAddInName
TidyUp:
  On Error GoTo 0
  Exit Sub
LocErr:
  Select Case ReportError(Err.Description, Err.Number, "RemoveOldCopy", "Class Module clsUpdate")
  Case vbRetry
    Resume
  Case vbIgnore
    Resume Next
  Case vbAbort
    Resume TidyUp
  End Select
End Sub

Public Function GetUpdate() As Boolean
  On Error Resume Next
  'If workbook has been saved readonly, we can safely delete the file!
  If ThisWorkbook.ReadOnly Then
    Err.Clear
                             
          
                    
                                     
            
    Kill CurrentAddinName
  End If
  LastUpdate = Now
  ThisWorkbook.SaveAs TempAddInName
  DoEvents
  Kill CurrentAddinName
  On Error GoTo 0
  DownloadFile DownloadName, CurrentAddinName
  If Err = 0 Then GetUpdate = True
End Function

Private Property Get CurrentAddinName() As String
  CurrentAddinName = msCurrentAddinName
End Property

Private Property Let CurrentAddinName(ByVal sCurrentAddinName As String)
  msCurrentAddinName = sCurrentAddinName
End Property

Private Property Get TempAddInName() As String
  TempAddInName = msTempAddInName
End Property

Private Property Let TempAddInName(ByVal sTempAddInName As String)
  msTempAddInName = sTempAddInName
End Property

Public Property Get DownloadName() As String
  DownloadName = msDownloadName
End Property

Public Property Let DownloadName(ByVal sDownloadName As String)
  msDownloadName = sDownloadName
End Property

Public Property Get CheckURL() As String
  CheckURL = msCheckURL
End Property

Public Property Let CheckURL(ByVal sCheckURL As String)
  msCheckURL = sCheckURL
End Property

Public Property Get LastUpdate() As Date
  Dim dtNow As Date
  On Error GoTo LocErr
  dtNow = Int(Now)
  mdtLastUpdate = CDate(GetSetting(AppName, "Updates", "LastUpdate", "0"))
  If mdtLastUpdate = 0 Then
    'Never checked for an update, save today!
    LastUpdate = dtNow
  End If
  LastUpdate = mdtLastUpdate
TidyUp:
  On Error GoTo 0
  Exit Property
LocErr:
  Select Case ReportError(Err.Description, Err.Number, "LastUpdate", "Class Module clsUpdate")
  Case vbRetry
    Resume
  Case vbIgnore
    Resume Next
  Case vbAbort
    Resume TidyUp
  End Select
End Property

Public Property Let LastUpdate(ByVal dtLastUpdate As Date)
  mdtLastUpdate = dtLastUpdate
  SaveSetting AppName, "Updates", "LastUpdate", CStr(CLng(mdtLastUpdate))
End Property

Public Property Get AppName() As String
  AppName = msAppName
End Property

Public Property Let AppName(ByVal sAppName As String)
  msAppName = sAppName
End Property

Public Property Get Manual() As Boolean
  Manual = mbManual
End Property

Public Property Let Manual(ByVal bManual As Boolean)
  mbManual = bManual
End Property

Public Property Get NewBuild() As String
  NewBuild = msNewBuild
End Property

Public Property Let NewBuild(ByVal sNewBuild As String)
  msNewBuild = sNewBuild
End Property



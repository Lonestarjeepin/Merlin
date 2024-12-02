Attribute VB_Name = "d_modUpdate"
Option Explicit

Dim mcUpdate As clsUpdate

Public Const GSAPPNAME As String = "Update Addin Demo"
Public Declare PtrSafe Function InternetGetConnectedState _
                         Lib "wininet.dll" (lpdwFlags As LongPtr, _
                                            ByVal dwReserved As LongPtr) As Boolean

Function IsConnected() As Boolean
    Dim Stat As LongPtr
    IsConnected = (InternetGetConnectedState(Stat, 0&) <> 0)
End Function

Sub ManualUpdate()
    On Error Resume Next
    Call CheckAndUpdate
    'Application.OnTime Now, "'" & ThisWorkbook.FullName & "'!CheckAndUpdate"
End Sub
Public Sub CheckAndUpdate(Optional bManual As Boolean = True)
   
    On Error GoTo LocErr
    Set mcUpdate = New clsUpdate
    If bManual Then
        Application.Cursor = xlWait
    End If
    With mcUpdate
        'Set intial values of class
        
'**************************************************************************************************************************************************
'**************************************************************************************************************************************************
'**************************************************************************************************************************************************
        'Current build
        .Build = 20
'**************************************************************************************************************************************************
'**************************************************************************************************************************************************
'**************************************************************************************************************************************************
        'Name of this app, probably a global variable, such as appname
        .AppName = "Merlin"
        'Get rid of possible old backup copy
        .RemoveOldCopy
        'URL which contains build # of new version
        .CheckURL = "https://merlinaddin.xyz/Merlin/Merlin_Build_Number.html"
        .DownloadName = "https://merlinaddin.xyz/Merlin/Merlin.xlam"

        'Started check automatically or manually?
        .Manual = bManual
        'Check once a week
        If (Now - .LastUpdate >= 7) Or bManual Then
            .LastUpdate = Int(Now)
            .DoUpdate
        End If
        Set mcUpdate = Nothing
    End With
TidyUp:
    On Error GoTo 0
    Application.Cursor = xlDefault
    Exit Sub
LocErr:
    Select Case ReportError(Err.Description, Err.Number, "CheckAndUpdate", "Module modUpdate")
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    Case vbAbort
        Resume TidyUp
    End Select
End Sub


'-------------------------------------------------------------------------
' Module    : modErrorHandler
' Company   : JKP Application Development Services (c) 2006
' Author    : Jan Karel Pieterse
' Created   : 17-1-2006
' Purpose   : Main error handling functions
'-------------------------------------------------------------------------

Function ReportError(Description As String, Number As LongPtr, ModuleName As String, _
                     ProcName As String)
'-------------------------------------------------------------------------
' Procedure : ReportError Created by Jan Karel Pieterse
' Company   : JKP Application Development Services (c) 2006
' Author    : Jan Karel Pieterse
' Created   : 17-1-2006
' Purpose   : Function to report errors and log them
'-------------------------------------------------------------------------
    Dim stMsg As String
    Dim l As LongPtr
    Dim lAnswer As LongPtr
    On Error Resume Next    ' in case of errors in here, continue
    stMsg = "Error " & Number & ": " & Description & " in " & ModuleName & "." & ProcName

    ' put details in Immediate window
    Debug.Print stMsg

    'output to a text file
    l = FreeFile()  ' get free file number
    Open ThisWorkbook.Path & Application.PathSeparator & GSAPPNAME & " Errors.Log" For Append As #l
    Print #l, Now, ThisWorkbook.Name, stMsg
    Close #l
    ' message for the user
    lAnswer = MsgBox("Oops, An error has occurred in " & GSAPPNAME & vbNewLine & "Error:" & vbNewLine & stMsg, vbAbortRetryIgnore + vbExclamation, GSAPPNAME & ": Error message")
    ' return vbRetry if debugging - causes Resume
    If lAnswer = vbRetry Then
        ReportError = vbRetry
        'Next statement will be ignored if project is protected
        Stop              ' press F8 multiple times to see the statement that caused the error
    ElseIf lAnswer = vbAbort Then
        ReportError = vbAbort
    Else
        ReportError = vbIgnore
    End If
TidyUp:
    Exit Function
LocErr:
    Select Case ReportError(Err.Description, Err.Number, "modErrorHandler", "ReportError")
    Case vbRetry
        Resume
    Case vbAbort
        Resume TidyUp
    Case vbIgnore
        Resume Next
    End Select
End Function


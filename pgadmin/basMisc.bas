Attribute VB_Name = "basMisc"
' pgAdmin - PostgreSQL db Administration/Management for Win32
' Copyright (C) 1998 - 2001, Dave Page

' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.

' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.

' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

Option Explicit

Sub Main()
On Error GoTo Err_Handler
Dim Res As String
Dim i As Long
  frmSplash.Show
  frmSplash.Refresh
  Logging = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Logging", 0)
  MaskPassword = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Mask Password", 1)
  LogFile = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Log File", "C:\pgAdmin.log")
  BBar = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Button Bar", 1)
  SQLPane = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "SQL Pane", 1)
  TextColours = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Text Colours", DEFAULT_TEXT_COLOURS)
  LogStartup
  
  'Load the installed exporters
  
  ReDim Exporters(0)
  Res = RegGetSubkey(HKEY_CLASSES_ROOT, "", i)
  Do Until Res = "Not Found"
    If InStr(1, Res, "pgAdmin_Exporter") Then
      Set Exporters(UBound(Exporters)) = CreateObject(Res)
      LogInitMsg "Loading Exporter: " & Res & " (" & Exporters(UBound(Exporters)).Description & " v" & Exporters(UBound(Exporters)).Version & ")"
      ReDim Preserve Exporters(UBound(Exporters) + 1)
Continue:
    End If
    i = i + 1
    Res = RegGetSubkey(HKEY_CLASSES_ROOT, "", i)
  Loop
  If UBound(Exporters) > 0 Then ReDim Preserve Exporters(UBound(Exporters) - 1)
  frmSplash.lblStatus.Caption = frmSplash.lblStatus.Caption & vbCrLf & "Loaded " & UBound(Exporters) + 1 & " Exporters successfully."
  Set fMainForm = New frmMain
  Load fMainForm
  If ActionCancelled = True Then
    ActionCancelled = False
    Unload fMainForm
    Unload frmSplash
    Exit Sub
  End If
  fMainForm.Show
  Unload frmSplash
  Exit Sub
Err_Handler:
  If Err.Number = -2147024770 Then
    LogInitMsg "Exporter: " & Res & " is registered but could not be found!"
    GoTo Continue
  ElseIf Err.Number = 13 Or Err.Number = 429 Then
    LogInitMsg "Exporter: " & Res & " is corrupt or invalid!"
    GoTo Continue
  Else
    Err.Raise Err.Number
    End
  End If
End Sub

Public Sub StartMsg(Msg As String)
Dim fNum As Integer
Dim X As Long
  If Logging = 1 Then
    fNum = FreeFile
    Open LogFile For Append As #fNum
    Print #fNum, Now & vbTab; Msg
    Close #fNum
  End If
  If Len(fMainForm.txtLog.Text) + Len(Now & " - " & Msg) > 32000 Then
    fMainForm.txtLog.Text = Mid(fMainForm.txtLog.Text, InStr(Len(Msg), fMainForm.txtLog.Text, vbCrLf) + 2, Len(fMainForm.txtLog.Text))
  End If
  X = Len(fMainForm.txtLog.Text)
  fMainForm.txtLog.Text = fMainForm.txtLog.Text & vbCrLf & Now & " - " & Msg
  fMainForm.txtLog.SelStart = X + 2
  fMainForm.MousePointer = vbHourglass
  fMainForm.StatusBar1.Panels(1).Text = Msg
  fMainForm.StatusBar1.Refresh
  QryTimer = Timer
End Sub

Public Sub LogMsg(Msg As String)
Dim fNum As Integer
Dim X As Long
  If Len(fMainForm.txtLog.Text) + Len(Now & " - " & Msg) > 32000 Then
    fMainForm.txtLog.Text = Mid(fMainForm.txtLog.Text, InStr(Len(Msg), fMainForm.txtLog.Text, vbCrLf) + 2, Len(fMainForm.txtLog.Text))
  End If
  X = Len(fMainForm.txtLog.Text)
  fMainForm.txtLog.Text = fMainForm.txtLog.Text & vbCrLf & Now & " - " & Msg
  fMainForm.txtLog.SelStart = X + 2
  If Logging <> 1 Then Exit Sub
  fNum = FreeFile
  Open LogFile For Append As #fNum
  Print #fNum, Now & vbTab & Msg
  Close #fNum
End Sub

Public Sub LogInitMsg(Msg As String)
Dim fNum As Integer
  If Logging <> 1 Then Exit Sub
  fNum = FreeFile
  Open LogFile For Append As #fNum
  Print #fNum, Now & vbTab & Msg
  Close #fNum
End Sub

Public Sub LogStartup()
Dim fNum As Integer
  If Logging <> 1 Then Exit Sub
  fNum = FreeFile
  Open LogFile For Append As #fNum
  Print #fNum, vbCrLf & "######################################################################"
  If DEVELOPMENT Then
    Print #fNum, "* pgAdmin v" & app.Major & "." & app.Minor & "." & app.Revision & " DEV: Startup - " & Format(Now, "yyyy-MM-dd hh:mm:ss")
  Else
    Print #fNum, "* pgAdmin v" & app.Major & "." & app.Minor & "." & app.Revision & ": Startup - " & Format(Now, "yyyy-MM-dd hh:mm:ss")
  End If
  Print #fNum, "######################################################################" & vbCrLf
  Close #fNum
End Sub

Public Sub EndMsg()
Dim fNum As Integer
Dim Msg As String
Dim X As Long
  Msg = "Done - " & Fix((Timer - QryTimer) * 100) / 100 & " Secs."
  If Mid(fMainForm.StatusBar1.Panels(1).Text, Len(fMainForm.StatusBar1.Panels(1).Text) - 4, 5) <> "Done." Then
    If Logging = 1 Then
      fNum = FreeFile
      Open LogFile For Append As #fNum
      Print #fNum, Now & vbTab & "Done - " & Fix((Timer - QryTimer) * 100) / 100 & " Secs."
      Close #fNum
    End If
    If Len(fMainForm.txtLog.Text) + Len(Now & " - " & Msg) > 32000 Then
      fMainForm.txtLog.Text = Mid(fMainForm.txtLog.Text, InStr(Len(Msg), fMainForm.txtLog.Text, vbCrLf) + 2, Len(fMainForm.txtLog.Text))
    End If
    X = Len(fMainForm.txtLog.Text)
    fMainForm.txtLog.Text = fMainForm.txtLog.Text & vbCrLf & Now & " - " & Msg
    fMainForm.txtLog.SelStart = X + 2
    fMainForm.StatusBar1.Panels(2).Text = Fix((Timer - QryTimer) * 100) / 100 & " Secs."
    fMainForm.StatusBar1.Panels(1).Text = fMainForm.StatusBar1.Panels(1).Text & " Done."
    fMainForm.StatusBar1.Refresh
  End If
  fMainForm.MousePointer = vbDefault
End Sub

Public Function CountChar(OrigString As String, FindChar As Integer)
On Error GoTo Err_Handler
Dim X As Integer
Dim Y As Integer
  Y = 0
  For X = 1 To Len(OrigString)
    If Mid(OrigString, X, 1) = Chr(FindChar) Then Y = Y + 1
  Next
  CountChar = Y
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err, "basMisc, CountChar"
End Function

Public Sub LogQuery(Query As String)
On Error GoTo Err_Handler
  If Tracking <> True Then Exit Sub
  If Mid(UCase(Query), 1, 6) = "VACUUM" Then Exit Sub
  If Mid(UCase(Query), 1, 6) = "SELECT" Then Exit Sub
  If Mid(UCase(Query), 1, 6) = "UPDATE" Then Exit Sub
  If Mid(UCase(Query), 1, 6) = "INSERT" Then Exit Sub
  If Mid(UCase(Query), 1, 6) = "DELETE" Then Exit Sub
  Query = Replace(Query, "\", "\\")
  Query = Replace(Query, "'", "\'")
  Query = Replace(Query, QUOTE, "\" & QUOTE)
  gConnection.Execute "INSERT INTO pgadmin_rev_log (username, version, query) VALUES ('" & Username & "', '" & TrackVer & "', '" & Query & "')"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "basMisc, LogQuery"
End Sub

Public Sub LogError(Error As ErrObject, Optional SubOrFunc As String)
Dim fNum As Integer
Dim X As Variant
  fNum = FreeFile
  Open LogFile For Append As #fNum
  Print #fNum, "*******************************************************"
  If DEVELOPMENT Then
    Print #fNum, "* pgAdmin v" & app.Major & "." & app.Minor & "." & app.Revision & " DEV: Error - " & Format(Now, "yyyy-MM-dd hh:mm:ss")
  Else
    Print #fNum, "* pgAdmin v" & app.Major & "." & app.Minor & "." & app.Revision & ": Error - " & Format(Now, "yyyy-MM-dd hh:mm:ss")
  End If
  Print #fNum, "*******************************************************"
  Print #fNum, ""
  Print #fNum, "Error Details"
  Print #fNum, "*************"
  Print #fNum, "Error Number: " & Error.Number
  Print #fNum, "Error Description: " & Error.Description
  Print #fNum, "Error Source: " & Error.Source
  Print #fNum, "Subroutine or Function: " & SubOrFunc
  Print #fNum, ""
  Print #fNum, "System Details"
  Print #fNum, "**************"
  Print #fNum, "Operating System: " & WinName & " v" & WinVer & " Build " & WinBuild
  Print #fNum, "Additional Info: " & WinInfo
  Print #fNum, ""
  Print #fNum, "Environment Details"
  Print #fNum, "*******************"
  Print #fNum, "Application Path: " & app.Path
  Print #fNum, "Datasource: " & Datasource
  Print #fNum, "Tracking: " & Tracking
  Print #fNum, "TrackVer: " & TrackVer
  If MaskPassword = 0 Then
    Print #fNum, "Connect: " & gConnection.ConnectionString
  Else
    Print #fNum, "Connect: " & Replace(gConnection.ConnectionString, "PWD=" & Password, "PWD=******")
  End If
  Print #fNum, "MDAC Version: " & gConnection.Version
  If gConnection.State = adStateOpen Then
    Print #fNum, "DBMS Version: " & gConnection.Properties("DBMS VERSION")
  End If
  Print #fNum, ""
  Close #fNum
  MsgBox "An error has occured and has been logged to " & LogFile & vbCrLf & vbCrLf & _
         "Error: " & Error.Number & vbCrLf & vbCrLf & Error.Description & vbCrLf & vbCrLf & "Function: " & SubOrFunc, vbExclamation, "Error"
End Sub

Public Function StartURL(URL As String) As Long
On Error GoTo Err_Handler
Dim Scr_hDC As Long
  Scr_hDC = GetDesktopWindow()
  StartURL = ShellExecute(Scr_hDC, "Open", URL, "", "C:\", SW_SHOWNORMAL)
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err, "basMisc, StartURL"
End Function

Public Function MoveRS(rs As Recordset, Records As Long) As Long
On Error GoTo Err_Handler
Dim X As Long
  If Records < 1 Then Exit Function
  If rs Is Nothing Then Exit Function
  For X = 1 To Records
    If rs.EOF <> True Then
      rs.MoveNext
    Else
      Exit For
    End If
  Next
  MoveRS = X
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err, "basMisc, MoveRS"
End Function

Public Function SuperUser() As Boolean
On Error GoTo Err_Handler
Dim rs As New Recordset
Static bNotFirstTime As Boolean
Static bSuperuser As Boolean
  If Not bNotFirstTime Then
    LogMsg "Executing: SELECT usesuper FROM pg_user WHERE usename = '" & Username & "'"
    rs.Open "SELECT usesuper FROM pg_user WHERE usename = '" & Username & "'", gConnection, adOpenForwardOnly
    If rs!usesuper = "1" Or rs!usesuper = True Then
      bSuperuser = True
      SuperUser = True
    Else
      bSuperuser = False
      SuperUser = False
    End If
  Else
    SuperUser = bSuperuser
  End If
  bNotFirstTime = True
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err, "basMisc, SuperUser"
End Function

Public Function ObjectExists(szName As String, otObject As ObjectTypes) As Long
On Error GoTo Err_Handler
Dim rs As New Recordset
  ObjectExists = 0
  Select Case otObject
    Case tTable
      LogMsg "Executing: SELECT c.oid FROM pg_class c, pg_tables t WHERE c.relname = t.tablename AND relname = '" & szName & "'"
      rs.Open "SELECT c.oid FROM pg_class c, pg_tables t WHERE c.relname = t.tablename AND relname = '" & szName & "'", gConnection
      If Not rs.EOF Then ObjectExists = rs!OID
    Case tIndex
      LogMsg "Executing: SELECT oid FROM pg_class WHERE relkind = 'i' AND relname = '" & szName & "'"
      rs.Open "SELECT oid FROM pg_class WHERE relkind = 'i' AND relname = '" & szName & "'", gConnection
      If Not rs.EOF Then ObjectExists = rs!OID
    Case tSequence
      LogMsg "Executing: SELECT oid FROM pg_class WHERE relkind = 'S' AND relname = '" & szName & "'"
      rs.Open "SELECT oid FROM pg_class WHERE relkind = 'S' AND relname = '" & szName & "'", gConnection
      If Not rs.EOF Then ObjectExists = rs!OID
    Case tFunction
      LogMsg "Executing: SELECT oid FROM pg_proc WHERE proname = '" & szName & "'"
      rs.Open "SELECT oid FROM pg_proc WHERE proname = '" & szName & "'", gConnection
      If Not rs.EOF Then ObjectExists = rs!OID
    Case tTrigger
      LogMsg "Executing: SELECT oid FROM pg_trigger WHERE tgname = '" & szName & "'"
      rs.Open "SELECT oid FROM pg_trigger WHERE tgname = '" & szName & "'", gConnection
      If Not rs.EOF Then ObjectExists = rs!OID
    Case tView
      LogMsg "Executing: SELECT c.oid FROM pg_class c, pg_views v WHERE c.relname = v.viewname AND relname = '" & szName & "'"
      rs.Open "SELECT c.oid FROM pg_class c, pg_views v WHERE c.relname = v.viewname AND relname = '" & szName & "'", gConnection
      If Not rs.EOF Then ObjectExists = rs!OID
    Case tLanguage
      LogMsg "Executing: SELECT oid FROM pg_language WHERE lanname = '" & szName & "'"
      rs.Open "SELECT oid FROM pg_language WHERE lanname = '" & szName & "'", gConnection
      If Not rs.EOF Then ObjectExists = rs!OID
  End Select
  If rs.State <> adStateClosed Then rs.Close
  Exit Function
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  If Err.Number <> 0 Then LogError Err, "basMisc, ObjectExists"
End Function

Public Sub Chk_DriverOptions()
On Error GoTo Err_Handler
  If InStr(1, gConnection.ConnectionString, "READONLY=0") = 0 Then MsgBox "This datasource is currently Read Only. Any attempts to modify the database will fail.", vbExclamation, "Warning"
  If InStr(1, gConnection.ConnectionString, "PROTOCOL=6.4") = 0 Then MsgBox "This datasource is not configured to use the PostgreSQL v6.4 communications protocol. Performance and functionality may be impaired.", vbExclamation, "Warning"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "basMisc, Chk_DriverOptions"
End Sub

Public Function Chk_dbVersion() As Integer
On Error Resume Next
Dim rs As New Recordset
  Chk_dbVersion = 0
  LogMsg "Executing: SELECT version()"
  rs.Open "SELECT version()", gConnection, adOpenForwardOnly
  LogMsg "Database: " & rs!Version
  If Val(Mid(rs!Version, 11, 14)) < MIN_PGSQL_VERSION Then
    Chk_dbVersion = 1
  Else
    Chk_dbVersion = 0
  End If
  Set rs = Nothing
End Function

Public Function DSN_Exists(szName As String) As Boolean
On Error Resume Next
Dim i As Integer
Dim sDSNItem As String * 1024
Dim sDRVItem As String * 1024
Dim sDSN As String
Dim sDRV As String
Dim iDSNLen As Integer
Dim iDRVLen As Integer
Dim lHenv As Long

  DSN_Exists = False
  If SQLAllocEnv(lHenv) <> -1 Then
    Do Until i <> SQL_SUCCESS
      sDSNItem = Space(1024)
      sDRVItem = Space(1024)
      i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
      sDSN = VBA.Left(sDSNItem, iDSNLen)
      sDRV = VBA.Left(sDRVItem, iDRVLen)
      If sDSN = szName Then DSN_Exists = True
    Loop
  End If
End Function

Public Function DSNServer() As String
On Error GoTo Err_Handler
Dim X As Integer
  X = InStr(1, gConnection.ConnectionString, "SERVER=")
  If X <> 0 Then DSNServer = Mid(gConnection.ConnectionString, X + 7, InStr(X + 7, gConnection.ConnectionString, ";") - (X + 7))
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err, "basMisc, DSNServer"
End Function

Public Function DSNPort() As String
On Error GoTo Err_Handler
Dim X As Integer
  X = InStr(1, gConnection.ConnectionString, "PORT=")
  If X <> 0 Then DSNPort = Mid(gConnection.ConnectionString, X + 5, InStr(X + 5, gConnection.ConnectionString, ";") - (X + 5))
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err, "basMisc, DSNPort"
End Function

Public Function dbSZ(szData As String) As String
  szData = Replace(szData, "\", "\\")
  szData = Replace(szData, "'", "\'")
  dbSZ = szData
End Function

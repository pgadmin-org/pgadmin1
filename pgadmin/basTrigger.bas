Attribute VB_Name = "basTrigger"
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
Option Compare Text

'****
'**** Triggers
'****
'****

Function cmp_Trigger_CreateSQL(ByVal szTrigger_name As String, ByVal szTrigger_table As String, ByVal szTrigger_function As String, ByVal szTrigger_arguments As String, ByVal szTrigger_foreach As String, ByVal szTrigger_Executes As String, ByVal szTrigger_event As String, Optional iTrigger_type As Integer) As String
' Two syntaxes
' cmp_Trigger_Create (szTrigger_name, szTrigger_table, szTrigger_function, szTrigger_ForEach, szTrigger_Executes, szTrigger_Event )
' cmp_Trigger_Create (szTrigger_name, szTrigger_table, szTrigger_function, "", "", "", szTrigger_type)

On Error GoTo Err_Handler
    Dim szQueryStr As String

    ' if trigger_type defined
    If Not (IsMissing(iTrigger_type)) Then
        If iTrigger_type <> 0 Then
            ' retrieve values from trigger
            
            If (iTrigger_type And 1) = 1 Then
              szTrigger_foreach = " Row"
            Else
              szTrigger_foreach = " Statement"
            End If
            
            If (iTrigger_type And 2) = 2 Then
              szTrigger_Executes = " Before"
            Else
              szTrigger_Executes = " After"
            End If
            
            If (iTrigger_type And 4) = 4 Then szTrigger_event = szTrigger_event & "Insert OR "
            If (iTrigger_type And 8) = 8 Then szTrigger_event = szTrigger_event & "Delete OR "
            If (iTrigger_type And 16) = 16 Then szTrigger_event = szTrigger_event & "Update OR "
            szTrigger_event = Left(szTrigger_event, Len(szTrigger_event) - 3)
        End If
    End If
     
    szQueryStr = "CREATE TRIGGER " & QUOTE & szTrigger_name & QUOTE & vbCrLf
    szQueryStr = szQueryStr & " " & szTrigger_Executes & " " & szTrigger_event & vbCrLf
    szQueryStr = szQueryStr & " ON " & QUOTE & szTrigger_table & QUOTE & " FOR EACH " & szTrigger_foreach & vbCrLf
    szQueryStr = szQueryStr & " EXECUTE PROCEDURE " & szTrigger_function & "(" & szTrigger_arguments & ")"
    
    cmp_Trigger_CreateSQL = szQueryStr
    Exit Function
    
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Trigger_CreateSQL"
End Function

Sub cmp_Trigger_Create(ByVal szTrigger_name As String, ByVal szTrigger_table As String, ByVal szTrigger_function As String, ByVal szTrigger_arguments As String, ByVal szTrigger_foreach As String, ByVal szTrigger_Executes As String, ByVal szTrigger_event As String, Optional iTrigger_type As Integer)
' Two syntaxes
' cmp_Trigger_Create (szTrigger_name, szTrigger_table, szTrigger_function, szTrigger_ForEach, szTrigger_Executes, szTrigger_Event )
' cmp_Trigger_Create (szTrigger_name, szTrigger_table, szTrigger_function, "", "", "", szTrigger_type)
    Dim szQueryStr As String
    
    If (IsMissing(iTrigger_type)) Then
      szQueryStr = cmp_Trigger_CreateSQL(szTrigger_name, szTrigger_table, szTrigger_function, szTrigger_arguments, szTrigger_foreach, szTrigger_Executes, szTrigger_event)
    Else
      szQueryStr = cmp_Trigger_CreateSQL(szTrigger_name, szTrigger_table, szTrigger_function, szTrigger_arguments, szTrigger_foreach, szTrigger_Executes, szTrigger_event, iTrigger_type)
    End If
    
    ' Log information
    LogMsg "Creating trigger " & szTrigger_name & "..."
    LogMsg "Executing: " & szQueryStr
      
    ' Execute drop query and close log
    szQueryStr = Replace(szQueryStr, vbCrLf, "")
    gConnection.Execute szQueryStr
    LogQuery szQueryStr
      
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Trigger_SQL"
  If Err.Number = -2147467259 Then MsgBox "Trigger " & szTrigger_name & " could not be compiled." & vbCrLf & "Check source code and compile again."
  bContinueCompilation = False
End Sub

Sub cmp_Trigger_DropIfExists(ByVal lngTrigger_OID As Long, Optional ByVal szTrigger_name As String, Optional ByVal szTrigger_table As String)
 On Error GoTo Err_Handler
    Dim szDropStr As String
    
    ' Test existence of trigger
    If cmp_Trigger_Exists(lngTrigger_OID, szTrigger_name & "", szTrigger_table & "") Then
        ' Retrieve name and table is we only know the OID
        If lngTrigger_OID <> 0 And ((szTrigger_name = "") Or (szTrigger_table = "")) Then cmp_Trigger_GetValues lngTrigger_OID, "", szTrigger_name, szTrigger_table
        
        ' Create drop query
        szDropStr = "DROP TRIGGER " & QUOTE & szTrigger_name & QUOTE & " ON " & szTrigger_table
               
        ' Log information
        LogMsg "Dropping trigger " & szTrigger_name & " on table " & szTrigger_table & "..."
        LogMsg "Executing: " & szDropStr
        
        ' Execute drop query and close log
        gConnection.Execute szDropStr
        LogQuery szDropStr
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Trigger_DropIfExists"
End Sub

Sub cmp_Trigger_GetValues(lngTrigger_OID As Long, Optional szTrigger_PostgreSQLtable As String, Optional szTrigger_name As String, Optional szTrigger_table As String, Optional szTrigger_function As String, Optional szTrigger_arguments As String, Optional szTrigger_foreach As String, Optional szTrigger_Executes As String, Optional szTrigger_event As String, Optional szTrigger_Comments As String)
 ' On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    Dim iTrigger_type As Integer
    
    ' Where should we get the values ?
    If IsMissing(szTrigger_PostgreSQLtable) Or (szTrigger_PostgreSQLtable = "") Then
        szTrigger_PostgreSQLtable = "pgadmin_triggers"
    End If
        
    ' Select query
    If lngTrigger_OID <> 0 Then
        szQueryStr = "SELECT * from " & szTrigger_PostgreSQLtable
        szQueryStr = szQueryStr & " WHERE trigger_OID = " & lngTrigger_OID
        LogMsg "Retrieving name and table from trigger OID =" & lngTrigger_OID & "..."
    Else
        If IsMissing(szTrigger_name) Then szTrigger_name = ""
        szQueryStr = "SELECT * from " & szTrigger_PostgreSQLtable & " WHERE "
        szQueryStr = szQueryStr & " trigger_name = '" & szTrigger_name & "' "
        If Not (IsMissing(szTrigger_table)) And szTrigger_table <> "" Then
            szQueryStr = szQueryStr & " AND trigger_table = '" & szTrigger_table & "'"
        End If
    End If
    
    LogMsg "Executing: " & szQueryStr
    
    ' open
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    
    If Not rsComp.EOF Then
        lngTrigger_OID = rsComp!trigger_oid
        If Not (IsMissing(szTrigger_name)) Then szTrigger_name = rsComp!trigger_name & ""
        If Not (IsMissing(szTrigger_table)) Then szTrigger_table = rsComp!trigger_table & ""
        If Not (IsMissing(szTrigger_function)) Then szTrigger_function = rsComp!trigger_function & ""
        If Not (IsMissing(szTrigger_arguments)) Then szTrigger_arguments = rsComp!trigger_arguments & ""
        iTrigger_type = rsComp!trigger_type
        If iTrigger_type <> 0 Then
            If Not (IsMissing(szTrigger_foreach)) Then
                If (iTrigger_type And 1) = 1 Then
                  szTrigger_foreach = "Row"
                Else
                  szTrigger_foreach = "Statement"
                End If
            End If
            
            If Not (IsMissing(szTrigger_Executes)) Then
                If (iTrigger_type And 2) = 2 Then
                  szTrigger_Executes = "Before"
                Else
                  szTrigger_Executes = "After"
                End If
            End If
            
            If Not (IsMissing(szTrigger_event)) Then
                If (iTrigger_type And 4) = 4 Then szTrigger_event = szTrigger_event & "Insert OR "
                If (iTrigger_type And 8) = 8 Then szTrigger_event = szTrigger_event & "Delete OR "
                If (iTrigger_type And 16) = 16 Then szTrigger_event = szTrigger_event & "Update OR "
                szTrigger_event = Left(szTrigger_event, Len(szTrigger_event) - 3)
            End If
        End If
        rsComp.Close
    Else
        lngTrigger_OID = 0
        If Not (IsMissing(szTrigger_name)) Then szTrigger_name = ""
        If Not (IsMissing(szTrigger_table)) Then szTrigger_table = ""
        If Not (IsMissing(szTrigger_function)) Then szTrigger_function = ""
        If Not (IsMissing(szTrigger_arguments)) Then szTrigger_arguments = ""
        If Not (IsMissing(szTrigger_foreach)) Then szTrigger_foreach = ""
        If Not (IsMissing(szTrigger_Executes)) Then szTrigger_Executes = ""
        If Not (IsMissing(szTrigger_event)) Then szTrigger_event = ""
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Trigger_GetValues"
End Sub

Function cmp_Trigger_Exists(ByVal lngTrigger_OID As Long, Optional ByVal szTrigger_name As String, Optional ByVal szTrigger_table As String) As Boolean
 On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
  
    cmp_Trigger_Exists = False
    
    If lngTrigger_OID <> 0 Then
        szQueryStr = "SELECT * FROM pgadmin_triggers"
        szQueryStr = szQueryStr & " WHERE Trigger_OID = " & lngTrigger_OID
        
        ' Logging
        LogMsg "Testing existence of trigger OID = " & lngTrigger_OID & "..."
    Else
        If szTrigger_table <> "" And szTrigger_name <> "" Then
            szQueryStr = "SELECT * FROM pgadmin_triggers"
            szQueryStr = szQueryStr & " WHERE Trigger_name = '" & szTrigger_name & "'"
            szQueryStr = szQueryStr & " AND Trigger_table = '" & szTrigger_table & "'"
            
            ' Logging
            LogMsg "Testing existence of trigger " & szTrigger_name & " on table " & szTrigger_table & "..."
        Else
            Exit Function
        End If
    End If
    
      ' retrieve name and arguments of function to drop
    LogMsg "Executing: " & szQueryStr

    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    
    If Not rsComp.EOF Then
        cmp_Trigger_Exists = True
        rsComp.Close
    End If
    
  Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Trigger_DropIfExists"
End Function

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

Function cmp_Trigger_CreateSQL(szTrigger_name As String, szTrigger_table As String, szTrigger_function As String, szTrigger_arguments As String, szTrigger_foreach As String, szTrigger_executes As String, szTrigger_event As String, Optional iTrigger_type As Integer) As String
' Two syntaxes
' cmp_Trigger_Create (szTrigger_name, szTrigger_table, szTrigger_function, szTrigger_ForEach, szTrigger_Executes, szTrigger_Event )
' cmp_Trigger_Create (szTrigger_name, szTrigger_table, szTrigger_function, "", "", "", szTrigger_type)

On Error GoTo Err_Handler
    Dim szQueryStr As String

    ' if trigger_type defined
    If Not (IsMissing(iTrigger_type)) And iTrigger_type <> 0 Then cmp_Trigger_Ctype iTrigger_type, szTrigger_foreach, szTrigger_executes, szTrigger_event
    szTrigger_arguments = Replace(szTrigger_arguments, "'", "''")
      
    szQueryStr = "CREATE TRIGGER " & QUOTE & szTrigger_name & QUOTE
    szQueryStr = szQueryStr & " " & szTrigger_executes & " " & szTrigger_event
    szQueryStr = szQueryStr & " ON " & QUOTE & szTrigger_table & QUOTE & " FOR EACH " & szTrigger_foreach
    szQueryStr = szQueryStr & " EXECUTE PROCEDURE " & szTrigger_function & "(" & szTrigger_arguments & ")"
    
    cmp_Trigger_CreateSQL = szQueryStr
    Exit Function
    
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_CreateSQL"
End Function

Sub cmp_Trigger_Create(szTrigger_PostgreSqlTable As String, ByVal szTrigger_name As String, ByVal szTrigger_table As String, ByVal szTrigger_function As String, ByVal szTrigger_arguments As String, ByVal szTrigger_foreach As String, ByVal szTrigger_executes As String, ByVal szTrigger_event As String, Optional iTrigger_type As Integer)
On Error GoTo Err_Handler
    
    Dim szQueryStr As String
    Dim szTrigger_oid As Long
    Dim szTrigger_Query_oid As Variant
    
    ' Where should we get the values ?
    If (szTrigger_PostgreSqlTable = "") Then szTrigger_PostgreSqlTable = "pgadmin_triggers"
    
    If (IsMissing(iTrigger_type)) Then
      szQueryStr = cmp_Trigger_CreateSQL(szTrigger_name, szTrigger_table, szTrigger_function, szTrigger_arguments, szTrigger_foreach, szTrigger_executes, szTrigger_event)
    Else
      szQueryStr = cmp_Trigger_CreateSQL(szTrigger_name, szTrigger_table, szTrigger_function, szTrigger_arguments, szTrigger_foreach, szTrigger_executes, szTrigger_event, iTrigger_type)
    End If
   
    If (szTrigger_PostgreSqlTable <> "pgadmin_triggers") Then
        szTrigger_arguments = Replace(szTrigger_arguments, "'", "''")
        szTrigger_arguments = Replace(szTrigger_arguments, vbCrLf, "\n")
        
        szQueryStr = "INSERT INTO " & szTrigger_PostgreSqlTable & " (Trigger_name, Trigger_table, Trigger_function, Trigger_arguments, Trigger_foreach, Trigger_executes, Trigger_event)"
        szQueryStr = szQueryStr & " VALUES ("
        szQueryStr = szQueryStr & "'" & szTrigger_name & "', "
        szQueryStr = szQueryStr & "'" & szTrigger_table & "', "
        szQueryStr = szQueryStr & "'" & szTrigger_function & "', "
        szQueryStr = szQueryStr & "'" & szTrigger_arguments & "', "
        szQueryStr = szQueryStr & "'" & szTrigger_foreach & "', "
        szQueryStr = szQueryStr & "'" & szTrigger_executes & "', "
        szQueryStr = szQueryStr & "'" & szTrigger_event & "' "
        szQueryStr = szQueryStr & ");"
    End If
    
    ' Log information
    LogMsg "Creating trigger " & szTrigger_name & " on " & szTrigger_table & " in " & szTrigger_PostgreSqlTable & "..."
    LogMsg "Executing: " & szQueryStr
      
    ' Execute drop query and close log
    szQueryStr = Replace(szQueryStr, vbCrLf, " ")
    gConnection.Execute szQueryStr
    LogQuery szQueryStr
      
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_SQL"
  If Err.Number = -2147467259 Then MsgBox "Trigger " & szTrigger_name & " could not be compiled." & vbCrLf & "Check source code and compile again."
  bContinueRebuilding = False
End Sub

Sub cmp_Trigger_DropIfExists(szTrigger_PostgreSqlTable As String, ByVal lngTrigger_oid As Long, Optional ByVal szTrigger_name As String, Optional ByVal szTrigger_table As String)
 On Error GoTo Err_Handler
    Dim szDropStr As String
    
    ' Where should we get the values ?
    If (szTrigger_PostgreSqlTable = "") Then szTrigger_PostgreSqlTable = "pgadmin_triggers"
    
    ' Test existence of trigger
    If cmp_Trigger_Exists(szTrigger_PostgreSqlTable, lngTrigger_oid, szTrigger_name & "", szTrigger_table & "") Then
        ' Retrieve name and table is we only know the OID
        If lngTrigger_oid <> 0 And ((szTrigger_name = "") Or (szTrigger_table = "")) Then cmp_Trigger_GetValues szTrigger_PostgreSqlTable, lngTrigger_oid, szTrigger_name, szTrigger_table
        
        ' Create drop query
        If (szTrigger_PostgreSqlTable = "pgadmin_triggers") Then
            szDropStr = "DROP TRIGGER " & QUOTE & szTrigger_name & QUOTE & " ON " & szTrigger_table
        Else
            szDropStr = "DELETE FROM " & szTrigger_PostgreSqlTable & " WHERE "
            szDropStr = szDropStr & "trigger_name='" & szTrigger_name & "' AND trigger_table='" & szTrigger_table & "'"
        End If
        
        ' Log information
        LogMsg "Dropping trigger " & szTrigger_name & " on table " & szTrigger_table & " in " & szTrigger_PostgreSqlTable & "..."
        LogMsg "Executing: " & szDropStr
        
        ' Execute drop query and close log
        gConnection.Execute szDropStr
        If (szTrigger_PostgreSqlTable = "pgadmin_triggers") Then LogQuery szDropStr
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_DropIfExists"
End Sub

Sub cmp_Trigger_GetValues(szTrigger_PostgreSqlTable As String, lngTrigger_oid As Long, Optional szTrigger_name As String, Optional szTrigger_table As String, Optional szTrigger_function As String, Optional szTrigger_arguments As String, Optional szTrigger_foreach As String, Optional szTrigger_executes As String, Optional szTrigger_event As String, Optional szTrigger_Comments As String)
On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    Dim iTrigger_type As Integer
    
    ' Where should we get the values ?
    If (szTrigger_PostgreSqlTable = "") Then szTrigger_PostgreSqlTable = "pgadmin_triggers"
      
    ' Select query
    If lngTrigger_oid <> 0 Then
        szQueryStr = "SELECT * from " & szTrigger_PostgreSqlTable
        szQueryStr = szQueryStr & " WHERE trigger_OID = " & lngTrigger_oid
    Else
        If IsMissing(szTrigger_name) Then szTrigger_name = ""
        szQueryStr = "SELECT * from " & szTrigger_PostgreSqlTable & " WHERE "
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
        If IsNull(rsComp!trigger_oid) Then
            lngTrigger_oid = 0
        Else
            lngTrigger_oid = rsComp!trigger_oid
        End If
        If Not (IsMissing(szTrigger_name)) Then szTrigger_name = rsComp!trigger_name & ""
        If Not (IsMissing(szTrigger_table)) Then szTrigger_table = rsComp!trigger_table & ""
        If Not (IsMissing(szTrigger_function)) Then szTrigger_function = rsComp!trigger_function & ""
        If Not (IsMissing(szTrigger_arguments)) Then szTrigger_arguments = rsComp!Trigger_arguments & ""
        
        If (szTrigger_PostgreSqlTable = "pgadmin_triggers") Then
            iTrigger_type = rsComp!Trigger_type
            If iTrigger_type <> 0 Then cmp_Trigger_Ctype iTrigger_type, szTrigger_foreach, szTrigger_executes, szTrigger_event
        Else
            If Not (IsMissing(szTrigger_foreach)) Then szTrigger_foreach = rsComp!Trigger_foreach & ""
            If Not (IsMissing(szTrigger_executes)) Then szTrigger_executes = rsComp!Trigger_executes & ""
            If Not (IsMissing(szTrigger_event)) Then szTrigger_event = rsComp!Trigger_event & ""
        End If
        rsComp.Close
    Else
        lngTrigger_oid = 0
        If Not (IsMissing(szTrigger_name)) Then szTrigger_name = ""
        If Not (IsMissing(szTrigger_table)) Then szTrigger_table = ""
        If Not (IsMissing(szTrigger_function)) Then szTrigger_function = ""
        If Not (IsMissing(szTrigger_arguments)) Then szTrigger_arguments = ""
        If Not (IsMissing(szTrigger_foreach)) Then szTrigger_foreach = ""
        If Not (IsMissing(szTrigger_executes)) Then szTrigger_executes = ""
        If Not (IsMissing(szTrigger_event)) Then szTrigger_event = ""
        If Not (IsMissing(szTrigger_foreach)) Then szTrigger_foreach = ""
        If Not (IsMissing(szTrigger_executes)) Then szTrigger_executes = ""
        If Not (IsMissing(szTrigger_event)) Then szTrigger_event = ""
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_GetValues"
End Sub

Function cmp_Trigger_Exists(szTrigger_PostgreSqlTable As String, ByVal lngTrigger_oid As Long, Optional ByVal szTrigger_name As String, Optional ByVal szTrigger_table As String) As Boolean
 On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
  
    ' Where should we get the values ?
    If (szTrigger_PostgreSqlTable = "") Then szTrigger_PostgreSqlTable = "pgadmin_triggers"
    cmp_Trigger_Exists = False
    
    If lngTrigger_oid <> 0 Then
        szQueryStr = "SELECT * FROM " & szTrigger_PostgreSqlTable
        szQueryStr = szQueryStr & " WHERE Trigger_OID = " & lngTrigger_oid
        
        ' Logging
        LogMsg "Testing existence of trigger OID = " & lngTrigger_oid & " in table " & szTrigger_PostgreSqlTable & "..."
    Else
        If szTrigger_table = "" Or szTrigger_name = "" Then Exit Function
        szQueryStr = "SELECT * FROM " & szTrigger_PostgreSqlTable
        szQueryStr = szQueryStr & " WHERE Trigger_name = '" & szTrigger_name & "'"
        szQueryStr = szQueryStr & " AND Trigger_table = '" & szTrigger_table & "'"
        
        ' Logging
        LogMsg "Testing existence of trigger " & szTrigger_name & " on table " & szTrigger_table & " in " & szTrigger_PostgreSqlTable & "..."
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
  If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_DropIfExists"
End Function

Public Sub cmp_Trigger_CopyToDev()
On Error GoTo Err_Handler
    Dim szQuery As String
    Dim rsComp As New Recordset
    Dim iUbound As Long
    Dim iLoop As Long
    
    Dim szTrigger() As Variant
    Dim szTrigger_name As String
    Dim szTrigger_table As String
    Dim szTrigger_function As String
    Dim szTrigger_arguments As String
    Dim szTrigger_foreach As String
    Dim szTrigger_executes As String
    Dim szTrigger_event As String
    Dim iTrigger_type As Integer
    
    szQuery = "TRUNCATE TABLE pgadmin_dev_triggers;" & _
    "  INSERT INTO pgadmin_dev_triggers SELECT * " & _
    "  FROM pgadmin_triggers " & _
    "  WHERE trigger_oid > " & LAST_SYSTEM_OID & _
    "  AND trigger_name NOT LIKE 'pgadmin_%' " & _
    "  AND trigger_name NOT LIKE 'pg_%' " & _
    "  AND trigger_name NOT LIKE 'RI_%' " & _
    "  ORDER BY trigger_name; " & _
    "  UPDATE pgadmin_dev_triggers SET trigger_iscompiled = 'f';"
    
    LogMsg "Copying pgadmin_triggers into pgadmin_dev_triggers..."
    LogMsg "Executing: " & szQuery
    gConnection.Execute szQuery
    
        ' initialize pgadmin_dev_view
    szQuery = "SELECT Trigger_name, Trigger_table, Trigger_function, Trigger_arguments, Trigger_type"
    szQuery = szQuery & " FROM pgadmin_dev_triggers ORDER BY trigger_oid"
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQuery, gConnection, adOpenDynamic
    
    If Not (rsComp.EOF) Then
        szQuery = ""
        szTrigger = rsComp.GetRows
        If rsComp.State <> adStateClosed Then rsComp.Close
        iUbound = UBound(szTrigger, 2)
            For iLoop = 0 To iUbound
                'Get values
                szTrigger_name = szTrigger(0, iLoop)
                szTrigger_table = szTrigger(1, iLoop)
                szTrigger_function = szTrigger(2, iLoop)
                szTrigger_arguments = szTrigger(3, iLoop)
                iTrigger_type = szTrigger(4, iLoop)
                cmp_Trigger_Ctype iTrigger_type, szTrigger_foreach, szTrigger_executes, szTrigger_event
                
                ' Update definition of view
                szQuery = szQuery & " UPDATE pgadmin_dev_triggers SET"
                szQuery = szQuery & " Trigger_foreach = '" & szTrigger_foreach & "', "
                szQuery = szQuery & " Trigger_executes = '" & szTrigger_executes & "', "
                szQuery = szQuery & " Trigger_event = '" & szTrigger_event & "' "
                szQuery = szQuery & " WHERE "
                szQuery = szQuery & " Trigger_name = '" & szTrigger_name & "'"
                szQuery = szQuery & " AND Trigger_table = '" & szTrigger_table & "'; "
            Next iLoop
            LogMsg "Executing: " & szQuery
            gConnection.Execute szQuery
    End If
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_CopyToDev"
End Sub

Sub cmp_Trigger_Ctype(iTrigger_type As Integer, szTrigger_foreach As String, szTrigger_executes As String, szTrigger_event As String)
On Error GoTo Err_Handler
        If iTrigger_type <> 0 Then
            ' retrieve values from trigger
            
            If (iTrigger_type And 1) = 1 Then
              szTrigger_foreach = "Row"
            Else
              szTrigger_foreach = "Statement"
            End If
            
            If (iTrigger_type And 2) = 2 Then
              szTrigger_executes = "Before"
            Else
              szTrigger_executes = "After"
            End If
            
            If (iTrigger_type And 4) = 4 Then szTrigger_event = szTrigger_event & "Insert OR "
            If (iTrigger_type And 8) = 8 Then szTrigger_event = szTrigger_event & "Delete OR "
            If (iTrigger_type And 16) = 16 Then szTrigger_event = szTrigger_event & "Update OR "
            szTrigger_event = Trim(Left(szTrigger_event, Len(szTrigger_event) - 3))
        Else
            szTrigger_foreach = ""
            szTrigger_executes = ""
            szTrigger_event = ""
        End If
        
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_Ctype"
End Sub

Public Sub cmp_Trigger_ParseName(szInput As String, szTrigger_name As String, szTrigger_table As String)
On Error GoTo Err_Handler
    Dim iInstr As Integer
    iInstr = InStr(szInput, "ON")
    If iInstr > 0 Then
        szTrigger_name = Left(szInput, iInstr - 2)
        szTrigger_table = Mid(szInput, iInstr + 3, Len(szInput) - iInstr - 2)
    Else
        szTrigger_name = szInput
        szTrigger_table = ""
    End If
    
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_ParseName"
End Sub

Public Sub cmp_Trigger_DropAll(Optional szTrigger_PostgreSqlTable As String)
On Error GoTo Err_Handler
    Dim szQuery As String
    Dim szTrigger() As Variant
    Dim iLoop As Long
    Dim iUbound As Long
    Dim rsTrigger As New Recordset
    Dim szTrigger_name As String
    Dim szTrigger_table As String
    
    If IsMissing(szTrigger_PostgreSqlTable) Or (szTrigger_PostgreSqlTable = "") Then szTrigger_PostgreSqlTable = "pgadmin_triggers"
        
    If (szTrigger_PostgreSqlTable = "pgadmin_triggers") Then
        szQuery = "SELECT trigger_name, trigger_table FROM pgadmin_triggers " & _
        "  WHERE trigger_oid > " & LAST_SYSTEM_OID & _
        "  AND trigger_name NOT LIKE 'pgadmin_%' " & _
        "  AND trigger_name NOT LIKE 'pg_%' " & _
        "  AND trigger_name NOT LIKE 'RI_%' " & _
        "  ORDER BY trigger_name; "
    
        LogMsg "Dropping all triggers in pgadmin_triggers..."
        LogMsg "Executing: " & szQuery
        
        If rsTrigger.State <> adStateClosed Then rsTrigger.Close
        rsTrigger.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
    
        If Not (rsTrigger.EOF) Then
            szTrigger = rsTrigger.GetRows
            rsTrigger.Close
            iUbound = UBound(szTrigger, 2)
                For iLoop = 0 To iUbound
                     szTrigger_name = szTrigger(0, iLoop)
                     szTrigger_table = szTrigger(1, iLoop)
                     cmp_Trigger_DropIfExists "", 0, szTrigger_name, szTrigger_table
                Next iLoop
            Erase szTrigger
        End If
    Else
        szQuery = "TRUNCATE " & szTrigger_PostgreSqlTable
        LogMsg "Truncating " & szTrigger_PostgreSqlTable & "..."
        LogMsg "Executing: " & szQuery
        gConnection.Execute szQuery
    End If
   
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basProject, cmp_Trigger_DropAll"
End Sub

Attribute VB_Name = "basTrigger"
' pgadmin - PostgreSQL db Administration/Management for Win32
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

Private szPro_Text As String
Private szDev_Text As String
Private szSys_Text As String

Private iPro_Index As Long
Private iDev_Index As Long
Private iSys_Index As Long

Private iPro_Count As Long
Private iDev_Count As Long
Private iSys_Count As Long

'****
'**** Triggers
'****

Function cmp_Trigger_CreateSQL(szTrigger_Name As String, szTrigger_Table As String, szTrigger_Function As String, szTrigger_Arguments As String, szTrigger_Foreach As String, szTrigger_Executes As String, szTrigger_Event As String) As String
On Error GoTo Err_Handler
    Dim szQuery As String
    
    cmp_Trigger_CreateSQL = ""
    
    If szTrigger_Name = "" Then Exit Function
    If szTrigger_Table = "" Then Exit Function
    If szTrigger_Function = "" Then Exit Function
    If szTrigger_Foreach = "" Then Exit Function
    If szTrigger_Executes = "" Then Exit Function
    If szTrigger_Event = "" Then Exit Function
    
    szQuery = "CREATE TRIGGER " & QUOTE & szTrigger_Name & QUOTE
    szQuery = szQuery & " " & szTrigger_Executes & " " & szTrigger_Event
    szQuery = szQuery & " ON " & QUOTE & szTrigger_Table & QUOTE & " FOR EACH " & szTrigger_Foreach
    szQuery = szQuery & " EXECUTE PROCEDURE " & szTrigger_Function & "(" & szTrigger_Arguments & ")"
    
    cmp_Trigger_CreateSQL = szQuery
    Exit Function
    
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_CreateSQL"
End Function

Sub cmp_Trigger_Create(szTrigger_PostgreSqlTable As String, ByVal szTrigger_Name As String, ByVal szTrigger_Table As String, ByVal szTrigger_Function As String, ByVal szTrigger_Arguments As String, ByVal szTrigger_Foreach As String, ByVal szTrigger_Executes As String, ByVal szTrigger_Event As String, ByVal szTrigger_Comments As String)
On Error GoTo Err_Handler
    Dim iTrigger_type As Integer
    
    iTrigger_type = cmp_Trigger_Ctype_ToInteger(szTrigger_Foreach, szTrigger_Executes, szTrigger_Event)
    If iTrigger_type = 0 Then Exit Sub
    
    Dim szQuery As String
    
    ' Where should we get the values ?
    If (szTrigger_PostgreSqlTable = "") Then szTrigger_PostgreSqlTable = "pgadmin_triggers"
   
    If (szTrigger_PostgreSqlTable = "pgadmin_triggers") Then
        szQuery = cmp_Trigger_CreateSQL(szTrigger_Name, szTrigger_Table, szTrigger_Function, szTrigger_Arguments, szTrigger_Foreach, szTrigger_Executes, szTrigger_Event)
    Else
        szTrigger_Arguments = Replace(szTrigger_Arguments, "'", "''")
        szTrigger_Comments = Replace(szTrigger_Comments, "'", "''")
        
        szQuery = "INSERT INTO " & szTrigger_PostgreSqlTable & " (Trigger_name, Trigger_table, Trigger_function, Trigger_arguments, Trigger_type, Trigger_comments)"
        szQuery = szQuery & " VALUES ("
        szQuery = szQuery & "'" & szTrigger_Name & "', "
        szQuery = szQuery & "'" & szTrigger_Table & "', "
        szQuery = szQuery & "'" & szTrigger_Function & "', "
        szQuery = szQuery & "'" & szTrigger_Arguments & "', "
        szQuery = szQuery & "'" & Str(iTrigger_type) & "', "
        szQuery = szQuery & "'" & szTrigger_Comments & "' "
        szQuery = szQuery & ");"
    End If
    
    ' Log information
    LogMsg "Creating trigger " & szTrigger_Name & " on " & szTrigger_Table & " in " & szTrigger_PostgreSqlTable & "..."
    LogMsg "Executing: " & szQuery
            
    'Execute
    gConnection.Execute szQuery
    If (szTrigger_PostgreSqlTable = "pgadmin_triggers") Then
        LogQuery szQuery
    
        ' Write comments
        szQuery = "COMMENT ON TRIGGER " & szTrigger_Name & " ON " & szTrigger_Table & " IS '" & szTrigger_Comments & "'"
        LogQuery szQuery
        gConnection.Execute szQuery
    End If
      
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_SQL"
If Err.Number = -2147467259 Then MsgBox "Trigger " & szTrigger_Name & " could not be compiled." & vbCrLf & "Check source code and compile again."
bContinueRebuilding = False
End Sub

Sub cmp_Trigger_DropIfExists(szTrigger_PostgreSqlTable As String, ByVal szTrigger_Name As String, ByVal szTrigger_Table As String)
On Error GoTo Err_Handler
    Dim szDropStr As String
    
    If szTrigger_Name = "" Then Exit Sub
    If szTrigger_Table = "" Then Exit Sub
    
    ' Where should we get the values ?
    If (szTrigger_PostgreSqlTable = "") Then szTrigger_PostgreSqlTable = "pgadmin_triggers"
    
    ' Test existence of trigger
    If cmp_Trigger_Exists(szTrigger_PostgreSqlTable, szTrigger_Name, szTrigger_Table) Then
        cmp_Trigger_Drop szTrigger_PostgreSqlTable, szTrigger_Name, szTrigger_Table
    End If

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_DropIfExists"
End Sub

Sub cmp_Trigger_Drop(szTrigger_PostgreSqlTable As String, ByVal szTrigger_Name As String, ByVal szTrigger_Table As String)
On Error GoTo Err_Handler
    Dim szDropStr As String
    
    If szTrigger_Name = "" Then Exit Sub
    If szTrigger_Table = "" Then Exit Sub
    
    ' Where should we get the values ?
    If (szTrigger_PostgreSqlTable = "") Then szTrigger_PostgreSqlTable = "pgadmin_triggers"
          
    ' Create drop query
    If (szTrigger_PostgreSqlTable = "pgadmin_triggers") Then
        szDropStr = "DROP TRIGGER " & QUOTE & szTrigger_Name & QUOTE & " ON " & szTrigger_Table
    Else
        szDropStr = "DELETE FROM " & szTrigger_PostgreSqlTable & " WHERE "
        szDropStr = szDropStr & "trigger_name='" & szTrigger_Name & "' AND trigger_table='" & szTrigger_Table & "'"
    End If
        
    ' Log information
    LogMsg "Dropping trigger " & szTrigger_Name & " on table " & szTrigger_Table & " in " & szTrigger_PostgreSqlTable & "..."
    LogMsg "Executing: " & szDropStr
    
    ' Execute drop query and close log
    gConnection.Execute szDropStr
    If (szTrigger_PostgreSqlTable = "pgadmin_triggers") Then LogQuery szDropStr

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_DropIfExists"
End Sub

Sub cmp_Trigger_GetValues(szTrigger_PostgreSqlTable As String, szTrigger_Name As String, szTrigger_Table As String, Optional szTrigger_Function As String, Optional szTrigger_Arguments As String, Optional szTrigger_Foreach As String, Optional szTrigger_Executes As String, Optional szTrigger_Event As String, Optional szTrigger_Comments As String)
On Error GoTo Err_Handler
    Dim szQuery As String
    Dim rsComp As New Recordset
    Dim iTrigger_type As Integer
    
    If szTrigger_Name = "" Then Exit Sub
    
    ' Where should we get the values ?
    If (szTrigger_PostgreSqlTable = "") Then szTrigger_PostgreSqlTable = "pgadmin_triggers"
      
    ' Select query
    szQuery = "SELECT * from " & szTrigger_PostgreSqlTable & " WHERE "
    szQuery = szQuery & " trigger_name = '" & szTrigger_Name & "' "
    szQuery = szQuery & " AND trigger_table = '" & szTrigger_Table & "'"

    LogMsg "Executing: " & szQuery
    
    ' open
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQuery, gConnection
    
    If Not rsComp.EOF Then
        szTrigger_Name = rsComp!trigger_name & ""
        szTrigger_Table = rsComp!trigger_table & ""
        szTrigger_Function = rsComp!trigger_function & ""
        szTrigger_Arguments = rsComp!Trigger_arguments & ""
        cmp_Trigger_Ctype_ToString rsComp!Trigger_type, szTrigger_Foreach, szTrigger_Executes, szTrigger_Event
        szTrigger_Comments = rsComp!Trigger_comments & ""
        rsComp.Close
    Else
        If Not (IsMissing(szTrigger_Name)) Then szTrigger_Name = ""
        If Not (IsMissing(szTrigger_Table)) Then szTrigger_Table = ""
        If Not (IsMissing(szTrigger_Function)) Then szTrigger_Function = ""
        If Not (IsMissing(szTrigger_Arguments)) Then szTrigger_Arguments = ""
        If Not (IsMissing(szTrigger_Foreach)) Then szTrigger_Foreach = ""
        If Not (IsMissing(szTrigger_Executes)) Then szTrigger_Executes = ""
        If Not (IsMissing(szTrigger_Event)) Then szTrigger_Event = ""
        If Not (IsMissing(szTrigger_Foreach)) Then szTrigger_Foreach = ""
        If Not (IsMissing(szTrigger_Executes)) Then szTrigger_Executes = ""
        If Not (IsMissing(szTrigger_Event)) Then szTrigger_Event = ""
        If Not (IsMissing(szTrigger_Comments)) Then szTrigger_Comments = ""
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_GetValues"
End Sub

Function cmp_Trigger_Exists(szTrigger_PostgreSqlTable As String, Optional ByVal szTrigger_Name As String, Optional ByVal szTrigger_Table As String) As Boolean
 On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
  
    If szTrigger_Table = "" Or szTrigger_Name = "" Then Exit Function
  
    ' Where should we get the values ?
    If (szTrigger_PostgreSqlTable = "") Then szTrigger_PostgreSqlTable = "pgadmin_triggers"
    cmp_Trigger_Exists = False
    
    szQueryStr = "SELECT * FROM " & szTrigger_PostgreSqlTable
    szQueryStr = szQueryStr & " WHERE Trigger_name = '" & szTrigger_Name & "'"
    szQueryStr = szQueryStr & " AND Trigger_table = '" & szTrigger_Table & "'"

    ' Logging
    LogMsg "Testing existence of trigger " & szTrigger_Name & " on table " & szTrigger_Table & " in " & szTrigger_PostgreSqlTable & "..."

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
  If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_Exists"
End Function

Sub cmp_Trigger_Ctype_ToString(iTrigger_type As Integer, szTrigger_Foreach As String, szTrigger_Executes As String, szTrigger_Event As String)
On Error GoTo Err_Handler
        szTrigger_Event = ""
        szTrigger_Foreach = ""
        szTrigger_Executes = ""
            
        If IsNull(iTrigger_type) Then iTrigger_type = 0
        
        If (iTrigger_type And 1) = 1 Then
          szTrigger_Foreach = "Row"
        Else
          szTrigger_Foreach = "Statement"
        End If
        
        If (iTrigger_type And 2) = 2 Then
          szTrigger_Executes = "Before"
        Else
          szTrigger_Executes = "After"
        End If
        
        szTrigger_Event = ""
        If (iTrigger_type And 4) = 4 Then szTrigger_Event = szTrigger_Event & "Insert OR "
        If (iTrigger_type And 8) = 8 Then szTrigger_Event = szTrigger_Event & "Delete OR "
        If (iTrigger_type And 16) = 16 Then szTrigger_Event = szTrigger_Event & "Update OR "
        
        If Len(szTrigger_Event) > 0 Then szTrigger_Event = Trim(Left(szTrigger_Event, Len(szTrigger_Event) - 3))

Err_Handler:
If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_Ctype_ToString"
End Sub

Function cmp_Trigger_Ctype_ToInteger(szTrigger_Foreach As String, szTrigger_Executes As String, szTrigger_Event As String) As Integer
On Error GoTo Err_Handler
        Dim iForEach As Integer
        Dim iExecutes As Integer
        Dim iInsert As Integer
        Dim iDelete As Integer
        Dim iUpdate As Integer
        
        iForEach = 0
        iExecutes = 0
        iInsert = 0
        iDelete = 0
        iUpdate = 0
        
        cmp_Trigger_Ctype_ToInteger = 0
        
        If szTrigger_Foreach = "" Then Exit Function
        If szTrigger_Executes = "" Then Exit Function
        If szTrigger_Event = "" Then Exit Function
        
        If InStr(szTrigger_Foreach, "Row") > 0 Then iForEach = 1
        If InStr(szTrigger_Executes, "Before") > 0 Then iExecutes = 2
        
        If InStr(szTrigger_Event, "Insert") > 0 Then iInsert = 4
        If InStr(szTrigger_Event, "Delete") > 0 Then iDelete = 8
        If InStr(szTrigger_Event, "Update") > 0 Then iUpdate = 16
        
        cmp_Trigger_Ctype_ToInteger = iForEach + iExecutes + iInsert + iDelete + iUpdate
        
Exit Function
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_Ctype_ToInteger"
End Function

Public Sub cmp_Trigger_ParseName(szInput As String, szTrigger_Name As String, szTrigger_Table As String)
On Error GoTo Err_Handler
    Dim iInstr As Integer
    iInstr = InStr(szInput, " ON ")
    If iInstr > 0 Then
        szTrigger_Name = Left(szInput, iInstr - 1)
        szTrigger_Table = Mid(szInput, iInstr + 4, Len(szInput) - iInstr - 2)
    Else
        szTrigger_Name = szInput
        szTrigger_Table = ""
    End If
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_ParseName"
End Sub

Public Sub cmp_Trigger_Move(szTrigger_source_table As String, szTrigger_target_table As String, szTrigger_Name As String, szTrigger_Table As String, Optional bPromptForReplace As Boolean)
On Error GoTo Err_Handler

Dim szTrigger_Function As String
Dim szTrigger_Arguments As String
Dim szTrigger_Foreach As String
Dim szTrigger_Executes As String
Dim szTrigger_Event As String
Dim szTrigger_Comments As String

    
    If IsMissing(bPromptForReplace) = True Then bPromptForReplace = True
    
    If szTrigger_source_table = "" Then szTrigger_source_table = "pgadmin_Triggers"
    If szTrigger_target_table = "" Then szTrigger_target_table = "pgadmin_Triggers"
    If szTrigger_source_table = szTrigger_target_table Then Exit Sub
    
    If cmp_Trigger_Exists(szTrigger_source_table, szTrigger_Name, szTrigger_Table) Then
        cmp_Trigger_GetValues szTrigger_source_table, szTrigger_Name, szTrigger_Table, szTrigger_Function, szTrigger_Arguments, szTrigger_Foreach, szTrigger_Executes, szTrigger_Event, szTrigger_Comments
        If cmp_Trigger_Exists(szTrigger_target_table, szTrigger_Name, szTrigger_Table) Then
            If (bPromptForReplace = False) Then
                cmp_Trigger_Drop szTrigger_target_table, szTrigger_Name, szTrigger_Table
                cmp_Trigger_Create szTrigger_target_table, szTrigger_Name, szTrigger_Table, szTrigger_Function, szTrigger_Arguments, szTrigger_Foreach, szTrigger_Executes, szTrigger_Event, szTrigger_Comments
            Else
               If MsgBox("Replace existing target Trigger " & vbCrLf & szTrigger_Name & " on table " & szTrigger_Table & " ?", vbYesNo) = vbYes Then
                cmp_Trigger_Drop szTrigger_target_table, szTrigger_Name, szTrigger_Table
                cmp_Trigger_Create szTrigger_target_table, szTrigger_Name, szTrigger_Table, szTrigger_Function, szTrigger_Arguments, szTrigger_Foreach, szTrigger_Executes, szTrigger_Event, szTrigger_Comments
               End If
            End If
        Else
            cmp_Trigger_Create szTrigger_target_table, szTrigger_Name, szTrigger_Table, szTrigger_Function, szTrigger_Arguments, szTrigger_Foreach, szTrigger_Executes, szTrigger_Event, szTrigger_Comments
        End If
        If bContinueRebuilding = True And szTrigger_target_table = "pgadmin_Triggers" Then
            cmp_Trigger_SetIsCompiled szTrigger_source_table, szTrigger_Name, szTrigger_Table
        End If
    End If
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_Move"
End Sub

Public Sub cmp_Trigger_SetIsCompiled(ByVal szTrigger_dev_table As String, ByVal szTrigger_Name As String, ByVal szTrigger_Table As String)
On Error GoTo Err_Handler
    Dim szQueryStr As String
    
   If szTrigger_Name & "" = "" Then Exit Sub
        
    szQueryStr = "UPDATE " & szTrigger_dev_table & " SET Trigger_iscompiled = 't'"
    szQueryStr = szQueryStr & " WHERE Trigger_name = '" & szTrigger_Name & "'"
    szQueryStr = szQueryStr & " AND Trigger_table = '" & szTrigger_Table & "'"
     
    LogMsg "Executing: " & szQueryStr
    gConnection.Execute szQueryStr
Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_SetIsCompiled"
End Sub


' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Tree
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub cmp_trigger_tree_export(Tree As TreeToy, cDialog As CommonDialog)
On Error GoTo Err_Handler

    Dim szExport As String
    Dim bExport As Boolean
    Dim szHeader As String
    
    Dim nodX As Node
    Dim szParentKey As String
    
    Dim szTrigger_pgTable As String
    Dim szTrigger_Name As String
    Dim szTrigger_Table As String
    Dim szTrigger_Function As String
    Dim szTrigger_Arguments As String
    Dim szTrigger_Foreach As String
    Dim szTrigger_Event As String
    Dim szTrigger_Executes As String
    Dim szTrigger_Comments As String
    
    bExport = False
    szExport = ""
        
    For Each nodX In Tree.Nodes
        If (nodX.Checked = True) Then
            If nodX.Parent Is Nothing Then
               szParentKey = nodX.Key
            Else
               szParentKey = nodX.Parent.Key
            End If
            
            If szParentKey = "Pro:" Or szParentKey = "Sys:" Then
                    szTrigger_pgTable = "pgadmin_triggers"
            Else
                    szTrigger_pgTable = gDevPostgresqlTables & "_Triggers"
            End If

            bExport = True
            
            cmp_Trigger_ParseName nodX.Text, szTrigger_Name, szTrigger_Table
            cmp_Trigger_GetValues szTrigger_pgTable, szTrigger_Name, szTrigger_Table, szTrigger_Function, szTrigger_Arguments, szTrigger_Foreach, szTrigger_Executes, szTrigger_Event, szTrigger_Comments
                
            If szTrigger_Name <> "" Then
                ' Header
                szExport = szExport & "/*" & vbCrLf
                szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
                szExport = szExport & "Trigger " & szTrigger_Name & " on table " & szTrigger_Table & vbCrLf
                If szTrigger_Comments <> "" Then szExport = szExport & szTrigger_Comments & vbCrLf
                szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
                szExport = szExport & "*/" & vbCrLf
                
                ' Trigger
                szExport = szExport & cmp_Trigger_CreateSQL(szTrigger_Name, szTrigger_Table, szTrigger_Function, szTrigger_Arguments, szTrigger_Foreach, szTrigger_Executes, szTrigger_Event) & vbCrLf & vbCrLf
            End If
        End If
    Next
 
    If bExport Then
        szHeader = "/*" & vbCrLf
        szHeader = szHeader & "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" & vbCrLf
        szHeader = szHeader & "The choice of a GNU generation " & vbCrLf
        szHeader = szHeader & "PostgreSQL     www.postgresql.org" & vbCrLf
        szHeader = szHeader & "pgadmin        www.greatbridge.org/project/pgadmin" & vbCrLf
        szHeader = szHeader & "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" & vbCrLf
        szHeader = szHeader & "*/" & vbCrLf & vbCrLf
        szExport = szHeader & szExport
        MsgExportToFile cDialog, szExport, "sql", "Export triggers"
    End If
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_trigger_tree_export"
End Sub

Sub cmp_Trigger_tree_copy_devtopro(Tree As TreeToy)
On Error GoTo Err_Handler

    Dim nodX As Node
    Dim szParentKey As String
    Dim bRefresh As Boolean
    
    Dim szTrigger_Name As String
    Dim szTrigger_Table As String
      
    Dim szMsgboxMessage As String
    
    If Tree.TreeCountChecked = 0 Then Exit Sub
    
    bRefresh = False
    bContinueRebuilding = True
    For Each nodX In Tree.Nodes
        If (nodX.Checked = True) Then
            If nodX.Parent Is Nothing Then
               szParentKey = nodX.Key
            Else
               szParentKey = nodX.Parent.Key
            End If

            If szParentKey = "Dev:" And bContinueRebuilding = True Then
                cmp_Trigger_ParseName nodX.Text, szTrigger_Name, szTrigger_Table
                cmp_Trigger_Move gDevPostgresqlTables & "_Triggers", "pgadmin_Triggers", szTrigger_Name, szTrigger_Table, True
                bRefresh = True
            End If
        End If
    Next
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_tree_copy_devtopro"
End Sub

Public Sub cmp_Trigger_tree_copy_protodev(Tree As TreeToy)
On Error GoTo Err_Handler
    Dim nodX As Node
    Dim szParentKey As String
    Dim bRefresh As Boolean
    
    Dim szTrigger_Name As String
    Dim szTrigger_Table As String
      
    If Tree.TreeCountChecked = 0 Then Exit Sub
    
    bRefresh = False
    For Each nodX In Tree.Nodes
        If (nodX.Checked = True) Then
            If nodX.Parent Is Nothing Then
               szParentKey = nodX.Key
            Else
               szParentKey = nodX.Parent.Key
            End If

            If szParentKey = "Pro:" Or szParentKey = "Sys:" Then
                  cmp_Trigger_ParseName nodX.Text, szTrigger_Name, szTrigger_Table
                  cmp_Trigger_Move "pgadmin_triggers", gDevPostgresqlTables & "_Triggers", szTrigger_Name, szTrigger_Table, True
                  bRefresh = True
            End If
        End If
    Next
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_tree_copy_protodev"
End Sub

Public Sub cmp_Trigger_tree_drop(Tree As TreeToy)
 On Error GoTo Err_Handler
    Dim szTrigger_pgTable As String
    Dim szTrigger_Table As String
    Dim szTrigger_Name As String
    Dim nodX As Node
    Dim bDrop As Boolean
    
    Dim szParentKey As String
    
    If Tree Is Nothing Then Exit Sub
       
    StartMsg "Dropping Trigger(s)..."
        
       For Each nodX In Tree.Nodes
             If (nodX.Checked = True) Then
                bDrop = False
                
                If (nodX.Parent Is Nothing) Then
                   szParentKey = nodX.Key
                Else
                   szParentKey = nodX.Parent.Key
                End If
    
                Select Case szParentKey
                    Case "Pro:"
                    szTrigger_pgTable = "pgadmin_Triggers"
                    bDrop = True
                
                    Case "Dev:"
                    szTrigger_pgTable = gDevPostgresqlTables & "_Triggers"
                    bDrop = True
                End Select
                     
                If bDrop = True Then
                    cmp_Trigger_ParseName nodX.Text, szTrigger_Name, szTrigger_Table
                    cmp_Trigger_DropIfExists szTrigger_pgTable, szTrigger_Name, szTrigger_Table
                End If
             End If
        Next
        Set nodX = Nothing
        
        EndMsg
    
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_tree_drop"
End Sub

Public Sub cmp_Trigger_tree_refresh(Tree As TreeToy, bShowSystem As Boolean, iPro_Index As Integer, iSys_Index As Integer, iDev_Index As Integer)
On Error GoTo Err_Handler

  Dim NodeX As Node
  Dim szQuery As String
  Dim szTrigger() As Variant
  Dim iLoop As Long
  Dim iUbound As Long
  
  Dim szTrigger_oid As String
  Dim szTrigger_Name As String
  Dim szTrigger_Table As String
  Dim szTrigger_Function As String
  Dim szTrigger_Arguments As String
  Dim szTrigger_type As String
  Dim szTrigger_Foreach As String
  Dim szTrigger_Executes As String
  Dim szTrigger_Event As String
  Dim szTrigger_Comments As String
  
  Dim szTrigger_iscompiled As Boolean
  Dim rsTrigger As New Recordset
  
  StartMsg "Retrieving Trigger Names..."
  
  Tree.Nodes.Clear
  iPro_Count = 0
  iDev_Count = 0
  iSys_Count = 0

  If DevMode = False Then
    szPro_Text = "User Triggers"
  Else
    szPro_Text = "Production Triggers"
  End If
  
  Set NodeX = Tree.Nodes.Add(, tvwChild, "Pro:", szPro_Text, 1)
  iPro_Index = NodeX.Index
  NodeX.Expanded = False
  
  szDev_Text = "Development Triggers"
  If DevMode = True Then
    Set NodeX = Tree.Nodes.Add(, tvwChild, "Dev:", szDev_Text, 1)
    iDev_Index = NodeX.Index
    NodeX.Expanded = False
  End If
  
  szSys_Text = "System Triggers"
  If bShowSystem = True Then
    Set NodeX = Tree.Nodes.Add(, tvwChild, "Sys:", szSys_Text, 1)
    iSys_Index = NodeX.Index
    NodeX.Expanded = False
  End If

 ' ---------------------------------------------------------------------
 ' Retrieve pgadmin_Triggers in one query
 ' ---------------------------------------------------------------------
  If rsTrigger.State <> adStateClosed Then rsTrigger.Close
  If bShowSystem = True Then
     szQuery = "SELECT Trigger_oid, Trigger_name, Trigger_table, Trigger_function, Trigger_arguments, Trigger_type FROM pgadmin_Triggers ORDER BY Trigger_name"
  Else
     szQuery = "SELECT Trigger_oid, Trigger_name, Trigger_table, Trigger_function, Trigger_arguments, Trigger_type FROM pgadmin_Triggers WHERE Trigger_oid > " & LAST_SYSTEM_OID & " AND Trigger_name NOT LIKE 'pgadmin_%' AND Trigger_name NOT LIKE 'pg_%' ORDER BY Trigger_name"
  End If
  LogMsg "Executing: " & szQuery
  rsTrigger.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
  
  If Not (rsTrigger.EOF) Then
    szTrigger = rsTrigger.GetRows
    iUbound = UBound(szTrigger, 2)
    For iLoop = 0 To iUbound
         szTrigger_oid = szTrigger(0, iLoop) & ""
         szTrigger_Name = szTrigger(1, iLoop) & ""
         szTrigger_Table = szTrigger(2, iLoop) & ""
         szTrigger_Function = szTrigger(3, iLoop) & ""
         szTrigger_Arguments = szTrigger(4, iLoop) & ""
         szTrigger_type = szTrigger(5, iLoop) & ""
         cmp_Trigger_Ctype_ToString CInt(szTrigger_type), szTrigger_Foreach, szTrigger_Executes, szTrigger_Event

         If CLng(szTrigger_oid) < LAST_SYSTEM_OID Or Left(szTrigger_Name, 8) = "pgadmin_" Or Left(szTrigger_Name, 3) = "pg_" Then
         ' ---------------------------------------------------------------------
         ' If it is a system Trigger, add it to "S:" System node
         ' ---------------------------------------------------------------------
            If szTrigger_Table <> "" Then
                Set NodeX = Tree.Nodes.Add("Sys:", tvwChild, "S:" & szTrigger_Name & " on " & szTrigger_Table, szTrigger_Name & " on " & szTrigger_Table, 2)
            Else
                Set NodeX = Tree.Nodes.Add("Sys:", tvwChild, "S:" & szTrigger_Name, szTrigger_Name, 2)
            End If
            iSys_Count = iSys_Count + 1
          Else
         ' ---------------------------------------------------------------------
         ' Else it is a user Trigger, add it to "P:" Production node
         ' ---------------------------------------------------------------------
            If szTrigger_Table <> "" Then
                Set NodeX = Tree.Nodes.Add("Pro:", tvwChild, "P:" & szTrigger_Name & " on " & szTrigger_Table, szTrigger_Name & " on " & szTrigger_Table, 4)
            Else
                Set NodeX = Tree.Nodes.Add("Pro:", tvwChild, "P:" & szTrigger_Name, szTrigger_Name, 4)
             End If
            iPro_Count = iPro_Count + 1
          End If
          NodeX.Tag = cmp_Trigger_CreateSQL(szTrigger_Name, szTrigger_Table, szTrigger_Function, szTrigger_Arguments, szTrigger_Foreach, szTrigger_Executes, szTrigger_Event)
          NodeX.Image = 4
    Next iLoop
  End If
  
  Tree.Nodes.Item(iPro_Index).Text = Tree.Nodes.Item(iPro_Index).Text & " (" & CStr(iPro_Count) & ")"
  If iSys_Count > 0 Then Tree.Nodes.Item(iSys_Index).Text = Tree.Nodes.Item(iSys_Index).Text & " (" & CStr(iSys_Count) & ")"

  Erase szTrigger
  
 ' ---------------------------------------------------------------------
 ' Retrieve pgadmin_dev_Triggers in one query
 ' ---------------------------------------------------------------------
 If DevMode = True Then
      If rsTrigger.State <> adStateClosed Then rsTrigger.Close
      szQuery = "SELECT Trigger_name, Trigger_table, Trigger_function, Trigger_arguments, Trigger_type, Trigger_iscompiled FROM " & gDevPostgresqlTables & "_Triggers" & " ORDER BY Trigger_name"
      LogMsg "Executing: " & szQuery
      rsTrigger.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
    
      If Not (rsTrigger.EOF) Then
        szTrigger = rsTrigger.GetRows
        iUbound = UBound(szTrigger, 2)
        iDev_Count = iUbound + 1
        For iLoop = 0 To iUbound
            szTrigger_Name = szTrigger(0, iLoop) & ""
            szTrigger_Table = szTrigger(1, iLoop) & ""
            szTrigger_Function = szTrigger(2, iLoop) & ""
            szTrigger_Arguments = szTrigger(3, iLoop) & ""
            szTrigger_type = szTrigger(4, iLoop) & ""
            
            If IsNull(szTrigger(5, iLoop)) Then
                szTrigger_iscompiled = False
             Else
                szTrigger_iscompiled = szTrigger(5, iLoop)
             End If
            
            If szTrigger_Table <> "" Then
                Set NodeX = Tree.Nodes.Add("Dev:", tvwChild, "D:" & szTrigger_Name & " on " & szTrigger_Table, szTrigger_Name & " on " & szTrigger_Table, 2)
             Else
                Set NodeX = Tree.Nodes.Add("Dev:", tvwChild, "D:" & szTrigger_Name, szTrigger_Name, 2)
            End If
            NodeX.Tag = cmp_Trigger_CreateSQL(szTrigger_Name, szTrigger_Table, szTrigger_Function, szTrigger_Arguments, szTrigger_Foreach, szTrigger_Executes, szTrigger_Event)
                       
            If szTrigger_iscompiled = False Then
                NodeX.Image = 3
            Else
                NodeX.Image = 2
            End If
        Next iLoop
      End If
      Erase szTrigger
  Tree.Nodes.Item(iDev_Index).Text = Tree.Nodes.Item(iDev_Index).Text & " (" & CStr(iDev_Count) & ")"
  End If
  
  Set rsTrigger = Nothing
    
  EndMsg
Exit Sub

Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "basTrigger, cmp_Trigger_tree_refresh"
End Sub

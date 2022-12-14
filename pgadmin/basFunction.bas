Attribute VB_Name = "basFunction"
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

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' General
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Function cmp_Function_Exists(szFunction_table As String, Optional ByVal szFunction_name As String, Optional ByVal szFunction_arguments As String) As Boolean
On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    If (szFunction_table = "") Then
        szFunction_table = "pgadmin_functions"
    End If
    cmp_Function_Exists = False
        
    If szFunction_name <> "" Then
        szQueryStr = "SELECT * FROM " & szFunction_table
        szQueryStr = szQueryStr & " WHERE Function_name = '" & szFunction_name & "'"
        szQueryStr = szQueryStr & " AND Function_arguments = '" & szFunction_arguments & "'"
        'Log
        LogMsg "Testing existence of function " & szFunction_name & " (" & szFunction_arguments & ") in " & szFunction_table & "..."
    Else
        Exit Function
    End If

    
    ' retrieve name and arguments of function to drop
    LogMsg "Executing: " & szQueryStr
 
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
        
    'Drop function if exists
    If Not rsComp.EOF Then
       cmp_Function_Exists = True
    End If
  Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_Exists"
End Function

Public Sub cmp_Function_Drop(szFunction_table As String, ByVal szFunction_name As String, ByVal szFunction_arguments As String)
On Error GoTo Err_Handler
    Dim szDropStr As String
    
    If (szFunction_table = "") Then
        szFunction_table = "pgadmin_functions"
    End If
    If szFunction_name = "" Then Exit Sub
    
    ' create drop query
    If (szFunction_table = "pgadmin_functions") Then
        szDropStr = "DROP FUNCTION " & QUOTE & szFunction_name & QUOTE & " (" & szFunction_arguments & ");"
    Else
        szDropStr = "DELETE FROM " & szFunction_table & " WHERE "
        szDropStr = szDropStr & "function_name='" & szFunction_name & "' AND function_arguments='" & szFunction_arguments & "'"
    End If
    
    ' Log information
    LogMsg "Dropping function " & szFunction_name & " (" & szFunction_arguments & ") in " & szFunction_table & "..."
    LogMsg "Executing: " & szDropStr
    
    ' Execute drop query and close log
    gConnection.Execute szDropStr
    If (szFunction_table = "pgadmin_functions") Then
        LogQuery szDropStr
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_DropIfExists"
End Sub

Public Sub cmp_Function_DropIfExists(szFunction_table As String, ByVal szFunction_name As String, ByVal szFunction_arguments As String)
On Error GoTo Err_Handler
    Dim szDropStr As String
    
    ' Development  -> szFunction_Table="pgadmin_dev_functions"
    ' Production   -> szFunction_Table="pgadmin_functions"
    If (szFunction_table = "") Then
        szFunction_table = "pgadmin_functions"
    End If
    If szFunction_name = "" Then Exit Sub
    
    'Drop function if exists
    If cmp_Function_Exists(szFunction_table, szFunction_name & "", szFunction_arguments & "") = True Then
        cmp_Function_Drop szFunction_table, szFunction_name & "", szFunction_arguments & ""
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_DropIfExists"
End Sub

Public Sub cmp_Function_Create(szFunction_table As String, ByVal szFunction_name As String, ByVal szFunction_arguments As String, ByVal szFunction_returns As String, ByVal szFunction_source As String, ByVal szFunction_language As String, ByVal szFunction_owner As String, ByVal szFunction_comments As String)
On Error GoTo Err_Handler
    Dim szQuery As String
    
    If (szFunction_table = "") Then
        szFunction_table = "pgadmin_functions"
    End If
    
    If (szFunction_table = "pgadmin_functions") Then
        szQuery = cmp_Function_CreateSQL(szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language)
    Else
        szFunction_source = Replace(szFunction_source, "'", "''")
        szFunction_comments = Replace(szFunction_comments, "'", "''")
        
        'szFunction_source = Replace(szFunction_source, vbCrLf, "\n")
        
        szQuery = "INSERT INTO " & szFunction_table & " (function_name, Function_arguments, Function_returns, Function_source, Function_language, Function_owner, Function_comments)"
        szQuery = szQuery & "VALUES ("
        szQuery = szQuery & "'" & szFunction_name & "', "
        szQuery = szQuery & "'" & szFunction_arguments & "', "
        szQuery = szQuery & "'" & szFunction_returns & "', "
        szQuery = szQuery & "'" & szFunction_source & "', "
        szQuery = szQuery & "'" & szFunction_language & "', "
        szQuery = szQuery & "'" & szFunction_owner & "', "
        szQuery = szQuery & "'" & szFunction_comments & "' "
        szQuery = szQuery & ");"
    End If
    
    'Log
    LogMsg "Creating function " & szFunction_name & "(" & szFunction_arguments & ") in " & szFunction_table & "..."
    LogMsg "Executing: " & szQuery
    
    'Execute
    gConnection.Execute szQuery
    If (szFunction_table = "pgadmin_functions") Then
        LogQuery szQuery
    
        ' Write comments
        szQuery = "COMMENT ON FUNCTION " & szFunction_name & " (" & szFunction_arguments & ") IS '" & Replace(szFunction_comments, "'", "''") & "'"
        LogQuery szQuery
        gConnection.Execute szQuery
    End If
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_Create"
If Err.Number = -2147467259 Then MsgBox "Function " & szFunction_name & " (" & szFunction_arguments & ") could not be compiled." & vbCrLf & "Check source code and compile again."
bContinueRebuilding = False
End Sub

Public Function cmp_Function_CreateSQL(ByVal szFunction_name As String, ByVal szFunction_argumentlist As String, ByVal szFunction_returns As String, ByVal szFunction_source As String, ByVal szFunction_language As String) As String
On Error GoTo Err_Handler
    Dim szCreateStr As String
    
    szCreateStr = "CREATE FUNCTION " & QUOTE & szFunction_name & "" & QUOTE & " ("
    szCreateStr = szCreateStr & szFunction_argumentlist & "" & ") " & vbCrLf
    szCreateStr = szCreateStr & "RETURNS " & szFunction_returns & " " & vbCrLf
    szCreateStr = szCreateStr & "AS '" & Replace(szFunction_source, "'", "''") & "' " & vbCrLf
    szCreateStr = szCreateStr & "LANGUAGE '" & szFunction_language & "'"
    
    cmp_Function_CreateSQL = szCreateStr
  Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_CreateSQL"
End Function

Public Sub cmp_Function_Move(szFunction_source_table As String, szFunction_target_table As String, szFunction_name As String, szFunction_arguments As String, Optional bPromptForReplace As Boolean)
On Error GoTo Err_Handler
    Dim szFunction_returns As String
    Dim szFunction_source As String
    Dim szFunction_language As String
    Dim szFunction_owner As String
    Dim szFunction_comments As String
    
    If IsMissing(bPromptForReplace) = True Then bPromptForReplace = True
    
    If szFunction_source_table = "" Then szFunction_source_table = "pgadmin_Functions"
    If szFunction_target_table = "" Then szFunction_target_table = "pgadmin_Functions"
    If szFunction_source_table = szFunction_target_table Then Exit Sub
    
    If cmp_Function_Exists(szFunction_source_table, szFunction_name, szFunction_arguments) Then
        cmp_Function_GetValues szFunction_source_table, szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language, szFunction_owner, szFunction_comments
        If cmp_Function_Exists(szFunction_target_table, szFunction_name, szFunction_arguments) Then
             If (bPromptForReplace = False) Then
                 cmp_Function_Drop szFunction_target_table, szFunction_name, szFunction_arguments
                 cmp_Function_Create szFunction_target_table, szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language, szFunction_owner, szFunction_comments
             Else
                If MsgBox("Replace existing target Function " & vbCrLf & szFunction_name & " ?", vbYesNo) = vbYes Then
                    cmp_Function_Drop szFunction_target_table, szFunction_name, szFunction_arguments
                    cmp_Function_Create szFunction_target_table, szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language, szFunction_owner, szFunction_comments
                End If
             End If
        Else
             cmp_Function_Create szFunction_target_table, szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language, szFunction_owner, szFunction_comments
        End If
        If bContinueRebuilding = True And szFunction_target_table = "pgadmin_Functions" Then
            cmp_Function_SetIsCompiled szFunction_source_table, szFunction_name, szFunction_arguments
        End If
    End If
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_Move"
End Sub


Sub cmp_Function_GetValues(szFunction_table As String, szFunction_name As String, szFunction_arguments As String, Optional szFunction_returns As String, Optional szFunction_source As String, Optional szFunction_language As String, Optional szFunction_owner As String, Optional szFunction_comments As String)
 On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    If (szFunction_table = "") Then szFunction_table = "pgadmin_functions"

    ' Select query

    szQueryStr = "SELECT * from " & szFunction_table
    szQueryStr = szQueryStr & " WHERE function_name = '" & szFunction_name & "'"
    szQueryStr = szQueryStr & " AND function_arguments = '" & szFunction_arguments & "'"
     
    ' open
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    
    If Not rsComp.EOF Then
        If Not (IsMissing(szFunction_name)) Then szFunction_name = rsComp!function_name & ""
        If Not (IsMissing(szFunction_arguments)) Then szFunction_arguments = rsComp!Function_arguments & ""
        If Not (IsMissing(szFunction_returns)) Then szFunction_returns = rsComp!Function_returns & ""
        If Not (IsMissing(szFunction_source)) Then szFunction_source = rsComp!function_source & ""
        If Not (IsMissing(szFunction_language)) Then szFunction_language = rsComp!function_language & ""
        If Not (IsMissing(szFunction_owner)) Then szFunction_owner = rsComp!function_owner & ""
        If Not (IsMissing(szFunction_comments)) Then szFunction_comments = rsComp!function_comments & ""
       
        If (szFunction_name <> "") And (szFunction_returns = "") Then szFunction_returns = "opaque"
        rsComp.Close
    Else
        If Not (IsMissing(szFunction_name)) Then szFunction_name = ""
        If Not (IsMissing(szFunction_arguments)) Then szFunction_arguments = ""
        If Not (IsMissing(szFunction_returns)) Then szFunction_returns = ""
        If Not (IsMissing(szFunction_source)) Then szFunction_source = ""
        If Not (IsMissing(szFunction_language)) Then szFunction_language = ""
        If Not (IsMissing(szFunction_owner)) Then szFunction_owner = ""
        If Not (IsMissing(szFunction_comments)) Then szFunction_comments = ""
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_GetValues"
End Sub

Public Sub cmp_Function_ParseName(szInput As String, szFunction_name As String, szFunction_arguments As String)
On Error GoTo Err_Handler

Dim iInstr As Integer
    iInstr = InStr(szInput, "(")
    If iInstr > 0 Then
        szFunction_name = Trim(Left(szInput, iInstr - 1))
        szFunction_arguments = Trim(Mid(szInput, iInstr + 1, Len(szInput) - iInstr - 1))
    Else
        szFunction_name = Trim(szInput)
        szFunction_arguments = ""
    End If
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Func_CopyToDev"
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Dependencies
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub cmp_Function_Dependency_Initialize(ByVal szDependency_table As String, ByVal szFunction_name As String)
On Error GoTo Err_Handler
    Dim szDependencyStr As String
    Dim rsComp As New Recordset
    
    ' Initialize function(child)->function(parent) dependencies
    szDependencyStr = "SELECT * FROM pgadmin_dev_functions WHERE "
    szDependencyStr = szDependencyStr & " function_source ~* '[[:<:]]" & szFunction_name & "[[:>:]]'; "
  
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szDependencyStr, gConnection, adOpenForwardOnly, adLockReadOnly
  
    If Not rsComp.EOF Then
        szDependencyStr = "INSERT INTO " & szDependency_table & " (dependency_parent_object, dependency_parent_name, dependency_child_object, dependency_child_name) "
        szDependencyStr = szDependencyStr & " SELECT 'function' AS dependency_parent_object, '" & szFunction_name & "' AS dependency_parent_name, 'function' AS dependency_child_object, function_name as dependency_child_name "
        szDependencyStr = szDependencyStr & " FROM pgadmin_dev_functions WHERE "
        szDependencyStr = szDependencyStr & " function_source ~* '[[:<:]]" & szFunction_name & "[[:>:]]'; "
        
        LogMsg "Executing: " & szDependencyStr
        gConnection.Execute szDependencyStr
    End If
    
    ' Initialize view(child)->function(parent) dependencies
    szDependencyStr = "SELECT * FROM pgadmin_dev_views WHERE "
    szDependencyStr = szDependencyStr & " view_definition ~* '[[:<:]]" & szFunction_name & "[[:>:]]'; "
  
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szDependencyStr, gConnection, adOpenForwardOnly, adLockReadOnly
  
    If Not rsComp.EOF Then
        szDependencyStr = "INSERT INTO " & szDependency_table & " (dependency_parent_object, dependency_parent_name, dependency_child_object, dependency_child_name) "
        szDependencyStr = szDependencyStr & " SELECT 'function' AS dependency_parent_object, '" & szFunction_name & "' AS dependency_parent_name, 'view' AS dependency_child_object, view_name as dependency_child_name "
        szDependencyStr = szDependencyStr & " FROM pgadmin_dev_views WHERE "
        szDependencyStr = szDependencyStr & " view_definition ~* '[[:<:]]" & szFunction_name & "[[:>:]]'; "
        
        LogMsg "Executing: " & szDependencyStr
        gConnection.Execute szDependencyStr
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_Dependency_Initialize"
End Sub

Public Sub cmp_Function_SetIsCompiled(ByVal szFunction_dev_table As String, ByVal szFunction_name As String, ByVal szFunction_arguments As String)
On Error GoTo Err_Handler
    Dim szQueryStr As String
    
   If szFunction_name & "" = "" Then Exit Sub
        
    szQueryStr = "UPDATE " & szFunction_dev_table & " SET function_iscompiled = 't'"
    szQueryStr = szQueryStr & " WHERE Function_name = '" & szFunction_name & "'"
    szQueryStr = szQueryStr & " AND Function_arguments = '" & szFunction_arguments & "'"
     
    LogMsg "Executing: " & szQueryStr
    gConnection.Execute szQueryStr
Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_SetIsCompiled"
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Tree
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Sub cmp_function_tree_copy_devtopro(Tree As TreeToy)
On Error GoTo Err_Handler

    Dim nodX As Node
    Dim szParentKey As String
    Dim bRefresh As Boolean
    
    Dim szFunction_name As String
    Dim szFunction_arguments As String
      
    Dim szMsgboxMessage As String
    
    If Tree.TreeCountChecked = 0 Then Exit Sub
    
    szMsgboxMessage = "WARNING!" & vbCrLf & vbCrLf & _
    "Compilation is intended for testing newly created function(s)." & vbCrLf & vbCrLf & _
    "Beware that if the required functions are used by other functions, " & vbCrLf & _
    "triggers or views, dependencies are broken. " & vbCrLf & vbCrLf & _
    "If you are not sure whether you might break dependencies" & vbCrLf & _
    "or not, use the Rebuild Project button instead." & vbCrLf & vbCrLf & _
    "Continue?"
    
    If MsgBox(szMsgboxMessage, vbYesNo) = vbYes Then
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
                    cmp_Function_ParseName nodX.Text, szFunction_name, szFunction_arguments
                    cmp_Function_Move gDevPostgresqlTables & "_functions", "pgadmin_functions", szFunction_name, szFunction_arguments, True
                    bRefresh = True
                End If
            End If
        Next
    End If
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "basFunction, cmp_function_tree_copy_devtopro"
End Sub

Public Sub cmp_function_tree_copy_protodev(Tree As TreeToy)
On Error GoTo Err_Handler
    Dim nodX As Node
    Dim szParentKey As String
    Dim bRefresh As Boolean
    
    Dim szFunction_name As String
    Dim szFunction_arguments As String
      
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
                  cmp_Function_ParseName nodX.Text, szFunction_name, szFunction_arguments
                  cmp_Function_Move "pgadmin_functions", gDevPostgresqlTables & "_functions", szFunction_name, szFunction_arguments, True
                  bRefresh = True
            End If
        End If
    Next
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "basFunction, cmp_function_tree_copy_protodev"
End Sub

Public Sub cmp_function_tree_export(Tree As TreeToy, cDialog As CommonDialog)
On Error GoTo Err_Handler
    Dim szExport As String
    Dim bExport As Boolean
    Dim szHeader As String
    
    Dim nodX As Node
    Dim szParentKey As String
    
    Dim szFunction_table As String
    Dim szFunction_name As String
    Dim szFunction_arguments As String
    Dim szFunction_returns As String
    Dim szFunction_source As String
    Dim szFunction_language As String
    Dim szFunction_owner As String
    Dim szFunction_comments As String
    
    If Tree Is Nothing Then Exit Sub
    
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
                    szFunction_table = "pgadmin_functions"
            Else
                    szFunction_table = gDevPostgresqlTables & "_functions"
            End If

            bExport = True
            
            cmp_Function_ParseName nodX.Text, szFunction_name, szFunction_arguments
            cmp_Function_GetValues szFunction_table, szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language, szFunction_owner, szFunction_comments
                
            If szFunction_name <> "" Then
                ' Header
                szExport = szExport & "/*" & vbCrLf
                szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
                szExport = szExport & "Function " & szFunction_name & "(" & szFunction_arguments & ")" & " -> " & szFunction_returns & vbCrLf
                If szFunction_comments <> "" Then szExport = szExport & szFunction_comments & vbCrLf
                szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
                szExport = szExport & "*/" & vbCrLf
                
                ' Function
                szExport = szExport & cmp_Function_CreateSQL(szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language) & vbCrLf & vbCrLf
            End If
        End If
    Next
    
    If bExport Then
        szHeader = "/*" & vbCrLf
        szHeader = szHeader & "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" & vbCrLf
        szHeader = szHeader & "The choice of a GNU generation " & vbCrLf
        szHeader = szHeader & "PostgreSQL     www.postgresql.org" & vbCrLf
        szHeader = szHeader & "pgAdmin        www.greatbridge.org/project/pgadmin" & vbCrLf
        szHeader = szHeader & "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" & vbCrLf
        szHeader = szHeader & "*/" & vbCrLf & vbCrLf
        szExport = szHeader & szExport
        MsgExportToFile cDialog, szExport, "sql", "Export functions"
    End If
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "basFunction, cmp_function_tree_export"
End Sub

Public Sub cmp_function_tree_drop(Tree As TreeToy)
 On Error GoTo Err_Handler
    Dim szFunction_table As String
    Dim szFunction_name As String
    Dim szFunction_arguments As String
    Dim nodX As Node
    Dim bDrop As Boolean
    
    Dim szParentKey As String
    
    If Tree Is Nothing Then Exit Sub
       
    StartMsg "Dropping Function(s)..."
        
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
                    szFunction_table = "pgadmin_functions"
                    bDrop = True
                
                    Case "Dev:"
                    szFunction_table = gDevPostgresqlTables & "_functions"
                    bDrop = True
                End Select
                     
                If bDrop = True Then
                    cmp_Function_ParseName nodX.Text, szFunction_name, szFunction_arguments
                    cmp_Function_DropIfExists szFunction_table, szFunction_name, szFunction_arguments
                End If
             End If
        Next
        Set nodX = Nothing
        
        EndMsg
    
Exit Sub
Err_Handler:
EndMsg
If Err.Number <> 0 Then LogError Err, "basFunction, cmp_function_tree_drop"
End Sub

Public Sub cmp_function_tree_refresh(Tree As TreeToy, bShowSystem As Boolean, iPro_Index As Integer, iSys_Index As Integer, iDev_Index As Integer)
On Error GoTo Err_Handler

  Dim NodeX As Node
  Dim szQuery As String
  Dim szFunc() As Variant
  Dim iLoop As Long
  Dim iUbound As Long
  
  Dim szFunction_oid As String
  Dim szFunction_name As String
  Dim szFunction_arguments As String
  Dim szFunction_returns As String
  Dim szFunction_source As String
  Dim szFunction_language As String
  Dim szFunction_iscompiled As Boolean
  
  Dim rsFunc As New Recordset
  
  StartMsg "Retrieving Function Names..."
  
  Tree.Nodes.Clear
  iDev_Count = 0
  iPro_Count = 0
  iSys_Count = 0
  
  ' Development functions
  szDev_Text = "Development functions"
  If DevMode = True Then
    Set NodeX = Tree.Nodes.Add(, tvwChild, "Dev:", szDev_Text, 1)
    iDev_Index = NodeX.Index
    NodeX.Expanded = False
  End If
  
  ' Production functions
  If DevMode = False Then
    szPro_Text = "User functions"
  Else
    szPro_Text = "Production functions"
  End If
  Set NodeX = Tree.Nodes.Add(, tvwChild, "Pro:", szPro_Text, 1)
  iPro_Index = NodeX.Index
  NodeX.Expanded = False
  
  ' System functions
  szSys_Text = "System functions"
  If bShowSystem = True Then
    Set NodeX = Tree.Nodes.Add(, tvwChild, "Sys:", szSys_Text, 1)
    iSys_Index = NodeX.Index
    NodeX.Expanded = False
  End If

 ' ---------------------------------------------------------------------
 ' Retrieve pgadmin_functions in one query
 ' ---------------------------------------------------------------------
  If rsFunc.State <> adStateClosed Then rsFunc.Close
  If bShowSystem = True Then
     szQuery = "SELECT function_oid, function_name, function_arguments, Function_returns, Function_source, Function_language FROM pgadmin_functions ORDER BY function_name"
  Else
     szQuery = "SELECT function_oid, function_name, function_arguments, Function_returns, Function_source, Function_language FROM pgadmin_functions WHERE function_oid > " & LAST_SYSTEM_OID & " AND function_name NOT LIKE 'pgadmin_%' AND function_name NOT LIKE 'pg_%' AND function_name NOT LIKE 'int4eq' ORDER BY function_name"
  End If
  LogMsg "Executing: " & szQuery
  rsFunc.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
  
  If Not (rsFunc.EOF) Then
    szFunc = rsFunc.GetRows
    iUbound = UBound(szFunc, 2)
    For iLoop = 0 To iUbound
         szFunction_oid = szFunc(0, iLoop) & ""
         szFunction_name = szFunc(1, iLoop) & ""
         szFunction_arguments = szFunc(2, iLoop) & ""
         szFunction_returns = szFunc(3, iLoop) & ""
         szFunction_source = szFunc(4, iLoop) & ""
         szFunction_language = szFunc(5, iLoop) & ""
         
         If CLng(szFunction_oid) < LAST_SYSTEM_OID Or Left(szFunction_name, 8) = "pgadmin_" Or Left(szFunction_name, 3) = "pg_" Or szFunction_name = "int4eq" Then
         ' ---------------------------------------------------------------------
         ' If it is a system function, add it to "S:" System node
         ' ---------------------------------------------------------------------
            iSys_Count = iSys_Count + 1
            If szFunction_arguments <> "" Then
                Set NodeX = Tree.Nodes.Add("Sys:", tvwChild, "S:" & szFunction_name & " (" & szFunction_arguments & ")", szFunction_name & " (" & szFunction_arguments & ")", 5)
            Else
                Set NodeX = Tree.Nodes.Add("Sys:", tvwChild, "S:" & szFunction_name, szFunction_name, 5)
            End If
        Else
         ' ---------------------------------------------------------------------
         ' Else it is a user function, add it to "P:" Production node
         ' ---------------------------------------------------------------------
            iPro_Count = iPro_Count + 1
            If szFunction_arguments <> "" Then
                Set NodeX = Tree.Nodes.Add("Pro:", tvwChild, "P:" & szFunction_name & "(" & szFunction_arguments & ")", szFunction_name & " (" & szFunction_arguments & ")", 4)
            Else
                Set NodeX = Tree.Nodes.Add("Pro:", tvwChild, "P:" & szFunction_name, szFunction_name, 4)
            End If
            If DevMode = False Then NodeX.Image = 6
        End If
        NodeX.Tag = cmp_Function_CreateSQL(szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language)
    Next iLoop
    Tree.Nodes.Item(iPro_Index).Text = Tree.Nodes.Item(iPro_Index).Text & " (" & CStr(iPro_Count) & ")"
    If iSys_Count > 0 Then Tree.Nodes.Item(iSys_Index).Text = Tree.Nodes.Item(iSys_Index).Text & " (" & CStr(iSys_Count) & ")"
  End If
  Erase szFunc
  
 ' ---------------------------------------------------------------------
 ' Retrieve pgadmin_dev_functions in one query
 ' ---------------------------------------------------------------------
 If DevMode = True Then
      If rsFunc.State <> adStateClosed Then rsFunc.Close
      szQuery = "SELECT function_oid, function_name, function_arguments, Function_returns, Function_source, Function_language, Function_iscompiled FROM " & gDevPostgresqlTables & "_functions" & " ORDER BY function_name"
      LogMsg "Executing: " & szQuery
      rsFunc.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
    
      If Not (rsFunc.EOF) Then
        szFunc = rsFunc.GetRows
        iUbound = UBound(szFunc, 2)
        iDev_Count = iUbound + 1
        For iLoop = 0 To iUbound
             szFunction_oid = szFunc(0, iLoop) & ""
             szFunction_name = szFunc(1, iLoop) & ""
             szFunction_arguments = szFunc(2, iLoop) & ""
             szFunction_returns = szFunc(3, iLoop) & ""
             szFunction_source = szFunc(4, iLoop) & ""
             szFunction_language = szFunc(5, iLoop) & ""
             If IsNull(szFunc(6, iLoop)) Then
                szFunction_iscompiled = False
             Else
                szFunction_iscompiled = szFunc(6, iLoop)
             End If
            If szFunction_arguments <> "" Then
                Set NodeX = Tree.Nodes.Add("Dev:", tvwChild, "D:" & szFunction_name & " (" & szFunction_arguments & ")", szFunction_name & " (" & szFunction_arguments & ")", 2)
            Else
                Set NodeX = Tree.Nodes.Add("Dev:", tvwChild, "D:" & szFunction_name, szFunction_name, 2)
            End If
            NodeX.Tag = cmp_Function_CreateSQL(szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language)
            If DevMode = True And szFunction_iscompiled = False Then NodeX.Image = 3
        Next iLoop
      End If
      Erase szFunc
      
    Tree.Nodes.Item(iDev_Index).Text = Tree.Nodes.Item(iDev_Index).Text & " (" & CStr(iDev_Count) & ")"
  End If
  
  Set rsFunc = Nothing
  Tree.Refresh
  
  EndMsg
Exit Sub

Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_function_tree_refresh"
End Sub

Attribute VB_Name = "basView"
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

'****
'**** Views
'****

Sub cmp_View_DropIfExists(szview_table As String, ByVal szView_name As String)
On Error GoTo Err_Handler
    Dim szDropStr As String
    
    ' Where should we get the values ?
    If (szview_table = "") Then szview_table = "pgadmin_views"
    If szView_name = "" Then Exit Sub
    
    ' Test existence of view
    If cmp_View_Exists(szview_table, szView_name) = True Then
        cmp_View_Drop szview_table, szView_name
    End If

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basView, cmp_View_DropIfExists"
End Sub

Sub cmp_View_Drop(szview_table As String, ByVal szView_name As String)
On Error GoTo Err_Handler
    Dim szDropStr As String
    
    ' Where should we get the values ?
    If (szview_table = "") Then szview_table = "pgadmin_views"
    If szView_name = "" Then Exit Sub
    
    ' create drop query
    If (szview_table = "pgadmin_views") Then
        szDropStr = "DROP VIEW " & QUOTE & szView_name & QUOTE
    Else
        szDropStr = "DELETE FROM " & szview_table & " WHERE view_name ='" & szView_name & "'"
    End If

    LogMsg "Executing: " & szDropStr
    gConnection.Execute szDropStr
    LogQuery szDropStr
        
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basView, cmp_View_DropIfExists"
End Sub

Function cmp_View_Exists(szview_table As String, ByVal szView_name As String) As Boolean
On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
  
    ' Where should we get the values ?
    If (szview_table = "") Then szview_table = "pgadmin_views"
    If szView_name = "" Then Exit Function
    
    cmp_View_Exists = False
    szQueryStr = "SELECT * FROM  " & szview_table
    szQueryStr = szQueryStr & " WHERE view_name = '" & szView_name & "' "

    
      ' retrieve name and arguments of function to drop
    LogMsg "Testing existence of view " & szView_name & " in " & szview_table & "..."
    LogMsg "Executing: " & szQueryStr

    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    

    If Not rsComp.EOF Then
        cmp_View_Exists = True
        rsComp.Close
    End If

Exit Function
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basView, cmp_View_DropIfExists"
End Function

Sub cmp_View_Create(szview_table As String, ByVal szView_name As String, ByVal szView_definition As String, ByVal szView_owner As String, ByVal szView_acl As String, ByVal szView_comments As String)
On Error GoTo Err_Handler
    Dim szQuery As String
    Dim szview_oid As Long
    Dim szView_query_oid As Variant
  
    ' Where should we get the values ?
    If (szview_table = "") Then szview_table = "pgadmin_views"
    
    If (szview_table = "pgadmin_views") Then
        szQuery = cmp_View_CreateSQL(szView_name, szView_definition)
    Else
        szView_definition = Trim(Replace(szView_definition, "'", "''"))
        szView_comments = Trim(Replace(szView_comments, "'", "''"))
    
        szQuery = "INSERT INTO " & szview_table & " (View_name, View_definition, View_comments) "
        szQuery = szQuery & "VALUES ("
        szQuery = szQuery & "'" & szView_name & "', "
        szQuery = szQuery & "'" & szView_definition & "', "
        szQuery = szQuery & "'" & szView_comments & "' "
        szQuery = szQuery & ");"
    End If
    
    LogMsg "Creating view " & szView_name & " in " & szview_table & "..."
    LogMsg "Executing: " & szQuery
    
    ' Execute drop query and close log
    gConnection.Execute szQuery
    
    If (szview_table = "pgadmin_views") Then
        LogQuery szQuery
    
        ' Write comments
        szQuery = "COMMENT ON VIEW " & szView_name & " IS '" & szView_comments & "'"
        LogQuery szQuery
        gConnection.Execute szQuery
    End If

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basView, cmp_Views_Create"
If Err.Number = -2147467259 Then MsgBox "View " & szView_name & " could not be compiled." & vbCrLf & "Check source code and compile again."
bContinueRebuilding = False
End Sub

Function cmp_View_CreateSQL(ByVal szView_name As String, ByVal szView_definition As String) As String
On Error GoTo Err_Handler
  Dim szQuery As String
    szQuery = Trim("CREATE VIEW " & QUOTE & szView_name & QUOTE & vbCrLf & " AS " & szView_definition)
    If Right(szQuery, 1) <> ";" Then szQuery = szQuery & ";"
    cmp_View_CreateSQL = szQuery

Exit Function
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basView, cmp_Views_Create"
End Function

Sub cmp_View_GetValues(szview_table As String, szView_name As String, Optional szView_definition As String, Optional szView_owner As String, Optional szView_acl As String, Optional szView_comments As String)
On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    ' Where should we get the values ?
    If (szview_table = "") Then szview_table = "pgadmin_views"
    If szView_name = "" Then Exit Sub
    
    ' Select query
    
    szQueryStr = "SELECT * from " & szview_table & " WHERE view_name = '" & szView_name & "'"
    LogMsg "Executing: " & szQueryStr
    
    ' open
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    
    If Not rsComp.EOF Then
        If Not (IsMissing(szView_name)) Then szView_name = rsComp!view_name & ""
        If Not (IsMissing(szView_owner)) Then szView_owner = rsComp!view_owner & ""
        If Not (IsMissing(szView_acl)) Then szView_acl = rsComp!view_acl & ""
        If (szview_table = "pgadmin_views") Then
            If Not (IsMissing(szView_definition)) Then szView_definition = cmp_View_GetViewDef(szView_name)
        Else
            If Not (IsMissing(szView_definition)) Then szView_definition = rsComp!view_definition & ""
        End If
        If Not (IsMissing(szView_comments)) Then szView_comments = rsComp!view_comments & ""
        rsComp.Close
    Else
        If Not (IsMissing(szView_name)) Then szView_name = ""
        If Not (IsMissing(szView_owner)) Then szView_owner = ""
        If Not (IsMissing(szView_acl)) Then szView_acl = ""
        If Not (IsMissing(szView_definition)) Then szView_definition = ""
        If Not (IsMissing(szView_comments)) Then szView_comments = ""
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basView, cmp_Views_GetValues"
End Sub

Public Function cmp_View_GetViewDef(ByVal lngView_Name As String) As String
On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsTemp As New Recordset
    cmp_View_GetViewDef = ""
    
    If lngView_Name = "" Then Exit Function
    
    szQueryStr = "SELECT pg_get_viewdef ('" & lngView_Name & "') as Result"
    'LogMsg "Executing: " & szQueryStr
    
    ' open
    If rsTemp.State <> adStateClosed Then rsTemp.Close
    rsTemp.Open szQueryStr, gConnection
    
    cmp_View_GetViewDef = ""
    If Not rsTemp.EOF Then
        cmp_View_GetViewDef = rsTemp!result & ""
    End If
    
    Exit Function
Err_Handler:
  cmp_View_GetViewDef = "Not a view"
End Function

Public Sub cmp_View_Move(szView_source_table As String, szView_target_table As String, szView_name As String, Optional bPromptForReplace As Boolean)
On Error GoTo Err_Handler
    Dim szView_definition As String
    Dim szView_owner As String
    Dim szView_acl As String
    Dim szView_comments As String
    
    If IsMissing(bPromptForReplace) = True Then bPromptForReplace = True
    
    If szView_source_table = "" Then szView_source_table = "pgadmin_views"
    If szView_target_table = "" Then szView_target_table = "pgadmin_views"
    If szView_source_table = szView_target_table Then Exit Sub
    
    If cmp_View_Exists(szView_source_table, szView_name) Then
        cmp_View_GetValues szView_source_table, szView_name, szView_definition, szView_owner, szView_acl, szView_comments
        If cmp_View_Exists(szView_target_table, szView_name) Then
             If (bPromptForReplace = False) Then
                 cmp_View_Drop szView_target_table, szView_name
                 cmp_View_Create szView_target_table, szView_name, szView_definition, szView_owner, szView_acl, szView_comments
             Else
                If MsgBox("Replace existing target view " & vbCrLf & szView_name & " ?", vbYesNo) = vbYes Then
                    cmp_View_Drop szView_target_table, szView_name
                    cmp_View_Create szView_target_table, szView_name, szView_definition, szView_owner, szView_acl, szView_comments
                End If
             End If
        Else
             cmp_View_Create szView_target_table, szView_name, szView_definition, szView_owner, szView_acl, szView_comments
        End If
        If bContinueRebuilding = True And szView_target_table = "pgadmin_views" Then
            cmp_View_SetIsCompiled szView_source_table, szView_name
        End If
    End If
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basFunction, cmp_View_CopyToDev_New"
End Sub

Public Sub cmp_View_SetIsCompiled(ByVal szview_dev_table As String, ByVal szView_name As String)
On Error GoTo Err_Handler
    Dim szQueryStr As String
    
    If szView_name & "" = "" Then Exit Sub
    
    szQueryStr = "UPDATE " & szview_dev_table & " SET view_iscompiled = 't'"
    szQueryStr = szQueryStr & " WHERE view_name = '" & szView_name & "'"
     
    LogMsg "Executing: " & szQueryStr
    gConnection.Execute szQueryStr
Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_SetIsCompiled"
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Dependencies
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub cmp_View_Dependency_Initialize(ByVal szDependency_table As String, ByVal szView_name As String)
On Error GoTo Err_Handler
    Dim szDependencyStr As String
    Dim rsComp As New Recordset
    
    ' Initialize view(child)->view(parent) dependencies
    szDependencyStr = "SELECT * FROM pgadmin_dev_views WHERE "
    szDependencyStr = szDependencyStr & " view_definition ~* '[[:<:]]" & szView_name & "[[:>:]]'; "
  
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szDependencyStr, gConnection, adOpenForwardOnly, adLockReadOnly
  
    If Not rsComp.EOF Then
        szDependencyStr = "INSERT INTO " & szDependency_table & " (dependency_parent_object, dependency_parent_name, dependency_child_object, dependency_child_name) "
        szDependencyStr = szDependencyStr & " SELECT 'view' AS dependency_parent_object, '" & szView_name & "' AS dependency_parent_name, 'view' AS dependency_child_object, view_name as dependency_child_name "
        szDependencyStr = szDependencyStr & " FROM pgadmin_dev_views WHERE "
        szDependencyStr = szDependencyStr & " view_definition ~* '[[:<:]]" & szView_name & "[[:>:]]'; "
        
        LogMsg "Executing: " & szDependencyStr
        gConnection.Execute szDependencyStr
    End If
    
    ' Initialize function(child)->view(parent) dependencies
    szDependencyStr = "SELECT * FROM pgadmin_dev_functions WHERE "
    szDependencyStr = szDependencyStr & " function_source ~* '[[:<:]]" & szView_name & "[[:>:]]'; "
  
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szDependencyStr, gConnection, adOpenForwardOnly, adLockReadOnly
  
    If Not rsComp.EOF Then
        szDependencyStr = "INSERT INTO " & szDependency_table & " (dependency_parent_object, dependency_parent_name, dependency_child_object, dependency_child_name) "
        szDependencyStr = szDependencyStr & " SELECT 'view' AS dependency_parent_object, '" & szView_name & "' AS dependency_parent_name, 'function' AS dependency_child_object, function_name as dependency_child_name "
        szDependencyStr = szDependencyStr & " FROM pgadmin_dev_functions WHERE "
        szDependencyStr = szDependencyStr & " function_source ~* '[[:<:]]" & szView_name & "[[:>:]]'; "
        
        LogMsg "Executing: " & szDependencyStr
        gConnection.Execute szDependencyStr
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_Dependency_Initialize"
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Tree
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Sub cmp_view_tree_copy_devtopro(Tree As TreeToy)
On Error GoTo Err_Handler

    Dim nodX As Node
    Dim szParentKey As String
    Dim bRefresh As Boolean
    
    Dim szView_name As String
    Dim szMsgboxMessage As String
    
    If Tree Is Nothing Then Exit Sub
    If Tree.TreeCountChecked = 0 Then Exit Sub
    
    szMsgboxMessage = "WARNING!" & vbCrLf & vbCrLf & _
    "Compilation is intended for testing newly created view(s)." & vbCrLf & vbCrLf & _
    "Beware that if the required views are used by other views, " & vbCrLf & _
    "triggers or functions, dependencies are broken. " & vbCrLf & vbCrLf & _
    "If you are not sure whether you might break dependencies" & vbCrLf & _
    "or not, use the Rebuild Project button instead." & vbCrLf & vbCrLf & _
    "Continue?"
    
    If MsgBox(szMsgboxMessage, vbYesNo) = vbYes Then
        bRefresh = False
        bContinueRebuilding = True
        For Each nodX In Tree.Nodes
            If (nodX.Checked = True) Then
                If nodX.Parent Is Nothing Then
                   szParentKey = "" ' skip
                Else
                   szParentKey = nodX.Parent.Key
                End If
    
                If szParentKey = "Dev:" And bContinueRebuilding = True Then
                    szView_name = nodX.Text
                    cmp_View_Move gDevPostgresqlTables & "_views", "pgadmin_views", szView_name, True
                    bRefresh = True
                End If
            End If
        Next
    End If
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "basview, cmp_view_tree_copy_devtopro"
End Sub

Public Sub cmp_view_tree_copy_protodev(Tree As TreeToy)
On Error GoTo Err_Handler
    Dim nodX As Node
    Dim szParentKey As String
    Dim bRefresh As Boolean
    
    Dim szView_name As String
      
    If Tree Is Nothing Then Exit Sub
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
                szView_name = nodX.Text
                cmp_View_Move "pgadmin_views", gDevPostgresqlTables & "_views", szView_name
                bRefresh = True
            End If
        End If
    Next
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "basview, cmp_view_tree_copy_protodev"
End Sub

Public Sub cmp_view_tree_export(Tree As TreeToy, cDialog As CommonDialog)
On Error GoTo Err_Handler

    Dim iLoop As Long
    Dim iListCount As Long
    Dim szExport As String
    Dim bExport As Boolean
    Dim szHeader As String
    
    Dim nodX As Node
    Dim szParentKey As String
    
    Dim szView_name As String
    Dim szView_owner As String
    Dim szView_acl As String
    Dim szView_comments As String
    Dim szView_definition As String
    
    Dim szview_table As String
    
    If Tree Is Nothing Then Exit Sub
    If Tree.TreeCountChecked = 0 Then Exit Sub
    
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
                    szview_table = "pgadmin_views"
            Else
                    szview_table = gDevPostgresqlTables & "_views"
            End If

            bExport = True
            szView_name = nodX.Text
            cmp_View_GetValues szview_table, szView_name, szView_definition, szView_owner, szView_acl, szView_comments

            If szView_name <> "" Then
    
                ' Header
                szExport = szExport & "/*" & vbCrLf
                szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
                szExport = szExport & "View " & szView_name & vbCrLf
                If szView_comments <> "" Then szExport = szExport & szView_comments & vbCrLf
                szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
                szExport = szExport & "*/" & vbCrLf
                
                ' Function
                szExport = szExport & cmp_View_CreateSQL(szView_name, szView_definition) & vbCrLf & vbCrLf
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
        MsgExportToFile cDialog, szExport, "sql", "Export views"
    End If
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "basview, cmp_view_tree_export"
End Sub

Public Sub cmp_view_tree_drop(Tree As TreeToy)
On Error GoTo Err_Handler
    Dim szview_table As String
    Dim szView_name As String
    Dim szview_arguments As String
    Dim nodX As Node
    Dim bDrop As Boolean
    
    Dim szParentKey As String
    
    If Tree Is Nothing Then Exit Sub
    If Tree.TreeCountChecked = 0 Then Exit Sub
       
    StartMsg "Dropping view(s)..."
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
                    szview_table = "pgadmin_views"
                    bDrop = True
                
                    Case "Dev:"
                    szview_table = gDevPostgresqlTables & "_views"
                    bDrop = True
                End Select
                     
                If bDrop = True Then
                    szView_name = nodX.Text
                    cmp_View_DropIfExists szview_table, szView_name
                End If
             End If
        Next
        Set nodX = Nothing
  EndMsg
    
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "basview, cmp_view_tree_drop"
End Sub

Public Sub cmp_view_tree_refresh(Tree As TreeToy, bShowSystem As Boolean, iPro_Index As Integer, iSys_Index As Integer, iDev_Index As Integer)
On Error GoTo Err_Handler

  Dim NodeX As Node
  Dim szQuery As String
  Dim szView() As Variant
  Dim iLoop As Long
  Dim iUbound As Long
  
  Dim szview_oid As String
  Dim szView_name As String
  Dim szView_definition As String
  Dim szview_iscompiled As String
  
  Dim rsView As New Recordset
  
  StartMsg "Retrieving view Names..."
  
  Tree.Nodes.Clear
  
  If DevMode = False Then
    szPro_Text = "User views"
  Else
    szPro_Text = "2 - Production views"
  End If
  
  Set NodeX = Tree.Nodes.Add(, tvwChild, "Pro:", szPro_Text, 1)
  iPro_Index = NodeX.Index
  NodeX.Expanded = False
  
  szDev_Text = "1 - Development views"
  If DevMode = True Then
    Set NodeX = Tree.Nodes.Add(, tvwChild, "Dev:", szDev_Text, 1)
    iDev_Index = NodeX.Index
    NodeX.Expanded = False
  End If
  
  szSys_Text = "3 - System views"
  If bShowSystem = True Then
    Set NodeX = Tree.Nodes.Add(, tvwChild, "Sys:", szSys_Text, 1)
    iSys_Index = NodeX.Index
    NodeX.Expanded = False
  End If

 ' ---------------------------------------------------------------------
 ' Retrieve pgadmin_views in one query
 ' ---------------------------------------------------------------------

  If rsView.State <> adStateClosed Then rsView.Close
  If bShowSystem = True Then
     szQuery = "SELECT view_oid, view_name FROM pgadmin_views ORDER BY view_name"
  Else
     szQuery = "SELECT view_oid, view_name FROM pgadmin_views WHERE view_oid > " & LAST_SYSTEM_OID & " AND view_name NOT LIKE 'pgadmin_%' AND view_name NOT LIKE 'pg_%' ORDER BY view_name"
  End If
  LogMsg "Executing: " & szQuery
  rsView.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
  
  If Not (rsView.EOF) Then
    szView = rsView.GetRows
    iUbound = UBound(szView, 2)
    For iLoop = 0 To iUbound
         szview_oid = szView(0, iLoop) & ""
         szView_name = szView(1, iLoop) & ""
         szView_definition = cmp_View_GetViewDef(szView_name)
         
         If CLng(szview_oid) < LAST_SYSTEM_OID Or Left(szView_name, 8) = "pgadmin_" Or Left(szView_name, 3) = "pg_" Then
         ' ---------------------------------------------------------------------
         ' If it is a system view, add it to "S:" System node
         ' ---------------------------------------------------------------------
            Set NodeX = Tree.Nodes.Add("Sys:", tvwChild, "S:" & szView_name, szView_name, 2)
        Else
         ' ---------------------------------------------------------------------
         ' Else it is a user view, add it to "P:" Production node
         ' ---------------------------------------------------------------------
            Set NodeX = Tree.Nodes.Add("Pro:", tvwChild, "P:" & szView_name, szView_name, 4)
        End If
        NodeX.Tag = cmp_View_CreateSQL(szView_name, szView_definition)
        NodeX.Image = 4
    Next iLoop
  End If
  Erase szView
  
 ' ---------------------------------------------------------------------
 ' Retrieve pgadmin_dev_views in one query
 ' ---------------------------------------------------------------------
 If DevMode = True Then
      If rsView.State <> adStateClosed Then rsView.Close
      szQuery = "SELECT view_oid, view_name, view_definition, view_iscompiled FROM " & gDevPostgresqlTables & "_views" & " ORDER BY view_name"
      LogMsg "Executing: " & szQuery
      rsView.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
    
      If Not (rsView.EOF) Then
        szView = rsView.GetRows
        iUbound = UBound(szView, 2)
        For iLoop = 0 To iUbound
            szview_oid = szView(0, iLoop) & ""
            szView_name = szView(1, iLoop) & ""
            szView_definition = szView(2, iLoop) & ""
            szview_iscompiled = szView(3, iLoop) & ""
            
            Set NodeX = Tree.Nodes.Add("Dev:", tvwChild, "D:" & szView_name, szView_name, 2)
            NodeX.Tag = cmp_View_CreateSQL(szView_name, szView_definition)
            
            If szview_iscompiled = "" Then
                NodeX.Image = 3
            Else
                NodeX.Image = 2
            End If
        Next iLoop
      End If
      Erase szView
  End If
  
  Set rsView = Nothing
    
  EndMsg
Exit Sub

Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "basview, cmp_view_tree_refresh"
End Sub

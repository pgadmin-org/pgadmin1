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

'****
'**** Views
'****

Sub cmp_View_DropIfExists(szview_table As String, ByVal lngView_oid As Long, Optional ByVal szview_name As String)
On Error GoTo Err_Handler
    Dim szDropStr As String
    
    ' Where should we get the values ?
    If (szview_table = "") Then szview_table = "pgadmin_views"
    
    ' Test existence of view
    If cmp_View_Exists(szview_table, lngView_oid, szview_name) = True Then
        cmp_View_Drop szview_table, 0, szview_name
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basView, cmp_View_DropIfExists"
End Sub

Sub cmp_View_Drop(szview_table As String, ByVal lngView_oid As Long, Optional ByVal szview_name As String)
On Error GoTo Err_Handler
    Dim szDropStr As String
    
    ' Where should we get the values ?
    If (szview_table = "") Then szview_table = "pgadmin_views"
    
    ' Test existence of view

    If szview_name = "" Then cmp_View_GetValues szview_table, lngView_oid, szview_name

    ' create drop query
    If (szview_table = "pgadmin_views") Then
        szDropStr = "DROP VIEW " & QUOTE & szview_name & QUOTE
        LogMsg "Executing: " & szDropStr
        gConnection.Execute szDropStr
        LogQuery szDropStr
    Else
        szDropStr = "DELETE FROM " & szview_table & " WHERE view_name ='" & szview_name & "'"
        LogMsg "Executing: " & szDropStr
        gConnection.Execute szDropStr
    End If
     
    LogMsg "Executing: " & szDropStr
    gConnection.Execute szDropStr
    LogQuery szDropStr
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basView, cmp_View_DropIfExists"
End Sub

Function cmp_View_Exists(szview_table As String, ByVal lngView_oid As Long, ByVal szview_name As String) As Boolean
On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
  
    ' Where should we get the values ?
    If (szview_table = "") Then szview_table = "pgadmin_views"
    
    cmp_View_Exists = False
    If lngView_oid <> 0 Then
        szQueryStr = "SELECT * FROM " & szview_table
        szQueryStr = szQueryStr & " WHERE view_OID = " & Str(lngView_oid)
    Else
        If szview_name <> "" Then
            szQueryStr = "SELECT * FROM  " & szview_table
            szQueryStr = szQueryStr & " WHERE view_name = '" & szview_name & "' "
        Else
            Exit Function
        End If
    End If
    
      ' retrieve name and arguments of function to drop
    LogMsg "Testing existence of view " & szview_name & " in " & szview_table & "..."
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

Sub cmp_View_Create(szview_table As String, ByVal szview_name As String, ByVal szview_definition As String)
On Error GoTo Err_Handler
    Dim szCreateStr As String
    Dim szview_oid As Long
    Dim szView_query_oid As Variant
  
    ' Where should we get the values ?
    If (szview_table = "") Then szview_table = "pgadmin_views"
    
    If (szview_table = "pgadmin_views") Then
        szCreateStr = cmp_View_CreateSQL(szview_name, szview_definition)
    Else
        szCreateStr = "INSERT INTO " & szview_table & " (View_name, View_definition)"
        szCreateStr = szCreateStr & "VALUES ("
        szCreateStr = szCreateStr & "'" & szview_name & "', "
        szCreateStr = szCreateStr & "'" & Replace(szview_definition, "'", "''") & "' "
        szCreateStr = szCreateStr & ");"
    End If
    
    LogMsg "Creating view " & szview_name & " in " & szview_table & "..."
    LogMsg "Executing: " & szCreateStr
    
    ' Execute drop query and close log
    gConnection.Execute szCreateStr
    LogQuery szCreateStr

  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basView, cmp_Views_Create"
  If Err.Number = -2147467259 Then MsgBox "View " & szview_name & " could not be compiled." & vbCrLf & "Check source code and compile again."
  bContinueRebuilding = False
End Sub

Function cmp_View_CreateSQL(ByVal szview_name As String, ByVal szview_definition As String) As String
On Error GoTo Err_Handler
  Dim szQuery As String
    szQuery = "CREATE VIEW " & szview_name & vbCrLf & " AS " & szview_definition & "; "
    cmp_View_CreateSQL = szQuery
  Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basView, cmp_Views_Create"
End Function

Sub cmp_View_GetValues(szview_table As String, lngView_oid As Long, Optional szview_name As String, Optional szview_definition As String, Optional szview_owner As String, Optional szView_acl As String, Optional szview_comments As String)
On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    ' Where should we get the values ?
    If (szview_table = "") Then szview_table = "pgadmin_views"
        
    ' Select query
    If lngView_oid <> 0 Then
        szQueryStr = "SELECT * from " & szview_table
        szQueryStr = szQueryStr & " WHERE view_OID = " & lngView_oid
    Else
        If IsMissing(szview_name) Then szview_name = ""
        szQueryStr = "SELECT * from " & szview_table & " WHERE view_name = '" & szview_name & "'"
    End If
    
    LogMsg "Executing: " & szQueryStr
    
    ' open
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    
    If Not rsComp.EOF Then
        If IsNull(rsComp!view_oid) Then
            lngView_oid = 0
        Else
            lngView_oid = rsComp!view_oid
        End If
        If Not (IsMissing(szview_name)) Then szview_name = rsComp!view_name & ""
        If Not (IsMissing(szview_owner)) Then szview_owner = rsComp!view_owner & ""
        If Not (IsMissing(szView_acl)) Then szView_acl = rsComp!view_acl & ""
        If (szview_table = "pgadmin_views") Then
            If Not (IsMissing(szview_definition)) Then szview_definition = cmp_View_GetViewDef(szview_name)
        Else
            If Not (IsMissing(szview_definition)) Then szview_definition = rsComp!view_definition & ""
        End If
        If Not (IsMissing(szview_comments)) Then szview_comments = rsComp!view_comments & ""
        rsComp.Close
    Else
        If Not (IsMissing(szview_name)) Then szview_name = ""
        If Not (IsMissing(szview_owner)) Then szview_owner = ""
        If Not (IsMissing(szView_acl)) Then szView_acl = ""
        If Not (IsMissing(szview_definition)) Then szview_definition = ""
        If Not (IsMissing(szview_comments)) Then szview_comments = ""
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
    LogMsg "Executing: " & szQueryStr
    
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

Public Sub cmp_View_Move(szView_source_table As String, szView_target_table As String, szview_name As String)
On Error GoTo Err_Handler
    Dim szview_definition As String
    
    If szView_source_table = "" Then szView_source_table = "pgadmin_views"
    If szView_target_table = "" Then szView_target_table = "pgadmin_views"
    If szView_source_table = szView_target_table Then Exit Sub
    
    If cmp_View_Exists(szView_source_table, 0, szview_name) Then
        cmp_View_GetValues szView_source_table, 0, szview_name, szview_definition
        
        If cmp_View_Exists(szView_target_table, 0, szview_name) = True Then
            If (MsgBox("Replace existing target view " & vbCrLf & szview_name & " ?", vbYesNo) = vbYes) Then
                cmp_View_Drop szView_target_table, 0, szview_name
                cmp_View_Create szView_target_table, szview_name, szview_definition
            End If
        Else
             cmp_View_Create szView_target_table, szview_name, szview_definition
        End If
    End If
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basFunction, cmp_View_CopyToDev_New"
End Sub

Public Sub cmp_View_DropAll(Optional szview_table As String)
On Error GoTo Err_Handler
    Dim szQuery As String
    Dim szView() As Variant
    Dim iLoop As Long
    Dim iUbound As Long
    Dim rsView As New Recordset
    Dim szview_name As String
    
    If IsMissing(szview_table) Or (szview_table = "") Then szview_table = "pgadmin_views"
        
    If (szview_table = "pgadmin_views") Then
        szQuery = "SELECT view_name FROM pgadmin_views " & _
        "  WHERE view_oid > " & LAST_SYSTEM_OID & _
        "  AND view_name NOT LIKE 'pgadmin_%' " & _
        "  AND view_name NOT LIKE 'pg_%' " & _
        "  ORDER BY view_name; "
        
        LogMsg "Dropping all views in pgadmin_views..."
        LogMsg "Executing: " & szQuery
        
        If rsView.State <> adStateClosed Then rsView.Close
        rsView.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
    
        If Not (rsView.EOF) Then
            szView = rsView.GetRows
            rsView.Close
            iUbound = UBound(szView, 2)
                For iLoop = 0 To iUbound
                     szview_name = szView(0, iLoop)
                     cmp_View_DropIfExists "", 0, szview_name
                Next iLoop
            Erase szView
        End If
    Else
        szQuery = "TRUNCATE " & szview_table
        LogMsg "Truncating " & szview_table & "..."
        LogMsg "Executing: " & szQuery
        gConnection.Execute szQuery
    End If
   
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basProject, cmp_View_DropAll"
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Tree
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Sub cmp_view_tree_copy_devtopro(Tree As TreeToy)
On Error GoTo Err_Handler

    Dim nodX As Node
    Dim szParentKey As String
    Dim bRefresh As Boolean
    
    Dim szview_name As String
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
                   szParentKey = nodX.Key
                Else
                   szParentKey = nodX.Parent.Key
                End If
    
                If szParentKey = "Dev:" And bContinueRebuilding = True Then
                    szview_name = nodX.Text
                    cmp_View_Move gDevPostgresqlTables & "_views", "pgadmin_views", szview_name
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
    
    Dim szview_name As String
      
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
                szview_name = nodX.Text
                cmp_View_Move "pgadmin_views", gDevPostgresqlTables & "_views", szview_name
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
    
    Dim szview_name As String
    Dim szview_owner As String
    Dim szView_acl As String
    Dim szview_comments As String
    Dim szview_definition As String
    
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
            szview_name = nodX.Text
            cmp_View_GetValues szview_table, 0, szview_name, szview_definition, szview_owner, szView_acl, szview_comments

            ' Header
            szExport = szExport & "/*" & vbCrLf
            szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
            szExport = szExport & "View " & szview_name & vbCrLf
            If szview_comments <> "" Then szExport = szExport & szview_comments & vbCrLf
            szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
            szExport = szExport & "*/" & vbCrLf
            
            ' Function
            szExport = szExport & cmp_View_CreateSQL(szview_name, szview_definition) & vbCrLf & vbCrLf
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
    Dim szview_name As String
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
                    szview_name = nodX.Text
                    cmp_View_DropIfExists szview_table, 0, szview_name
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

Public Sub cmp_view_tree_refresh(Tree As TreeToy, bShowSystem As Boolean)
On Error GoTo Err_Handler

  Dim NodeX As Node
  Dim szQuery As String
  Dim szView() As Variant
  Dim iLoop As Long
  Dim iUbound As Long
  
  Dim szview_oid As String
  Dim szview_name As String
  Dim szview_definition As String
  
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
  
  szSys_Text = "System views"
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
         szview_name = szView(1, iLoop) & ""
         szview_definition = cmp_View_GetViewDef(szview_name)
         
         If CLng(szview_oid) < LAST_SYSTEM_OID Or Left(szview_name, 8) = "pgadmin_" Or Left(szview_name, 3) = "pg_" Then
         ' ---------------------------------------------------------------------
         ' If it is a system view, add it to "S:" System node
         ' ---------------------------------------------------------------------
            Set NodeX = Tree.Nodes.Add("Sys:", tvwChild, "S:" & szview_name, szview_name, 2)
            NodeX.Tag = cmp_View_CreateSQL(szview_name, szview_definition)
        Else
         ' ---------------------------------------------------------------------
         ' Else it is a user view, add it to "P:" Production node
         ' ---------------------------------------------------------------------
            Set NodeX = Tree.Nodes.Add("Pro:", tvwChild, "P:" & szview_name, szview_name, 4)
            NodeX.Tag = cmp_View_CreateSQL(szview_name, szview_definition)
        End If
    Next iLoop
  End If
  Erase szView
  
 ' ---------------------------------------------------------------------
 ' Retrieve pgadmin_dev_views in one query
 ' ---------------------------------------------------------------------
 If DevMode = True Then
      If rsView.State <> adStateClosed Then rsView.Close
      szQuery = "SELECT view_oid, view_name, view_definition FROM " & gDevPostgresqlTables & "_views" & " ORDER BY view_name"
      LogMsg "Executing: " & szQuery
      rsView.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
    
      If Not (rsView.EOF) Then
        szView = rsView.GetRows
        iUbound = UBound(szView, 2)
        For iLoop = 0 To iUbound
            szview_oid = szView(0, iLoop) & ""
            szview_name = szView(1, iLoop) & ""
            szview_definition = szView(2, iLoop) & ""

            Set NodeX = Tree.Nodes.Add("Dev:", tvwChild, "D:" & szview_name, szview_name, 2)
            NodeX.Tag = cmp_View_CreateSQL(szview_name, szview_definition)
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

Public Sub cmp_view_tree_activatebuttons(Tree As TreeToy, iSelected As Integer, sz_key As String, bShowSystem As Boolean)
'On Error GoTo Err_Handler
    Dim szSelectedNodeRoot As String
    Dim nodX As Node
    Dim szMode As String
    Dim bExpanded As Boolean
    ' Count checked items
    iSelected = Tree.TreeCountChecked
    
    ' Find the mode of the selected item
    sz_key = ""
    If iSelected > 0 Then
        If Tree.SelectedItem.Parent Is Nothing Then
            sz_key = Tree.SelectedItem.Key
        Else
            sz_key = Tree.SelectedItem.Parent.Key
        End If
        
        Select Case sz_key
            Case "Pro:"
            If DevMode = True Then
                Tree.TreeSetChildren Tree.Nodes.Item(iDev_Index), False
            End If
            If bShowSystem = True Then
                Tree.TreeSetChildren Tree.Nodes.Item(iSys_Index), False
            End If
            
            Case "Dev:"
            Tree.TreeSetChildren Tree.Nodes.Item(iPro_Index), False
            If bShowSystem = True Then
                Tree.TreeSetChildren Tree.Nodes.Item(iSys_Index), False
            End If
            
            Case "Sys:"
            If DevMode = True Then
                Tree.TreeSetChildren Tree.Nodes.Item(iDev_Index), False
            End If
            Tree.TreeSetChildren Tree.Nodes.Item(iPro_Index), False
        End Select
    End If
  
    iSelected = Tree.TreeCountChecked
      
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basview, cmp_view_tree_activatebuttons"
End Sub


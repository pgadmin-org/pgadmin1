Attribute VB_Name = "basIndex"
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

Public Function cmp_Index_CreateSQL(ByVal szindex_name As String, ByVal bIsUnique As Boolean, ByVal szindex_table As String, ByVal szindex_type As String, ByVal szIndex_fields As String) As String
On Error GoTo Err_Handler
    Dim szCreateStr As String
    
    szCreateStr = "CREATE " & IIf(bIsUnique, "UNIQUE", "") & "INDEX" & _
    QUOTE & szindex_name & QUOTE & " ON " & _
    QUOTE & szindex_table & QUOTE & " USING " & szindex_type & " (" & szIndex_fields & ")"
     
    cmp_Index_CreateSQL = szCreateStr

Exit Function
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basIndex, cmp_Index_CreateSQL"
End Function

Public Function cmp_Index_Exists(szIndex_pgtable As String, Optional ByVal szindex_name As String) As Boolean
On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    If (szIndex_pgtable = "") Then
        szIndex_pgtable = "pgadmin_indexes"
    End If
    cmp_Index_Exists = False
        
    If szindex_name <> "" Then
        szQueryStr = "SELECT * FROM " & szIndex_pgtable
        szQueryStr = szQueryStr & " WHERE Index_name = '" & szindex_name & "'"
        'Log
        LogMsg "Testing existence of index " & szindex_name & " in " & szIndex_pgtable & "..."
    Else
        Exit Function
    End If

    
    ' retrieve name and arguments of function to drop
    LogMsg "Executing: " & szQueryStr
 
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
        
    'Drop function if exists
    If Not rsComp.EOF Then
       cmp_Index_Exists = True
    End If
  Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basIndex, cmp_Index_Exists"
End Function

Public Sub cmp_Index_Drop(szIndex_pgtable As String, ByVal szindex_name As String)
On Error GoTo Err_Handler
    Dim szDropStr As String
    
    If (szIndex_pgtable = "") Then
        szIndex_pgtable = "pgadmin_indexes"
    End If
    If szindex_name = "" Then Exit Sub
    
    ' create drop query
    If (szIndex_pgtable = "pgadmin_Indexes") Then
        szDropStr = "DROP INDEX " & QUOTE & szindex_name & QUOTE & ";"
    Else
        szDropStr = "DELETE FROM " & szIndex_pgtable & " WHERE "
        szDropStr = szDropStr & "Index_name='" & szindex_name & "'"
    End If
    
    ' Log information
    LogMsg "Dropping index " & szindex_name & " in " & szIndex_pgtable & "..."
    LogMsg "Executing: " & szDropStr
    
    ' Execute drop query and close log
    gConnection.Execute szDropStr
    If (szIndex_pgtable = "pgadmin_indexes") Then
        LogQuery szDropStr
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basIndex, cmp_Index_DropIfExists"
End Sub

Public Sub cmp_Index_DropIfExists(szIndex_pgtable As String, ByVal szindex_name As String)
On Error GoTo Err_Handler
    Dim szDropStr As String
    
    ' Development  -> szIndex_pgtable="pgadmin_dev_Indexes"
    ' Production   -> szIndex_pgtable="pgadmin_Indexes"
    If (szIndex_pgtable = "") Then
        szIndex_pgtable = "pgadmin_indexes"
    End If
    If szindex_name = "" Then Exit Sub
    
    'Drop function if exists
    If cmp_Index_Exists(szIndex_pgtable, szindex_name & "") = True Then
        cmp_Index_Drop szIndex_pgtable, szindex_name & ""
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basIndex, cmp_Index_DropIfExists"
End Sub

Public Sub cmp_Index_Create(ByVal szIndex_pgtable As String, ByVal szindex_name As String, ByVal bIndex_is_unique As Boolean, ByVal szindex_table As String, ByVal szindex_type As String, ByVal szIndex_fields As String, ByVal szindex_comments As String)
On Error GoTo Err_Handler
    Dim szQuery As String
    
    If (szIndex_pgtable = "") Then
        szIndex_pgtable = "pgadmin_indexes"
    End If
    
    If (szIndex_pgtable = "pgadmin_Indexes") Then
       szQuery = cmp_Index_CreateSQL(szindex_name, bIndex_is_unique, szindex_table, szindex_type, szIndex_fields)
    Else
        'szIndex_source = Replace(szIndex_source, "'", "''")
        'szIndex_comments = Replace(szIndex_comments, "'", "''")
   
        'szIndex_source = Replace(szIndex_source, vbCrLf, "\n")
        
        'szQuery = "INSERT INTO " & szIndex_table & " (Index_name, Index_arguments, Index_returns, Index_source, Index_language, Index_owner, Index_comments)"
        'szQuery = szQuery & "VALUES ("
        'szQuery = szQuery & "'" & szIndex_name & "', "
        'szQuery = szQuery & "'" & szIndex_arguments & "', "
        'szQuery = szQuery & "'" & szIndex_returns & "', "
        'szQuery = szQuery & "'" & szIndex_source & "', "
        'szQuery = szQuery & "'" & szIndex_language & "', "
        'szQuery = szQuery & "'" & szIndex_owner & "', "
        'szQuery = szQuery & "'" & szIndex_comments & "' "
        'szQuery = szQuery & ");"
   End If
    
    'Log
    LogMsg "Creating function " & szindex_name & ") in " & szIndex_pgtable & "..."
    LogMsg "Executing: " & szQuery
    
    'Execute
    gConnection.Execute szQuery
    If (szIndex_pgtable = "pgadmin_indexes") Then
        LogQuery szQuery
    
    ' Write comments
        szQuery = "COMMENT ON INDEX " & szindex_name & ") IS '" & Replace(szindex_comments, "'", "''") & "'"
       LogQuery szQuery
        gConnection.Execute szQuery
    End If
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basIndex, cmp_Index_Create"
If Err.Number = -2147467259 Then MsgBox "Index " & szindex_name & " could not be compiled." & vbCrLf & "Check source code and compile again."
bContinueRebuilding = False
End Sub

Public Sub cmp_Index_Move(ByVal szIndex_source_table As String, ByVal szIndex_target_table As String, ByVal szindex_name As String, Optional bPromptForReplace As Boolean)
On Error GoTo Err_Handler
    Dim bIndex_is_unique As Boolean
    Dim bIndex_is_primary As Boolean
    Dim bIndex_is_lossy As Boolean
    Dim szindex_table As String
    Dim szindex_type As String
    Dim szIndex_fields As String
    Dim szindex_definition As String
    Dim szindex_comments As String

    If IsMissing(bPromptForReplace) = True Then bPromptForReplace = True
    
    If szIndex_source_table = "" Then szIndex_source_table = "pgadmin_Indexes"
    If szIndex_target_table = "" Then szIndex_target_table = "pgadmin_Indexes"
    If szIndex_source_table = szIndex_target_table Then Exit Sub
    
    If cmp_Index_Exists(szIndex_source_table, szindex_name) Then
        cmp_Index_GetValues szIndex_source_table, szindex_name, bIndex_is_unique, _
        bIndex_is_primary, bIndex_is_lossy, szindex_table, _
        szindex_type, szIndex_fields, szindex_definition, szindex_comments
        
        If cmp_Index_Exists(szIndex_target_table, szindex_name) Then
             If (bPromptForReplace = False) Then
                 cmp_Index_Drop szIndex_target_table, szindex_name
                 cmp_Index_Create szIndex_target_table, szindex_name, bIndex_is_unique, _
                 szindex_table, szindex_type, szIndex_fields, szindex_comments
             Else
                If MsgBox("Replace existing target Index " & vbCrLf & szindex_name & " ?", vbYesNo) = vbYes Then
                 cmp_Index_Drop szIndex_target_table, szindex_name
                 cmp_Index_Create szIndex_target_table, szindex_name, bIndex_is_unique, _
                 szindex_table, szindex_type, szIndex_fields, szindex_comments
                End If
             End If
        Else
             cmp_Index_Create szIndex_target_table, szindex_name, bIndex_is_unique, _
                 szindex_table, szindex_type, szIndex_fields, szindex_comments
        End If
        If bContinueRebuilding = True And szIndex_target_table = "pgadmin_Indexes" Then
            cmp_Index_SetIsCompiled szIndex_source_table, szindex_name
        End If
    End If
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basIndex, cmp_Index_Move"
End Sub


Sub cmp_Index_GetValues(ByVal szIndex_pgtable As String, szindex_name As String, _
 bIndex_is_unique As Boolean, bIndex_is_primary As Boolean, bIndex_is_lossy As Boolean, _
 szindex_table As String, szindex_type As String, szIndex_fields As String, szindex_definition As String, _
 szindex_comments As String)
 
 On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    If (szIndex_pgtable = "") Then szIndex_pgtable = "pgadmin_Indexes"

    ' Select query

    szQueryStr = "SELECT * from " & szIndex_pgtable
    szQueryStr = szQueryStr & " WHERE Index_name = '" & szindex_name & _
    "' ORDER BY column_position"
     
    ' open
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    
    If Not rsComp.EOF Then
        szindex_name = rsComp!Index_name & ""
        bIndex_is_unique = rsComp!index_is_unique
        bIndex_is_primary = rsComp!index_is_primary
        bIndex_is_lossy = rsComp!index_is_lossy
        szindex_table = rsComp!index_table & ""
        szindex_type = rsComp!Index_type & ""
        szindex_definition = rsComp!Index_definition & ""
        szindex_comments = rsComp!index_comments & ""
       
        While Not (rsComp.EOF)
            If Not (IsNull(szIndex_fields)) Then szIndex_fields = szIndex_fields & ", "
            szIndex_fields = szIndex_fields & rsComp!Column_name
        Wend
        rsComp.Close
    Else
        szindex_name = ""
        bIndex_is_unique = Null
        bIndex_is_primary = Null
        bIndex_is_lossy = Null
        szindex_table = ""
        szindex_type = ""
        szindex_definition = ""
        szindex_comments = ""
        szIndex_fields = ""
    End If

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basIndex, cmp_Index_GetValues"
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Dependencies
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub cmp_Index_SetIsCompiled(ByVal szIndex_pgtable As String, ByVal szindex_name As String)
On Error GoTo Err_Handler
    Dim szQueryStr As String
    
   If szindex_name & "" = "" Then Exit Sub
        
    szQueryStr = "UPDATE " & szIndex_pgtable & " SET Index_iscompiled = 't'"
    szQueryStr = szQueryStr & " WHERE Index_name = '" & szindex_name & "'"
     
    LogMsg "Executing: " & szQueryStr
    gConnection.Execute szQueryStr
Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basIndex, cmp_Index_SetIsCompiled"
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Tree
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Sub cmp_index_tree_copy_devtopro(Tree As TreeToy)
On Error GoTo Err_Handler

    Dim nodX As Node
    Dim szParentKey As String
    Dim bRefresh As Boolean
    
    Dim szindex_name As String
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
                szindex_name = nodX.Text & ""
                cmp_Index_Move gDevPostgresqlTables & "_indexes", "pgadmin_Indexes", szindex_name, True
                bRefresh = True
            End If
        End If
    Next
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "basindex, cmp_index_tree_copy_devtopro"
End Sub

Public Sub cmp_index_tree_copy_protodev(Tree As TreeToy)
On Error GoTo Err_Handler
    Dim nodX As Node
    Dim szParentKey As String
    Dim bRefresh As Boolean
    
    Dim szindex_name As String
      
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
                  szindex_name = nodX.Text & ""
                  cmp_Index_Move "pgadmin_Indexes", gDevPostgresqlTables & "_Indexes", szindex_name, True
                  bRefresh = True
            End If
        End If
    Next
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "basindex, cmp_index_tree_copy_protodev"
End Sub

'Public Sub cmp_index_tree_export(Tree As TreeToy, cDialog As CommonDialog)
'On Error GoTo Err_Handler
'    Dim szExport As String
'    Dim bExport As Boolean
'    Dim szHeader As String
'
'    Dim nodX As Node
'    Dim szParentKey As String
'
'    Dim szindex_table As String
'    Dim szindex_name As String
'    Dim szindex_arguments As String
'    Dim szindex_returns As String
'    Dim szindex_source As String
'    Dim szindex_language As String
'    Dim szindex_owner As String
'    Dim szindex_comments As String
'
'    If Tree Is Nothing Then Exit Sub
'
 '   bExport = False
'    szExport = ""
'
'    For Each nodX In Tree.Nodes
'        If (nodX.Checked = True) Then
'            If nodX.Parent Is Nothing Then
'               szParentKey = nodX.Key
'            Else
'               szParentKey = nodX.Parent.Key
'            End If
'
'            If szParentKey = "Pro:" Or szParentKey = "Sys:" Then
'                    szindex_table = "pgadmin_Indexes"
'            Else
'                    szindex_table = gDevPostgresqlTables & "_Indexes"
'            End If''
'
'            bExport = True
'
'            cmp_Index_ParseName nodX.Text, szindex_name, szindex_arguments
'            cmp_Index_GetValues szindex_table, szindex_name, szindex_arguments, szindex_returns, szindex_source, szindex_language, szindex_owner, szindex_comments
'
'            If szindex_name <> "" Then
                ' Header
'                szExport = szExport & "/*" & vbCrLf
'                szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
'                szExport = szExport & "index " & szindex_name & "(" & szindex_arguments & ")" & " -> " & szindex_returns & vbCrLf
'                If szindex_comments <> "" Then szExport = szExport & szindex_comments & vbCrLf
'                szExport = szExport & "-------------------------------------------------------------------" & vbCrLf
'                szExport = szExport & "*/" & vbCrLf
'
'                ' index
'                szExport = szExport & cmp_Index_CreateSQL(szindex_name, szindex_arguments, szindex_returns, szindex_source, szindex_language) & vbCrLf & vbCrLf
'            End If
'        End If
'    Next
'
'    If bExport Then
'        szHeader = "/*" & vbCrLf
'        szHeader = szHeader & "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" & vbCrLf
'        szHeader = szHeader & "The choice of a GNU generation " & vbCrLf
'        szHeader = szHeader & "PostgreSQL     www.postgresql.org" & vbCrLf
'        szHeader = szHeader & "pgAdmin        www.greatbridge.org/project/pgadmin" & vbCrLf
'        szHeader = szHeader & "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" & vbCrLf
'        szHeader = szHeader & "*/" & vbCrLf & vbCrLf
'        szExport = szHeader & szExport
'        MsgExportToFile cDialog, szExport, "sql", "Export Indexes"
'    End If
'
'Exit Sub
'Err_Handler: If Err.Number <> 0 Then LogError Err, "basindex, cmp_index_tree_export"
'End Sub'

Public Sub cmp_index_tree_drop(Tree As TreeToy)
 On Error GoTo Err_Handler
    Dim szindex_table As String
    Dim szindex_name As String
    Dim nodX As Node
    Dim bDrop As Boolean
    
    Dim szParentKey As String
    
    If Tree Is Nothing Then Exit Sub
       
    StartMsg "Dropping indexe(s)..."
        
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
                    szindex_table = "pgadmin_Indexes"
                    bDrop = True
                
                    Case "Dev:"
                    szindex_table = gDevPostgresqlTables & "_Indexes"
                    bDrop = True
                End Select
                     
                If bDrop = True Then
                    szindex_name = nodX.Text & ""
                    cmp_Index_DropIfExists szindex_table, szindex_name
                End If
             End If
        Next
        Set nodX = Nothing
        
        EndMsg
    
Exit Sub
Err_Handler:
EndMsg
If Err.Number <> 0 Then LogError Err, "basindex, cmp_index_tree_drop"
End Sub

Public Sub cmp_index_tree_refresh(Tree As TreeToy, bShowSystem As Boolean, iPro_Index As Integer, iSys_Index As Integer, iDev_Index As Integer)
On Error GoTo Err_Handler

  Dim NodeX As Node
  Dim szQuery As String
  Dim szIndex() As Variant
  Dim iLoop As Long
  Dim iUbound As Long
  
  Dim szindex_oid As String
  Dim szindex_name As String
  Dim bIndex_is_unique As Boolean
  Dim bIndex_is_primary As Boolean
  Dim bIndex_is_lossy As Boolean
  Dim szindex_table As String
  Dim szindex_type As String
  Dim szIndex_fields As String
  Dim szindex_comments As String
  Dim szindex_iscompiled As Boolean
  Dim szindex_definition As String
  
  Dim rsIndex As New Recordset

  StartMsg "Retrieving index Names..."
  
  iPro_Count = 0
  iDev_Count = 0
  iSys_Count = 0
  Tree.Nodes.Clear
  
  If DevMode = False Then
    szPro_Text = "Indexes"
  Else
    szPro_Text = "Production Indexes"
  End If
  
  Set NodeX = Tree.Nodes.Add(, tvwChild, "Pro:", szPro_Text, 1)
  iPro_Index = NodeX.Index
  NodeX.Expanded = False
  
  szDev_Text = "Development Indexes"
  If DevMode = True Then
    Set NodeX = Tree.Nodes.Add(, tvwChild, "Dev:", szDev_Text, 1)
    iDev_Index = NodeX.Index
    NodeX.Expanded = False
  End If
  
  szSys_Text = "System Indexes"
  If bShowSystem = True Then
    Set NodeX = Tree.Nodes.Add(, tvwChild, "Sys:", szSys_Text, 1)
    iSys_Index = NodeX.Index
    NodeX.Expanded = False
  End If

 ' ---------------------------------------------------------------------
 ' Retrieve pgadmin_Indexes in one query
 ' ---------------------------------------------------------------------
  If rsIndex.State <> adStateClosed Then rsIndex.Close
  If bShowSystem = True Then
     szQuery = "SELECT DISTINCT ON(index_name) index_oid, index_name FROM pgadmin_indexes ORDER BY index_name"
  Else
     szQuery = "SELECT DISTINCT ON(index_name) index_oid, index_name FROM pgadmin_indexes WHERE index_oid > " & LAST_SYSTEM_OID & " AND index_name NOT LIKE 'pgadmin_%' AND index_name NOT LIKE 'pg_%' ORDER BY index_name"
  End If
  LogMsg "Executing: " & szQuery
  rsIndex.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
  
  If Not (rsIndex.EOF) Then
    szIndex = rsIndex.GetRows
    iUbound = UBound(szIndex, 2)
    For iLoop = 0 To iUbound
         szindex_oid = szIndex(0, iLoop) & ""
         szindex_name = szIndex(1, iLoop) & ""
         
         cmp_Index_GetValues "pgadmin_dev_indexes", szindex_name, bIndex_is_unique, _
         bIndex_is_primary, bIndex_is_lossy, szindex_table, szindex_type, szIndex_fields, _
         szindex_definition, szindex_comments
           
         If CLng(szindex_oid) < LAST_SYSTEM_OID Or Left(szindex_name, 8) = "pgadmin_" Or Left(szindex_name, 3) = "pg_" Or szindex_name = "int4eq" Then
            ' ---------------------------------------------------------------------
            ' If it is a system index, add it to "S:" System node
            ' ---------------------------------------------------------------------
            Set NodeX = Tree.Nodes.Add("Sys:", tvwChild, "S:" & szindex_name, szindex_name, 5)
            iSys_Count = iSys_Count + 1
          Else
            ' ---------------------------------------------------------------------
            ' Else it is a user index, add it to "P:" Production node
            ' ---------------------------------------------------------------------
            Set NodeX = Tree.Nodes.Add("Pro:", tvwChild, "P:" & szindex_name, szindex_name, 4)
            iPro_Count = iPro_Count + 1
            If DevMode = False Then NodeX.Image = 6
         End If
        
    Next iLoop
  End If
  Erase szIndex
  
 ' ---------------------------------------------------------------------
 ' Retrieve pgadmin_dev_indexs in one query
 ' ---------------------------------------------------------------------
 If DevMode = True Then
      If rsIndex.State <> adStateClosed Then rsIndex.Close
      szQuery = "SELECT DISTINCT ON(index_name) index_oid, index_name FROM " & gDevPostgresqlTables & "_indexs" & " ORDER BY index_name"
      LogMsg "Executing: " & szQuery
      rsIndex.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
    
      If Not (rsIndex.EOF) Then
        szIndex = rsIndex.GetRows
        iUbound = UBound(szIndex, 2)
        iDev_Count = iUbound + 1
        For iLoop = 0 To iUbound
             szindex_oid = szIndex(0, iLoop) & ""
             szindex_name = szIndex(1, iLoop) & ""
             
            Set NodeX = Tree.Nodes.Add("Dev:", tvwChild, "D:" & szindex_name, szindex_name, 2)
            If DevMode = True And szindex_iscompiled = False Then NodeX.Image = 3
        Next iLoop
      End If
      Erase szIndex
  End If
  
  Set rsIndex = Nothing
    
  EndMsg
Exit Sub

Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "basindex, cmp_index_tree_refresh"
End Sub


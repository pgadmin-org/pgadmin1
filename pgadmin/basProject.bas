Attribute VB_Name = "basProject"
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
'**** Project
'****
'****

Public Sub comp_Project_BackupViews(Optional ByVal Function_OldName, Optional ByVal Function_NewName, Optional ByVal Table_OldName, Optional ByVal Table_NewName)
'On Error GoTo Err_Handler
    Dim szQuery As String
    Dim szView_definition As String
    Dim rsComp As New Recordset
    
    ' pgadmin_dev_functions, pgadmin_dev_triggers, pgadmin_dev_views are temporary tables.
    ' We first copy pgadmin_functions, pgadmin_triggers, pgadmin_views into them
      
    szQuery = "TRUNCATE TABLE pgadmin_dev_views;" & _
    "  INSERT INTO pgadmin_dev_views SELECT * from " & _
    "  pgadmin_views " & _
    "  WHERE view_oid > " & LAST_SYSTEM_OID & _
    "  AND view_name NOT LIKE 'pgadmin_%' " & _
    "  AND view_name NOT LIKE 'pg_%' " & _
    "  ORDER BY view_name; " & _
    "  UPDATE pgadmin_dev_views SET view_iscompiled = 'f';"
    LogMsg "Initializing pgadmin_dev_views..."
    LogMsg "Executing: " & szQuery
    gConnection.Execute szQuery
       
    ' initialize pgadmin_dev_view
    szQuery = "SELECT * FROM pgadmin_dev_views ORDER BY view_oid"
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQuery, gConnection, adOpenDynamic
    
    While True = False
    'While Not rsComp.EOF
        'Copy view definition
        'szView_definition = Replace(cmp_View_GetViewDef(rsComp!view_name), "'", "''")
        
        'Rename underlying functions if needed
        If Not (IsMissing(Function_NewName)) And Not (IsMissing(Function_OldName)) Then
            If (Function_OldName <> "") And (Function_NewName <> "") And (Function_NewName <> Function_OldName) Then
                szView_definition = Replace(szView_definition, Function_OldName, Function_NewName)
            End If
        End If
        
        ' Rename underlying table if needed
        If Not (IsMissing(Table_NewName)) And Not (IsMissing(Table_OldName)) Then
            If (Table_OldName <> "") And (Table_NewName <> "") And (Table_NewName <> Table_OldName) Then
                szView_definition = Replace(szView_definition, Table_OldName, Table_NewName)
            End If
        End If
        
        ' Update definition of view
        szQuery = "UPDATE pgadmin_dev_views SET view_definition = '" & szView_definition & "' WHERE view_oid = '" & rsComp!view_oid & "'"
        gConnection.Execute szQuery
        rsComp.MoveNext
    Wend

    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, comp_View_DevBackup"
End Sub

Public Sub comp_Project_BackupTriggers(Optional ByVal Function_OldName, Optional ByVal Function_NewName, Optional ByVal Table_OldName, Optional ByVal Table_NewName)
'On Error GoTo Err_Handler
    Dim szQuery As String
    Dim szTrigger_function As String
    Dim rsComp As New Recordset
    
    ' pgadmin_dev_functions, pgadmin_dev_triggers, pgadmin_dev_views are temporary tables.
    ' We first copy pgadmin_functions, pgadmin_triggers, pgadmin_views into them
    
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
    
    ' Rename functions if needed
    If Not (IsMissing(Function_NewName)) And Not (IsMissing(Function_OldName)) Then
        If (Function_OldName <> "") And (Function_NewName <> "") And (Function_NewName <> Function_OldName) Then
            ' Looking for triggers based on Function_OldName
            szQuery = "SELECT pgadmin_dev_triggers WHERE trigger_function ILIKE '" & Function_OldName & "';"
            LogMsg "Looking for triggers based on function " & Function_OldName & "..."
            LogMsg "Executing: " & szQuery
            
            If rsComp.State <> adStateClosed Then rsComp.Close
            rsComp.Open szQuery, gConnection
            
            ' If found, replace Function_OldName by Function_NewName
            If Not rsComp.EOF Then
                szQuery = "UPDATE pgadmin_dev_triggers " & _
                " SET   trigger_function = '" & Function_NewName & "'" & _
                " WHERE trigger_function ILIKE '" & Function_OldName & "';"
                
                LogMsg "Renaming trigger underlying function " & Function_OldName & " into " & Function_NewName & "..."
                LogMsg "Executing: " & szQuery
                gConnection.Execute szQuery
            End If
        End If
   End If
   
    ' Rename tables if needed
    If Not (IsMissing(Table_NewName)) And Not (IsMissing(Table_OldName)) Then
        If (Table_OldName <> "") And (Table_NewName <> "") And (Table_NewName <> Table_OldName) Then
            ' Looking for triggers based on Function_OldName
            szQuery = "SELECT pgadmin_dev_tables WHERE trigger_table ILIKE '" & Table_OldName & "';"
            LogMsg "Looking for triggers based on table " & Table_OldName & "..."
            LogMsg "Executing: " & szQuery
            
            If rsComp.State <> adStateClosed Then rsComp.Close
            rsComp.Open szQuery, gConnection
            
            ' If found, replace Function_OldName by Function_NewName
            If Not rsComp.EOF Then
                szQuery = "UPDATE pgadmin_dev_triggers " & _
                " SET   trigger_table = '" & Table_NewName & "'" & _
                " WHERE trigger_table ILIKE '" & Table_OldName & "';"
                
                LogMsg "Renaming trigger underlying table " & Table_OldName & " into " & Table_NewName & "..."
                LogMsg "Executing: " & szQuery
                gConnection.Execute szQuery
            End If
        End If
    End If
    
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, comp_Project_Initialize"
End Sub

Public Sub comp_Project_BackupFunctions(Optional ByVal Function_OldName, Optional ByVal Function_NewName)
'On Error GoTo Err_Handler
    Dim szQuery As String
    Dim rsComp As New Recordset
    Dim szFunction_source As String
    
    ' pgadmin_dev_functions, pgadmin_dev_triggers, pgadmin_dev_views are temporary tables.
    ' We first copy pgadmin_functions, pgadmin_triggers, pgadmin_views into them
    
    szQuery = "TRUNCATE TABLE pgadmin_dev_functions;" & _
    "  INSERT INTO pgadmin_dev_functions SELECT * " & _
    "  FROM pgadmin_functions " & _
    "  WHERE function_name NOT LIKE '%_call_handler' " & _
    "  AND function_name NOT LIKE 'pgadmin_%' " & _
    "  AND function_name NOT LIKE 'pg_%' " & _
    "  AND function_oid > " & LAST_SYSTEM_OID & _
    "  ORDER BY function_oid ;" & _
    "  UPDATE pgadmin_dev_functions SET function_iscompiled = 'f';" & _
    "  UPDATE pgadmin_dev_functions SET function_returns = 'opaque' WHERE function_returns = NULL;"
    LogMsg "Initializing pgadmin_dev_functions..."
    LogMsg "Executing: " & szQuery
    gConnection.Execute szQuery
        
    szQuery = "TRUNCATE TABLE pgadmin_dev_dependencies;"
    LogMsg "Initializing pgadmin_dev_dependencies..."
    LogMsg "Executing: " & szQuery
    gConnection.Execute szQuery
    
    ' Then, we fill the pgadmin_dev_dependencies table
    szQuery = "SELECT * FROM pgadmin_dev_functions ORDER BY function_OID"
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQuery, gConnection, adOpenDynamic
    
    While Not rsComp.EOF
        cmp_Function_Dependency_Initialize rsComp!function_oid, rsComp!Function_name
        rsComp.MoveNext
    Wend
    
     ' Change function_name if needed
    If Not (IsMissing(Function_NewName)) And Not (IsMissing(Function_OldName)) Then
        If (Function_OldName <> "") And (Function_NewName <> "") And (Function_NewName <> Function_OldName) Then
            ' Looking for triggers based on Function_OldName
            szQuery = "SELECT pgadmin_dev_functions WHERE function_name ILIKE '" & Function_OldName & "';"
            LogMsg "Looking for functions named " & Function_OldName & "..."
            LogMsg "Executing: " & szQuery
            
            If rsComp.State <> adStateClosed Then rsComp.Close
            rsComp.Open szQuery, gConnection
            
            ' If found, replace Function_OldName by Function_NewName
            If Not rsComp.EOF Then
                szQuery = "UPDATE pgadmin_dev_functions " & _
                " SET   function_function = '" & Function_NewName & "'" & _
                " WHERE function_function ILIKE '" & Function_OldName & "';"
                
                LogMsg "Updating function_name in " & Function_OldName & " with " & Function_NewName & "..."
                LogMsg "Executing: " & szQuery
                gConnection.Execute szQuery
            End If
        End If
        
         ' Change function_source if needed
        If (Function_OldName <> "") And (Function_NewName <> "") And (Function_NewName <> Function_OldName) Then
            ' Looking for triggers based on Function_OldName
            szQuery = "SELECT pgadmin_dev_functions WHERE function_source ILIKE '%" & Function_OldName & "%';"
            LogMsg "Looking for functions containing function_source " & Function_OldName & "..."
            LogMsg "Executing: " & szQuery
            
            If rsComp.State <> adStateClosed Then rsComp.Close
            rsComp.Open szQuery, gConnection
            
            ' If found, replace Function_OldName by Function_NewName
            While Not rsComp.EOF
                szFunction_source = Replace(rsComp!Function_source, "'", "''")
                szFunction_source = Replace(szFunction_source, Function_OldName, Function_NewName)
                szQuery = "UPDATE pgadmin_dev_functions SET szFunction_source = '" & szFunction_source & "'"
                szQuery = szQuery & " WHERE function_oid = " & Str(rsComp!function_oid)
                
                LogMsg "Updating function_source in " & Function_OldName & " with " & szFunction_source & "..."
                LogMsg "Executing: " & szQuery
                
                gConnection.Execute szQuery
                rsComp.MoveNext
            Wend
        End If
    End If
        
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, comp_Project_BackupFunctions"
End Sub


Public Sub comp_Project_Initialize()
On Error GoTo Err_Handler
   comp_Project_BackupFunctions
   comp_Project_BackupViews
   comp_Project_BackupTriggers
   
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, comp_Project_Initialize"
End Sub

Public Function comp_Project_FindNextFunctionToCompile() As Long
On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    
    szQueryStr = "SELECT * From pgadmin_dev_functions WHERE function_iscompiled = 'f' ORDER BY function_oid"
    
    LogMsg "Looking for next function to compile..."
    LogMsg "Executing: " & szQueryStr
    
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection, adOpenDynamic
    
    comp_Project_FindNextFunctionToCompile = 0
    While Not rsComp.EOF
        If cmp_Function_HasSatisfiedDependencies(rsComp!function_oid) = True Then
            comp_Project_FindNextFunctionToCompile = rsComp!function_oid
            LogMsg "Next vailable function to compile has OID = " & Str(comp_Project_FindNextFunctionToCompile) & "..."
            Exit Function
        End If
        rsComp.MoveNext
    Wend
   
    Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, comp_Project_FindNextFunctionToCompile"
End Function

Public Sub comp_Project_RebuildTriggers()
On Error GoTo Err_Handler
    Dim rsTrigger As New Recordset
    Dim szQueryStr As String
    ' Obviously this does not work
    
    szQueryStr = "SELECT * From pgadmin_dev_triggers"
    
    LogMsg "Now relinking triggers..."
    LogMsg "Executing: " & szQueryStr
    
    If rsTrigger.State <> adStateClosed Then rsTrigger.Close
    rsTrigger.Open szQueryStr, gConnection, adOpenDynamic
    
    ' All triggers carry functions_OID that have been deleted
    ' Therefore, we cannot stop and must compile all triggers
    While Not rsTrigger.EOF
        cmp_Trigger_DropIfExists rsTrigger!trigger_oid, rsTrigger!trigger_name, rsTrigger!trigger_table
        cmp_Trigger_Create rsTrigger!trigger_name, rsTrigger!trigger_table, rsTrigger!trigger_function & "", rsTrigger!trigger_arguments & "", "", "", "", rsTrigger!trigger_type
        rsTrigger.MoveNext
    Wend

    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, comp_Project_RebuildTriggers"
End Sub

Public Sub comp_Project_RebuildViews()
On Error GoTo Err_Handler
    Dim rsViews As New Recordset
    Dim szQueryStr As String
    Dim szViewDefinition As String
    
    szQueryStr = "SELECT * From pgadmin_dev_views"
    
    LogMsg "Now relinking views..."
    LogMsg "Executing: " & szQueryStr
    
    If rsViews.State <> adStateClosed Then rsViews.Close
    rsViews.Open szQueryStr, gConnection, adOpenDynamic
    
    While Not rsViews.EOF
        cmp_View_DropIfExists rsViews!view_oid, rsViews!view_name
        cmp_View_Create rsViews!view_name, rsViews!view_definition
        rsViews.MoveNext
    Wend

    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, comp_Project_RebuildTriggers"
End Sub

Public Sub comp_Project_Compile()
On Error GoTo Err_Handler
    Dim lngNextFunctionToCompile_OID As Long
    
    bContinueCompilation = True
    lngNextFunctionToCompile_OID = comp_Project_FindNextFunctionToCompile
    While (lngNextFunctionToCompile_OID > 0) And (bContinueCompilation = True)
        cmp_Function_Compile (lngNextFunctionToCompile_OID)
        lngNextFunctionToCompile_OID = comp_Project_FindNextFunctionToCompile
    Wend
      
    ' We must always relink triggers and views
    ' even if function compilation was aborted
    comp_Project_RebuildTriggers
    comp_Project_RebuildViews
    
    If bContinueCompilation = True Then MsgBox ("Rebuilding successfull")
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, comp_Project_Compile"
End Sub

Attribute VB_Name = "basProject"
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
'**** Project
'****
'****

Public Sub cmp_Project_Initialize()
On Error GoTo Err_Handler
    Dim szDependency() As Variant
    Dim iLoop As Long
    Dim iUbound As Long
    Dim szQuery As String
    Dim rsDependency As New Recordset
    Dim szParent_name As String
    
    ' Initialize dependencies
    
    szQuery = "TRUNCATE TABLE pgadmin_dev_dependencies; " & _
    "UPDATE pgadmin_dev_functions SET function_iscompiled = 'f'; " & _
    "UPDATE pgadmin_dev_views SET view_iscompiled = 'f'; " & _
    "UPDATE pgadmin_dev_triggers SET trigger_iscompiled = 'f'; "

    LogMsg "Initializing pgadmin_dev_dependencies..."
    LogMsg "Executing: " & szQuery
    gConnection.Execute szQuery
    
    ' Initialize dependencies on functions
    szQuery = "SELECT function_name FROM pgadmin_dev_functions ORDER BY function_name"
    If rsDependency.State <> adStateClosed Then rsDependency.Close
    rsDependency.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
 
    If Not (rsDependency.EOF) Then
      szDependency = rsDependency.GetRows
      rsDependency.Close
      iUbound = UBound(szDependency, 2)
      For iLoop = 0 To iUbound
          szParent_name = szDependency(0, iLoop)
          cmp_Function_Dependency_Initialize "pgadmin_dev_dependencies", szParent_name
      Next iLoop
      Erase szDependency
    End If

    ' Initialize dependencies on views
    szQuery = "SELECT view_name FROM pgadmin_dev_views ORDER BY view_name"
    If rsDependency.State <> adStateClosed Then rsDependency.Close
    rsDependency.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
 
    If Not (rsDependency.EOF) Then
      szDependency = rsDependency.GetRows
      rsDependency.Close
      iUbound = UBound(szDependency, 2)
      For iLoop = 0 To iUbound
          szParent_name = szDependency(0, iLoop)
          cmp_View_Dependency_Initialize "pgadmin_dev_dependencies", szParent_name
      Next iLoop
      Erase szDependency
    End If
    
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basProject, cmp_Project_Initialize"
End Sub

Public Function cmp_Project_FindNextFunctionToCompile() As String
On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsFunc As New Recordset
    Dim szFunction_name As String
    Dim szFunction_arguments As String
    
    szQueryStr = "SELECT function_name, function_arguments " & _
    " FROM pgadmin_dev_functions" & _
    " WHERE function_iscompiled = 'f' AND function_name " & _
    " NOT IN" & _
    " (" & _
    " SELECT dependency_child_name FROM pgadmin_dev_dependencies d, pgadmin_dev_functions f" & _
    " WHERE d.dependency_parent_name = f.function_name " & _
    " AND f.function_iscompiled = 'f'" & _
    " ) " & _
    " ORDER BY function_name"
       
    LogMsg "Looking for next function to compile..."
    LogMsg "Executing: " & szQueryStr
    
    If rsFunc.State <> adStateClosed Then rsFunc.Close
    rsFunc.Open szQueryStr, gConnection, adOpenForwardOnly, adLockReadOnly
    
    cmp_Project_FindNextFunctionToCompile = ""
    If Not (rsFunc.EOF) Then
      szFunction_name = rsFunc(0) & ""
      szFunction_arguments = rsFunc(1) & ""
      cmp_Project_FindNextFunctionToCompile = szFunction_name & "(" & szFunction_arguments & ")"
    End If
   
    rsFunc.Close
    Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basProject, cmp_Project_FindNextFunctionToCompile"
End Function

Public Sub cmp_Project_Move_Functions(szFunction_source_table As String, szFunction_source_clause As String, szFunction_target_table As String)
On Error GoTo Err_Handler
    Dim rsFunc As New Recordset
    Dim szQuery As String
    Dim szFunction As Variant
    Dim szFunction_name As String
    Dim szFunction_arguments As String
    Dim iUbound As Long
    Dim iLoop As Long
    
    If ObjectExists(szFunction_source_table, tTable) = 0 Then Exit Sub
    
    If ObjectExists(szFunction_target_table, tTable) = 0 And ObjectExists(szFunction_target_table, tView) = 0 Then
        szQuery = "CREATE TABLE " & szFunction_target_table & " AS SELECT * FROM " & szFunction_source_table & " LIMIT 1 ; TRUNCATE TABLE " & szFunction_target_table
        If Not SuperuserChk Then Exit Sub
        LogMsg "Executing: " & szQuery
        gConnection.Execute szQuery
        LogMsg "Executing: GRANT all ON " & szFunction_target_table & " TO public"
        gConnection.Execute "GRANT all ON " & szFunction_target_table & " TO public"
    End If
    
    szQuery = "SELECT function_name, function_arguments FROM " & szFunction_source_table
    If szFunction_source_clause <> "" Then szQuery = szQuery & " WHERE " & szFunction_source_clause
    
    LogMsg "Now rebuilding functions..."
    LogMsg "Executing: " & szQuery
    
    If rsFunc.State <> adStateClosed Then rsFunc.Close
    rsFunc.Open szQuery, gConnection, adOpenForwardOnly
    
    If Not (rsFunc.EOF) Then
        szFunction = rsFunc.GetRows
        iUbound = UBound(szFunction, 2)
        rsFunc.Close
        For iLoop = 0 To iUbound
            szFunction_name = szFunction(0, iLoop) & ""
            szFunction_arguments = szFunction(1, iLoop) & ""
            cmp_Function_Move szFunction_source_table, szFunction_target_table, szFunction_name, szFunction_arguments, False
        Next
    End If
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basProject, cmp_Project_Move_Functions"
End Sub

Public Sub cmp_Project_Move_Triggers(szTrigger_source_table As String, szTrigger_source_clause As String, szTrigger_target_table As String)
On Error GoTo Err_Handler
    Dim szQuery As String
    Dim szTrigger() As Variant
    Dim iLoop As Long
    Dim iUbound As Long
    Dim rsTriggers As New Recordset
    Dim szTrigger_Name As String
    Dim szTrigger_Table As String
 
    If ObjectExists(szTrigger_source_table, tTable) = 0 Then Exit Sub
    
     If ObjectExists(szTrigger_target_table, tTable) = 0 And ObjectExists(szTrigger_target_table, tView) = 0 Then
        szQuery = "CREATE TABLE " & szTrigger_target_table & " AS SELECT * FROM " & szTrigger_source_table & " LIMIT 1 ; TRUNCATE TABLE " & szTrigger_target_table
        If Not SuperuserChk Then Exit Sub
        LogMsg "Executing: " & szQuery
        gConnection.Execute szQuery
        LogMsg "Executing: GRANT all ON " & szTrigger_target_table & " TO public"
        gConnection.Execute "GRANT all ON " & szTrigger_target_table & " TO public"
    End If
    
    szQuery = "SELECT trigger_name, trigger_table FROM " & szTrigger_source_table
    If szTrigger_source_clause <> "" Then szQuery = szQuery & " WHERE " & szTrigger_source_clause
    
    LogMsg "Now relinking triggers..."
    LogMsg "Executing: " & szQuery
    
    If rsTriggers.State <> adStateClosed Then rsTriggers.Close
    rsTriggers.Open szQuery, gConnection, adOpenForwardOnly
    
    If Not (rsTriggers.EOF) Then
        szTrigger = rsTriggers.GetRows
        iUbound = UBound(szTrigger, 2)
        rsTriggers.Close
        For iLoop = 0 To iUbound
            szTrigger_Name = szTrigger(0, iLoop) & ""
            szTrigger_Table = szTrigger(1, iLoop) & ""
            cmp_Trigger_Move szTrigger_source_table, szTrigger_target_table, szTrigger_Name, szTrigger_Table, False
        Next
    End If
 
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basProject, cmp_Project_Move_Triggers"
End Sub

Public Function cmp_Project_FindNextViewToCompile() As String
On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsFunc As New Recordset
    Dim szView_name As String
    
    szQueryStr = "SELECT view_name " & _
    " FROM pgadmin_dev_views" & _
    " WHERE view_iscompiled = 'f' AND view_name " & _
    " NOT IN" & _
    " (" & _
    " SELECT dependency_child_name FROM pgadmin_dev_dependencies d, pgadmin_dev_views v" & _
    " WHERE d.dependency_parent_name = v.view_name " & _
    " AND v.view_iscompiled = 'f'" & _
    " ) " & _
    " ORDER BY view_name"
    
    LogMsg "Looking for next view to compile..."
    LogMsg "Executing: " & szQueryStr
    
    If rsFunc.State <> adStateClosed Then rsFunc.Close
    rsFunc.Open szQueryStr, gConnection, adOpenForwardOnly, adLockReadOnly
    
    cmp_Project_FindNextViewToCompile = ""
    If Not (rsFunc.EOF) Then
      szView_name = rsFunc(0) & ""
      cmp_Project_FindNextViewToCompile = szView_name
      LogMsg "Next vailable View to compile is " & cmp_Project_FindNextViewToCompile & "..."
    End If
    rsFunc.Close

Exit Function
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basProject, cmp_Project_FindNextViewToCompile"
End Function

Public Sub cmp_Project_Move_Views(szView_source_table As String, szView_source_clause As String, szView_target_table As String)
On Error GoTo Err_Handler
    Dim rsViews As New Recordset
    Dim szQuery As String
    Dim szView As Variant
    Dim szView_name As String
    Dim iUbound As Long
    Dim iLoop As Long
    
    If ObjectExists(szView_source_table, tTable) = 0 Then Exit Sub
    
    If ObjectExists(szView_target_table, tTable) = 0 And ObjectExists(szView_target_table, tView) = 0 Then
        szQuery = "CREATE TABLE " & szView_target_table & " AS SELECT * FROM " & szView_source_table & " LIMIT 1 ; TRUNCATE TABLE " & szView_target_table
        If Not SuperuserChk Then Exit Sub
        LogMsg "Executing: " & szQuery
        gConnection.Execute szQuery
        LogMsg "Executing: GRANT all ON " & szView_target_table & " TO public"
        gConnection.Execute "GRANT all ON " & szView_target_table & " TO public"
    End If
    
    szQuery = "SELECT view_name FROM " & szView_source_table
    If szView_source_clause <> "" Then szQuery = szQuery & " WHERE " & szView_source_clause
    
    LogMsg "Now relinking views..."
    LogMsg "Executing: " & szQuery
    
    If rsViews.State <> adStateClosed Then rsViews.Close
    rsViews.Open szQuery, gConnection, adOpenForwardOnly
    
    If Not (rsViews.EOF) Then
        szView = rsViews.GetRows
        iUbound = UBound(szView, 2)
        rsViews.Close
        For iLoop = 0 To iUbound
            szView_name = szView(0, iLoop) & ""
            cmp_View_Move szView_source_table, szView_target_table, szView_name, False
        Next
    End If
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basProject, cmp_Project_Move_Views"
End Sub

Public Sub cmp_Project_Compile()
On Error GoTo Err_Handler
    Dim szNextObject_name As String
    Dim szFunction_name As String
    Dim szFunction_arguments As String
    Dim szView_name As String
    
    bContinueRebuilding = True
    szFunction_name = "Go on"
    szView_name = "Go on"
    
    While ((szFunction_name <> "") Or (szView_name <> "")) And (bContinueRebuilding = True)

    ' Rebuild functions
    '
        While (szFunction_name <> "") And (bContinueRebuilding = True)
            szNextObject_name = cmp_Project_FindNextFunctionToCompile
            cmp_Function_ParseName szNextObject_name, szFunction_name, szFunction_arguments
            If szFunction_name <> "" Then
                cmp_Function_Move "pgadmin_dev_functions", "pgadmin_functions", szFunction_name, szFunction_arguments, False
            End If
        Wend
    
        ' Rebuild views
        '
        While (szView_name <> "") And (bContinueRebuilding = True)
            szNextObject_name = cmp_Project_FindNextViewToCompile
            szView_name = szNextObject_name & ""
            If szView_name <> "" Then
                cmp_View_Move "pgadmin_dev_Views", "pgadmin_Views", szView_name, False
            End If
        Wend
        
        szFunction_name = cmp_Project_FindNextFunctionToCompile
        szView_name = cmp_Project_FindNextViewToCompile
    Wend
    
     ' Rebuild triggers
    If bContinueRebuilding = True Then cmp_Project_Move_Triggers gDevPostgresqlTables & "_triggers", "", "pgadmin_triggers"
    
    If bContinueRebuilding = True Then
        MsgBox ("Rebuilding successfull")
    End If

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "basProject, cmp_Project_Compile"
End Sub

Public Function cmp_Project_IsRebuilt() As Boolean
Dim iResult As Long

    cmp_Project_IsRebuilt = False
    
    iResult = RsExecuteGetResult("SELECT COUNT (*) FROM pgadmin_dev_functions WHERE function_oid = NULL")
    If iResult > 0 Then Exit Function
    
    iResult = RsExecuteGetResult("SELECT COUNT (*) FROM pgadmin_dev_triggers WHERE trigger_oid = NULL")
    If iResult > 0 Then Exit Function
    
    iResult = RsExecuteGetResult("SELECT COUNT (*) FROM pgadmin_dev_views WHERE view_oid = NULL")
    If iResult > 0 Then Exit Function
    
    cmp_Project_IsRebuilt = True
End Function

Public Sub cmp_Project_Rebuild()
On Error GoTo Err_Handler
    If MsgBox("Please confirm you wish to continue.", vbYesNo + vbQuestion, _
            "Rebuild project") = vbYes Then
        cmp_Project_Initialize
        cmp_Project_Compile
    End If
    
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basProject, cmp_Project_Rebuild"
End Sub


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
'On Error GoTo Err_Handler
    Dim szFunc() As Variant
    Dim iLoop As Long
    Dim iUbound As Long
    Dim szQuery As String
    Dim rsFunc As New Recordset
    Dim szFunction_source As String
    Dim szFunction_name As String
    
    ' Initialize dependencies
    
    szQuery = "TRUNCATE TABLE pgadmin_dev_dependencies; UPDATE pgadmin_dev_functions SET function_iscompiled = 'f';"
    LogMsg "Initializing pgadmin_dev_dependencies..."
    LogMsg "Executing: " & szQuery
    gConnection.Execute szQuery
    
    szQuery = "SELECT function_name FROM pgadmin_dev_functions ORDER BY function_oid"
    If rsFunc.State <> adStateClosed Then rsFunc.Close
    rsFunc.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
 
    If Not (rsFunc.EOF) Then
      szFunc = rsFunc.GetRows
      rsFunc.Close
      iUbound = UBound(szFunc, 2)
      For iLoop = 0 To iUbound
          szFunction_name = szFunc(0, iLoop)
          cmp_Function_Dependency_Initialize "pgadmin_dev_dependencies", "pgadmin_dev_functions", szFunction_name
      Next iLoop
      Erase szFunc
    End If

    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basProject, cmp_Project_Initialize"
End Sub

Public Function cmp_Project_FindNextFunctionToCompile() As String
'On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim szFunc() As Variant
    Dim iLoop As Long
    Dim iUbound As Long
    Dim rsFunc As New Recordset
    Dim szFunction_name As String
    Dim szFunction_arguments As String
    
    szQueryStr = "SELECT function_name, function_arguments From pgadmin_dev_functions WHERE function_iscompiled = 'f' ORDER BY function_oid"
    
    LogMsg "Looking for next function to compile..."
    LogMsg "Executing: " & szQueryStr
    
    If rsFunc.State <> adStateClosed Then rsFunc.Close
    rsFunc.Open szQueryStr, gConnection, adOpenForwardOnly, adLockReadOnly
    
    cmp_Project_FindNextFunctionToCompile = ""
    If Not (rsFunc.EOF) Then
      szFunc = rsFunc.GetRows
      rsFunc.Close
      iUbound = UBound(szFunc, 2)
      For iLoop = 0 To iUbound
           szFunction_name = szFunc(0, iLoop)
           szFunction_arguments = szFunc(1, iLoop)
           If cmp_Function_HasSatisfiedDependencies("pgadmin_dev_functions", "pgadmin_dev_dependencies", szFunction_name) = True Then
                cmp_Project_FindNextFunctionToCompile = szFunction_name & "(" & szFunction_arguments & ")"
                LogMsg "Next vailable function to compile is " & cmp_Project_FindNextFunctionToCompile & "..."
                Exit Function
            End If
      Next iLoop
      Erase szFunc
    End If
   
    Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basProject, cmp_Project_FindNextFunctionToCompile"
End Function

Public Sub cmp_Project_RebuildTriggers()
'On Error GoTo Err_Handler
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
        cmp_Trigger_DropIfExists "", 0, rsTrigger!trigger_name, rsTrigger!trigger_table
        cmp_Trigger_Create "", rsTrigger!trigger_name, rsTrigger!trigger_table, rsTrigger!trigger_function & "", rsTrigger!Trigger_arguments & "", rsTrigger!Trigger_foreach & "", rsTrigger!Trigger_executes & "", rsTrigger!Trigger_event & ""
        rsTrigger.MoveNext
    Wend

    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basProject, cmp_Project_RebuildTriggers"
End Sub

Public Sub cmp_Project_RebuildViews()
'On Error GoTo Err_Handler
    Dim rsViews As New Recordset
    Dim szQueryStr As String
    Dim szViewDefinition As String
    
    szQueryStr = "SELECT * From pgadmin_dev_views"
    
    LogMsg "Now relinking views..."
    LogMsg "Executing: " & szQueryStr
    
    If rsViews.State <> adStateClosed Then rsViews.Close
    rsViews.Open szQueryStr, gConnection, adOpenDynamic
    
    While Not rsViews.EOF
        cmp_View_DropIfExists "", 0, rsViews!view_name
        cmp_View_Create "", rsViews!view_name, rsViews!view_definition
        rsViews.MoveNext
    Wend

    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basProject, cmp_Project_RebuildViews"
End Sub

Public Sub cmp_Project_Compile()
'On Error GoTo Err_Handler
    Dim szNextFunctionToCompile_name As String
    Dim szFunction_name As String
    Dim szFunction_arguments As String
    
    bContinueRebuilding = True
    szNextFunctionToCompile_name = cmp_Project_FindNextFunctionToCompile
    cmp_Function_ParseName szNextFunctionToCompile_name, szFunction_name, szFunction_arguments
    
    While (szFunction_name <> "") And (bContinueRebuilding = True)
        cmp_Function_Compile "pgadmin_dev_functions", szFunction_name, szFunction_arguments
        szNextFunctionToCompile_name = cmp_Project_FindNextFunctionToCompile
        cmp_Function_ParseName szNextFunctionToCompile_name, szFunction_name, szFunction_arguments
    Wend
      
    ' We must always relink triggers and views
    ' even if function compilation was aborted
    cmp_Project_RebuildTriggers
    cmp_Project_RebuildViews
    
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
'On Error GoTo Err_Handler
    If MsgBox("Rebuilding feature does not keep comments and ACL." & vbCrLf & "Please confirm you wish to continue.", vbYesNo + vbQuestion, _
            "Rebuild project") = vbYes Then
        cmp_Project_Initialize
        cmp_Project_Compile
    End If
    
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basProject, cmp_Project_Rebuild"
End Sub


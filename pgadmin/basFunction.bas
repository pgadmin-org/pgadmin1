Attribute VB_Name = "basFunction"
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
'****
'****
'****

Public Function cmp_Function_Exists(ByVal szFunction_OID As Long, Optional ByVal szFunction_name As String, Optional ByVal szFunction_arguments As String) As Boolean
On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    cmp_Function_Exists = False
        
    If szFunction_OID <> 0 Then
        szQueryStr = "SELECT * FROM pgadmin_functions "
        szQueryStr = szQueryStr & "WHERE Function_OID = " & szFunction_OID
        
        ' Log
        LogMsg "Testing existence of function OID = " & szFunction_OID & "..."
    Else
        If szFunction_name <> "" Then
            szQueryStr = "SELECT * FROM pgadmin_functions "
            szQueryStr = szQueryStr & "WHERE Function_name = '" & szFunction_name & "' "
            If szFunction_arguments <> "" Then
                szQueryStr = szQueryStr & "AND Function_arguments = '" & szFunction_arguments & "'"
            Else
                szQueryStr = szQueryStr & "AND Function_arguments = '' "
            End If
            'Log
            LogMsg "Testing existence of function " & szFunction_name & " (" & szFunction_arguments & ")..."
        Else
            Exit Function
        End If
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
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Function_Exists"
End Function

Public Sub cmp_Function_DropIfExists(ByVal szFunction_OID As Long, Optional ByVal szFunction_name As String, Optional ByVal szFunction_arguments As String)
On Error GoTo Err_Handler
    Dim szDropStr As String
    
    'Drop function if exists
    If cmp_Function_Exists(szFunction_OID, szFunction_name & "", szFunction_arguments & "") = True Then
        ' Retrieve function name and arguments if we only know the OID
        If szFunction_OID <> 0 Then cmp_Function_GetValues szFunction_OID, "", szFunction_name, szFunction_arguments
        
        ' create drop query
        szDropStr = "DROP FUNCTION " & QUOTE & szFunction_name & QUOTE & " (" & szFunction_arguments & ");"
                
        ' Log information
        LogMsg "Dropping function " & szFunction_name & " (" & szFunction_arguments & ")..."
        LogMsg "Executing: " & szDropStr
        
        ' Execute drop query and close log
        gConnection.Execute szDropStr
        LogQuery szDropStr
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Function_DropIfExists"
End Sub

Public Sub cmp_Function_Create(ByVal szFunction_name As String, ByVal szFunction_argumentlist As String, ByVal szFunction_returns As String, ByVal szFunction_source As String, ByVal szFunction_language As String)
On Error GoTo Err_Handler
    Dim szCreateStr As String

    szCreateStr = cmp_Function_CreateSQL(szFunction_name, szFunction_argumentlist, szFunction_returns, szFunction_source, szFunction_language)
    
    'Log
    LogMsg "Creating function " & szFunction_name & "(" & szFunction_argumentlist & ") ..."
    LogMsg "Executing: " & szCreateStr
    
    'Execute
    gConnection.Execute szCreateStr
    LogQuery szCreateStr
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Function_Create"
  If Err.Number = -2147467259 Then MsgBox "Function " & szFunction_name & " (" & szFunction_argumentlist & ") could not be compiled." & vbCrLf & "Check source code and compile again."
  bContinueCompilation = False
End Sub

Public Function cmp_Function_CreateSQL(ByVal szFunction_name As String, ByVal szFunction_argumentlist As String, ByVal szFunction_returns As String, ByVal szFunction_source As String, ByVal szFunction_language As String) As String
On Error GoTo Err_Handler
    Dim szCreateStr As String

    szCreateStr = "CREATE FUNCTION " & QUOTE & szFunction_name & "" & QUOTE & " ("
    szCreateStr = szCreateStr & szFunction_argumentlist & "" & ") " & vbCrLf
    szCreateStr = szCreateStr & "RETURNS " & szFunction_returns & " " & vbCrLf
    szCreateStr = szCreateStr & "AS '" & vbCrLf
    szCreateStr = szCreateStr & szFunction_source & vbCrLf
    szCreateStr = szCreateStr & "' " & vbCrLf
    szCreateStr = szCreateStr & "LANGUAGE '" & szFunction_language & "'"

    cmp_Function_CreateSQL = szCreateStr
  Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Function_CreateSQL"
End Function

Public Sub cmp_Function_Compile(ByVal lngFunction_OID As Long)
On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    Dim szFunction_name As String
    Dim szFunction_arguments As String
    Dim szFunction_returns As String
    Dim szFunction_language As String
    Dim szFunction_source As String

    ' Retrieve function
    cmp_Function_GetValues lngFunction_OID, "", szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language
 
    ' Compile function if exists
    If szFunction_name <> "" Then
        ' Attempt to create a temporary function to see if it compiles
        LogMsg "Checking if " & szFunction_name & " (" & szFunction_arguments & ") can be compiled ..."
        cmp_Function_DropIfExists 0, "pgadmin_fake__" & Left(szFunction_name, 15), szFunction_arguments
        cmp_Function_Create "pgadmin_fake__" & Left(szFunction_name, 15), szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language
        cmp_Function_DropIfExists 0, "pgadmin_fake__" & Left(szFunction_name, 15), szFunction_arguments
    
       If bContinueCompilation = True Then
            ' If it does, compile the real function
            cmp_Function_DropIfExists lngFunction_OID
            cmp_Function_Create szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language
        
           ' Tell PgAdmin that the function was compiled
            cmp_Function_SetIsCompiled szFunction_name, szFunction_arguments
            LogMsg szFunction_name & " (" & szFunction_arguments & ") was successfuly compiled."
        End If
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Function_Compile"
End Sub

Public Sub cmp_Function_Dependency_Initialize(ByVal lngFunction_OID As Long, ByVal szFunction_name As String)
On Error GoTo Err_Handler
    Dim szDependencyStr As String
    Dim rsComp As New Recordset
    
    ' Drop existing dependencies
    cmp_Function_Dependency_DropIfExists lngFunction_OID
    
    ' Scan pgadmin_dev_functions for dependencies
     
    szDependencyStr = "SELECT * FROM pgadmin_dev_functions WHERE function_source ILIKE '%" & szFunction_name & "%'"
    szDependencyStr = szDependencyStr & " AND Function_OID <> " & lngFunction_OID
    
    ' Log
    LogMsg "Scanning pgadmin_dev_functions for dependencies ..."
    LogMsg "Executing: " & szDependencyStr
    
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szDependencyStr, gConnection, adOpenDynamic
  
    ' Write dependencies in pgadmin_dev_dependencies
    If Not rsComp.EOF Then
        szDependencyStr = "INSERT INTO pgadmin_dev_dependencies (dependency_from, dependency_to) "
        szDependencyStr = szDependencyStr & " SELECT " & Str(lngFunction_OID) & " AS dependency_from, pgadmin_dev_functions.function_OID as dependency_to "
        szDependencyStr = szDependencyStr & " FROM pgadmin_dev_functions WHERE "
        szDependencyStr = szDependencyStr & " function_source ilike '%" & szFunction_name & "%' "
        szDependencyStr = szDependencyStr & " AND function_OID <> " & lngFunction_OID
        
        ' Log
        LogMsg "Writing dependencies..."
        LogMsg "Executing: " & szDependencyStr
        
        gConnection.Execute szDependencyStr
    End If
    
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Function_Dependency_Initialize"
End Sub

Public Sub cmp_Function_Dependency_DropIfExists(ByVal lngFunction_OID As Long)
On Error GoTo Err_Handler
    Dim szDependencyStr As String
    Dim rsComp As New Recordset
    
    szDependencyStr = "SELECT * FROM pgadmin_dev_dependencies WHERE dependency_from = " & lngFunction_OID
    
    ' retrieve name and arguments of function to drop
    LogMsg "Testing existence of function OID = " & lngFunction_OID
    LogMsg "Executing: " & szDependencyStr
 
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szDependencyStr, gConnection
        
    'Drop function if exists
    If Not rsComp.EOF Then
       szDependencyStr = "DELETE FROM pgadmin_dev_dependencies WHERE dependency_from = " & Str(lngFunction_OID)
       
       LogMsg "Dropping dependencies on function OID = " & Str(lngFunction_OID)
       LogMsg "Executing: " & szDependencyStr
    
       gConnection.Execute szDependencyStr
    End If
Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Function_Dependency_DropIfExists"
End Sub

Public Sub cmp_Function_SetIsCompiled(ByVal szFunction_name As String, ByVal szFunction_arguments As String)
On Error GoTo Err_Handler
    Dim szQueryStr As String

    szQueryStr = "UPDATE pgadmin_dev_functions SET function_iscompiled = 't'"
    szQueryStr = szQueryStr & " WHERE Function_name = '" & szFunction_name & "'"
    szQueryStr = szQueryStr & " AND Function_arguments = '" & szFunction_arguments & "'"
     
    LogMsg "Setting function " & szFunction_name & " (" & szFunction_arguments & "" & ") to IsCompiled=TRUE..."
    LogMsg "Executing: " & szQueryStr
    
    gConnection.Execute szQueryStr
Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Function_SetIsCompiled"
End Sub

Public Function cmp_Function_HasSatisfiedDependencies(ByVal lngFunction_OID As Long) As Boolean
    On Error GoTo Err_Handler
    
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    ' Test existence of unsatisfied dependencies
    szQueryStr = "SELECT pgadmin_dev_functions.Function_OID, pgadmin_dev_functions.function_name, pgadmin_dev_functions_1.function_iscompiled"
    szQueryStr = szQueryStr & " From pgadmin_dev_functions"
    szQueryStr = szQueryStr & "    INNER JOIN pgadmin_dev_dependencies"
    szQueryStr = szQueryStr & "    ON pgadmin_dev_functions.Function_OID = pgadmin_dev_dependencies.dependency_from"
    szQueryStr = szQueryStr & "    INNER JOIN pgadmin_dev_functions AS pgadmin_dev_functions_1"
    szQueryStr = szQueryStr & "    ON pgadmin_dev_dependencies.dependency_to =  pgadmin_dev_functions_1.Function_OID"
    szQueryStr = szQueryStr & "    WHERE ((pgadmin_dev_functions.Function_OID = " & Str(lngFunction_OID) & ") AND (pgadmin_dev_functions_1.function_iscompiled = 'f'));"
    
    LogMsg "Testing existence of satisfied dependencies on function OID = " & lngFunction_OID
    LogMsg "Executing: " & szQueryStr
  
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    
    cmp_Function_HasSatisfiedDependencies = False
    If rsComp.EOF Then
        cmp_Function_HasSatisfiedDependencies = True
    End If
    
    Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Function_HasSatisfiedDependencies"
End Function

Sub cmp_Function_GetValues(lngFunction_OID As Long, Optional szFunction_PostgreSQLtable As String, Optional szFunction_name As String, Optional szFunction_arguments As String, Optional szFunction_returns As String, Optional szFunction_source As String, Optional szFunction_language As String, Optional szFunction_owner As String, Optional szFunction_comments As String)
 On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    ' Where should we get the values ?
        If IsMissing(szFunction_PostgreSQLtable) Or (szFunction_PostgreSQLtable = "") Then
            szFunction_PostgreSQLtable = "pgadmin_functions"
        End If

    ' Select query
    If lngFunction_OID <> 0 Then
        szQueryStr = "SELECT * from " & szFunction_PostgreSQLtable
        szQueryStr = szQueryStr & " WHERE function_OID = " & lngFunction_OID
    Else
        If IsMissing(szFunction_name) Then szFunction_name = ""
            szQueryStr = "SELECT * from " & szFunction_PostgreSQLtable
            szQueryStr = szQueryStr & " WHERE function_name = '" & szFunction_name & "'"
            If Not (IsMissing(szFunction_arguments)) Then
                szQueryStr = szQueryStr & " AND function_arguments = '" & szFunction_arguments & "'"
            End If
    End If
    
    ' Log information
    LogMsg "Retrieving information from function OID =" & lngFunction_OID & " in table " & szFunction_PostgreSQLtable & "..."
    LogMsg "Executing: " & szQueryStr
    
    ' open
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    
    If Not rsComp.EOF Then
        lngFunction_OID = rsComp!function_oid
        If Not (IsMissing(szFunction_name)) Then szFunction_name = rsComp!Function_name & ""
        If Not (IsMissing(szFunction_arguments)) Then szFunction_arguments = rsComp!Function_arguments & ""
        If Not (IsMissing(szFunction_returns)) Then szFunction_returns = rsComp!Function_returns & ""
        If Not (IsMissing(szFunction_source)) Then szFunction_source = rsComp!Function_source & ""
        If Not (IsMissing(szFunction_language)) Then szFunction_language = rsComp!Function_language & ""
        If Not (IsMissing(szFunction_owner)) Then szFunction_owner = rsComp!function_owner & ""
        If Not (IsMissing(szFunction_comments)) Then szFunction_comments = rsComp!function_comments & ""
       
        If (szFunction_name <> "") And (szFunction_returns = "") Then szFunction_returns = "opaque"
        szFunction_source = Replace(szFunction_source, "'", "''")
        rsComp.Close
    Else
        lngFunction_OID = 0
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
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Function_GetValues"
End Sub

Sub cmp_Function_GetCurrentValues(ByVal lngFunction_OID As Long, Optional szFunction_name As String, Optional szFunction_arguments As String, Optional szFunction_returns As String, Optional szFunction_source As String, Optional szFunction_language As String, Optional szFunction_owner As String)
 On Error GoTo Err_Handler
    cmp_Function_GetValues lngFunction_OID, "pgadmin_functions", szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language, szFunction_owner
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Function_GetCurrentValues"
End Sub

Sub cmp_Function_GetDevValues(ByVal lngFunction_OID As Long, Optional szFunction_name As String, Optional szFunction_arguments As String, Optional szFunction_returns As String, Optional szFunction_source As String, Optional szFunction_language As String, Optional szFunction_owner As String)
 On Error GoTo Err_Handler
    cmp_Function_GetValues lngFunction_OID, "pgadmin_dev_functions", szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language, szFunction_owner
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Function_GetCurrentValues"
End Sub

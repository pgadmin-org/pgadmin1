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

Public Function cmp_Function_Exists(szFunction_PostgreSqlTable As String, ByVal szFunction_oid As Long, Optional ByVal szFunction_name As String, Optional ByVal szFunction_arguments As String) As Boolean
On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    ' Development  -> szFunction_PostgreSqlTable="pgadmin_dev_functions"
    ' Production   -> szFunction_PostgreSqlTable="pgadmin_functions"
    If (szFunction_PostgreSqlTable = "") Then
        szFunction_PostgreSqlTable = "pgadmin_functions"
    End If
    cmp_Function_Exists = False
        
    If szFunction_oid <> 0 Then
        szQueryStr = "SELECT * FROM " & szFunction_PostgreSqlTable
        szQueryStr = szQueryStr & " WHERE Function_OID = " & szFunction_oid
        
        ' Log
        LogMsg "Testing existence of function OID = " & szFunction_oid & "..."
    Else
        If szFunction_name <> "" Then
            szQueryStr = "SELECT * FROM " & szFunction_PostgreSqlTable
            szQueryStr = szQueryStr & " WHERE Function_name = '" & szFunction_name & "'"
            szQueryStr = szQueryStr & " AND Function_arguments = '" & szFunction_arguments & "'"
            'Log
            LogMsg "Testing existence of function " & szFunction_name & " (" & szFunction_arguments & ") in " & szFunction_PostgreSqlTable & "..."
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
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_Exists"
End Function

Public Sub cmp_Function_DropIfExists(szFunction_PostgreSqlTable As String, ByVal szFunction_oid As Long, Optional ByVal szFunction_name As String, Optional ByVal szFunction_arguments As String)
On Error GoTo Err_Handler
    Dim szDropStr As String
    
    ' Development  -> szFunction_PostgreSqlTable="pgadmin_dev_functions"
    ' Production   -> szFunction_PostgreSqlTable="pgadmin_functions"
    If (szFunction_PostgreSqlTable = "") Then
        szFunction_PostgreSqlTable = "pgadmin_functions"
    End If
    
    'Drop function if exists
    If cmp_Function_Exists(szFunction_PostgreSqlTable, szFunction_oid, szFunction_name & "", szFunction_arguments & "") = True Then
        ' Retrieve function name and arguments if we only know the OID
        If szFunction_oid <> 0 Then cmp_Function_GetValues szFunction_PostgreSqlTable, szFunction_oid, szFunction_name, szFunction_arguments
        
        ' create drop query
        If (szFunction_PostgreSqlTable = "pgadmin_functions") Then
            szDropStr = "DROP FUNCTION " & QUOTE & szFunction_name & QUOTE & " (" & szFunction_arguments & ");"
        Else
            szDropStr = "DELETE FROM " & szFunction_PostgreSqlTable & " WHERE "
            szDropStr = szDropStr & "function_name='" & szFunction_name & "' AND function_arguments='" & szFunction_arguments & "'"
        End If
        
        ' Log information
        LogMsg "Dropping function " & szFunction_name & " (" & szFunction_arguments & ") in " & szFunction_PostgreSqlTable & "..."
        LogMsg "Executing: " & szDropStr
        
        ' Execute drop query and close log
        gConnection.Execute szDropStr
        If (szFunction_PostgreSqlTable = "pgadmin_functions") Then LogQuery szDropStr
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_DropIfExists"
End Sub

Public Sub cmp_Function_Create(szFunction_PostgreSqlTable As String, ByVal szFunction_name As String, ByVal szFunction_arguments As String, ByVal szFunction_returns As String, ByVal szFunction_source As String, ByVal szFunction_language As String)
On Error GoTo Err_Handler
    Dim szCreateStr As String
    Dim szFunction_query_oid As Variant
    Dim szFunction_oid As Long
    
    ' Development  -> szFunction_PostgreSqlTable="pgadmin_dev_functions"
    ' Production   -> szFunction_PostgreSqlTable="pgadmin_functions"
    If (szFunction_PostgreSqlTable = "") Then
        szFunction_PostgreSqlTable = "pgadmin_functions"
    End If
    
    If (szFunction_PostgreSqlTable = "pgadmin_functions") Then
        szCreateStr = cmp_Function_CreateSQL(szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language)
    Else
        szFunction_source = Replace(szFunction_source, "'", "''")
        szFunction_source = Replace(szFunction_source, vbCrLf, "\n")
        
        szCreateStr = "INSERT INTO " & szFunction_PostgreSqlTable & " (function_name, Function_arguments, Function_returns, Function_source, Function_language)"
        szCreateStr = szCreateStr & "VALUES ("
        szCreateStr = szCreateStr & "'" & szFunction_name & "', "
        szCreateStr = szCreateStr & "'" & szFunction_arguments & "', "
        szCreateStr = szCreateStr & "'" & szFunction_returns & "', "
        szCreateStr = szCreateStr & "'" & szFunction_source & "', "
        szCreateStr = szCreateStr & "'" & szFunction_language & "' "
        szCreateStr = szCreateStr & ");"
    End If
    
    'Log
    LogMsg "Creating function " & szFunction_name & "(" & szFunction_arguments & ") in " & szFunction_PostgreSqlTable & "..."
    LogMsg "Executing: " & szCreateStr
    
    'Execute
    gConnection.Execute szCreateStr
    If (szFunction_PostgreSqlTable = "pgadmin_functions") Then LogQuery szCreateStr
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_Create"
  If Err.Number = -2147467259 Then MsgBox "Function " & szFunction_name & " (" & szFunction_arguments & ") could not be compiled." & vbCrLf & "Check source code and compile again."
  bContinueRebuilding = False
End Sub

Public Function cmp_Function_CreateSQL(ByVal szFunction_name As String, ByVal szFunction_argumentlist As String, ByVal szFunction_returns As String, ByVal szFunction_source As String, ByVal szFunction_language As String) As String
On Error GoTo Err_Handler
    Dim szCreateStr As String

    szFunction_source = Replace(szFunction_source, "'", "''")
    
    szCreateStr = "CREATE FUNCTION " & QUOTE & szFunction_name & "" & QUOTE & " ("
    szCreateStr = szCreateStr & szFunction_argumentlist & "" & ") " & vbCrLf
    szCreateStr = szCreateStr & "RETURNS " & szFunction_returns & " " & vbCrLf
    szCreateStr = szCreateStr & "AS '" & szFunction_source & vbCrLf
    szCreateStr = szCreateStr & "' " & vbCrLf
    szCreateStr = szCreateStr & "LANGUAGE '" & szFunction_language & "'"
    
    cmp_Function_CreateSQL = szCreateStr
  Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_CreateSQL"
End Function

Public Sub cmp_Function_Compile(ByVal szFunction_name As String, ByVal szFunction_arguments As String)
On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    Dim lngFunction_oid As Long
    Dim szFunction_returns As String
    Dim szFunction_language As String
    Dim szFunction_source As String
    
    ' Retrieve function
    lngFunction_oid = 0
    cmp_Function_GetValues "pgadmin_dev_functions", lngFunction_oid, szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language
 
    ' Compile function if exists
    If szFunction_name <> "" Then
    
        ' Attempt to create a temporary function to see if it compiles
        LogMsg "Checking if " & szFunction_name & " (" & szFunction_arguments & ") can be compiled ..."
        cmp_Function_DropIfExists "", 0, "pgadmin_fake__" & Left(szFunction_name, 15), szFunction_arguments
        cmp_Function_Create "", "pgadmin_fake__" & Left(szFunction_name, 15), szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language
 
        If (cmp_Function_Exists("", 0, "pgadmin_fake__" & Left(szFunction_name, 15), szFunction_arguments) = True) Then
            cmp_Function_DropIfExists "", 0, "pgadmin_fake__" & Left(szFunction_name, 15), szFunction_arguments
        Else
            bContinueRebuilding = False
            MsgBox "Function " & szFunction_name & "(" & szFunction_arguments & ") could not be compiled." & vbCrLf & "Check source code and rebuild project again.", vbOKOnly
        End If
        
       If bContinueRebuilding = True Then
            cmp_Function_DropIfExists "", 0, szFunction_name, szFunction_arguments
            cmp_Function_Create "", szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language
        
           ' Tell PgAdmin that the function was compiled
            If bContinueRebuilding = True Then
                cmp_Function_SetIsCompiled szFunction_name, szFunction_arguments
                LogMsg szFunction_name & " (" & szFunction_arguments & ") was successfuly compiled."
            End If
        End If
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_Compile"
  bContinueRebuilding = False
End Sub

Public Sub cmp_Function_Dependency_Initialize(ByVal szFunction_name As String)
On Error GoTo Err_Handler
    Dim szDependencyStr As String
    Dim rsComp As New Recordset
    
    ' Scan pgadmin_dev_functions for dependencies
    szDependencyStr = "SELECT * FROM pgadmin_dev_functions WHERE function_source ILIKE '%" & szFunction_name & "%'"
     
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szDependencyStr, gConnection, adOpenDynamic
  
    ' Write dependencies in pgadmin_dev_dependencies
    If Not rsComp.EOF Then
        szDependencyStr = "INSERT INTO pgadmin_dev_dependencies (dependency_to, dependency_from) "
        szDependencyStr = szDependencyStr & " SELECT '" & szFunction_name & "' AS dependency_to, function_name as dependency_from "
        szDependencyStr = szDependencyStr & " FROM pgadmin_dev_functions WHERE "
        szDependencyStr = szDependencyStr & " function_source ilike '%" & szFunction_name & "%'; "
        
        ' Log
        LogMsg "Executing: " & szDependencyStr
        
        gConnection.Execute szDependencyStr
    End If
    
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_Dependency_Initialize"
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
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_SetIsCompiled"
End Sub

Public Function cmp_Function_HasSatisfiedDependencies(ByVal szFunction_name As String) As Boolean
    On Error GoTo Err_Handler
    
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    ' Test existence of unsatisfied dependencies
    szQueryStr = "SELECT pgadmin_dev_functions.function_name, pgadmin_dev_functions.function_arguments, pgadmin_dev_functions_1.function_iscompiled"
    szQueryStr = szQueryStr & " From pgadmin_dev_functions"
    szQueryStr = szQueryStr & "    INNER JOIN pgadmin_dev_dependencies"
    szQueryStr = szQueryStr & "    ON pgadmin_dev_functions.Function_name = pgadmin_dev_dependencies.dependency_from"
    szQueryStr = szQueryStr & "    INNER JOIN pgadmin_dev_functions AS pgadmin_dev_functions_1"
    szQueryStr = szQueryStr & "    ON pgadmin_dev_dependencies.dependency_to =  pgadmin_dev_functions_1.Function_name"
    szQueryStr = szQueryStr & "    WHERE ((pgadmin_dev_functions.Function_name = '" & szFunction_name & "') AND (pgadmin_dev_functions_1.function_iscompiled = 'f'));"
    
    LogMsg "Testing existence of satisfied dependencies on function " & szFunction_name
    LogMsg "Executing: " & szQueryStr
  
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    
    cmp_Function_HasSatisfiedDependencies = False
    If rsComp.EOF Then
        cmp_Function_HasSatisfiedDependencies = True
    End If
    
    Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Function_HasSatisfiedDependencies"
End Function

Sub cmp_Function_GetValues(szFunction_PostgreSqlTable As String, lngFunction_oid As Long, Optional szFunction_name As String, Optional szFunction_arguments As String, Optional szFunction_returns As String, Optional szFunction_source As String, Optional szFunction_language As String, Optional szFunction_owner As String, Optional szFunction_comments As String)
 On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    If (szFunction_PostgreSqlTable = "") Then szFunction_PostgreSqlTable = "pgadmin_functions"

    ' Select query
    If lngFunction_oid <> 0 Then
        szQueryStr = "SELECT * from " & szFunction_PostgreSqlTable
        szQueryStr = szQueryStr & " WHERE function_OID = " & lngFunction_oid
    Else
        If IsMissing(szFunction_name) Then szFunction_name = ""
            szQueryStr = "SELECT * from " & szFunction_PostgreSqlTable
            szQueryStr = szQueryStr & " WHERE function_name = '" & szFunction_name & "'"
            If Not (IsMissing(szFunction_arguments)) Then
                szQueryStr = szQueryStr & " AND function_arguments = '" & szFunction_arguments & "'"
            End If
    End If
     
    ' open
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    
    If Not rsComp.EOF Then
        If IsNull(rsComp!function_oid) Then
            lngFunction_oid = 0
        Else
            lngFunction_oid = rsComp!function_oid
        End If
        If Not (IsMissing(szFunction_name)) Then szFunction_name = rsComp!function_name & ""
        If Not (IsMissing(szFunction_arguments)) Then szFunction_arguments = rsComp!Function_arguments & ""
        If Not (IsMissing(szFunction_returns)) Then szFunction_returns = rsComp!Function_returns & ""
        If Not (IsMissing(szFunction_source)) Then szFunction_source = rsComp!Function_source & ""
        If Not (IsMissing(szFunction_language)) Then szFunction_language = rsComp!Function_language & ""
        If Not (IsMissing(szFunction_owner)) Then szFunction_owner = rsComp!function_owner & ""
        If Not (IsMissing(szFunction_comments)) Then szFunction_comments = rsComp!function_comments & ""
       
        If (szFunction_name <> "") And (szFunction_returns = "") Then szFunction_returns = "opaque"
        rsComp.Close
    Else
        lngFunction_oid = 0
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

Public Sub cmp_Function_CopyToDev()
On Error GoTo Err_Handler
    Dim szQuery As String
    Dim rsComp As New Recordset
    Dim szFunction_source As String
    
    szQuery = "TRUNCATE TABLE pgadmin_dev_functions;" & _
    "  INSERT INTO pgadmin_dev_functions SELECT * " & _
    "  FROM pgadmin_functions " & _
    "  WHERE function_name NOT LIKE '%_call_handler' " & _
    "  AND function_name NOT LIKE 'pgadmin_%' " & _
    "  AND function_name NOT LIKE 'pg_%' " & _
    "  AND function_oid > " & LAST_SYSTEM_OID & _
    "  ORDER BY function_oid ;" & _
    "  UPDATE pgadmin_dev_functions SET function_iscompiled = 't';" & _
    "  UPDATE pgadmin_dev_functions SET function_returns = 'opaque' WHERE function_returns = NULL;" & _
    "  TRUNCATE TABLE pgadmin_dev_dependencies;"
    
    LogMsg "Copying pgadmin_functions to pgadmin_dev_functions..."
    LogMsg "Executing: " & szQuery
    gConnection.Execute szQuery
        
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basFunction, cmp_Func_CopyToDev"
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

Public Sub cmp_Function_DropAll(Optional szFunction_PostgreSqlTable As String)
On Error GoTo Err_Handler
    Dim szQuery As String
    Dim szFunc() As Variant
    Dim iLoop As Long
    Dim iUbound As Long
    Dim rsFunc As New Recordset
    Dim szFunction_name As String
    Dim szFunction_arguments As String
    
    If IsMissing(szFunction_PostgreSqlTable) Or (szFunction_PostgreSqlTable = "") Then szFunction_PostgreSqlTable = "pgadmin_functions"
        
    If (szFunction_PostgreSqlTable = "pgadmin_functions") Then
        szQuery = " SELECT function_name, function_arguments " & _
        "  FROM pgadmin_functions " & _
        "  WHERE function_name NOT LIKE '%_call_handler' " & _
        "  AND function_name NOT LIKE 'pgadmin_%' " & _
        "  AND function_name NOT LIKE 'pg_%' " & _
        "  AND function_oid > " & LAST_SYSTEM_OID & _
        "  ORDER BY function_oid ;"
        
        LogMsg "Dropping all functions in pgadmin_functions..."
        LogMsg "Executing: " & szQuery
        
        If rsFunc.State <> adStateClosed Then rsFunc.Close
        rsFunc.Open szQuery, gConnection, adOpenForwardOnly, adLockReadOnly
    
        If Not (rsFunc.EOF) Then
            szFunc = rsFunc.GetRows
            rsFunc.Close
            iUbound = UBound(szFunc, 2)
                For iLoop = 0 To iUbound
                     szFunction_name = szFunc(0, iLoop)
                     szFunction_arguments = szFunc(1, iLoop)
                     cmp_Function_DropIfExists "", 0, szFunction_name, szFunction_arguments
                Next iLoop
            Erase szFunc
        End If
    Else
        szQuery = "TRUNCATE " & szFunction_PostgreSqlTable
        LogMsg "Truncate " & szFunction_PostgreSqlTable & "..."
        LogMsg "Executing: " & szQuery
        gConnection.Execute szQuery
    End If
   
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basProject, cmp_Function_DropAll"
End Sub

Attribute VB_Name = "basCompiler"
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
'**** Views
'****

Sub cmp_View_DropIfExists(ByVal lngView_OID As Long, Optional ByVal szView_Name As String)
 On Error GoTo Err_Handler
    Dim szDropStr As String
  
    ' Test existence of view
    If cmp_View_Exists(lngView_OID, szView_Name & "") = True Then
    
        If szView_Name = "" Then cmp_View_GetValues lngView_OID, "", szView_Name
    
        ' create drop query
        szDropStr = "DROP VIEW " & QUOTE & szView_Name & QUOTE
               
        ' Log information
        LogMsg "Dropping view " & szView_Name & "..."
        LogMsg "Executing: " & szDropStr
        
        ' Execute drop query and close log
        gConnection.Execute szDropStr
        LogQuery szDropStr
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_View_DropIfExists"
End Sub

Function cmp_View_Exists(ByVal lngView_OID As Long, ByVal szView_Name As String) As Boolean
 On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
  
    cmp_View_Exists = False
    If lngView_OID <> 0 Then
        szQueryStr = "SELECT * FROM pgadmin_views "
        szQueryStr = szQueryStr & "WHERE view_OID = " & Str(lngView_OID)
    Else
        If szView_Name <> "" Then
            szQueryStr = "SELECT * FROM pgadmin_views "
            szQueryStr = szQueryStr & "WHERE view_name = '" & szView_Name & "' "
        Else
            Exit Function
        End If
    End If
    
      ' retrieve name and arguments of function to drop
    LogMsg "Testing existence of view " & szView_Name & "..."
    LogMsg "Executing: " & szQueryStr

    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    

    If Not rsComp.EOF Then
        cmp_View_Exists = True
        rsComp.Close
    End If
  Exit Function
  
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_View_DropIfExists"
End Function

Sub cmp_View_Create(ByVal szView_Name As String, ByVal szView_Definition As String)
On Error GoTo Err_Handler
  Dim szCreateStr As String

    szCreateStr = cmp_View_CreateSQL(szView_Name, szView_Definition)
    LogMsg "Creating view " & szView_Name & "..."
    LogMsg "Executing: " & szCreateStr
    
    ' Execute drop query and close log
    gConnection.Execute szCreateStr
    LogQuery szCreateStr

  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Views_Create"
  If Err.Number = -2147467259 Then MsgBox "View " & szView_Name & " could not be compiled." & vbCrLf & "Check source code and compile again."
  bContinueCompilation = False
End Sub

Function cmp_View_CreateSQL(ByVal szView_Name As String, ByVal szView_Definition As String) As String
On Error GoTo Err_Handler
  Dim szQuery As String

    szQuery = "CREATE VIEW " & szView_Name & " AS " & szView_Definition & "; "
    cmp_View_CreateSQL = szQuery
  Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Views_Create"
End Function

Sub cmp_View_GetValues(ByVal lngView_OID As Long, Optional szView_PostgreSQLtable As String, Optional szView_Name As String, Optional szView_Definition As String, Optional szView_Owner As String, Optional szView_Acl As String)
 On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    ' Where should we get the values ?
    If IsMissing(szView_PostgreSQLtable) Or (szView_PostgreSQLtable = "") Then
        szView_PostgreSQLtable = "pgadmin_views"
    End If
        
    ' Select query
    If lngView_OID <> 0 Then
        szQueryStr = "SELECT * from " & szView_PostgreSQLtable
        szQueryStr = szQueryStr & " WHERE view_OID = " & lngView_OID
        LogMsg "Retrieving values from view OID =" & lngView_OID & "..."
    Else
        ' to be written
        Exit Sub
    End If
    
    LogMsg "Executing: " & szQueryStr
    
    ' open
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    
    If Not rsComp.EOF Then
        If Not (IsMissing(szView_Name)) Then szView_Name = rsComp!view_name & ""
        If Not (IsMissing(szView_Owner)) Then szView_Owner = rsComp!view_owner & ""
        If Not (IsMissing(szView_Acl)) Then szView_Acl = rsComp!view_acl & ""
        If Not (IsMissing(szView_Definition)) Then szView_Definition = cmp_View_GetViewDef(szView_Name)
        rsComp.Close
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Views_GetValues"
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
    
    If Not rsTemp.EOF Then
        cmp_View_GetViewDef = rsTemp!Result
    End If
    
    Exit Function
Err_Handler:
  cmp_View_GetViewDef = "Not a view"
End Function


'****
'**** Triggers
'****
'****

Function cmp_Trigger_CreateSQL(ByVal szTrigger_name As String, ByVal szTrigger_table As String, ByVal szTrigger_function As String, ByVal szTrigger_arguments As String, ByVal szTrigger_ForEach As String, ByVal szTrigger_Executes As String, ByVal szTrigger_event As String, Optional iTrigger_type As Integer) As String
' Two syntaxes
' cmp_Trigger_Create (szTrigger_name, szTrigger_table, szTrigger_function, szTrigger_ForEach, szTrigger_Executes, szTrigger_Event )
' cmp_Trigger_Create (szTrigger_name, szTrigger_table, szTrigger_function, "", "", "", szTrigger_type)

On Error GoTo Err_Handler
    Dim szQueryStr As String

    ' if trigger_type defined
    If Not (IsMissing(iTrigger_type)) Then
        If iTrigger_type <> 0 Then
            ' retrieve values from trigger
            
            If (iTrigger_type And 1) = 1 Then
              szTrigger_ForEach = " Row"
            Else
              szTrigger_ForEach = " Statement"
            End If
            
            If (iTrigger_type And 2) = 2 Then
              szTrigger_Executes = " Before"
            Else
              szTrigger_Executes = " After"
            End If
            
            If (iTrigger_type And 4) = 4 Then szTrigger_event = szTrigger_event & "Insert OR "
            If (iTrigger_type And 8) = 8 Then szTrigger_event = szTrigger_event & "Delete OR "
            If (iTrigger_type And 16) = 16 Then szTrigger_event = szTrigger_event & "Update OR "
            szTrigger_event = Left(szTrigger_event, Len(szTrigger_event) - 3)
        End If
    End If
     
    szQueryStr = "CREATE TRIGGER " & QUOTE & szTrigger_name & QUOTE
    szQueryStr = szQueryStr & " " & szTrigger_Executes & " " & szTrigger_event
    szQueryStr = szQueryStr & " ON " & QUOTE & szTrigger_table & QUOTE & " FOR EACH " & szTrigger_ForEach
    szQueryStr = szQueryStr & " EXECUTE PROCEDURE " & szTrigger_function & "(" & szTrigger_arguments & ")"
    
    cmp_Trigger_CreateSQL = szQueryStr
    Exit Function
    
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Trigger_CreateSQL"
End Function

Sub cmp_Trigger_Create(ByVal szTrigger_name As String, ByVal szTrigger_table As String, ByVal szTrigger_function As String, ByVal szTrigger_arguments As String, ByVal szTrigger_ForEach As String, ByVal szTrigger_Executes As String, ByVal szTrigger_event As String, Optional iTrigger_type As Integer)
' Two syntaxes
' cmp_Trigger_Create (szTrigger_name, szTrigger_table, szTrigger_function, szTrigger_ForEach, szTrigger_Executes, szTrigger_Event )
' cmp_Trigger_Create (szTrigger_name, szTrigger_table, szTrigger_function, "", "", "", szTrigger_type)
    Dim szQueryStr As String
    
    If (IsMissing(iTrigger_type)) Then
      szQueryStr = cmp_Trigger_CreateSQL(szTrigger_name, szTrigger_table, szTrigger_function, szTrigger_arguments, szTrigger_ForEach, szTrigger_Executes, szTrigger_event)
    Else
      szQueryStr = cmp_Trigger_CreateSQL(szTrigger_name, szTrigger_table, szTrigger_function, szTrigger_arguments, szTrigger_ForEach, szTrigger_Executes, szTrigger_event, iTrigger_type)
    End If
    
    ' Log information
    LogMsg "Creating trigger " & szTrigger_name & "..."
    LogMsg "Executing: " & szQueryStr
      
    ' Execute drop query and close log
    gConnection.Execute szQueryStr
    LogQuery szQueryStr
      
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Trigger_SQL"
  If Err.Number = -2147467259 Then MsgBox "Trigger " & szTrigger_name & " could not be compiled." & vbCrLf & "Check source code and compile again."
  bContinueCompilation = False
End Sub

Sub cmp_Trigger_DropIfExists(ByVal lngTrigger_OID As Long, Optional ByVal szTrigger_name As String, Optional ByVal szTrigger_table As String)
 On Error GoTo Err_Handler
    Dim szDropStr As String
    
    ' Test existence of trigger
    If cmp_Trigger_Exists(lngTrigger_OID, szTrigger_name & "", szTrigger_table & "") Then
        ' Retrieve name and table is we only know the OID
        If lngTrigger_OID <> 0 And ((szTrigger_name = "") Or (szTrigger_table = "")) Then cmp_Trigger_GetValues lngTrigger_OID, "", szTrigger_name, szTrigger_table
        
        ' Create drop query
        szDropStr = "DROP TRIGGER " & QUOTE & szTrigger_name & QUOTE & " ON " & szTrigger_table
               
        ' Log information
        LogMsg "Dropping trigger " & szTrigger_name & " on table " & szTrigger_table & "..."
        LogMsg "Executing: " & szDropStr
        
        ' Execute drop query and close log
        gConnection.Execute szDropStr
        LogQuery szDropStr
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Trigger_DropIfExists"
End Sub

Sub cmp_Trigger_GetValues(ByVal lngTrigger_OID As Long, Optional szTrigger_PostgreSQLtable As String, Optional szTrigger_name As String, Optional szTrigger_table As String, Optional szTrigger_function As String, Optional szTrigger_arguments As String, Optional szTrigger_ForEach As String, Optional szTrigger_Executes As String, Optional szTrigger_event As String)
 On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    Dim iTrigger_type As Integer
    
    ' Where should we get the values ?
    If IsMissing(szTrigger_PostgreSQLtable) Or (szTrigger_PostgreSQLtable = "") Then
        szTrigger_PostgreSQLtable = "pgadmin_triggers"
    End If
        
    ' Select query
    If lngTrigger_OID <> 0 Then
        
        szQueryStr = "SELECT * from " & szTrigger_PostgreSQLtable
        szQueryStr = szQueryStr & " WHERE trigger_OID = " & lngTrigger_OID
        LogMsg "Retrieving name and table from trigger OID =" & lngTrigger_OID & "..."
    Else
        ' to be written
        Exit Sub
    End If
    
    LogMsg "Executing: " & szQueryStr
    
    ' open
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    
    If Not rsComp.EOF Then
        If Not (IsMissing(szTrigger_name)) Then szTrigger_name = rsComp!trigger_name & ""
        If Not (IsMissing(szTrigger_table)) Then szTrigger_table = rsComp!trigger_table & ""
        If Not (IsMissing(szTrigger_function)) Then szTrigger_function = rsComp!trigger_function & ""
        If Not (IsMissing(szTrigger_arguments)) Then szTrigger_arguments = rsComp!trigger_arguments & ""
        iTrigger_type = rsComp!trigger_type
        If iTrigger_type <> 0 Then
            If Not (IsMissing(szTrigger_ForEach)) Then
                If (iTrigger_type And 1) = 1 Then
                  szTrigger_ForEach = "Row"
                Else
                  szTrigger_ForEach = "Statement"
                End If
            End If
            
            If Not (IsMissing(szTrigger_Executes)) Then
                If (iTrigger_type And 2) = 2 Then
                  szTrigger_Executes = "Before"
                Else
                  szTrigger_Executes = "After"
                End If
            End If
            
            If Not (IsMissing(szTrigger_event)) Then
                If (iTrigger_type And 4) = 4 Then szTrigger_event = szTrigger_event & "Insert OR "
                If (iTrigger_type And 8) = 8 Then szTrigger_event = szTrigger_event & "Delete OR "
                If (iTrigger_type And 16) = 16 Then szTrigger_event = szTrigger_event & "Update OR "
                szTrigger_event = Left(szTrigger_event, Len(szTrigger_event) - 3)
            End If
        End If
        rsComp.Close
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Trigger_GetValues"
End Sub

Function cmp_Trigger_Exists(ByVal lngTrigger_OID As Long, Optional ByVal szTrigger_name As String, Optional ByVal szTrigger_table As String) As Boolean
 On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
  
    cmp_Trigger_Exists = False
    
    If lngTrigger_OID <> 0 Then
        szQueryStr = "SELECT * FROM pgadmin_triggers"
        szQueryStr = szQueryStr & " WHERE Trigger_OID = " & lngTrigger_OID
        
        ' Logging
        LogMsg "Testing existence of trigger OID = " & lngTrigger_OID & "..."
    Else
        If szTrigger_table <> "" And szTrigger_name <> "" Then
            szQueryStr = "SELECT * FROM pgadmin_triggers"
            szQueryStr = szQueryStr & " WHERE Trigger_name = '" & szTrigger_name & "'"
            szQueryStr = szQueryStr & " AND Trigger_table = '" & szTrigger_table & "'"
            
            ' Logging
            LogMsg "Testing existence of trigger " & szTrigger_name & " on table " & szTrigger_table & "..."
        Else
            Exit Function
        End If
    End If
    
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
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Trigger_DropIfExists"
End Function

'****
'**** Function
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
    szCreateStr = szCreateStr & szFunction_argumentlist & "" & ") "
    szCreateStr = szCreateStr & "RETURNS " & szFunction_returns & " "
    szCreateStr = szCreateStr & "AS '" & szFunction_source & "' "
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

Sub cmp_Function_GetValues(ByVal lngFunction_OID As Long, Optional szFunction_PostgreSQLtable As String, Optional szFunction_name As String, Optional szFunction_arguments As String, Optional szFunction_returns As String, Optional szFunction_source As String, Optional szFunction_language As String, Optional szFunction_owner As String)
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
        If IsMissing(szFunction_name) Or IsMissing(szFunction_name) Then Exit Sub
        If szFunction_name <> "" Then
            szQueryStr = "SELECT * from " & szFunction_PostgreSQLtable
            szQueryStr = szQueryStr & " WHERE function_name = '" & szFunction_name & "'"
            If szFunction_arguments <> "" Then
                szQueryStr = szQueryStr & " AND function_arguments = '" & szFunction_arguments & "'"
            End If
        End If
    End If
    
    ' Log information
    LogMsg "Retrieving information from function OID =" & lngFunction_OID & " in table " & szFunction_PostgreSQLtable & "..."
    LogMsg "Executing: " & szQueryStr
    
    ' open
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    
    If Not rsComp.EOF Then
        If Not (IsMissing(szFunction_name)) Then szFunction_name = rsComp!Function_name & ""
        If Not (IsMissing(szFunction_arguments)) Then szFunction_arguments = rsComp!Function_arguments & ""
        If Not (IsMissing(szFunction_returns)) Then szFunction_returns = rsComp!Function_returns & ""
        If Not (IsMissing(szFunction_source)) Then szFunction_source = rsComp!Function_source & ""
        If Not (IsMissing(szFunction_language)) Then szFunction_language = rsComp!Function_language & ""
        If Not (IsMissing(szFunction_owner)) Then szFunction_owner = rsComp!function_owner & ""
       
        If (szFunction_name <> "") And (szFunction_returns = "") Then szFunction_returns = "opaque"
        szFunction_source = Replace(szFunction_source, "'", "''")
        rsComp.Close
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

'****
'**** Table
'****
'****
Public Sub cmp_Table_DropIfExists(ByVal szTable_name As String)
    On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    If cmp_Table_Exists(szTable_name) Then
        szQueryStr = "DROP TABLE " & QUOTE & szTable_name & QUOTE
        
        'Log
        LogMsg "Dropping table " & szTable_name
        LogMsg "Executing: " & szQueryStr
        
        gConnection.Execute szQueryStr
        LogQuery szQueryStr
    End If
    
      Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Table_DropIfExists"
End Sub

Public Function cmp_Table_Exists(ByVal szTable_name As String) As Boolean
    On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    szQueryStr = "SELECT * FROM pgadmin_tables WHERE Table_name = '" & szTable_name & "'"
    ' Log
    LogMsg "Testing existence of table " & szTable_name & "..."
    LogMsg "Executing: SELECT * FROM pgadmin_tables WHERE Table_name = " & szTable_name
  
    ' Test existence of the table
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection, adOpenDynamic
    
    cmp_Table_Exists = False
    If Not rsComp.EOF Then
        cmp_Table_Exists = True
    End If
    
      Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Table_Exists"
End Function

'****
'**** Project
'****
'****

Public Sub comp_Project_BackupViews(Optional ByVal Function_OldName, Optional ByVal Function_NewName, Optional ByVal Table_OldName, Optional ByVal Table_NewName)
'On Error GoTo Err_Handler
    Dim szQuery As String
    Dim szView_Definition As String
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
    
    While Not rsComp.EOF
        'Copy view definition
        szView_Definition = Replace(cmp_View_GetViewDef(rsComp!view_name), "'", "''")
        
        ' Rename underlying functions if needed
        If Not (IsMissing(Function_NewName)) And Not (IsMissing(Function_OldName)) Then
            If (Function_OldName <> "") And (Function_NewName <> "") And (Function_NewName <> Function_OldName) Then
                szView_Definition = Replace(szView_Definition, Function_OldName, Function_NewName)
            End If
        End If
        
        ' Rename underlying table if needed
        If Not (IsMissing(Table_NewName)) And Not (IsMissing(Table_OldName)) Then
            If (Table_OldName <> "") And (Table_NewName <> "") And (Table_NewName <> Table_OldName) Then
                szView_Definition = Replace(szView_Definition, Table_OldName, Table_NewName)
            End If
        End If
        
        ' Update definition of view
        szQuery = "UPDATE pgadmin_dev_views SET view_definition = '" & szView_Definition & "' WHERE view_oid = '" & rsComp!view_oid & "'"
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
    "  AND trigger_name NOT LIKE 'RI_ConstraintTrigger_%' " & _
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
        cmp_Function_Dependency_Initialize rsComp!function_OID, rsComp!Function_name
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
                szQuery = szQuery & " WHERE function_oid = " & Str(rsComp!function_OID)
                
                LogMsg "Updating function_source in " & Function_OldName & " with " & szFunction_source & "..."
                LogMsg "Executing: " & szQuery
                
                gConnection.Execute szQuery
                rsComp.MoveNext
            Wend
        End If
    End If
        
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, comp_Project_Initialize"
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
        If cmp_Function_HasSatisfiedDependencies(rsComp!function_OID) = True Then
            comp_Project_FindNextFunctionToCompile = rsComp!function_OID
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

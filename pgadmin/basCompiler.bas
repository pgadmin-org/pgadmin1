Attribute VB_Name = "basCompiler"
Option Explicit
Dim bContinueCompilation As Boolean

'****
'**** Views
'****

Sub cmp_View_DropIfExists(ByVal szView_name As String)
 'On Error GoTo Err_Handler
    Dim szDropStr As String
  
    ' Test existence of view
    If cmp_View_Exists(szView_name) = True Then
        ' create drop query
        szDropStr = "DROP VIEW " & QUOTE & szView_name & QUOTE
               
        ' Log information
        LogMsg "Dropping view " & szView_name & "..."
        LogMsg "Executing: " & szDropStr
        
        ' Execute drop query and close log
        gConnection.Execute szDropStr
        LogQuery szDropStr
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_View_DropIfExists"
End Sub

Function cmp_View_Exists(ByVal szView_name As String) As Boolean
 'On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
  
    szQueryStr = "SELECT * FROM pgadmin_dev_views "
    szQueryStr = szQueryStr & "WHERE view_name = '" & szView_name & "' "
    
      ' retrieve name and arguments of function to drop
    LogMsg "Testing existence of view " & szView_name & "..."
    LogMsg "Executing: " & szQueryStr

    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
    
    cmp_View_Exists = False
    If Not rsComp.EOF Then
        cmp_View_Exists = True
        rsComp.Close
    End If
  Exit Function
  
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_View_DropIfExists"
End Function

Sub cmp_View_Create(ByVal szView_name As String, ByVal szView_definition As String)
'On Error GoTo Err_Handler
    Dim szCreateStr As String

  szCreateStr = "CREATE VIEW " & szView_name & " AS " & szView_definition
  
      ' Log information
  LogMsg "Creating view " & szView_name & "..."
  LogMsg "Executing: " & szCreateStr
    
    ' Execute drop query and close log
    gConnection.Execute szCreateStr
    LogQuery szCreateStr
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Views_Create"
End Sub

'****
'**** Triggers
'****
'****

Sub cmp_Trigger_Create(ByVal szTrigger_name As String, ByVal szTrigger_table As String, ByVal szTrigger_function As String, ByVal szTrigger_arguments As String, ByVal szTrigger_ForEach As String, ByVal szTrigger_Executes As String, ByVal szTrigger_Event As String, Optional iTrigger_type As Integer)
' Two syntaxes
' cmp_Trigger_Create (szTrigger_name, szTrigger_table, szTrigger_function, szTrigger_ForEach, szTrigger_Executes, szTrigger_Event )
' cmp_Trigger_Create (szTrigger_name, szTrigger_table, szTrigger_function, "", "", "", szTrigger_type)

'On Error GoTo Err_Handler
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
            
            If (iTrigger_type And 4) = 4 Then szTrigger_Event = szTrigger_Event & "Insert OR "
            If (iTrigger_type And 8) = 8 Then szTrigger_Event = szTrigger_Event & "Delete OR "
            If (iTrigger_type And 16) = 16 Then szTrigger_Event = szTrigger_Event & "Update OR "
            szTrigger_Event = Left(szTrigger_Event, Len(szTrigger_Event) - 3)
        End If
    End If
  
    szQueryStr = "CREATE TRIGGER " & QUOTE & szTrigger_name & QUOTE
    szQueryStr = szQueryStr & " " & szTrigger_Executes & " " & szTrigger_Event
    szQueryStr = szQueryStr & " ON " & QUOTE & szTrigger_table & QUOTE & " FOR EACH " & szTrigger_ForEach
    szQueryStr = szQueryStr & " EXECUTE PROCEDURE " & szTrigger_function & "(" & szTrigger_arguments & ")"
    
        ' Log information
    LogMsg "Creating trigger " & szTrigger_name & "..."
    LogMsg "Executing: " & szQueryStr
      
      ' Execute drop query and close log
      gConnection.Execute szQueryStr
      LogQuery szQueryStr
      
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Trigger_SQL"
End Sub

Sub cmp_Trigger_DropIfExists(ByVal lngTrigger_OID As Long, Optional ByVal szTrigger_name As String, Optional ByVal szTrigger_table As String)
 'On Error GoTo Err_Handler
    Dim szDropStr As String
    
    ' Test existence of trigger
    If cmp_Trigger_Exists(lngTrigger_OID, szTrigger_name & "", szTrigger_table & "") Then
        ' Retrieve name and table is we only know the OID
        If lngTrigger_OID <> 0 Then cmp_Trigger_GetInfo lngTrigger_OID, szTrigger_name, szTrigger_table
        
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

Sub cmp_Trigger_GetInfo(ByVal lngTrigger_OID As Long, Optional szTrigger_name As String, Optional szTrigger_table As String, Optional szTrigger_function As String, Optional szTrigger_arguments As String, Optional szTrigger_ForEach As String, Optional szTrigger_Executes As String, Optional szTrigger_Event As String)
 'On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    Dim iTrigger_type As Integer
    
    If lngTrigger_OID <> 0 Then
        ' create drop query
        szQueryStr = "SELECT * from pgadmin_triggers"
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
        If Not (IsMissing(szTrigger_name)) Then szTrigger_name = rsComp!trigger_name
        If Not (IsMissing(szTrigger_table)) Then szTrigger_table = rsComp!trigger_table
        If Not (IsMissing(szTrigger_function)) Then szTrigger_function = rsComp!trigger_function
        If Not (IsMissing(szTrigger_arguments)) Then szTrigger_arguments = rsComp!trigger_arguments
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
            
            If Not (IsMissing(szTrigger_Event)) Then
                If (iTrigger_type And 4) = 4 Then szTrigger_Event = szTrigger_Event & "Insert OR "
                If (iTrigger_type And 8) = 8 Then szTrigger_Event = szTrigger_Event & "Delete OR "
                If (iTrigger_type And 16) = 16 Then szTrigger_Event = szTrigger_Event & "Update OR "
                szTrigger_Event = Left(szTrigger_Event, Len(szTrigger_Event) - 3)
            End If
        End If
        rsComp.Close
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Trigger_GetInfo"
End Sub

Function cmp_Trigger_Exists(ByVal lngTrigger_OID As Long, Optional ByVal szTrigger_name As String, Optional ByVal szTrigger_table As String) As Boolean
 'On Error GoTo Err_Handler
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
'On Error GoTo Err_Handler
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
            szQueryStr = szQueryStr & "AND Function_arguments = '" & szFunction_arguments & "'"
            
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
'On Error GoTo Err_Handler
    Dim szDropStr As String
    
    'Drop function if exists
    If cmp_Function_Exists(szFunction_OID, szFunction_name & "", szFunction_arguments & "") = True Then
        ' Retrieve function name and arguments if we only know the OID
        If szFunction_OID <> 0 Then cmp_Function_GetInfo szFunction_OID, szFunction_name, szFunction_arguments
        
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
'On Error GoTo Err_Handler
    Dim szCreateStr As String

  szCreateStr = "CREATE FUNCTION " & QUOTE & szFunction_name & "" & QUOTE & " ("
  szCreateStr = szCreateStr & szFunction_argumentlist & "" & ") "
  szCreateStr = szCreateStr & "RETURNS " & szFunction_returns & " "
  szCreateStr = szCreateStr & "AS '" & szFunction_source & "' "
  szCreateStr = szCreateStr & "LANGUAGE '" & szFunction_language & "'"
  
  'Log
  LogMsg "Creating function " & szFunction_name & "(" & szFunction_argumentlist & ") ..."
  LogMsg "Executing: " & szCreateStr
  
  'Execute
  gConnection.Execute szCreateStr
  LogQuery szCreateStr
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Function_Create"
  MsgBox "Function " & szFunction_name & " (" & szFunction_argumentlist & ") could not be compiled. Check source code of the function and compile again."
  bContinueCompilation = False
End Sub

Public Sub cmp_Function_Compile(ByVal lngFunction_OID As Long)
'On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    Dim szFunction_name As String
    Dim szFunction_arguments As String
    Dim szFunction_returns As String
    Dim szFunction_language As String
    Dim szFunction_source As String

    ' Retrive latest version of function
    szQueryStr = "SELECT * FROM pgadmin_functions WHERE function_OID = " & Str(lngFunction_OID)
    
     'Log
    LogMsg "Retrieving latest version of function OID = " & lngFunction_OID & " ..."
    LogMsg "Executing: " & szQueryStr
  
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQueryStr, gConnection
  
    ' Compile function if exists
    If Not rsComp.EOF Then
        szFunction_name = rsComp!Function_name & ""
        szFunction_arguments = rsComp!Function_arguments & ""
        szFunction_returns = rsComp!Function_returns & ""
        If szFunction_returns = "" Then szFunction_returns = "opaque" 'strange
        szFunction_language = rsComp!Function_language & ""
        szFunction_source = Replace(rsComp!Function_source, "'", "''") & ""
        
        ' Attempt to create a temporary function to see if it compiles
        LogMsg "Checking if " & szFunction_name & " (" & szFunction_arguments & ") can be compiled ..."
        cmp_Function_DropIfExists 0, "pgadmin_dev_temp_function", szFunction_arguments
        cmp_Function_Create "pgadmin_dev_temp_function", szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language
    
       If bContinueCompilation = True Then
            ' If it does, compile the real function
            cmp_Function_DropIfExists 0, szFunction_name, szFunction_arguments
            cmp_Function_Create szFunction_name, szFunction_arguments, szFunction_returns, szFunction_source, szFunction_language
        
           ' Tell PgAdmin that the function was compiled
            cmp_Function_SetIsCompiled szFunction_name, szFunction_arguments
            LogMsg szFunction_name & " (" & szFunction_arguments & ") was successfuly compiled."
        End If
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Function_Compile"
  MsgBox "Function " & szFunction_name & " (" & szFunction_arguments & ") could not be compiled. Check source code of the function and compile again."
  bContinueCompilation = False
End Sub


Public Sub cmp_Function_Dependency_Initialize(ByVal lngFunction_OID As Long, ByVal szFunction_name As String)
'On Error GoTo Err_Handler
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
'On Error GoTo Err_Handler
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
'On Error GoTo Err_Handler
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
    'On Error GoTo Err_Handler
    
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

Sub cmp_Function_GetInfo(ByVal lngFunction_OID As Long, Optional szFunction_name As String, Optional szFunction_arguments As String, Optional szFunction_returns As String, Optional szFunction_source As String, Optional szFunction_language As String)
 'On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    If lngFunction_OID <> 0 Then
            ' create drop query
        szQueryStr = "SELECT * from pgadmin_functions"
        szQueryStr = szQueryStr & " WHERE function_OID = " & lngFunction_OID
               
        ' Log information
        LogMsg "Retrieving information from function OID =" & lngFunction_OID & "..."
        LogMsg "Executing: " & szQueryStr
        
        ' open
        If rsComp.State <> adStateClosed Then rsComp.Close
        rsComp.Open szQueryStr, gConnection
        
        If Not rsComp.EOF Then
            If Not (IsMissing(szFunction_name)) Then szFunction_name = rsComp!Function_name
            If Not (IsMissing(szFunction_arguments)) Then szFunction_arguments = rsComp!Function_arguments
            If Not (IsMissing(szFunction_returns)) Then szFunction_returns = rsComp!Function_returns
            If Not (IsMissing(szFunction_source)) Then szFunction_source = rsComp!Function_source
            If Not (IsMissing(szFunction_language)) Then szFunction_language = rsComp!Function_language
            rsComp.Close
        End If
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, cmp_Function_GetInfo"
End Sub
'****
'**** Table
'****
'****
Public Sub cmp_Table_DropIfExists(ByVal szTable_name As String)
    'On Error GoTo Err_Handler
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
    'On Error GoTo Err_Handler
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

Public Sub comp_Project_Initialize()
'On Error GoTo Err_Handler
    Dim InitializeStr As String
    Dim rsComp As New Recordset
    
    Chk_HelperObjects
    
    ' Fill pgadmin_dev_dependencies table
    InitializeStr = "SELECT * FROM pgadmin_dev_functions ORDER BY function_OID"
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open InitializeStr, gConnection, adOpenDynamic
    
    While Not rsComp.EOF
        cmp_Function_Dependency_Initialize rsComp!function_OID, rsComp!Function_name
        rsComp.MoveNext
    Wend
    
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, comp_Project_Initialize"
End Sub

Public Function comp_Project_FindNextFunctionToCompile() As Long
'On Error GoTo Err_Handler
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

Public Sub comp_Project_RelinkTriggers()
'On Error GoTo Err_Handler
    Dim rsTrigger As New Recordset
    Dim szQueryStr As String
    ' Obviously this does not work
    
    szQueryStr = "SELECT * From pgadmin_dev_triggers"
    
    LogMsg "Now relinking triggers..."
    LogMsg "Executing: " & szQueryStr
    
    If rsTrigger.State <> adStateClosed Then rsTrigger.Close
    rsTrigger.Open szQueryStr, gConnection, adOpenDynamic
    
    While Not rsTrigger.EOF
        ' Drop trigger if exists and then recreate it
        cmp_Trigger_DropIfExists rsTrigger!trigger_OID, rsTrigger!trigger_name, rsTrigger!trigger_table
        cmp_Trigger_Create rsTrigger!trigger_name, rsTrigger!trigger_table, rsTrigger!trigger_function, rsTrigger!trigger_arguments, "", "", "", rsTrigger!trigger_type
        rsTrigger.MoveNext
    Wend

    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, comp_Project_RelinkTriggers"
End Sub

Public Sub comp_Project_RelinkViews()
'On Error GoTo Err_Handler
    Dim rsViews As New Recordset
    Dim szQueryStr As String
    ' Obviously this does not work
    
    szQueryStr = "SELECT * From pgadmin_dev_views"
    
    LogMsg "Now relinking views..."
    LogMsg "Executing: " & szQueryStr
    
    If rsViews.State <> adStateClosed Then rsViews.Close
    rsViews.Open szQueryStr, gConnection, adOpenDynamic
    
    While Not rsViews.EOF
        ' Drop view if exists and then recreate it
        cmp_View_DropIfExists rsViews!view_name
        cmp_View_Create rsViews!view_name, rsViews!view_definition
        rsViews.MoveNext
    Wend

    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, comp_Project_RelinkTriggers"
End Sub

Public Sub comp_Project_Compile()
'On Error GoTo Err_Handler
    Dim lngNextFunctionToCompile_OID As Long
    
    bContinueCompilation = True
    lngNextFunctionToCompile_OID = comp_Project_FindNextFunctionToCompile
    While (lngNextFunctionToCompile_OID > 0) And (bContinueCompilation = True)
        cmp_Function_Compile (lngNextFunctionToCompile_OID)
        lngNextFunctionToCompile_OID = comp_Project_FindNextFunctionToCompile
    Wend
    
    comp_Project_RelinkTriggers
    comp_Project_RelinkViews
    
    If bContinueCompilation = True Then MsgBox ("Rebuilding successfull")
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basCompiler, comp_Project_Compile"
End Sub

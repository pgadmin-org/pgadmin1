Attribute VB_Name = "basView"
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

Sub cmp_View_DropIfExists(szView_PostgreSQLtable As String, ByVal lngView_oid As Long, Optional ByVal szView_name As String)
 On Error GoTo Err_Handler
    Dim szDropStr As String
    
    ' Where should we get the values ?
    If (szView_PostgreSQLtable = "") Then szView_PostgreSQLtable = "pgadmin_views"
    
    ' Test existence of view
    If cmp_View_Exists(szView_PostgreSQLtable, lngView_oid, szView_name & "") = True Then
    
        If szView_name = "" Then cmp_View_GetValues szView_PostgreSQLtable, lngView_oid, "", szView_name
    
        ' create drop query
        If (szView_PostgreSQLtable = "pgadmin_views") Then
            szDropStr = "DROP VIEW " & QUOTE & szView_name & QUOTE
        Else
            szDropStr = "DELETE FROM " & szView_PostgreSQLtable & " WHERE view_name ='" & szView_name & "'"
        End If
         
        ' Log information
        LogMsg "Dropping view " & szView_name & " in " & szView_PostgreSQLtable & "..."
        LogMsg "Executing: " & szDropStr
        
        ' Execute drop query and close log
        gConnection.Execute szDropStr
        LogQuery szDropStr
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basView, cmp_View_DropIfExists"
End Sub

Function cmp_View_Exists(szView_PostgreSQLtable As String, ByVal lngView_oid As Long, ByVal szView_name As String) As Boolean
 On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
  
    ' Where should we get the values ?
    If (szView_PostgreSQLtable = "") Then szView_PostgreSQLtable = "pgadmin_views"
    
    cmp_View_Exists = False
    If lngView_oid <> 0 Then
        szQueryStr = "SELECT * FROM " & szView_PostgreSQLtable
        szQueryStr = szQueryStr & " WHERE view_OID = " & Str(lngView_oid)
    Else
        If szView_name <> "" Then
            szQueryStr = "SELECT * FROM  " & szView_PostgreSQLtable
            szQueryStr = szQueryStr & " WHERE view_name = '" & szView_name & "' "
        Else
            Exit Function
        End If
    End If
    
      ' retrieve name and arguments of function to drop
    LogMsg "Testing existence of view " & szView_name & " in " & szView_PostgreSQLtable & "..."
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

Sub cmp_View_Create(szView_PostgreSQLtable As String, ByVal szView_name As String, ByVal szView_definition As String)
On Error GoTo Err_Handler
    Dim szCreateStr As String
    Dim szView_oid As Long
    Dim szView_query_oid As Variant
  
    ' Where should we get the values ?
    If (szView_PostgreSQLtable = "") Then szView_PostgreSQLtable = "pgadmin_views"
    
    If (szView_PostgreSQLtable = "pgadmin_views") Then
        szCreateStr = cmp_View_CreateSQL(szView_name, szView_definition)
    Else
        szCreateStr = "INSERT INTO " & szView_PostgreSQLtable & " (View_name, View_definition)"
        szCreateStr = szCreateStr & "VALUES ("
        szCreateStr = szCreateStr & "'" & szView_name & "', "
        szCreateStr = szCreateStr & "'" & szView_definition & "' "
        szCreateStr = szCreateStr & ");"
    End If
    
    LogMsg "Creating view " & szView_name & " in " & szView_PostgreSQLtable & "..."
    LogMsg "Executing: " & szCreateStr
    
    ' Execute drop query and close log
    gConnection.Execute szCreateStr
    LogQuery szCreateStr

  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basView, cmp_Views_Create"
  If Err.Number = -2147467259 Then MsgBox "View " & szView_name & " could not be compiled." & vbCrLf & "Check source code and compile again."
  bContinueRebuilding = False
End Sub

Function cmp_View_CreateSQL(ByVal szView_name As String, ByVal szView_definition As String) As String
On Error GoTo Err_Handler
  Dim szQuery As String

    szQuery = "CREATE VIEW " & szView_name & vbCrLf & " AS " & szView_definition & "; "
    cmp_View_CreateSQL = szQuery
  Exit Function
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basView, cmp_Views_Create"
End Function

Sub cmp_View_GetValues(szView_PostgreSQLtable As String, lngView_oid As Long, Optional szView_name As String, Optional szView_definition As String, Optional szView_owner As String, Optional szView_acl As String, Optional szView_comments As String)
 On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    ' Where should we get the values ?
    If (szView_PostgreSQLtable = "") Then szView_PostgreSQLtable = "pgadmin_views"
        
    ' Select query
    If lngView_oid <> 0 Then
        szQueryStr = "SELECT * from " & szView_PostgreSQLtable
        szQueryStr = szQueryStr & " WHERE view_OID = " & lngView_oid
    Else
        If IsMissing(szView_name) Then szView_name = ""
        szQueryStr = "SELECT * from " & szView_PostgreSQLtable & " WHERE view_name = '" & szView_name & "'"
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
        If Not (IsMissing(szView_name)) Then szView_name = rsComp!view_name & ""
        If Not (IsMissing(szView_owner)) Then szView_owner = rsComp!view_owner & ""
        If Not (IsMissing(szView_acl)) Then szView_acl = rsComp!view_acl & ""
        If (szView_PostgreSQLtable = "pgadmin_views") Then
            If Not (IsMissing(szView_definition)) Then szView_definition = cmp_View_GetViewDef(szView_name)
        Else
            If Not (IsMissing(szView_definition)) Then szView_definition = rsComp!view_definition
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

Public Sub comp_View_CopyToDev()
On Error GoTo Err_Handler
    Dim szQuery As String
    Dim szView() As Variant
    Dim szView_definition As String
    Dim szView_name As String
    Dim iUbound As Long
    Dim iLoop As Long
    Dim rsComp As New Recordset
    
    
    szQuery = "TRUNCATE TABLE pgadmin_dev_views;" & _
    "  INSERT INTO pgadmin_dev_views SELECT * from " & _
    "  pgadmin_views " & _
    "  WHERE view_oid > " & LAST_SYSTEM_OID & _
    "  AND view_name NOT LIKE 'pgadmin_%' " & _
    "  AND view_name NOT LIKE 'pg_%' " & _
    "  ORDER BY view_name; " & _
    "  UPDATE pgadmin_dev_views SET view_iscompiled = 'f';"
    LogMsg "Copying pgadmin_views to pgadmin_dev_views..."
    LogMsg "Executing: " & szQuery
    gConnection.Execute szQuery
       
    ' initialize pgadmin_dev_view
    szQuery = "SELECT view_name FROM pgadmin_dev_views ORDER BY view_oid"
    If rsComp.State <> adStateClosed Then rsComp.Close
    rsComp.Open szQuery, gConnection, adOpenDynamic
    
    If Not (rsComp.EOF) Then
        szQuery = ""
        szView = rsComp.GetRows
        If rsComp.State <> adStateClosed Then rsComp.Close
        iUbound = UBound(szView, 2)
            For iLoop = 0 To iUbound
                 'Get view definition
                 szView_name = szView(0, iLoop)
                 szView_definition = Replace(cmp_View_GetViewDef(szView_name), "'", "''")
                
                ' Update definition of view
                szQuery = szQuery & "UPDATE pgadmin_dev_views SET view_definition = '" & szView_definition & "' WHERE view_name = '" & szView_name & "'; "
            Next iLoop
            LogMsg "Executing: " & szQuery
            gConnection.Execute szQuery
    End If
    
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basView, comp_View_CopyToDev"
End Sub

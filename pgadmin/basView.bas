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

Sub cmp_View_DropIfExists(szView_PostgreSqlTable As String, ByVal lngView_oid As Long, Optional ByVal szView_name As String)
 On Error GoTo Err_Handler
    Dim szDropStr As String
    
    ' Where should we get the values ?
    If (szView_PostgreSqlTable = "") Then szView_PostgreSqlTable = "pgadmin_views"
    
    ' Test existence of view
    If cmp_View_Exists(szView_PostgreSqlTable, lngView_oid, szView_name & "") = True Then
    
        If szView_name = "" Then cmp_View_GetValues szView_PostgreSqlTable, lngView_oid, "", szView_name
    
        ' create drop query
        If (szView_PostgreSqlTable = "pgadmin_views") Then
            szDropStr = "DROP VIEW " & QUOTE & szView_name & QUOTE
        Else
            szDropStr = "DELETE FROM " & szView_PostgreSqlTable & " WHERE view_name ='" & szView_name & "'"
        End If
         
        ' Log information
        LogMsg "Dropping view " & szView_name & " in " & szView_PostgreSqlTable & "..."
        LogMsg "Executing: " & szDropStr
        
        ' Execute drop query and close log
        gConnection.Execute szDropStr
        LogQuery szDropStr
    End If
  Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basView, cmp_View_DropIfExists"
End Sub

Function cmp_View_Exists(szView_PostgreSqlTable As String, ByVal lngView_oid As Long, ByVal szView_name As String) As Boolean
 On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
  
    ' Where should we get the values ?
    If (szView_PostgreSqlTable = "") Then szView_PostgreSqlTable = "pgadmin_views"
    
    cmp_View_Exists = False
    If lngView_oid <> 0 Then
        szQueryStr = "SELECT * FROM " & szView_PostgreSqlTable
        szQueryStr = szQueryStr & " WHERE view_OID = " & Str(lngView_oid)
    Else
        If szView_name <> "" Then
            szQueryStr = "SELECT * FROM  " & szView_PostgreSqlTable
            szQueryStr = szQueryStr & " WHERE view_name = '" & szView_name & "' "
        Else
            Exit Function
        End If
    End If
    
      ' retrieve name and arguments of function to drop
    LogMsg "Testing existence of view " & szView_name & " in " & szView_PostgreSqlTable & "..."
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

Sub cmp_View_Create(szView_PostgreSqlTable As String, ByVal szView_name As String, ByVal szView_definition As String)
On Error GoTo Err_Handler
    Dim szCreateStr As String
    Dim szView_oid As Long
    Dim szView_query_oid As Variant
  
    ' Where should we get the values ?
    If (szView_PostgreSqlTable = "") Then szView_PostgreSqlTable = "pgadmin_views"
    
    If (szView_PostgreSqlTable = "pgadmin_views") Then
        szCreateStr = cmp_View_CreateSQL(szView_name, szView_definition)
    Else
        szCreateStr = "INSERT INTO " & szView_PostgreSqlTable & " (View_name, View_definition)"
        szCreateStr = szCreateStr & "VALUES ("
        szCreateStr = szCreateStr & "'" & szView_name & "', "
        szCreateStr = szCreateStr & "'" & Replace(szView_definition, "'", "''") & "' "
        szCreateStr = szCreateStr & ");"
    End If
    
    LogMsg "Creating view " & szView_name & " in " & szView_PostgreSqlTable & "..."
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

Sub cmp_View_GetValues(szView_PostgreSqlTable As String, lngView_oid As Long, Optional szView_name As String, Optional szView_definition As String, Optional szView_owner As String, Optional szView_acl As String, Optional szView_comments As String)
 On Error GoTo Err_Handler
    Dim szQueryStr As String
    Dim rsComp As New Recordset
    
    ' Where should we get the values ?
    If (szView_PostgreSqlTable = "") Then szView_PostgreSqlTable = "pgadmin_views"
        
    ' Select query
    If lngView_oid <> 0 Then
        szQueryStr = "SELECT * from " & szView_PostgreSqlTable
        szQueryStr = szQueryStr & " WHERE view_OID = " & lngView_oid
    Else
        If IsMissing(szView_name) Then szView_name = ""
        szQueryStr = "SELECT * from " & szView_PostgreSqlTable & " WHERE view_name = '" & szView_name & "'"
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
        If (szView_PostgreSqlTable = "pgadmin_views") Then
            If Not (IsMissing(szView_definition)) Then szView_definition = cmp_View_GetViewDef(szView_name)
        Else
            If Not (IsMissing(szView_definition)) Then szView_definition = rsComp!view_definition & ""
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
    
    cmp_View_GetViewDef = ""
    If Not rsTemp.EOF Then
        cmp_View_GetViewDef = rsTemp!result & ""
    End If
    
    Exit Function
Err_Handler:
  cmp_View_GetViewDef = "Not a view"
End Function

Public Sub cmp_View_CopyToDev()
On Error GoTo Err_Handler
    Dim szQuery As String
    Dim rsTemp As New Recordset
    Dim szView As Variant
    Dim iUbound As Long
    Dim iLoop As Long
    
    Dim lView_oid As Long
    Dim szView_name As String
    Dim szView_definition As String
    Dim szView_owner As String
    Dim szView_acl As String
    Dim szView_comments As String

    LogMsg "Exporting pgadmin_views to pgadmin_dev_views..."
    
    szQuery = "SELECT view_name, pgadmin_get_viewdef(view_name) as view_definition FROM pgadmin_views WHERE view_name NOT LIKE 'pgadmin_%' AND view_name NOT LIKE 'pg_%' ORDER BY view_name"
    LogMsg "Executing: " & szQuery
    rsTemp.Open szQuery, gConnection, adOpenDynamic
   
    If Not (rsTemp.EOF) Then
      szView = rsTemp.GetRows
      iUbound = UBound(szView, 2)
      For iLoop = 0 To iUbound
           szView_name = szView(0, iLoop)
           szView_definition = szView(1, iLoop)
           
           lView_oid = 0
           cmp_View_GetValues "pgadmin_views", lView_oid, szView_name, szView_definition, szView_owner, szView_acl, szView_comments
           cmp_View_DropIfExists "pgadmin_dev_views", 0, szView_name
           cmp_View_Create "pgadmin_dev_views", szView_name, szView_definition
      Next iLoop
      rsTemp.Close
      Erase szView
    End If
    
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basView, cmp_View_CopyToDev"
End Sub

Public Sub cmp_View_DropAll(Optional szView_PostgreSqlTable As String)
On Error GoTo Err_Handler
    Dim szQuery As String
    Dim szView() As Variant
    Dim iLoop As Long
    Dim iUbound As Long
    Dim rsView As New Recordset
    Dim szView_name As String
    
    If IsMissing(szView_PostgreSqlTable) Or (szView_PostgreSqlTable = "") Then szView_PostgreSqlTable = "pgadmin_views"
        
    If (szView_PostgreSqlTable = "pgadmin_views") Then
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
                     szView_name = szView(0, iLoop)
                     cmp_View_DropIfExists "", 0, szView_name
                Next iLoop
            Erase szView
        End If
    Else
        szQuery = "TRUNCATE " & szView_PostgreSqlTable
        LogMsg "Truncating " & szView_PostgreSqlTable & "..."
        LogMsg "Executing: " & szQuery
        gConnection.Execute szQuery
    End If
   
    Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err, "basProject, cmp_View_DropAll"
End Sub


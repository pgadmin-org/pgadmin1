Attribute VB_Name = "basTable"
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
  If Err.Number <> 0 Then LogError Err, "basTable, cmp_Table_DropIfExists"
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
  If Err.Number <> 0 Then LogError Err, "basTable, cmp_Table_Exists"
End Function

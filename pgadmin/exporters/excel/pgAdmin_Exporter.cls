VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgAdmin_Exporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Implements pgExporter

Private Property Get pgExporter_Description() As String
  pgExporter_Description = "OLE Link to MS Excel"
End Property

Private Property Get pgExporter_Version() As String
  pgExporter_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Private Property Get pgExporter_Author() As String
  pgExporter_Author = "Written by Dave Page" & vbCrLf & _
                      "http://www.greatbridge.org/project/pgadmin/"
End Property

Private Sub pgExporter_Export(rsData As Variant, Optional szConfig As String)
On Error Resume Next
Dim X As Integer
Dim Y As Integer
  Dim ExcelApp As Object
  Set ExcelApp = CreateObject("excel.application")
  ExcelApp.Visible = True
  ExcelApp.Workbooks.Add
  
  'Create Header
  For X = 0 To rsData.Fields.Count - 1
    ExcelApp.Cells(1, X + 1).NumberFormat = "@"
    ExcelApp.Cells(1, X + 1).Font.Bold = True
    ExcelApp.Cells(1, X + 1).Value = rsData.Fields(X).Name
  Next

  'Enter Data
  
  rsData.MoveFirst
  Y = 2
  While Not rsData.EOF
    For X = 0 To rsData.Fields.Count - 1
      ExcelApp.Cells(Y, X + 1).NumberFormat = "@"
      ExcelApp.Cells(Y, X + 1).Value = rsData.Fields(X).Value
    Next
    rsData.MoveNext
    Y = Y + 1
  Wend
  Screen.MousePointer = vbNormal

  'Autofit
  
  For X = 0 To rsData.Fields.Count - 1
    ExcelApp.Columns(GetCol(X + 1)).AutoFit
  Next
  
End Sub

Private Function GetCol(ColNum As Integer) As String
Dim Char1 As String
Dim Char2 As String
Dim X As Double
  If ColNum < 1 Or ColNum > 256 Then Exit Function
  X = ColNum / 26
  If X < 1 Then
    Char1 = ""
  Else
    Char1 = Chr(96 + Int(X))
  End If
  Char2 = Chr(96 + (ColNum - (Int(X) * 26)))
  GetCol = Char1 & Char2
End Function



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

Private Property Get pgExporter_Author() As String
  pgExporter_Author = "Written by Dave Page" & vbCrLf & _
                      "http://www.greatbridge.org/project/pgadmin/"
End Property

Private Property Get pgExporter_Description() As String
  pgExporter_Description = "ASCII Text File"
End Property

Private Property Get pgExporter_Version() As String
  pgExporter_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Private Sub pgExporter_Export(rsData As Variant, Optional szConfig As String)
Dim fNum As Integer
Dim X As Integer
Dim Y As Integer
Dim lCount As Long
Dim szFilename As String
Dim szQuote As String
Dim szDelimiter As String
Dim bTrailing As Boolean
Dim szData As String
Dim szField As String

  'Currently we don't support unattended operation...
  
  If szConfig = "" Then
    frmExport.Show vbModal
    szConfig = frmExport.txtFileName
  End If
  szFilename = frmExport.szFilename
  szQuote = frmExport.szQuote
  szDelimiter = frmExport.szDelimiter
  bTrailing = frmExport.bTrailing
  frmExport.Hide
  Unload frmExport
  DoEvents
  
  If szFilename = "" Then
    MsgBox "No filename specified - Export aborted!", vbExclamation, "Error"
    Exit Sub
  End If
  
  'Open the file
  Screen.MousePointer = vbArrowHourglass
  fNum = FreeFile
  Open szFilename For Output As #fNum
  
  'Loop through the records, and save line by line.
  
  lCount = 0
  Load frmStatus
  frmStatus.Show
  While Not rsData.EOF
    For X = 0 To rsData.Fields.Count - 1
      If frmExport.lvSubMap.ListItems.Count > 0 Then
        For Y = 1 To frmExport.lvSubMap.ListItems.Count
          szField = Replace(rsData.Fields(X).Value & "", frmExport.lvSubMap.ListItems(Y).Text, frmExport.lvSubMap.ListItems(Y).SubItems(1))
        Next
      Else
        szField = rsData.Fields(X).Value & ""
      End If
      szData = szData & szQuote & szField & szQuote & szDelimiter
    Next
    If bTrailing = 0 Then
      szData = Mid(szData, 1, Len(szData) - Len(szDelimiter))
    End If
    Print #fNum, szData
    lCount = lCount + 1
    frmStatus.lblCount = lCount
    frmStatus.Refresh
    DoEvents
    szData = ""
    rsData.MoveNext
  Wend
  rsData.Close
   
  'Close the file
  
  Set rsData = Nothing
  Close #fNum
  Screen.MousePointer = vbNormal
  Unload frmStatus
  
  'If szConfig is blank (ie. not unattended) then display results.
  MsgBox "Exported " & lCount & " records to " & szFilename, vbInformation, "Data Export"
End Sub



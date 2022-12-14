VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSQLOutput 
   Caption         =   "SQL Output"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   Icon            =   "frmSQLOutput.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   8160
   Begin VB.PictureBox picTools 
      Height          =   465
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   4905
      TabIndex        =   5
      Top             =   1215
      Width           =   4965
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1845
         TabIndex        =   8
         ToolTipText     =   "Delete the selected record."
         Top             =   45
         Width           =   825
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   330
         Left            =   945
         TabIndex        =   7
         ToolTipText     =   "Edit the selected record."
         Top             =   45
         Width           =   825
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   330
         Left            =   45
         TabIndex        =   6
         ToolTipText     =   "Add a new record."
         Top             =   45
         Width           =   825
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   330
         Left            =   1395
         TabIndex        =   10
         ToolTipText     =   "Add a new record."
         Top             =   45
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   330
         Left            =   45
         TabIndex        =   9
         ToolTipText     =   "Add a new record."
         Top             =   45
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "0 Records"
         Height          =   195
         Left            =   2745
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox picEdit 
      Height          =   1005
      Left            =   0
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4515
      Begin VB.PictureBox picScroll 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   0
         ScaleHeight     =   59
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   242
         TabIndex        =   2
         Top             =   0
         Width           =   3630
         Begin VB.TextBox txtField 
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   0
            Left            =   45
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   225
            Width           =   3300
         End
         Begin VB.Label lblField 
            AutoSize        =   -1  'True
            Caption         =   "Field Label"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   4
            Top             =   45
            Width           =   765
         End
      End
      Begin VB.VScrollBar scScroll 
         Height          =   780
         LargeChange     =   100
         Left            =   3960
         SmallChange     =   10
         TabIndex        =   1
         Top             =   0
         Width           =   240
      End
   End
   Begin MSComctlLib.ListView lvData 
      Height          =   1185
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   2090
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmSQLOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim rsSQL As New Recordset
Dim szTable As String
Dim szWhere As String
Dim bUpdateable As Boolean

Private Sub cmdAdd_Click()
On Error GoTo Err_Handler
  BuildEditBox
  lblInfo.Caption = "Add Record"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQLOutput, cmdAdd_Click"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
  HideEditBox
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQLOutput, cmdCancel_Click"
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Err_Handler
Dim x As Integer
Dim y As Integer
Dim Z As Integer
Dim szCriteria As String
Dim rsCount As New Recordset
Dim szQuery As String
Dim szValues() As String
Dim szKeys() As String
Dim bFlag As Boolean
  If MsgBox("Are you sure you wish to delete the selected record?", vbQuestion + vbYesNo, "Delete Record?") = vbNo Then Exit Sub
  
  'Build the most concise WHERE clause we can. adDate and adDBDate fields should be
  'formatted as ISO dates.
  For x = 0 To lvData.ColumnHeaders.Count - 1
    If x = 0 Then
      If lvData.SelectedItem.Text <> "" Then
        Select Case Val(Mid(lvData.ColumnHeaders(x + 1).Key, InStr(1, lvData.ColumnHeaders(x + 1).Key, ":") + 1))
          Case adDate
            szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.Text, "yyyy-MM-dd") & "' AND "
          Case adDBDate
            szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.Text, "yyyy-MM-dd") & "' AND "
            Case adDBTimeStamp
            szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.Text, "yyyy-MM-dd hh:mm:ss") & "' AND "
          Case Else
            szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & dbSZ(lvData.SelectedItem.Text) & "' AND "
        End Select
      End If
    Else
      If lvData.SelectedItem.SubItems(x) <> "" Then
        Select Case Val(Mid(lvData.ColumnHeaders(x + 1).Key, InStr(1, lvData.ColumnHeaders(x + 1).Key, ":") + 1))
          Case adDate
            szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.SubItems(x), "yyyy-MM-dd") & "' AND "
          Case adDBDate
            szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.SubItems(x), "yyyy-MM-dd") & "' AND "
          Case adDBTimeStamp
            szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.SubItems(x), "yyyy-MM-dd hh:mm:ss") & "' AND "
          Case Else
            szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & dbSZ(lvData.SelectedItem.SubItems(x)) & "' AND "
        End Select
      End If
    End If
  Next
  
  'Find out how many records would be affected. Abort if zero, delete if 1 or
  'give the option to delete if > 1
  StartMsg "Counting matching records..."
  If Len(szCriteria) > 5 Then szCriteria = Mid(szCriteria, 1, Len(szCriteria) - 5)
  szQuery = "SELECT count(*) AS count FROM " & szTable & " WHERE " & szCriteria
  If szWhere <> "" Then szQuery = szQuery & " AND " & szWhere
  LogMsg "Executing: " & szQuery
  rsCount.Open szQuery, gConnection, adOpenForwardOnly
  
  'Prepare the delete query for later
  szQuery = "DELETE FROM " & szTable & " WHERE " & szCriteria
  If szWhere <> "" Then szQuery = szQuery & " AND " & szWhere
  EndMsg
  If Not rsCount.EOF Then
    Select Case rsCount!Count
      Case 0
        MsgBox "Could not locate the record for deletion in the database!", vbExclamation, "Error"
        GoTo Done
      Case 1
        StartMsg "Deleting record..."
        LogMsg "Executing: " & szQuery
        gConnection.Execute szQuery
        lvData.ListItems.Remove (lvData.SelectedItem.Index)
        GoTo Done
      Case Else
        If MsgBox("The selected record could not be uniquely identified. " & rsCount!Count & " records match, and will all be deleted if you proceed. Do you wish to continue?", vbQuestion + vbYesNo, "Delete Multiple Records") = vbNo Then Exit Sub
        StartMsg "Deleting records..."
        LogMsg "Executing: " & szQuery
        gConnection.Execute szQuery
        
        'Get all the values in the selected row, then iterate through all rows and delete matching
        ReDim szValues(lvData.ColumnHeaders.Count - 1)
        For x = 0 To lvData.ColumnHeaders.Count - 1
          If x = 0 Then
            szValues(x) = lvData.SelectedItem.Text
          Else
            szValues(x) = lvData.SelectedItem.SubItems(x)
          End If
        Next x
        
        'Delete matching rows.
        For x = lvData.ListItems.Count To 1 Step -1
          bFlag = False
          For y = 1 To lvData.ColumnHeaders.Count - 1
            If szValues(y) <> lvData.ListItems(x).SubItems(y) Then
              bFlag = True
              Exit For
            End If
          Next y
          If Not (bFlag Or szValues(0) <> lvData.ListItems(x).Text) Then
            lvData.ListItems.Remove lvData.ListItems(x).Index
          End If
        Next x
        GoTo Done
    End Select
  End If
Done:
  EndMsg
  If lvData.ListItems.Count > 0 Then
    lblInfo.Caption = "Record " & lvData.SelectedItem.Index & " of " & lvData.ListItems.Count
  Else
    lblInfo.Caption = "Record 0 of 0"
  End If
  If lvData.ListItems.Count > 0 And bUpdateable = True Then cmdEdit.Enabled = True
  If lvData.ListItems.Count > 0 And bUpdateable = True Then cmdDelete.Enabled = True
  If rsCount.State <> adStateClosed Then rsCount.Close
  Set rsCount = Nothing
  Exit Sub
Err_Handler:
  If rsCount.State <> adStateClosed Then rsCount.Close
  Set rsCount = Nothing
  If Err.Number <> 0 Then LogError Err, "frmSQLOutput, cmdDelete_Click"
End Sub

Private Sub cmdSave_Click()
On Error GoTo Err_Handler
Dim szQuery As String
Dim szColumns As String
Dim szValues As String
Dim szCriteria As String
Dim szCells() As String
Dim x As Integer
Dim y As Integer
Dim Z As Integer
Dim bFlag As Boolean
Dim itmX As ListItem
Dim rsCount As New Recordset
  If lblInfo.Caption = "Add Record" Then
    'Add new record
    'First build lists of columns and values
    LogMsg "Adding new record..."
    For x = 0 To lblField.Count - 1
      If txtField(x).Text <> "" Then
        szColumns = szColumns & QUOTE & lblField(x).Caption & QUOTE & ", "
        Select Case Val(Mid(lvData.ColumnHeaders(x + 1).Key, InStr(1, lvData.ColumnHeaders(x + 1).Key, ":") + 1))
          Case adDate
            szValues = szValues & "'" & Format(txtField(x).Text, "yyyy-MM-dd") & "', "
          Case adDBDate
            szValues = szValues & "'" & Format(txtField(x).Text, "yyyy-MM-dd") & "', "
          Case adDBTimeStamp
            szValues = szValues & "'" & Format(txtField(x).Text, "yyyy-MM-dd hh:mm:ss") & "', "
          Case Else
            szValues = szValues & "'" & dbSZ(txtField(x).Text) & "', "
        End Select
      End If
    Next x
    
    'Check the data, then trim the ', ' from the end of each string and create the SQL query
    If szColumns = "" Then
      EndMsg
      MsgBox "No data has been entered!", vbExclamation, "Error"
      Exit Sub
    End If
    If Len(szColumns) > 2 Then szColumns = "(" & Mid(szColumns, 1, Len(szColumns) - 2) & ")"
    If Len(szValues) > 2 Then szValues = "(" & Mid(szValues, 1, Len(szValues) - 2) & ")"
    szQuery = "INSERT INTO " & szTable & " " & szColumns & " VALUES " & szValues
    
    'Execute the query
    LogMsg "Executing: " & szQuery
    gConnection.Execute szQuery
    
    'Now add the record to the grid. If the query failed, we won't get to here 'cos
    'we'll be in the error handler.
    Set itmX = lvData.ListItems.Add(, , txtField(0).Text)
    For x = 1 To lblField.Count - 1
      itmX.SubItems(x) = txtField(x).Text
    Next x
    GoTo Done
  Else
    'Update record
    'First build lists of columns and values
    For x = 0 To lblField.Count - 1
      If txtField(x).Tag = "Y" Then
        Select Case Val(Mid(lvData.ColumnHeaders(x + 1).Key, InStr(1, lvData.ColumnHeaders(x + 1).Key, ":") + 1))
          Case adDate
            szValues = szValues & QUOTE & lblField(x).Caption & QUOTE & " = '" & Format(txtField(x).Text, "yyyy-MM-dd") & "', "
          Case adDBDate
            szValues = szValues & QUOTE & lblField(x).Caption & QUOTE & " = '" & Format(txtField(x).Text, "yyyy-MM-dd") & "', "
          Case adDBTimeStamp
            szValues = szValues & QUOTE & lblField(x).Caption & QUOTE & " = '" & Format(txtField(x).Text, "yyyy-MM-dd hh:mm:ss") & "', "
          Case Else
            szValues = szValues & QUOTE & lblField(x).Caption & QUOTE & " = '" & dbSZ(txtField(x).Text) & "', "
        End Select
      End If
    Next x
    
    'Build the most concise WHERE clause we can. adDate and adDBDate fields should be
    'formatted as ISO dates.
    For x = 0 To lvData.ColumnHeaders.Count - 1
      If x = 0 Then
        If lvData.SelectedItem.Text <> "" Then
          Select Case Val(Mid(lvData.ColumnHeaders(x + 1).Key, InStr(1, lvData.ColumnHeaders(x + 1).Key, ":") + 1))
            Case adDate
              szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.Text, "yyyy-MM-dd") & "' AND "
            Case adDBDate
              szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.Text, "yyyy-MM-dd") & "' AND "
            Case adDBTimeStamp
              szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.Text, "yyyy-MM-dd hh:mm:ss") & "' AND "
            Case Else
              szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & dbSZ(lvData.SelectedItem.Text) & "' AND "
          End Select
        End If
      Else
        If lvData.SelectedItem.SubItems(x) <> "" Then
          Select Case Val(Mid(lvData.ColumnHeaders(x + 1).Key, InStr(1, lvData.ColumnHeaders(x + 1).Key, ":") + 1))
            Case adDate
              szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.SubItems(x), "yyyy-MM-dd") & "' AND "
            Case adDBDate
              szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.SubItems(x), "yyyy-MM-dd") & "' AND "
            Case adDBTimeStamp
              szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & Format(lvData.SelectedItem.SubItems(x), "yyyy-MM-dd hh:mm:ss") & "' AND "
            Case Else
              szCriteria = szCriteria & QUOTE & lvData.ColumnHeaders(x + 1).Text & QUOTE & " = '" & dbSZ(lvData.SelectedItem.SubItems(x)) & "' AND "
          End Select
        End If
      End If
    Next

    'Check the data
    If szValues = "" Then
      EndMsg
      MsgBox "No data has been modified!", vbExclamation, "Error"
      Exit Sub
    End If
    
    'Find out how many records would be affected. Abort if zero, update if 1 or
    'give the option to update all if > 1
    StartMsg "Counting matching records..."
    If Len(szValues) > 2 Then szValues = Mid(szValues, 1, Len(szValues) - 2)
    If Len(szCriteria) > 5 Then szCriteria = Mid(szCriteria, 1, Len(szCriteria) - 5)
    szQuery = "SELECT count(*) AS count FROM " & szTable & " WHERE " & szCriteria
    If szWhere <> "" Then szQuery = szQuery & " AND " & szWhere
    LogMsg "Executing: " & szQuery
    rsCount.Open szQuery, gConnection, adOpenForwardOnly

    'Prepare the update query for later
    szQuery = "UPDATE " & szTable & " SET " & szValues & " WHERE " & szCriteria
    If szWhere <> "" Then szQuery = szQuery & " AND " & szWhere
    EndMsg
    If Not rsCount.EOF Then
      Select Case rsCount!Count
        Case 0
          MsgBox "Could not locate the record for updating in the database!", vbExclamation, "Error"
          GoTo Done
        Case 1
          StartMsg "Updating record..."
          LogMsg "Executing: " & szQuery
          gConnection.Execute szQuery
          'Update the grid
          For x = 0 To lblField.Count - 1
            If x = 0 Then
              lvData.SelectedItem.Text = txtField(x).Text
            Else
              lvData.SelectedItem.SubItems(x) = txtField(x).Text
            End If
          Next x
          GoTo Done
        Case Else
          If MsgBox("The selected record could not be uniquely identified. " & rsCount!Count & " records match, and will all be updated if you proceed. Do you wish to continue?", vbQuestion + vbYesNo, "Update Multiple Records") = vbNo Then Exit Sub
          StartMsg "Updating records..."
          LogMsg "Executing: " & szQuery
          gConnection.Execute szQuery

          'Get all the values in the selected row, then iterate through all rows and update matching
          ReDim szCells(lvData.ColumnHeaders.Count - 1)
          For x = 0 To lvData.ColumnHeaders.Count - 1
            If x = 0 Then
              szCells(x) = lvData.SelectedItem.Text
            Else
              szCells(x) = lvData.SelectedItem.SubItems(x)
            End If
          Next x
          
          'Update matching rows.
          For x = lvData.ListItems.Count To 1 Step -1
            bFlag = False
            For y = 1 To lvData.ColumnHeaders.Count - 1
              If szCells(y) <> lvData.ListItems(x).SubItems(y) Then
                bFlag = True
                Exit For
              End If
            Next y
            If Not (bFlag Or szCells(0) <> lvData.ListItems(x).Text) Then
              For Z = 0 To lblField.Count - 1
                If Z = 0 Then
                  lvData.ListItems(x).Text = txtField(Z).Text
                Else
                  lvData.ListItems(x).SubItems(Z) = txtField(Z).Text
                End If
              Next Z
            End If
          Next x
          GoTo Done
      End Select
    End If
  End If
Done:
  EndMsg
  HideEditBox
  If lvData.ListItems.Count > 0 Then
    lblInfo.Caption = "Record " & lvData.SelectedItem.Index & " of " & lvData.ListItems.Count
  Else
    lblInfo.Caption = "Record 0 of 0"
  End If
  If lvData.ListItems.Count > 0 And bUpdateable = True Then cmdEdit.Enabled = True
  If lvData.ListItems.Count > 0 And bUpdateable = True Then cmdDelete.Enabled = True
  If rsCount.State <> adStateClosed Then rsCount.Close
  Set rsCount = Nothing
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmSQLOutput, cmdSave_Click"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Set rsSQL = Nothing
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQLOutput, Form_Load"
End Sub

Public Sub Display(rsQuery As Recordset)
On Error GoTo Err_Handler
Dim iStart As Integer
Dim iEnd As Integer
Dim iTemp As Integer
Dim x As Integer
Dim szQuery As String
Dim szChar As String
Dim szBits() As String
Dim bInQuotes As Boolean
Dim bFlag As Boolean
  Set rsSQL = rsQuery

  'Figure out if the query is updateable. This is the case if:
  '1) There is one and only one table
  '2) The table is not actually a view
  'We must also get the tablename, and any WHERE clause to help
  'with update queries.
  
  'Start by converting any spaces inside double quotes to tabs which
  'should never appear in the SQL
  szQuery = ""
  bInQuotes = False
  For x = 1 To Len(rsSQL.Source)
    szChar = Mid(rsSQL.Source, x, 1)
    If szChar = QUOTE Then
      szQuery = szQuery & QUOTE
      bInQuotes = Not bInQuotes
    ElseIf szChar = " " And bInQuotes Then
      szQuery = szQuery & vbTab
    Else
      szQuery = szQuery & szChar
    End If
  Next x
  
  'Find the FROM clause. If it is inside single quotes then we
  'should try again - it won't in doubles as there are no spaces
  'in doubles anymore.
  iStart = 0
  bFlag = False
  bInQuotes = False
  While bFlag = False
    iStart = InStr(iStart + 1, UCase(szQuery), " FROM ")
    If iStart = 0 Then 'No FROMs found
      bFlag = True
    Else 'Found a FROM, check it's not in quotes
      For x = 1 To iStart
        If Mid(szQuery, x, 1) = "'" Then bInQuotes = Not bInQuotes
      Next x
      If Not bInQuotes Then bFlag = True
    End If
  Wend
  
  'If FROM is not found then we must have a tableless query
  '(eg. SELECT version()), otherwise increment iStart past the FROM
  If iStart = 0 Then
    szTable = ""
    szWhere = ""
    bUpdateable = False
    GoTo GotInfo
  Else
    iStart = iStart + 6
  End If
  
  'Find the end of the FROM clause. This will be delimited by one of the
  'following, or the end of the string:
  'WHERE GROUP HAVING UNION INTERSECT EXCEPT ORDER FOR LIMIT
  iEnd = InStr(iStart, UCase(szQuery), " WHERE ")
  iTemp = InStr(iStart, UCase(szQuery), " GROUP ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " HAVING ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " UNION ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " INTERSECT ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " EXCEPT ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " ORDER ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " FOR ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " LIMIT ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  If iEnd = 0 Then iEnd = Len(szQuery) + 1

  'Split the FROM clause by space. We can then iterate through each element of
  'the array to figure out whether we have more than one table. The following
  'conditions could determine that we have more than one table:
  '1) A trailing , on any element
  '2) An element containing JOIN INNER OUTER LEFT RIGHT FULL CROSS or [(]SELECT
  szBits = Split(Mid(szQuery, iStart, iEnd - iStart), " ")
  For x = 0 To UBound(szBits)
    If Right(szBits(x), 1) = "," Then
      szTable = ""
      szWhere = ""
      bUpdateable = False
      GoTo GotInfo
    End If
    If UCase(szBits(x)) = "JOIN" Or _
       UCase(szBits(x)) = "INNER" Or _
       UCase(szBits(x)) = "OUTER" Or _
       UCase(szBits(x)) = "LEFT" Or _
       UCase(szBits(x)) = "RIGHT" Or _
       UCase(szBits(x)) = "FULL" Or _
       UCase(szBits(x)) = "CROSS" Or _
       UCase(szBits(x)) = "SELECT" Or _
       UCase(szBits(x)) = "(SELECT" Then
      szTable = ""
      szWhere = ""
      bUpdateable = False
      GoTo GotInfo
    End If
  Next x

  'If we got this far then we should only have one table so we should
  'get it's name. It should be the first item in the array unless
  'ONLY was specified

  If UCase(szBits(0)) = "ONLY" Then
    szTable = Replace(szBits(1), vbTab, " ")
  Else
    szTable = Replace(szBits(0), vbTab, " ")
  End If
  
  'Check to see if our table is actually a view. If it is then we can't
  'update :-(
  If ObjectExists(szTable, tView) <> 0 Then
    szTable = ""
    szWhere = ""
    bUpdateable = False
    GoTo GotInfo
  End If
  
  'Yippee!
  bUpdateable = True
  
  'As we're updateable we should also extract any WHERE clause
  'to add to any query based updates we may do. This will help
  'us to locate the exact record that the user wants to update.
  iStart = 0
  bFlag = False
  bInQuotes = False
  While bFlag = False
    iStart = InStr(iStart + 1, UCase(szQuery), " WHERE ")
    If iStart = 0 Then 'No WHEREs found
      bFlag = True
    Else 'Found a WHERE, check it's not in quotes
      For x = 1 To iStart
        If Mid(szQuery, x, 1) = "'" Then bInQuotes = Not bInQuotes
      Next x
      If Not bInQuotes Then bFlag = True
    End If
  Wend

  'If WHERE is not found then we must have an 'all records' query
  'otherwise increment iStart past the WHERE
  If iStart = 0 Then
    szWhere = ""
    GoTo GotInfo
  Else
    iStart = iStart + 7
  End If
  
  'Find the end of the WHERE clause. This will be delimited by one of the
  'following, or the end of the string:
  'GROUP HAVING UNION INTERSECT EXCEPT ORDER FOR LIMIT
  iEnd = InStr(iStart, UCase(szQuery), " GROUP ")
  iTemp = InStr(iStart, UCase(szQuery), " HAVING ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " UNION ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " INTERSECT ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " EXCEPT ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " ORDER ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " FOR ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  iTemp = InStr(iStart, UCase(szQuery), " LIMIT ")
  If iTemp <> 0 And iTemp < iEnd Then iEnd = iTemp
  If iEnd = 0 Then iEnd = Len(szQuery) + 1
  
  szWhere = Trim(Mid(szQuery, iStart, iEnd - iStart))

GotInfo:

  'Setup the form
  Me.Caption = "SQL Output - " & rsQuery.Source
  If bUpdateable Then
    cmdEdit.Enabled = True
    cmdAdd.Enabled = True
    cmdDelete.Enabled = True
  Else
    cmdEdit.Enabled = False
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
  End If
  LoadGrid
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQLOutput, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 And Me.ScaleHeight > 0 Then
    If Me.WindowState = 0 Then
      If Me.Width < 5820 Then Me.Width = 5820
      If Me.Height < 3600 Then Me.Height = 3600
    End If
    
    picTools.Visible = True
    picTools.Width = Me.ScaleWidth
    picTools.Top = Me.ScaleHeight - picTools.Height
    lvData.Width = Me.ScaleWidth
    lvData.Height = Me.ScaleHeight - picTools.Height
    picEdit.Height = lvData.Height
    picEdit.Width = lvData.Width
    picScroll.Width = picEdit.ScaleWidth - scScroll.Width
    scScroll.Left = picScroll.Width
    scScroll.Height = picEdit.ScaleHeight

  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQLOutput, Form_Resize"
End Sub

Private Sub LoadGrid()
On Error GoTo Err_Handler
Dim x As Long
Dim itmX As ListItem

  cmdSave.Visible = False
  cmdCancel.Visible = False
  
  'Load Headers
  StartMsg "Loading Data..."
  lvData.ColumnHeaders.Clear
  For x = 0 To rsSQL.Fields.Count - 1
    lvData.ColumnHeaders.Add , "C" & x & ":" & rsSQL.Fields(x).Type, rsSQL.Fields(x).Name & ""
  Next x
      
  'Load Data
  lvData.ListItems.Clear
  lblInfo.Caption = "Record 0 of 0"
  If Not (rsSQL.EOF And rsSQL.BOF) Then
    While Not rsSQL.EOF
    
      'Add the listitem
      Select Case rsSQL.Fields(0).Type
        Case adDBTime
          Set itmX = lvData.ListItems.Add(, , Format(rsSQL.Fields(0).Value & "", "ttttt"))
        Case Else
          Set itmX = lvData.ListItems.Add(, , rsSQL.Fields(0).Value & "")
      End Select
        
      'Add the extra fields
      For x = 1 To rsSQL.Fields.Count - 1
        Select Case rsSQL.Fields(x).Type
          Case adDBTime
            itmX.SubItems(x) = Format(rsSQL.Fields(x).Value & "", "ttttt")
          Case Else
            itmX.SubItems(x) = rsSQL.Fields(x).Value & ""
        End Select
      Next
      rsSQL.MoveNext
    Wend
    lblInfo.Caption = "Record " & lvData.SelectedItem.Index & " of " & lvData.ListItems.Count
  End If
  
  'Set Buttons
  If lvData.ListItems.Count > 0 And bUpdateable = True Then cmdEdit.Enabled = True
  If lvData.ListItems.Count > 0 And bUpdateable = True Then cmdDelete.Enabled = True
  
  EndMsg
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQLOutput, LoadGrid"
End Sub

Private Sub lvData_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Err_Handler
  lblInfo.Caption = "Record " & lvData.SelectedItem.Index & " of " & lvData.ListItems.Count
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQLOutput, lvData_ItemClick"
End Sub

Private Sub BuildEditBox()
On Error GoTo Err_Handler
Dim x As Integer
  lblField(0).Top = 3
  txtField(0).Top = lblField(0).Top + lblField(0).Height
  txtField(0).Width = picScroll.Width - 6
  lblField(0).Caption = lvData.ColumnHeaders(1).Text
  If lblField(0).Caption = "oid" Or _
     lblField(0).Caption = "cmax" Or _
     lblField(0).Caption = "xmax" Or _
     lblField(0).Caption = "cmin" Or _
     lblField(0).Caption = "xmin" Or _
     lblField(0).Caption = "ctid" Then
    txtField(0).Locked = True
  Else
    txtField(0).Locked = False
  End If
  For x = 2 To lvData.ColumnHeaders.Count
    Load lblField(x - 1)
    Load txtField(x - 1)
    lblField(x - 1).Visible = True
    txtField(x - 1).Visible = True
    lblField(x - 1).Top = txtField(x - 2).Top + txtField(x - 2).Height + 1
    txtField(x - 1).Top = lblField(x - 1).Top + lblField(x - 1).Height
    txtField(x - 1).Width = picScroll.Width - 6
    txtField(x - 1).TabIndex = txtField(x - 2).TabIndex + 1
    lblField(x - 1).Caption = lvData.ColumnHeaders(x).Text
    If lblField(x - 1).Caption = "oid" Or _
       lblField(0).Caption = "cmax" Or _
       lblField(0).Caption = "xmax" Or _
       lblField(0).Caption = "cmin" Or _
       lblField(0).Caption = "xmin" Or _
       lblField(0).Caption = "ctid" Then
      txtField(x - 1).Locked = True
    Else
      txtField(x - 1).Locked = False
    End If
  Next
  picScroll.Height = txtField(x - 2).Top + txtField(x - 2).Height + 1
  picEdit.Visible = True
  scScroll.Max = picScroll.ScaleHeight - picEdit.ScaleHeight
  cmdAdd.Visible = False
  cmdEdit.Visible = False
  cmdDelete.Visible = False
  cmdSave.Visible = True
  cmdCancel.Visible = True
  txtField(0).SetFocus
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQLOutput, BuildEditBox"
End Sub

Private Sub cmdEdit_Click()
On Error GoTo Err_Handler
Dim x As Long
  BuildEditBox
  For x = 0 To lvData.ColumnHeaders.Count - 1
    If x = 0 Then
      txtField(x).Text = lvData.SelectedItem.Text
    Else
      txtField(x).Text = lvData.SelectedItem.SubItems(x)
    End If
    txtField(x).Tag = ""
  Next
  lblInfo.Caption = "Edit Record"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQLOutput, cmdEdit_Click"
End Sub

Private Sub scScroll_Change()
On Error GoTo Err_Handler
  picScroll.Top = -scScroll.Value
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQLOutput, scScroll_Change"
End Sub

Private Sub txtField_Change(Index As Integer)
On Error GoTo Err_Handler
  txtField(Index).Tag = "Y"
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQLOutput, txtField_Change"
End Sub

Private Sub txtField_GotFocus(Index As Integer)
On Error GoTo Err_Handler
Dim x As Long
  For x = 0 To txtField.Count - 1
    If x = Index Then
      txtField(x).BackColor = &H8000000E
    Else
      txtField(x).BackColor = &H8000000F
    End If
  Next
  If txtField(Index).Top + txtField(Index).Height > picEdit.ScaleHeight - picScroll.Top Then
    If lblField(Index).Top > scScroll.Max Then
      picScroll.Top = scScroll.Max
      scScroll.Value = scScroll.Max
    Else
      picScroll.Top = -lblField(Index).Top
      scScroll.Value = -picScroll.Top
    End If
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQLOutput, txtField_GotFocus"
End Sub

Private Sub HideEditBox()
On Error GoTo Err_Handler
Dim x As Integer
  txtField(0).Text = ""
  txtField(0).Tag = ""
  For x = 2 To lvData.ColumnHeaders.Count
    Unload lblField(x - 1)
    Unload txtField(x - 1)
  Next
  cmdAdd.Visible = True
  cmdEdit.Visible = True
  cmdDelete.Visible = True
  cmdSave.Visible = False
  cmdCancel.Visible = False
  picEdit.Visible = False
  If lvData.ListItems.Count > 0 Then
    lblInfo.Caption = "Record " & lvData.SelectedItem.Index & " of " & lvData.ListItems.Count
  Else
    lblInfo.Caption = "Record 0 of 0"
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmSQLOutput, HideEditBox"
End Sub


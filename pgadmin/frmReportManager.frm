VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.1#0"; "HighlightBox.ocx"
Begin VB.Form frmReportManager 
   Caption         =   "Report Manager"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   Icon            =   "frmReportManager.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   8205
   Begin VB.Frame fraReport 
      Caption         =   "Report Details"
      Height          =   4020
      Left            =   4500
      TabIndex        =   5
      Top             =   0
      Width           =   3660
      Begin HighlightBox.HBX txtDescription 
         Height          =   3075
         Left            =   90
         TabIndex        =   7
         Top             =   855
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   5424
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Caption         =   "Description"
      End
      Begin VB.TextBox txtAuthor 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   495
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Author"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   270
         Width           =   465
      End
   End
   Begin MSComctlLib.ImageList ilBrowser 
      Left            =   405
      Top             =   2430
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportManager.frx":0442
            Key             =   "Manager"
            Object.Tag             =   "Manager"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportManager.frx":0894
            Key             =   "Folder"
            Object.Tag             =   "Folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportManager.frx":116E
            Key             =   "Report"
            Object.Tag             =   "Report"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvReports 
      Height          =   4005
      Left            =   1485
      TabIndex        =   3
      Top             =   0
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   7064
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "ilBrowser"
      Appearance      =   1
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Remove the selected report from pgAdmin"
      Top             =   765
      Width           =   1365
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Add a new report to pgAdmin."
      Top             =   405
      Width           =   1365
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "View the selected report."
      Top             =   45
      Width           =   1365
   End
End
Attribute VB_Name = "frmReportManager"
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

Private Sub trvReports_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Err_Handler
  If Button = 2 Then PopupMenu fMainForm.mnuCTXReportManager
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmReportManager, trvReports_MouseUp"
End Sub

Public Sub cmdAdd_Click()
On Error GoTo Err_Handler
  Load frmReportAdd
  frmReportAdd.Show
  frmReportAdd.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmReportManager, cmdAdd_Click"
End Sub

Public Sub cmdRemove_Click()
On Error GoTo Err_Handler
Dim szData As String
Dim fNum As Integer
Dim x As Integer
  If Mid(trvReports.SelectedItem.Key, 1, 4) <> "REP:" Then
    MsgBox "You must select a report!", vbExclamation, "Error"
    Exit Sub
  End If
  If MsgBox("Are you sure you wish to remove the selected report?", vbQuestion + vbYesNo, "Remove Report?") = vbNo Then Exit Sub
  StartMsg "Writing " & app.Path & "\Reports\Reports.dat"
  Dim fso As New FileSystemObject
  Dim fReports As File
  If fso.FileExists(app.Path & "\Reports\Reports.bak") Then
    Set fReports = fso.GetFile(app.Path & "\Reports\Reports.bak")
    fReports.Delete True
  End If
  Set fReports = fso.GetFile(app.Path & "\Reports\Reports.dat")
  fReports.Move (app.Path & "\Reports\Reports.bak")
  szData = ""
  fNum = FreeFile
  Open app.Path & "\Reports\Reports.dat" For Binary Access Write As #fNum
  For x = 1 To UBound(rptList)
    If x <> CInt(Mid(trvReports.SelectedItem.Key, 5, InStr(5, trvReports.SelectedItem.Key, ":") - 5)) Then
      szData = szData & rptList(x).szName & Chr(253) & rptList(x).szCategory & Chr(253) & rptList(x).szFile & Chr(253) & rptList(x).szSQL & Chr(253) & rptList(x).szAuthor & Chr(253) & rptList(x).szDescription & Chr(253)
      If rptList(x).bShowTree = True Then
        szData = szData & "1" & Chr(253)
      Else
        szData = szData & "0" & Chr(253)
      End If
      If rptList(x).bRefreshTables = True Then
        szData = szData & "1" & Chr(253)
      Else
        szData = szData & "0" & Chr(253)
      End If
      If rptList(x).bRefreshSequences = True Then
        szData = szData & "1" & Chr(254)
      Else
        szData = szData & "0" & Chr(254)
      End If
    End If
  Next
  Put #fNum, , szData
  Close #fNum
  Refresh_List
  EndMsg
  Refresh_List
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmReportManager, cmdView_Click"
End Sub

Public Sub cmdView_Click()
On Error GoTo Err_Handler
Dim rptViewer As New frmReportViewer
Dim iReport As Integer
  If Mid(trvReports.SelectedItem.Key, 1, 4) <> "REP:" Then
    MsgBox "You must select a report!", vbExclamation, "Error"
    Exit Sub
  End If
  iReport = CInt(Mid(trvReports.SelectedItem.Key, 5, InStr(5, trvReports.SelectedItem.Key, ":") - 5))
  Load rptViewer
  rptViewer.Show
  rptViewer.ZOrder 0
  If rptList(iReport).bRefreshSequences = True Then Update_SequenceCache
  If rptList(iReport).bRefreshTables = True Then Update_TableCache
  rptViewer.RunReport rptList(iReport).szFile, rptList(iReport).szSQL, rptList(iReport).bShowTree
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmReportManager, cmdView_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 8325
  Me.Height = 4455
  Refresh_List
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmReportManager, Form_Load"
End Sub

Public Sub Refresh_List()
On Error GoTo Err_Handler
Dim NodeX As Node
Dim fNum As Integer
Dim szData As String
Dim szEntries As Variant
Dim szEntry As Variant
Dim x As Integer
Dim y As Integer
Dim bFound As Boolean
  ReDim rptList(0)
  fNum = FreeFile
  Open app.Path & "\Reports\Reports.dat" For Binary Access Read As #fNum
  szData = Input(LOF(fNum), #fNum)
  Close #fNum
  If Len(szData) < 16 Then
      MsgBox "The Report Data file (" & app.Path & "\Reports\Reports.dat) is corrupt or missing!", vbCritical, "Error"
      Exit Sub
    End If
  szEntries = Split(szData, Chr(254), , vbBinaryCompare)
  For x = 0 To UBound(szEntries) - 1
    szEntry = Split(szEntries(x), Chr(253), , vbBinaryCompare)
    If UBound(szEntry) <> 8 Then
      MsgBox "The Report Data file (" & app.Path & "\Reports\Reports.dat) is corrupt!", vbCritical, "Error"
      Exit Sub
    End If
    ReDim Preserve rptList(UBound(rptList) + 1)
    rptList(UBound(rptList)).szName = szEntry(0)
    rptList(UBound(rptList)).szCategory = szEntry(1)
    If Mid(szEntry(2), 2, 2) <> ":\" Then
      rptList(UBound(rptList)).szFile = app.Path & "\Reports\" & szEntry(2)
    Else
      rptList(UBound(rptList)).szFile = szEntry(2)
    End If
    rptList(UBound(rptList)).szSQL = szEntry(3)
    rptList(UBound(rptList)).szAuthor = szEntry(4)
    rptList(UBound(rptList)).szDescription = szEntry(5)
    If szEntry(6) = "1" Then
      rptList(UBound(rptList)).bShowTree = True
    Else
      rptList(UBound(rptList)).bShowTree = False
    End If
    If szEntry(7) = "1" Then
      rptList(UBound(rptList)).bRefreshTables = True
    Else
      rptList(UBound(rptList)).bRefreshTables = False
    End If
    If szEntry(8) = "1" Then
      rptList(UBound(rptList)).bRefreshSequences = True
    Else
      rptList(UBound(rptList)).bRefreshSequences = False
    End If
  Next
  trvReports.Nodes.Clear
  Set NodeX = trvReports.Nodes.Add(, tvwChild, "ROOT::", "Categories", 1)
  For x = 1 To UBound(rptList)
    bFound = False
    For y = 1 To trvReports.Nodes.Count
      If trvReports.Nodes(y).Key = "CAT::" & rptList(x).szCategory Then
        bFound = True
        Exit For
      End If
    Next
    If bFound = False Then
      Set NodeX = trvReports.Nodes.Add("ROOT::", tvwChild, "CAT::" & rptList(x).szCategory, rptList(x).szCategory, 2)
    End If
    Set NodeX = trvReports.Nodes.Add("CAT::" & rptList(x).szCategory, tvwChild, "REP:" & x & ":" & rptList(x).szName, rptList(x).szName, 3)
  Next
  trvReports.Nodes(1).Expanded = True
  trvReports.Nodes(1).Selected = True
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmReportManager, Refresh_List"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  txtDescription.Minimise
  If Me.WindowState <> 1 Then
    If Me.Width < 8325 Then Me.Width = 8325
    If Me.Height < 4455 Then Me.Height = 4455
    trvReports.Height = Me.ScaleHeight
    trvReports.Width = Me.ScaleWidth - trvReports.Left - fraReport.Width - 25
    fraReport.Left = trvReports.Left + trvReports.Width + 25
    fraReport.Height = Me.ScaleHeight
    txtDescription.Height = fraReport.Height - txtDescription.Top - 100
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmReportManager, Form_Resize"
End Sub

Private Sub trvReports_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
Dim iReport As Integer
  Node.Expanded = True
  If Mid(Node.Key, 1, 4) = "REP:" Then
    iReport = CInt(Mid(Node.Key, 5, InStr(5, Node.Key, ":") - 5))
    txtAuthor.Text = rptList(iReport).szAuthor
    txtDescription.Text = rptList(iReport).szDescription
  Else
    txtAuthor.Text = ""
    txtDescription.Text = ""
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmReportManager, trvReports_NodeClick"
End Sub

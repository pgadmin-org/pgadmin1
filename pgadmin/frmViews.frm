VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Object = "{44DFA8BA-326E-4D0F-8941-25E814743439}#1.0#0"; "TreeToys.ocx"
Begin VB.Form frmViews 
   Caption         =   "Views"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   Icon            =   "frmViews.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4575
   ScaleWidth      =   8205
   Begin VB.CommandButton cmdRebuild 
      BackColor       =   &H80000018&
      Caption         =   "Rebuild &Project"
      Height          =   330
      Left            =   45
      TabIndex        =   19
      ToolTipText     =   "Checks and rebuilds project dependencies."
      Top             =   3510
      Width           =   1410
   End
   Begin VB.CommandButton cmdCopyDevToPro 
      BackColor       =   &H80000018&
      Caption         =   "Compile unsafe"
      Height          =   330
      Left            =   45
      TabIndex        =   18
      ToolTipText     =   "Compiles a repository function."
      Top             =   3870
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmdCopyProToDev 
      BackColor       =   &H80000018&
      Caption         =   "Load->Developt"
      Height          =   330
      Left            =   45
      TabIndex        =   17
      ToolTipText     =   "Compiles a repository function."
      Top             =   4230
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmdExportView 
      Caption         =   "Export View"
      Height          =   330
      Left            =   45
      TabIndex        =   15
      ToolTipText     =   "Modify the selected View."
      Top             =   1125
      Width           =   1410
   End
   Begin VB.CommandButton cmdModifyView 
      Caption         =   "&Modify View"
      Height          =   330
      Left            =   45
      TabIndex        =   14
      ToolTipText     =   "Modify the selected View."
      Top             =   405
      Width           =   1410
   End
   Begin VB.CommandButton cmdViewData 
      Caption         =   "&View Data"
      Height          =   330
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "Edit the comment for the selected View."
      Top             =   2205
      Width           =   1410
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "&Edit Comment"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Edit the comment for the selected View."
      Top             =   1485
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   45
      TabIndex        =   4
      ToolTipText     =   "Refresh the list of Views."
      Top             =   1845
      Width           =   1410
   End
   Begin VB.CommandButton cmdDropView 
      Caption         =   "&Drop View"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Delete the selected View."
      Top             =   765
      Width           =   1410
   End
   Begin VB.CommandButton cmdCreateView 
      Caption         =   "&Create View"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Create a new View."
      Top             =   45
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show System:"
      Height          =   525
      Left            =   45
      TabIndex        =   10
      Top             =   2970
      Width           =   1380
      Begin VB.CheckBox chkSystem 
         Caption         =   "Views"
         Height          =   225
         Left            =   135
         TabIndex        =   5
         ToolTipText     =   "Select to view system views"
         Top             =   225
         Width           =   1065
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   45
      Top             =   2610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select SQL File"
      Filter          =   "All Files (*.*)|*.*"
   End
   Begin TreeToys.TreeToy trvBrowser 
      Height          =   4560
      Left            =   1485
      TabIndex        =   16
      Top             =   0
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   8043
      NodeTips        =   1
      BorderStyle     =   1
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      Indentation     =   299,906
      LabelEdit       =   1
      LineStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ilBrowser 
      Left            =   540
      Top             =   2610
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViews.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViews.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViews.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViews.frx":08D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViews.frx":0A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViews.frx":0B8C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraDetails 
      Caption         =   "View Details"
      Height          =   4560
      Left            =   4500
      TabIndex        =   8
      Top             =   0
      Width           =   3660
      Begin HighlightBox.HBX txtDefinition 
         Height          =   1815
         Left            =   90
         TabIndex        =   20
         Top             =   1215
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   3201
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
         Caption         =   "Definition"
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   2670
      End
      Begin VB.TextBox txtACL 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   855
         Width           =   2670
      End
      Begin VB.TextBox txtOwner 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   2670
      End
      Begin HighlightBox.HBX txtComments 
         Height          =   1365
         Left            =   90
         TabIndex        =   21
         Top             =   3105
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   2408
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
         Caption         =   "Comments"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ACL"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   11
         Top             =   900
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   9
         Top             =   585
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmViews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Public dragNode As Node, dropNode As Node

Private iPro_Index As Integer
Private iDev_Index As Integer
Private iSys_Index As Integer

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Form
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++s++++++++++++++++++++++++++

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 8325
  Me.Height = 4980
  txtDefinition.Wordlist = TextColours
  Set trvBrowser.ImageList = ilBrowser
  cmdRefresh_Click

Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  txtDefinition.Minimise
  txtComments.Minimise
  If Me.WindowState <> 1 And Me.ScaleHeight > 0 Then
    If Me.WindowState = 0 Then
      If Me.Width < 8325 Then Me.Width = 8325
      If Me.Height < 4980 Then Me.Height = 4980
    End If
    
    
        trvBrowser.Height = Me.ScaleHeight
        trvBrowser.Width = Me.ScaleWidth - trvBrowser.Left - fraDetails.Width - 25
        fraDetails.Left = trvBrowser.Left + trvBrowser.Width + 25
        fraDetails.Height = Me.ScaleHeight
        txtComments.Height = fraDetails.Height - txtComments.Top - 100

  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, Form_Resize"
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Buttons
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub cmdExportView_Click()
On Error GoTo Err_Handler

    cmp_view_tree_export trvBrowser, CommonDialog1
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, cmdExportView_Click"
End Sub

Public Sub cmdModifyView_Click()
On Error GoTo Err_Handler

Dim szView_name As String

szView_name = trvBrowser.SelectedItem.Text & ""
If szView_name <> "" Then
    gView_Name = szView_name
    Load frmAddView
    frmAddView.Show
End If

Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, cmdModifyView_Click"
End Sub

Private Sub cmdRebuild_Click()
On Error GoTo Err_Handler
    cmp_Project_Rebuild
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, cmdRebuildProject_Click"
End Sub

Private Sub cmdCopyDevToPro_Click()
On Error GoTo Err_Handler
    cmp_view_tree_copy_devtopro trvBrowser
    cmdRefresh_Click
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdCopyDevToPro_Click"
End Sub

Private Sub cmdCopyProToDev_Click()
On Error GoTo Err_Handler
    cmp_view_tree_copy_protodev trvBrowser
    cmdRefresh_Click
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmFunctions, cmdCopyProToDev_Click"
End Sub

Private Sub chkSystem_Click()
On Error GoTo Err_Handler
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, ChkSystem_Click"
End Sub

Public Sub cmdViewData_Click()
On Error GoTo Err_Handler
Dim Response As Integer
Dim Tuples As Long
Dim rsQuery As New Recordset
Dim szView_name As String
Dim szQuery As String

  szView_name = trvBrowser.SelectedItem.Text & ""
  If szView_name = "" Then
    MsgBox "You must select a view to view!", vbExclamation, "Error"
    Exit Sub
  End If
  szQuery = "SELECT count(*) As records FROM " & QUOTE & szView_name & QUOTE
  Tuples = CLng(RsExecuteGetResult(szQuery))
  If Tuples > 1000 Then
    Response = MsgBox("That table contains " & Tuples & " rows which may take some time to load! Do you wish to continue?", _
    vbExclamation + vbYesNo, "Warning")
    If Response = vbNo Then Exit Sub
  End If
  
  Dim DataForm As New frmSQLOutput
  LogMsg "Executing: SELECT * FROM " & QUOTE & szView_name & QUOTE
  rsQuery.Open "SELECT * FROM " & QUOTE & szView_name & QUOTE, gConnection, adOpenForwardOnly, adLockReadOnly
  Load DataForm
  DataForm.Display rsQuery
  DataForm.Show
  DataForm.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, cmdViewData_Click"
End Sub

Public Sub cmdComment_Click()
On Error GoTo Err_Handler
'Dim szView_name As String

    'szView_name = trvBrowser.SelectedItem.Text & ""
    
    'If szView_name = "" Then
    '  MsgBox "You must select a View to edit the comment for.", vbExclamation, "Error"
    '  Exit Sub
    'End If
    'Load frmComments
    'frmComments.Setup "frmViews", QUOTE & szView_name & QUOTE, txtOID.Text
    'frmComments.Show
  
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, cmdComment_Click"
End Sub

Public Sub cmdCreateView_Click()
On Error GoTo Err_Handler
  gView_Name = ""
  Load frmAddView
  frmAddView.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmViews, cmdCreateView_Click"
End Sub

Public Sub cmdDropView_Click()
On Error GoTo Err_Handler
Dim szText As String

szText = trvBrowser.TreeTextChecked
If MsgBox("Are you sure you wish to drop:" & vbCrLf & vbCrLf & szText, vbYesNo + vbQuestion, _
            "Confirm View(s) Delete") = vbYes Then
        StartMsg "Dropping View..."
        
        cmp_view_tree_drop trvBrowser
        cmdRefresh_Click
  End If
  
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmViews, cmdDropView_Click"
End Sub

Public Sub cmdRefresh_Click()
On Error GoTo Err_Handler

cmp_view_tree_refresh trvBrowser, CBool(chkSystem), iPro_Index, iSys_Index, iDev_Index

cmdCopyDevToPro.Visible = DevMode
cmdCopyProToDev.Visible = DevMode
cmdRebuild.Visible = DevMode

CmdViewButton

Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmViews, cmdRefresh_Click"
End Sub

Public Sub CmdViewButton()
On Error GoTo Err_Handler

cmdButtonActivate trvBrowser, CBool(chkSystem), iPro_Index, iSys_Index, iDev_Index, cmdCreateView, cmdModifyView, cmdDropView, cmdExportView, cmdComment, cmdRefresh, cmdViewData

Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmViews, CmdViewButton"
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Treeview
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub trvBrowser_NodeCheck(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler

trvBrowser.FreezeCtl
trvBrowser.TreeSelectiveCheck Node
trvBrowser.UnFreezeCtl
    
CmdViewButton

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmviews, trvBrowser_NodeCheck"
End Sub

Private Sub trvBrowser_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler

    Dim szview_table As String
    Dim szView_name As String
    Dim szView_definition As String
    Dim szView_owner As String
    Dim szView_acl As String
    Dim szView_comments As String
        
    '----------------------------------------------------------------------------------
    ' Retrieve view name and arguments from List
    '----------------------------------------------------------------------------------
    Dim szRoot As String
    If Node.Text <> "" Then
        szView_name = Node.Text
        szRoot = Left(Node.Key, 2)
        If szRoot = "P:" Or szRoot = "S:" Then
            szview_table = "pgadmin_views"
        Else
            szview_table = gDevPostgresqlTables & "_views"
        End If
    Else
        szView_name = ""
    End If
    '----------------------------------------------------------------------------------
    ' Lookup database
    '----------------------------------------------------------------------------------
    StartMsg "Retrieving view Info..."
    cmp_View_GetValues szview_table, szView_name, szView_definition, szView_owner, szView_acl, szView_comments
    
    txtOwner.Text = szView_owner
    txtACL.Text = szView_acl
    txtName.Text = szView_name
    txtDefinition.Text = szView_definition
    txtComments.Text = szView_comments
    
    
    CmdViewButton
    EndMsg
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmviews, trvBrowser_NodeClick"
End Sub

Private Sub trvBrowser_dblClick()
On Error GoTo Err_Handler

    If trvBrowser.SelectedItem.Parent Is Nothing Then Exit Sub
    If DevMode = True And trvBrowser.SelectedItem.Parent.Key = "Pro:" Or trvBrowser.SelectedItem.Parent.Key = "Sys:" Then Exit Sub

    cmdModifyView_Click
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmviews, trvBrowser_dblClick"
End Sub

Private Sub trvBrowser_OLEStartDrag(Data As MSComctlLib.DataObject, _
AllowedEffects As Long)
On Error GoTo Err_Handler

Set dragNode = trvBrowser.SelectedItem

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmviews, trvBrowser_OLEStartDrag"
End Sub

Private Sub trvBrowser_MouseDown(Button As Integer, Shift As Integer, _
x As Single, y As Single)
On Error GoTo Err_Handler

    Set dragNode = trvBrowser.HitTest(x, y)
    Set dropNode = Nothing
    If Not (dragNode Is Nothing) Then
        dragNode.Selected = True
        trvBrowser_NodeClick dragNode
    End If
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmviews, trvBrowser_MouseDown"
End Sub

Private Sub trvBrowser_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Err_Handler

Dim sz_drag_key As String
Dim sz_drop_key As String

    Set dropNode = trvBrowser.HitTest(x, y)

    If Not (dragNode Is Nothing) And Not (dropNode Is Nothing) Then
        If dragNode.Key <> dropNode.Key Then
            If dragNode.Parent Is Nothing Then
               sz_drag_key = dragNode.Key
            Else
               sz_drag_key = dragNode.Parent.Key
            End If
            
            If dropNode.Parent Is Nothing Then
               sz_drop_key = dropNode.Key
            Else
               sz_drop_key = dropNode.Parent.Key
            End If
            
            Select Case sz_drag_key
                Case "Pro:"
                    If (sz_drop_key = "Dev:") Then
                            cmdCopyProToDev_Click
                    End If
                
                Case "Sys:"
                    If (sz_drop_key = "Dev:") Then
                            cmdCopyProToDev_Click
                    End If
                    
                 Case "Dev:"
                    If (sz_drop_key = "Pro:" Or sz_drop_key = "Sys:") Then
                       cmdCopyDevToPro_Click
                    End If
            End Select
         Else
            If Not (dropNode Is Nothing) Then
                trvBrowser_NodeCheck dropNode
            End If
        End If
    End If
    
    Set dragNode = Nothing
    Set dropNode = Nothing
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmviews, trvBrowser_OLEDragDrop"
End Sub

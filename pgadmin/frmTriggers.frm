VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Object = "{44DFA8BA-326E-4D0F-8941-25E814743439}#1.0#0"; "TreeToys.ocx"
Begin VB.Form frmTriggers 
   Caption         =   "Triggers"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   Icon            =   "frmTriggers.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7266.234
   ScaleMode       =   0  'User
   ScaleWidth      =   9740.402
   Begin VB.CommandButton cmdRebuild 
      BackColor       =   &H80000018&
      Caption         =   "Rebuild &Project"
      Height          =   330
      Left            =   45
      TabIndex        =   24
      ToolTipText     =   "Checks and rebuilds project dependencies."
      Top             =   3555
      Width           =   1410
   End
   Begin VB.CommandButton cmdCopyDevToPro 
      BackColor       =   &H80000018&
      Caption         =   "Compile unsafe"
      Height          =   330
      Left            =   45
      TabIndex        =   23
      ToolTipText     =   "Compiles a repository function."
      Top             =   3915
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmdCopyProToDev 
      BackColor       =   &H80000018&
      Caption         =   "Load->Developt"
      Height          =   330
      Left            =   45
      TabIndex        =   22
      ToolTipText     =   "Compiles a repository function."
      Top             =   4275
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmdExportTrig 
      Caption         =   "Export Trigger"
      Height          =   330
      Left            =   45
      TabIndex        =   18
      ToolTipText     =   "Modify the selected trigger."
      Top             =   1125
      Width           =   1410
   End
   Begin VB.CommandButton cmdModifyTrig 
      Caption         =   "&Modify Trigger"
      Height          =   330
      Left            =   45
      TabIndex        =   17
      ToolTipText     =   "Modify the selected trigger."
      Top             =   405
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show System:"
      Height          =   525
      Left            =   45
      TabIndex        =   16
      Top             =   2970
      Width           =   1380
      Begin VB.CheckBox chkSystem 
         Caption         =   "Triggers"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Select to view system triggers."
         Top             =   225
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "&Edit Comment"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Edit the comment for the selected Trigger."
      Top             =   1485
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "Refresh the list of Triggers."
      Top             =   1845
      Width           =   1410
   End
   Begin VB.CommandButton cmdDropTrig 
      Caption         =   "&Drop Trigger"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Delete the selected Trigger."
      Top             =   765
      Width           =   1410
   End
   Begin VB.CommandButton cmdCreateTrig 
      Caption         =   "&Create Trigger"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Create a new Trigger."
      Top             =   45
      Width           =   1410
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   45
      Top             =   2250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select SQL File"
      Filter          =   "All Files (*.*)|*.*"
   End
   Begin TreeToys.TreeToy trvBrowser 
      Height          =   5505
      Left            =   1530
      TabIndex        =   21
      Top             =   45
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   9710
      NodeTips        =   1
      BorderStyle     =   1
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HideSelection   =   0   'False
      Indentation     =   566.929
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
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
      Top             =   2250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriggers.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriggers.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriggers.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriggers.frx":08D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Trigger Details"
      Height          =   5550
      Left            =   4680
      TabIndex        =   10
      Top             =   0
      Width           =   4200
      Begin HighlightBox.HBX txtComments 
         Height          =   3255
         Left            =   90
         TabIndex        =   25
         Top             =   2205
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   5741
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
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   270
         Width           =   3300
      End
      Begin VB.TextBox txtForEach 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1845
         Width           =   3300
      End
      Begin VB.TextBox txtEvent 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1530
         Width           =   3300
      End
      Begin VB.TextBox txtExecutes 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1215
         Width           =   3300
      End
      Begin VB.TextBox txtFunction 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   900
         Width           =   3300
      End
      Begin VB.TextBox txtTable 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   585
         Width           =   3300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   20
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Table"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   630
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Function"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   14
         Top             =   945
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Executes"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   13
         Top             =   1260
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Event"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   12
         Top             =   1575
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "For Each"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   11
         Top             =   1890
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmTriggers"
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
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 9000
  Me.Height = 6000
  
  Set trvBrowser.ImageList = ilBrowser
  cmdRefresh_Click

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmTriggers, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  txtComments.Minimise
  If Me.WindowState <> 1 And Me.ScaleHeight > 0 Then
    If Me.WindowState = 0 Then
      If Me.Width < 9000 Then Me.Width = 9000
      If Me.Height < 6000 Then Me.Height = 6000
    End If
    
    trvBrowser.Height = Me.ScaleHeight
    trvBrowser.Width = Me.ScaleWidth - trvBrowser.Left - fraDetails.Width - 25
    fraDetails.Left = trvBrowser.Left + trvBrowser.Width + 25
    fraDetails.Height = Me.ScaleHeight
    txtComments.Height = fraDetails.Height - txtComments.Top - 100

  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, Form_Resize"
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Buttons
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub cmdCopyDevToPro_Click()
On Error GoTo Err_Handler
    cmp_Trigger_tree_copy_devtopro trvBrowser
    cmdRefresh_Click
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdCopyDevToPro_Click"
End Sub

Private Sub cmdCopyProToDev_Click()
On Error GoTo Err_Handler
    cmp_Trigger_tree_copy_protodev trvBrowser
    cmdRefresh_Click
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdCopyProToDev_Click"
End Sub

Private Sub cmdExportTrig_Click()
On Error GoTo Err_Handler
    cmp_trigger_tree_export trvBrowser, CommonDialog1
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdExportTrig_Click"
End Sub

Public Sub cmdModifyTrig_Click()
On Error GoTo Err_Handler
    If txtName <> "" Then
        ModifyTrigger txtName & " ON " & txtTable
    End If
    
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdModifyTrig_Click"
End Sub

Private Sub ModifyTrigger(szTrigger As String)
On Error GoTo Err_Handler

If szTrigger <> "" Then
    ' Get name and arguments
    cmp_Trigger_ParseName szTrigger, gTrigger_Name, gTrigger_Table
     
    ' Load form
    Load frmAddTrigger
    frmAddTrigger.Show
End If

Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, ModifyTrigger"
End Sub

Private Sub cmdRebuild_Click()
On Error GoTo Err_Handler

    cmp_Project_Rebuild
    cmdRefresh_Click
Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdRebuild_Click"
End Sub

Private Sub chkSystem_Click()
On Error GoTo Err_Handler
  cmdRefresh_Click

Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, chkFunctions_Click"
End Sub

Public Sub cmdComment_Click()
On Error GoTo Err_Handler
  'If txtOID.Text = "" Then
  '  MsgBox "You must select a function to edit the comment for.", vbExclamation, "Error"
  '  Exit Sub
  'End If
  'Load frmComments
  'frmComments.Setup "frmTriggers", QUOTE & txtName.Text & QUOTE & "(" & txtTable.Text & ")", Val(txtOID.Text)
  'frmComments.Show

Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdComment_Click"
End Sub

Public Sub cmdCreateTrig_Click()
On Error GoTo Err_Handler
  ' This means we will create the function
  gTrigger_Name = ""
  gTrigger_Table = ""
  
  ' Load form
  Load frmAddTrigger
  frmAddTrigger.Show

Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdCreateTrig_Click"
End Sub

Public Sub cmdDropTrig_Click()
On Error GoTo Err_Handler
Dim szText As String

szText = trvBrowser.TreeTextChecked
If MsgBox("Are you sure you wish to drop:" & vbCrLf & vbCrLf & szText, vbYesNo + vbQuestion, _
            "Connfirm?") = vbYes Then
    cmp_Trigger_tree_drop trvBrowser
    cmdRefresh_Click
End If

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdDropTrig_Click"
End Sub

Public Sub cmdRefresh_Click()
On Error GoTo Err_Handler

cmp_Trigger_tree_refresh trvBrowser, CBool(chkSystem), iPro_Index, iSys_Index, iDev_Index

cmdCopyDevToPro.Visible = DevMode
cmdCopyProToDev.Visible = DevMode
cmdRebuild.Visible = DevMode

CmdTriggerButton

  Exit Sub
Err_Handler:
EndMsg
If Err.Number <> 0 Then LogError Err, "frmTriggers, cmdRefresh_Click"
End Sub

Public Sub CmdTriggerButton()
On Error GoTo Err_Handler

cmdButtonActivate trvBrowser, CBool(chkSystem), iPro_Index, iSys_Index, iDev_Index, cmdCreateTrig, cmdModifyTrig, cmdDropTrig, cmdExportTrig, cmdComment, cmdRefresh

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmTriggers, CmdFuncButton"
End Sub

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Treeview
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub trvBrowser_NodeCheck(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler

trvBrowser.FreezeCtl
trvBrowser.TreeSelectiveCheck Node
trvBrowser.UnFreezeCtl
    
CmdTriggerButton

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmTriggers, trvBrowser_NodeCheck"
End Sub

Private Sub trvBrowser_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_Handler
    
    Dim szTrigger_pgTable As String
    Dim szTrigger_Name As String
    Dim szTrigger_Table As String
    Dim szTrigger_Function As String
    Dim szTrigger_Arguments As String
    Dim szTrigger_Foreach As String
    Dim szTrigger_Event As String
    Dim szTrigger_Executes As String
    Dim szTrigger_Comments As String
    
    '----------------------------------------------------------------------------------
    ' Retrieve Trigger name and arguments from List
    '----------------------------------------------------------------------------------
    Dim szRoot As String
    If Node.Text <> "" Then
        cmp_Trigger_ParseName Node.Text, szTrigger_Name, szTrigger_Table
        szRoot = Left(Node.Key, 2)
        If szRoot = "P:" Or szRoot = "S:" Then
            szTrigger_pgTable = "pgadmin_Triggers"
        Else
            szTrigger_pgTable = gDevPostgresqlTables & "_Triggers"
        End If
    Else
        szTrigger_Name = ""
    End If
    '----------------------------------------------------------------------------------
    ' Lookup database
    '----------------------------------------------------------------------------------
    StartMsg "Retrieving Trigger Info..."
    cmp_Trigger_GetValues szTrigger_pgTable, szTrigger_Name, szTrigger_Table, szTrigger_Function, szTrigger_Arguments, szTrigger_Foreach, szTrigger_Executes, szTrigger_Event, szTrigger_Comments
    
   
    txtName.Text = szTrigger_Name
    txtTable.Text = szTrigger_Table
    If szTrigger_Function <> "" Then
        txtFunction.Text = szTrigger_Function & " (" & szTrigger_Arguments & ")"
    Else
        txtFunction.Text = ""
    End If
    txtForEach.Text = szTrigger_Foreach
    txtExecutes.Text = szTrigger_Executes
    txtEvent.Text = szTrigger_Event
    txtComments.Text = szTrigger_Comments
       
    CmdTriggerButton
    EndMsg
    
Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmTriggers, trvBrowser_NodeClick"
End Sub

Private Sub trvBrowser_dblClick()
On Error GoTo Err_Handler

    If trvBrowser.SelectedItem.Parent Is Nothing Then Exit Sub
    If DevMode = True And trvBrowser.SelectedItem.Parent.Key = "Pro:" Or trvBrowser.SelectedItem.Parent.Key = "Sys:" Then Exit Sub

    cmdModifyTrig_Click

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmTriggers, trvBrowser_dblClick"
End Sub

Private Sub trvBrowser_OLEStartDrag(Data As MSComctlLib.DataObject, _
AllowedEffects As Long)
On Error GoTo Err_Handler

Set dragNode = trvBrowser.SelectedItem

Exit Sub
Err_Handler:
If Err.Number <> 0 Then LogError Err, "frmTriggers, trvBrowser_OLEStartDrag"
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
If Err.Number <> 0 Then LogError Err, "frmTriggers, trvBrowser_MouseDown"
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
If Err.Number <> 0 Then LogError Err, "frmTriggers, trvBrowser_OLEDragDrop"
End Sub


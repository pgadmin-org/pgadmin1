VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{44DFA8BA-326E-4D0F-8941-25E814743439}#1.0#0"; "TreeToys.ocx"
Begin VB.Form frmMain 
   Caption         =   "TreeToys Demo"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10065
   LinkTopic       =   "frmMain"
   ScaleHeight     =   7920
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin TreeToys.TreeToy TreeToy1 
      Height          =   7845
      Left            =   2790
      TabIndex        =   15
      Top             =   45
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   13838
      Indentation     =   566.929
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
   Begin VB.CheckBox Check3 
      Caption         =   $"frmMain.frx":0000
      Height          =   495
      Left            =   60
      TabIndex        =   14
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "NodeTips"
      Height          =   2235
      Left            =   45
      TabIndex        =   7
      Top             =   0
      Width           =   1335
      Begin VB.OptionButton Option1 
         Caption         =   "Key"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1155
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show Icons"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1860
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Path"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1140
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Text"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tag"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   540
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "None"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   225
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Scroll Tips"
      Height          =   2235
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      Begin VB.OptionButton Option2 
         Caption         =   "Key"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1155
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show Icons"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1155
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Path"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1140
         Width           =   1155
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Text"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1155
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Tag"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   1155
      End
      Begin VB.OptionButton Option2 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   270
      Top             =   3015
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":003C
            Key             =   "keyPar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0358
            Key             =   "keyChd"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' TreeToy
' Copyright (C) 2001, Jaroslaw Zwierz, AVE
' Poland, 04-628 Warszawa, ul. Alpejska 38
' tel./fax (+ 48 22) 815 68 99
' email: jerry@ave.com.pl

' Converted to an ActiveX Control by Dave Page (dpage@vale-housing.co.uk)
' for the pgAdmin project (http://www.greatbridge.org/project/pgadmin/)

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

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        TreeToy1.ShowIconsInNodeTips = True
    Else
        TreeToy1.ShowIconsInNodeTips = False
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = vbChecked Then
        TreeToy1.ShowIconsInScrollTips = True
    Else
        TreeToy1.ShowIconsInScrollTips = False
    End If
End Sub

Private Sub Check3_Click()

Dim nodX As Node

    'Reset current selection
    TreeToy1.FreezeCtl
    For Each nodX In TreeToy1.Nodes
        nodX.Bold = False
        nodX.Checked = False
    Next
    TreeToy1.UnFreezeCtl
    
End Sub

Private Sub Form_Load()
    'Start ToolTip Bonanza
    Set TreeToy1.ImageList = ImageList1
    AddSomeData TreeToy1
    

    
    
End Sub


Private Sub AddSomeData(t As treetoy)

Dim Node1 As Node
Dim Node2 As Node
Dim Node3 As Node
Dim Node4 As Node
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
  
  ' Fill up the treeview...
  For i = 1 To 5
    Set Node1 = TreeToy1.Nodes.Add(, , , "Root" & i, 1)
    Node1.Tag = "Essai Essai Essai Essai Essai Essai Essai"
    Node1.Tag = Node1.Tag & vbCrLf & "Essai Essai Essai Essai Essai "
    Node1.Tag = Node1.Tag & vbCrLf & "Essai Essai Essai Essai Essai "
    Node1.Tag = Node1.Tag & vbCrLf & "Essai Essai Essai Essai Essai "
    Node1.Tag = Node1.Tag & vbCrLf & "Essai Essai Essai Essai Essai "
    Node1.Tag = Node1.Tag & vbCrLf & "Essai Essai Essai Essai Essai "
    Node1.Tag = Node1.Tag & vbCrLf & "Essai Essai Essai Essai Essai "
    Node1.Tag = Node1.Tag & vbCrLf & "Essai Essai Essai Essai Essai "
    Node1.Text = Node1.Text
    For j = 1 To 5
      Set Node2 = TreeToy1.Nodes.Add(Node1.Index, tvwChild, , "Root" & i & "Child" & j, 2)
      Node2.Tag = Node1.Tag & vbCr & Node2.Text
      For k = 1 To 5
        Set Node3 = TreeToy1.Nodes.Add(Node2.Index, tvwChild, , "GrandChild" & (16 * (i - 1)) + (4 * (j - 1)) + k, 2)
        Node3.Tag = Node2.Tag & vbCr & Node3.Text
        For l = 1 To 5
            Set Node4 = TreeToy1.Nodes.Add(Node3.Index, tvwChild, , "SubChild" & (16 * (i - 1)) + (4 * (j - 1)) + k + l - 1, 2)
            Node4.Tag = Node3.Tag & vbCr & Node4.Text
        Next
      Next
    Next
  
    
  Next
    
    For Each Node1 In t.Nodes
        Node1.Expanded = True
    Next
    
End Sub


Private Sub Form_Resize()

Dim sw As Long
Dim sh As Long

    With TreeToy1
        sw = ScaleWidth - .Left - 60
        If sw > 0 Then
            .Width = sw
        End If
        
        sh = ScaleHeight - .Top - 60
        If sh > 0 Then
            .Height = sh
        End If
    End With
    
End Sub

Private Sub Option1_Click(Index As Integer)

    TreeToy1.NodeTips = Index
    
End Sub

Private Sub Option2_Click(Index As Integer)

    TreeToy1.ScrollTips = Index
    
End Sub

Private Sub TreeToy1_NodeCheck(ByVal Node As Node)

    If Check3.Value = vbChecked Then
        TreeToy1.FreezeCtl
        TreeToy1.TreeSelectiveCheck Node
        TreeToy1.UnFreezeCtl
    End If
    
End Sub

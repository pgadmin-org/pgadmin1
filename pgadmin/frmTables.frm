VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmTables 
   Caption         =   "Tables"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   Icon            =   "frmTables.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleWidth      =   8205
   Begin VB.CommandButton cmdSerialize 
      Caption         =   "&Serialize Column"
      Height          =   330
      Left            =   45
      TabIndex        =   75
      ToolTipText     =   "Convert this column into a (non primary key) serial column."
      Top             =   2205
      Width           =   1380
   End
   Begin VB.CommandButton cmdAddColumn 
      Caption         =   "&Add Column"
      Height          =   330
      Left            =   45
      TabIndex        =   5
      ToolTipText     =   "Add a new column to the selected table."
      Top             =   1845
      Width           =   1380
   End
   Begin VB.CommandButton cmdRenColumn 
      Caption         =   "Re&name Column"
      Height          =   330
      Left            =   45
      TabIndex        =   4
      ToolTipText     =   "Rename the selected column"
      Top             =   1485
      Width           =   1380
   End
   Begin VB.CommandButton cmdData 
      Caption         =   "View Da&ta"
      Height          =   330
      Left            =   45
      TabIndex        =   3
      ToolTipText     =   "View the data in the selected table"
      Top             =   1125
      Width           =   1380
   End
   Begin VB.CommandButton cmdRenTable 
      Caption         =   "&Rename Table"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Rename the selected table"
      Top             =   405
      Width           =   1380
   End
   Begin VB.CommandButton cmdDropTable 
      Caption         =   "&Drop Table"
      Height          =   330
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Drop the selected table"
      Top             =   765
      Width           =   1380
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "&Edit Comment"
      Height          =   330
      Left            =   45
      TabIndex        =   6
      ToolTipText     =   "Edit the comment for the selected object."
      Top             =   2565
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show System:"
      Height          =   705
      Left            =   45
      TabIndex        =   42
      Top             =   3285
      Width           =   1380
      Begin VB.CheckBox chkFields 
         Caption         =   "Columns"
         Height          =   225
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Select to view system columns"
         Top             =   420
         Width           =   1065
      End
      Begin VB.CheckBox chkTables 
         Caption         =   "Tables"
         Height          =   225
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Select to view system tables"
         Top             =   210
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Re&fresh"
      Height          =   330
      Left            =   45
      TabIndex        =   7
      ToolTipText     =   "Reload the table definitions from the database"
      Top             =   2925
      Width           =   1380
   End
   Begin VB.CommandButton cmdAddTable 
      Caption         =   "&Create Table"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Create a new table"
      Top             =   45
      Width           =   1380
   End
   Begin MSComctlLib.ImageList ilBrowser 
      Left            =   1680
      Top             =   2775
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTables.frx":030A
            Key             =   "Server"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTables.frx":0624
            Key             =   "Table"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTables.frx":093E
            Key             =   "Column"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTables.frx":0C58
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTables.frx":0DB2
            Key             =   "ForeignKey"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTables.frx":0F0C
            Key             =   "PrimaryKey"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTables.frx":1066
            Key             =   "Unique"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvBrowser 
      Height          =   4005
      Left            =   1485
      TabIndex        =   10
      ToolTipText     =   "Browse the table and column definitions"
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
   Begin VB.Frame fraUnique 
      Caption         =   "Unique Constraint Details"
      Height          =   3975
      Left            =   4500
      TabIndex        =   72
      Top             =   0
      Width           =   3660
      Begin HighlightBox.HBX txtUniqueColumns 
         Height          =   3300
         Left            =   90
         TabIndex        =   34
         Top             =   585
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   5821
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
         Caption         =   "Columns"
      End
      Begin VB.TextBox txtUniqueOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   225
         Width           =   2400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Index OID"
         Height          =   195
         Index           =   35
         Left            =   90
         TabIndex        =   74
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.Frame fraPrimary 
      Caption         =   "Primary Key Details"
      Height          =   3975
      Left            =   4500
      TabIndex        =   71
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtPrimaryOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   225
         Width           =   2400
      End
      Begin HighlightBox.HBX txtPrimaryColumns 
         Height          =   3300
         Left            =   90
         TabIndex        =   76
         Top             =   585
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   5821
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
         Caption         =   "Columns"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Index OID"
         Height          =   195
         Index           =   33
         Left            =   90
         TabIndex        =   73
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.Frame fraForeign 
      Caption         =   "Foreign Key Details"
      Height          =   3975
      Left            =   4500
      TabIndex        =   68
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtForeignTable 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   540
         Width           =   2400
      End
      Begin VB.TextBox txtForeignOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   225
         Width           =   2400
      End
      Begin HighlightBox.HBX txtLocalColumns 
         Height          =   1500
         Left            =   90
         TabIndex        =   38
         Top             =   855
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   2646
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
         Caption         =   "Local Columns"
      End
      Begin HighlightBox.HBX txtForeignColumns 
         Height          =   1455
         Left            =   90
         TabIndex        =   39
         Top             =   2430
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   2566
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
         Caption         =   "Foreign Columns"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Foreign Table"
         Height          =   195
         Index           =   27
         Left            =   90
         TabIndex        =   70
         Top             =   585
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trigger OID"
         Height          =   195
         Index           =   25
         Left            =   90
         TabIndex        =   69
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.Frame fraCheck 
      Caption         =   "Check Details"
      Height          =   3975
      Left            =   4500
      TabIndex        =   66
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtCheckOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   225
         Width           =   2400
      End
      Begin HighlightBox.HBX txtCheckDefinition 
         Height          =   3300
         Left            =   90
         TabIndex        =   41
         Top             =   585
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   5821
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Check OID"
         Height          =   195
         Index           =   28
         Left            =   90
         TabIndex        =   67
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.Frame fraColumn 
      Caption         =   "Column Details"
      Height          =   3975
      Left            =   4500
      TabIndex        =   43
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtNotNull 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1485
         Width           =   2400
      End
      Begin VB.TextBox txtDefault 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1800
         Width           =   2400
      End
      Begin VB.TextBox txtLength 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   855
         Width           =   2400
      End
      Begin VB.TextBox txtNumber 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   540
         Width           =   2400
      End
      Begin VB.TextBox txtColOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   225
         Width           =   2400
      End
      Begin VB.TextBox txtType 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1170
         Width           =   2400
      End
      Begin HighlightBox.HBX txtColComments 
         Height          =   1725
         Left            =   90
         TabIndex        =   32
         Top             =   2160
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   3043
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
         Caption         =   "Default"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   49
         Top             =   1845
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Not Null"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   48
         Top             =   1530
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   195
         Index           =   18
         Left            =   90
         TabIndex        =   47
         Top             =   1215
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Length"
         Height          =   195
         Index           =   19
         Left            =   90
         TabIndex        =   46
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Number"
         Height          =   195
         Index           =   20
         Left            =   90
         TabIndex        =   45
         Top             =   585
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   21
         Left            =   90
         TabIndex        =   44
         Top             =   270
         Width           =   285
      End
   End
   Begin VB.Frame fraDatasource 
      Caption         =   "Datasource Details"
      Height          =   3975
      Left            =   4500
      TabIndex        =   55
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtCompiler 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1485
         Width           =   2400
      End
      Begin VB.TextBox txtPlatform 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1170
         Width           =   2400
      End
      Begin VB.TextBox txtdbVer 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   855
         Width           =   2400
      End
      Begin VB.TextBox txtUsername 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   2400
      End
      Begin VB.TextBox txtTimeOut 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   540
         Width           =   2400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Compiler"
         Height          =   195
         Index           =   23
         Left            =   90
         TabIndex        =   60
         Top             =   1530
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Platform"
         Height          =   195
         Index           =   22
         Left            =   90
         TabIndex        =   59
         Top             =   1215
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DBMS"
         Height          =   195
         Index           =   17
         Left            =   90
         TabIndex        =   58
         Top             =   900
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         Height          =   195
         Index           =   15
         Left            =   90
         TabIndex        =   57
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Timeout"
         Height          =   195
         Index           =   14
         Left            =   90
         TabIndex        =   56
         Top             =   585
         Width           =   570
      End
   End
   Begin VB.Frame fraTable 
      Caption         =   "Table Details"
      Height          =   3975
      Left            =   4500
      TabIndex        =   50
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtPrimaryKey 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2745
         Width           =   2400
      End
      Begin VB.TextBox txtIndexes 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1485
         Width           =   2400
      End
      Begin VB.TextBox txtRules 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1800
         Width           =   2400
      End
      Begin VB.TextBox txtShared 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2115
         Width           =   2400
      End
      Begin VB.TextBox txtTriggers 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2430
         Width           =   2400
      End
      Begin VB.TextBox txtPermissions 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   855
         Width           =   2400
      End
      Begin VB.TextBox txtRows 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1170
         Width           =   2400
      End
      Begin VB.TextBox txtOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   225
         Width           =   2400
      End
      Begin VB.TextBox txtOwner 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   540
         Width           =   2400
      End
      Begin HighlightBox.HBX txtComments 
         Height          =   825
         Left            =   90
         TabIndex        =   25
         Top             =   3060
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   1455
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
         Caption         =   "Primary Key?"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   65
         Top             =   2790
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Shared?"
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   64
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rules?"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   63
         Top             =   1845
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Indexes?"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   62
         Top             =   1530
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Triggers?"
         Height          =   195
         Index           =   16
         Left            =   90
         TabIndex        =   61
         Top             =   2475
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ACL"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   54
         Top             =   900
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   53
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   52
         Top             =   585
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rows"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   51
         Top             =   1215
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmTables"
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
Dim rsTables As New Recordset
Dim rsFields As New Recordset
Dim rsChecks As New Recordset
Dim rsForeign As New Recordset
Dim rsPrimary As New Recordset
Dim rsUnique As New Recordset

Private Sub cmdSerialize_Click()
'-------------------------------------------------------------------------------
' Routine: AddSequenceToField
' Description: Create a sequence and attach it to this field
' Created by: Kirk Roybal ' Date-Time: 3/6/2001 4:42:29 AM
' Machine: KIRK
' set the next sequence id to the highest value contained in the data + 1.
'-------------------------------------------------------------------------------
'on error GoTo Err_Handler
Dim sTableName As String
Dim sFieldName As String
Dim sSQL As String
Dim rsNextVal As New Recordset
Dim lNextVal As Long

  ' exit for any crap reasons
  If Left(trvBrowser.SelectedItem.Key, 1) <> "F" Then
    MsgBox "That object is not a column!", vbExclamation, "Error"
    Exit Sub
  End If
  If Left(trvBrowser.SelectedItem.Parent.Text, 3) = "pg_" Or Left(trvBrowser.SelectedItem.Parent.Text, 8) = "pgadmin_" Then
    MsgBox "That is a system table!", vbExclamation, "Error"
    Exit Sub
  End If

  'DJP - Check the field selected is a suitable type and doesn't already have a default.
  'The last item clicked on the treeview must have been the column,
  'so txtType *should* contain the type for the correct column.
  If Left(txtType.Text, 3) <> "int" And _
     Left(txtType.Text, 5) <> "float" And _
     Left(txtType.Text, 7) <> "numeric" Then
    MsgBox "Only intX, floatX and numeric columns can be serialized!", vbExclamation, "Error"
    Exit Sub
  End If
  If txtDefault.Text <> "" Then
    MsgBox "You cannot serialize a column which has a default value!", vbExclamation, "Error"
    Exit Sub
  End If
  
  'DJP - Confirm action.
  If MsgBox("Are you sure you wish to serialize " & trvBrowser.SelectedItem.Text & "?", vbYesNo + vbQuestion, "Confirm Serialize Column") = vbNo Then Exit Sub
  
  sFieldName = trvBrowser.SelectedItem.Text
  sTableName = trvBrowser.SelectedItem.Parent.Text

  StartMsg "Creating Sequence..."

  sSQL = "SELECT MAX(" & QUOTE & sFieldName & QUOTE & ") FROM " & QUOTE & sTableName & QUOTE
  rsNextVal.Open sSQL, gConnection, adOpenForwardOnly, adLockReadOnly
  If Not (rsNextVal.BOF And rsNextVal.EOF) Then
    lNextVal = Val(rsNextVal(0) & "") + 1
    If lNextVal = 0 Then
      lNextVal = 1
    End If
  End If
  If rsNextVal.State <> adStateClosed Then rsNextVal.Close
  Set rsNextVal = Nothing
  sSQL = "CREATE SEQUENCE " & QUOTE & sTableName & "_" & sFieldName & "_seq" & QUOTE
  LogMsg "Executing: " & sSQL
  gConnection.Execute sSQL, , adCmdText
  LogQuery sSQL
  sSQL = "SELECT setval('" & sTableName & "_" & sFieldName & "_seq', " & lNextVal & ");" & vbCrLf
 
  ' MsgBox sSQL
  LogMsg "Executing: " & sSQL
  gConnection.Execute sSQL, , adCmdText
  LogQuery sSQL
  sSQL = "ALTER TABLE " & QUOTE & sTableName & QUOTE
  sSQL = sSQL & " ALTER COLUMN " & QUOTE & sFieldName & QUOTE
  sSQL = sSQL & " SET DEFAULT nextval('" & sTableName & "_" & sFieldName & "_seq');"
  LogMsg "Executing: " & sSQL
  gConnection.Execute sSQL, , adCmdText
  LogQuery sSQL
  cmdRefresh_Click
  EndMsg
  Exit Sub

Err_Handler:
  If rsNextVal.State <> adStateClosed Then rsNextVal.Close
  Set rsNextVal = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTables, cmdSerialize_Click"
End Sub

Private Sub trvBrowser_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Err_Handler
  If Button = 2 Then PopupMenu fMainForm.mnuCTXTables
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTables, trvBrowser_MouseUp"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Set rsTables = Nothing
  Set rsFields = Nothing
  Set rsChecks = Nothing
  Set rsForeign = Nothing
  Set rsPrimary = Nothing
  Set rsUnique = Nothing
End Sub

Private Sub chkFields_Click()
On Error GoTo Err_Handler
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTables, chkFields_Click"
End Sub

Private Sub chkTables_Click()
On Error GoTo Err_Handler
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTables, chkTables_Click"
End Sub

Public Sub cmdAddColumn_Click()
On Error GoTo Err_Handler
  If Left(trvBrowser.SelectedItem.Key, 1) <> "T" Then
    MsgBox "That object is not a table!", vbExclamation, "Error"
    Exit Sub
  End If
  If Left(trvBrowser.SelectedItem.Text, 3) = "pg_" Or Left(trvBrowser.SelectedItem.Text, 8) = "pgadmin_" Then
    MsgBox "That is a system table!", vbExclamation, "Error"
    Exit Sub
  End If
  Load frmAddColumn
  frmAddColumn.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTables, cmdAddColumn_Click"
End Sub

Public Sub cmdAddTable_Click()
On Error GoTo Err_Handler
  Load frmAddTable
  frmAddTable.Show
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTables, cmdAddTable_Click"
End Sub

Public Sub cmdComment_Click()
On Error GoTo Err_Handler
  If fraTable.Visible Then
    If txtOID.Text = "" Then
      MsgBox "You must select a table to edit the comment for.", vbExclamation, "Error"
      Exit Sub
    End If
    Load frmComments
    frmComments.Setup "frmTables", QUOTE & trvBrowser.SelectedItem.Text & QUOTE, Val(txtOID.Text)
    frmComments.Show
  ElseIf fraColumn.Visible Then
    If txtColOID.Text = "" Then
      MsgBox "You must select a column to edit the comment for.", vbExclamation, "Error"
      Exit Sub
    End If
    Load frmComments
    frmComments.Setup "frmTables - Column", QUOTE & trvBrowser.SelectedItem.Parent.Text & QUOTE & "." & QUOTE & trvBrowser.SelectedItem.Text & QUOTE, Val(txtColOID.Text)
    frmComments.Show
  Else
    MsgBox "You must select a table or column to edit the comment for.", vbExclamation, "Error"
    Exit Sub
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTables, cmdComment_Click"
End Sub

Public Sub cmdDropTable_Click()
On Error GoTo Err_Handler
  If Left(trvBrowser.SelectedItem.Key, 1) <> "T" Then
    MsgBox "That object is not a table!", vbExclamation, "Error"
    Exit Sub
  Else
  If Left(trvBrowser.SelectedItem.Text, 3) = "pg_" Or Left(trvBrowser.SelectedItem.Text, 8) = "pgadmin_" Then
    MsgBox "That is a system table!", vbExclamation, "Error"
    Exit Sub
  End If
    If MsgBox("Are you sure you wish to delete " & trvBrowser.SelectedItem.Text & "?", vbYesNo + vbQuestion, _
              "Confirm Table Delete") = vbYes Then
      StartMsg "Dropping Table..."
      fMainForm.txtSQLPane.Text = "DROP TABLE " & QUOTE & trvBrowser.SelectedItem.Text & QUOTE
      LogMsg "Executing: DROP TABLE " & QUOTE & trvBrowser.SelectedItem.Text & QUOTE
      gConnection.Execute "DROP TABLE " & QUOTE & trvBrowser.SelectedItem.Text & QUOTE
      LogQuery "DROP TABLE " & QUOTE & trvBrowser.SelectedItem.Text & QUOTE
      trvBrowser.Nodes.Remove trvBrowser.SelectedItem.Key
      EndMsg
    End If
  End If
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTables, cmdDropTable_Click"
End Sub

Public Sub cmdRefresh_Click()
'on error GoTo Err_Handler
Dim NodeX As Node
Dim rsDesc As New Recordset
Dim iUbound As Long
Dim iLoop As Long
Dim szGetRows As Variant
Dim lngTable_oid As Long
Dim szTable_name As String
Dim szQuery As String

  fraTable.Visible = False
  fraColumn.Visible = False
  fraDatasource.Visible = False
  fraCheck.Visible = False
  fraForeign.Visible = False
  fraPrimary.Visible = False
  fraUnique.Visible = False
  
  Me.Refresh
  txtUsername.Text = Username
  txtTimeOut.Text = gConnection.ConnectionTimeout
  
  StartMsg "Retrieving Table Definitions..."
  If rsFields.State <> adStateClosed Then rsFields.Close
  If rsChecks.State <> adStateClosed Then rsChecks.Close
  If rsForeign.State <> adStateClosed Then rsForeign.Close
  If rsPrimary.State <> adStateClosed Then rsPrimary.Close
  If rsUnique.State <> adStateClosed Then rsUnique.Close
  If rsTables.State <> adStateClosed Then rsTables.Close
  If chkTables.Value = 1 Then
    szQuery = "SELECT DISTINCT ON(table_name) table_name, table_oid, table_owner, table_acl, table_has_indexes, table_has_rules, table_is_shared, table_has_triggers, table_has_primarykey, table_comments FROM pgadmin_tables ORDER BY table_name"
    LogMsg "Executing: " & szQuery
    rsTables.Open szQuery, gConnection, adOpenStatic
  Else
    szQuery = "SELECT DISTINCT ON(table_name) table_name, table_oid, table_owner, table_acl, table_has_indexes, table_has_rules, table_is_shared, table_has_triggers, table_has_primarykey, table_comments FROM pgadmin_tables WHERE table_oid > " & LAST_SYSTEM_OID & " AND table_name NOT LIKE 'pgadmin_%' AND table_name NOT LIKE 'pg_%' ORDER BY table_name"
    LogMsg "Executing: " & szQuery
    rsTables.Open szQuery, gConnection, adOpenStatic
  End If
  
  trvBrowser.Nodes.Clear
  Set NodeX = trvBrowser.Nodes.Add(, tvwChild, "D:" & Datasource, Datasource, 1)
  If Not (rsTables.EOF) Then
    szGetRows = rsTables.GetRows
    iUbound = UBound(szGetRows, 2)
    For iLoop = 0 To iUbound
        szTable_name = szGetRows(0, iLoop)
        lngTable_oid = szGetRows(1, iLoop)
        Set NodeX = trvBrowser.Nodes.Add("D:" & Datasource, tvwChild, "T:" & lngTable_oid, szTable_name, 2)
    Next iLoop
    rsTables.MoveFirst
    Erase szGetRows
  End If
      
  trvBrowser.Nodes(1).Expanded = True
  trvBrowser.Nodes(1).Selected = True
  LogMsg "Executing: SELECT version()"
  rsDesc.Open "SELECT version()", gConnection, adOpenForwardOnly
  txtdbVer.Text = Mid(rsDesc!Version, 1, InStr(1, rsDesc!Version, " on ") - 1)
  txtPlatform.Text = Mid(rsDesc!Version, InStr(1, rsDesc!Version, " on") + 4, InStr(1, rsDesc!Version, ", compiled by ") - InStr(1, rsDesc!Version, " on") - 4)
  txtCompiler.Text = Mid(rsDesc!Version, InStr(1, rsDesc!Version, ", compiled by ") + 14, Len(rsDesc!Version))
  fraDatasource.Visible = True
  Set rsDesc = Nothing
  EndMsg
  Exit Sub
Err_Handler:
  Set rsDesc = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTables, cmdRefresh_Click"
End Sub

Public Sub cmdRenColumn_Click()
On Error GoTo Err_Handler
Dim NewName As String
Dim AlterStr As String
  If Left(trvBrowser.SelectedItem.Key, 1) <> "F" Then
    MsgBox "That object is not a column!", vbExclamation, "Error"
    Exit Sub
  End If
  If Left(trvBrowser.SelectedItem.Parent.Text, 3) = "pg_" Or Left(trvBrowser.SelectedItem.Parent.Text, 8) = "pgadmin_" Then
    MsgBox "That is a system table!", vbExclamation, "Error"
    Exit Sub
  End If
  NewName = InputBox("Enter a new column name: ", "Rename Column", trvBrowser.SelectedItem.Text)
  If NewName = "" Or NewName = trvBrowser.SelectedItem.Text Then Exit Sub
  AlterStr = "ALTER TABLE " & QUOTE & trvBrowser.SelectedItem.Parent.Text & QUOTE & " RENAME COLUMN " & _
              QUOTE & trvBrowser.SelectedItem.Text & QUOTE & " TO " & QUOTE & NewName & QUOTE
  fMainForm.txtSQLPane.Text = AlterStr
  StartMsg "Renaming Column..."
  LogMsg "Executing: " & AlterStr
  gConnection.Execute AlterStr
  LogQuery AlterStr
  cmdRefresh_Click
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTables, cmdRenColumn_Click"
End Sub

Public Sub cmdRenTable_Click()
On Error GoTo Err_Handler
Dim NewName As String
Dim AlterStr As String
  If Left(trvBrowser.SelectedItem.Key, 1) <> "T" Then
    MsgBox "That object is not a table!", vbExclamation, "Error"
    Exit Sub
  End If
  If Left(trvBrowser.SelectedItem.Text, 3) = "pg_" Or Left(trvBrowser.SelectedItem.Text, 8) = "pgadmin_" Then
    MsgBox "That is a system table!", vbExclamation, "Error"
    Exit Sub
  End If
  NewName = InputBox("Enter a new table name: ", "Rename Table", trvBrowser.SelectedItem.Text)
  If NewName = "" Or NewName = trvBrowser.SelectedItem.Text Then Exit Sub
  AlterStr = "ALTER TABLE " & QUOTE & trvBrowser.SelectedItem.Text & QUOTE & " RENAME TO " & QUOTE & NewName & QUOTE
  StartMsg "Renaming Table..."
  fMainForm.txtSQLPane.Text = AlterStr
  LogMsg "Executing: " & AlterStr
  gConnection.Execute AlterStr
  LogQuery AlterStr
  cmdRefresh_Click
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTables, cmdRenTable_Click"
End Sub

Public Sub cmdData_Click()
On Error GoTo Err_Handler
Dim Response As Integer
Dim Tuples As Long
Dim rsQuery As New Recordset
  If Left(trvBrowser.SelectedItem.Key, 1) <> "T" Then
    MsgBox "That object is not a table!", vbExclamation, "Error"
    Exit Sub
  End If
  If trvBrowser.SelectedItem.Text = "pg_xactlock" Or trvBrowser.SelectedItem.Text = "pg_log" Or trvBrowser.SelectedItem.Text = "pg_variable" Then
    MsgBox "You cannot view pg_log, pg_variable or pg_xactlock - it's a PostgreSQL thing!", vbExclamation, "Error"
    Exit Sub
  End If
  If rsQuery.State <> adStateClosed Then rsQuery.Close
  LogMsg "Executing: SELECT count(*) As records FROM " & QUOTE & trvBrowser.SelectedItem.Text & QUOTE
  rsQuery.Open "SELECT count(*) As records FROM " & QUOTE & trvBrowser.SelectedItem.Text & QUOTE, gConnection, adOpenForwardOnly
  If Not rsQuery.EOF Then
    Tuples = rsQuery!Records
  Else
    Tuples = 0
  End If
  If rsQuery.State <> adStateClosed Then rsQuery.Close
  If Tuples > 1000 Then
    Response = MsgBox("That table contains " & Tuples & " rows which may take some time to load! Do you wish to continue?", _
    vbExclamation + vbYesNo, "Warning")
    If Response = vbNo Then Exit Sub
  End If
  Dim DataForm As New frmSQLOutput
  rsQuery.Open "SELECT * FROM " & QUOTE & trvBrowser.SelectedItem.Text & QUOTE, gConnection, adOpenDynamic, adLockPessimistic
  Load DataForm
  DataForm.Display rsQuery
  DataForm.Show
  DataForm.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTables, cmdData_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 8325
  Me.Height = 4455
  txtCheckDefinition.Wordlist = TextColours
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTables, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  txtComments.Minimise
  txtColComments.Minimise
  txtUniqueColumns.Minimise
  txtPrimaryColumns.Minimise
  txtLocalColumns.Minimise
  txtForeignColumns.Minimise
  txtCheckDefinition.Minimise
  If Me.WindowState <> 1 And Me.ScaleHeight > 0 Then
    If Me.WindowState = 0 Then
      If frmTables.Width < 8325 Then frmTables.Width = 8325
      If frmTables.Height < 4455 Then frmTables.Height = 4455
    End If
    
    trvBrowser.Height = frmTables.ScaleHeight
    trvBrowser.Width = frmTables.ScaleWidth - trvBrowser.Left - fraDatasource.Width - 25
    fraDatasource.Left = trvBrowser.Left + trvBrowser.Width + 25
    fraDatasource.Height = Me.ScaleHeight
    txtComments.Height = fraDatasource.Height - txtComments.Top - 100
    txtColComments.Height = fraDatasource.Height - txtColComments.Top - 100
    txtCheckDefinition.Height = fraDatasource.Height - txtCheckDefinition.Top - 100
    txtForeignColumns.Height = fraDatasource.Height - txtForeignColumns.Top - 100
    txtPrimaryColumns.Height = fraDatasource.Height - txtPrimaryColumns.Top - 100
    txtUniqueColumns.Height = fraDatasource.Height - txtUniqueColumns.Top - 100
    fraTable.Left = fraDatasource.Left
    fraTable.Height = fraDatasource.Height
    fraColumn.Left = fraDatasource.Left
    fraColumn.Height = fraDatasource.Height
    fraCheck.Left = fraDatasource.Left
    fraCheck.Height = fraDatasource.Height
    fraForeign.Left = fraDatasource.Left
    fraForeign.Height = fraDatasource.Height
    fraPrimary.Left = fraDatasource.Left
    fraPrimary.Height = fraDatasource.Height
    fraUnique.Left = fraDatasource.Left
    fraUnique.Height = fraDatasource.Height

  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTables, Form_Resize"
End Sub

Private Sub trvBrowser_NodeClick(ByVal Node As MSComctlLib.Node)
'on error GoTo Err_Handler
Dim NodeX As Node
Dim lOID As Long
Dim rsTemp As New Recordset
Dim x As Long
Dim y As Long
Dim szHex As String
Dim szKey As String
Dim szArgString As String
Dim szArgs() As String

Dim szQuery As String
Dim iLoop As Long
Dim iUbound As Long
Dim szGetRows() As Variant

             
  'If a table was clicked, set the data in the grid, and create children
  'if necessary
    
  Me.Refresh
  Select Case Mid(Node.Key, 1, 1)
  Case "T"
    StartMsg "Retrieving Table Info..."
    fraDatasource.Visible = False
    fraColumn.Visible = False
    fraCheck.Visible = False
    fraForeign.Visible = False
    fraPrimary.Visible = False
    fraUnique.Visible = False
    
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' Tables
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim szTable_name As String
    Dim lTable_OID As Long
    Dim szTable_oid As String
    Dim szTable_owner As String
    Dim szTable_acl As String
    Dim szTable_has_indexes As String
    Dim szTable_has_rules As String
    Dim szTable_is_shared As String
    Dim szTable_has_triggers As String
    Dim szTable_has_primarykey As String
    Dim szTable_comments As String

    If Not (rsTables.EOF) Then
        szGetRows = rsTables.GetRows
        iUbound = UBound(szGetRows, 2)
        For iLoop = 0 To iUbound
            szTable_name = szGetRows(0, iLoop) & ""
            If szTable_name = Node.Text Then
                lTable_OID = Int(szGetRows(1, iLoop))
                szTable_oid = szGetRows(1, iLoop) & ""
                szTable_owner = szGetRows(2, iLoop) & ""
                szTable_acl = szGetRows(3, iLoop) & ""
                szTable_has_indexes = szGetRows(4, iLoop) & ""
                szTable_has_rules = szGetRows(5, iLoop) & ""
                szTable_is_shared = szGetRows(6, iLoop) & ""
                szTable_has_triggers = szGetRows(7, iLoop) & ""
                szTable_has_primarykey = szGetRows(8, iLoop) & ""
                szTable_comments = szGetRows(9, iLoop) & ""
            
                lOID = lTable_OID
                txtOID.Text = szTable_oid
                txtOwner.Text = szTable_owner
                txtPermissions.Text = szTable_acl
                If Node.Text = "pg_log" Or Node.Text = "pg_variable" Or Node.Text = "pg_xactlock" Then
                  txtRows.Text = "Unknown"
                Else
                  Dim result As Variant
                  result = RsExecuteGetResult("SELECT count(*) As records FROM " & QUOTE & Node.Text & QUOTE)
                  If Not (IsNull(result)) Then
                    txtRows.Text = result
                  Else
                    txtRows.Text = "Unknown"
                  End If
                  If rsTemp.State <> adStateClosed Then rsTemp.Close
                End If
                txtIndexes.Text = szTable_has_indexes
                txtRules.Text = szTable_has_rules
                txtShared.Text = szTable_is_shared
                txtTriggers.Text = szTable_has_triggers
                txtPrimaryKey.Text = szTable_has_primarykey
                txtComments.Text = szTable_comments
            End If
        Next iLoop
        Erase szGetRows
        rsTables.MoveFirst
    End If
  
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' Colums
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim szFields_Table_oid As String
    Dim szFields_Table_name As String
    Dim szFields_Column_name As String
    Dim szFields_Column_oid As String
    Dim szFields_column_position As String
    Dim szFields_column_type As String
    Dim szFields_column_length As String
    Dim szFields_column_not_null As String
    Dim szFields_column_default As String
    Dim szFields_column_comments As String
    
    If rsFields.State = adStateClosed Then
      If chkFields.Value = 1 Then
        szQuery = "SELECT table_oid, table_name, column_name, column_oid, column_position, column_type, column_length, column_not_null, column_default, column_comments FROM pgadmin_tables ORDER BY column_position"
        LogMsg "Executing: " & szQuery
        rsFields.Open szQuery, gConnection, adOpenStatic
      Else
        szQuery = "SELECT table_oid, table_name, column_name, column_oid, column_position, column_type, column_length, column_not_null, column_default, column_comments FROM pgadmin_tables WHERE column_position > 0 ORDER BY column_position"
        LogMsg "Executing: " & szQuery
        rsFields.Open szQuery, gConnection, adOpenStatic
      End If
      On Error Resume Next

      If Not (rsFields.EOF) Then
            szGetRows = rsFields.GetRows
            iUbound = UBound(szGetRows, 2)
            For iLoop = 0 To iUbound
                szFields_Table_oid = szGetRows(0, iLoop) & ""
                szFields_Column_oid = szGetRows(3, iLoop) & ""
                szFields_Column_name = szGetRows(2, iLoop) & ""
                Set NodeX = trvBrowser.Nodes.Add("T:" & szFields_Table_oid, tvwChild, "F:" & szFields_Column_oid, szFields_Column_name, 3)
            Next iLoop
            Erase szGetRows
            rsFields.MoveFirst
       End If
      
      On Error GoTo Err_Handler
    End If
    
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' Checks
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim szCheck_table_oid As String
    Dim szCheck_oid As String
    Dim szCheck_name As String
    Dim szCheck_definition As String
    
    If rsChecks.State = adStateClosed Then
      szQuery = "SELECT check_oid, check_name, check_table_oid, check_table_name, check_definition, check_comments FROM pgadmin_checks ORDER BY check_table_name, check_name;"
      LogMsg "Executing: " & szQuery
      rsChecks.Open szQuery, gConnection, adOpenStatic
      On Error Resume Next
      If Not (rsChecks.EOF) Then
            szGetRows = rsChecks.GetRows
            iUbound = UBound(szGetRows, 2)
            For iLoop = 0 To iUbound
                szCheck_oid = szGetRows(0, iLoop) & ""
                szCheck_name = szGetRows(1, iLoop) & ""
                szCheck_table_oid = szGetRows(2, iLoop) & ""
                
                If szCheck_name & "" = "" Then
                  Set NodeX = trvBrowser.Nodes.Add("T:" & szCheck_table_oid, tvwChild, "C:" & szCheck_oid, "Unamed Check", 4)
                Else
                  Set NodeX = trvBrowser.Nodes.Add("T:" & szCheck_table_oid, tvwChild, "C:" & szCheck_oid, szCheck_name, 4)
                End If
            Next iLoop
            Erase szGetRows
            rsChecks.MoveFirst
       End If
      On Error GoTo Err_Handler
    End If
    
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' Foreign keys
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim szForeign_tgrelid As String
    Dim szForeign_oid As String
    Dim szForeign_tgconstrname As String
    
    If rsForeign.State = adStateClosed Then
      szQuery = "SELECT oid, tgrelid, tgconstrname, tgnargs, tgargs, pgadmin_get_desc(oid) AS comments FROM pg_trigger WHERE tgisconstraint = TRUE AND tgtype = 21"
      LogMsg szQuery
      rsForeign.Open szQuery, gConnection, adOpenStatic
      On Error Resume Next
      
      If Not (rsForeign.EOF) Then
            szGetRows = rsForeign.GetRows
            iUbound = UBound(szGetRows, 2)
            For iLoop = 0 To iUbound
                szForeign_tgrelid = szGetRows(1, iLoop) & ""
                szForeign_oid = szGetRows(0, iLoop) & ""
                szForeign_tgconstrname = szGetRows(2, iLoop) & ""
                
                If rsForeign!tgconstrname & "" = "" Then
                  Set NodeX = trvBrowser.Nodes.Add("T:" & szForeign_tgrelid, tvwChild, "O:" & szForeign_oid, "Unamed Foreign Key", 5)
                Else
                  Set NodeX = trvBrowser.Nodes.Add("T:" & szForeign_tgrelid, tvwChild, "O:" & szForeign_oid, szForeign_tgconstrname, 5)
                End If
            Next iLoop
            Erase szGetRows
            rsForeign.MoveFirst
       End If
      On Error GoTo Err_Handler
    End If
    
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' Primary keys
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim szPrimary_index_table As String
    Dim szPrimary_index_oid As String
    Dim szPrimary_index_name As String
    Dim szPrimary_column_name As String
    
    If rsPrimary.State = adStateClosed Then
      szQuery = "SELECT index_oid, index_name, index_table, column_name, index_comments FROM pgadmin_indexes WHERE index_is_primary = 'Yes'"
      LogMsg szQuery
      rsPrimary.Open szQuery, gConnection, adOpenStatic
      On Error Resume Next
     
      If Not (rsPrimary.EOF) Then
            szGetRows = rsPrimary.GetRows
            iUbound = UBound(szGetRows, 2)
            
            For iLoop = 0 To iUbound
                szPrimary_index_table = szGetRows(2, iLoop) & ""
                szPrimary_index_oid = szGetRows(0, iLoop) & ""
                szPrimary_index_name = szGetRows(1, iLoop) & ""
                
                'pgadmin_indexes only has the table name so we need to get the Node Key first
                szKey = ""
                For x = 1 To trvBrowser.Nodes.Count
                  If (trvBrowser.Nodes(x).Text = szPrimary_index_table & "") And (Mid(trvBrowser.Nodes(x).Key, 1, 1) = "T") Then
                    szKey = trvBrowser.Nodes(x).Key
                    Exit For
                  End If
                Next x
                If szKey <> "" Then Set NodeX = trvBrowser.Nodes.Add(szKey, tvwChild, "P:" & szPrimary_index_oid, szPrimary_index_name, 6)
            Next iLoop
            Erase szGetRows
            rsPrimary.MoveFirst
       End If
      On Error GoTo Err_Handler
    End If
  
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' Unique Constraints
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim szUnique_index_table As String
    Dim szUnique_index_oid As String
    Dim szUnique_index_name As String
    Dim szUnique_column_name As String
    
    If rsUnique.State = adStateClosed Then
      'Note, as Primary Keys are inherently unique, exclude them here.
      szQuery = "SELECT index_oid, index_name, index_table, column_name, index_comments FROM pgadmin_indexes WHERE index_is_unique = 'Yes' AND index_is_primary = 'No'"
      LogMsg szQuery
      rsUnique.Open szQuery, gConnection, adOpenStatic
      On Error Resume Next
      
        If Not (rsUnique.EOF) Then
            szGetRows = rsUnique.GetRows
            iUbound = UBound(szGetRows, 2)
            For iLoop = 0 To iUbound
                szUnique_index_table = szGetRows(2, iLoop) & ""
                szUnique_index_oid = szGetRows(0, iLoop) & ""
                szUnique_index_name = szGetRows(1, iLoop) & ""
                
                'pgadmin_indexes only has the table name so we need to get the Node Key first
                szKey = ""
                For x = 1 To trvBrowser.Nodes.Count
                  If (trvBrowser.Nodes(x).Text = szUnique_index_table & "") And (Mid(trvBrowser.Nodes(x).Key, 1, 1) = "T") Then
                    szKey = trvBrowser.Nodes(x).Key
                    Exit For
                  End If
                Next x
                If szKey <> "" Then Set NodeX = trvBrowser.Nodes.Add(szKey, tvwChild, "U:" & szUnique_index_oid, szUnique_index_name, 7)
            Next iLoop
            Erase szGetRows
            rsUnique.MoveFirst
       End If
      On Error GoTo Err_Handler
    End If
  
    EndMsg
    fraTable.Visible = True
    
  'If a field was clicked then display the field data.
  
  Case "F"
    StartMsg "Retrieving Attribute Definition..."
    fraDatasource.Visible = False
    fraTable.Visible = False
    fraCheck.Visible = False
    fraForeign.Visible = False
    fraPrimary.Visible = False
    fraUnique.Visible = False
        
    If Not (rsFields.EOF) Then
            szGetRows = rsFields.GetRows
            iUbound = UBound(szGetRows, 2)
            For iLoop = 0 To iUbound
                 szUnique_index_table = szGetRows(2, iLoop) & ""
                    
                 szFields_Table_oid = szGetRows(0, iLoop) & ""
                 szFields_Table_name = szGetRows(1, iLoop) & ""
                 szFields_Column_name = szGetRows(2, iLoop) & ""
                 szFields_Column_oid = szGetRows(3, iLoop) & ""
                 szFields_column_position = szGetRows(4, iLoop) & ""
                 szFields_column_type = szGetRows(5, iLoop) & ""
                 szFields_column_length = szGetRows(6, iLoop) & ""
                 szFields_column_not_null = szGetRows(7, iLoop) & ""
                 szFields_column_default = szGetRows(8, iLoop) & ""
                 szFields_column_comments = szGetRows(9, iLoop) & ""
                       
                 If szFields_Column_name = Node.Text And szFields_Table_name = Node.Parent.Text Then
                    txtColOID.Text = szFields_Column_oid
                    txtNumber.Text = szFields_column_position
                    If szFields_column_type & "" = "numeric" Then
                      szHex = Hex((Int(szFields_column_length) - 4) And &HFFFF)
                      txtLength.Text = CLng("&H" & Mid(szHex, 1, Len(szHex) - 4)) & "," & CLng("&H" & Mid(szHex, Len(szHex) - 3, Len(szHex)))
                    Else
                      txtLength.Text = szFields_column_length
                    End If
                    txtNotNull.Text = szFields_column_not_null
                    txtType.Text = szFields_column_type
                    txtDefault.Text = szFields_column_default
                    txtColComments.Text = szFields_column_comments
                End If
                   
            Next iLoop
            Erase szGetRows
            rsFields.MoveFirst
       End If
    fraColumn.Visible = True
    
  Case "C" 'Check

    StartMsg "Retrieving Check Definition..."
    fraDatasource.Visible = False
    fraTable.Visible = False
    fraColumn.Visible = False
    fraForeign.Visible = False
    fraPrimary.Visible = False
    fraUnique.Visible = False
 
    If Not (rsChecks.EOF) Then
         szGetRows = rsChecks.GetRows
         iUbound = UBound(szGetRows, 2)
         For iLoop = 0 To iUbound
                szCheck_oid = szGetRows(0, iLoop) & ""
                szCheck_definition = szGetRows(4, iLoop) & ""
                
                If CLng(szCheck_oid) = CLng(Mid(Node.Key, 3)) Then
                    txtCheckOID.Text = szCheck_oid
                    txtCheckDefinition.Text = szCheck_definition
                    iLoop = iUbound
                End If
         Next iLoop
         Erase szGetRows
         rsChecks.MoveFirst
    End If
    fraCheck.Visible = True
    
  Case "O" 'Foreign Key
  
    StartMsg "Retrieving Foreign Key Definition..."
    fraDatasource.Visible = False
    fraTable.Visible = False
    fraColumn.Visible = False
    fraCheck.Visible = False
    fraPrimary.Visible = False
    fraUnique.Visible = False
    txtForeignTable.Text = ""
    txtForeignColumns.Text = ""
    txtLocalColumns.Text = ""
    While Not rsForeign.EOF
      If rsForeign!OID = CLng(Mid(Node.Key, 3)) Then
        txtForeignOID.Text = rsForeign!OID & ""
        If rsForeign!tgnargs >= 6 Then
          For x = 0 To rsForeign.Fields("tgargs").ActualSize - 1
            szArgString = szArgString & Chr(rsForeign!tgargs(x))
          Next x
          szArgs = Split(szArgString, Chr(0))
          txtForeignTable.Text = szArgs(2)
          y = 1
          For x = 4 To UBound(szArgs) Step 2
            If szArgs(x) <> "" Then
              txtLocalColumns.Text = txtLocalColumns.Text & y & ") " & szArgs(x) & vbCrLf
              txtForeignColumns.Text = txtForeignColumns.Text & y & ") " & szArgs(x + 1) & vbCrLf
              y = y + 1
            End If
          Next x
        End If
        rsForeign.MoveLast
      End If
      rsForeign.MoveNext
    Wend
    If rsForeign.BOF <> True Then rsForeign.MoveFirst
    fraForeign.Visible = True
    
  Case "P" 'Primary Key

    StartMsg "Retrieving Primary Key Definition..."
    fraDatasource.Visible = False
    fraTable.Visible = False
    fraColumn.Visible = False
    fraCheck.Visible = False
    fraForeign.Visible = False
    fraUnique.Visible = False
    txtPrimaryColumns.Text = ""
    
    If Not (rsPrimary.EOF) Then
         szGetRows = rsPrimary.GetRows
         iUbound = UBound(szGetRows, 2)
         For iLoop = 0 To iUbound
            szPrimary_index_oid = szGetRows(0, iLoop) & ""
            szPrimary_column_name = szGetRows(3, iLoop) & ""
                
            If CLng(szPrimary_index_oid) = CLng(Mid(Node.Key, 3)) Then
              txtPrimaryOID.Text = szPrimary_index_oid
              txtPrimaryColumns.Text = txtPrimaryColumns.Text & szPrimary_column_name & vbCrLf
            End If
         Next iLoop
         Erase szGetRows
         rsPrimary.MoveFirst
    End If
    fraPrimary.Visible = True
    
  Case "U" 'Unique Constraint

    StartMsg "Retrieving Unique Constraint Definitions..."
    fraDatasource.Visible = False
    fraTable.Visible = False
    fraColumn.Visible = False
    fraCheck.Visible = False
    fraForeign.Visible = False
    fraPrimary.Visible = False
    txtUniqueColumns.Text = ""
    
    If Not (rsUnique.EOF) Then
         szGetRows = rsUnique.GetRows
         iUbound = UBound(szGetRows, 2)
         For iLoop = 0 To iUbound
            szUnique_index_oid = szGetRows(0, iLoop) & ""
            szUnique_column_name = szGetRows(3, iLoop) & ""
                
            If CLng(szUnique_index_oid) = CLng(Mid(Node.Key, 3)) Then
              txtUniqueOID.Text = szUnique_index_oid
              txtUniqueColumns.Text = txtUniqueColumns.Text & szUnique_column_name & vbCrLf
            End If
         Next iLoop
         Erase szGetRows
         rsUnique.MoveFirst
    End If
    fraUnique.Visible = True
    
  Case "D" 'Datasource

    fraTable.Visible = False
    fraColumn.Visible = False
    fraCheck.Visible = False
    fraForeign.Visible = False
    fraPrimary.Visible = False
    fraUnique.Visible = False
    txtUsername.Text = Username
    txtTimeOut.Text = gConnection.CommandTimeout
    fraDatasource.Visible = True
    LogMsg "Executing: SELECT version()"
    rsTemp.Open "SELECT version()", gConnection, adOpenForwardOnly
    txtdbVer.Text = Mid(rsTemp!Version, 1, InStr(1, rsTemp!Version, " on ") - 1)
    txtPlatform.Text = Mid(rsTemp!Version, InStr(1, rsTemp!Version, " on") + 4, InStr(1, rsTemp!Version, ", compiled by ") - InStr(1, rsTemp!Version, " on") - 4)
    txtCompiler.Text = Mid(rsTemp!Version, InStr(1, rsTemp!Version, ", compiled by ") + 14, Len(rsTemp!Version))
  
  End Select
  
  'This stuff can always be done.
  
  Set rsTemp = Nothing
  Node.Expanded = True
  EndMsg
  Exit Sub
Err_Handler:
  Set rsTemp = Nothing
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmTables, trvBrowser_NodeClick"
End Sub


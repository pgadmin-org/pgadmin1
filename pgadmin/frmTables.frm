VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
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
      TabIndex        =   91
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
      TabIndex        =   47
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
   Begin VB.Frame fraColumn 
      Caption         =   "Column Details"
      Height          =   3975
      Left            =   4500
      TabIndex        =   48
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
      Begin VB.TextBox txtColComments 
         BackColor       =   &H8000000F&
         Height          =   1500
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   2385
         Width           =   3480
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Default"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   55
         Top             =   1845
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Not Null"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   54
         Top             =   1530
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Index           =   13
         Left            =   90
         TabIndex        =   53
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   195
         Index           =   18
         Left            =   90
         TabIndex        =   52
         Top             =   1215
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Length"
         Height          =   195
         Index           =   19
         Left            =   90
         TabIndex        =   51
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Number"
         Height          =   195
         Index           =   20
         Left            =   90
         TabIndex        =   50
         Top             =   585
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   21
         Left            =   90
         TabIndex        =   49
         Top             =   270
         Width           =   285
      End
   End
   Begin VB.Frame fraDatasource 
      Caption         =   "Datasource Details"
      Height          =   3975
      Left            =   4500
      TabIndex        =   62
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
         TabIndex        =   67
         Top             =   1530
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Platform"
         Height          =   195
         Index           =   22
         Left            =   90
         TabIndex        =   66
         Top             =   1215
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DBMS"
         Height          =   195
         Index           =   17
         Left            =   90
         TabIndex        =   65
         Top             =   900
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         Height          =   195
         Index           =   15
         Left            =   90
         TabIndex        =   64
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Timeout"
         Height          =   195
         Index           =   14
         Left            =   90
         TabIndex        =   63
         Top             =   585
         Width           =   570
      End
   End
   Begin VB.Frame fraTable 
      Caption         =   "Table Details"
      Height          =   3975
      Left            =   4500
      TabIndex        =   56
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
      Begin VB.TextBox txtComments 
         BackColor       =   &H8000000F&
         Height          =   555
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   3330
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Primary Key?"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   72
         Top             =   2790
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Shared?"
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   71
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rules?"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   70
         Top             =   1845
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Indexes?"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   69
         Top             =   1530
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Triggers?"
         Height          =   195
         Index           =   16
         Left            =   90
         TabIndex        =   68
         Top             =   2475
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ACL"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   61
         Top             =   900
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   60
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   59
         Top             =   585
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rows"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   58
         Top             =   1215
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Index           =   12
         Left            =   90
         TabIndex        =   57
         Top             =   3105
         Width           =   735
      End
   End
   Begin VB.Frame fraUnique 
      Caption         =   "Unique Constraint Details"
      Height          =   3975
      Left            =   4500
      TabIndex        =   85
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtUniqueColumns 
         BackColor       =   &H8000000F&
         Height          =   1815
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   810
         Width           =   3480
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
      Begin VB.TextBox txtUniqueComments 
         BackColor       =   &H8000000F&
         Height          =   1005
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   2880
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Columns"
         Height          =   195
         Index           =   37
         Left            =   90
         TabIndex        =   90
         Top             =   585
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Index OID"
         Height          =   195
         Index           =   35
         Left            =   90
         TabIndex        =   89
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Index           =   31
         Left            =   90
         TabIndex        =   88
         Top             =   2655
         Width           =   735
      End
   End
   Begin VB.Frame fraPrimary 
      Caption         =   "Primary Key Details"
      Height          =   3975
      Left            =   4500
      TabIndex        =   83
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtPrimaryComments 
         BackColor       =   &H8000000F&
         Height          =   1005
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   2880
         Width           =   3480
      End
      Begin VB.TextBox txtPrimaryOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   225
         Width           =   2400
      End
      Begin VB.TextBox txtPrimaryColumns 
         BackColor       =   &H8000000F&
         Height          =   1815
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   810
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Index           =   34
         Left            =   90
         TabIndex        =   87
         Top             =   2655
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Index OID"
         Height          =   195
         Index           =   33
         Left            =   90
         TabIndex        =   86
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Columns"
         Height          =   195
         Index           =   26
         Left            =   90
         TabIndex        =   84
         Top             =   585
         Width           =   600
      End
   End
   Begin VB.Frame fraForeign 
      Caption         =   "Foreign Key Details"
      Height          =   3975
      Left            =   4500
      TabIndex        =   77
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtLocalColumns 
         BackColor       =   &H8000000F&
         Height          =   600
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   1125
         Width           =   3480
      End
      Begin VB.TextBox txtForeignColumns 
         BackColor       =   &H8000000F&
         Height          =   600
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Top             =   1980
         Width           =   3480
      End
      Begin VB.TextBox txtForeignTable 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   540
         Width           =   2400
      End
      Begin VB.TextBox txtForeignOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   225
         Width           =   2400
      End
      Begin VB.TextBox txtForeignComments 
         BackColor       =   &H8000000F&
         Height          =   1050
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   2835
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Local Columns"
         Height          =   195
         Index           =   30
         Left            =   90
         TabIndex        =   82
         Top             =   900
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Foreign Columns"
         Height          =   195
         Index           =   29
         Left            =   90
         TabIndex        =   81
         Top             =   1755
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Foreign Table"
         Height          =   195
         Index           =   27
         Left            =   90
         TabIndex        =   80
         Top             =   585
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trigger OID"
         Height          =   195
         Index           =   25
         Left            =   90
         TabIndex        =   79
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   78
         Top             =   2610
         Width           =   735
      End
   End
   Begin VB.Frame fraCheck 
      Caption         =   "Check Details"
      Height          =   3975
      Left            =   4500
      TabIndex        =   73
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtCheckComments 
         BackColor       =   &H8000000F&
         Height          =   1500
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Top             =   2385
         Width           =   3480
      End
      Begin VB.TextBox txtCheckOID 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   225
         Width           =   2400
      End
      Begin VB.TextBox txtCheckDefinition 
         BackColor       =   &H8000000F&
         Height          =   1320
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Top             =   810
         Width           =   3480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   76
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Check OID"
         Height          =   195
         Index           =   28
         Left            =   90
         TabIndex        =   75
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Check Definition"
         Height          =   195
         Index           =   24
         Left            =   90
         TabIndex        =   74
         Top             =   585
         Width           =   1170
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
'On Error GoTo Err_Handler
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

Private Sub trvBrowser_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    OID = txtOID.Text
  ElseIf fraColumn.Visible Then
    If txtColOID.Text = "" Then
      MsgBox "You must select a column to edit the comment for.", vbExclamation, "Error"
      Exit Sub
    End If
    OID = txtColOID.Text
  ElseIf fraCheck.Visible Then
    If txtCheckOID.Text = "" Then
      MsgBox "You must select a Check Constraint to edit the comment for.", vbExclamation, "Error"
      Exit Sub
    End If
    OID = txtCheckOID.Text
  ElseIf fraForeign.Visible Then
    If txtForeignOID.Text = "" Then
      MsgBox "You must select a Foreign Key to edit the comment for.", vbExclamation, "Error"
      Exit Sub
    End If
    OID = txtForeignOID.Text
  ElseIf fraPrimary.Visible Then
    If txtPrimaryOID.Text = "" Then
      MsgBox "You must select a Primary Key to edit the comment for.", vbExclamation, "Error"
      Exit Sub
    End If
    OID = txtPrimaryOID.Text
  ElseIf fraUnique.Visible Then
    If txtUniqueOID.Text = "" Then
      MsgBox "You must select a Unique Constraint to edit the comment for.", vbExclamation, "Error"
      Exit Sub
    End If
    OID = txtUniqueOID.Text
  End If
  CallingForm = "frmTables"
  Load frmComments
  frmComments.Show
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
On Error GoTo Err_Handler
Dim NodeX As Node
Dim rsDesc As New Recordset

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
    LogMsg "Executing: SELECT DISTINCT ON(table_name) table_name, table_oid, table_owner, table_acl, table_has_indexes, table_has_rules, table_is_shared, table_has_triggers, table_has_primarykey, table_comments FROM pgadmin_tables ORDER BY table_name"
    rsTables.Open "SELECT DISTINCT ON(table_name) table_name, table_oid, table_owner, table_acl, table_has_indexes, table_has_rules, table_is_shared, table_has_triggers, table_has_primarykey, table_comments FROM pgadmin_tables ORDER BY table_name", gConnection, adOpenStatic
  Else
    LogMsg "Executing: SELECT DISTINCT ON(table_name) table_name, table_oid, table_owner, table_acl, table_has_indexes, table_has_rules, table_is_shared, table_has_triggers, table_has_primarykey, table_comments FROM pgadmin_tables WHERE table_oid > " & LAST_SYSTEM_OID & " AND table_name NOT LIKE 'pgadmin_%' AND table_name NOT LIKE 'pg_%' ORDER BY table_name"
    rsTables.Open "SELECT DISTINCT ON(table_name) table_name, table_oid, table_owner, table_acl, table_has_indexes, table_has_rules, table_is_shared, table_has_triggers, table_has_primarykey, table_comments FROM pgadmin_tables WHERE table_oid > " & LAST_SYSTEM_OID & " AND table_name NOT LIKE 'pgadmin_%' AND table_name NOT LIKE 'pg_%' ORDER BY table_name", gConnection, adOpenStatic
  End If
  trvBrowser.Nodes.Clear
  Set NodeX = trvBrowser.Nodes.Add(, tvwChild, "D:" & Datasource, Datasource, 1)
  While Not rsTables.EOF
    Set NodeX = trvBrowser.Nodes.Add("D:" & Datasource, tvwChild, "T:" & rsTables!table_oid, rsTables!table_name, 2)
    rsTables.MoveNext
  Wend
  If rsTables.BOF <> True Then rsTables.MoveFirst
      
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
  Me.Width = 8325
  Me.Height = 4455
  LogMsg "Loading Form: " & Me.Name
  cmdRefresh_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmTables, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
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
    txtCheckComments.Height = fraDatasource.Height - txtCheckComments.Top - 100
    txtForeignComments.Height = fraDatasource.Height - txtForeignComments.Top - 100
    txtPrimaryComments.Height = fraDatasource.Height - txtPrimaryComments.Top - 100
    txtUniqueComments.Height = fraDatasource.Height - txtUniqueComments.Top - 100
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
On Error GoTo Err_Handler
Dim NodeX As Node
Dim lOID As Long
Dim rsTemp As New Recordset
Dim X As Integer
Dim Y As Integer
Dim szKey As String
Dim szArgString As String
Dim szArgs() As String

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
    While Not rsTables.EOF
      If rsTables!table_name = Node.Text Then
        lOID = rsTables!table_oid
        txtOID.Text = rsTables!table_oid & ""
        txtOwner.Text = rsTables!table_owner & ""
        txtPermissions.Text = rsTables!table_acl & ""
        If Node.Text = "pg_log" Or Node.Text = "pg_variable" Or Node.Text = "pg_xactlock" Then
          txtRows.Text = "Unknown"
        Else
          If rsTemp.State <> adStateClosed Then rsTemp.Close
          LogMsg "Executing: SELECT count(*) As records FROM " & QUOTE & Node.Text & QUOTE
          rsTemp.Open "SELECT count(*) As records FROM " & QUOTE & Node.Text & QUOTE, gConnection, adOpenForwardOnly
          If Not rsTemp.EOF Then
            txtRows.Text = rsTemp!Records
          Else
            txtRows.Text = "Unknown"
          End If
          If rsTemp.State <> adStateClosed Then rsTemp.Close
        End If
        txtIndexes.Text = rsTables!table_has_indexes & ""
        txtRules.Text = rsTables!table_has_rules & ""
        txtShared.Text = rsTables!table_is_shared & ""
        txtTriggers.Text = rsTables!table_has_triggers & ""
        txtPrimaryKey.Text = rsTables!table_has_primarykey & ""
        txtComments.Text = rsTables!table_comments & ""
        rsTables.MoveLast
      End If
      rsTables.MoveNext
    Wend
    If rsTables.BOF <> True Then rsTables.MoveFirst
  
    'Get Columns
    If rsFields.State = adStateClosed Then
      If chkFields.Value = 1 Then
        LogMsg "Executing: SELECT table_oid, table_name, column_name, column_oid, column_position, column_type, column_length, column_not_null, column_default, column_comments FROM pgadmin_tables ORDER BY column_position"
        rsFields.Open "SELECT table_oid, table_name, column_name, column_oid, column_position, column_type, column_length, column_not_null, column_default, column_comments FROM pgadmin_tables ORDER BY column_position", gConnection, adOpenStatic
      Else
        LogMsg "Executing: SELECT table_oid, table_name, column_name, column_oid, column_position, column_type, column_length, column_not_null, column_default, column_comments FROM pgadmin_tables WHERE column_position > 0 ORDER BY column_position"
        rsFields.Open "SELECT table_oid, table_name, column_name, column_oid, column_position, column_type, column_length, column_not_null, column_default, column_comments FROM pgadmin_tables WHERE column_position > 0 ORDER BY column_position", gConnection, adOpenStatic
      End If
      On Error Resume Next
      While Not rsFields.EOF
        Set NodeX = trvBrowser.Nodes.Add("T:" & rsFields!table_oid, tvwChild, "F:" & rsFields!column_oid, rsFields!column_name, 3)
      rsFields.MoveNext
      Wend
      On Error GoTo Err_Handler
      If rsFields.BOF <> True Then rsFields.MoveFirst
    End If
    
    'Get Checks
    If rsChecks.State = adStateClosed Then
      LogMsg "Executing: SELECT * FROM pgadmin_checks ORDER BY check_table_name, check_name"
      rsChecks.Open "SELECT * FROM pgadmin_checks ORDER BY check_table_name, check_name", gConnection, adOpenStatic
      On Error Resume Next
      While Not rsChecks.EOF
        If rsChecks!check_name & "" = "" Then
          Set NodeX = trvBrowser.Nodes.Add("T:" & rsChecks!check_table_oid, tvwChild, "C:" & rsChecks!check_oid, "Unamed Check", 4)
        Else
          Set NodeX = trvBrowser.Nodes.Add("T:" & rsChecks!check_table_oid, tvwChild, "C:" & rsChecks!check_oid, rsChecks!check_name, 4)
        End If
        rsChecks.MoveNext
      Wend
      On Error GoTo Err_Handler
      If rsChecks.BOF <> True Then rsChecks.MoveFirst
    End If
    
    'Get Foreign Keys
    If rsForeign.State = adStateClosed Then
      LogMsg "Executing: SELECT tgrelid, tgconstrname, tgnargs, tgargs, CASE WHEN oid <= " & LAST_SYSTEM_OID & " THEN pgadmin_get_pgdesc(oid) ELSE pgadmin_get_desc(oid) END AS comments FROM pg_trigger WHERE tgisconstraint = TRUE AND tgtype = 21"
      rsForeign.Open "SELECT oid, tgrelid, tgconstrname, tgnargs, tgargs, CASE WHEN oid <= " & LAST_SYSTEM_OID & " THEN pgadmin_get_pgdesc(oid) ELSE pgadmin_get_desc(oid) END AS comments FROM pg_trigger WHERE tgisconstraint = TRUE AND tgtype = 21", gConnection, adOpenStatic
      On Error Resume Next
      While Not rsForeign.EOF
        If rsForeign!tgconstrname & "" = "" Then
          Set NodeX = trvBrowser.Nodes.Add("T:" & rsForeign!tgrelid, tvwChild, "O:" & rsForeign!OID, "Unamed Foreign Key", 5)
        Else
          Set NodeX = trvBrowser.Nodes.Add("T:" & rsForeign!tgrelid, tvwChild, "O:" & rsForeign!OID, rsForeign!tgconstrname, 5)
        End If
        rsForeign.MoveNext
      Wend
      On Error GoTo Err_Handler
      If rsForeign.BOF <> True Then rsForeign.MoveFirst
    End If
    
    'Get Primary Keys
    If rsPrimary.State = adStateClosed Then
      LogMsg "Executing: SELECT index_oid, index_name, index_table, column_name, index_comments FROM pgadmin_indexes WHERE index_is_primary = 'Yes'"
      rsPrimary.Open "SELECT index_oid, index_name, index_table, column_name, index_comments FROM pgadmin_indexes WHERE index_is_primary = 'Yes'", gConnection, adOpenStatic
      On Error Resume Next
      While Not rsPrimary.EOF
        'pgadmin_indexes only has the table name so we need to get the Node Key first
        szKey = ""
        For X = 1 To trvBrowser.Nodes.Count
          If (trvBrowser.Nodes(X).Text = rsPrimary!index_table & "") And (Mid(trvBrowser.Nodes(X).Key, 1, 1) = "T") Then
            szKey = trvBrowser.Nodes(X).Key
            Exit For
          End If
        Next X
        If szKey <> "" Then Set NodeX = trvBrowser.Nodes.Add(szKey, tvwChild, "P:" & rsPrimary!index_oid, rsPrimary!index_name, 6)
        rsPrimary.MoveNext
      Wend
      On Error GoTo Err_Handler
      If rsPrimary.BOF <> True Then rsPrimary.MoveFirst
    End If
  
    'Get Unique Constraints
    If rsUnique.State = adStateClosed Then
      'Note, as Primary Keys are inherently unique, exclude them here.
      LogMsg "Executing: SELECT index_oid, index_name, index_table, column_name, index_comments FROM pgadmin_indexes WHERE index_is_unique = 'Yes' AND index_is_primary = 'No'"
      rsUnique.Open "SELECT index_oid, index_name, index_table, column_name, index_comments FROM pgadmin_indexes WHERE index_is_unique = 'Yes' AND index_is_primary = 'No'", gConnection, adOpenStatic
      On Error Resume Next
      While Not rsUnique.EOF
        'pgadmin_indexes only has the table name so we need to get the Node Key first
        szKey = ""
        For X = 1 To trvBrowser.Nodes.Count
          If (trvBrowser.Nodes(X).Text = rsUnique!index_table & "") And (Mid(trvBrowser.Nodes(X).Key, 1, 1) = "T") Then
            szKey = trvBrowser.Nodes(X).Key
            Exit For
          End If
        Next X
        If szKey <> "" Then Set NodeX = trvBrowser.Nodes.Add(szKey, tvwChild, "U:" & rsUnique!index_oid, rsUnique!index_name, 7)
        rsUnique.MoveNext
      Wend
      On Error GoTo Err_Handler
      If rsUnique.BOF <> True Then rsUnique.MoveFirst
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
    While Not rsFields.EOF
      If rsFields!column_name = Node.Text And rsFields!table_name = Node.Parent.Text Then
        txtColOID.Text = rsFields!column_oid & ""
        txtNumber.Text = rsFields!column_position & ""
        If rsFields!column_type & "" = "numeric" Then
          X = Hex((rsFields!column_length - 4) And &HFFFF)
          txtLength.Text = CLng("&H" & Mid(X, 1, Len(X) - 4)) & "," & CLng("&H" & Mid(X, Len(X) - 3, Len(X)))
        Else
          txtLength.Text = rsFields!column_length & ""
        End If
        txtNotNull.Text = rsFields!column_not_null & ""
        txtType.Text = rsFields!column_type & ""
        txtDefault.Text = rsFields!column_default & ""
        txtColComments.Text = rsFields!column_comments & ""
        rsFields.MoveLast
      End If
      rsFields.MoveNext
    Wend
    If rsFields.BOF <> True Then rsFields.MoveFirst
    fraColumn.Visible = True
    
  Case "C" 'Check

    StartMsg "Retrieving Check Definition..."
    fraDatasource.Visible = False
    fraTable.Visible = False
    fraColumn.Visible = False
    fraForeign.Visible = False
    fraPrimary.Visible = False
    fraUnique.Visible = False
    While Not rsChecks.EOF
      If rsChecks!check_oid = CLng(Mid(Node.Key, 3)) Then
        txtCheckOID.Text = rsChecks!check_oid & ""
        txtCheckDefinition.Text = rsChecks!check_definition & ""
        txtCheckComments.Text = rsChecks!check_comments & ""
        rsChecks.MoveLast
      End If
      rsChecks.MoveNext
    Wend
    If rsChecks.BOF <> True Then rsChecks.MoveFirst
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
          For X = 0 To rsForeign.Fields("tgargs").ActualSize - 1
            szArgString = szArgString & Chr(rsForeign!tgargs(X))
          Next X
          szArgs = Split(szArgString, Chr(0))
          txtForeignTable.Text = szArgs(2)
          Y = 1
          For X = 4 To UBound(szArgs) Step 2
            If szArgs(X) <> "" Then
              txtLocalColumns.Text = txtLocalColumns.Text & Y & ") " & szArgs(X) & vbCrLf
              txtForeignColumns.Text = txtForeignColumns.Text & Y & ") " & szArgs(X + 1) & vbCrLf
              Y = Y + 1
            End If
          Next X
        End If
        txtForeignComments.Text = rsForeign!Comments & ""
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
    While Not rsPrimary.EOF
      If rsPrimary!index_oid = CLng(Mid(Node.Key, 3)) Then
        txtPrimaryOID.Text = rsPrimary!index_oid & ""
        txtPrimaryComments.Text = rsPrimary!index_comments & ""
        txtPrimaryColumns.Text = txtPrimaryColumns.Text & rsPrimary!column_name & vbCrLf
      End If
      rsPrimary.MoveNext
    Wend
    If rsPrimary.BOF <> True Then rsPrimary.MoveFirst
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
    While Not rsUnique.EOF
      If rsUnique!index_oid = CLng(Mid(Node.Key, 3)) Then
        txtUniqueOID.Text = rsUnique!index_oid & ""
        txtUniqueComments.Text = rsUnique!index_comments & ""
        txtUniqueColumns.Text = txtUniqueColumns.Text & rsUnique!column_name & vbCrLf
      End If
      rsUnique.MoveNext
    Wend
    If rsUnique.BOF <> True Then rsUnique.MoveFirst
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


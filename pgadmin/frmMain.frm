VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{001ECB85-1072-11D2-AD1C-C0924EC1BE27}#5.1#0"; "sbarvb.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H80000010&
   Caption         =   "pgAdmin"
   ClientHeight    =   7380
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8775
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSQLPane 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   1425
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   8775
      TabIndex        =   2
      Top             =   5685
      Width           =   8775
      Begin HighlightBox.TBX txtLog 
         Height          =   1275
         Left            =   3780
         TabIndex        =   10
         Top             =   135
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   2249
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
         Caption         =   "Log View"
      End
      Begin VB.CommandButton cmdSQL 
         Caption         =   "&SQL"
         Default         =   -1  'True
         Height          =   330
         Left            =   45
         TabIndex        =   4
         ToolTipText     =   "Copy the contents of the SQL pane to the clipboard."
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "&Copy SQL"
         Height          =   330
         Left            =   45
         TabIndex        =   3
         ToolTipText     =   "Copy the contents of the SQL pane to the clipboard."
         Top             =   540
         Width           =   1095
      End
      Begin HighlightBox.HBX txtSQLPane 
         Height          =   1275
         Left            =   1215
         TabIndex        =   9
         Top             =   135
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   2249
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
         Caption         =   "SQL View"
         Wordlist        =   $"frmMain.frx":030A
         RightMargin     =   1.00000e5
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   8010
         TabIndex        =   8
         Top             =   1125
         Width           =   645
      End
      Begin VB.Label lblSQLBar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         Enabled         =   0   'False
         ForeColor       =   &H80000005&
         Height          =   105
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   8745
      End
      Begin VB.Image imgLogo 
         Height          =   1260
         Left            =   6345
         Picture         =   "frmMain.frx":03E4
         Stretch         =   -1  'True
         Top             =   135
         Width           =   2400
      End
   End
   Begin VB.PictureBox picSideBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   5685
      Left            =   0
      ScaleHeight     =   5685
      ScaleWidth      =   1200
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      Begin SideBarVB.SideBar sbMain 
         Height          =   5820
         Left            =   0
         TabIndex        =   5
         Top             =   135
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   10266
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483636
         ForeColor       =   -2147483639
         SmallImageList  =   "ilSideBar"
         LargeImageList  =   "ilSideBar"
      End
      Begin VB.Label lblToolbar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         Enabled         =   0   'False
         ForeColor       =   &H80000005&
         Height          =   105
         Left            =   0
         TabIndex        =   6
         Top             =   45
         Width           =   1410
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7110
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8758
            Text            =   "Status"
            TextSave        =   "Status"
            Key             =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Key             =   "Timer"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Key             =   "Mode"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Key             =   "Database"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Key             =   "User"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilSideBar 
      Left            =   1350
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18C0
            Key             =   "Tune db"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BDA
            Key             =   "Triggers"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EF4
            Key             =   "Languages"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":220E
            Key             =   "Databases"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2528
            Key             =   "rExec"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2842
            Key             =   "Tracking"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B5C
            Key             =   "Migration"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E76
            Key             =   "Users"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3190
            Key             =   "Groups"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34AA
            Key             =   "Vacuum"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37C4
            Key             =   "Tables"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3ADE
            Key             =   "Indexes"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DF8
            Key             =   "Import"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4112
            Key             =   "Privileges"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":442C
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4746
            Key             =   "SQL"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A60
            Key             =   "Sequences"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D7A
            Key             =   "Functions"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5094
            Key             =   "Views"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53AE
            Key             =   "Psql"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":56C8
            Key             =   "Reports"
            Object.Tag             =   "Reports"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":59E2
            Key             =   "Datasources"
            Object.Tag             =   "Datasources"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5CFC
            Key             =   "Manager"
            Object.Tag             =   "Manager"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6016
            Key             =   "Exporters"
            Object.Tag             =   "Exporters"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuSelectdb 
         Caption         =   "&Select db"
      End
      Begin VB.Menu mnuFileChangePassword 
         Caption         =   "&Change Password"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuFilePrinter 
         Caption         =   "&Printer"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSchema 
      Caption         =   "&Schema"
      Begin VB.Menu mnuSchemaDatabases 
         Caption         =   "&Databases"
      End
      Begin VB.Menu mnuSchemaTables 
         Caption         =   "&Tables"
      End
      Begin VB.Menu mnuSchemaIndexes 
         Caption         =   "&Indexes"
      End
      Begin VB.Menu mnuSchemaViews 
         Caption         =   "&Views"
      End
      Begin VB.Menu mnuSchemaSequences 
         Caption         =   "&Sequences"
      End
      Begin VB.Menu mnuSchemaTriggers 
         Caption         =   "&Triggers"
      End
      Begin VB.Menu mnuSchemaFunctions 
         Caption         =   "&Functions"
      End
      Begin VB.Menu mnuSchemaLanguages 
         Caption         =   "&Languages"
      End
      Begin VB.Menu mnuSchemaPrivileges 
         Caption         =   "&Privileges"
      End
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "S&ystem"
      Begin VB.Menu mnuSystemTunedb 
         Caption         =   "&Tune db"
      End
      Begin VB.Menu mnuSystemVacuum 
         Caption         =   "&Vacuum"
      End
      Begin VB.Menu mnuSystemAnalyze 
         Caption         =   "&Analyze"
      End
      Begin VB.Menu mnuSystemUsers 
         Caption         =   "&Users"
      End
      Begin VB.Menu mnuSystemGroups 
         Caption         =   "&Groups"
      End
      Begin VB.Menu mnuSystemTracking 
         Caption         =   "T&racking"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsSQL 
         Caption         =   "&SQL"
      End
      Begin VB.Menu mnuToolsImport 
         Caption         =   "&Import"
      End
      Begin VB.Menu mnuToolsMigration 
         Caption         =   "&Database Migration"
      End
      Begin VB.Menu mnuToolsReports 
         Caption         =   "&Reports"
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuUtilitiesDatasources 
         Caption         =   "&Datasources"
      End
      Begin VB.Menu mnuUtilitiesrExec 
         Caption         =   "&rExec"
      End
      Begin VB.Menu mnuUtilitiesPsql 
         Caption         =   "&Psql"
      End
      Begin VB.Menu mnuUtilitiesExporters 
         Caption         =   "&Exporters"
      End
   End
   Begin VB.Menu mnuAdvanced 
      Caption         =   "&Advanced"
      Begin VB.Menu mnuAdvancedSequence 
         Caption         =   "Refresh &Sequence Cache"
      End
      Begin VB.Menu mnuAdvancedTable 
         Caption         =   "Refresh &Table Cache"
      End
      Begin VB.Menu mnuAdvancedCreateAll 
         Caption         =   "&Create missing pgAdmin Server Side Objects"
      End
      Begin VB.Menu mnuAdvancedDropAll 
         Caption         =   "&Drop all pgAdmin Server Side Objects"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowViewBB 
         Caption         =   "&View ButtonBar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuWindowViewSQLPane 
         Caption         =   "V&iew SQL Pane"
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile V&ertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuCTXDatabases 
      Caption         =   "Databases"
      Visible         =   0   'False
      Begin VB.Menu mnuCTXDatabase_Create 
         Caption         =   "Create Database"
      End
      Begin VB.Menu mnuCTXDatabase_Drop 
         Caption         =   "Drop Database"
      End
      Begin VB.Menu Sep16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCTXDatabase_UserDSN 
         Caption         =   "Create &User DSN"
      End
      Begin VB.Menu mnuCTXDatabase_SystemDSN 
         Caption         =   "Create &System DSN"
      End
      Begin VB.Menu Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCTXDatabase_Comment 
         Caption         =   "Edit Comment"
      End
      Begin VB.Menu mnuCTXDatabase_Refresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuCTXExporters 
      Caption         =   "Exporters"
      Visible         =   0   'False
      Begin VB.Menu mnuCTXExporters_Refresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu Sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCTXExporters_Install 
         Caption         =   "Install Exporter"
      End
      Begin VB.Menu mnuCTXExporters_Uninstall 
         Caption         =   "Uninstall Exporter"
      End
   End
   Begin VB.Menu mnuCTXReportManager 
      Caption         =   "Report Manager"
      Visible         =   0   'False
      Begin VB.Menu mnuCTXReportManager_View 
         Caption         =   "View"
      End
      Begin VB.Menu mnuCTXReportManager_Add 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuCTXReportManager_Remove 
         Caption         =   "Remove"
      End
   End
   Begin VB.Menu mnuCTXFunctions 
      Caption         =   "Functions"
      Visible         =   0   'False
      Begin VB.Menu mnuCTXFunctions_Add 
         Caption         =   "Add Function"
      End
      Begin VB.Menu mnuCTXFunctions_Modify 
         Caption         =   "Modify Function"
      End
      Begin VB.Menu mnuCTXFunctions_Drop 
         Caption         =   "Drop Function"
      End
      Begin VB.Menu Sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCTXFunctions_Comment 
         Caption         =   "Edit Comment"
      End
      Begin VB.Menu mnuCTXFunctions_Refresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuCTXGroups 
      Caption         =   "Groups"
      Visible         =   0   'False
      Begin VB.Menu mnuCTXGroups_Refresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCTXGroups_Create 
         Caption         =   "Create Group"
      End
      Begin VB.Menu mnuCTXGroups_Drop 
         Caption         =   "Drop Group"
      End
   End
   Begin VB.Menu mnuCTXIndexes 
      Caption         =   "Indexes"
      Visible         =   0   'False
      Begin VB.Menu mnuCTXIndexes_Create 
         Caption         =   "Create Index"
      End
      Begin VB.Menu mnuCTXIndexes_Drop 
         Caption         =   "Drop Index"
      End
      Begin VB.Menu Sep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCTXIndexes_Comment 
         Caption         =   "Edit Comment"
      End
      Begin VB.Menu mnuCTXIndexes_Refresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuCTXLanguages 
      Caption         =   "Languages"
      Visible         =   0   'False
      Begin VB.Menu mnuCTXLanguages_Create 
         Caption         =   "Create Language"
      End
      Begin VB.Menu mnuCTXLanguages_Drop 
         Caption         =   "Drop Language"
      End
      Begin VB.Menu Sep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCTXLanguages_Refresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuCTXSequences 
      Caption         =   "Sequences"
      Visible         =   0   'False
      Begin VB.Menu mnuCTXSequences_Create 
         Caption         =   "Create Sequence"
      End
      Begin VB.Menu mnuCTXSequences_Drop 
         Caption         =   "Drop Sequence"
      End
      Begin VB.Menu Sep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCTXSequences_Comment 
         Caption         =   "Edit Comment"
      End
      Begin VB.Menu mnuCTXSequences_Refresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuCTXTables 
      Caption         =   "Tables"
      Visible         =   0   'False
      Begin VB.Menu mnuCTXTables_Create 
         Caption         =   "Create Table"
      End
      Begin VB.Menu mnuCTXTables_Rename 
         Caption         =   "Rename Table"
      End
      Begin VB.Menu mnuCTXTables_Drop 
         Caption         =   "Drop Table"
      End
      Begin VB.Menu Sep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCTXTables_RenameColumn 
         Caption         =   "Rename Column"
      End
      Begin VB.Menu mnuCTXTables_AddColumn 
         Caption         =   "Add Column"
      End
      Begin VB.Menu Sep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCTXTables_Comment 
         Caption         =   "Edit Comment"
      End
      Begin VB.Menu mnuCTXTables_Data 
         Caption         =   "View Data"
      End
      Begin VB.Menu mnuCTXTables_Refresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuCTXTriggers 
      Caption         =   "Triggers"
      Visible         =   0   'False
      Begin VB.Menu mnuCTXTriggers_Create 
         Caption         =   "Create Trigger"
      End
      Begin VB.Menu mnuCTXTriggers_Modify 
         Caption         =   "Modify Trigger"
      End
      Begin VB.Menu mnuCTXTriggers_Drop 
         Caption         =   "Drop Trigger"
      End
      Begin VB.Menu Sep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCTXTriggers_Comment 
         Caption         =   "Edit Comment"
      End
      Begin VB.Menu mnuCTXTriggers_Refresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuCTXUsers 
      Caption         =   "Users"
      Visible         =   0   'False
      Begin VB.Menu mnuCTXUsers_Refresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu Sep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCTXUsers_Create 
         Caption         =   "Create User"
      End
      Begin VB.Menu mnuCTXUsers_Drop 
         Caption         =   "Drop User"
      End
      Begin VB.Menu mnuCTXUsers_Modify 
         Caption         =   "Modify User"
      End
   End
   Begin VB.Menu mnuCTXViews 
      Caption         =   "Views"
      Visible         =   0   'False
      Begin VB.Menu mnuCTXViews_Create 
         Caption         =   "Create View"
      End
      Begin VB.Menu mnuCTXViews_Modify 
         Caption         =   "Modify View"
      End
      Begin VB.Menu mnuCTXViews_Drop 
         Caption         =   "Drop View"
      End
      Begin VB.Menu Sep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCTXViews_Comment 
         Caption         =   "Edit Comment"
      End
      Begin VB.Menu mnuCTXViews_Data 
         Caption         =   "View Data"
      End
      Begin VB.Menu mnuCTXViews_Refresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Private Sub cmdCopy_Click()
On Error GoTo Err_Handler
  Clipboard.SetText txtSQLPane.Text
  StatusBar1.Panels("Status").Text = "SQL Copied to clipboard."
  StatusBar1.Refresh
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, cmdCopy_Click"
End Sub

Private Sub cmdSQL_Click()
On Error GoTo Err_Handler
  mnuToolsSQL_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, cmdSQL_Click"
End Sub

Private Sub MDIForm_Activate()
On Error GoTo Err_Handler
  txtLog.SelStart = Len(txtLog.Text)
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, MDIForm_Activate"
End Sub

Private Sub MDIForm_DragDrop(Source As Control, x As Single, y As Single)
On Error Resume Next
  If Source.Name = "picSideBar" Then
    If x < Source.Width Then
      Source.Align = vbAlignLeft
      Exit Sub
    End If
    If x > (fMainForm.Width - Source.Width) Then
      Source.Align = vbAlignRight
      Exit Sub
    End If
  ElseIf Source.Name = "picSQLPane" Then
    If y < Source.Height Then
      Source.Align = vbAlignTop
      Exit Sub
    End If
    If y > (fMainForm.Height - (2 * Source.Height)) Then
      Source.Align = vbAlignBottom
      Exit Sub
    End If
  End If
End Sub

Private Sub MDIForm_Load()
On Error GoTo Err_Handler
Dim Prn As String
Dim rs As New Recordset
Dim x As Printer
  LogInitMsg "Loading Form: " & Me.Name
  ActionCancelled = False
  
  'Load the default settings
  
  picSideBar.Align = CInt(RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Button Bar Pos", 3))
  picSQLPane.Align = CInt(RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "SQL Pane Pos", 2))
  txtSQLPane.Wordlist = TextColours

  'Allow the user to login to the correct datasource
  
  Load frmODBCLogon
  frmODBCLogon.Show vbModal, Me
  If DEVELOPMENT Then
    lblVersion.Caption = "Version " & app.Major & "." & app.Minor & "." & app.Revision & " DEV"
  Else
    lblVersion.Caption = "Version " & app.Major & "." & app.Minor & "." & app.Revision
  End If
  With sbMain
    .AddFolder "Schema", "Schema"
    .AddFolder "System", "System"
    .AddFolder "Tools", "Tools"
    .AddFolder "Utilities", "Utilities"
    .Folders("Schema").AddItem "Databases", "Databases", "Databases"
    .Folders("Schema").AddItem "Tables", "Tables", "Tables"
    .Folders("Schema").AddItem "Indexes", "Indexes", "Indexes"
    .Folders("Schema").AddItem "Views", "Views", "Views"
    .Folders("Schema").AddItem "Sequences", "Sequences", "Sequences"
    .Folders("Schema").AddItem "Triggers", "Triggers", "Triggers"
    .Folders("Schema").AddItem "Functions", "Functions", "Functions"
    .Folders("Schema").AddItem "Languages", "Languages", "Languages"
    .Folders("Schema").AddItem "Privileges", "Privileges", "Privileges"
    .Folders("System").AddItem "Vacuum", "Vacuum", "Vacuum"
    .Folders("System").AddItem "Analyze", "Analyze", "Vacuum"
    .Folders("System").AddItem "Tune db", "Tune db", "Tune db"
    .Folders("System").AddItem "Users", "Users", "Users"
    .Folders("System").AddItem "Groups", "Groups", "Groups"
    .Folders("System").AddItem "Tracking", "Tracking", "Tracking"
    .Folders("Tools").AddItem "SQL", "SQL", "SQL"
    .Folders("Tools").AddItem "Import", "Import", "Import"
    .Folders("Tools").AddItem "Migration", "Migration", "Migration"
    .Folders("Tools").AddItem "Reports", "Reports", "Reports"
    .Folders("Utilities").AddItem "Datasources", "Datasources", "Datasources"
    .Folders("Utilities").AddItem "rExec", "rExec", "rExec"
    .Folders("Utilities").AddItem "Psql", "Psql", "Psql"
    .Folders("Utilities").AddItem "Exporters", "Exporters", "Exporters"
  End With
  If Chk_dbVersion <> 0 Then
    MsgBox "This version of pgAdmin requires PostgreSQL v" & MIN_PGSQL_VERSION & " or higher!", vbCritical, "Error"
    Force_Selectdb
  End If
  If ActionCancelled <> True Then Chk_DriverOptions
  If ActionCancelled <> True Then Chk_HelperObjects
  If ActionCancelled <> True Then
    If rs.State <> adStateClosed Then rs.Close
    LogMsg "Executing: SELECT param_value FROM pgadmin_param WHERE param_id = 2"
    rs.Open "SELECT param_value FROM pgadmin_param WHERE param_id = 2", gConnection, adOpenForwardOnly
    If rs!param_value = "Y" Then
      Tracking = True
      If rs.State <> adStateClosed Then rs.Close
      LogMsg "Executing: SELECT param_value FROM pgadmin_param WHERE param_id = 3"
      rs.Open "SELECT param_value FROM pgadmin_param WHERE param_id = 3", gConnection, adOpenForwardOnly
      TrackVer = Val(rs!param_value)
    Else
      Tracking = False
    End If
    If rs.State <> adStateClosed Then rs.Close
  
    'What mode are we running in (Development or Production)?
    LogMsg "Executing: SELECT param_value FROM pgadmin_param WHERE param_id = 4"
    rs.Open "SELECT param_value FROM pgadmin_param WHERE param_id = 4", gConnection, adOpenForwardOnly
    If rs!param_value = "Y" Then
      DevMode = True
      StatusBar1.Panels("Mode").Text = "Development Mode"
    Else
      DevMode = False
      StatusBar1.Panels("Mode").Text = "Production Mode"
    End If
    gDevPostgresqlTables = "pgadmin_dev"
  End If
  
  If rs.State <> adStateClosed Then rs.Close
  
  Prn = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Printer", "")
  For Each x In Printers
    If x.DeviceName = Prn Then
      Set Printer = x
      Exit For
    End If
  Next
  Set rs = Nothing
  Screen.MousePointer = vbNormal
  EndMsg
  Exit Sub
Err_Handler:
  Set rs = Nothing
  Screen.MousePointer = vbNormal
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmMain, MDIForm_Load"
End Sub

Private Sub MDIForm_Resize()
On Error GoTo Err_Handler
  If fMainForm.WindowState <> 1 Then
    If Me.Width < 4800 Then Me.Width = 4800
    If Me.Height < 3500 Then Me.Height = 3500
    If BBar = 1 Then
        mnuWindowViewBB.Checked = True
        picSideBar.Visible = True
    Else
        mnuWindowViewBB.Checked = False
        picSideBar.Visible = False
    End If
    If SQLPane = 1 Then
        mnuWindowViewSQLPane.Checked = True
        picSQLPane.Visible = True
    Else
        mnuWindowViewSQLPane.Checked = False
        picSQLPane.Visible = False
    End If
    Me.Arrange vbArrangeIcons
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, MDIForm_Resize"
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error GoTo Err_Handler
  If mnuWindowViewBB.Checked = True Then
    RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "Button Bar", ValString, "1"
  Else
    RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "Button Bar", ValString, "0"
  End If
  If mnuWindowViewSQLPane.Checked = True Then
    RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "SQL Pane", ValString, "1"
  Else
    RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "SQL Pane", ValString, "0"
  End If
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "Button Bar Pos", ValDWord, picSideBar.Align
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "SQL Pane Pos", ValDWord, picSQLPane.Align
  DoEvents
  If gConnection.State <> adStateClosed Then gConnection.Close
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, MDIForm_Unload"
End Sub


Private Sub mnuAdvancedCreateAll_Click()
On Error GoTo Err_Handler
  If Not SuperUser Then
    MsgBox "Only Superusers can create pgAdmin Server Side Objects!", vbExclamation, "Error"
    Exit Sub
  End If
  If MsgBox("You are about to create all missing pgAdmin Server Side Objects. Are you sure you want to continue?", vbQuestion + vbYesNo, "Delete System Objects?") = vbNo Then Exit Sub
  StartMsg "Creating missing pgAdmin Server Side Objects..."
  Chk_HelperObjects
  EndMsg
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmMain, mnuAdvancedCreateAll_Click"
End Sub

Private Sub mnuAdvancedDropAll_Click()
On Error GoTo Err_Handler
Dim x As Integer
  If Not SuperUser Then
    MsgBox "Only Superusers can drop pgAdmin Server Side Objects!", vbExclamation, "Error"
    Exit Sub
  End If
  x = MsgBox("Do you want to drop the Revision Tracking Log, System Table and Description Table as well (these may contain data which will be lost).", vbQuestion + vbDefaultButton2 + vbYesNoCancel, "Delete System Objects?")
  If x = vbCancel Then Exit Sub
  If x = vbYes Then
    If MsgBox("This action cannot be undone. Are you sure you want to continue?", vbQuestion + vbYesNo, "Confirm Delete") = vbNo Then Exit Sub
    Drop_Objects True
  ElseIf x = vbNo Then
    Drop_Objects False
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuAdvancedDropAll_Click"
End Sub

Private Sub mnuAdvancedSequence_Click()
On Error GoTo Err_Handler
  If MsgBox("You are about to refresh the Sequence Cache. Are you sure you wish to continue?", vbQuestion + vbYesNo, "Delete System Objects?") = vbNo Then Exit Sub
  Update_SequenceCache
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuAdvancedSequence_Click"
End Sub

Private Sub mnuAdvancedTable_Click()
On Error GoTo Err_Handler
  If MsgBox("You are about to refresh the Table Cache. Are you sure you wish to continue?", vbQuestion + vbYesNo, "Delete System Objects?") = vbNo Then Exit Sub
  Update_TableCache
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuAdvancedTable_Click"
End Sub

Private Sub mnuCTXDatabase_Comment_Click()
  frmDatabases.cmdComment_Click
End Sub

Private Sub mnuCTXDatabase_Create_Click()
  frmDatabases.cmdCreatedb_Click
End Sub

Private Sub mnuCTXDatabase_Drop_Click()
  frmDatabases.cmdDropdb_Click
End Sub

Private Sub mnuCTXDatabase_Refresh_Click()
  frmDatabases.cmdRefresh_Click
End Sub

Private Sub mnuCTXDatabase_SystemDSN_Click()
  frmDatabases.cmdSystemDSN_Click
End Sub

Private Sub mnuCTXDatabase_UserDSN_Click()
  frmDatabases.cmdUserDSN_Click
End Sub

Private Sub mnuCTXExporters_Install_Click()
  frmExporters.cmdInstall_Click
End Sub

Private Sub mnuCTXExporters_Refresh_Click()
  frmExporters.cmdRefresh_Click
End Sub

Private Sub mnuCTXExporters_Uninstall_Click()
  frmExporters.cmdUninstall_Click
End Sub

Private Sub mnuCTXFunctions_Add_Click()
  frmFunctions.cmdCreateFunc_Click
End Sub

Private Sub mnuCTXFunctions_Modify_Click()
  frmFunctions.cmdModifyFunc_Click
End Sub

Private Sub mnuCTXFunctions_Comment_Click()
  frmFunctions.cmdComment_Click
End Sub

Private Sub mnuCTXFunctions_Drop_Click()
  frmFunctions.cmdDropFunc_Click
End Sub

Private Sub mnuCTXFunctions_Refresh_Click()
  frmFunctions.cmdRefresh_Click
End Sub

Private Sub mnuCTXGroups_Create_Click()
  frmGroups.cmdCreate_Click
End Sub

Private Sub mnuCTXGroups_Drop_Click()
  frmGroups.cmdDrop_Click
End Sub

Private Sub mnuCTXGroups_Refresh_Click()
  frmGroups.cmdRefresh_Click
End Sub

Private Sub mnuCTXIndexes_Comment_Click()
  frmIndexes.cmdComment_Click
End Sub

Private Sub mnuCTXIndexes_Create_Click()
  frmIndexes.cmdAddIndex_Click
End Sub

Private Sub mnuCTXIndexes_Drop_Click()
  frmIndexes.cmdDropIndex_Click
End Sub

Private Sub mnuCTXIndexes_Refresh_Click()
  frmIndexes.cmdRefresh_Click
End Sub

Private Sub mnuCTXLanguages_Create_Click()
  frmLanguages.cmdCreateLang_Click
End Sub

Private Sub mnuCTXLanguages_Drop_Click()
  frmLanguages.cmdDropLang_Click
End Sub

Private Sub mnuCTXLanguages_Refresh_Click()
  frmLanguages.cmdRefresh_Click
End Sub

Private Sub mnuCTXReportManager_Add_Click()
  frmReportManager.cmdAdd_Click
End Sub

Private Sub mnuCTXReportManager_Remove_Click()
  frmReportManager.cmdRemove_Click
End Sub

Private Sub mnuCTXReportManager_View_Click()
  frmReportManager.cmdView_Click
End Sub

Private Sub mnuCTXSequences_Comment_Click()
  frmSequences.cmdComment_Click
End Sub

Private Sub mnuCTXSequences_Create_Click()
  frmSequences.cmdCreateSeq_Click
End Sub

Private Sub mnuCTXSequences_Drop_Click()
  frmSequences.cmdDropSeq_Click
End Sub

Private Sub mnuCTXSequences_Refresh_Click()
  frmSequences.cmdRefresh_Click
End Sub

Private Sub mnuCTXTables_AddColumn_Click()
  frmTables.cmdAddColumn_Click
End Sub

Private Sub mnuCTXTables_Comment_Click()
  frmTables.cmdComment_Click
End Sub

Private Sub mnuCTXTables_Create_Click()
  frmTables.cmdAddTable_Click
End Sub

Private Sub mnuCTXTables_Data_Click()
  frmTables.cmdData_Click
End Sub

Private Sub mnuCTXTables_Drop_Click()
  frmTables.cmdDropTable_Click
End Sub

Private Sub mnuCTXTables_Refresh_Click()
  frmTables.cmdRefresh_Click
End Sub

Private Sub mnuCTXTables_Rename_Click()
  frmTables.cmdRenTable_Click
End Sub

Private Sub mnuCTXTables_RenameColumn_Click()
  frmTables.cmdRenColumn_Click
End Sub

Private Sub mnuCTXTriggers_Comment_Click()
  frmTriggers.cmdComment_Click
End Sub

Private Sub mnuCTXTriggers_Create_Click()
  frmTriggers.cmdCreateTrig_Click
End Sub

Private Sub mnuCTXTriggers_Modify_Click()
  frmTriggers.cmdModifyTrig_Click
End Sub

Private Sub mnuCTXTriggers_Drop_Click()
  frmTriggers.cmdDropTrig_Click
End Sub

Private Sub mnuCTXTriggers_Refresh_Click()
  frmTriggers.cmdRefresh_Click
End Sub

Private Sub mnuCTXUsers_Create_Click()
  frmUsers.cmdCreate_Click
End Sub

Private Sub mnuCTXUsers_Drop_Click()
  frmUsers.cmdDrop_Click
End Sub

Private Sub mnuCTXUsers_Modify_Click()
  frmUsers.cmdModify_Click
End Sub

Private Sub mnuCTXUsers_Refresh_Click()
  frmUsers.cmdRefresh_Click
End Sub

Private Sub mnuCTXViews_Comment_Click()
  frmViews.cmdComment_Click
End Sub

Private Sub mnuCTXViews_Create_Click()
  frmViews.cmdCreateView_Click
End Sub

Private Sub mnuCTXViews_Modify_Click()
  frmViews.cmdModifyView_Click
End Sub

Private Sub mnuCTXViews_Data_Click()
  frmViews.cmdViewData_Click
End Sub

Private Sub mnuCTXViews_Drop_Click()
  frmViews.cmdDropView_Click
End Sub

Private Sub mnuCTXViews_Refresh_Click()
  frmViews.cmdRefresh_Click
End Sub

Private Sub mnuFileChangePassword_Click()
On Error GoTo Err_Handler
  Load frmPassword
  frmPassword.Show
  frmPassword.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuFileChangePassword_Click"
End Sub

Private Sub mnuFileOptions_Click()
On Error GoTo Err_Handler
  Load frmOptions
  frmOptions.Show
  frmOptions.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuFileOptions_Click"
End Sub

Private Sub mnuFilePrinter_Click()
On Error GoTo Err_Handler
  Load frmPrinter
  frmPrinter.Show
  frmPrinter.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuFilePrinter_Click"
End Sub

Private Sub mnuSchemaDatabases_Click()
On Error GoTo Err_Handler
  Load frmDatabases
  frmDatabases.Show
  frmDatabases.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuSchemaDatabases_Click"
End Sub

Private Sub mnuSchemaFunctions_Click()
On Error GoTo Err_Handler
  Load frmFunctions
  frmFunctions.Show
  frmFunctions.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuSchemaFunctions_Click"
End Sub

Private Sub mnuSchemaLanguages_Click()
On Error GoTo Err_Handler
  Load frmLanguages
  frmLanguages.Show
  frmLanguages.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuSchemaLanguages_Click"
End Sub

Private Sub mnuSchemaSequences_Click()
On Error GoTo Err_Handler
  Load frmSequences
  frmSequences.Show
  frmSequences.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuSchemaSequences_Click"
End Sub

Private Sub mnuSchemaTriggers_Click()
On Error GoTo Err_Handler
  Load frmTriggers
  frmTriggers.Show
  frmTriggers.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuSchemaTriggers_Click"
End Sub

Private Sub mnuSchemaViews_Click()
On Error GoTo Err_Handler
  Load frmViews
  frmViews.Show
  frmViews.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuSchemaViews_Click"
End Sub

Private Sub mnuSystemAnalyze_Click()
On Error GoTo Err_Handler
Dim Response As Integer

  Response = MsgBox("WARNING: Database vacuuming should only be performed when there is no one using the database." & vbCrLf & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo)
  If Response = vbNo Then Exit Sub
  
  StartMsg "Vacuuming & Analyzing the database..."
  LogMsg "Executing: VACUUM ANALYZE"
  gConnection.Execute "VACUUM ANALYZE"
  MsgBox "The database has been vacuumed and analyzed.", vbInformation
  EndMsg
  Exit Sub
  
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmMain, mnuSystemAnalyze_Click"
End Sub

Private Sub mnuSystemGroups_Click()
On Error GoTo Err_Handler
  Load frmGroups
  frmGroups.Show
  frmGroups.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuSystemGroups_Click"
End Sub

Private Sub mnuSystemTracking_Click()
On Error GoTo Err_Handler
  Load frmTracking
  frmTracking.Show
  frmTracking.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuSystemTracking_Click"
End Sub

Private Sub mnuSelectdb_Click()
On Error GoTo Err_Handler
Dim x As Integer
Dim Response As Integer

  Response = MsgBox("Selecting a new database will close all currently open windows." & vbCrLf & "Do you wish to continue?", vbExclamation + vbYesNo, "Select db")
  If Response = vbNo Then Exit Sub
                    
  For x = Forms.Count - 1 To 0 Step -1
    If Forms(x).Name <> "frmMain" Then
      Unload Forms(x)
    End If
  Next
  
  ActionCancelled = False

  'Allow the user to login to the correct datasource
  
  Dim formODBCLogon As New frmODBCLogon
  Load formODBCLogon
  formODBCLogon.Show vbModal, Me
  
  'Check the database version
  
  Chk_DriverOptions
  If Chk_dbVersion <> 0 Then
    MsgBox "This version of pgAdmin requires PostgreSQL v" & MIN_PGSQL_VERSION & " or higher!", vbCritical, "Error"
    Force_Selectdb
  End If
  Chk_HelperObjects
  
  StatusBar1.Panels("Status").Text = "Ready"
  StatusBar1.Panels("Database").Text = "Connected to: " & Datasource
  StatusBar1.Panels("User").Text = "Username: " & Username
  StatusBar1.Refresh
  Screen.MousePointer = vbNormal
  Exit Sub
Err_Handler:
  Screen.MousePointer = vbNormal
  If Err.Number <> 0 Then LogError Err, "frmMain, mnuSelectdb_Click"
End Sub

Private Sub Force_Selectdb()
On Error GoTo Err_Handler
Dim x As Integer
             
  For x = Forms.Count - 1 To 0 Step -1
    If Forms(x).Name <> "frmMain" Then
      Unload Forms(x)
    End If
  Next
  
  ActionCancelled = False

  'Allow the user to login to the correct datasource
  
  Dim formODBCLogon As New frmODBCLogon
  Load formODBCLogon
  formODBCLogon.Show vbModal, Me
  
  'Check the database version
  
  Chk_DriverOptions
  If Chk_dbVersion <> 0 Then
    MsgBox "This version of pgAdmin requires PostgreSQL v" & MIN_PGSQL_VERSION & " or higher!", vbCritical, "Error"
    Force_Selectdb
  End If
  
  StatusBar1.Panels("Status").Text = "Ready"
  StatusBar1.Panels("Database").Text = "Connected to: " & Datasource
  StatusBar1.Panels("User").Text = "Username: " & Username
  StatusBar1.Refresh
  Screen.MousePointer = vbNormal
  Exit Sub
Err_Handler:
  Screen.MousePointer = vbNormal
  If Err.Number <> 0 Then LogError Err, "frmMain, Force_Selectdb"
End Sub
Private Sub mnuSystemUsers_Click()
On Error GoTo Err_Handler
  Load frmUsers
  frmUsers.Show
  frmUsers.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuSystemUsers_Click"
End Sub

Private Sub mnuToolsMigration_Click()
On Error GoTo Err_Handler
  Load frmMigration
  frmMigration.Show
  frmMigration.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuToolsMigration_Click"
End Sub

Private Sub mnuToolsImport_Click()
On Error GoTo Err_Handler
  Load frmImport
  frmImport.Show
  frmImport.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuToolsImport_Click"
End Sub

Private Sub mnuToolsReports_Click()
On Error GoTo Err_Handler
  Load frmReportManager
  frmReportManager.Show
  frmReportManager.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnutoolsReports_Click"
End Sub

Private Sub mnuUtilitiesDatasources_Click()
On Error GoTo Err_Handler
Dim Scr_hDC As Long
Dim x As Long
  Scr_hDC = GetDesktopWindow()
  LogMsg "Executing: Opening the ODBC Datasource Manager..."
  x = ShellExecute(Scr_hDC, "Open", "rundll32.exe", "shell32.dll,Control_RunDLL odbccp32.cpl", "C:\", SW_SHOWNORMAL)
  If x <= 32 Then
    MsgBox "An error occured opening the 32Bit ODBC Datasource Manager.", vbCritical, "Error!"
    LogMsg "Could not open the ODBC Datasource Manager (Error: " & x & ")."
    End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuUtilitiesDatasources_Click"
End Sub

Private Sub mnuUtilitiesExporters_Click()
On Error GoTo Err_Handler
  Load frmExporters
  frmExporters.Show
  frmExporters.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuUtilitiesExporters_Click"
End Sub

Private Sub mnuUtilitiesPsql_Click()
On Error GoTo Err_Handler
  Load frmPsql
  frmPsql.Show
  frmPsql.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuUtilitiesPsql_Click"
End Sub

Private Sub mnuUtilitiesrExec_Click()
On Error GoTo Err_Handler
  Load frmRexec
  frmRexec.Show
  frmRexec.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuUtilitiesrExec_Click"
End Sub

Private Sub mnuToolsSQL_Click()
On Error GoTo Err_Handler
Dim SQL As New frmSQL
Dim x As Integer
Dim y As Integer
  y = 1
  For x = 0 To Forms.Count - 1
    If Mid(Forms(x).Caption, 1, 6) = "SQL - " Then
      If InStr(1, Forms(x).Caption, "(") <> 0 Then
        If Mid(Forms(x).Caption, 7, InStr(7, Forms(x).Caption, " ") - 7) = y Then y = Mid(Forms(x).Caption, 7, InStr(7, Forms(x).Caption, " ") - 7) + 1
      Else
        If Mid(Forms(x).Caption, 7) = y Then y = Mid(Forms(x).Caption, 7) + 1
      End If
    End If
  Next
  Load SQL
  SQL.Show
  SQL.Caption = "SQL - " & y
  SQL.szTitle = "SQL - " & y
  SQL.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuToolsSQL_Click"
End Sub

Private Sub mnuSystemTunedb_Click()
On Error GoTo Err_Handler
  Load frmTunedb
  If ActionCancelled = False Then
    frmTunedb.Show
    frmTunedb.ZOrder 0
  Else
    ActionCancelled = False
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuSystemTunedb_Click"
End Sub

Private Sub mnuSystemVacuum_Click()
On Error GoTo Err_Handler
Dim Response As Integer

  Response = MsgBox("WARNING: Database vacuuming should only be performed when there is no one using the database." & vbCrLf & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo)
  If Response = vbNo Then Exit Sub
  
  StartMsg "Vacuuming database..."
  LogMsg "Executing: VACUUM"
  gConnection.Execute "VACUUM"
  MsgBox "The database has been vacuumed.", vbInformation
  EndMsg
  Exit Sub
  
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmMain, mnuSystemVacuum_Click"
End Sub

Private Sub mnuHelpAbout_Click()
On Error GoTo Err_Handler
  frmAbout.Show vbModal, Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuHelpAbout_Click"
End Sub

Private Sub mnuSchemaIndexes_Click()
On Error GoTo Err_Handler
  Load frmIndexes
  frmIndexes.Show
  frmIndexes.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuSchemaIndexes_Click"
End Sub

Private Sub mnuSchemaPrivileges_Click()
On Error GoTo Err_Handler
  Load frmPrivileges
  frmPrivileges.Show
  frmPrivileges.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuSchemaPrivilieges_Click"
End Sub

Private Sub mnuSchemaTables_Click()
On Error GoTo Err_Handler
  Load frmTables
  frmTables.Show
  frmTables.ZOrder 0
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuSchemaTables_Click"
End Sub

Private Sub mnuWindowArrangeIcons_Click()
On Error GoTo Err_Handler
  Me.Arrange vbArrangeIcons
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuWindowArrangeIcons_Click"
End Sub

Private Sub mnuWindowCascade_Click()
On Error GoTo Err_Handler
  Me.Arrange vbCascade
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuWindowCascade_Click"
End Sub

Private Sub mnuWindowTileHorizontal_Click()
On Error GoTo Err_Handler
  Me.Arrange vbTileHorizontal
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuWindowTileHorizontal_Click"
End Sub

Private Sub mnuWindowTileVertical_Click()
On Error GoTo Err_Handler
  Me.Arrange vbTileVertical
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuWindowTileVertical_Click"
End Sub

Private Sub mnuFileExit_Click()
Dim szMessage As String
Dim szAnswer As Variant

If cmp_Project_IsRebuilt = False Then
   szMessage = "Your project needs to be rebuilt to apply changes in " & vbCrLf & _
   "functions, triggers and views. Apply changes ?"
   szAnswer = MsgBox(szMessage, vbYesNo, "Rebuild project")
   If szAnswer = vbYes Then
        cmp_Project_Rebuild
        If bContinueRebuilding = False Then Exit Sub
   End If
End If

On Error GoTo Err_Handler
  Unload Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuFileExit_Click"
  End
End Sub

Private Sub mnuWindowViewSQLPane_Click()
On Error GoTo Err_Handler
  If picSQLPane.Visible = True Then
    picSQLPane.Visible = False
    mnuWindowViewSQLPane.Checked = False
    SQLPane = 0
  Else
    picSQLPane.Visible = True
    mnuWindowViewSQLPane.Checked = True
    SQLPane = 1
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuWindowViewSQLPane_Click"
End Sub

Private Sub picSQLPane_Resize()
On Error GoTo Err_Handler
  If fMainForm.WindowState <> 1 Then
    If txtSQLPane.Maximised Then txtSQLPane.Minimise
    If txtLog.Maximised Then txtLog.Minimise
    lblSQLBar.Top = 15
    lblSQLBar.Left = 15
    lblSQLBar.Width = picSQLPane.Width - 30
    txtSQLPane.Height = picSQLPane.Height - (lblSQLBar.Height + 15)
    txtSQLPane.Width = ((picSQLPane.Width - imgLogo.Width - txtSQLPane.Left) / 5) * 2
    txtLog.Height = picSQLPane.Height - (lblSQLBar.Height + 15)
    txtLog.Left = txtSQLPane.Left + txtSQLPane.Width
    txtLog.Width = (txtSQLPane.Width / 2) * 3
    imgLogo.Left = picSQLPane.Width - imgLogo.Width
    lblVersion.Top = picSQLPane.Height - (lblVersion.Height * 1.5)
    lblVersion.Left = picSQLPane.Width - lblVersion.Width - 50
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, picSQLPane_Resize"
End Sub

Private Sub picSideBar_Resize()
On Error GoTo Err_Handler
  If fMainForm.WindowState <> 1 Then
    lblToolbar.Top = 15
    lblToolbar.Left = 15
    lblToolbar.Width = picSideBar.Width - 30
    sbMain.Top = lblToolbar.Height + 15
    sbMain.Left = 0
    sbMain.Height = picSideBar.Height - (lblToolbar.Height + 15)
    sbMain.Width = picSideBar.Width
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, picSideBar_Resize"
End Sub

Private Sub mnuWindowViewBB_Click()
On Error GoTo Err_Handler
  If picSideBar.Visible = True Then
    picSideBar.Visible = False
    mnuWindowViewBB.Checked = False
    BBar = 0
  Else
    picSideBar.Visible = True
    mnuWindowViewBB.Checked = True
    BBar = 1
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, mnuWindowViewBB_Click"
End Sub

Private Sub sbMain_FolderItemClick(FolderItem As SideBarVB.FolderItem)
On Error GoTo Err_Handler
  Select Case FolderItem.Key
    Case "Databases"
      mnuSchemaDatabases_Click
    Case "Tables"
      mnuSchemaTables_Click
    Case "Indexes"
      mnuSchemaIndexes_Click
    Case "Views"
      mnuSchemaViews_Click
    Case "Sequences"
      mnuSchemaSequences_Click
    Case "Triggers"
      mnuSchemaTriggers_Click
    Case "Functions"
      mnuSchemaFunctions_Click
    Case "Languages"
      mnuSchemaLanguages_Click
    Case "Privileges"
      mnuSchemaPrivileges_Click
    Case "Tracking"
      mnuSystemTracking_Click
    Case "Vacuum"
      mnuSystemVacuum_Click
    Case "Analyze"
      mnuSystemAnalyze_Click
    Case "Tune db"
      mnuSystemTunedb_Click
    Case "Users"
      mnuSystemUsers_Click
    Case "Groups"
      mnuSystemGroups_Click
    Case "SQL"
      mnuToolsSQL_Click
    Case "Import"
      mnuToolsImport_Click
    Case "Migration"
      mnuToolsMigration_Click
    Case "Reports"
      mnuToolsReports_Click
    Case "Datasources"
      mnuUtilitiesDatasources_Click
    Case "rExec"
      mnuUtilitiesrExec_Click
    Case "Psql"
      mnuUtilitiesPsql_Click
    Case "Exporters"
      mnuUtilitiesExporters_Click
  End Select
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, sbMain_FolderItemClick"
End Sub

Private Sub sbMain_FolderItemHiLite(FolderItem As SideBarVB.FolderItem)
On Error GoTo Err_Handler
  Select Case FolderItem.Key
    Case "Databases"
      sbMain.ToolTipText = "Open the Database browser."
    Case "Tables"
      sbMain.ToolTipText = "Open the Table browser."
    Case "Indexes"
      sbMain.ToolTipText = "Open the Index browser."
    Case "Views"
      sbMain.ToolTipText = "Open the View browser."
    Case "Sequences"
      sbMain.ToolTipText = "Open the Sequence browser."
    Case "Triggers"
      sbMain.ToolTipText = "Open the Trigger browser."
    Case "Functions"
      sbMain.ToolTipText = "Open the Function browser."
    Case "Languages"
      sbMain.ToolTipText = "Open the Language browser."
    Case "Privileges"
      sbMain.ToolTipText = "Open the Privileges configuration dialogue."
    Case "Tracking"
      sbMain.ToolTipText = "Open the Revision Tracking dialogue."
    Case "Vacuum"
      sbMain.ToolTipText = "Vacuum the database."
    Case "Analyze"
      sbMain.ToolTipText = "Vacuum and Analyze the database."
    Case "Tune db"
      sbMain.ToolTipText = "Tune the database for MS Access & Visual Basic."
    Case "Users"
      sbMain.ToolTipText = "Open the User Manager."
    Case "Groups"
      sbMain.ToolTipText = "Open the User Group Manager."
    Case "SQL"
      sbMain.ToolTipText = "Enter arbitary SQL queries."
    Case "Import"
      sbMain.ToolTipText = "Run the Data Import Wizard."
    Case "Migration"
      sbMain.ToolTipText = "Open the Database Migration Wizard."
    Case "Reports"
      sbMain.ToolTipText = "Run the report browser."
    Case "Datasources"
      sbMain.ToolTipText = "Run the 32Bit ODBC Datasource Manager."
    Case "rExec"
      sbMain.ToolTipText = "Execute OS commands on a remote host."
    Case "Psql"
      sbMain.ToolTipText = "Execute Psql."
    Case "Exporters"
      sbMain.ToolTipText = "Run the Exporter Manager."
  End Select
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmMain, sbMain_FolderItemHiLite"
End Sub

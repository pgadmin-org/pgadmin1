VERSION 5.00
Begin VB.Form frmODBCLogon 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "pgAdmin Logon"
   ClientHeight    =   1560
   ClientLeft      =   2850
   ClientTop       =   1710
   ClientWidth     =   4110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmODBCLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   300
      Left            =   3375
      TabIndex        =   1
      ToolTipText     =   "Logon with the parameters entered"
      Top             =   90
      Width           =   660
   End
   Begin VB.ComboBox cboDSNList 
      Height          =   315
      ItemData        =   "frmODBCLogon.frx":030A
      Left            =   1035
      List            =   "frmODBCLogon.frx":030C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select a PostgreSQL datasource."
      Top             =   90
      Width           =   2280
   End
   Begin VB.TextBox txtPWD 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1035
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "Enter your password on the remote databae"
      Top             =   780
      Width           =   3015
   End
   Begin VB.TextBox txtUID 
      Height          =   300
      Left            =   1035
      TabIndex        =   2
      ToolTipText     =   "Enter you username on the remote database"
      Top             =   450
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   1845
      TabIndex        =   4
      ToolTipText     =   "Cancel the logon and exit the program"
      Top             =   1170
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   2970
      TabIndex        =   5
      ToolTipText     =   "Logon with the parameters entered"
      Top             =   1170
      Width           =   1065
   End
   Begin VB.Label lblLogon 
      AutoSize        =   -1  'True
      Caption         =   "&Password:"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   8
      Top             =   825
      Width           =   735
   End
   Begin VB.Label lblLogon 
      AutoSize        =   -1  'True
      Caption         =   "&Username:"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   480
      Width           =   765
   End
   Begin VB.Label lblLogon 
      AutoSize        =   -1  'True
      Caption         =   "&Datasource:"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   6
      Top             =   135
      Width           =   870
   End
End
Attribute VB_Name = "frmODBCLogon"
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

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
  ActionCancelled = True
  Unload Me
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmODBCLogon, cmdCancel_Click"
  End
End Sub

Private Sub cmdNew_Click()
On Error GoTo Err_Handler
  Load frmNewDSN
  frmNewDSN.Show vbModal, Me
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmODBCLogon, cmdNew_Click"
End Sub

Public Sub cmdOK_Click()
On Error GoTo Err_Handler
    
  Screen.MousePointer = vbHourglass
  Connect = "DSN=" & cboDSNList.Text & ";UID=" & txtUID.Text & ";PWD=" & txtPWD.Text
  If gConnection.State <> adStateClosed Then gConnection.Close
  gConnection.Open Connect
  Datasource = cboDSNList.Text
  Username = txtUID.Text
  Password = txtPWD.Text
  If DEVELOPMENT Then
    fMainForm.Caption = Datasource & " - pgAdmin v" & app.Major & "." & app.Minor & "." & app.Revision & " DEV"
  Else
    fMainForm.Caption = Datasource & " - pgAdmin v" & app.Major & "." & app.Minor & "." & app.Revision
  End If
  fMainForm.StatusBar1.Panels(1).Text = "Ready"
  fMainForm.StatusBar1.Panels(3).Text = "Connected to: " & Datasource
  fMainForm.StatusBar1.Panels(4).Text = "Username: " & Username
  fMainForm.StatusBar1.Refresh
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "DataSource", ValString, cboDSNList.Text
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin", "Username", ValString, txtUID.Text
  Unload Me
Err_Handler:
  If Err = -2147217843 Then
    Screen.MousePointer = vbNormal
    fMainForm.StatusBar1.Panels(1).Text = "Not connected."
    MsgBox "Incorrect Password - Connection to datasource: " & cboDSNList.Text & _
           " failed!", vbCritical, "Connection Error"
    txtPWD.SelStart = 0
    txtPWD.SelLength = Len(txtPWD.Text)
    txtPWD.SetFocus
  Else
    Screen.MousePointer = vbNormal
    If Err.Number <> 0 Then LogError Err, "frmODBCLogon, cmdOK_Click"
    Exit Sub
  End If
End Sub

Private Sub Form_Activate()
On Error GoTo Err_Handler
  If cboDSNList.Text <> "" And txtUID.Text <> "" And txtPWD.Text <> "" Then cmdOK_Click
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmODBCLogon, Form_Activate"
  End
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Dim LastDSN As String
Dim X As Integer
Dim Options() As String
    Options = Split(Command, " ")
    X = 0
    GetDSNsAndDrivers
    txtUID.Text = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "Username", LCase(WinUserName))
    If UBound(Options) >= 0 Then
      LastDSN = Options(0)
    Else
      LastDSN = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin", "DataSource", "")
    End If
    If LastDSN <> "" Then
      Do While X <> cboDSNList.ListCount
        If UCase(cboDSNList.List(X)) = UCase(LastDSN) Then cboDSNList.ListIndex = X
        X = X + 1
      Loop
    End If
    If UBound(Options) >= 1 Then txtUID.Text = Options(1)
    If UBound(Options) >= 2 And SecondLogon = False Then txtPWD.Text = Options(2)
    SecondLogon = True
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmODBCLogon, Form_Load"
  End
End Sub

Public Sub GetDSNsAndDrivers()
On Error Resume Next
Dim i As Integer
Dim sDSNItem As String * 1024
Dim sDRVItem As String * 1024
Dim sDSN As String
Dim sDRV As String
Dim iDSNLen As Integer
Dim iDRVLen As Integer
Dim lHenv As Long         'handle to the environment

  'Clear the list (Bug #193)
  cboDSNList.Clear
  
  'get the DSNs
  If SQLAllocEnv(lHenv) <> -1 Then
    Do Until i <> SQL_SUCCESS
      sDSNItem = Space(1024)
      sDRVItem = Space(1024)
      i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
      sDSN = VBA.Left(sDSNItem, iDSNLen)
      sDRV = VBA.Left(sDRVItem, iDRVLen)
      If sDSN <> Space(iDSNLen) And sDRV = "PostgreSQL" Then cboDSNList.AddItem sDSN
    Loop
  End If
  cboDSNList.ListIndex = 0
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  txtPWD.SetFocus
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmODBCLogon, Form_Resize"
End Sub

Private Sub txtPWD_GotFocus()
On Error GoTo Err_Handler
  txtPWD.SelStart = 0
  txtPWD.SelLength = Len(txtPWD.Text)
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmODBCLogon, txtPWD_GotFocus"
End Sub

Private Sub txtUID_GotFocus()
On Error GoTo Err_Handler
  txtUID.SelStart = 0
  txtUID.SelLength = Len(txtUID.Text)
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmODBCLogon, txtUID_GotFocus"
End Sub

VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmReportViewer 
   Caption         =   "Report Viewer"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9150
   Icon            =   "frmReportViewer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   9150
   Begin CRVIEWERLibCtl.CRViewer crViewer 
      Height          =   6225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9105
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmReportViewer"
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
Dim rs As New Recordset

Private Sub Form_Load()
On Error GoTo Err_Handler
  LogMsg "Loading Form: " & Me.Name
  Me.Width = 8435
  Me.Height = 4455
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmReportViewer, Form_Load"
End Sub

Private Sub Form_Resize()
On Error GoTo Err_Handler
  If Me.WindowState <> 1 Then
    CRViewer.Height = Me.ScaleHeight
    CRViewer.Width = Me.ScaleWidth
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err, "frmReportViewer, Form_Resize"
End Sub

Public Function RunReport(szFile As String, szSQL As String, bShowTree As Boolean) As Integer
On Error GoTo Err_Handler
  StartMsg "Preparing Report..."
  Dim CryRpt As craxdrt.Report
  Dim CryApp As New craxdrt.Application
  If rs.State <> adStateClosed Then rs.Close
  LogMsg "Executing: " & szSQL
  rs.Open szSQL, gConnection
  Set CryRpt = CryApp.OpenReport(szFile, crOpenReportByTempCopy)
  If CryRpt.HasSavedData = True Then CryRpt.DiscardSavedData
  CryRpt.DisplayProgressDialog = True
  CryRpt.Database.SetDataSource rs, 3, 1
  If bShowTree = True Then
    CRViewer.EnableGroupTree = True
    CRViewer.DisplayGroupTree = True
  Else
    CRViewer.EnableGroupTree = False
    CRViewer.DisplayGroupTree = False
  End If
  CRViewer.ReportSource = CryRpt
  CRViewer.ViewReport
  Set CryRpt = Nothing
  Set CryApp = Nothing
  EndMsg
  Exit Function
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err, "frmReportViewer, RunReport"
End Function

Attribute VB_Name = "basGlobal"
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

Public Type rptDef
  szCategory As String
  szName As String
  szFile As String
  szSQL As String
  szAuthor As String
  szDescription As String
  bShowTree As Boolean
  bRefreshTables As Boolean
  bRefreshSequences As Boolean
End Type
Public Enum ObjectTypes
  tTable = 0
  tIndex = 1
  tSequence = 2
  tFunction = 3
  tTrigger = 4
  tView = 5
  tLanguage = 6
End Enum

Public Const SSO_VERSION = 2
Public Const DEVELOPMENT = False
Public Const QUOTE = """"
Public Const LAST_SYSTEM_OID = 18655
Public Const MIN_PGSQL_VERSION = 7

Public gConnection As New Connection
Public ActionCancelled As Boolean
Public fMainForm As frmMain
Public Datasource As String
Public Username As String
Public Connect As String
Public Password As String
Public QryTimer As Single
Public BBar As Variant
Public SQLPane As Variant
Public Tracking As Boolean
Public TrackVer As Single
Public Logging As Variant
Public MaskPassword As Variant
Public LogFile As String
Public OID As String
Public CallingForm As String
Public SecondLogon As Boolean
Public rptList() As rptDef
Public Exporters() As pgExporter


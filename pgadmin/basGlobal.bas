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

Public Const SSO_VERSION = 2.04
Public Const DEVELOPMENT = True
Public Const QUOTE = """"
Public Const LAST_SYSTEM_OID = 18655
Public Const MIN_PGSQL_VERSION = 7
Public Const DEFAULT_TEXT_COLOURS = "alter|0|0|16711680;comment|0|0|16711680;create|0|0|16711680;delete|0|0|16711680;drop|0|0|16711680;explain|0|0|16711680;grant|0|0|16711680;insert|0|0|16711680;revoke|0|0|16711680;" & _
                                    "select|0|0|16711680;update|0|0|16711680;vacuum|0|0|16711680;aggregate|0|0|255;constraint|0|0|255;database|0|0|255;function|0|0|255;group|0|0|255;index|0|0|255;" & _
                                    "language|0|0|255;operator|0|0|255;rule|0|0|255;sequence|0|0|255;table|0|0|255;trigger|0|0|255;abort|0|0|11998061;begin|0|0|11998061;" & _
                                    "checkpoint|0|0|11998061;close|0|0|11998061;cluster|0|0|11998061;commit|0|0|11998061;copy|0|0|11998061;declare|0|0|11998061;end|0|0|11998061;fetch|0|0|11998061;listen|0|0|11998061;" & _
                                    "load|0|0|11998061;lock|0|0|11998061;move|0|0|11998061;notify|0|0|11998061;reindex|0|0|11998061;reset|0|0|11998061;rollback|0|0|11998061;set|0|0|11998061;show|0|0|11998061;truncate|0|0|11998061;" & _
                                    "unlisten|0|0|11998061;asc|0|0|32768;ascending|0|0|32768;by|0|0|32768;desc|0|0|32768;descending|0|0|32768;from|0|0|32768;having|0|0|32768;into|0|0|32768;" & _
                                    "on|0|0|32768;order|0|0|32768;using|0|0|32768;where|0|0|32768;"

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
Public TextColours As String
Public OID As String
Public CallingForm As String
Public SecondLogon As Boolean
Public rptList() As rptDef
Public Exporters() As pgExporter
Public gPostgresOBJ_OID As Long
Public bContinueCompilation As Boolean

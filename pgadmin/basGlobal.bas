Attribute VB_Name = "basGlobal"
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

Public Const SSO_VERSION = 7#
Public Const DEVELOPMENT = True
Public Const QUOTE = """"
Public Const LAST_SYSTEM_OID = 18655
Public Const MIN_PGSQL_VERSION = 7
Public Const DEFAULT_TEXT_COLOURS = "ALTER|0|0|16711680;COMMENT|0|0|16711680;CREATE|0|0|16711680;DELETE|0|0|16711680;DROP|0|0|16711680;EXPLAIN|0|0|16711680;GRANT|0|0|16711680;INSERT|0|0|16711680;REVOKE|0|0|16711680;" & _
                                    "SELECT|0|0|16711680;UPDATE|0|0|16711680;VACUUM|0|0|16711680;AGGREGATE|0|0|255;CONSTRAINT|0|0|255;DATABASE|0|0|255;FUNCTION|0|0|255;GROUP|0|0|255;INDEX|0|0|255;" & _
                                    "LANGUAGE|0|0|255;OPERATOR|0|0|255;RULE|0|0|255;SEQUENCE|0|0|255;TABLE|0|0|255;TRIGGER|0|0|255;ABORT|0|0|11998061;BEGIN|0|0|11998061;" & _
                                    "CHECKPOINT|0|0|11998061;CLOSE|0|0|11998061;CLUSTER|0|0|11998061;COMMIT|0|0|11998061;COPY|0|0|11998061;DECLARE|0|0|11998061;FETCH|0|0|11998061;LISTEN|0|0|11998061;" & _
                                    "LOAD|0|0|11998061;LOCK|0|0|11998061;MOVE|0|0|11998061;NOTIFY|0|0|11998061;REINDEX|0|0|11998061;RESET|0|0|11998061;ROLLBACK|0|0|11998061;SET|0|0|11998061;SHOW|0|0|11998061;TRUNCATE|0|0|11998061;" & _
                                    "UNLISTEN|0|0|11998061;AND|0|0|32768;AS|0|0|32768;ASC|0|0|32768;ASCENDING|0|0|32768;BY|0|0|32768;CASE|0|0|32768;DESC|0|0|32768;DESCENDING|0|0|32768;ELSE|0|0|32768;FROM|0|0|32768;END|0|0|32768;HAVING|0|0|32768;INTO|0|0|32768;" & _
                                    "ON|0|0|32768;OR|0|0|32768;ORDER|0|0|32768;THEN|0|0|32768;USING|0|0|32768;WHEN|0|0|32768;WHERE|0|0|32768;"

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
Public SecondLogon As Boolean
Public rptList() As rptDef
Public Exporters() As pgExporter
Public DevMode As Boolean 'True = Development & Production mode, False = Production Mode only

' Global variables used to open AddForms
Public gPostgresOBJ_OID As Long

Public gFunction_Name As String
Public gFunction_Arguments As String

Public gTrigger_Name As String
Public gTrigger_Table As String

Public gView_Name As String

' Global variable used to control rebuilding
Public bContinueRebuilding As Boolean 'Stops rebuilding if false
Public gDevConnection As String 'Connection to repository
Public gDevPostgresqlTables As String



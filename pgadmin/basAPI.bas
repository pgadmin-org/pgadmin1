Attribute VB_Name = "basAPI"
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

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
Public Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Public Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)

Public Const SW_SHOWNORMAL = 1
Public Const SQL_SUCCESS As Long = 0
Public Const SQL_FETCH_NEXT As Long = 1
Public Const ODBC_ADD_DSN = 1            ' Add data source
Public Const ODBC_CONFIG_DSN = 2         ' Configure (edit) data source
Public Const ODBC_REMOVE_DSN = 3         ' Remove data source
Public Const ODBC_ADD_SYS_DSN = 4        ' Add a system DSN
Public Const ODBC_CONFIG_SYS_DSN = 5     ' Configure a system DSN
Public Const ODBC_REMOVE_SYS_DSN = 6     ' Remove a system DSN
Public Const ODBC_REMOVE_DEFAULT_DSN = 7 ' Remove the default DSN

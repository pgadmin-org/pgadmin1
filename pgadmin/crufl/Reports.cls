VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Public Function NumericFieldLength(lTypeMod As Long) As String
Dim szTemp As String
  szTemp = Hex((lTypeMod - 4) And &HFFFF)
  NumericFieldLength = CInt("&H" & Mid(szTemp, 1, Len(szTemp) - 4)) & "." & CInt("&H" & Mid(szTemp, Len(szTemp) - 3, Len(szTemp)))
End Function

Public Function TriggerForEach(iType As Integer) As String
  If (iType And 1) = 1 Then
    TriggerForEach = "Row"
  Else
    TriggerForEach = "Statement"
  End If
End Function

Public Function TriggerExecutes(iType As Integer) As String
  If (iType And 2) = 2 Then
    TriggerExecutes = "Before"
  Else
    TriggerExecutes = "After"
  End If
End Function

Public Function TriggerEvent(iType As Integer) As String
Dim szTemp As String
  If (iType And 4) = 4 Then szTemp = szTemp & "Insert "
  If (iType And 8) = 8 Then szTemp = szTemp & "Delete "
  If (iType And 16) = 16 Then szTemp = szTemp & "Update "
  TriggerEvent = RTrim(szTemp)
End Function

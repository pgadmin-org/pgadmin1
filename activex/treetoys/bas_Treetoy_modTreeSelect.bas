Attribute VB_Name = "bas_TreeToy_modTreeSelect"
' TreeToy
' Copyright (C) 2001, Jaroslaw Zwierz, AVE
' Poland, 04-628 Warszawa, ul. Alpejska 38
' tel./fax (+ 48 22) 815 68 99
' email: jerry@ave.com.pl

' Converted to an ActiveX Control by Dave Page (dpage@vale-housing.co.uk)
' for the pgAdmin project (http://www.greatbridge.org/project/pgadmin/)

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

Private Enum NodeCheck
    nodChecked = True
    nodUnchecked = False
    nodPartial = 1
End Enum

Public Function iTreeTextChecked(Tree As TreeView) As String
Dim nodX As Node
Dim iCountChecked As Integer
Dim szDelimiter As String

iTreeTextChecked = ""
iCountChecked = iTreeCountChecked(Tree)

If iCountChecked < 10 Then
    szDelimiter = vbCrLf
Else
    szDelimiter = ", "
End If

    For Each nodX In Tree.Nodes
        If nodX.Checked = True And Not (nodX.Parent Is Nothing) Then
            If iTreeTextChecked <> "" Then iTreeTextChecked = iTreeTextChecked & szDelimiter
            iTreeTextChecked = iTreeTextChecked & nodX.Parent.Text & "->" & nodX.Text
        End If
    Next
End Function

Public Function iTreeCountChecked(Tree As TreeView) As Integer
Dim nodX As Node
Dim iLoop As Integer
iLoop = 0

    For Each nodX In Tree.Nodes
        If nodX.Checked = True Then
            iLoop = iLoop + 1
        End If
    Next
iTreeCountChecked = iLoop
End Function

Public Function iTreeGetRoot(Node As MSComctlLib.Node) As String
    Dim szSelectedItem As String
    Dim szModeSplit() As String
    
    szSelectedItem = Node.FullPath
    iTreeGetRoot = ""
    
    If szSelectedItem <> "" Then
        szModeSplit = Split(szSelectedItem, "\")
        iTreeGetRoot = szModeSplit(0)
    End If
End Function


Public Sub iTreeSelectiveCheck(Node As MSComctlLib.Node)
    
    Node.Bold = False
    
    iTreeSetChildren Node, Node.Checked
    
    If TreeCheckSibling(Node, Node.Checked) Then
        TreeSetParents Node, Node.Checked
    Else
        TreeSetParents Node, nodPartial
    End If
    
End Sub

Public Sub iTreeSingleCheck(Tree As TreeView, Node As MSComctlLib.Node)
    
Dim nodX As Node
    If Node.Checked Then
        For Each nodX In Tree.Nodes
            If nodX.Index <> Node.Index And nodX.Checked Then
                nodX.Checked = False
            End If
        Next
    End If
End Sub

Public Sub iTreeSetChildren(Node As MSComctlLib.Node, bCheck As Boolean)

Dim nodX As Node
        
    If Node.Children = 0 Then
        Exit Sub
    End If
    
    Set nodX = Node.Child
    Do Until nodX Is Nothing
        nodX.Bold = False
        nodX.Checked = bCheck
        
        If nodX.Children > 0 Then
            iTreeSetChildren nodX, bCheck
        End If
        
        Set nodX = nodX.Next
    Loop
    
End Sub

Private Sub TreeSetParents(Node As MSComctlLib.Node, ByVal nCheck As NodeCheck)
    
Dim nodX As Node

    If (Node.Parent Is Nothing) Then
        Exit Sub
    End If
    
    Set nodX = Node.Parent
    Select Case nCheck
        Case nodChecked
            nodX.Checked = True
            nodX.Bold = False
        Case nodUnchecked
            nodX.Checked = False
            nodX.Bold = False
        Case nodPartial
            nodX.Checked = False
            nodX.Bold = True
    End Select
    
    If nCheck = nodPartial Then
        TreeSetParents nodX, nodPartial
    Else
        If TreeCheckSibling(nodX, nodX.Checked) Then
            TreeSetParents nodX, nodX.Checked
        Else
            TreeSetParents nodX, nodPartial
        End If
    End If
        
End Sub

Private Function TreeCheckSibling(Node As MSComctlLib.Node, ByVal bCheck As Boolean) As Boolean
    
Dim nodX As Node

    TreeCheckSibling = True
    
    Set nodX = Node.FirstSibling
    
    Do Until nodX Is Nothing
        If nodX.Checked <> bCheck Then
            TreeCheckSibling = False
            Exit Do
        End If
        
        Set nodX = nodX.Next
    Loop
    
End Function


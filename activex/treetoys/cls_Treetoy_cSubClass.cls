VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_TreeToy_cSubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private colMessage As New Collection
Private mHwnd As Long
Private mProcess As MessageProcess

Public Event WndProc(Msg As Long, wParam As Long, lParam As Long, nResult As Long)

Public Enum MessageProcess
    mpSendAndProcess = 0
    mpProcessAndSend
    mpProcessAndEat
End Enum

Property Let MessageProcessing(nValue As MessageProcess)
    mProcess = nValue
End Property

Property Get MessageProcessing() As MessageProcess
    MessageProcessing = mProcess
End Property


Property Let hWnd(hWndValue As Long)
    
    If hWndValue <> 0 Then
        
        If mHwnd <> 0 Then
            UnSubClass Me
        End If
        
        mHwnd = hWndValue
        bas_TreeToy_mSubClass.SubClass Me
        
        
    End If
    
End Property

Property Get hWnd() As Long
    
    hWnd = mHwnd
    
End Property

Public Sub AttachMessage(Msg As Long)

Dim i As Long

    On Error GoTo ERRTAG
    
    If Message(Msg) Then
        Exit Sub
    Else
        colMessage.Add Msg, "key" & Msg
    End If
        
    Exit Sub
ERRTAG:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
    Resume ERREND
ERREND:

End Sub

Public Sub DetachMessage(Msg As Long, bHook As Boolean)

Dim i As Long

    On Error GoTo ERRTAG
    
    If Message(Msg) Then
        colMessage.Remove "key" & Msg
    End If
        
    Exit Sub
ERRTAG:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
    Resume ERREND
ERREND:

End Sub

Public Function Message(Msg As Long) As Boolean

Dim i As Integer
    
    For i = 1 To colMessage.Count
        If colMessage(i) = Msg Then
            Message = True
            Exit Function
        End If
    Next
    
End Function


Private Sub Class_Terminate()

    Do Until colMessage.Count = 0
        colMessage.Remove 1
    Loop
    Set colMessage = Nothing
    
    If mHwnd <> 0 Then
        UnSubClass Me
    End If
    
End Sub

Friend Function RaiseWndProc(Msg As Long, wParam As Long, lParam As Long) As Long

Dim lRet As Long

    RaiseEvent WndProc(Msg, wParam, lParam, lRet)
    RaiseWndProc = lRet
    
End Function

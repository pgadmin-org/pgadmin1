Attribute VB_Name = "bas_TreeToy_mSubClass"
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

'This is common module, so we have to keep track of each
'cls_TreeToy_cSubClass instance to call correct Window Procedure
'Do not call this procedures outside cls_TreeToy_cSubClass

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Const GWL_WNDPROC = (-4)

Private Type SCInfo
    ProcOld As Long
    cSC As cls_TreeToy_cSubClass
End Type

Private arrSubClassInfo() As SCInfo
Private arrSubClassInfoCount As Long

Public Sub SubClass(cSC As cls_TreeToy_cSubClass)
'Do not use outside off cls_TreeToy_cSubClass
Dim i As Long

    For i = 0 To arrSubClassInfoCount - 1
        If arrSubClassInfo(i).cSC.hWnd = cSC.hWnd Then
            'Already subclassed
            Exit Sub
        End If
    Next
            
    arrSubClassInfoCount = arrSubClassInfoCount + 1
    ReDim Preserve arrSubClassInfo(arrSubClassInfoCount)
    
    With arrSubClassInfo(arrSubClassInfoCount - 1)
        Set .cSC = cSC
        .ProcOld = GetWindowLong(.cSC.hWnd, GWL_WNDPROC)
        SetWindowLong .cSC.hWnd, GWL_WNDPROC, AddressOf MyProc
    End With
    
End Sub

Public Sub UnSubClass(cSC As cls_TreeToy_cSubClass)
'Do not use outside off cls_TreeToy_cSubClass
Dim hWnd As Long
Dim i As Long
Dim j As Long
    
    hWnd = cSC.hWnd
    
    For i = 0 To arrSubClassInfoCount - 1
        If arrSubClassInfo(i).cSC.hWnd = hWnd Then
            SetWindowLong hWnd, GWL_WNDPROC, arrSubClassInfo(i).ProcOld
            
            'Remove item from array
            arrSubClassInfoCount = arrSubClassInfoCount - 1
            For j = i To arrSubClassInfoCount
                arrSubClassInfo(j) = arrSubClassInfo(j + 1)
            Next j
            ReDim Preserve arrSubClassInfo(arrSubClassInfoCount)
            
            Exit For
        End If
    Next
    
End Sub


Private Function MyProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim i As Long

    For i = 0 To arrSubClassInfoCount - 1
        With arrSubClassInfo(i)
            If .cSC.hWnd = hWnd Then
                If .cSC.Message(Msg) Then
                    If .cSC.MessageProcessing = mpSendAndProcess Then
                        'Send original message to window before processing
                        MyProc = CallWindowProc(.ProcOld, hWnd, Msg, wParam, lParam)
                    End If
                
                    'Fire WndProc event of cSC (Custom processing)
                    MyProc = .cSC.RaiseWndProc(Msg, wParam, lParam)
                    
                    If .cSC.MessageProcessing = mpProcessAndSend Then
                        'Send original message to window after processing
                        MyProc = CallWindowProc(.ProcOld, hWnd, Msg, wParam, lParam)
                    End If
                
                Else
                    'Call original window procedure
                    MyProc = CallWindowProc(.ProcOld, hWnd, Msg, wParam, lParam)
                End If
                
                Exit For
            End If
        End With
    Next
    
End Function


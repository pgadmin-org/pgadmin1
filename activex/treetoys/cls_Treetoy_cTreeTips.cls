VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_TreeToy_cTreeTips"
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

'Properties
Public ShowIconsInNodeTips As Boolean
Public ShowIconsInScrollTips As Boolean

Private mNodeTips As TipType
Private mScrollTips As TipType

Private mTree As MSComctlLib.TreeView

Private WithEvents cSub As cls_TreeToy_cSubClass
Attribute cSub.VB_VarHelpID = -1

'Other
Private mTPPX As Long
Private mTPPY As Long
Private ProcOld As Long

'General API
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Any, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal bytes As Long)
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As tagTRACKMOUSEEVENT) As Long

'Messages & Flags
Private Const WM_MOUSEMOVE = &H200
Private Const WM_VSCROLL = &H115
Private Const SB_ENDSCROLL = 8
Private Const SB_THUMBTRACK = 5
Private Const WM_MOUSEHOVER = &H2A1&
Private Const WM_MOUSELEAVE = &H2A3&
Private Const TME_HOVER = &H1&
Private Const TME_LEAVE = &H2&
Private Const TME_QUERY = &H40000000
Private Const TME_CANCEL = &H80000000
Private Const HOVER_DEFAULT = &HFFFFFFFF

'Types
Private Type tagTRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Neccessary Treeview Definitions
Private Const TV_FIRST      As Long = &H1100
Private Const TVM_GETNEXTITEM  As Long = (TV_FIRST + 10)
Private Const TVM_GETITEMRECT = (TV_FIRST + 4)
Private Const TVGN_FIRSTVISIBLE = &H5

Public Property Set Tree(TreeView1 As MSComctlLib.TreeView)

    Set mTree = TreeView1
    
    Set cSub = New cls_TreeToy_cSubClass
    
    cSub.hWnd = mTree.hWnd
    cSub.AttachMessage WM_MOUSEMOVE
    cSub.AttachMessage WM_MOUSELEAVE
    cSub.AttachMessage WM_VSCROLL
    
End Property

Public Property Let NodeTips(Value As TipType)

    mNodeTips = Value
    
End Property

Public Property Get NodeTips() As TipType

    NodeTips = mNodeTips
    
End Property

Public Property Let ScrollTips(Value As TipType)

    mScrollTips = Value
    
End Property

Public Property Get ScrollTips() As TipType

    ScrollTips = mScrollTips
    
End Property

Private Sub Class_Initialize()

    'Property cache
    mTPPX = Screen.TwipsPerPixelX
    mTPPY = Screen.TwipsPerPixelY
    
End Sub

Private Sub Class_Terminate()

    If Not mTree Is Nothing Then
        Set cSub = Nothing
        Set mTree = Nothing
    End If
    
End Sub

Private Function LowWord(ByVal Value As Long) As Integer
    CopyMemory LowWord, Value, 2
End Function

Private Function HiWord(ByVal Value As Long) As Integer
    CopyMemory HiWord, ByVal VarPtr(Value) + 2, 2
End Function


Private Sub StartTracking(hWnd As Long)

Dim tET As tagTRACKMOUSEEVENT
Dim lR As Long

    On Error Resume Next
    
    'Fires message procesing (WM_MOUSEHOVER & WM_MOUSELEAVE) by Window given by hwnd
    tET.cbSize = Len(tET)
    tET.dwFlags = TME_LEAVE Or TME_HOVER
    tET.dwHoverTime = HOVER_DEFAULT
    tET.hwndTrack = hWnd
    
    lR = TrackMouseEvent(tET)
    
End Sub


Private Function TreeView_GetNextItem(hWnd As Long, hItem As Long, Flag As Long) As Long

   TreeView_GetNextItem = SendMessageAny(hWnd, TVM_GETNEXTITEM, Flag, ByVal hItem)
   
End Function


Private Function GetFirstVisibleNode() As MSComctlLib.Node

Dim hItem As Long
Dim rc As RECT
Dim ret As Boolean
Dim nodX As Node
Dim fItemRect As Long
Dim Flag As Long

    'Get First Visible Item
    hItem = TreeView_GetNextItem(mTree.hWnd, 0, TVGN_FIRSTVISIBLE)
    If hItem = 0 Then
        Exit Function
    End If
    
    rc.Left = hItem
    fItemRect = 1
    
    ret = SendMessageAny(mTree.hWnd, TVM_GETITEMRECT, ByVal fItemRect, rc)

    If ret Then
        Set GetFirstVisibleNode = mTree.HitTest((rc.Left) * Screen.TwipsPerPixelX, (rc.Top) * Screen.TwipsPerPixelY)
    End If
    
End Function

Private Sub cSub_WndProc(Msg As Long, wParam As Long, lParam As Long, nResult As Long)

Static bVisible As Boolean

Dim nodX As Node
Dim x As Long
Dim y As Long
Dim TipText As String
Dim TipImage As StdPicture

    On Error Resume Next

    'Process Messages
    If mNodeTips <> ttNone Then
        Select Case Msg
            Case WM_MOUSELEAVE
                Unload frmToolTip
                
                        
            Case WM_MOUSEMOVE
                StartTracking mTree.hWnd
                x = LowWord(lParam) * mTPPX
                y = HiWord(lParam) * mTPPY
                
                Set nodX = mTree.HitTest(x, y)
                If Not nodX Is Nothing Then
                    Select Case mNodeTips
                        Case ttTag
                            TipText = nodX.Tag
                        Case ttText
                            TipText = nodX.Text
                        Case ttPath
                            TipText = nodX.FullPath
                        Case ttKey
                            TipText = nodX.Key
                    End Select
                    
                    If ShowIconsInNodeTips Then
                        If Not IsEmpty(nodX.Image) Then
                            Set TipImage = mTree.ImageList.ListImages(nodX.Image).ExtractIcon
                        End If
                    End If
                    
                    If TipText <> "" Then frmToolTip.ShowToolTip TipText, TipImage
                Else
                    Unload frmToolTip
                End If
        End Select
    End If
    
    If mScrollTips <> ttNone Then
                 
        If Msg = WM_VSCROLL Then
            Select Case LowWord(wParam)

                Case SB_THUMBTRACK
                    Set nodX = GetFirstVisibleNode()
                    If Not nodX Is Nothing Then
                        Select Case mScrollTips
                            Case ttTag
                                TipText = nodX.Tag
                            Case ttText
                                TipText = nodX.Text
                            Case ttPath
                                TipText = nodX.FullPath
                            Case ttKey
                                TipText = nodX.Key
                        End Select
                        
                        If TipText = "" Then Exit Sub
                        
                        If ShowIconsInScrollTips Then
                            If Not IsEmpty(nodX.Image) Then
                                Set TipImage = mTree.ImageList.ListImages(nodX.Image).ExtractIcon
                            End If
                        End If
                        
                        If bVisible Then
                            If ShowIconsInScrollTips Then
                                frmToolTip.SetImage TipImage, True
                            End If
                            frmToolTip.SetText TipText
                            
                        Else
                            If TipText <> "" Then frmToolTip.ShowToolTip TipText, TipImage
                            bVisible = True
                        End If
                    End If
                    
              
                Case SB_ENDSCROLL
                    Unload frmToolTip
                    bVisible = False
            End Select
        End If
    End If
    
End Sub

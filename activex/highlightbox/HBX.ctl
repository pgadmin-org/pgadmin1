VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl HBX 
   ClientHeight    =   1065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   PropertyPages   =   "HBX.ctx":0000
   ScaleHeight     =   1065
   ScaleWidth      =   2865
   ToolboxBitmap   =   "HBX.ctx":0019
   Begin RichTextLib.RichTextBox rtbstring 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   180
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"HBX.ctx":0113
   End
   Begin VB.Image imgdn 
      Height          =   150
      Left            =   2670
      Picture         =   "HBX.ctx":019E
      Top             =   30
      Width           =   150
   End
   Begin VB.Image imgup 
      Height          =   150
      Left            =   2670
      Picture         =   "HBX.ctx":04F8
      Top             =   30
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000009&
      Height          =   225
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   2475
   End
   Begin VB.Shape shpBar 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   0
      Top             =   0
      Width           =   2865
   End
End
Attribute VB_Name = "HBX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' HBX - Auto Highlighting Expanding text box
' Copyright (C) 2001, Mark Yeatman

' This library is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public
' License as published by the Free Software Foundation; either
' version 2.1 of the License, or (at your option) any later version.

' This library is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.

' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

Option Explicit
'Default Property Values:
Const m_def_Wordlist = "0"

Dim m_ForeColor As Long
Dim m_BackStyle As Boolean
Dim m_ControlBarVisible As Boolean
Dim m_MaximisedWidth As Variant
Dim iControlTop As Integer
Dim iControlLeft As Integer

Const m_def_ControlBarVisible = 0
Const m_def_MaximisedHeight = 0
Const m_def_ForeColor = 0
Const m_def_BackStyle = False
Const m_def_MaximisedWidth = 0

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=rtbstring,rtbstring,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=rtbstring,rtbstring,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=rtbstring,rtbstring,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=rtbstring,rtbstring,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=rtbstring,rtbstring,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=rtbstring,rtbstring,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=rtbstring,rtbstring,-1,MouseUp

Public Type WordStyle
  szColour As String
  szRTFstring As String
  szString As String
  bBold As Boolean
  bItalic As Boolean
End Type


Dim szRTBFontinfo As String
Dim szRTBWordinfo As String
Dim szRTBColours As String
Const DELIMCHARS = " []{}()'"""
'Property Variables:
Dim m_Wordlist As String
Dim m_MaximisedHeight As Variant
Dim bMaximised As Boolean

Private Sub rtbstring_Change()
Static lPrevLen As Long
Dim lCount As Long
lCount = Len(rtbstring.Text)
If (lCount > lPrevLen + 2) Or (lCount < lPrevLen - 2) Then QR
lPrevLen = lCount
End Sub

Private Sub ColourWord() ' This is a coulour a word at a time test
On Error Resume Next
Dim lWordend As Long
Dim lWordstart As Long
Dim szChar As String
Dim szTemp As String

lWordend = rtbstring.SelStart
If lWordend < 1 Then Exit Sub

lWordstart = lWordend
szChar = Mid(rtbstring.Text, lWordstart - 1, 1)
If szChar = Chr(10) Then
 MsgBox "Return"
End If

While (CharInstr(szChar, DELIMCHARS) = False) And szChar <> Chr(10)
  If lWordstart = 1 Then GoTo fred
  lWordstart = lWordstart - 1
  szChar = Mid(rtbstring.Text, lWordstart, 1)
Wend

fred:
If lWordstart - 1 = 0 Then
  szTemp = Mid(rtbstring.Text, lWordstart, (lWordend - lWordstart))
  rtbstring.SelStart = lWordstart - 1
  rtbstring.SelLength = (lWordend - lWordstart)
Else
  rtbstring.SelStart = lWordstart + 1
  rtbstring.SelLength = (lWordend - lWordstart)
  szTemp = Mid(rtbstring.Text, rtbstring.SelStart, rtbstring.SelLength)
  rtbstring.SelStart = lWordstart
  rtbstring.SelLength = (lWordend - lWordstart)
End If

If WordCache(szTemp).szString & "" <> "" Then
  rtbstring.SelBold = WordCache(szTemp).bBold
  rtbstring.SelItalic = WordCache(szTemp).bItalic
  rtbstring.SelColor = Val(WordCache(szTemp).szColour) 'RGB(0, 0, 255)
End If

rtbstring.SelStart = lWordend + 1 'Reset cursor position
rtbstring.SelBold = False
rtbstring.SelItalic = False
rtbstring.SelColor = RGB(0, 0, 0)

End Sub
Private Sub QR()
On Error Resume Next
Dim lCurpos As Long
Dim Stringlist() As String
Dim szOutputstring As String
Dim szTemp As String
Dim szChar As String
Dim lArraypos As Long
Dim szData As String
Dim lX As Long
ReDim Stringlist(Len(rtbstring.Text))

lCurpos = rtbstring.SelStart

szTemp = ""
szData = rtbstring.Text
For lX = 1 To Len(szData)
  szChar = Mid(szData, lX, 1)
   If CharInstr(szChar, DELIMCHARS) = False And szChar <> Chr(10) Then
    szTemp = szTemp & szChar
  Else
    If szChar = Chr(10) Then szChar = "\par "
    If szTemp <> "" Then
      Stringlist(lArraypos) = szTemp
      lArraypos = lArraypos + 1
    End If
    Stringlist(lArraypos) = szChar
    lArraypos = lArraypos + 1
    szTemp = ""
  End If
Next lX

If szTemp <> "" Then Stringlist(lArraypos) = szTemp

For lX = 0 To UBound(Stringlist) - 1
  If Stringlist(lX) <> "" Then Stringlist(lX) = SearchCache(Stringlist(lX))
Next lX
  
For lX = 0 To UBound(Stringlist)
  If Stringlist(lX) <> "" Then szOutputstring = szOutputstring & Stringlist(lX)
Next lX

rtbstring.TextRTF = szRTBFontinfo & szRTBColours & szRTBWordinfo & szOutputstring & " \par }"
rtbstring.SelStart = lCurpos
End Sub
Friend Sub BuildCache()
  Dim szStrings() As String
  Dim szValues() As String
  Dim szColours() As String
  Dim szRTBtmp As String
  Dim iLoop As Integer
  Dim WordDisplay As WordStyle
  Dim colRGB As T_RGB
  Dim iX As Integer
  Dim bColour As Boolean
  
  rtbstring.Text = "|"
  szRTBFontinfo = Mid(rtbstring.TextRTF, 1, (InStr(1, rtbstring.TextRTF, "|") - 1))
  szRTBWordinfo = Mid(szRTBFontinfo, InStr(1, szRTBFontinfo, "}}") + 2)
  szRTBFontinfo = Mid(szRTBFontinfo, 1, InStr(1, szRTBFontinfo, "}}") + 1)
  ReDim szColours(0)
  
  rtbstring.TextRTF = ""
  Set WordCache = New Collection
  szRTBColours = "{\colortbl ;\red0\green0\blue0"
  szStrings = Split(m_Wordlist, ";")
  
  For iLoop = 0 To UBound(szStrings) - 1
    szValues = Split(szStrings(iLoop), "|")
    
    szRTBtmp = szValues(0)   ' Remove the ucase this line from if you upper case require not
    WordDisplay.szString = szValues(0)
    If szValues(1) = "1" Then
      szRTBtmp = "\b " & szRTBtmp & "\b "
      WordDisplay.bBold = True
    Else
      WordDisplay.bBold = False
    End If
    
    If szValues(2) = "1" Then
      szRTBtmp = "\i " & szRTBtmp & "\i "
      WordDisplay.bItalic = True
    Else
      WordDisplay.bItalic = False
    End If
        
    WordDisplay.szColour = szValues(3)
    
    
    'Search for colour
    
    bColour = False
    For iX = 0 To UBound(szColours)
      If szColours(iX) = szValues(3) Then
        bColour = True
        Exit For
      End If
    Next iX
         
    If bColour = False Then
      szColours(UBound(szColours)) = szValues(3)
      WordDisplay.szRTFstring = ("\cf" & UBound(szColours) + 2) & " " & szRTBtmp & "\cf0 "
      ReDim Preserve szColours(UBound(szColours) + 1)
    Else
      WordDisplay.szRTFstring = ("\cf" & iX + 2) & " " & szRTBtmp & "\cf0 "
    End If
    WordCache.Add WordDisplay, szValues(0)

    szRTBtmp = ""
  Next iLoop
  
  For iX = 0 To UBound(szColours) - 1
    colRGB = get_RGB("" & szColours(iX))
    szRTBColours = szRTBColours & ";" & "\red" & colRGB.R & "\green" & colRGB.G & "\blue" & colRGB.B
  Next iX
          
  szRTBColours = szRTBColours & ";}"
End Sub
Private Function findControl(szName As String) As Integer
Dim iX As Integer
  For iX = 0 To UserControl.ParentControls.Count - 1
    If UserControl.ParentControls.Item(iX).Name = szName Then
      findControl = iX
    End If
  Next iX
End Function

Private Sub imgdn_Click()
Dim ifoundcontrol As Integer
  
  MinimisedHeight = UserControl.Height
  MinimisedWidth = UserControl.Width
  LastTop = UserControl.Extender.Top
  Lastleft = UserControl.Extender.Left
  
  imgdn.Visible = False
  imgup.Visible = True
  ifoundcontrol = findControl(UserControl.Ambient.DisplayName)
  If ifoundcontrol = 0 Then Exit Sub
  iControlLeft = UserControl.ParentControls.Item(ifoundcontrol).Left
  iControlTop = UserControl.ParentControls.Item(ifoundcontrol).Top

  If m_MaximisedHeight = 0 Then
    UserControl.Height = UserControl.Parent.Height
    UserControl.ParentControls.Item(ifoundcontrol).Move 0, 0
  Else
    If m_MaximisedHeight > MinimisedHeight Then UserControl.Height = m_MaximisedHeight
  End If
  
  If m_MaximisedWidth = 0 Then
    UserControl.Width = UserControl.Parent.Width - 100
    UserControl.ScaleLeft = 0
    Else
    If m_MaximisedWidth > MinimisedWidth Then UserControl.Width = m_MaximisedWidth
  End If
  rtbstring.Height = (UserControl.Height - (3 * shpBar.Height))
  bMaximised = True
  UserControl.ParentControls.Item(ifoundcontrol).ZOrder 0
End Sub

Private Sub imgup_Click()
Dim ifoundcontrol As Integer
  imgup.Visible = False
  imgdn.Visible = True
  ifoundcontrol = findControl(UserControl.Ambient.DisplayName)
  If ifoundcontrol = 0 Then
    imgup.Visible = True
    imgdn.Visible = False
    Exit Sub
  End If
  
  UserControl.ParentControls.Item(ifoundcontrol).Move iControlLeft, iControlTop
  UserControl.Height = MinimisedHeight
  UserControl.Width = MinimisedWidth
  bMaximised = False
  End Sub

Private Sub rtbstring_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or _
   KeyCode = vbKeyReturn Or _
   KeyCode = vbKeySpace Or _
   KeyCode = vbKeyDelete Then ColourWord

RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Resize()
  If UserControl.Width < 10 Then UserControl.Width = 10
  If UserControl.Height < 500 Then UserControl.Height = 500
  shpBar.Width = UserControl.Width
  rtbstring.Top = 0 + shpBar.Height
  rtbstring.Width = UserControl.Width
  rtbstring.Height = (UserControl.Height - shpBar.Height)
  imgup.Left = UserControl.Width - 200
  imgdn.Left = UserControl.Width - 200
  
  If bMaximised = True Then imgup_Click

  'ShowBar m_ControlBarVisible
End Sub
Private Sub ShowBar(bSho As Boolean)

    If bSho = True Then
    shpBar.Height = 195
    rtbstring.Top = 0 + shpBar.Height
    rtbstring.Height = UserControl.Height
  Else
    shpBar.Height = 0
    rtbstring.Top = 0 + shpBar.Height
    rtbstring.Height = UserControl.Height
  End If
End Sub

Private Sub rtbstring_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub rtbstring_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub rtbstring_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub rtbstring_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub rtbstring_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub rtbstring_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub
Public Property Get MaximisedWidth() As Variant
  MaximisedWidth = m_MaximisedWidth
End Property

Public Property Let MaximisedWidth(ByVal New_MaximisedWidth As Variant)
  m_MaximisedWidth = New_MaximisedWidth
  PropertyChanged "MaximisedWidth"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_MaximisedWidth = m_def_MaximisedWidth
  Set UserControl.Font = Ambient.Font
  m_MaximisedHeight = m_def_MaximisedHeight
  m_Wordlist = m_def_Wordlist
  BuildCache
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  rtbstring.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
  rtbstring.Enabled = PropBag.ReadProperty("Enabled", True)
  Set rtbstring.Font = PropBag.ReadProperty("Font", Ambient.Font)
  rtbstring.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
  rtbstring.SelText = PropBag.ReadProperty("SelText", "")
  rtbstring.SelStart = PropBag.ReadProperty("SelStart", 0)
  rtbstring.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
  rtbstring.Locked = PropBag.ReadProperty("Locked", False)
  rtbstring.Text = PropBag.ReadProperty("Text", "HBX")
  rtbstring.RightMargin = PropBag.ReadProperty("RightMargin", 0)
  rtbstring.Locked = PropBag.ReadProperty("Locked", False)
  rtbstring.MaxLength = PropBag.ReadProperty("MaxLength", 0)
  m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
  m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
  lblCaption.Caption = PropBag.ReadProperty("Caption", "")
  rtbstring.Text = PropBag.ReadProperty("Text", "RichTextBox1")
  lblCaption.Caption = PropBag.ReadProperty("Caption", "")
  m_ControlBarVisible = PropBag.ReadProperty("ControlBarVisible", m_def_ControlBarVisible)
  m_Wordlist = PropBag.ReadProperty("Wordlist", m_def_Wordlist)
  BuildCache
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("BackColor", rtbstring.BackColor, &H80000005)
  Call PropBag.WriteProperty("Enabled", rtbstring.Enabled, True)
  Call PropBag.WriteProperty("Font", rtbstring.Font, Ambient.Font)
  Call PropBag.WriteProperty("Locked", rtbstring.Locked, False)
  Call PropBag.WriteProperty("MaximisedHeight", m_MaximisedHeight, m_def_MaximisedHeight)
  Call PropBag.WriteProperty("BorderStyle", rtbstring.BorderStyle, 1)
  Call PropBag.WriteProperty("MaximisedWidth", m_MaximisedWidth, m_def_MaximisedWidth)
  Call PropBag.WriteProperty("SelText", rtbstring.SelText, "")
  Call PropBag.WriteProperty("SelStart", rtbstring.SelStart, 0)
  Call PropBag.WriteProperty("ToolTipText", rtbstring.ToolTipText, "")
  Call PropBag.WriteProperty("Caption", lblCaption.Caption, "")
  Call PropBag.WriteProperty("MaxLength", rtbstring.MaxLength, 0)
  Call PropBag.WriteProperty("Text", rtbstring.Text, "HBX")
  Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
  Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
  Call PropBag.WriteProperty("RightMargin", rtbstring.RightMargin, 0)
  Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
  Call PropBag.WriteProperty("Locked", rtbstring.Locked, False)
  Call PropBag.WriteProperty("MaxLength", rtbstring.MaxLength, 0)
  Call PropBag.WriteProperty("ToolTipText", rtbstring.ToolTipText, "")
  Call PropBag.WriteProperty("ScrollBars", rtbstring.ScrollBars, 0)
  Call PropBag.WriteProperty("Text", rtbstring.Text, "RichTextBox1")
  Call PropBag.WriteProperty("Caption", lblCaption.Caption, "")
  Call PropBag.WriteProperty("ControlBarVisible", m_ControlBarVisible, m_def_ControlBarVisible)
  Call PropBag.WriteProperty("Wordlist", m_Wordlist, m_def_Wordlist)
End Sub
Public Function Maximise() As Variant
imgup_Click
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
  BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
  UserControl.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtbstring,rtbstring,-1,Locked
Public Property Get Locked() As Boolean
  Locked = rtbstring.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
  rtbstring.Locked() = New_Locked
  PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtbstring,rtbstring,-1,MaxLength
Public Property Get MaxLength() As Long
  MaxLength = rtbstring.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
  rtbstring.MaxLength() = New_MaxLength
  PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtbstring,rtbstring,-1,ToolTipText
Public Property Get ToolTipText() As String
  ToolTipText = rtbstring.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
  rtbstring.ToolTipText() = New_ToolTipText
  PropertyChanged "ToolTipText"
End Property

Public Property Get MaximisedHeight() As Variant
  MaximisedHeight = m_MaximisedHeight
End Property

Public Property Let MaximisedHeight(ByVal New_MaximisedHeight As Variant)
  m_MaximisedHeight = New_MaximisedHeight
  PropertyChanged "MaximisedHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtbstring,rtbstring,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
  BackColor = rtbstring.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  rtbstring.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtbstring,rtbstring,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = rtbstring.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  rtbstring.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtbstring,rtbstring,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = rtbstring.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set rtbstring.Font = New_Font
  PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtbstring,rtbstring,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a control."
  rtbstring.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtbstring,rtbstring,-1,Text
Public Property Get Text() As String
  Text = rtbstring.Text
End Property

Public Property Let Text(ByVal New_Text As String)
  rtbstring.Text() = New_Text
  PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
  Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
  lblCaption.Caption() = New_Caption
  PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ControlBarVisible() As Boolean
Attribute ControlBarVisible.VB_Description = "Shows / Hides the Control Bar"
  ControlBarVisible = m_ControlBarVisible
End Property

Public Property Let ControlBarVisible(ByVal New_ControlBarVisible As Boolean)
  m_ControlBarVisible = New_ControlBarVisible
  PropertyChanged "ControlBarVisible"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub HUP()
Attribute HUP.VB_Description = "Causes a Re-Colouring of the control"
QR
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Wordlist() As String
  Wordlist = m_Wordlist
End Property

Public Property Let Wordlist(ByVal New_Wordlist As String)
  m_Wordlist = New_Wordlist
  PropertyChanged "Wordlist"
End Property


VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl HBX 
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2790
   PropertyPages   =   "HBX.ctx":0000
   ScaleHeight     =   1215
   ScaleWidth      =   2790
   ToolboxBitmap   =   "HBX.ctx":0019
   Begin RichTextLib.RichTextBox rtbString 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"HBX.ctx":0113
   End
End
Attribute VB_Name = "HBX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' HighlightBOX - Auto highlighting text box
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

Const m_def_ForeColor = 0
Const m_def_BackStyle = False
Const m_def_MultiLine = 0
Const m_def_WordList = ""
Dim m_ForeColor As Long
Dim m_BackStyle As Boolean
Dim m_MultiLine As Boolean
Dim m_WordList As String
Event Click() 'MappingInfo=rtbString,rtbString,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=rtbString,rtbString,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=rtbString,rtbString,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=rtbString,rtbString,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=rtbString,rtbString,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=rtbString,rtbString,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses a mouse button."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=rtbString,rtbString,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=rtbString,rtbString,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user presses and releases a mouse button."

Public Type WordStyle
  szColour As String
  bBold As Boolean
  bItalic As Boolean
End Type

Dim WordCache As Collection


Private Sub rtbstring_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or _
  KeyCode = vbKeySpace Or _
  KeyCode = vbKeyReturn Or _
  KeyCode = vbKeyUp Or _
  KeyCode = vbKeyDown Or _
  KeyCode = vbKeyLeft Or _
  KeyCode = vbKeyRight Then ColourWord
RaiseEvent KeyUp(KeyCode, Shift)

End Sub
Private Sub ColourWord()

Dim lWordend As Long
Dim lWordstart As Long
Dim szColour As String
Dim bBold As Boolean
Dim bItalic As Boolean
Dim szChar As String

lWordend = rtbString.SelStart - 1
If lWordend < 1 Then Exit Sub

lWordstart = lWordend
'search back for start of word

'While Mid(rtbString.Text, lWordstart, 1) <> " "
szChar = Mid(rtbString.Text, lWordstart, 1)
While szChar <> " " And szChar <> ")" And _
      szChar <> "(" And szChar <> "}" And _
      szChar <> "{" And szChar <> "]" And _
      szChar <> "[" And szChar <> ";" And _
      szChar <> ":" And szChar <> Chr(10)
  
  lWordstart = lWordstart - 1
  If lWordstart = 0 Then GoTo fred
  szChar = Mid(rtbString.Text, lWordstart, 1)
Wend

fred:
rtbString.SelStart = lWordstart
rtbString.SelLength = lWordend - lWordstart

CheckString Mid(rtbString.Text, rtbString.SelStart + 1, rtbString.SelLength), szColour, bBold, bItalic
rtbString.SelBold = bBold
rtbString.SelItalic = bItalic
rtbString.SelColor = Val(szColour) 'RGB(0, 0, 255)
      
rtbString.SelStart = lWordend + 1 'Reset cursor position



End Sub



Private Sub ColourBox()

Dim lCount As Long
Dim lRestart As Long
Dim lWordlen As Long
Dim lCursorPos As Long
Dim lSelectedLength As Long
Dim szMidString As String
Dim szWordString As String

Dim szColour As String
Dim bBold As Boolean
Dim bItalic As Boolean

  lRestart = 1
  lWordlen = 0
  lCursorPos = rtbString.SelStart
  lSelectedLength = rtbString.SelLength
  
  'Clear all to black
  rtbString.SelStart = 1
  rtbString.SelLength = Len(rtbString.Text)
  'rtbString.SelColor = RGB(0, 0, 0)
  'rtbString.SelBold = False
  'rtbString.SelItalic = False
  'rtbString.SelStart = 1
  'rtbString.SelLength = 0
  
  For lCount = 1 To Len(rtbString.Text)
    If Mid(rtbString.Text, lCount, 1) = " " Or Mid(rtbString.Text, lCount, 1) = ")" Or _
      Mid(rtbString.Text, lCount, 1) = "(" Or Mid(rtbString.Text, lCount, 1) = "}" Or _
      Mid(rtbString.Text, lCount, 1) = "{" Or Mid(rtbString.Text, lCount, 1) = "]" Or _
      Mid(rtbString.Text, lCount, 1) = "[" Or Mid(rtbString.Text, lCount, 1) = ";" Or _
      Mid(rtbString.Text, lCount, 1) = ":" Or Mid(rtbString.Text, lCount, 1) = Chr(10) Then
      szMidString = Mid(rtbString.Text, lRestart, lWordlen)
      szMidString = Trim(szMidString)
      DoEvents
      CheckString szMidString, szColour, bBold, bItalic
      rtbString.SelStart = lRestart - 1
      rtbString.SelLength = lWordlen
      rtbString.SelBold = bBold
      rtbString.SelItalic = bItalic
      rtbString.SelColor = Val(szColour) 'RGB(0, 0, 255)
      'rtbString.SelStart = lCursorPos
      'rtbString.SelLength = lSelectedLength
      
      lRestart = lCount + 1
      lWordlen = 0
    Else
      lWordlen = lWordlen + 1
      DoEvents
    End If
    
  Next lCount
  
  rtbString.SelStart = lCursorPos
  rtbString.SelLength = lSelectedLength
End Sub
Private Sub CheckString(szCheck As String, szColour As String, bBold As Boolean, bItalic As Boolean)
  On Error Resume Next
      
  'lookup - if exists
  szColour = 0
  bBold = False
  bItalic = False
  
   
  'search WordCache
  szColour = WordCache(szCheck).szColour
  bBold = WordCache(szCheck).bBold
  bItalic = WordCache(szCheck).bItalic

End Sub
Private Sub UserControl_Resize()
If UserControl.Height < 315 Then UserControl.Height = 315
If UserControl.Width < 500 Then UserControl.Width = 500
rtbString.Width = UserControl.Width
rtbString.Height = UserControl.Height
End Sub
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color of an object."
  BackColor = rtbString.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  rtbString.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = rtbString.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
  rtbString.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = rtbString.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
  Set rtbString.Font = New_Font
  PropertyChanged "Font"
End Property
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = rtbString.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
  rtbString.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a control."
  rtbString.Refresh
End Sub
Private Sub rtbString_Click()
  RaiseEvent Click
End Sub
Private Sub rtbString_DblClick()
  RaiseEvent DblClick
End Sub
Private Sub rtbString_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub rtbString_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub rtbString_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub
Private Sub rtbString_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub
Private Sub rtbString_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub
Private Sub UserControl_InitProperties()
  m_WordList = m_def_WordList
  m_ForeColor = m_def_ForeColor
  m_BackStyle = m_def_BackStyle
  m_MultiLine = m_def_MultiLine
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  rtbString.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
  rtbString.Enabled = PropBag.ReadProperty("Enabled", True)
  Set rtbString.Font = PropBag.ReadProperty("Font", Ambient.Font)
  rtbString.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
  m_WordList = PropBag.ReadProperty("WordList", m_def_WordList)
  rtbString.Text = PropBag.ReadProperty("Text", "HBX")
  m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
  m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
  m_MultiLine = PropBag.ReadProperty("MultiLine", m_def_MultiLine)
  rtbString.RightMargin = PropBag.ReadProperty("RightMargin", 0)
  rtbString.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
  rtbString.Locked = PropBag.ReadProperty("Locked", False)
  BuildCache
End Sub
Friend Sub BuildCache()
  
  
  Dim szStrings() As String
  Dim szValues() As String
  Dim iLoop As Integer
  Dim WordDisplay As WordStyle
  
  Set WordCache = New Collection
  
  szStrings = Split(m_WordList, ";")
  For iLoop = 0 To UBound(szStrings) - 1
    szValues = Split(szStrings(iLoop), "|")
    
    If szValues(1) = "1" Then 'Bold
      WordDisplay.bBold = True
    Else
      WordDisplay.bBold = False
    End If
       
    If szValues(2) = "1" Then ' Italic
      WordDisplay.bItalic = True
    Else
      WordDisplay.bItalic = False
    End If
    
    WordDisplay.szColour = szValues(3)
    
    WordCache.Add WordDisplay, szValues(0)
  Next iLoop
  
  
  

End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackColor", rtbString.BackColor, &H80000005)
  Call PropBag.WriteProperty("Enabled", rtbString.Enabled, True)
  Call PropBag.WriteProperty("Font", rtbString.Font, Ambient.Font)
  Call PropBag.WriteProperty("BorderStyle", rtbString.BorderStyle, 1)
  Call PropBag.WriteProperty("WordList", m_WordList, m_def_WordList)
  Call PropBag.WriteProperty("Text", rtbString.Text, "HBX")
  Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
  Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
  Call PropBag.WriteProperty("ScrollBars", rtbString.ScrollBars, 0)
  Call PropBag.WriteProperty("MultiLine", m_MultiLine, m_def_MultiLine)
  Call PropBag.WriteProperty("RightMargin", rtbString.RightMargin, 0)
  Call PropBag.WriteProperty("ToolTipText", rtbString.ToolTipText, "")
  Call PropBag.WriteProperty("Locked", rtbString.Locked, False)
End Sub
Public Property Get Wordlist() As String
Attribute Wordlist.VB_Description = "Pipe Sectioned, Semicolon Delimited array list of words"
Attribute Wordlist.VB_ProcData.VB_Invoke_Property = "Properties"
  Wordlist = m_WordList
End Property
Public Property Let Wordlist(ByVal New_WordList As String)
  m_WordList = New_WordList
  BuildCache
  PropertyChanged "WordList"
End Property
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in an object."
  Text = rtbString.Text
End Property
Public Property Let Text(ByVal New_Text As String)
  rtbString.Text() = New_Text
  PropertyChanged "Text"
  rtbstring_KeyUp 32, 0
End Property
Public Sub HUP()
Attribute HUP.VB_Description = "Causes the Control to refresh and search the string. "
ColourBox
End Sub
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
  ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As Long)
  m_ForeColor = New_ForeColor
  PropertyChanged "ForeColor"
End Property
Public Property Get BackStyle() As Boolean
Attribute BackStyle.VB_Description = "Returns/sets a value that indicating whether the selected object's verbs will be displayed in a popup menu when the right mouse button is clicked."
  BackStyle = m_BackStyle
End Property
Public Property Let BackStyle(ByVal New_BackStyle As Boolean)
  m_BackStyle = New_BackStyle
  PropertyChanged "BackStyle"
End Property

Public Property Get MultiLine() As Boolean
  MultiLine = m_MultiLine
End Property
Public Property Let MultiLine(ByVal New_MultiLine As Boolean)
Attribute MultiLine.VB_Description = "Returns/sets a value indicating whether a RichTextBox control can accept and display multiple lines of text."
  m_MultiLine = New_MultiLine
  PropertyChanged "MultiLine"
End Property
Public Property Get RightMargin() As Single
Attribute RightMargin.VB_Description = "Sets the right margin used for textwrap, centering, etc."
  RightMargin = rtbString.RightMargin
End Property
Public Property Let RightMargin(ByVal New_RightMargin As Single)
  rtbString.RightMargin() = New_RightMargin
  PropertyChanged "RightMargin"
End Property
Public Property Get ToolTipText() As String
  ToolTipText = rtbString.ToolTipText
End Property
Public Property Let ToolTipText(ByVal New_ToolTipText As String)
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
  rtbString.ToolTipText() = New_ToolTipText
  PropertyChanged "ToolTipText"
End Property
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets a value indicating whether the contents in a RichTextBox control can be edited."
  Locked = rtbString.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
  rtbString.Locked() = New_Locked
  PropertyChanged "Locked"
End Property

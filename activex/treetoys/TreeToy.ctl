VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.UserControl TreeToy 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "TreeToy.ctx":0000
   Begin MSComctlLib.TreeView trvMain 
      Height          =   3300
      Left            =   225
      TabIndex        =   0
      Top             =   135
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   5821
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "TreeToy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Public Enum TipType
    ttNone = 0
    ttTag
    ttText
    ttPath
    ttKey
End Enum

Dim cTT As New cls_TreeToy_cTreeTips

'Event Declarations:
Event AfterLabelEdit(Cancel As Integer, NewString As String)
Attribute AfterLabelEdit.VB_Description = "Occurs after a user edits the label of the currently selected Node or ListItem object."
Event BeforeLabelEdit(Cancel As Integer)
Attribute BeforeLabelEdit.VB_Description = "Occurs when a user attempts to edit the label of the currently selected ListItem or Node object."
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event Collapse(ByVal Node As Node)
Attribute Collapse.VB_Description = "Generated when any Node object in a TreeView control is collapsed."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
Event Expand(ByVal Node As Node)
Attribute Expand.VB_Description = "Occurs when a Node object in a TreeView control is expanded; that is, when its child nodes become visible."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus"
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event NodeCheck(ByVal Node As Node)
Attribute NodeCheck.VB_Description = "Occurs when Checkboxes = True and a Node object is checked/unchecked."
Event NodeClick(ByVal Node As Node)
Attribute NodeClick.VB_Description = "Occurs when a Node object is clicked."
Event Validate(Cancel As Boolean)
Attribute Validate.VB_Description = "Occurs when a control loses focus to a control that causes validation."

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  cTT.ShowIconsInNodeTips = PropBag.ReadProperty("ShowIconsInNodeTips", False)
  cTT.ShowIconsInScrollTips = PropBag.ReadProperty("ShowIconsInScrollTips", False)
  cTT.NodeTips = PropBag.ReadProperty("NodeTips", ttNone)
  cTT.ScrollTips = PropBag.ReadProperty("ScrollTips", ttNone)
  trvMain.Appearance = PropBag.ReadProperty("Appearance", 1)
  trvMain.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
  trvMain.CausesValidation = PropBag.ReadProperty("CausesValidation", True)
  trvMain.Checkboxes = PropBag.ReadProperty("Checkboxes", False)
  trvMain.Enabled = PropBag.ReadProperty("Enabled", True)
  trvMain.FullRowSelect = PropBag.ReadProperty("FullRowSelect", False)
  trvMain.HideSelection = PropBag.ReadProperty("HideSelection", True)
  trvMain.HotTracking = PropBag.ReadProperty("HotTracking", False)
  trvMain.Indentation = PropBag.ReadProperty("Indentation", 566.9291)
  trvMain.LabelEdit = PropBag.ReadProperty("LabelEdit", 0)
  trvMain.LineStyle = PropBag.ReadProperty("LineStyle", 0)
  trvMain.PathSeparator = PropBag.ReadProperty("PathSeparator", "\")
  trvMain.Scroll = PropBag.ReadProperty("Scroll", True)
  trvMain.SingleSel = PropBag.ReadProperty("SingleSel", False)
  trvMain.Sorted = PropBag.ReadProperty("Sorted", False)
  trvMain.Style = PropBag.ReadProperty("Style", 7)
  Set trvMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("ShowIconsInNodeTips", cTT.ShowIconsInNodeTips, False)
  Call PropBag.WriteProperty("ShowIconsInScrollTips", cTT.ShowIconsInScrollTips, False)
  Call PropBag.WriteProperty("NodeTips", cTT.NodeTips, ttNone)
  Call PropBag.WriteProperty("ScrollTips", cTT.ScrollTips, ttNone)
  Call PropBag.WriteProperty("Appearance", trvMain.Appearance, 1)
  Call PropBag.WriteProperty("BorderStyle", trvMain.BorderStyle, 0)
  Call PropBag.WriteProperty("CausesValidation", trvMain.CausesValidation, True)
  Call PropBag.WriteProperty("Checkboxes", trvMain.Checkboxes, False)
  Call PropBag.WriteProperty("Enabled", trvMain.Enabled, True)
  Call PropBag.WriteProperty("FullRowSelect", trvMain.FullRowSelect, False)
  Call PropBag.WriteProperty("HideSelection", trvMain.HideSelection, True)
  Call PropBag.WriteProperty("HotTracking", trvMain.HotTracking, False)
  Call PropBag.WriteProperty("Indentation", trvMain.Indentation, 566.9291)
  Call PropBag.WriteProperty("LabelEdit", trvMain.LabelEdit, 0)
  Call PropBag.WriteProperty("LineStyle", trvMain.LineStyle, 0)
  Call PropBag.WriteProperty("PathSeparator", trvMain.PathSeparator, "\")
  Call PropBag.WriteProperty("Scroll", trvMain.Scroll, True)
  Call PropBag.WriteProperty("SingleSel", trvMain.SingleSel, False)
  Call PropBag.WriteProperty("Sorted", trvMain.Sorted, False)
  Call PropBag.WriteProperty("Style", trvMain.Style, 7)
  Call PropBag.WriteProperty("Font", trvMain.Font, Ambient.Font)
End Sub

Private Sub UserControl_Initialize()
  Set cTT.Tree = trvMain
End Sub

Private Sub UserControl_Resize()
  trvMain.Top = 0
  trvMain.Left = 0
  trvMain.Width = UserControl.Width
  trvMain.Height = UserControl.Height
End Sub

'Methods
Public Sub FreezeCtl()
Attribute FreezeCtl.VB_Description = "Freezes the TreeView Control."
  iFreezeCtl trvMain
End Sub

Public Sub UnFreezeCtl()
Attribute UnFreezeCtl.VB_Description = "Unfreezes the Treeview Control."
  iUnFreezeCtl trvMain
End Sub

Public Function TreeCountChecked() As Integer
Attribute TreeCountChecked.VB_Description = "Returns the number of Nodes currently checked."
  TreeCountChecked = iTreeCountChecked(trvMain)
End Function

Public Function TreeGetRoot(Node As MSComctlLib.Node) As String
Attribute TreeGetRoot.VB_Description = "Returns the root Node key for the specified Node."
  TreeGetRoot = iTreeGetRoot(Node)
End Function

Public Sub TreeSelectiveCheck(Node As MSComctlLib.Node)
Attribute TreeSelectiveCheck.VB_Description = "Selectively check the specified node."
  iTreeSelectiveCheck Node
End Sub

Public Sub TreeSingleCheck(Node As MSComctlLib.Node)
Attribute TreeSingleCheck.VB_Description = "If checked itself, makes the specified node the only checked node."
  iTreeSingleCheck trvMain, Node
End Sub

Public Sub TreeSetChildren(Node As MSComctlLib.Node, Check As Boolean)
Attribute TreeSetChildren.VB_Description = "Recursively set Child Nodes."
  iTreeSetChildren Node, Check
End Sub

'Properties
Public Property Get ShowIconsInNodeTips() As Boolean
Attribute ShowIconsInNodeTips.VB_Description = "Sets/Returns whether icons will be displayed in Node tooltips."
  ShowIconsInNodeTips = cTT.ShowIconsInNodeTips
End Property

Public Property Let ShowIconsInNodeTips(ByVal New_ShowIconsInNodeTips As Boolean)
  cTT.ShowIconsInNodeTips = New_ShowIconsInNodeTips
  PropertyChanged "ShowIconsInNodeTips"
End Property

Public Property Get ShowIconsInScrollTips() As Boolean
Attribute ShowIconsInScrollTips.VB_Description = "Sets/Returns whether icons will be displayed in Scroll tooltips."
  ShowIconsInScrollTips = cTT.ShowIconsInScrollTips
End Property

Public Property Let ShowIconsInScrollTips(ByVal New_ShowIconsInScrollTips As Boolean)
  cTT.ShowIconsInScrollTips = New_ShowIconsInScrollTips
  PropertyChanged "ShowIconsInScrollTips"
End Property

Public Property Get NodeTips() As TipType
Attribute NodeTips.VB_Description = "Sets/Returns the source of the Node ToolTips."
  NodeTips = cTT.NodeTips
End Property

Public Property Let NodeTips(ByVal New_NodeTips As TipType)
  cTT.NodeTips = New_NodeTips
  PropertyChanged "NodeTips"
End Property

Public Property Get ScrollTips() As TipType
Attribute ScrollTips.VB_Description = "Sets/Returns the source of the Scroll ToolTips."
  ScrollTips = cTT.ScrollTips
End Property

Public Property Let ScrollTips(ByVal New_ScrollTips As TipType)
  cTT.ScrollTips = New_ScrollTips
  PropertyChanged "ScrollTips"
End Property

'
'Mapped properties/events are below!
'

Public Property Get SelectedItem() As MSComctlLib.Node
Attribute SelectedItem.VB_Description = "Returns the currently selected Node."
  Set SelectedItem = trvMain.SelectedItem
End Property

Public Property Get Nodes() As MSComctlLib.Nodes
Attribute Nodes.VB_Description = "Returns a reference to a collection of Node objects."
  Set Nodes = trvMain.Nodes
End Property

Public Property Get ImageList() As Object
Attribute ImageList.VB_Description = "Sets/Returns the ImageList Control used by the TreeView."
  Set ImageList = trvMain.ImageList
End Property

Public Property Set ImageList(ByVal New_Imagelist As Object)
  Set trvMain.ImageList = New_Imagelist
End Property

Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not controls, Forms or an MDIForm are painted at run time with 3-D effects."
  Appearance = trvMain.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
  trvMain.Appearance() = New_Appearance
  PropertyChanged "Appearance"
End Property

Public Property Get BorderStyle() As MSComctlLib.BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = trvMain.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As MSComctlLib.BorderStyleConstants)
  trvMain.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

Public Property Get CausesValidation() As Boolean
Attribute CausesValidation.VB_Description = "Returns/sets whether validation occurs on the control which lost focus."
  CausesValidation = trvMain.CausesValidation
End Property

Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
  trvMain.CausesValidation() = New_CausesValidation
  PropertyChanged "CausesValidation"
End Property

Public Property Get Checkboxes() As Boolean
Attribute Checkboxes.VB_Description = "Returns/sets a value which determines if the control displays a checkbox next to each item in the tree."
  Checkboxes = trvMain.Checkboxes
End Property

Public Property Let Checkboxes(ByVal New_Checkboxes As Boolean)
  trvMain.Checkboxes() = New_Checkboxes
  PropertyChanged "Checkboxes"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = trvMain.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  trvMain.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_Description = "Returns/sets a value which determines if the entire row of the selected item is highlighted and clicking anywhere on an item's row causes it to be selected."
  FullRowSelect = trvMain.FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
  trvMain.FullRowSelect() = New_FullRowSelect
  PropertyChanged "FullRowSelect"
End Property

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Determines whether the selected item will display as selected when the TreeView loses focus"
  HideSelection = trvMain.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
  trvMain.HideSelection() = New_HideSelection
  PropertyChanged "HideSelection"
End Property

Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Returns/sets a value which determines if items are highlighted as the mousepointer passes over them."
  HotTracking = trvMain.HotTracking
End Property

Public Property Let HotTracking(ByVal New_HotTracking As Boolean)
  trvMain.HotTracking() = New_HotTracking
  PropertyChanged "HotTracking"
End Property

Public Property Get Indentation() As Single
Attribute Indentation.VB_Description = "Returns/sets the width of the indentation for a TreeView control."
  Indentation = trvMain.Indentation
End Property

Public Property Let Indentation(ByVal New_Indentation As Single)
  trvMain.Indentation() = New_Indentation
  PropertyChanged "Indentation"
End Property

Public Property Get LabelEdit() As LabelEditConstants
Attribute LabelEdit.VB_Description = "Returns/sets a value that determines if a user can edit the label of a ListItem or Node object."
  LabelEdit = trvMain.LabelEdit
End Property

Public Property Let LabelEdit(ByVal New_LabelEdit As LabelEditConstants)
  trvMain.LabelEdit() = New_LabelEdit
  PropertyChanged "LabelEdit"
End Property

Public Property Get LineStyle() As TreeLineStyleConstants
Attribute LineStyle.VB_Description = "Returns/sets the style of lines displayed between Node objects."
  LineStyle = trvMain.LineStyle
End Property

Public Property Let LineStyle(ByVal New_LineStyle As TreeLineStyleConstants)
  trvMain.LineStyle() = New_LineStyle
  PropertyChanged "LineStyle"
End Property

Public Property Get PathSeparator() As String
Attribute PathSeparator.VB_Description = "Returns/sets the delimiter string used for the path returned by the FullPath property."
  PathSeparator = trvMain.PathSeparator
End Property

Public Property Let PathSeparator(ByVal New_PathSeparator As String)
  trvMain.PathSeparator() = New_PathSeparator
  PropertyChanged "PathSeparator"
End Property

Public Property Get Scroll() As Boolean
Attribute Scroll.VB_Description = "Returns/sets a value which determines if the TreeView displays scrollbars and allows scrolling (vertical and horizontal)."
  Scroll = trvMain.Scroll
End Property

Public Property Let Scroll(ByVal New_Scroll As Boolean)
  trvMain.Scroll() = New_Scroll
  PropertyChanged "Scroll"
End Property

Public Property Get SingleSel() As Boolean
Attribute SingleSel.VB_Description = "Returns/sets a value which determines if selecting a new item in the tree expands that item and collapses the previously selected item."
  SingleSel = trvMain.SingleSel
End Property

Public Property Let SingleSel(ByVal New_SingleSel As Boolean)
  trvMain.SingleSel() = New_SingleSel
  PropertyChanged "SingleSel"
End Property

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
  Sorted = trvMain.Sorted
End Property

Public Property Let Sorted(ByVal New_Sorted As Boolean)
  trvMain.Sorted() = New_Sorted
  PropertyChanged "Sorted"
End Property

Public Property Get Style() As TreeStyleConstants
Attribute Style.VB_Description = "Displays a hierarchical list of Node objects, each of which consists of a label and an optional bitmap."
  Style = trvMain.Style
End Property

Public Property Let Style(ByVal New_Style As TreeStyleConstants)
  trvMain.Style() = New_Style
  PropertyChanged "Style"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = trvMain.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set trvMain.Font = New_Font
  PropertyChanged "Font"
End Property

Private Sub trvMain_AfterLabelEdit(Cancel As Integer, NewString As String)
  RaiseEvent AfterLabelEdit(Cancel, NewString)
End Sub

Private Sub trvMain_BeforeLabelEdit(Cancel As Integer)
  RaiseEvent BeforeLabelEdit(Cancel)
End Sub

Private Sub trvMain_Click()
  RaiseEvent Click
End Sub

Private Sub trvMain_Collapse(ByVal Node As Node)
  RaiseEvent Collapse(Node)
End Sub

Private Sub trvMain_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub trvMain_Expand(ByVal Node As Node)
  RaiseEvent Expand(Node)
End Sub

Private Sub trvMain_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub trvMain_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub trvMain_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub trvMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub trvMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub trvMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub trvMain_NodeCheck(ByVal Node As Node)
  RaiseEvent NodeCheck(Node)
End Sub

Private Sub trvMain_NodeClick(ByVal Node As Node)
  RaiseEvent NodeClick(Node)
End Sub

Private Sub trvMain_Validate(Cancel As Boolean)
  RaiseEvent Validate(Cancel)
End Sub

Public Function GetVisibleCount() As Long
Attribute GetVisibleCount.VB_Description = "Returns the number of Node objects that fit in the internal area of a TreeView control."
  GetVisibleCount = trvMain.GetVisibleCount
End Function

Public Function HitTest(x As Single, y As Single) As Node
Attribute HitTest.VB_Description = "Returns a reference to the ListItem object or Node object located at the coordinates of x and y. Used with drag and drop operations."
  Set HitTest = trvMain.HitTest(x, y)
End Function

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a form or control."
  trvMain.Refresh
End Sub

Public Sub StartLabelEdit()
Attribute StartLabelEdit.VB_Description = "Begins a label editing operation on a ListItem or Node object."
  trvMain.StartLabelEdit
End Sub

Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
  trvMain.OLEDrag
End Sub


VERSION 5.00
Begin VB.Form frmToolTip 
   Appearance      =   0  'Flat
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   225
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   225
   ScaleWidth      =   1455
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tooltip"
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   15
      Width           =   480
   End
End
Attribute VB_Name = "frmToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

'-------------------------------------------------------------------------------
'APIs for finding the cursor dimensions:

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long

Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'-------------------------------------------------------------------------------
'APIs for showing the window without activating it:

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOWNOACTIVATE = 4

'-------------------------------------------------------------------------------
'APIs for making the window a top-most window:

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10

'-------------------------------------------------------------------------------
'Variables to store Twip measurements:

Private mTPPX As Single
Private mTPPY As Single

Private mnMarginWidth As Single
Private mnMarginHeight As Single
Public Sub ShowToolTip(Text As String, Optional Image As StdPicture, Optional MarginWidth As Long = 2, Optional MarginHeight As Long = 2)
    
    'Property cache
    mTPPX = Screen.TwipsPerPixelX
    mTPPY = Screen.TwipsPerPixelY

    'Calculate the margin values
    mnMarginWidth = MarginWidth * mTPPX
    mnMarginHeight = MarginHeight * mTPPY

    'Set the image, if any.
    SetImage Image, True
    
    'Set the text.
    SetText Text, True
    
    'Size and position the controls and the form.
    SetSize
    SetPosition
    
    'Show the form, but don't activate it.
    ShowWindow hWnd, SW_SHOWNOACTIVATE
    
End Sub

Public Sub SetImage(Image As StdPicture, Optional bNoResize As Boolean)

Dim bUseImage As Boolean

    'The image may be Nothing ...
    If Not Image Is Nothing Then
        '... or it may be empty
        If Image.Type <> vbPicTypeNone Then
            bUseImage = True
        End If
    End If

    If bUseImage Then
        'Set the image. This will automatically size the Image control, because
        '   its Stretch property is set to False.
        Set Image1.Picture = Image
        'Show the image
        Image1.Visible = True
    Else
        'Hide the image
        Image1.Visible = False
    End If
    
    If Not bNoResize Then
        SetSize
    End If
    
End Sub

Public Sub SetText(Text As String, Optional bNoResize As Boolean)
    
    'Label1.AutoSize = True
    Label1.Caption = Text
    
    If Not bNoResize Then
        SetSize
    End If
    
End Sub

Private Sub SetSize()

    'Set the position of the controls and the height of the form.
    If Image1.Visible Then
        'Position the controls horizontally.

        'Put a margin to the left of the image.
        Image1.Left = mnMarginWidth
        'Put a margin between the image and the text.
        Label1.Left = Image1.Width + (2 * mnMarginWidth)

        'Position the controls vertically and set the form's height.
        If Image1.Height > Label1.Height Then
            'Put a margin above the image
            Image1.Top = mnMarginHeight
            'Put a margin below the image, and allow room for the form's border.
            Height = Image1.Height + (2 * mnMarginHeight) + BorderHeight
            'Put the text in the middle of the remaining space.
            Label1.Top = (ScaleHeight - Label1.Height) / 2
        Else
            'Put a margin above the text
            Label1.Top = mnMarginHeight
            'Put a margin below the text, and allow room for the form's border.
            Height = Label1.Height + (2 * mnMarginHeight) + BorderHeight
            'Put the image in the middle of the remaining space.
            Image1.Top = (ScaleHeight - Label1.Height) / 2
        End If
    Else
        'Put a margin to the left of and above the label.
        Label1.Move mnMarginWidth, mnMarginWidth
        'Put a margin below the label, and allow room for the form's border.
        Height = Label1.Height + (2 * mnMarginHeight) + BorderHeight
    End If

    'Set the width of the form, and allow room for the form's border.
    Width = Label1.Left + Label1.Width + mnMarginWidth + BorderWidth

End Sub

Private Sub SetPosition()

Dim nCursorLeft As Single
Dim nCursorTop As Single
Dim nCursorBottom As Single

    
    'Get the interesting cursor dimensions. GetCursorDimensions gets six bits of
    '   information, but we only care about a few, so we use the named argument
    '   syntax to specify only the ones we care about.
    GetCursorDimensions Left:=nCursorLeft, Top:=nCursorTop, Bottom:=nCursorBottom

    'Position the form horizontally
    If nCursorLeft + Width <= Screen.Width Then
        'Line up the form with the mouse pointer, if it will fit.
        Left = nCursorLeft
    Else
        'If it won't fit, then put it as far right as possible.
        Left = Screen.Width - Width
    End If

    'Position the form vertically
    If nCursorBottom + Height <= Screen.Height Then
        'Put the form under the cursor, if it will fit.
        Top = nCursorTop + 300 'nCursorBottom
    Else
        'If it won't fit, then put it above the cursor.
        Top = nCursorTop - Height
    End If
    
End Sub


Private Property Get BorderWidth() As Single
    
    'Find out how much space is needed for the form's border.
    BorderWidth = Width - ScaleWidth
    
End Property
Private Property Get BorderHeight() As Single
    
    BorderHeight = Height - ScaleHeight
    
End Property

Private Sub GetCursorDimensions(Optional PointerX As Single, Optional PointerY As Single, Optional Left As Single, Optional Top As Single, Optional Right As Single, Optional Bottom As Single)

Dim ptCursor As POINTAPI
Dim hCursor As Long
Dim udtIconInfo As ICONINFO
Dim nMultiplier As Single
Dim udtBitmapInfo As BITMAPINFO

    'Find the pointer position
    If GetCursorPos(ptCursor) = 0 Then
        Err.Raise 5, , "GetCursorPos failed."
    End If

    'Get a handle to the current cursor
    hCursor = GetCursor
    
    'Get the icon information for the current cursor
    If GetIconInfo(hCursor, udtIconInfo) = 0 Then
        Err.Raise 5, , "GetIconInfo failed."
    End If
    
    If udtIconInfo.hbmMask = 0 Then
        Err.Raise 5, , "GetIconInfo returned an invalid hbmMask."
    End If
    
    'If the hbmColor member is zero, then this is a black and white cursor.
    If udtIconInfo.hbmColor = 0 Then
        'If this is a black and white cursor, then the bitmap is actually twice
        '   the height of the cursor.
        nMultiplier = 0.5
    Else
        'If this is a color bitmap, then the bitmap is the actual height of the
        '   cursor
        nMultiplier = 1
        'Release the color bitmap created by GetIconInfo.
        DeleteObject udtIconInfo.hbmColor
    End If
    
    'Initialize the biSize member so that Windows knows what kind of structure
    '   it has to fill.
    udtBitmapInfo.bmiHeader.biSize = Len(udtBitmapInfo.bmiHeader)
    
    'Get the bitmap information
    If GetDIBits(hDC, udtIconInfo.hbmMask, 0, 0, ByVal 0, udtBitmapInfo, 0) = 0 Then
        Err.Raise 5, , "GetDIBits failed."
    End If
    
    'Release the mask bitmap created by GetIconInfo
    DeleteObject udtIconInfo.hbmMask
    
    'Adjust the height for a black and white cursor, if necessary (see above)
    udtBitmapInfo.bmiHeader.biHeight = udtBitmapInfo.bmiHeader.biHeight * nMultiplier

    'Calculate the actual return values
    With ptCursor
        'Simply convert the pointer position to twips
        PointerX = .x * mTPPX
        PointerY = .y * mTPPY
        
        'Back up from pointer position to beginning edge of cursor, then convert to
        '   twips
        Left = (.x - udtIconInfo.xHotspot) * mTPPX
        Top = (.y - udtIconInfo.yHotspot) * mTPPY
    End With
    
    'Find extent of cursor in twips. Add to beginning edge of cursor, already in
    '   twips.
    Right = (udtBitmapInfo.bmiHeader.biWidth * mTPPX) + Left
    Bottom = (udtBitmapInfo.bmiHeader.biHeight * mTPPY) + Top
    
End Sub

Private Sub Form_Load()
    
    'Make this a topmost window, as is appropriate for a tooltip
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    
End Sub





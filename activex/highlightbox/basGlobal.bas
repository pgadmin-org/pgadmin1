Attribute VB_Name = "basFunctions"
Option Explicit
Public Type T_RGB
  R As Integer
  G As Integer
  B As Integer
End Type

Public MinimisedHeight As Integer
Public MinimisedWidth As Integer
Public LastTop As Integer
Public Lastleft As Integer
Public MaximisedState As Boolean
Public WordCache As Collection
Public Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long
Public Function get_RGB(LColour As Long) As T_RGB
Dim szHEX As String
szHEX = Hex(LColour)
While Len(szHEX) < 6
  szHEX = "0" & szHEX
Wend
get_RGB.R = CInt("&H" & Mid(szHEX, 5, 2))
get_RGB.G = CInt("&H" & Mid(szHEX, 3, 2))
get_RGB.B = CInt("&H" & Mid(szHEX, 1, 2))
End Function

Public Function CharInstr(szChar As String, szString As String) As Boolean

Dim lX As Long
For lX = 1 To Len(szString)
  If Mid(szString, lX, 1) = szChar Then
    CharInstr = True
    Exit For
  End If
Next lX

End Function
Public Function SearchCache(szLookup As String) As String
Dim szSearchcache As String
On Error Resume Next

  szSearchcache = WordCache(szLookup).szRTFstring
  If szSearchcache = "" Then szSearchcache = szLookup
  SearchCache = szSearchcache
  
End Function

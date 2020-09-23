Attribute VB_Name = "Module1"
Option Explicit

Declare Function GetStretchBltMode Lib "gdi32" (ByVal hdc As Long) As Long

Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long

'Public Const COLORONCOLOR = 3
Public Const HALFTONE = 4

' Use:-
'   If SetStretchBltMode(Picture1.hdc, COLORONCOLOR) = 0 Then
'      MsgBox "SetStretchBltMode error ", vbCritical, " "
'      End
'   End If

'----------------------------------------------------------------
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
ByVal dwRop As Long) As Long

'----------------------------------------------------------------
'To fill BITMAP structure
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" _
(ByVal hObject As Long, ByVal Lenbmp As Long, dimbmp As Any) As Long

Public Type BITMAP
   bmType As Long              ' Type of bitmap
   bmWidth As Long             ' Pixel width
   bmHeight As Long            ' Pixel height
   bmWidthBytes As Long        ' Byte width = 3 x Pixel width
   bmPlanes As Integer         ' Color depth of bitmap
   bmBitsPixel As Integer      ' Bits per pixel, must be 16 or 24
   bmBits As Long              ' This is the pointer to the bitmap data  !!!
End Type
Public bmp As BITMAP

Public FileSpec$()
Public NumPicsSelected As Long


Public Sub Extract_Store_FileSpecs(ByVal AFileSpec$, ByVal StartPicNum As Long)
'Public FileSpec$(), NumPicsSelected

' AFileSpec$ = multi-selected filespec string
' eg "C:\Program Files\Common\File1.bmp"  ' One file
' eg "C:\Program Files\Common|File1.bmp|File2.jpg"  ' Two files, | = Null char

Dim pNull1 As Long  ' Instr pointers
Dim pNull2 As Long
Dim APath$
Dim FName$

   NumPicsSelected = StartPicNum
   
   pNull2 = 1
   pNull1 = InStr(pNull2, AFileSpec$, vbNullChar)
   If pNull1 = 0 Then
      NumPicsSelected = NumPicsSelected + 1
      ReDim FileSpec$(NumPicsSelected)
      FileSpec$(NumPicsSelected) = AFileSpec$
   Else  ' pNull1<>0
      
      APath$ = Left$(AFileSpec$, pNull1 - 1) & "\"
      pNull1 = pNull1 + 1
      
      Do
         pNull2 = InStr(pNull1, AFileSpec$, vbNullChar)
         If pNull2 = 0 Then
            NumPicsSelected = NumPicsSelected + 1
            ReDim Preserve FileSpec$(NumPicsSelected)
            FName$ = Mid$(AFileSpec$, pNull1, Len(AFileSpec$) - pNull1 + 1)
            FileSpec$(NumPicsSelected) = FName$
            Exit Do
         Else
            NumPicsSelected = NumPicsSelected + 1
            ReDim Preserve FileSpec$(NumPicsSelected)
            FName$ = Mid$(AFileSpec$, pNull1, pNull2 - pNull1)
            FileSpec$(NumPicsSelected) = FName$
         End If
         pNull1 = pNull2 + 1
      Loop
   
   End If

End Sub


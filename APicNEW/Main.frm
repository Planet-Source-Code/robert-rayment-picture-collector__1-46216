VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7515
   DrawWidth       =   2
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   501
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   10560
      ScaleHeight     =   2955
      ScaleWidth      =   1230
      TabIndex        =   1
      Top             =   225
      Width           =   1230
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      Height          =   3315
      Left            =   60
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   0
      Top             =   360
      Width           =   2460
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save As"
      End
      Begin VB.Menu zbrk2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Picture Collector by Robert Rayment June 2003

' Updated: with New look dialog (Thanks to RedBird77)
'          aspect ratio maintained

Option Explicit

Option Base 1

Dim CommonDialog1 As New OSDialog

Dim PathSpec$
Dim i As Long  ' Gen loop counter
Dim A$         ' Gen String


Private Sub Form_Load()

PathSpec$ = App.Path
If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"

Me.Caption = "Picture Collector by Robert Rayment"
Me.Show

End Sub

Private Sub SHOWPICS()

On Error Resume Next

Dim NumRows As Long
Dim NumCols As Long
Dim NW As Long
Dim NH As Long
Dim PW As Long
Dim PH As Long
Dim py As Long
Dim px As Long
Dim W As Long
Dim H As Long
Dim row As Long
Dim col As Long
Dim res As Long
Dim p As Long
Dim pdot As Long

Dim OldMode As Long


NumRows = NumPicsSelected \ 10 + 1  ' sometimes 1 too many
' 64x64  68x78
NW = 64        ' 64
NH = 64        ' 64
PW = NW + 4    ' 68
PH = NH + 14   ' 78
NumCols = 10

Picture1.Width = 10 * PW + 12
Picture1.Height = NumRows * PH + 20

OldMode = GetStretchBltMode(Picture1.hdc) ' OldMode =1 Win98

If SetStretchBltMode(Picture1.hdc, HALFTONE) = 0 Then
   MsgBox "SetStretchBltMode error ", vbCritical, " "
   End
End If
' NB HALFTONE makes no difference in Win98 but better & slower
' in WinXP
   
   i = 1
   For row = 1 To NumRows
      py = 2 + (row - 1) * PH
   For col = 1 To NumCols
   
      NW = 64        ' 64
      NH = 64        ' 64
      
      px = 2 + (col - 1) * PW
      
      Picture2.Picture = LoadPicture
      Picture2.Picture = LoadPicture(FileSpec$(i))
      Picture2.Refresh
      
      res = GetObjectAPI(Picture2.Image, Len(bmp), bmp)
      W = bmp.bmWidth
      H = bmp.bmHeight
      
   '      ' Sometimes NEEDS BlackNess first !!!
   '   ' Actually almost anything other than vbSrcCopy
   '   If StretchBlt(Picture1.hdc, px, py, _
   '      NW, NH, Picture2.hdc, _
   '      0&, 0&, W, H, &H42) = 0 Then
   '      MsgBox "StretchBlt error , blacken", vbCritical, " "
   '      End
   '   End If
   '   Picture1.Refresh
      
      ' Maintain aspect ratio
      If W >= H Then
         NH = NH * (H / W)
      Else
         NW = NW * (W / H)
      End If
      
      If StretchBlt(Picture1.hdc, px, py, _
         NW, NH, Picture2.hdc, _
         0&, 0&, W, H, vbSrcCopy) = 0 Then
         MsgBox "StretchBlt error ", vbCritical, " "
         End
      End If
      Picture1.Refresh
      
      Picture1.CurrentY = py + NH + 1
      Picture1.CurrentX = px
      A$ = FileSpec$(i)
      p = InStrRev(A$, "\")
      If p = 0 Then
         pdot = InStr(1, A$, ".")
         A$ = Left$(A$, pdot - 1)
      Else
         A$ = Right$(A$, Len(A$) - p)
         pdot = InStr(1, A$, ".")
         A$ = Left$(A$, pdot - 1)
      End If
         
      Picture1.Print A$
      
      i = i + 1
      If i > NumPicsSelected Then Exit For
   Next col
   If i > NumPicsSelected Then Exit For
   Next row

SetStretchBltMode Picture1.hdc, OldMode

End Sub


Private Sub mnuOpen_Click()
Dim bpp As Integer
Dim Title$, Filt$, InDir$
Dim AFileSpec$
Dim StartNumPicsSelected

   ' LOAD STANDARD VB PICTURES INTO picDisplay ONLY
   
   MousePointer = vbDefault
   
   Title$ = "Load picture(s) Multi-select"
   Filt$ = "Pics bmp,jpg,gif,wmf,emf|*.bmp;*.jpg;*.gif;*.wmf;*.emf"
   InDir$ = AFileSpec$
   
   CommonDialog1.ShowOpen AFileSpec$, Title$, Filt$, InDir$, "", Me.hWnd, True  ', True
   
   If Len(AFileSpec$) = 0 Then
      MsgBox " No files selected." & vbCr & " or too many files" & vbCr & " for string length.", vbInformation, " "
      Close
      NumPicsSelected = 0
      Exit Sub
   End If
   
   Cls
   Print AFileSpec$
   Me.Refresh
   
   Picture1.Cls
   
   StartNumPicsSelected = 0
   Extract_Store_FileSpecs AFileSpec$, StartNumPicsSelected
   
   Cls
   Print " NumPicsSelected =" & Str$(NumPicsSelected)
   
   SHOWPICS

End Sub

Private Sub mnuSave_Click()
Dim Title$, Filt$, InDir$
Dim AFileSpec$

   '  SAVE 24bpp BMP
   AFileSpec$ = ""
   Title$ = "Save stored pic as a bmp"
   Filt$ = "Save bmp|*.bmp"
   InDir$ = AFileSpec$
   
   CommonDialog1.ShowSave AFileSpec$, Title$, Filt$, InDir$, "", Me.hWnd
   
   If Len(AFileSpec$) = 0 Then
      Close
      Exit Sub
   End If
   
   SavePicture Picture1.Image, AFileSpec$

End Sub
Private Sub Form_Unload(Cancel As Integer)
Set CommonDialog1 = Nothing
Unload Me
End
End Sub

Private Sub mnuExit_Click()
Set CommonDialog1 = Nothing
Unload Me
End
End Sub


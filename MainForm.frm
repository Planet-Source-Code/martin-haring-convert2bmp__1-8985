VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MainForm 
   AutoRedraw      =   -1  'True
   Caption         =   "Convert2Bmp"
   ClientHeight    =   4035
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlg1 
      Left            =   4920
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3195
      Left            =   105
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   3135
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu fOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu Saveas 
         Caption         =   "&Save As..."
      End
      Begin VB.Menu CloseMe 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "&Help"
      Begin VB.Menu AboutMe 
         Caption         =   "&About Convert2Bmp..."
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AboutMe_Click()
frmAbout.Show 1


End Sub

Private Sub CloseMe_Click()
 Unload Me
End Sub

Private Sub fOpen_Click()
 cdlg1.CancelError = True
On Error GoTo ErrHandler
cdlg1.FileName = ""
cdlg1.DialogTitle = "Open Image File"
cdlg1.Filter = "Image File(*.bmp,*.ico,*.jpg,*.gif,*.cur)|*.bmp;*.ico;*.jpg;*.gif;*.cur|All Files (*.*)|*.*"
cdlg1.FilterIndex = 1



cdlg1.ShowOpen
Set Picture1.Picture = LoadPicture(cdlg1.FileName)

 Exit Sub
ErrHandler:
  Exit Sub
End Sub

Private Sub Form_Resize()
Dim w As Integer, h As Integer
w = Me.width
h = Me.height
Picture1.Move 0, 0, w, h

End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim vFN
 For Each vFN In Data.Files
  Set Picture1.Picture = LoadPicture(vFN)
 Next vFN
 End If
End Sub

Private Sub Saveas_Click()
 cdlg1.CancelError = True
On Error GoTo ErrHandler
cdlg1.FileName = ""

cdlg1.DialogTitle = "Save To Bitmap"
cdlg1.Filter = "BitMap File|*.bmp"
cdlg1.DefaultExt = ".bmp"
cdlg1.ShowSave
SavePicture Picture1.Image, cdlg1.FileName

 Exit Sub
ErrHandler:
End Sub

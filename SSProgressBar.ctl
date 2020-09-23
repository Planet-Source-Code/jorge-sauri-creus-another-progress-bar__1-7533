VERSION 5.00
Begin VB.UserControl SSProgressBar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BackStyle       =   0  'Transparent
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   ClipControls    =   0   'False
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   735
   ScaleWidth      =   4575
   ToolboxBitmap   =   "SSProgressBar.ctx":0000
   Begin VB.PictureBox imgSkin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      Picture         =   "SSProgressBar.ctx":0312
      ScaleHeight     =   705
      ScaleWidth      =   4545
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.Image imgProgress 
         Height          =   375
         Index           =   1
         Left            =   120
         Picture         =   "SSProgressBar.ctx":4425
         Stretch         =   -1  'True
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblPercent 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4335
      End
   End
End
Attribute VB_Name = "SSProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public ImageType As Integer
Public Max As Integer
Private Value As Integer
Public Border As Boolean

Private Sub UserControl_Initialize()
Dim i As Integer
Dim inc As Byte

Border = False

If Border = True Then inc = 10 Else inc = 0

For i = 2 To 10
    Load imgProgress(i)
    If Border = True Then
        imgProgress(i).BorderStyle = 1
    Else
        imgProgress(i).BorderStyle = 0
    End If
    imgProgress(i).Visible = True
    imgProgress(i).Picture = imgProgress(1).Picture
    imgProgress(i).Left = imgProgress(i - 1).Left + imgProgress(i - 1).Width + inc
    imgProgress(i).Top = imgProgress(1).Top
Next i

'imgskin.Width = imgskin.Left + imgProgress(9).Left + imgProgress(9).Width + 100


Max = 100
Value = 0

End Sub


Public Sub SetValue(ByRef Valor As Integer)
Dim auxCont As Integer
Dim auxPercent As Integer

Dim i As Integer

If Valor > Max Then Valor = Max

auxCont = Int((10 * Valor) / Max)

If auxCont > 10 Then auxCont = 10

For i = 1 To auxCont
    imgProgress(i).Picture = frmSprites.imgBars(ImageType).Picture
Next i

auxPercent = Int((100 * Valor) / Max)

lblPercent.Caption = Str(auxPercent) + "%"

End Sub



Public Sub LoadPic(file As String)
Const EXTERNAL = 7


frmSprites.imgBars(EXTERNAL).Picture = LoadPicture(file)
ImageType = EXTERNAL

End Sub



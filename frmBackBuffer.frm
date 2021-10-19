VERSION 5.00
Begin VB.Form frmBackBuffer 
   AutoRedraw      =   -1  'True
   Caption         =   "Back buffering"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdAutoRedraw 
      Caption         =   "Draw with Autoredraw"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdDrawAuto 
      Caption         =   "Draw with backbuffer"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   3000
      Picture         =   "frmBackBuffer.frx":0000
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   6
      Top             =   3960
      Width           =   3060
   End
   Begin VB.PictureBox picSprite1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1560
      Picture         =   "frmBackBuffer.frx":1D502
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   4560
      Width           =   540
   End
   Begin VB.PictureBox picSprite2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2280
      Picture         =   "frmBackBuffer.frx":1E144
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   4560
      Width           =   540
   End
   Begin VB.PictureBox picMask1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   1560
      Picture         =   "frmBackBuffer.frx":1ED86
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   3960
      Width           =   540
   End
   Begin VB.PictureBox picMask2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2280
      Picture         =   "frmBackBuffer.frx":1F9C8
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   3960
      Width           =   540
   End
   Begin VB.PictureBox picFrontBuffer 
      Height          =   3135
      Left            =   120
      ScaleHeight     =   3075
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.PictureBox picBackBuffer 
      Height          =   3135
      Left            =   3360
      ScaleHeight     =   205
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1200
      TabIndex        =   13
      Top             =   5520
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Time elapsed:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   5520
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Front buffer"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   960
      TabIndex        =   9
      Top             =   3480
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Background Sprite"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3720
      TabIndex        =   8
      Top             =   7080
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Back buffer"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3840
      TabIndex        =   7
      Top             =   3480
      Width           =   825
   End
End
Attribute VB_Name = "frmBackBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Chapter 1
'Backbuffer vs Refresh method test
'


Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim BackGroundWidth As Long
Dim Backgroundheight As Long
Dim SpriteWidth As Long
Dim Spriteheight As Long

Dim Sprite1Y As Long
Dim Sprite1X As Long
Dim Sprite2Y As Long
Dim Sprite2X As Long
Dim bRefresh As Boolean
Dim Running As Boolean

Private Sub cmdAutoRedraw_Click()

bRefresh = True
picBackBuffer.AutoRedraw = True
RunMain

End Sub

Private Sub cmdDrawAuto_Click()
picFrontBuffer.AutoRedraw = False
picBackBuffer.AutoRedraw = False

bRefresh = False

RunMain


End Sub
Private Sub RunMain()
Dim Start1 As Long, Finish As Long
Dim i As Integer

Start1 = Timer

Do

    'Update movement variables
    Sprite1Y = Sprite1Y + 2
    Sprite1X = Sprite1X + 2
    
    Sprite2X = Sprite2X + 1
    Sprite2Y = Sprite2Y + 1
    
    'Keep the sprites inside the picture box
    If Sprite1Y > BackGroundWidth Then
        Sprite1Y = 0
    End If
    
    If Sprite1X > BackGroundWidth Then
        Sprite1X = 0
    End If
        
    If Sprite2Y > BackGroundWidth Then
        Sprite2Y = 0
    End If
    
    If Sprite2X > BackGroundWidth Then
        Sprite2X = 0
    End If
    '------------------
    
    'Draw the things
    DrawBackGround
    
    DrawMasks
    
    DrawSprites
    
    If Not bRefresh Then
    
        DrawToFrontBuffer
    
    End If
    
    
    If i = 1000 Then
           
        Finish = Timer
        lblTime.Caption = Finish - Start1
        i = 0
        Start1 = Timer
    Else
        i = i + 1
    End If
    
    
    DoEvents
    
Loop While Running

Unload Me

End Sub

Private Sub cmdExit_Click()

Running = False

End Sub

Private Sub Form_Load()


picBackBuffer.Move picBackBuffer.Left, picBackBuffer.Top, picBackground.Width, picBackground.Height
picFrontBuffer.Move picFrontBuffer.Left, picFrontBuffer.Top, picBackground.Width, picBackground.Height

'Set the dimensions
BackGroundWidth = picBackground.ScaleWidth
Backgroundheight = picBackground.ScaleHeight
SpriteWidth = picSprite1.ScaleWidth
Spriteheight = picSprite1.ScaleHeight

Sprite1Y = 10
Sprite1X = 10
Sprite2X = 10
Sprite2Y = 50

Running = True

End Sub


'Draws the background picture to the back buffer
Private Sub DrawBackGround()

'Draw background to back buffer
BitBlt picBackBuffer.hDC, 0, 0, BackGroundWidth, Backgroundheight, picBackground.hDC, 0, 0, vbSrcCopy

End Sub
'Draws the masks
Private Sub DrawMasks()
'Draw the masks
BitBlt picBackBuffer.hDC, Sprite1X, Sprite1Y, SpriteWidth, Spriteheight, picMask1.hDC, 0, 0, vbSrcAnd
BitBlt picBackBuffer.hDC, Sprite2X, Sprite2Y, SpriteWidth, Spriteheight, picMask2.hDC, 0, 0, vbSrcAnd

End Sub

'Draws the sprites
Private Sub DrawSprites()

BitBlt picBackBuffer.hDC, Sprite1X, Sprite1Y, SpriteWidth, Spriteheight, picSprite1.hDC, 0, 0, vbSrcPaint
BitBlt picBackBuffer.hDC, Sprite2X, Sprite2Y, SpriteWidth, Spriteheight, picSprite2.hDC, 0, 0, vbSrcPaint

If bRefresh Then
    picBackBuffer.Refresh
End If

End Sub
'Draws the back buffer to the front buffer
Private Sub DrawToFrontBuffer()

BitBlt picFrontBuffer.hDC, 0, 0, BackGroundWidth, Backgroundheight, picBackBuffer.hDC, 0, 0, vbSrcCopy

End Sub

Private Sub Form_Unload(Cancel As Integer)

'flag the end
Running = False

End Sub

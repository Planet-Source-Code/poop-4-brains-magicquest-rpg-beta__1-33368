VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Magic Quest RPG Game"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   359
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrFPS 
      Interval        =   1000
      Left            =   840
      Top             =   3240
   End
   Begin VB.PictureBox objm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   3
      Left            =   4560
      Picture         =   "frmGame.frx":0000
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   56
      Top             =   3480
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox obj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   3
      Left            =   3960
      Picture         =   "frmGame.frx":1EA2
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   55
      Top             =   3480
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox obj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   2
      Left            =   5760
      Picture         =   "frmGame.frx":3D44
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   52
      Top             =   3360
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox objm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   2
      Left            =   6360
      Picture         =   "frmGame.frx":5BE6
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   51
      Top             =   3360
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   14
      Left            =   3840
      Picture         =   "frmGame.frx":7A88
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   49
      Top             =   4320
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   13
      Left            =   3000
      Picture         =   "frmGame.frx":7EBA
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   43
      Top             =   3480
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   12
      Left            =   2760
      Picture         =   "frmGame.frx":82EC
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   42
      Top             =   3480
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   11
      Left            =   3000
      Picture         =   "frmGame.frx":871E
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   41
      Top             =   3240
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   10
      Left            =   2760
      Picture         =   "frmGame.frx":8B50
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   40
      Top             =   3240
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   9
      Left            =   2160
      Picture         =   "frmGame.frx":8F82
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   36
      Top             =   3960
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   8
      Left            =   2160
      Picture         =   "frmGame.frx":93B4
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   35
      Top             =   3600
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   7
      Left            =   2160
      Picture         =   "frmGame.frx":97E6
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   34
      Top             =   3240
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox buffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   7560
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   32
      Top             =   1560
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox objm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   1
      Left            =   7080
      Picture         =   "frmGame.frx":9C18
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   31
      Top             =   1320
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox objm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   7080
      Picture         =   "frmGame.frx":BABA
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   30
      Top             =   960
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox obj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1080
      Index           =   1
      Left            =   6480
      Picture         =   "frmGame.frx":BEEC
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   29
      Top             =   1320
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox obj 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   6720
      Picture         =   "frmGame.frx":DD8E
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   28
      Top             =   960
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   6
      Left            =   5400
      Picture         =   "frmGame.frx":E1C0
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox enemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   6720
      Picture         =   "frmGame.frx":E5F2
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   26
      Top             =   3360
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox enemy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Index           =   0
      Left            =   6720
      Picture         =   "frmGame.frx":F294
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   23
      Top             =   2760
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox charm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   5880
      Picture         =   "frmGame.frx":FE76
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox chars 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   5880
      Picture         =   "frmGame.frx":119B8
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   5
      Left            =   5040
      Picture         =   "frmGame.frx":134FA
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   19
      Top             =   3720
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   4
      Left            =   5040
      Picture         =   "frmGame.frx":1392C
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   18
      Top             =   3360
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   3
      Left            =   5040
      Picture         =   "frmGame.frx":13D5E
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   2
      Left            =   5040
      Picture         =   "frmGame.frx":14190
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   5040
      Picture         =   "frmGame.frx":1496A
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Tile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   5040
      Picture         =   "frmGame.frx":14D9C
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Timer tmrM 
      Interval        =   10
      Left            =   120
      Top             =   3000
   End
   Begin VB.PictureBox main 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   5415
      TabIndex        =   6
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit Game"
         Height          =   255
         Left            =   1320
         TabIndex        =   58
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start Game"
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "By Kevin Fleet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Magic Quest"
         BeginProperty Font 
            Name            =   "Tempus Sans ITC"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.PictureBox battle 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   3255
      TabIndex        =   4
      Top             =   0
      Width           =   3255
      Begin VB.Timer tmrMagic 
         Interval        =   100
         Left            =   2520
         Top             =   1680
      End
      Begin VB.CommandButton cmdMagic 
         Caption         =   "Magic!"
         Height          =   255
         Left            =   1200
         TabIndex        =   48
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run!"
         Height          =   255
         Left            =   2160
         TabIndex        =   44
         Top             =   2400
         Width           =   975
      End
      Begin VB.PictureBox meter2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   2640
         ScaleHeight     =   5
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   25
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox meter1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   360
         ScaleHeight     =   5
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   24
         Top             =   600
         Width           =   495
      End
      Begin VB.Timer tmrMeters 
         Interval        =   10
         Left            =   600
         Top             =   1680
      End
      Begin VB.Timer tmrAIAttack 
         Interval        =   1000
         Left            =   1080
         Top             =   1680
      End
      Begin VB.Timer tmrAttack2 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   2040
         Top             =   1680
      End
      Begin VB.Timer tmrAttack 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   1560
         Top             =   1680
      End
      Begin VB.CommandButton cmdAttack 
         Caption         =   "Attack!"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   975
      End
      Begin VB.Image imgMagic 
         Height          =   480
         Left            =   2640
         Picture         =   "frmGame.frx":151CE
         Top             =   840
         Width           =   480
      End
      Begin VB.Image imgAttack2 
         Height          =   150
         Left            =   1680
         Picture         =   "frmGame.frx":154D8
         Top             =   1080
         Width           =   900
      End
      Begin VB.Image imgAttack 
         Height          =   150
         Left            =   720
         Picture         =   "frmGame.frx":15C22
         Top             =   1080
         Width           =   900
      End
      Begin VB.Image imgEnemy 
         Height          =   465
         Left            =   2640
         Picture         =   "frmGame.frx":1636C
         Top             =   840
         Width           =   465
      End
      Begin VB.Image imgPlayer 
         Height          =   495
         Left            =   360
         Picture         =   "frmGame.frx":16F4E
         Top             =   840
         Width           =   450
      End
   End
   Begin VB.PictureBox market 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   3255
      TabIndex        =   8
      Top             =   -120
      Width           =   3255
      Begin VB.Timer tmrMoney 
         Left            =   1800
         Top             =   1200
      End
      Begin VB.ListBox lstMCost 
         Height          =   255
         Left            =   1320
         TabIndex        =   39
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ListBox lstICost 
         Height          =   255
         Left            =   1320
         TabIndex        =   38
         Top             =   1920
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdSell 
         Caption         =   "Sell"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ListBox lstStuff 
         Height          =   1230
         Left            =   1920
         TabIndex        =   16
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuy 
         Caption         =   "Buy"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ListBox lstBuy 
         Height          =   1425
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblCost 
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   54
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cost:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   53
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "What do you want to buy?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1440
         TabIndex        =   50
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblMoney 
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2040
         TabIndex        =   45
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Money:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         Top             =   2520
         Width           =   735
      End
   End
   Begin VB.PictureBox menu 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton cmdDrop 
         Caption         =   "Drop Item"
         Height          =   255
         Left            =   1080
         TabIndex        =   57
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdUseItem 
         Caption         =   "Use Item"
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ListBox lstItems 
         Height          =   1815
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.PictureBox stats 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3240
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   33
      Top             =   0
      Width           =   2175
      Begin VB.CommandButton cmdMenu 
         Caption         =   "Menu"
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   2520
         Width           =   1575
      End
   End
   Begin VB.Label lblM 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Welcome to Magic Quest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   46
      Top             =   2880
      Width           =   5415
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAttack_Click()
imgAttack.Visible = True
tmrAttack.Enabled = True
End Sub

Private Sub cmdBuy_Click()
If lstBuy.ListIndex < 0 Then Exit Sub

Dim Cost

Cost = lstMCost.List(lstBuy.ListIndex)

If Cost < P.Pack.money Then

    P.Pack.money = P.Pack.money - Cost
    
    If AddPackItem(lstBuy.List(lstBuy.ListIndex), lstMCost.List(lstBuy.ListIndex)) = False Then
    P.Pack.money = P.Pack.money + Cost
    End If

    lstStuff.Clear

    For I = 1 To 20
        If P.Pack.Item(I).Act = True Then
            lstStuff.AddItem P.Pack.Item(I).Caption
        End If
    Next I

    frmGame.lblMoney.Caption = P.Pack.money & "$"

Else

    MsgBox ("You dont have enough money to buy item!"), , "MagicQuest"

End If
End Sub

Private Sub cmdDrop_Click()
If lstItems.ListIndex < 0 Then Exit Sub

If MsgBox("Are you sure you want to drop the item " & lstItems.List(lstItems.ListIndex) & "?", vbYesNo, "Drop Item") = vbNo Then Exit Sub

P.Pack.Item(lstItems.ListIndex + 1).Act = False
lstItems.RemoveItem lstItems.ListIndex
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdMagic_Click()
tmrMagic.Enabled = True
imgMagic.Visible = True
End Sub

Private Sub cmdMenu_Click()
Select Case cmdMenu.Caption
    Case "Menu"
        LoadMenu
        cmdMenu.Caption = "Back to Game"
    Case "Back to Game"
        main.Visible = False
        board.Visible = True
        battle.Visible = False
        menu.Visible = False
        market.Visible = False
        cmdMenu.Caption = "Menu"
End Select
End Sub

Private Sub cmdMenu_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then KeyCode = 0
End Sub

Private Sub cmdRun_Click()
If Enem.MaxHP > Int(Rnd * 5) + 10 Then Exit Sub

main.Visible = False
board.Visible = True
battle.Visible = False
menu.Visible = False
market.Visible = False
End Sub

Private Sub cmdSell_Click()
If lstStuff.ListIndex < 0 Then Exit Sub

Dim I As Long

I = lstStuff.ListIndex

If MsgBox("Are you sure want to sell " & lstStuff.List(I) & " for " & lstICost.List(I), vbYesNo, "Sell") = vbNo Then Exit Sub

P.Pack.money = P.Pack.money + Val(lstICost.List(I))
P.Pack.Item(I + 1).Act = False

lstStuff.Clear

For I = 1 To 20
    If P.Pack.Item(I).Act = True Then
        lstStuff.AddItem P.Pack.Item(I).Caption
    End If
Next I

frmGame.lblMoney.Caption = P.Pack.money & "$"
End Sub

Private Sub cmdStart_Click()
NewGame
main.Visible = False
board.Visible = True
battle.Visible = False
menu.Visible = False
market.Visible = False
stats.Visible = True
End Sub

Private Sub cmdUseItem_Click()
If lstItems.ListIndex < 0 Then Exit Sub
UseItem UCase(lstItems.List(lstItems.ListIndex)), lstItems.ListIndex
lstItems.RemoveItem lstItems.ListIndex
End Sub

Private Sub Form_Load()
Dim C, I As Long, per As Long, place, text

Randomize

stats.Visible = False
main.Visible = True
board.Visible = False
battle.Visible = False
menu.Visible = False
market.Visible = False

MessageTime = 10 + Len(lblM.Caption)

Me.Visible = True

Do

    If C > 1 And GetTickCount > 300 Then
        C = 0
        FPS = FPS + 1
        
        If board.Visible = True Then

            buffer.Cls

            MovePlayer
            MoveObjs

            For I = 0 To M.Mapheight * M.Mapwidth

                If isValue(T(I).Type, 2) = True Then
    
                    If T(I).AniC > 10 Then
                        T(I).AniC = 0
                        T(I).Ani = T(I).Ani + 1: If T(I).Ani > 1 Then T(I).Ani = 0
                    Else
                        T(I).AniC = T(I).AniC + 1
                    End If
    
                End If
    
           E.DrawObj buffer.hdc, T(I).X, T(I).Y, TileWidth, TileHeight, Tile(T(I).Type).hdc, T(I).Ani * 18, 0, vbSrcCopy
    
    Next I

    For I = 1 To 20
       If o(I).Act = True Then
            E.DrawObj buffer.hdc, o(I).X, o(I).Y, TileWidth, TileHeight, objm(o(I).T).hdc, o(I).Step * 18, o(I).dir * 18, Mask
            E.DrawObj buffer.hdc, o(I).X, o(I).Y, TileWidth, TileHeight, obj(o(I).T).hdc, o(I).Step * 18, o(I).dir * 18, Sprite
        End If
    Next I

    E.DrawObj buffer.hdc, P.X, P.Y, 18, 18, charm.hdc, P.Step * 18, P.dir * 18, Mask
    E.DrawObj buffer.hdc, P.X, P.Y, 18, 18, chars.hdc, P.Step * 18, P.dir * 18, Sprite

    board.Cls
    E.DrawObj board.hdc, 0, 0, board.ScaleWidth, board.ScaleHeight, buffer.hdc, (P.X + 9) - (board.ScaleWidth \ 2), (P.Y + 9) - (board.ScaleHeight \ 2), vbSrcCopy
End If

If main.Visible = False Then

    stats.Cls

    For place = 1 To 10

        Select Case place
            Case 1
                per = P.B.Health \ P.B.MaxHealth
                per = per * 100
    
                 text = "Health: " & P.B.Health & " Out of " & P.B.MaxHealth
    
            Case 2: text = "Money: " & P.Pack.money & "$"
             Case 3: text = "Magic: " & P.Pack.mDust
            Case 4: text = ""
            Case 5: text = "Stats"
            Case 6: text = "Attack: " & P.B.Attack
            Case 7: text = "Defense: " & P.B.Defense
            Case 8: text = "Exp: " & P.B.Exp & " Out of " & P.B.TExp
            Case 9: text = "Level: " & P.B.Level
            Case 10: text = ""
            
        End Select
    
        stats.CurrentX = stats.ScaleWidth \ 2 - stats.TextWidth(text) \ 2
        stats.CurrentY = stats.TextHeight("|") * place + 10
        stats.Print text
        
    Next place
    
End If

Else

    C = C + 1
    
End If

DoEvents

Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstBuy_Click()
lblCost.Caption = Val(lstMCost.List(lstBuy.ListIndex)) & "$"
End Sub

Private Sub tmrAIAttack_Timer()
If battle.Visible = False Then Exit Sub
imgAttack2.Visible = True
tmrAttack2.Enabled = True
End Sub

Private Sub tmrAttack_Timer()
If battle.Visible = False Then Exit Sub
imgAttack.Visible = False
tmrAttack.Enabled = False

Enem.HP = Enem.HP - IIf((P.B.Attack - Enem.Defense) < 0, 0, (P.B.Attack - Enem.Defense))

If Enem.HP <= 0 Then

    P.B.Exp = P.B.Exp + (Enem.Attack + Enem.Defense)

    main.Visible = False
    board.Visible = True
    battle.Visible = False
    menu.Visible = False
    market.Visible = False

    If Int(Rnd * 5) = Int(Rnd * 4) Then
    
        Dim money
        money = Int(Rnd * 20) + 10
        P.Pack.money = P.Pack.money + money

        lblM.Caption = money & "$ was found on the ground..."
        MessageTime = 10 + Len(lblM.Caption)
    End If

    If P.B.Exp >= P.B.TExp Then
        P.B.Exp = 0
        P.B.Level = P.B.Level + 1
        P.B.Attack = P.B.Attack + 2
        P.B.Defense = P.B.Defense + 1
        P.B.Health = P.B.Health + 5
        P.B.MaxHealth = P.B.MaxHealth + 5
        P.B.TExp = P.B.TExp + (Int(Rnd * 10) + 1)
    End If
End If
End Sub

Private Sub tmrAttack2_Timer()
If battle.Visible = False Then Exit Sub
imgAttack2.Visible = False
tmrAttack2.Enabled = False

If Int(Rnd * 3) = 0 Then Exit Sub

P.B.Health = P.B.Health - IIf((Enem.Attack - P.B.Defense) < 0, 0, (Enem.Attack - P.B.Defense))

If P.B.Health <= 0 Then
    P.B.Health = 50

    main.Visible = False
    board.Visible = True
    battle.Visible = False
    menu.Visible = False
    market.Visible = False

    LoadMap "house.dat", 1, 2
End If
End Sub

Private Sub tmrFPS_Timer()
Me.Caption = "Magic Quest RPG Game - " & FPS & " FPS"
FPS = 0
End Sub

Private Sub tmrM_Timer()
If MessageTime > 0 Then

    MessageTime = MessageTime - 1
    
    If MessageTime <= 0 Then
        lblM.Caption = ""
    End If

Else

lblM.Caption = ""
End If
End Sub

Private Sub tmrMagic_Timer()
tmrMagic.Enabled = False
imgMagic.Visible = False

If battle.Visible = True And P.Pack.mDust > 0 Then

    P.Pack.mDust = P.Pack.mDust - 2: If P.Pack.mDust < 0 Then P.Pack.mDust = 0

    Enem.HP = Enem.HP - IIf(((P.B.Attack * 2) - Enem.Defense) < 0, 0, (P.B.Attack - Enem.Defense))

    If Int(Rnd * 5) = Int(Rnd * 4) Then
        Dim money
        money = Int(Rnd * 20) + 10
        lblM.Caption = money & "$ was found on the ground..."
        MessageTime = 10 + Len(lblM.Caption)
    End If

    If Enem.HP <= 0 Then
        P.B.Exp = P.B.Exp + (Enem.Attack + Enem.Defense)
        main.Visible = False
        board.Visible = True
        battle.Visible = False
        menu.Visible = False
        market.Visible = False
        
        If P.B.Exp >= P.B.TExp Then
            P.B.Exp = 0
            P.B.Level = P.B.Level + 1
            P.B.Attack = P.B.Attack + 2
            P.B.Defense = P.B.Defense + 1
            P.B.Health = P.B.Health + 5
            P.B.MaxHealth = P.B.MaxHealth + 5
            P.B.TExp = P.B.TExp + (Int(Rnd * 10) + 1)
        End If
    End If
End If
End Sub

Private Sub tmrMeters_Timer()
If battle.Visible = False Then Exit Sub

Dim per As Double

meter1.Cls
meter2.Cls

per = (P.B.Health \ P.B.MaxHealth) * 100 / 100
meter1.Line (0, 0)-(meter1.ScaleWidth * per, meter1.ScaleHeight), vbGreen, BF
per = (Enem.HP \ Enem.MaxHP) * 100 / 100
meter2.Line (0, 0)-(meter2.ScaleWidth * per, meter2.ScaleHeight), vbGreen, BF
End Sub

Private Sub tmrMoney_Timer()
lblMoney.Caption = P.Pack.money
End Sub

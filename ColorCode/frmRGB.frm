VERSION 5.00
Begin VB.Form frmRGB 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DNA  Color Conversion"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   8040
      Top             =   360
   End
   Begin VB.Frame frameCurColor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Colore corrente"
      ForeColor       =   &H00404040&
      Height          =   2775
      Left            =   0
      TabIndex        =   16
      Top             =   960
      Width           =   6135
      Begin VB.TextBox txtVbQbValue 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   67
         Top             =   2400
         Width           =   1335
      End
      Begin VB.HScrollBar hLite 
         Height          =   255
         Left            =   120
         Max             =   255
         Min             =   1
         TabIndex        =   62
         Top             =   1320
         Value           =   1
         Width           =   3615
      End
      Begin VB.HScrollBar sldBlue 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   61
         Top             =   960
         Width           =   3615
      End
      Begin VB.HScrollBar sldGreen 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   60
         Top             =   600
         Width           =   3615
      End
      Begin VB.HScrollBar sldRed 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   59
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txtHtmlValue 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtVbValue 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   52
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtLongValue 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtHexValue 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2400
         Width           =   975
      End
      Begin VB.PictureBox picCurColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   5865
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Trascina il colore nella lista dei preferiti"
         Top             =   1680
         Width           =   5895
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Const."
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   4680
         TabIndex        =   68
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblDL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5400
         TabIndex        =   64
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Dark <> Light"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   3960
         TabIndex        =   63
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Html Code"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1200
         TabIndex        =   58
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "VB Color Value"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   3360
         TabIndex        =   53
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Long Value"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   2280
         TabIndex        =   43
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblR1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5400
         TabIndex        =   42
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblG1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5400
         TabIndex        =   41
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblB1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5400
         TabIndex        =   40
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4560
         TabIndex        =   34
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4560
         TabIndex        =   33
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4560
         TabIndex        =   32
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblBlueLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "&Blue"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   3960
         TabIndex        =   22
         Top             =   990
         Width           =   495
      End
      Begin VB.Label lblGreenLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "&Green"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   3960
         TabIndex        =   21
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblRedLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "&Red"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   3960
         TabIndex        =   20
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblValueLbl 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hex Code"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   1335
      End
   End
   Begin VB.Frame frameFavorites 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Colori preferiti"
      ForeColor       =   &H00404040&
      Height          =   1215
      Left            =   6240
      TabIndex        =   0
      Top             =   3720
      Width           =   2295
      Begin VB.PictureBox picFavorite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   1800
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox picFavorite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   1800
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox picFavorite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   1800
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox picFavorite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   1440
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox picFavorite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   1440
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox picFavorite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   1440
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox picFavorite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   1080
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox picFavorite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   1080
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox picFavorite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   1080
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox picFavorite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   720
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox picFavorite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   720
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox picFavorite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   720
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox picFavorite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   360
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox picFavorite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   360
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox picFavorite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   360
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblRp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Conversioni 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Conversioni"
      ForeColor       =   &H00404040&
      Height          =   2775
      Left            =   6240
      TabIndex        =   35
      Top             =   960
      Width           =   2295
      Begin VB.CommandButton cmdRandom 
         Caption         =   "Random"
         Height          =   255
         Left            =   1200
         TabIndex        =   71
         Top             =   2400
         Width           =   975
      End
      Begin VB.CheckBox chkCapture 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Capture"
         Height          =   255
         Left            =   1200
         TabIndex        =   70
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox hexColor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   54
         Text            =   "000000"
         Top             =   2880
         Width           =   2040
      End
      Begin VB.TextBox decColor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   51
         Text            =   "0"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox hexBlue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   50
         Text            =   "00"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox hexGreen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   49
         Text            =   "00"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox hexRed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         MaxLength       =   2
         TabIndex        =   48
         Text            =   "00"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox decBlue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1680
         TabIndex        =   47
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox decGreen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   46
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox decRed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   45
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ERRORE !!!"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   720
         TabIndex        =   65
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblColorBox 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   120
         TabIndex        =   56
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "HEX   #"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   100
         TabIndex        =   39
         Top             =   880
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "LONG &&"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   100
         TabIndex        =   38
         Top             =   1240
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "R-G-B"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   100
         TabIndex        =   37
         Top             =   520
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CAPTURE COLOR"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1680
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   9375
      TabIndex        =   29
      Top             =   0
      Width           =   9375
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Developer's CodeBook 2001"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   2655
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   9360
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "CodeBook Color Conversion Tool"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   360
         TabIndex        =   30
         Top             =   360
         Width           =   6015
      End
      Begin VB.Image Image2 
         Height          =   555
         Left            =   7560
         Stretch         =   -1  'True
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame frameOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opzioni di copia automatica"
      ForeColor       =   &H00404040&
      Height          =   1215
      Left            =   0
      TabIndex        =   23
      Top             =   3720
      Width           =   6135
      Begin VB.CheckBox chkAutoCopy 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Copia automaticamente nella ClipBoard nei valori :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   360
         Width           =   5655
      End
      Begin VB.OptionButton OptVB 
         BackColor       =   &H00808080&
         Caption         =   "Visual Basic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   55
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton optHtm 
         BackColor       =   &H00808080&
         Caption         =   "Html"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   66
         Top             =   840
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.TextBox txtNumChars 
         Height          =   285
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   27
         Text            =   "6"
         Top             =   1440
         Width           =   495
      End
      Begin VB.OptionButton optHex 
         BackColor       =   &H00808080&
         Caption         =   "Hex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton optDecimal 
         BackColor       =   &H00808080&
         Caption         =   "Long"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblNumberSizeLbl 
         Caption         =   "Numero Caratteri"
         Height          =   255
         Left            =   3600
         TabIndex        =   26
         Top             =   1440
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC& Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long)

Private mbCapture As Boolean
Private mbNoChange As Boolean
Private mlCurColor As Long
Dim initVals(4) As Long

Private Type ColorRGB
  Red As Long
  Green As Long
  Blue As Long
End Type

Private Type PointAPI
    x   As Long
    y   As Long
End Type

Const HexString As String = "0123456789ABCDEF"

Private Sub chkCapture_Click()
' Avvio Routine cattura colore
  If chkCapture.Value = 1 Then
     Label6.Caption = "CAPUTURE COLOR"
     Timer1.Interval = 10
     Timer1.Enabled = True
     mbCapture = True
     Screen.MousePointer = 2
     Call ReleaseCapture
     Call SetCapture(Me.hwnd)
  End If
End Sub

Private Sub cmdRandom_Click()
' Colorazione Random
  Label6.Caption = "RANDOM COLOR"
  Randomize Timer
  setCurrentColor RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
End Sub

Private Sub decBlue_Change()
' Codifica dolore RGB(B)
  On Error GoTo err
  sldBlue.Value = Val(decBlue)
  decBlue_Validate False
err:
  If err.Number <> 0 Then Label10.Visible = True
End Sub

Private Sub decGreen_Change()
' Codifica dolore RGB(G)
  On Error GoTo err
  sldGreen.Value = Val(decGreen)
  decGreen_Validate False
err:
  If err.Number <> 0 Then Label10.Visible = True
End Sub

Private Sub decRed_Change()
' Codifica dolore RGB(R)
  On Error GoTo err
  sldRed.Value = Val(decRed)
  decRed_Validate False
err:
  If err.Number <> 0 Then Label10.Visible = True
End Sub

Private Sub decColor_Change()
' Codifica colore in formato VB
  On Error GoTo err
  decColor_Validate False
  txtVbValue.Text = "&H00" & hexBlue & hexGreen & hexRed & "&"
err:
  If err.Number <> 0 Then Label10.Visible = True
End Sub

Private Sub Form_Load()
' Carica gli ultimi settaggi
  'Image2.Picture = LoadResImage("LOGO", "LOGO")
  loadPreviousSettings
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' Rilascio del Mouse per la cattura del colore
  Dim lColor  As Long
  Dim hDeskDC As Long
  Dim ptMouse As PointAPI
  If mbCapture Then
     Call ReleaseCapture
     Call GetCursorPos(ptMouse)
     hDeskDC = GetDC(0)
     lColor = GetPixel(hDeskDC, ptMouse.x, ptMouse.y)
     If lColor <> -1 Then
        mlCurColor = lColor
        Call ResetControls
     End If
     Timer1.Enabled = False
     Screen.MousePointer = vbDefault
     chkCapture.Value = 0
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
' Salva i settaggi
  saveCurrentSettings
End Sub

Private Sub hexRed_Change()
' Modifica Colore HEX(R)
  On Error GoTo err
  If Len(hexRed) > 1 Then hexRed_Validate False
err:
  If err.Number <> 0 Then Label10.Visible = True
End Sub

Private Sub hexGreen_Change()
' Modifica Colore HEX(G)
  On Error GoTo err
  If Len(hexGreen) > 1 Then hexGreen_Validate False
err:
  If err.Number <> 0 Then Label10.Visible = True
End Sub

Private Sub hexBlue_Change()
' Modifica Colore HEX(B)
  On Error GoTo err
  If Len(hexBlue) > 1 Then hexBlue_Validate False
err:
  If err.Number <> 0 Then Label10.Visible = True
End Sub

Private Sub hLite_Change()
' Sincronizzazione delle Slide per incremento/decremento luminosita
  On Error Resume Next
  Dim iRed        As Integer
  Dim iGreen      As Integer
  Dim iBlue       As Integer
  Dim iChange     As Integer
  Dim lColor      As Long
  Static iOldVal  As Integer
  lblDL.Caption = hLite.Value
  If Not mbNoChange Then
     iChange = hLite.Value - iOldVal
     lColor = RGB(sldRed.Value, sldGreen.Value, sldBlue.Value)
     iRed = (lColor And &HFF&)
     iGreen = (lColor And &HFF00&) / &H100
     iBlue = (lColor And &HFF0000) / &H10000
     iRed = iRed + iChange
     iGreen = iGreen + iChange
     iBlue = iBlue + iChange
     If iRed > 255 Then
        iRed = 255
     ElseIf iRed < 0 Then
        iRed = 0
     End If
     If iGreen > 255 Then
        iGreen = 255
     ElseIf iGreen < 0 Then
        iGreen = 0
     End If
     If iBlue > 255 Then
        iBlue = 255
     ElseIf iBlue < 0 Then
        iBlue = 0
     End If
     mlCurColor = RGB(iRed, iGreen, iBlue)
     Call ResetControls
  End If
  iOldVal = hLite.Value
End Sub

Private Sub Hlite_Scroll()
' Attiva il refresh sullo scrolling
  Call hLite_Change
End Sub

Private Sub picFavorite_Click(Index As Integer)
' Carica il colore scelto nei colori preferiti
  setCurrentColor picFavorite(Index).BackColor
End Sub

Private Sub picFavorite_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
' Modifica del colore dei preferiti
  If Source.Name = "picCurColor" Then
     picFavorite(Index).BackColor = Source.BackColor
  End If
End Sub

Private Sub sldRed_Change()
' Modifica valori RGB(R) su Slide
  Label10.Visible = False
  setCurrentColor RGB(sldRed.Value, sldGreen.Value, sldBlue.Value)
  lblR.BackColor = RGB(sldRed.Value, 0, 0)
  lblR1.Caption = sldRed.Value
End Sub

Private Sub sldGreen_Change()
' Modifica valori RGB(G) su Slide
  Label10.Visible = False
  setCurrentColor RGB(sldRed.Value, sldGreen.Value, sldBlue.Value)
  lblG.BackColor = RGB(0, sldGreen.Value, 0)
  lblG1.Caption = sldGreen.Value
End Sub

Private Sub sldBlue_Change()
' Modifica valori RGB(B) su Slide
  Label10.Visible = False
  setCurrentColor RGB(sldRed.Value, sldGreen.Value, sldBlue.Value)
  lblB.BackColor = RGB(0, 0, sldBlue.Value)
  lblB1.Caption = sldBlue.Value
End Sub

Private Sub sldBlue_Scroll()
' Modifica valore scrolling
  sldBlue_Change
End Sub

Private Sub sldGreen_Scroll()
' Modifica valore scrolling
  sldGreen_Change
End Sub

Private Sub sldRed_Scroll()
' Modifica valore scrolling
  sldRed_Change
End Sub

Private Sub txtNumChars_Change()
' Verifica valori immessi
  If Val(txtNumChars.Text) = 0 Then
     Beep
     txtNumChars.Text = "1"
     txtNumChars.SelStart = 0
     txtNumChars.SelLength = Len(txtNumChars.Text)
  End If
End Sub

Private Sub txtNumChars_GotFocus()
' Selezione valori codifica Long
  txtNumChars.SelStart = 0
  txtNumChars.SelLength = Len(txtNumChars.Text)
End Sub

Private Sub txtNumChars_KeyPress(KeyAscii As Integer)
' Validazione input utente
  If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then
     Beep
     KeyAscii = 0
  End If
End Sub

Private Sub txtNumChars_LostFocus()
' Assegnazione valori impostati
  setCurrentColor picCurColor.BackColor
End Sub

Private Sub txtHexValue_GotFocus()
' Selezione valori HEX
  txtHexValue.SelStart = 0
  txtHexValue.SelLength = Len(txtHexValue.Text)
End Sub

Private Function getValueString(ByVal Color As Long) As String
' Elaborazione valori colore
  Dim rval As String, hexval As String, R As Byte, G As Byte, B As Byte
  Dim rVal1 As String
  rVal1 = Trim(Str(Color))
' Formato DEC
' aggiunta degli zeri neccessari
  If Len(rVal1) > Val(txtNumChars.Text) Then
     txtNumChars.Text = Trim(Str(Len(rVal1)))
  Else
     rVal1 = String(Val(txtNumChars.Text) - Len(rVal1), "0") & rVal1
  End If
  R = Color Mod &H100: Color = Color \ &H100
  G = Color Mod &H100: Color = Color \ &H100
  B = Color Mod &H100
' Formato HEX
  rval = String(2 - Len(Hex(R)), "0") & Hex(R)
  rval = rval & String(2 - Len(Hex(G)), "0") & Hex(G)
  rval = rval & String(2 - Len(Hex(B)), "0") & Hex(B)
  getValueString = rval
  txtLongValue.Text = "&" & rVal1
End Function

Private Sub setCurrentColor(ByVal Color As Long)
' impostazione colore corrente e copia sulla clipBoard
  Dim R As Byte, G As Byte, B As Byte
  Clipboard.Clear
  picCurColor.BackColor = Color
  lblRp.BackColor = Color
  txtHexValue.Text = "#" & getValueString(Color)
  txtHtmlValue.Text = "#" & Right$(getValueString(Color), 2) & _
                            Mid$(getValueString(Color), 3, 2) & _
                            Left$(getValueString(Color), 2)
  If chkAutoCopy.Value = 1 Then
     If optDecimal.Value Then
        Clipboard.SetText txtLongValue.Text
     ElseIf optHex.Value Then
        Clipboard.SetText txtHexValue.Text
     ElseIf OptVB.Value Then
        Clipboard.SetText txtVbValue.Text
     ElseIf optHtm.Value Then
        Clipboard.SetText txtHtmlValue.Text
     Else
        'non e' possibile...
     End If
  End If
' Aggiornamento dei componenti red, green, blue
  R = Color Mod &H100: Color = Color \ &H100
  G = Color Mod &H100: Color = Color \ &H100
  B = Color Mod &H100
  If Me.ActiveControl Is Nothing Then
     decRed.Text = Trim(Str(R))
     decGreen.Text = Trim(Str(G))
     decBlue.Text = Trim(Str(B))
  Else
     If Me.ActiveControl.Name <> "decRed" Then decRed.Text = Trim(Str(R))
     If Me.ActiveControl.Name <> "decGreen" Then decGreen.Text = Trim(Str(G))
     If Me.ActiveControl.Name <> "decBlue" Then decBlue.Text = Trim(Str(B))
  End If
  txtVbQbValue.Text = "QbColor(" & spy_colorqb & ")"
  Label12.Caption = "Const(" & spy_colorvb$ & ")"
  'txtVbConstValue.Text =

End Sub

Private Sub loadPreviousSettings()
' Caricamento parametri precedenti
  Dim i As Integer
  For i = picFavorite.LBound To picFavorite.UBound
      picFavorite(i).BackColor = GetSetting(App.Title, "ColorFavorites", _
      Trim(Str(i)), QBColor(i + 1))
  Next
  setCurrentColor GetSetting(App.Title, "ColorSettings", "Current Color", 0)
  chkAutoCopy.Value = GetSetting(App.Title, "ColorSettings", "AutoCopy", 1)
  optDecimal.Value = GetSetting(App.Title, "ColorSettings", "Decimal", False)
  optHex.Value = GetSetting(App.Title, "ColorSettings", "Hexadecimal", True)
  optHtm.Value = GetSetting(App.Title, "ColorSettings", "Html", False)
  OptVB.Value = GetSetting(App.Title, "ColorSettings", "VB", False)
  txtNumChars.Text = GetSetting(App.Title, "ColorSettings", "Number of Characters", "6")
End Sub

Private Sub saveCurrentSettings()
' Salvataggio parametri correnti
  Dim i As Integer
  For i = picFavorite.LBound To picFavorite.UBound
      SaveSetting App.Title, "ColorFavorites", Trim(Str(i)), picFavorite(i).BackColor
  Next
  SaveSetting App.Title, "ColorSettings", "Current Color", picCurColor.BackColor
  SaveSetting App.Title, "ColorSettings", "AutoCopy", chkAutoCopy.Value
  SaveSetting App.Title, "ColorSettings", "Decimal", optDecimal.Value
  SaveSetting App.Title, "ColorSettings", "Hexadecimal", optHex.Value
  SaveSetting App.Title, "ColorSettings", "Html", optHtm.Value
  SaveSetting App.Title, "ColorSettings", "VB", OptVB.Value
  SaveSetting App.Title, "ColorSettings", "Number of Characters", txtNumChars.Text
End Sub

Private Sub decBlue_GotFocus()
' Selezione valori Correnti(B)
  Label10.Visible = False
  Label6.Caption = "RGB TO LONG - RGB TO HEX"
  initVals(2) = Val(decBlue.Text)
  decBlue.SelStart = 0
  decBlue.SelLength = Len(decBlue.Text)
End Sub

Private Sub decBlue_KeyPress(KeyAscii As Integer)
' Validazione Input Utente
  If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
     If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys vbTab
     ElseIf KeyAscii = Asc(vbBack) Then
        ' fine...
     Else
        KeyAscii = 0
     End If
  End If
End Sub

Private Sub decBlue_Validate(Cancel As Boolean)
' Modifica valori impostati su blue
  On Error GoTo err
  If Len(decBlue.Text) = 0 Then decBlue.Text = "0"
  If Val(decBlue.Text) > 255 Then
     Beep
     decBlue.Text = CStr(initVals(2))
     decBlue.SetFocus
  Else
     initVals(2) = CLng(decBlue.Text)
     reFreshAll initVals(0), initVals(1), initVals(2)
  End If
err:
  If err.Number <> 0 Then Label10.Visible = True
End Sub

Private Sub decColor_GotFocus()
' Conversione da LONG a RGB/HEX
  Label10.Visible = False
  Label6.Caption = " LONG TO RGB - LONG TO HEX"
  initVals(3) = Val(decColor.Text)
  decColor.SelStart = 0
  decColor.SelLength = Len(decColor.Text)
End Sub

Private Sub decColor_KeyPress(KeyAscii As Integer)
' Validazione Input Utente
  If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
     If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys vbTab
     ElseIf KeyAscii = Asc(vbBack) Then
        ' fine...
     Else
        KeyAscii = 0
     End If
  End If
End Sub

Private Sub decColor_Validate(Cancel As Boolean)
' Validazione Input Utente
  On Error GoTo err
  If Len(decColor.Text) = 0 Then decColor.Text = "0"
  If Val(decColor.Text) > 16777215 Then
     Beep
     decColor.Text = CStr(initVals(3))
     decColor.SetFocus
  Else
     initVals(3) = CLng(decColor.Text)
     initVals(0) = initVals(3) And 255
     initVals(1) = (initVals(3) And 65280) \ 256&
     initVals(2) = (initVals(3) And 16711680) \ 65535
     reFreshAll initVals(0), initVals(1), initVals(2)
  End If
err:
  If err.Number <> 0 Then Label10.Visible = True
End Sub

Private Sub decGreen_GotFocus()
' Selezione Valori correnti(G)
  Label10.Visible = False
  Label6.Caption = " RGB TO LONG - RGB TO HEX"
  initVals(1) = Val(decGreen.Text)
  decGreen.SelStart = 0
  decGreen.SelLength = Len(decGreen.Text)
End Sub

Private Sub decGreen_KeyPress(KeyAscii As Integer)
' Validazione Input Utente
  If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
     If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys vbTab
     ElseIf KeyAscii = Asc(vbBack) Then
        ' fine...
     Else
        KeyAscii = 0
     End If
  End If
End Sub

Private Sub decGreen_Validate(Cancel As Boolean)
' Validazione Input Utente
  On Error GoTo err
  If Len(decGreen.Text) = 0 Then decGreen.Text = "0"
  If Val(decGreen.Text) > 255 Then
     Beep
     decGreen.Text = CStr(initVals(1))
     decGreen.SetFocus
  Else
     initVals(1) = CLng(decGreen.Text)
     reFreshAll initVals(0), initVals(1), initVals(2)
  End If
err:
  If err.Number <> 0 Then Label10.Visible = True
End Sub

Private Sub decRed_GotFocus()
' Selezione Valori Correnti(R)
  Label10.Visible = False
  Label6.Caption = " RGB TO LONG - RGB TO HEX"
  initVals(0) = Val(decRed.Text)
  decRed.SelStart = 0
  decRed.SelLength = Len(decRed.Text)
End Sub

Private Sub decRed_KeyPress(KeyAscii As Integer)
' Validazione Input Utente(R)
  If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
     If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys vbTab
     ElseIf KeyAscii = Asc(vbBack) Then
        ' fine...
     Else
        KeyAscii = 0
     End If
  End If
End Sub

Private Sub decRed_Validate(Cancel As Boolean)
' Validazione Input Utente(R)
  On Error GoTo err
  If Len(decRed.Text) = 0 Then decRed.Text = "0"
  If Val(decRed.Text) > 255 Then
     Beep
     decRed.Text = CStr(initVals(0))
     decRed.SetFocus
  Else
     initVals(0) = CLng(decRed.Text)
     reFreshAll initVals(0), initVals(1), initVals(2)
  End If
err:
  If err.Number <> 0 Then Label10.Visible = True
End Sub

Private Sub reFreshAll(Optional Red As Long, Optional Green As Long, _
        Optional Blue As Long, Optional cValue As Variant)
' Refresh colorazione elementi
  Dim Color As Long
  If Not IsMissing(cValue) Then
     Color = cValue
     Red = Color And 255
     Green = (Color And 65280) \ 256&
     Blue = (Color And 16711680) \ 65535
  Else
     Color = Red + Green * 256 + Blue * 256 * 256
  End If
  decRed.Text = CStr(Red)
  decGreen.Text = CStr(Green)
  decBlue.Text = CStr(Blue)
  hexRed.Text = PadL(Dec2Hex(Red), 2, "0")
  hexGreen.Text = PadL(Dec2Hex(Green), 2, "0")
  hexBlue.Text = PadL(Dec2Hex(Blue), 2, "0")
  hexColor.Text = (PadL(Dec2Hex(Color), 6, "0"))
  decColor.Text = CStr(Color)
  initVals(0) = Red
  initVals(1) = Green
  initVals(2) = Blue
  initVals(3) = Color
End Sub

Public Function Dec2Hex(deci As Long) As String
' Codifica Decimale -> HEX
  Dim ix As Long, iY As Long
  iY = deci \ 16
  ix = deci Mod 16
  Dec2Hex = Mid$(HexString, ix + 1, 1)
  Do While iY > 0
     ix = iY Mod 16
     Dec2Hex = Mid$(HexString, ix + 1, 1) & Dec2Hex
     iY = iY \ 16
  Loop
End Function

Private Function Hex2Dec(ByVal Hexi As String) As String
' Codifica HEX -> Decimale
  Dim ix As Long, iMult As Long, iY As Long, iDig As String
  iMult = 1
  Hexi = Trim$(Hexi)
  If Len(Hexi) = 0 Then Hexi = "0"
  For ix = 1 To Len(Hexi)
      iDig = Mid$(Hexi, Len(Hexi) - ix + 1, 1)
      If InStr(1, "0123456789", iDig) Then
         iY = iY + iMult * CLng(iDig)
      Else
         iY = iY + iMult * CLng(Asc(iDig) - Asc("A") + 10)
      End If
      iMult = iMult * 16
  Next ix
  Hex2Dec = CStr(iY)
End Function

Private Function PadL(sInput As String, iLen As Long, _
        Optional fillChar As String = " ") As String
' Conteggio spazi vuoti da riempire con 0
  If Len(sInput) <= iLen Then
     If fillChar = " " Then
        PadL = Space$(iLen - Len(sInput)) + sInput
     Else
        PadL = String$(iLen - Len(sInput), fillChar) + sInput
     End If
  Else
     PadL = Left$(sInput, iLen)
  End If
End Function

Private Sub hexBlue_GotFocus()
' Selezione Valori Correnti Hex(B)
  Label10.Visible = False
  Label6.Caption = " HEX TO RGB - HEX TO LONG"
  initVals(2) = Val(Hex2Dec(hexBlue.Text))
  hexBlue.SelStart = 0
  hexBlue.SelLength = Len(hexBlue.Text)
End Sub

Private Sub hexBlue_KeyPress(KeyAscii As Integer)
' Validazione Input Hex(B)
  If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") _
     Or KeyAscii >= Asc("A") And KeyAscii <= Asc("F") Then
  ElseIf KeyAscii >= Asc("a") And KeyAscii <= Asc("f") Then
     KeyAscii = Asc(UCase$(Chr$(KeyAscii))) ' upper case
  Else
     If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys vbTab
     ElseIf KeyAscii = Asc(vbBack) Then
        ' fine...
     Else
        KeyAscii = 0
     End If
  End If
End Sub

Private Sub hexBlue_Validate(Cancel As Boolean)
' Validazione Input Hex(B)
  On Error GoTo err
  If Len(hexBlue.Text) = 0 Then hexBlue.Text = "00"
  initVals(2) = CLng(Hex2Dec(hexBlue.Text))
  reFreshAll initVals(0), initVals(1), initVals(2)
err:
  If err.Number <> 0 Then Label10.Visible = True
End Sub

Private Sub hexColor_KeyPress(KeyAscii As Integer)
' Validazione Input Hex(C)
  If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") _
     Or KeyAscii >= Asc("A") And KeyAscii <= Asc("F") Then
  ElseIf KeyAscii >= Asc("a") And KeyAscii <= Asc("f") Then
     KeyAscii = Asc(UCase$(Chr$(KeyAscii))) ' upper case
  Else
     If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys vbTab
     ElseIf KeyAscii = Asc(vbBack) Then
        ' fine...
     Else
        KeyAscii = 0
     End If
  End If
End Sub

Private Sub hexColor_Validate(Cancel As Boolean)
' Validazione Input Hex(C)
  If Len(hexColor.Text) = 0 Then hexColor.Text = "000000"
  If Val(Hex2Dec(hexColor.Text)) > 16777215 Then
     Beep
     hexColor.Text = PadL(Dec2Hex(initVals(3)), 6, "0")
     hexColor.SetFocus
  Else
     initVals(0) = initVals(3) And 255
     initVals(1) = (initVals(3) And 65280) \ 256&
     initVals(2) = (initVals(3) And 16711680) \ 65535
     initVals(3) = CLng(Hex2Dec(hexColor.Text))
     reFreshAll initVals(2), initVals(1), initVals(0)
  End If
End Sub

Private Sub hexGreen_GotFocus()
' Selezione Valori Hex(G)
  Label10.Visible = False
  Label6.Caption = " HEX TO RGB - HEX TO LOMG"
  initVals(1) = Val(Hex2Dec(hexGreen.Text))
  hexGreen.SelStart = 0
  hexGreen.SelLength = Len(hexGreen.Text)
End Sub

Private Sub hexGreen_KeyPress(KeyAscii As Integer)
' Validazione Input Hex(G)
  If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") _
     Or KeyAscii >= Asc("A") And KeyAscii <= Asc("F") Then
  ElseIf KeyAscii >= Asc("a") And KeyAscii <= Asc("f") Then
     KeyAscii = Asc(UCase$(Chr$(KeyAscii))) ' upper case
  Else
     If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys vbTab
     ElseIf KeyAscii = Asc(vbBack) Then
        ' fine...
     Else
        KeyAscii = 0
     End If
  End If
End Sub

Private Sub hexGreen_Validate(Cancel As Boolean)
' Validazione Input Hex(G)
  On Error GoTo err
  If Len(hexGreen.Text) = 0 Then hexGreen.Text = "00"
  initVals(1) = CLng(Hex2Dec(hexGreen.Text))
  reFreshAll initVals(0), initVals(1), initVals(2)
err:
  If err.Number <> 0 Then Label10.Visible = True
End Sub

Private Sub hexred_GotFocus()
' Conversione da HEX a RGB/LONG
  Label10.Visible = False
  Label6.Caption = " HEX TO RGB - HEX TO LONG"
  initVals(0) = Val(Hex2Dec(hexRed.Text))
  hexRed.SelStart = 0
  hexRed.SelLength = Len(hexRed.Text)
End Sub

Private Sub hexRed_KeyPress(KeyAscii As Integer)
' Validazione Input Hex(R)
  If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") _
     Or KeyAscii >= Asc("A") And KeyAscii <= Asc("F") Then
  ElseIf KeyAscii >= Asc("a") And KeyAscii <= Asc("f") Then
     KeyAscii = Asc(UCase$(Chr$(KeyAscii))) ' upper case
  Else
     If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys vbTab
     ElseIf KeyAscii = Asc(vbBack) Then
        ' fine...
     Else
        KeyAscii = 0
     End If
  End If
End Sub

Private Sub hexRed_Validate(Cancel As Boolean)
' Validazione Input Hex(R)
  On Error GoTo err
  If Len(hexRed.Text) = 0 Then hexRed.Text = "00"
  initVals(0) = CLng(Hex2Dec(hexRed.Text))
  reFreshAll initVals(0), initVals(1), initVals(2)
err:
  If err.Number <> 0 Then Label10.Visible = True
End Sub

Private Sub ResetControls()
' Setta tutti i controlli
  On Error Resume Next
  Dim lColor  As Long
  Dim iIdx    As Integer
  Dim iRed    As Integer
  Dim iGreen  As Integer
  Dim iBlue   As Integer
  Dim sRed    As String
  Dim sGreen  As String
  Dim sBlue   As String
  Dim sHex    As String
  mbNoChange = True
' Verifica se e' un colore di sistema
' Restituisce -1 se non e' un colore di sistema
  lColor = ConvertToSysColor(mlCurColor)
' Converte i colori di sistema se possibile
' Un Long di &H80000000& o maggiore e' un numero negativo
' I colori di sistema in VB vanno da &H80000000& a &H80000018& (da -2147483648 a -2147483624)
' Tutti gli altri colori vanno da &H00000000& a &H00FFFFFF& (da 0 a 16777215).
  If lColor < -1 Then  'e' un colore di sistema
     txtVbValue.Text = "&H" & Hex$(lColor) & "&"
     lColor = GetSysColor(lColor And &HFF&)
  Else
     'Non e' un colore di sistema
      lColor = mlCurColor
  End If
' Mostra il colore nella Color box.
  lblColorBox.BackColor = lColor
' Estrazione dei valori RGB
  iRed = (lColor And &HFF&)
  iGreen = (lColor And &HFF00&) / &H100
  iBlue = (lColor And &HFF0000) / &H10000
  decRed.Text = CStr(iRed)
  decGreen = CStr(iGreen)
  decBlue = CStr(iBlue)
  mbNoChange = False
End Sub

Private Function ConvertToSysColor(ByVal lColor As Long) As Long
' Cerca il colore passato nei colori di sistema
  Dim lIdx As Long
  Dim sHex As String
  If lColor < 0 Then
    'E' un colore di sistema
     ConvertToSysColor = lColor
  Else
     For lIdx = 0 To 24
         If GetSysColor(lIdx) = lColor Then
           'Corrispondenza esatta
            sHex = Hex$(lIdx)
            If Len(sHex) < 2 Then
               sHex = "0" & sHex
            End If
            ConvertToSysColor = Val("&H800000" & sHex)
            Exit For
         End If
     Next
     If lIdx > 24 Then
       'Corrispondenza non trovata
        ConvertToSysColor = -1
     End If
  End If
End Function

Private Sub Timer1_Timer()
' Copia il colore nella Label
  lblColorBox.BackColor = spy_color
End Sub

Public Function spy_color() As Long
' Rileva i colori sotto il puntatore del mouse
  On Error Resume Next
  Dim CursorPos As PointAPI
  Dim rDC As Long, rPixel As Long
  Call GetCursorPos(CursorPos)
  rDC& = GetDC(0&)
  rPixel& = GetPixel(rDC&, CursorPos.x, CursorPos.y)
  Call ReleaseDC(0&, rDC&)
  spy_color& = rPixel&
End Function

Public Function spy_colorqb() As String
' Legge e codifica i colori qbasic
  Dim sColor As Long
  sColor& = picCurColor.BackColor
  Select Case sColor&
     Case QBColor(0):  spy_colorqb$ = "0"
     Case QBColor(1):  spy_colorqb$ = "1"
     Case QBColor(2):  spy_colorqb$ = "2"
     Case QBColor(3):  spy_colorqb$ = "3"
     Case QBColor(4):  spy_colorqb$ = "4"
     Case QBColor(5):  spy_colorqb$ = "5"
     Case QBColor(6):  spy_colorqb$ = "6"
     Case QBColor(7):  spy_colorqb$ = "7"
     Case QBColor(8):  spy_colorqb$ = "8"
     Case QBColor(9):  spy_colorqb$ = "9"
     Case QBColor(10): spy_colorqb$ = "10"
     Case QBColor(11): spy_colorqb$ = "11"
     Case QBColor(12): spy_colorqb$ = "12"
     Case QBColor(13): spy_colorqb$ = "13"
     Case QBColor(14): spy_colorqb$ = "14"
     Case QBColor(15): spy_colorqb$ = "15"
     Case Else:        spy_colorqb$ = "n/a"
  End Select
End Function

Public Function spy_colorvb() As String
' Legge e codifica le costanti Colore di vb
  Dim sColor As Long
  sColor& = picCurColor.BackColor
  Select Case sColor&
    Case vbBlack:   spy_colorvb$ = "vbBlack"
    Case vbRed:     spy_colorvb$ = "vbRed"
    Case vbGreen:   spy_colorvb$ = "vbGreen"
    Case vbYellow:  spy_colorvb$ = "vbYellow"
    Case vbBlue:    spy_colorvb$ = "vbBlue"
    Case vbMagenta: spy_colorvb$ = "vbMagenta"
    Case vbCyan:    spy_colorvb$ = "vbCyan"
    Case vbWhite:   spy_colorvb$ = "vbWhite"
    Case Else:      spy_colorvb$ = "n/a"
  End Select
End Function

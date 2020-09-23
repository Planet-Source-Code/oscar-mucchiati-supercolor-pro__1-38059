VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E5E5E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ColorPad Pro"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   2
      Left            =   2362
      Locked          =   -1  'True
      TabIndex        =   62
      Text            =   "&H0000FFFF"
      Top             =   5190
      Width           =   1050
   End
   Begin VB.PictureBox cMix 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   2
      Left            =   2362
      ScaleHeight     =   645
      ScaleWidth      =   1020
      TabIndex        =   61
      Top             =   4470
      Width           =   1050
   End
   Begin VB.TextBox txtMix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Index           =   0
      Left            =   4605
      Locked          =   -1  'True
      TabIndex        =   57
      Text            =   "[empty]"
      Top             =   5175
      Width           =   1050
   End
   Begin VB.PictureBox cMix 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   0
      Left            =   4605
      ScaleHeight     =   675
      ScaleWidth      =   1050
      TabIndex        =   56
      Top             =   4455
      Width           =   1050
      Begin VB.Shape rec 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Height          =   645
         Left            =   0
         Top             =   0
         Width           =   1030
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add Color"
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   180
         TabIndex        =   63
         Top             =   225
         Width           =   705
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   180
      Left            =   1215
      TabIndex        =   41
      Top             =   4950
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   318
      _Version        =   393216
      Arrows          =   65536
      Max             =   100
      Orientation     =   1179649
      Value           =   75
   End
   Begin VB.PictureBox cMix 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   1
      Left            =   120
      ScaleHeight     =   645
      ScaleWidth      =   1020
      TabIndex        =   38
      Top             =   4470
      Width           =   1050
   End
   Begin VB.TextBox txtMix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "&H000000FF"
      Top             =   5190
      Width           =   1050
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000BFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   6060
      ScaleHeight     =   645
      ScaleWidth      =   1545
      TabIndex        =   36
      Top             =   4470
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   6060
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "&H0000BFFF"
      Top             =   5190
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00EFE5E0&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   7110
      Picture         =   "frmMain.frx":57E2
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   33
      ToolTipText     =   "Press and keep pressed left mouse button to capture colors on screen."
      Top             =   6345
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00EFE5E0&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   7125
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   24
      ToolTipText     =   "Press and keep pressed left mouse button to capture colors on screen."
      Top             =   720
      Width           =   510
   End
   Begin VB.OptionButton optType 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E0&
      Caption         =   "&Hexadecimal"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   2685
      Width           =   1215
   End
   Begin VB.OptionButton optType 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E0&
      Caption         =   "&Web (HTML)"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   2445
      Width           =   1215
   End
   Begin VB.OptionButton optType 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E0&
      Caption         =   "&Visual Basic"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   2205
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optType 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E0&
      Caption         =   "&Long"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   1965
      Width           =   1215
   End
   Begin VB.PictureBox cL 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6750
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   2640
      Width           =   285
   End
   Begin VB.TextBox txtL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   6120
      TabIndex        =   17
      Text            =   "100"
      Top             =   2640
      Width           =   615
   End
   Begin VB.PictureBox cS 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6750
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   2280
      Width           =   285
   End
   Begin VB.TextBox txtS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   6120
      TabIndex        =   14
      Text            =   "100"
      Top             =   2280
      Width           =   615
   End
   Begin VB.PictureBox cH 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6750
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   1920
      Width           =   285
   End
   Begin VB.TextBox txtH 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   6120
      TabIndex        =   11
      Text            =   "359"
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox cB 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6750
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   9
      Top             =   1440
      Width           =   285
   End
   Begin VB.TextBox txtB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   6120
      TabIndex        =   8
      Text            =   "255"
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox cG 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6750
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   1080
      Width           =   285
   End
   Begin VB.TextBox txtG 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   6120
      TabIndex        =   5
      Text            =   "255"
      Top             =   1080
      Width           =   615
   End
   Begin VB.PictureBox cR 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6750
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   720
      Width           =   285
   End
   Begin VB.TextBox txtR 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   6120
      TabIndex        =   2
      Text            =   "255"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtColor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.PictureBox pctColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   120
      ScaleHeight     =   645
      ScaleWidth      =   1545
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.PictureBox pctTitlebar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawStyle       =   2  'Dot
      DrawWidth       =   16887
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   26
      Top             =   0
      Width           =   7725
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "© by OMAA.net"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   6030
         TabIndex        =   32
         Top             =   360
         Width           =   1065
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmMain.frx":5AEC
         Top             =   45
         Width           =   480
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SuperColor Pro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   780
         TabIndex        =   27
         Top             =   90
         Width           =   2115
      End
      Begin VB.Image imgClose 
         Height          =   480
         Left            =   7140
         Picture         =   "frmMain.frx":613F
         Top             =   45
         Width           =   480
      End
   End
   Begin MSComCtl2.FlatScrollBar hsR 
      Height          =   285
      Left            =   2700
      TabIndex        =   49
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      _Version        =   393216
      Arrows          =   65536
      LargeChange     =   10
      Max             =   255
      Orientation     =   1179649
   End
   Begin MSComCtl2.FlatScrollBar hsG 
      Height          =   285
      Left            =   2700
      TabIndex        =   50
      Top             =   1080
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      _Version        =   393216
      Arrows          =   65536
      LargeChange     =   10
      Max             =   255
      Orientation     =   1179649
   End
   Begin MSComCtl2.FlatScrollBar hsB 
      Height          =   285
      Left            =   2700
      TabIndex        =   51
      Top             =   1440
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      _Version        =   393216
      Arrows          =   65536
      LargeChange     =   10
      Max             =   255
      Orientation     =   1179649
   End
   Begin MSComCtl2.FlatScrollBar hsH 
      Height          =   285
      Left            =   2700
      TabIndex        =   52
      Top             =   1920
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      _Version        =   393216
      Arrows          =   65536
      LargeChange     =   10
      Max             =   359
      Orientation     =   1179649
   End
   Begin MSComCtl2.FlatScrollBar hsS 
      Height          =   285
      Left            =   2700
      TabIndex        =   53
      Top             =   2280
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      _Version        =   393216
      Arrows          =   65536
      LargeChange     =   10
      Max             =   100
      Orientation     =   1179649
   End
   Begin MSComCtl2.FlatScrollBar hsL 
      Height          =   285
      Left            =   2700
      TabIndex        =   54
      Top             =   2640
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      _Version        =   393216
      Arrows          =   65536
      LargeChange     =   10
      Max             =   100
      Orientation     =   1179649
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar2 
      Height          =   180
      Left            =   3450
      TabIndex        =   55
      Top             =   4935
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   318
      _Version        =   393216
      Enabled         =   0   'False
      Arrows          =   65536
      Max             =   100
      Orientation     =   1179649
      Value           =   75
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   3465
      TabIndex        =   64
      Top             =   4725
      Width           =   105
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3870
      TabIndex        =   60
      Top             =   4485
      Width           =   255
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "75%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   165
      Left            =   3855
      TabIndex        =   59
      Top             =   5175
      Width           =   300
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   3645
      TabIndex        =   58
      Top             =   4725
      Width           =   90
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   30
      X2              =   30
      Y1              =   261
      Y2              =   284
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   135
      Picture         =   "frmMain.frx":6E09
      Top             =   3975
      Width           =   240
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Addictive mixing (default)."
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   360
      TabIndex        =   48
      Top             =   5580
      Width           =   1890
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subtractive mixing."
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   2610
      TabIndex        =   47
      Top             =   5580
      Width           =   1380
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2415
      TabIndex        =   46
      Top             =   5580
      Width           =   105
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   45
      Top             =   5580
      Width           =   120
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   1425
      TabIndex        =   44
      Top             =   4740
      Width           =   90
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1260
      TabIndex        =   43
      Top             =   4710
      Width           =   120
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "75%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   165
      Left            =   1635
      TabIndex        =   42
      Top             =   5190
      Width           =   300
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5715
      TabIndex        =   40
      Top             =   4500
      Width           =   255
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1650
      TabIndex        =   39
      Top             =   4500
      Width           =   255
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      X1              =   8
      X2              =   504
      Y1              =   245
      Y2              =   245
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next version will allow you to mix different colors."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   525
      TabIndex        =   34
      Top             =   3885
      Width           =   3570
   End
   Begin VB.Image imgOpen 
      Height          =   105
      Left            =   6750
      Picture         =   "frmMain.frx":7218
      Top             =   3510
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image imgClosed 
      Height          =   105
      Left            =   6930
      Picture         =   "frmMain.frx":725C
      Top             =   3510
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image imgMixOpen 
      Height          =   105
      Left            =   7485
      Picture         =   "frmMain.frx":72A0
      Tag             =   "0"
      ToolTipText     =   "Open advanced mixing properties..."
      Top             =   3510
      Width           =   150
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   8
      X2              =   504
      Y1              =   203
      Y2              =   203
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Get the value from clipboard (if it's valid)."
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   960
      TabIndex        =   31
      Top             =   3405
      Width           =   2970
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CTRL + V"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   3405
      Width           =   750
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copies the value to clipboard (or you can double-click on the window color top-left)."
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   960
      TabIndex        =   29
      Top             =   3165
      Width           =   6015
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CTRL + C"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   3165
      Width           =   750
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   8
      X2              =   504
      Y1              =   122
      Y2              =   122
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Capture"
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   7088
      TabIndex        =   25
      Top             =   1200
      Width           =   585
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Luminance : "
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1800
      TabIndex        =   19
      Top             =   2685
      Width           =   900
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Saturation : "
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1800
      TabIndex        =   16
      Top             =   2325
      Width           =   900
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hue : "
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1800
      TabIndex        =   13
      Top             =   1965
      Width           =   900
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Blue : "
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1800
      TabIndex        =   10
      Top             =   1485
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Green : "
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1800
      TabIndex        =   7
      Top             =   1125
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Red : "
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1800
      TabIndex        =   4
      Top             =   765
      Width           =   900
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
        
Private outType As Integer
Private iR As Integer, iG As Integer, iB As Integer
Private iH As Integer, iSa As Integer, iL As Integer
Private cOut As New clsColors, bNoRefresh As Boolean

'*** SOME API CALLS & TYPES ***
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2


Private Sub cB_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub


Private Sub cG_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub


Private Sub cH_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub

Private Sub cL_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub


Private Sub cR_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub


Private Sub cS_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub


Private Sub FlatScrollBar1_Change()
    FlatScrollBar1.Value = 75
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    'Check if CTRL+C or CTRL+V has been pressed (and then relased)
    
    If Shift <> 2 Then Exit Sub 'CTRL must be pressed
    
    Dim sText As String
    
    Select Case KeyCode
        
        Case 67: ' CTRL+C pressed
            SetClipboard
            
        Case 86: ' CTRL+V pressed
            sText = LCase(Trim(Clipboard.GetText(vbCFText)))
            If Len(sText) <= 0 Then Exit Sub
            
            SetValue sText
            Beep
            
    End Select
    
    KeyCode = 0
    
End Sub

Private Sub HIDE_TITLEBAR()

    'This subroutine hides the titlebar of the form
    'It simply changes the class bits of the window
    
    Dim lngResult  As Long
    Dim hwnd As Long
    
    hwnd = Me.hwnd                                      'HWND to this form
    
    lngResult = GetWindowLong(hwnd, -16)                'Get the class bits
    SetWindowLong hwnd, -16, lngResult And Not &HC00000 'Modify the class bits with logic operators

End Sub

Private Sub Form_Load()
    
    'Let's do some preparing work
    
    HIDE_TITLEBAR
    
    Label13.Caption = "Next version will allow you to mix different colors." & vbCr & "EXAMPLE:"
    Me.Height = 3680
    
    hsR_Change
    hsG_Change
    hsB_Change
    
    outType = 1
    
    Update 0
    
    Picture2.MouseIcon = Picture1.Picture
    Picture2.Picture = Picture1.Picture
    
    doGradient pctTitlebar.hDC, 100, 0, pctTitlebar.ScaleWidth, pctTitlebar.ScaleHeight
    
    lblCopyright.ForeColor = RGB(0, 0, 128)
    
End Sub
Private Sub FormDrag(TheForm_HWND As Long)
    
    'A simple way to move forms with no titlebar around
    
    ReleaseCapture
    Call SendMessage(TheForm_HWND, &HA1, 2, 0&)
    
End Sub

Private Sub MeNotOnTop()
    
    'Set the form to normal state
    SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
End Sub

Private Sub MeOnTop()
    
    'Set the form to TOPMOST state: this window will be shown over any other
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
End Sub

Private Function GetCursorColor() As Long
    
    'This function returs the color value (long) of the pixel under the cursor
    
    Dim Pnt As POINTAPI
    Dim hDC_screen As Long
    
    'Get current cursor position
    GetCursorPos Pnt
    
    'Get the screen DC
    hDC_screen = GetDC(0)
    
    'Get the color value (long)
    GetCursorColor = GetPixel(hDC_screen, Pnt.x, Pnt.y)

End Function

Private Sub Form_Resize()
    
    On Error Resume Next
    Me.Cls
    Me.Line (0, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight - 1), RGB(128, 128, 228), B
    
End Sub

Private Sub hsB_Change()
    cOut.Blu = hsB.Value
    Update 3
End Sub

Private Sub hsB_Scroll()
    cOut.Blu = hsB.Value
    Update 3
End Sub


Private Sub hsG_Change()
    cOut.Green = hsG.Value
    Update 2
End Sub

Private Sub hsG_Scroll()
    cOut.Green = hsG.Value
    Update 2
End Sub


Private Sub hsH_Change()
    cOut.h = hsH.Value
    Update 4
End Sub

Private Sub hsH_Scroll()
    cOut.h = hsH.Value
    Update 4
End Sub


Private Sub hsL_Change()
    cOut.v = hsL.Value
    Update 6
End Sub

Private Sub hsL_Scroll()
    cOut.v = hsL.Value
    Update 6
End Sub


Private Sub hsR_Change()
    cOut.Red = hsR.Value
    Update 1
End Sub


Private Sub hsR_Scroll()
    cOut.Red = hsR.Value
    Update 1
End Sub


Private Sub Update(ByVal Index As Integer)

    'This routine refreshes the values of scrollbars and set the new color
    
    Dim sOut As String
    
    If bNoRefresh = True Then Exit Sub  'Do not refresh twice
    bNoRefresh = True                   'Do not refresh twice
    
    If Index > 0 Then
        If Index <= 3 Then
            hsH.Value = cOut.h
            hsS.Value = cOut.s
            hsL.Value = cOut.v
        Else
            hsR.Value = cOut.Red
            hsG.Value = cOut.Green
            hsB.Value = cOut.Blu
        End If
    Else
        hsR.Value = cOut.Red
        hsG.Value = cOut.Green
        hsB.Value = cOut.Blu
    End If

    'Gets the new values from the class
    iR = cOut.Red
    iG = cOut.Green
    iB = cOut.Blu
    iH = cOut.h
    iSa = cOut.s
    iL = cOut.v

    'Refreshes controls
    cR.BackColor = RGB(iR, 0, 0)
        txtR.Text = iR
    cG.BackColor = RGB(0, iG, 0)
        txtG.Text = iG
    cB.BackColor = RGB(0, 0, iB)
        txtB.Text = iB
    
    cH.BackColor = RGB(iH / 360 * 255, iH / 360 * 255, iH / 360 * 255)
        txtH.Text = iH
    cS.BackColor = RGB(iSa, iSa, iSa)
        txtS.Text = iSa
    cL.BackColor = RGB(iL, iL, iL)
        txtL.Text = iL
    
    'Set the color
    pctColor.BackColor = cOut.RGB_long
    
    'Select the proper text output
    Select Case outType
        Case 0: sOut = cOut.RGB_long
        Case 1: sOut = cOut.VBColor
        Case 2: sOut = cOut.WebColor
        Case 3: sOut = "h" & cOut.HexColor
    End Select
    txtColor.Text = sOut
    
    bNoRefresh = False  'Refresh completed
    
End Sub
Private Function MinMax(ByVal v As Integer, ByVal iMin As Integer, ByVal iMax As Integer) As Integer

    'The function returns the value between iMin and iMax
    
    If v < iMin Then v = iMin
    If v > iMax Then v = iMax
    
    MinMax = v
    
End Function

Private Sub hsS_Change()
    cOut.s = hsS.Value
    Update 5
End Sub

Private Sub hsS_Scroll()
    cOut.s = hsS.Value
    Update 5
End Sub



Private Sub imgClose_Click()
    Unload Me
End Sub


Private Sub imgMixOpen_Click()

    'Clicking on the control will open/close the advanced configuring controls
    
    If imgMixOpen.Tag = "0" Then
        Set imgMixOpen.Picture = imgOpen.Picture
        imgMixOpen.Tag = "1"
        Me.Height = Me.Height + 2150
        Me.Top = Me.Top - 1075
    Else
        Set imgMixOpen.Picture = imgClosed.Picture
        imgMixOpen.Tag = "0"
        Me.Height = Me.Height - 2150
        Me.Top = Me.Top + 1075
    End If
    
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    pctTitlebar_MouseMove Button, Shift, x, y
End Sub


Private Sub optType_Click(Index As Integer)
    outType = Index
    Update 0
End Sub

Private Sub optType_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub


Private Sub pctColor_DblClick()
    
    SetClipboard
    
End Sub

Private Sub pctColor_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub


Private Sub pctTitlebar_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub

Private Sub pctTitlebar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 1 Then Exit Sub
    FormDrag Me.hwnd

End Sub


Private Sub Picture2_KeyUp(KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 1 Then Exit Sub
    
    'Start to capture the cursor position & color. Set the correct mouse icon (a cross).
    ReleaseCapture
    SetCapture Picture2.hwnd
    Picture2.BackColor = RGB(255, 0, 0)
    
    Picture2.Picture = Me.Picture
    Picture2.MousePointer = 99
    
End Sub


Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 1 Then Exit Sub
    
    'Capturing the cursor position & color...
    cOut.RGB_long = GetCursorColor
    
    hsR.Value = cOut.Red
    hsG.Value = cOut.Green
    hsB.Value = cOut.Blu
    
    Update 0
    
    SetCapture Picture2.hwnd
    
End Sub


Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 1 Then Exit Sub
    
    'Ends capturing the cursor position & color. Set the icon back to the default (arrow)
    ReleaseCapture
    Picture2.BackColor = &HC0C0C0
    
    Picture2.MousePointer = 0
    Picture2.Picture = Picture1.Picture
    
End Sub


Private Sub txtB_LostFocus()
    hsB.Value = MinMax(Val(txtB.Text), 0, 255)
End Sub


Private Sub txtColor_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then    'ENTER key pressed: check if the text entered is correct
        SetValue txtColor.Text
        Update 0
    End If
    
End Sub


Private Sub txtG_LostFocus()
    hsG.Value = MinMax(Val(txtG.Text), 0, 255)
End Sub


Private Sub txtH_LostFocus()
    hsH.Value = MinMax(Val(txtH.Text), 0, 359)
End Sub


Private Sub txtL_LostFocus()
    hsL.Value = MinMax(Val(txtL.Text), 0, 100)
End Sub


Private Sub txtR_LostFocus()
    hsR.Value = MinMax(Val(txtR.Text), 0, 255)
End Sub


Private Sub txtS_LostFocus()
    hsS.Value = MinMax(Val(txtS.Text), 0, 100)
End Sub



Private Sub SetValue(ByVal sText As String)
        
    'This subroutine gets a string and check if the value in it is
    'an allowed "color value":
    
    On Error Resume Next
    
    sText = LCase(Trim(sText))
    
    Select Case Left(sText, 1)
        Case "&" 'Starts with a "&": VB color
            'cOut.VBColor = stet
            optType(1).Value = True
            If Mid(sText, 2, 1) = "h" Then
                cOut.RGB_long = Val(sText)
            End If
            
        Case "#" 'Starts with a "#": Web color
            optType(2).Value = True
            cOut.WebColor = sText
            
        Case "h" 'Starts with "h": Hex color
            sText = Right(sText, Len(sText) - 1)
            If IsNumeric(sText) Then
                optType(3).Value = True
                cOut.RGB_long = Val("&h" & sText)
            End If
            
        Case Else 'Can be a LONG value (if numeric)
            If Not IsNumeric(sText) Then Exit Sub
            optType(0).Value = True
            cOut.RGB_long = Val(sText)
            
    End Select
    
    bNoRefresh = False  'Force the refresh
    Update 0            'Refresh the color

End Sub

Private Sub SetClipboard()
    
    'Set clipboard text
    On Error Resume Next
    
    Dim sText As String
    
    Clipboard.Clear
    sText = txtColor.Text
    Clipboard.SetText sText, vbCFText
    Beep

End Sub

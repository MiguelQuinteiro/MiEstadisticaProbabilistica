VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ConProgressBar 
   BackColor       =   &H00FFFFC0&
   Caption         =   "STAT LAB"
   ClientHeight    =   9660
   ClientLeft      =   3555
   ClientTop       =   2865
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Papyrus"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "StatLab01.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9660
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   9360
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame6 
      Caption         =   "SOLUCIONES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   33
      Top             =   4800
      Width           =   8895
      Begin VB.TextBox lblProb_F 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox lblProb_E 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox lblProb_D 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox lblProb_C 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox lblProb_B 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox lblProb_A 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox lblProb_A2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox lblProb_B2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox lblProb_C2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox lblProb_D2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox lblProb_E2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox lblProb_F2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3240
         Width           =   3255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   375
         Left            =   5400
         OleObjectBlob   =   "StatLab01.frx":0CFA
         TabIndex        =   34
         Top             =   360
         Width           =   3255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   375
         Left            =   3720
         OleObjectBlob   =   "StatLab01.frx":0D92
         TabIndex        =   35
         Top             =   360
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "StatLab01.frx":0E02
         TabIndex        =   36
         Top             =   360
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "StatLab01.frx":0E88
         TabIndex        =   49
         Top             =   840
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "StatLab01.frx":0F14
         TabIndex        =   50
         Top             =   1320
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "StatLab01.frx":0FA4
         TabIndex        =   51
         Top             =   1800
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "StatLab01.frx":1036
         TabIndex        =   52
         Top             =   2280
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "StatLab01.frx":10C2
         TabIndex        =   53
         Top             =   2760
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "StatLab01.frx":114A
         TabIndex        =   54
         Top             =   3240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "DATOS INICIALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtExtracciones 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   56
         Text            =   "10000"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtAzules 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Text            =   "9"
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtBlancas 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Text            =   "3"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtRojas 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Text            =   "8"
         Top             =   1080
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   735
         Left            =   960
         OleObjectBlob   =   "StatLab01.frx":11E2
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   735
         Left            =   960
         OleObjectBlob   =   "StatLab01.frx":124A
         TabIndex        =   13
         Top             =   1920
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   735
         Left            =   960
         OleObjectBlob   =   "StatLab01.frx":12B6
         TabIndex        =   14
         Top             =   2760
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "StatLab01.frx":1320
         TabIndex        =   55
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "RESULTADO DE LAS EXTRACCIONES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   6735
      Begin VB.Frame Frame1 
         Caption         =   "Bola 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   6375
         Begin VB.TextBox lbl_B01 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   1320
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   " "
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox lbl_B01 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   3000
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   " "
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox lbl_B01 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   4680
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   " "
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Bola 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   6375
         Begin VB.TextBox lbl_B02 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   1320
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   " "
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox lbl_B02 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   3000
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   " "
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox lbl_B02 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   4680
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   " "
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Bola 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   6375
         Begin VB.TextBox lbl_B03 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   1320
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   " "
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox lbl_B03 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   3000
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   " "
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox lbl_B03 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   4680
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   " "
            Top             =   240
            Width           =   1575
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   375
         Left            =   1440
         OleObjectBlob   =   "StatLab01.frx":139E
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   375
         Left            =   3120
         OleObjectBlob   =   "StatLab01.frx":1406
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   375
         Left            =   4800
         OleObjectBlob   =   "StatLab01.frx":1472
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblRojas 
         Height          =   375
         Left            =   1440
         OleObjectBlob   =   "StatLab01.frx":14DC
         TabIndex        =   27
         Top             =   3960
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   375
         Left            =   1440
         OleObjectBlob   =   "StatLab01.frx":153A
         TabIndex        =   28
         Top             =   3600
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   375
         Left            =   3120
         OleObjectBlob   =   "StatLab01.frx":15AE
         TabIndex        =   29
         Top             =   3600
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   375
         Left            =   4800
         OleObjectBlob   =   "StatLab01.frx":1626
         TabIndex        =   30
         Top             =   3600
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblBlancas 
         Height          =   375
         Left            =   3120
         OleObjectBlob   =   "StatLab01.frx":169C
         TabIndex        =   31
         Top             =   3960
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblAzules 
         Height          =   375
         Left            =   4800
         OleObjectBlob   =   "StatLab01.frx":16FA
         TabIndex        =   32
         Top             =   3960
         Width           =   1575
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8400
      OleObjectBlob   =   "StatLab01.frx":1758
      Top             =   0
   End
   Begin VB.CommandButton cmdLimpiar 
      BackColor       =   &H00FF8080&
      Caption         =   "LIMPIAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MaskColor       =   &H0000C0C0&
      TabIndex        =   1
      Top             =   8640
      Width           =   4335
   End
   Begin VB.CommandButton cmdExtraer 
      BackColor       =   &H00FFC0C0&
      Caption         =   "EXTRAER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MaskColor       =   &H0000C0C0&
      TabIndex        =   0
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00FF8080&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      MaskColor       =   &H0000C0C0&
      TabIndex        =   2
      Top             =   8640
      Width           =   4335
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "Opciones"
      Begin VB.Menu mnuSkin 
         Caption         =   "Skin"
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "B-Studio"
            Index           =   0
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Comander"
            Index           =   1
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Cool Breeze"
            Index           =   2
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Copper"
            Index           =   3
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Corona"
            Index           =   4
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "DogmaX"
            Index           =   5
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Droid"
            Index           =   6
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Green"
            Index           =   7
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Gris"
            Index           =   8
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "KOZ"
            Index           =   9
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "LaST v1-2"
            Index           =   10
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "LongHorn"
            Index           =   11
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Mac"
            Index           =   12
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Media"
            Index           =   13
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Messenger"
            Index           =   14
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Metallic"
            Index           =   15
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "MMD"
            Index           =   16
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Neo"
            Index           =   17
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Office"
            Index           =   18
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Office2007"
            Index           =   19
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Orange_Graf"
            Index           =   20
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Paper"
            Index           =   21
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "SknR"
            Index           =   22
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "SoftCrystal"
            Index           =   23
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "St"
            Index           =   24
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "TopSecret"
            Index           =   25
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Tp"
            Index           =   26
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Web-II"
            Index           =   27
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Winamp 5"
            Index           =   28
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Zega"
            Index           =   29
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Zhelezo"
            Index           =   30
         End
         Begin VB.Menu mnuCambiarSkin 
            Caption         =   "Zippo"
            Index           =   31
         End
      End
   End
End
Attribute VB_Name = "ConProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *****************************************************************************
' * PROYECTO   : SISTEMA DE ESTADÍSTICA PROBABILÍSTICA
' * FORMULARIO : Formulario Principal
' * AUTORES    : Miguel Quinteiro
' * FECHA      : 25 de Abril de 2008
' * ***************************************************************************


'AL CARGAR EL FORMULARIO
Private Sub Form_Load()
  Open "Skin.txt" For Input As #1
  Line Input #1, A$
  Close #1
  Mi_Skin = "\Skins\" + A$
  Aplicar_skin Me

  ProgressBar1.Align = vbAlignBottom
  ProgressBar1.Visible = False
End Sub


'BOTÓN DE COMANDO PARA EXTRAER
Private Sub cmdExtraer_Click()
  Dim Counter As Integer
  Dim Workarea(10000) As String
  ProgressBar1.Min = LBound(Workarea)
  ProgressBar1.Max = UBound(Workarea)
  ProgressBar1.Visible = True

  Randomize (timmer)

  'Establece Min como valor de Value.
  ProgressBar1.Value = ProgressBar1.Min

  Rojo = 0: Rojo01 = 0: Rojo02 = 0: Rojo03 = 0
  Blanco = 0: Blanco01 = 0: Blanco02 = 0: Blanco03 = 0
  Azul = 0: Azul01 = 0: Azul02 = 0: Azul03 = 0

  lbl_B01(0) = Rojo01: lbl_B01(1) = Rojo02: lbl_B01(2) = Rojo03:
  lbl_B02(0) = Blanco01: lbl_B02(1) = Blanco02: lbl_B02(2) = Blanco03:
  lbl_B03(0) = Azul01: lbl_B03(1) = Azul02: lbl_B03(2) = Azul03

  Problema_A = 0
  lblProb_A = Problema_A

  VaRojas = Val(txtRojas)
  VaBlancas = Val(txtBlancas)
  VaAzules = Val(txtAzules)
  VaTotal = VaRojas + VaBlancas + VaAzules

  vaExtracciones = Val(txtExtracciones)

  For casos = 1 To vaExtracciones

    'TRES BOLAS DIFERENTES
    primera = 0
    segunda = 0
    tercera = 0
    Do Until (primera <> segunda And segunda <> tercera And tercera <> primera)
      primera = Int(Rnd(1) * (VaTotal) + 1)
      segunda = Int(Rnd(2) * (VaTotal) + 1)
      tercera = Int(Rnd(3) * (VaTotal) + 1)
    Loop
    '***********************************************************
    division = vaExtracciones / 10000
    Counter = casos / division
    '-------------------------------------------------------
    Workarea(Counter) = "Valor inicial"  '& Counter
    ProgressBar1.Value = Counter
    '--------------------------------------------------------
    '     Debug.Print Counter
    '****************************************************************

    'PRIMERA BOLA EXTRAIDA
    If (primera > VaRojas + VaBlancas) Then
      Azul = Azul + 1
      Azul01 = Azul01 + 1
      Bola01 = "A"
    End If
    If (primera <= VaRojas) Then
      Rojo = Rojo + 1
      Rojo01 = Rojo01 + 1
      Bola01 = "R"
    End If
    If (primera > VaRojas) And (primera <= VaRojas + VaBlancas) Then
      Blanco = Blanco + 1
      Blanco01 = Blanco01 + 1
      Bola01 = "B"
    End If

    'SEGUNDA BOLA EXTRAIDA
    If (segunda <= VaRojas) Then
      Rojo = Rojo + 1
      Rojo02 = Rojo02 + 1
      Bola02 = "R"
    End If
    If (segunda > VaRojas) And (segunda <= VaRojas + VaBlancas) Then
      Blanco = Blanco + 1
      Blanco02 = Blanco02 + 1
      Bola02 = "B"
    End If
    If (segunda > VaRojas + VaBlancas) Then
      Azul = Azul + 1
      Azul02 = Azul02 + 1
      Bola02 = "A"
    End If

    'TERCERA BOLA EXTRAIDA
    If (tercera > VaRojas + VaBlancas) Then
      Azul = Azul + 1
      Azul03 = Azul03 + 1
      Bola03 = "A"
    End If
    If (tercera <= VaRojas) Then
      Rojo = Rojo + 1
      Rojo03 = Rojo03 + 1
      Bola03 = "R"
    End If
    If (tercera > VaRojas) And (tercera <= VaRojas + VaBlancas) Then
      Blanco = Blanco + 1
      Blanco03 = Blanco03 + 1
      Bola03 = "B"
    End If

    'Resolución de Problemas
    ' A)
    If Bola01 = "R" And Bola02 = "R" And Bola03 = "R" Then Problema_A = Problema_A + 1
    ' B)
    If Bola01 = "B" And Bola02 = "B" And Bola03 = "B" Then Problema_B = Problema_B + 1
    'C)
    If Bola01 = "B" And Bola02 = "R" And Bola03 = "R" Then Problema_C = Problema_C + 1
    If Bola01 = "R" And Bola02 = "B" And Bola03 = "R" Then Problema_C = Problema_C + 1
    If Bola01 = "R" And Bola02 = "R" And Bola03 = "B" Then Problema_C = Problema_C + 1
    ' D)
    If Bola01 = "B" Or Bola02 = "B" Or Bola03 = "B" Then Problema_D = Problema_D + 1
    ' E)
    If Bola01 <> Bola02 And Bola01 <> Bola03 And Bola02 <> Bola03 Then Problema_E = Problema_E + 1
    ' F)
    If Bola01 = "R" And Bola02 = "B" And Bola03 = "A" Then Problema_F = Problema_F + 1

    ' Cálculo Empírico de la Probabilidad
    Problema_A2 = Problema_A / vaExtracciones
    Problema_B2 = Problema_B / vaExtracciones
    Problema_C2 = Problema_C / vaExtracciones
    Problema_D2 = Problema_D / vaExtracciones
    Problema_E2 = Problema_E / vaExtracciones
    Problema_F2 = Problema_F / vaExtracciones
  Next casos

  ' Impresión de Resultados

  ' Primera Bola
  lbl_B01(0) = Rojo01
  lbl_B01(1) = Blanco01
  lbl_B01(2) = Azul01

  ' Segunda Bola
  lbl_B02(0) = Rojo02
  lbl_B02(1) = Blanco02
  lbl_B02(2) = Azul02

  ' Tercera Bola
  lbl_B03(0) = Rojo03
  lbl_B03(1) = Blanco03
  lbl_B03(2) = Azul03

  ' Totales por Bolas
  lblRojas = Rojo
  lblBlancas = Blanco
  lblAzules = Azul

  ' Casos por Problemas
  lblProb_A = Problema_A
  lblProb_B = Problema_B
  lblProb_C = Problema_C
  lblProb_D = Problema_D
  lblProb_E = Problema_E
  lblProb_F = Problema_F

  ' Probabilidad por Problemas
  lblProb_A2 = Problema_A2
  lblProb_B2 = Problema_B2
  lblProb_C2 = Problema_C2
  lblProb_D2 = Problema_D2
  lblProb_E2 = Problema_E2
  lblProb_F2 = Problema_F2

  ProgressBar1.Visible = False
  ProgressBar1.Value = ProgressBar1.Min
End Sub


'BOTÓN DE COMANDO PARA LIMPIAR
Private Sub cmdLimpiar_Click()
' Borra todas las salidas

' Primera Bola
  lbl_B01(0) = ""
  lbl_B01(1) = ""
  lbl_B01(2) = ""

  ' Segunda Bola
  lbl_B02(0) = ""
  lbl_B02(1) = ""
  lbl_B02(2) = ""

  ' Tercera Bola
  lbl_B03(0) = ""
  lbl_B03(1) = ""
  lbl_B03(2) = ""

  ' Totales por Bolas
  lblRojas = ""
  lblBlancas = ""
  lblAzules = ""

  ' Casos por Problemas
  lblProb_A = ""
  lblProb_B = ""
  lblProb_C = ""
  lblProb_D = ""
  lblProb_E = ""
  lblProb_F = ""

  ' Probabilidad por Problemas
  lblProb_A2 = ""
  lblProb_B2 = ""
  lblProb_C2 = ""
  lblProb_D2 = ""
  lblProb_E2 = ""
  lblProb_F2 = ""

  txtExtracciones.SetFocus
End Sub

'BOTÓN DE COMANDO PARA SALIR
Private Sub cmdSalir_Click()
  End
End Sub


'MENÚ DE OPCIÓN PARA CAMBIER DE SKIN
Private Sub MNUCambiarSkin_Click(Index As Integer)
  Select Case Index
  Case Is = 0
    Mi_Skin = "\Skins\B-Studio.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "B-Studio.skn"
    Close #1

  Case Is = 1
    Mi_Skin = "\Skins\Comander.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Comander.skn"
    Close #1

  Case Is = 2
    Mi_Skin = "\Skins\Cool Breeze.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Cool Breeze.skn"
    Close #1

  Case Is = 3
    Mi_Skin = "\Skins\Copper.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Copper.skn"
    Close #1

  Case Is = 4
    Mi_Skin = "\Skins\Corona.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Corona.skn"
    Close #1

  Case Is = 5
    Mi_Skin = "\Skins\DogmaX.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "DogmaX.skn"
    Close #1

  Case Is = 6
    Mi_Skin = "\Skins\Droid.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Droid.skn"
    Close #1

  Case Is = 7
    Mi_Skin = "\Skins\Green.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Green.skn"
    Close #1

  Case Is = 8
    Mi_Skin = "\Skins\Gris.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Gris.skn"
    Close #1

  Case Is = 9
    Mi_Skin = "\Skins\KOZ.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "KOZ.skn"
    Close #1

  Case Is = 10
    Mi_Skin = "\Skins\LaST v1-2.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "LaST v1-2.skn"
    Close #1

  Case Is = 11
    Mi_Skin = "\Skins\LongHorn.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "LongHorn.skn"
    Close #1

  Case Is = 12
    Mi_Skin = "\Skins\Mac.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "green.skn"
    Close #1

  Case Is = 13
    Mi_Skin = "\Skins\Media.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Media.skn"
    Close #1

  Case Is = 14
    Mi_Skin = "\Skins\Messenger.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Messenger.skn"
    Close #1

  Case Is = 15
    Mi_Skin = "\Skins\Metallic.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Metallic.skn"
    Close #1

  Case Is = 16
    Mi_Skin = "\Skins\MMD.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "MMD.skn"
    Close #1

  Case Is = 17
    Mi_Skin = "\Skins\Neo.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Neo.skn"
    Close #1

  Case Is = 18
    Mi_Skin = "\Skins\Office.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Office.skn"
    Close #1

  Case Is = 19
    Mi_Skin = "\Skins\Office2007.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Office2007.skn"
    Close #1

  Case Is = 20
    Mi_Skin = "\Skins\Orange_Graf.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Orange_Graf.skn"
    Close #1

  Case Is = 21
    Mi_Skin = "\Skins\Paper.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Paper.skn"
    Close #1

  Case Is = 22
    Mi_Skin = "\Skins\SknR.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "SknR.skn"
    Close #1

  Case Is = 23
    Mi_Skin = "\Skins\SoftCrystal.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "SoftCrystal.skn"
    Close #1

  Case Is = 24
    Mi_Skin = "\Skins\St.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "St.skn"
    Close #1

  Case Is = 25
    Mi_Skin = "\Skins\TopSecret.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "TopSecret.skn"
    Close #1

  Case Is = 26
    Mi_Skin = "\Skins\Tp.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Tp.skn"
    Close #1

  Case Is = 27
    Mi_Skin = "\Skins\Web-II.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Web-II.skn"
    Close #1

  Case Is = 28
    Mi_Skin = "\Skins\Winamp 5.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Winamp 5.skn"
    Close #1

  Case Is = 29
    Mi_Skin = "\Skins\Zega.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Zega.skn"
    Close #1

  Case Is = 30
    Mi_Skin = "\Skins\Zhelezo.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Zhelezo.skn"
    Close #1

  Case Is = 31
    Mi_Skin = "\Skins\Zippo.skn"
    Skin1.LoadSkin App.Path & Mi_Skin
    Skin1.ApplySkin ConProgressBar.hWnd
    Open "Skin.txt" For Output As #1
    Print #1, "Zippo.skn"
    Close #1

  End Select
End Sub

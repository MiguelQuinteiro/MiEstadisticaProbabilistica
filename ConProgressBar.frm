VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form ConProgressBar 
   Caption         =   "Form2"
   ClientHeight    =   6195
   ClientLeft      =   3555
   ClientTop       =   2565
   ClientWidth     =   8235
   LinkTopic       =   "Form2"
   ScaleHeight     =   6195
   ScaleWidth      =   8235
   Begin VB.TextBox lblProb_F2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4560
      TabIndex        =   46
      Text            =   " "
      Top             =   5400
      Width           =   3495
   End
   Begin VB.TextBox lblProb_E2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4560
      TabIndex        =   45
      Text            =   " "
      Top             =   4920
      Width           =   3495
   End
   Begin VB.TextBox lblProb_D2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4560
      TabIndex        =   44
      Text            =   " "
      Top             =   4440
      Width           =   3495
   End
   Begin VB.TextBox lblProb_C2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4560
      TabIndex        =   43
      Text            =   " "
      Top             =   3960
      Width           =   3495
   End
   Begin VB.TextBox lblProb_B2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4560
      TabIndex        =   42
      Text            =   " "
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox lblProb_A2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4560
      TabIndex        =   41
      Text            =   " "
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox lblProb_A 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3000
      TabIndex        =   40
      Text            =   " "
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox lblProb_B 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3000
      TabIndex        =   39
      Text            =   " "
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox lblProb_C 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3000
      TabIndex        =   38
      Text            =   " "
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox lblProb_D 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3000
      TabIndex        =   37
      Text            =   " "
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox lblProb_E 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3000
      TabIndex        =   36
      Text            =   " "
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox lblProb_F 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3000
      TabIndex        =   35
      Text            =   " "
      Top             =   5400
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   5880
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
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
      Left            =   240
      TabIndex        =   14
      Text            =   "10000"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Height          =   495
      Left            =   1920
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtRojas 
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
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Text            =   "8"
      Top             =   960
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
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Text            =   "3"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtAzules 
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
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Text            =   "9"
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox lbl_B01 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   3240
      TabIndex        =   9
      Text            =   " "
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox lbl_B01 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   8
      Text            =   " "
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox lbl_B01 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   6600
      TabIndex        =   7
      Text            =   " "
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox lbl_B02 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   3240
      TabIndex        =   6
      Text            =   " "
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox lbl_B02 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   5
      Text            =   " "
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox lbl_B02 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   6600
      TabIndex        =   4
      Text            =   " "
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox lbl_B03 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   3240
      TabIndex        =   3
      Text            =   " "
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox lbl_B03 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   2
      Text            =   " "
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox lbl_B03 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   6600
      TabIndex        =   1
      Text            =   " "
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "PROBABILIDAD EMPÍRICA  =  f/N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   54
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label16 
      Caption         =   "CASOS (f)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   53
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "a) LAS TRES SEAN ROJAS:"
      Height          =   255
      Left            =   240
      TabIndex        =   52
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label11 
      Caption         =   "b) LAS TRES SEAN BLANCAS:"
      Height          =   255
      Left            =   240
      TabIndex        =   51
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label12 
      Caption         =   "c) DOS ROJAS Y UNA BLANCA:"
      Height          =   255
      Left            =   240
      TabIndex        =   50
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label13 
      Caption         =   "d) AL MENOS UNA BLANCA:"
      Height          =   255
      Left            =   240
      TabIndex        =   49
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Label Label14 
      Caption         =   "e) UNA DE CADA COLOR:"
      Height          =   255
      Left            =   240
      TabIndex        =   48
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label15 
      Caption         =   "f) ORDEN: ROJA, BLANCA, AZUL:"
      Height          =   255
      Left            =   240
      TabIndex        =   47
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Blancas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   33
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Azules"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   32
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Extracciones (N)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   " Bolas Inicial  COLOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblRojas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   29
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblBlancas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      TabIndex        =   28
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblAzules 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6600
      TabIndex        =   27
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Total Azules"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6720
      TabIndex        =   26
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Total Blancas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   25
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Total Rojas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   24
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Rojas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   23
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bola 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bola 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bola 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Rojas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   18
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Blancas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   17
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Azules"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   7080
      TabIndex        =   16
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Problemas Propuestos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   2535
   End
End
Attribute VB_Name = "ConProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

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
      Workarea(Counter) = "Valor inicial" '& Counter
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

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
   ProgressBar1.Align = vbAlignBottom
   ProgressBar1.Visible = False
   Command1.Caption = "EXTRAER"
End Sub



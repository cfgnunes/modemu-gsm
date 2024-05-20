VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Principal 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ModEmuGSM - Emulador de Modem GSM"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   ForeColor       =   &H00000000&
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   542
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_mensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   9
      Left            =   5640
      TabIndex        =   58
      TabStop         =   0   'False
      Text            =   "Nivel: 90% Bomba1: ON Bomba2: ON"
      Top             =   5040
      Width           =   3135
   End
   Begin VB.CheckBox chk_alocacao 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   8520
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   5400
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txt_remetente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   9
      Left            =   5640
      TabIndex        =   56
      TabStop         =   0   'False
      Text            =   "+553412341237"
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox txt_mensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   8
      Left            =   5640
      TabIndex        =   53
      TabStop         =   0   'False
      Text            =   "Nivel: 90% Bomba1: ON Bomba2: ON"
      Top             =   4080
      Width           =   3135
   End
   Begin VB.CheckBox chk_alocacao 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   8520
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   4440
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txt_remetente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   8
      Left            =   5640
      TabIndex        =   51
      TabStop         =   0   'False
      Text            =   "+553412341236"
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox txt_mensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   7
      Left            =   5640
      TabIndex        =   48
      TabStop         =   0   'False
      Text            =   "Nivel: 90% Bomba1: ON Bomba2: ON"
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CheckBox chk_alocacao 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   8520
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   3480
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txt_remetente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   7
      Left            =   5640
      TabIndex        =   46
      TabStop         =   0   'False
      Text            =   "+553412341235"
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox txt_mensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   5640
      TabIndex        =   43
      TabStop         =   0   'False
      Text            =   "Nivel: 90% Bomba1: ON Bomba2: ON"
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CheckBox chk_alocacao 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   8520
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2520
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txt_remetente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   5640
      TabIndex        =   41
      TabStop         =   0   'False
      Text            =   "+553412341234"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox txt_mensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   5640
      TabIndex        =   38
      TabStop         =   0   'False
      Text            =   "Nivel: 100% Bomba1: OFF Bomba2: ON"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.CheckBox chk_alocacao 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   8520
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1560
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txt_remetente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   5640
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "+553412345673"
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox txt_remetente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   33
      TabStop         =   0   'False
      Text            =   "+553491036588"
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CheckBox chk_alocacao 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   4080
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5400
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txt_mensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   31
      TabStop         =   0   'False
      Text            =   "Nivel: 80% Bomba1: ON Bomba2: OFF"
      Top             =   5040
      Width           =   3135
   End
   Begin VB.TextBox txt_mensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Nivel: 60% Bomba1: OFF Bomba2: ON"
      Top             =   4080
      Width           =   3135
   End
   Begin VB.CheckBox chk_alocacao 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4440
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txt_remetente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "+553488014709"
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox txt_mensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Nivel: 40% Bomba1: ON Bomba2: OFF"
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CheckBox chk_alocacao 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3480
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txt_remetente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "+553491086502"
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox txt_mensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Nivel: 20% Bomba1: OFF Bomba2: ON"
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CheckBox chk_alocacao 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2520
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txt_remetente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "+553491086564"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox txt_log 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1245
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6240
      Width           =   8775
   End
   Begin VB.TextBox txt_portaconfig 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7440
      TabIndex        =   1
      Text            =   "9600,n,8,1"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txt_portanumero 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7440
      TabIndex        =   0
      Text            =   "5"
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txt_remetente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "+553488126274"
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CheckBox chk_alocacao 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1560
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox txt_mensagem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Nivel: 00% Bomba1: ON Bomba2: OFF"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   5760
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5040
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
   End
   Begin VB.Label btn_chk_none 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desmarcar tudo"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2880
      MousePointer    =   99  'Custom
      TabIndex        =   62
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label btn_chk_all 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marcar tudo"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      MousePointer    =   99  'Custom
      TabIndex        =   61
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label label_remetente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remetente:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   4560
      TabIndex        =   60
      Top             =   5400
      Width           =   975
   End
   Begin VB.Shape linha_borda 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      Height          =   855
      Index           =   9
      Left            =   4560
      Top             =   4920
      Width           =   4335
   End
   Begin VB.Label label_mensagem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   4560
      TabIndex        =   59
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label label_remetente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remetente:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   4560
      TabIndex        =   55
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape linha_borda 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      Height          =   855
      Index           =   8
      Left            =   4560
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Label label_mensagem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   4560
      TabIndex        =   54
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label label_remetente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remetente:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   4560
      TabIndex        =   50
      Top             =   3480
      Width           =   975
   End
   Begin VB.Shape linha_borda 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      Height          =   855
      Index           =   7
      Left            =   4560
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label label_mensagem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   4560
      TabIndex        =   49
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label label_remetente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remetente:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   45
      Top             =   2520
      Width           =   975
   End
   Begin VB.Shape linha_borda 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      Height          =   855
      Index           =   6
      Left            =   4560
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Label label_mensagem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   44
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label label_remetente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remetente:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   40
      Top             =   1560
      Width           =   975
   End
   Begin VB.Shape linha_borda 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      Height          =   855
      Index           =   5
      Left            =   4560
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label label_mensagem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   4560
      TabIndex        =   39
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label label_mensagem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   35
      Top             =   5040
      Width           =   975
   End
   Begin VB.Shape linha_borda 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      Height          =   855
      Index           =   4
      Left            =   120
      Top             =   4920
      Width           =   4335
   End
   Begin VB.Label label_remetente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remetente:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   34
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label btn_start 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Iniciar"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4440
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label btn_stop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Parar"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6000
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label label_remetente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remetente:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   30
      Top             =   4440
      Width           =   975
   End
   Begin VB.Shape linha_borda 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      Height          =   855
      Index           =   3
      Left            =   120
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Label label_mensagem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   29
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label label_remetente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remetente:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Top             =   3480
      Width           =   975
   End
   Begin VB.Shape linha_borda 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      Height          =   855
      Index           =   2
      Left            =   120
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label label_mensagem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   27
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label label_remetente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remetente:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   2520
      Width           =   975
   End
   Begin VB.Shape linha_borda 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      Height          =   855
      Index           =   1
      Left            =   120
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Label label_mensagem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   25
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Autor: Cristiano Fraga G. Nunes"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   600
      Width           =   6135
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ModEmuGSM - Emulador de Modem GSM v1.0"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   360
      Width           =   6135
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Relatório de mensagens enviadas pelo modem:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   6000
      Width           =   5175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Configuração:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6240
      TabIndex        =   21
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Porta Comm:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6240
      TabIndex        =   20
      Top             =   240
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      Height          =   855
      Left            =   6120
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label label_mensagem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   975
   End
   Begin VB.Shape linha_borda 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      Height          =   855
      Index           =   0
      Left            =   120
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label label_remetente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Remetente:"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'para lembrar:
'vbCr = #0D
'vbCrLf = #0D & #0A
Option Explicit

Private Sub btn_chk_all_Click()
    Dim x As Integer
    For x = 1 To 10
     chk_alocacao(x - 1).Value = 1
    Next x
End Sub

Private Sub btn_chk_none_Click()
    Dim x As Integer
    For x = 1 To 10
        chk_alocacao(x - 1).Value = 0
    Next x
End Sub

Private Sub btn_start_Click()
    txt_portanumero.Enabled = False
    txt_portaconfig.Enabled = False
    btn_start.Enabled = False
    btn_stop.Enabled = True
    MSComm1.CommPort = Val(txt_portanumero.Text)
    MSComm1.Settings = txt_portaconfig
    sub_abre_serial
    Timer1.Interval = 1
End Sub

Private Sub btn_stop_Click()
    Timer1.Interval = 0
    btn_stop.Enabled = False
    btn_start.Enabled = True
    txt_portanumero.Enabled = True
    txt_portaconfig.Enabled = True
    sub_fecha_serial
End Sub

Sub sub_abre_serial()
    MSComm1.PortOpen = True
End Sub

Sub sub_fecha_serial()
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sub_fecha_serial
    End
End Sub

Private Sub Timer1_Timer()
    Dim x As Integer
    Dim texto_mensagem_enviada As String
    Dim destinatario_mensagem_enviada As String
    Dim leitura_da_mensagem As String
    Dim Results As String

    Picture1.BackColor = vbRed
    Do
        DoEvents
        If Timer1.Interval <> 0 Then
            Results = Results & MSComm1.Input
        Else
            Exit Do
        End If
    Loop Until InStr(Results, vbCr) > 0
    Picture1.BackColor = vbGreen

    leitura_da_mensagem = UCase(Results)

    If InStr(leitura_da_mensagem, "AT" & vbCr) > 0 Then
        MSComm1.Output = vbCrLf & "OK" & vbCrLf
    End If

    If InStr(leitura_da_mensagem, "ATE0" & vbCr) > 0 Then
        MSComm1.Output = vbCrLf & "OK" & vbCrLf
    End If

    If InStr(leitura_da_mensagem, "AT+CMGF=1" & vbCr) > 0 Then
        MSComm1.Output = vbCrLf & "OK" & vbCrLf
    End If

    For x = 1 To 10
        If InStr(leitura_da_mensagem, "AT+CMGR=" & x & vbCr) > 0 Then
            If chk_alocacao(x - 1).Value = 1 Then
                MSComm1.Output = vbCrLf & "+CMGR: " & Chr(&H22) & "REC UNREAD" & Chr(&H22) & "," _
                & Chr(&H22) & txt_remetente(x - 1).Text & Chr(&H22) & ",," & Chr(&H22) & "08/04/19,10:25:22+92" & Chr(&H22) & vbCrLf & txt_mensagem(x - 1).Text & vbCrLf & vbCrLf & "OK" & vbCrLf
            Else
                MSComm1.Output = vbCrLf & "+CMGR: 0,,0" & vbCrLf & vbCrLf & "OK" & vbCrLf
            End If
        End If
    Next x

    If InStr(leitura_da_mensagem, "AT+CMGS=") > 0 Then
        destinatario_mensagem_enviada = Mid(leitura_da_mensagem, (InStr(leitura_da_mensagem, Chr(&H22)) + 1), (InStrRev(leitura_da_mensagem, Chr(&H22)) - InStr(leitura_da_mensagem, Chr(&H22)) - 1))
        MSComm1.Output = vbCrLf & "> "
        Results = ""
        Picture1.BackColor = vbBlue
        Do
            DoEvents
            If Timer1.Interval <> 0 Then
                Results = Results & MSComm1.Input
            Else
                Exit Do
            End If
        Loop Until InStr(Results, Chr(&H1A)) > 0
        Picture1.BackColor = vbGreen
        texto_mensagem_enviada = Left(Results, Len(Results) - 1)
        MSComm1.Output = vbCrLf & "OK" & vbCrLf
        txt_log.Text = txt_log.Text & "Enviado: " & texto_mensagem_enviada & " -> " & destinatario_mensagem_enviada & vbCrLf
    End If

    For x = 1 To 10
        If InStr(leitura_da_mensagem, "AT+CMGD=" & x & vbCr) > 0 Then
            chk_alocacao(x - 1).Value = 0
            MSComm1.Output = vbCrLf & "OK" & vbCrLf
        End If
    Next x
End Sub

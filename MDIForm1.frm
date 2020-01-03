VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Zoologico"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "MDIForm1"
   MousePointer    =   1  'Arrow
   Moveable        =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   5775
      Left            =   0
      Picture         =   "MDIForm1.frx":0000
      ScaleHeight     =   5715
      ScaleWidth      =   9900
      TabIndex        =   0
      Top             =   0
      Width           =   9960
      Begin VB.CommandButton btnForm3 
         Caption         =   "&Eliminar de la Base de Datos del Zoo"
         Height          =   615
         Left            =   4320
         TabIndex        =   5
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton btnForm2 
         Caption         =   "&Agregar Animal"
         Height          =   615
         Left            =   5760
         TabIndex        =   2
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton btnForm1 
         Caption         =   "&Mostrar Info"
         Height          =   615
         Left            =   2880
         TabIndex        =   1
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Elija las opcion que desee para ser derivado"
         BeginProperty Font 
            Name            =   "Liberation Mono"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1200
         TabIndex        =   4
         Top             =   960
         Width           =   7620
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bienvenido a la Base de Datos del Zoo"
         BeginProperty Font 
            Name            =   "Liberation Mono"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   420
         TabIndex        =   3
         Top             =   360
         Width           =   8895
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnForm1_Click()
frmMostrar.Show
End Sub

Private Sub btnForm2_Click()
frmAgregar.Show
End Sub

Private Sub btnForm3_Click()
frmEliminar.Show
End Sub

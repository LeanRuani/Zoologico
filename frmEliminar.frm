VERSION 5.00
Begin VB.Form frmEliminar 
   Caption         =   "Eliminar Especies"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   Picture         =   "frmEliminar.frx":0000
   ScaleHeight     =   4740
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4755
      Left            =   0
      Picture         =   "frmEliminar.frx":E2CF
      ScaleHeight     =   4755
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   0
      Width           =   7185
      Begin VB.CommandButton btnBoton 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   495
         Left            =   3960
         TabIndex        =   3
         Top             =   2520
         Width           =   2055
      End
      Begin VB.ComboBox Especies 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   2415
      End
      Begin VB.ListBox Animales 
         Height          =   2205
         ItemData        =   "frmEliminar.frx":1C59E
         Left            =   1200
         List            =   "frmEliminar.frx":1C5A5
         TabIndex        =   1
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione una de las Especies"
         BeginProperty Font 
            Name            =   "Liberation Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   3600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Animales"
         BeginProperty Font 
            Name            =   "Liberation Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmEliminar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
cargarEspecie
End Sub

Private Sub btnBoton_Click()
If (Animales.ListIndex < 0) Then
    MsgBox "Por favor seleccion un animal"
Else
EliminarAlimentos
EliminarAnimales
End If
End Sub

'Carga del Especie
Private Sub cargarEspecie()
Dim qs As String
Dim cnx As New ADODB.Connection
Dim sql As String
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset

qs = "DSN=local"
cnx.Open qs
sql = "SELECT * FROM especies ORDER BY nombre "
cmd.ActiveConnection = cnx
cmd.CommandText = sql
cmd.CommandType = adCmdText

Set rs = cmd.Execute
Do While Not rs.EOF
DoEvents
    Especies.AddItem rs("nombre")
    Especies.ItemData(Especies.ListCount - 1) = rs("id")
    rs.MoveNext
Loop
rs.Close
cnx.Close
End Sub

'Muestra los animales de la especie seleccionada
Private Sub Especies_Click()
Dim id_especie As Integer: id_especie = Especies.ItemData(Especies.ListIndex)
Dim qs As String
Dim cnx As New ADODB.Connection
Dim sql As String
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Animales.Clear

qs = "dsn=local"
cnx.Open qs
sql = "SELECT * FROM animales WHERE id_especie = " & id_especie & " ORDER BY nombre ;"
cmd.ActiveConnection = cnx
cmd.CommandText = sql
cmd.CommandType = adCmdText

Set rs = cmd.Execute
Do While Not rs.EOF
DoEvents
        Animales.AddItem rs("nombre")
        Animales.ItemData(Animales.ListCount - 1) = rs("id")
        rs.MoveNext
Loop
rs.Close
cnx.Close
End Sub

'Elimina la fila del id_animal en la tabla animales_alimentos
Private Sub EliminarAlimentos()
Dim id_animal As Integer: id_animal = Animales.ItemData(Animales.ListIndex)
Dim qs As String
Dim cnx As New ADODB.Connection
Dim sql As String
Dim cmd As New ADODB.Command

qs = "DSN=local"
cnx.Open qs
sql = "DELETE FROM animales_alimentos WHERE id_animal = " & id_animal & ";"
cmd.ActiveConnection = cnx
cmd.CommandText = sql
cmd.CommandType = adCmdText
cmd.Execute
cnx.Close
End Sub

'Elimina la fila del animal de la tabla Animales
Private Sub EliminarAnimales()
Dim id_animal As Integer: id_animal = Animales.ItemData(Animales.ListIndex)
Dim qs As String
Dim cnx As New ADODB.Connection
Dim sql As String
Dim cmd As New ADODB.Command
qs = "DSN=local"
cnx.Open qs
sql = "DELETE FROM animales  WHERE id = " & id_animal & ";"
cmd.ActiveConnection = cnx
cmd.CommandText = sql
cmd.CommandType = adCmdText
cmd.Execute
cnx.Close
Especies_Click
End Sub

VERSION 5.00
Begin VB.Form frmMostrar 
   Caption         =   "Zoologico"
   ClientHeight    =   5820
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   5895
      Left            =   0
      Picture         =   "Base de Datos.frx":0000
      ScaleHeight     =   5835
      ScaleWidth      =   8355
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.ListBox Animales 
         Height          =   1620
         ItemData        =   "Base de Datos.frx":9DFC
         Left            =   5160
         List            =   "Base de Datos.frx":9E03
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox Especies 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   2415
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
         Height          =   255
         Left            =   5280
         TabIndex        =   5
         Top             =   360
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
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3600
      End
      Begin VB.Label lblAnimal 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   2175
         Left            =   840
         TabIndex        =   3
         Top             =   3120
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmMostrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
cargarEspecie
End Sub

'Carga las especies
Private Sub cargarEspecie()
Dim qs As String
Dim cnx As New ADODB.Connection
Dim sql As String
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset

qs = "DSN=local"
cnx.Open qs
sql = "SELECT * FROM especies ORDER BY nombre"
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
sql = "SELECT*FROM animales WHERE id_especie= " & id_especie & " ORDER BY nombre ;"
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

'Muestra la informacion del animal
Private Sub Animales_Click()
Dim id_animal As Integer: id_animal = Animales.ItemData(Animales.ListIndex)
Dim qs As String
Dim cnx As New ADODB.Connection
Dim sql As String
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset

qs = "dsn=local"
cnx.Open qs
sql = "SELECT a.id AS id, a.nombre AS nombre, a.sexo AS sexo, a.f_nacimiento AS f_nacimiento, z.id AS zona, c.nombre AS nombreC, al.nombre AS nombreA FROM animales a LEFT JOIN especies e ON e.id = a.id_especie LEFT JOIN cuidadores c ON c.id = e.id_cuidador LEFT JOIN zonas z ON z.id_especie = e.id LEFT JOIN animales_alimentos aa ON a.id = aa.id_animal LEFT JOIN alimentos al ON al.id = aa.id_alimento WHERE a.id = " & id_animal & ";"
cmd.ActiveConnection = cnx
cmd.CommandText = sql
cmd.CommandType = adCmdText

Set rs = cmd.Execute
Do While Not rs.EOF
DoEvents
    lblAnimal.Caption = "ID: " & rs("id") & vbNewLine & "Sexo: " & rs("sexo") & vbNewLine & "Fecha de nacimiento: " & rs("f_nacimiento") & vbNewLine & "Zona: " & rs("zona") & vbNewLine & "Nombre cuidador: " & rs("nombreC") & vbNewLine & "Alimento: " & rs("NombreA") & ""
    rs.MoveNext
Loop
rs.Close
cnx.Close
End Sub


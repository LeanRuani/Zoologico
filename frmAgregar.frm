VERSION 5.00
Begin VB.Form frmAgregar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agrega detalles del animal"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9060
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAgregar.frx":0000
   ScaleHeight     =   4665
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Height          =   4665
      Left            =   0
      Picture         =   "frmAgregar.frx":FF0E
      ScaleHeight     =   4605
      ScaleWidth      =   9030
      TabIndex        =   0
      Top             =   0
      Width           =   9090
      Begin VB.ComboBox Especie 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox cbmSexo 
         Height          =   315
         ItemData        =   "frmAgregar.frx":1C0E1
         Left            =   3240
         List            =   "frmAgregar.frx":1C0E3
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtNombre 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3840
         TabIndex        =   4
         ToolTipText     =   "Nombre del Animal"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3840
         TabIndex        =   3
         ToolTipText     =   "Fecha de Nacimento"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton btnGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   3240
         Width           =   5055
      End
      Begin VB.ComboBox Alimentos 
         Height          =   315
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Seleccione su Sexo"
         BeginProperty Font 
            Name            =   "Liberation Mono"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   3240
         TabIndex        =   12
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Seleccione su Especie"
         BeginProperty Font 
            Name            =   "Liberation Mono"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   2355
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre"
         BeginProperty Font 
            Name            =   "Liberation Mono"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2280
         TabIndex        =   10
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Año-Mes-Dia"
         BeginProperty Font 
            Name            =   "Liberation Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4200
         TabIndex        =   9
         Top             =   2760
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha de Nacimiento"
         BeginProperty Font 
            Name            =   "Liberation Mono"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1560
         TabIndex        =   8
         Top             =   2280
         Width           =   2025
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Seleccione su Alimento"
         BeginProperty Font 
            Name            =   "Liberation Mono"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5760
         TabIndex        =   7
         Top             =   240
         Width           =   2325
      End
   End
End
Attribute VB_Name = "frmAgregar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public id_animal As Integer

'Carga las Especies y en el combo carga sus items
Private Sub Form_Load()
cargarEspecie
cargarAlimentos
With cbmSexo
    .AddItem "Hembra"
    .AddItem "Macho"
End With
End Sub

'Inserta en la tabla animales lo que el usuario pone
Private Sub btnGrabar_Click()
If (Especie.ListIndex < 0) Then
    MsgBox "Ingrese todos los datos del animal para poder guardarlo"
Else
Dim id_especie As Integer: id_especie = Especie.ItemData(Especie.ListIndex)
Dim nombre As String: nombre = txtNombre.Text
Dim fecha As String: fecha = txtFecha.Text
Dim qs As String
Dim cnx As New ADODB.Connection
Dim sql As String
Dim cmd As New ADODB.Command

If (txtNombre.Text = "" Or txtFecha.Text = "" Or cbmSexo.ListIndex < 0 Or Alimentos.ListIndex < 0) Then
        MsgBox "Ingrese todos los datos del animal para poder guardarlo"
    Else
        qs = "DSN=local"
        cnx.Open qs
        sql = "INSERT INTO animales (id_especie, nombre, sexo, f_nacimiento)VALUES(" & id_especie & ",'" & nombre & "', '" & cbmSexo & "','" & fecha & "');"
        cmd.ActiveConnection = cnx
        cmd.CommandText = sql
        cmd.CommandType = adCmdText
        cmd.Execute
        cnx.Close
        guardarAnimal
        guardarAn_Al
        MsgBox "El Animal se guardo en la Base de Datos"
        Limpieza
End If
End If
End Sub

'Carga la Especie
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
    Especie.AddItem rs("nombre")
    Especie.ItemData(Especie.ListCount - 1) = rs("id")
    rs.MoveNext
Loop
rs.Close
cnx.Close
End Sub

'Muestra el nombre de los alimentos de la table Alimentos
Private Sub cargarAlimentos()
Dim qs As String
Dim cnx As New ADODB.Connection
Dim sql As String
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset

qs = "DSN=local"
cnx.Open qs
sql = "SELECT * FROM alimentos ORDER BY nombre"
cmd.ActiveConnection = cnx
cmd.CommandText = sql
cmd.CommandType = adCmdText

Set rs = cmd.Execute
Do While Not rs.EOF
DoEvents
    Alimentos.AddItem rs("nombre")
    Alimentos.ItemData(Alimentos.ListCount - 1) = rs("id")
    rs.MoveNext
Loop
rs.Close
cnx.Close
End Sub

'Guarda el alimento y el animal seleccionado
Private Sub guardarAn_Al()
Dim id_alimento As Integer: id_alimento = Alimentos.ItemData(Alimentos.ListIndex)
Dim qs As String
Dim cnx As New ADODB.Connection
Dim sql As String
Dim cmd As New ADODB.Command

qs = "DSN=local"
cnx.Open qs
sql = "INSERT INTO animales_alimentos (id_animal, id_alimento)VALUES(" & id_animal & ", " & id_alimento & ");"
cmd.ActiveConnection = cnx
cmd.CommandText = sql
cmd.CommandType = adCmdText
cmd.Execute
cnx.Close
End Sub

'Guarda el animal en la tabla animales_alimentos
Private Sub guardarAnimal()
Dim qs As String
Dim cnx As New ADODB.Connection
Dim sql As String
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
qs = "DSN=local"
cnx.Open qs
sql = "SELECT * FROM animales ORDER BY id DESC LIMIT 1 "
cmd.ActiveConnection = cnx
cmd.CommandText = sql
cmd.CommandType = adCmdText
Set rs = cmd.Execute
id_animal = rs("id")
rs.Close
cnx.Close

End Sub

'Guarda el alimento en la tabla animales_alimentos
Private Sub guardarAlimento()
Dim id_alimento As Integer: id_alimento = Alimentos.ItemData(Alimentos.ListIndex)
Dim qs As String
Dim cnx As New ADODB.Connection
Dim sql As String
Dim cmd As New ADODB.Command

qs = "DSN=local"
cnx.Open qs
sql = "SELECT * FROM alimentos al INNER JOIN animales_alimentos aa ON al.id = aa.id_alimento WHERE al.id = " & id_alimento & ";"
cmd.ActiveConnection = cnx
cmd.CommandText = sql
cmd.CommandType = adCmdText
cmd.Execute
cnx.Close
End Sub

Private Sub Limpieza()
Alimentos.Clear
cargarAlimentos
Especie.Clear
cargarEspecie
cbmSexo.Clear
With cbmSexo
    .AddItem "Hembra"
    .AddItem "Macho"
End With
txtNombre = ""
txtFecha = ""
End Sub


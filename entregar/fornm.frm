VERSION 5.00
Begin VB.Form fornm 
   Caption         =   "EDAD"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   360
   End
   Begin VB.CommandButton salir 
      Caption         =   "SALIR"
      Height          =   735
      Left            =   8040
      TabIndex        =   6
      Top             =   7200
      Width           =   2895
   End
   Begin VB.CommandButton lim 
      Caption         =   "LIMPIAR"
      Height          =   735
      Left            =   4560
      TabIndex        =   5
      Top             =   7200
      Width           =   2895
   End
   Begin VB.CommandButton calcular 
      Caption         =   "CALCULAR"
      Height          =   735
      Left            =   1080
      TabIndex        =   4
      Top             =   7200
      Width           =   2655
   End
   Begin VB.TextBox eds 
      Height          =   975
      Left            =   4560
      TabIndex        =   3
      Top             =   5160
      Width           =   5055
   End
   Begin VB.TextBox fcn 
      Height          =   855
      Left            =   4560
      TabIndex        =   2
      Top             =   2520
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "00:00:00"
      Height          =   615
      Left            =   12720
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label eda 
      Caption         =   "EDAD"
      Height          =   855
      Left            =   4560
      TabIndex        =   1
      Top             =   3840
      Width           =   5055
   End
   Begin VB.Label inge 
      Caption         =   "FECHA"
      Height          =   855
      Left            =   4560
      TabIndex        =   0
      Top             =   1200
      Width           =   5055
   End
End
Attribute VB_Name = "fornm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub calcular_Click()
Dim feshnao As Date, edad As Integer
feshnao = CDate(fcn)
edad = CInt((Date - feshnao) / 365)
eds = Str(edad) & “AÑOS”
End Sub

Private Sub lim_Click()
fcn = “”
eds = “”
fcn.SetFocus
End Sub

Private Sub salir_Click()
End
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Format(Time, "hh:mm:ss")
End Sub

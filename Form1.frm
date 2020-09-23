VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Esempio Genera/Elimina Stampa"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Dead print"
      Enabled         =   0   'False
      Height          =   420
      Left            =   2160
      TabIndex        =   1
      Top             =   2115
      Width           =   1950
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Make print"
      Height          =   420
      Left            =   90
      TabIndex        =   0
      Top             =   2115
      Width           =   1950
   End
   Begin VB.Label Label1 
      Height          =   1950
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   4020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdStampa$

Private Sub Command1_Click()
    
    '
    ' Genero una stampa
    '
    IdStampa$ = InizioStampa()
    If Len(IdStampa$) Then
        Printer.Print "Una scritta"
        Printer.Print "giusto per stampare"
        Printer.Print "qualcosa!  ^_^"
        Printer.EndDoc
        Command1.Enabled = False
        Command2.Enabled = True
    End If

End Sub

Private Sub Command2_Click()

    '
    ' Elimino la stampa generata dallo SPOOL
    '
    If Len(IdStampa$) Then
        Call EliminaStampa(IdStampa$)
        IdStampa$ = ""
        Command1.Enabled = True
        Command2.Enabled = False
    End If

End Sub

Private Sub Form_Load()

    Label1 = "Make a print and (quickly) you can dead it!"


End Sub

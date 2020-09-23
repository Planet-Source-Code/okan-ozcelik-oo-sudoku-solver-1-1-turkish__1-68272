VERSION 5.00
Begin VB.Form frmÝþlemVar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bekleyin..."
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdÝptal 
      Caption         =   "Ýptal"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Bilgi 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmÝþlemVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   OÖ Sudoku Çözer 1.1
'   Yapým Mart 2007

'   Telif Hakký © Okan ÖZÇELÝK
'   Geliþtiren Okan ÖZÇELÝK
'   okan_ozcelik@yahoo.com.tr

'   Bu yazýlým kodlarý üzerinde deðiþiklik yapýlmamalý.
'   Bu yazýlým kodlarý sadece örnek olarak incelenmeli.

Option Explicit

Const bçDeneniyor As String = "Çözüm için tahminler deneniyor."

Private Sub cmdÝptal_Click()
frmÇözer.ÝþlemVar = 1
End Sub

Private Sub Form_Load()
frmÇözer.ÝþlemVar = 2
Bilgi.Caption = bçDeneniyor
frmÇözer.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmÇözer.Enabled = True
End Sub

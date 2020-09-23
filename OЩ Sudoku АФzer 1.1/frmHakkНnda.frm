VERSION 5.00
Begin VB.Form frmHakkýnda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hakkýnda"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTamam 
      Caption         =   "Tamam"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblhBilgi 
      Caption         =   "Label1"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmHakkýnda"
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

Public PAdý As String

Private Sub HakkýndaOluþtur()
Const PYapým As String = "Yapým Mart 2007"
Const PGeliþtiren As String = "Okan ÖZÇELÝK"
Const Peposta As String = "okan_ozcelik@yahoo.com.tr"
Const PBilgi As String = "Bu yazýlým kodlarý üzerinde deðiþiklik yapýlmamalý." & _
vbCrLf & "Bu yazýlým kodlarý sadece örnek olarak incelenmeli."
Dim hBilgi(7) As String
Dim BilgiSýra As Byte

hBilgi(0) = PAdý
hBilgi(1) = PYapým
hBilgi(2) = ""
hBilgi(3) = "Telif Hakký © " & PGeliþtiren
hBilgi(4) = "Geliþtiren " & PGeliþtiren
hBilgi(5) = "e-posta " & Peposta
hBilgi(6) = ""
hBilgi(7) = PBilgi

lblhBilgi = ""
For BilgiSýra = 0 To UBound(hBilgi)
    lblhBilgi = lblhBilgi & hBilgi(BilgiSýra) & vbCrLf
Next

End Sub

Private Sub cmdTamam_Click()
Unload Me
End Sub

Private Sub Form_Load()
HakkýndaOluþtur
End Sub

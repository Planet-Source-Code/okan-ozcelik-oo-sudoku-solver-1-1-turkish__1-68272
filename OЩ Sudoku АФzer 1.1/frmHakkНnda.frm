VERSION 5.00
Begin VB.Form frmHakk�nda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hakk�nda"
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
Attribute VB_Name = "frmHakk�nda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   O� Sudoku ��zer 1.1
'   Yap�m Mart 2007

'   Telif Hakk� � Okan �Z�EL�K
'   Geli�tiren Okan �Z�EL�K
'   okan_ozcelik@yahoo.com.tr

'   Bu yaz�l�m kodlar� �zerinde de�i�iklik yap�lmamal�.
'   Bu yaz�l�m kodlar� sadece �rnek olarak incelenmeli.


Option Explicit

Public PAd� As String

Private Sub Hakk�ndaOlu�tur()
Const PYap�m As String = "Yap�m Mart 2007"
Const PGeli�tiren As String = "Okan �Z�EL�K"
Const Peposta As String = "okan_ozcelik@yahoo.com.tr"
Const PBilgi As String = "Bu yaz�l�m kodlar� �zerinde de�i�iklik yap�lmamal�." & _
vbCrLf & "Bu yaz�l�m kodlar� sadece �rnek olarak incelenmeli."
Dim hBilgi(7) As String
Dim BilgiS�ra As Byte

hBilgi(0) = PAd�
hBilgi(1) = PYap�m
hBilgi(2) = ""
hBilgi(3) = "Telif Hakk� � " & PGeli�tiren
hBilgi(4) = "Geli�tiren " & PGeli�tiren
hBilgi(5) = "e-posta " & Peposta
hBilgi(6) = ""
hBilgi(7) = PBilgi

lblhBilgi = ""
For BilgiS�ra = 0 To UBound(hBilgi)
    lblhBilgi = lblhBilgi & hBilgi(BilgiS�ra) & vbCrLf
Next

End Sub

Private Sub cmdTamam_Click()
Unload Me
End Sub

Private Sub Form_Load()
Hakk�ndaOlu�tur
End Sub

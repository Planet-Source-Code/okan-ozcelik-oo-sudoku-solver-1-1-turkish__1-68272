VERSION 5.00
Begin VB.UserControl Bölüm 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txtHücre 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      MaxLength       =   1
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Bölüm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   OÖ Sudoku Çözer 1.1
'   Yapým Mart 2007

'   Telif Hakký © Okan ÖZÇELÝK
'   Geliþtiren Okan ÖZÇELÝK
'   okan_ozcelik@yahoo.com.tr

'   Bu yazýlým kodlarý üzerinde deðiþiklik yapýlmamalý.
'   Bu yazýlým kodlarý sadece örnek olarak incelenmeli.


'3x3 hücre bloðunu barýndýran bölüm üretmek için.

Option Explicit
Const Boþluk As Single = 30

Public Enum HücreBiçimiSeç
    hbEtkiYok = 0
    hbGiriþ = 1
    hbBulmaca = 2
    hbÇözüm = 3
    hbTahmin = 4
    hbGiriþYanlýþ = 5
End Enum

Public Event GiriþYapýldý(hIndex As Byte)

Public Sub HücreEtkinleþtir(ByVal Index As Byte)
txtHücre(Index).SetFocus
End Sub

Public Function Hücre(ByVal Index As Byte, Optional ByVal hRakam As Byte = 10, _
                                            Optional ByVal hBiçim As HücreBiçimiSeç = hbEtkiYok) As Byte
Const rGiriþ As Long = &H80000005
Const rBulmaca As Long = &H8000000F
Const rÇözüm As Long = &H80000018
Const rTahmin As Long = &HC0FFFF
Const rGiriþYanlýþ As Long = &HFF&
If txtHücre(Index).Text = "" Then Hücre = 0 Else Hücre = CByte(txtHücre(Index).Text)
If Not hRakam = 10 Then txtHücre(Index).Text = CStr(hRakam) 'hRakam=10 etkisiz olacaktýr,
If hBiçim = hbEtkiYok Then Exit Function
With txtHücre(Index)
    .FontBold = True
    .Locked = True
    Select Case hBiçim
        Case hbGiriþ
            .BackColor = rGiriþ
            .Locked = False
        Case hbBulmaca
            .BackColor = rBulmaca
        Case hbÇözüm
            .BackColor = rÇözüm
            .FontBold = False
        Case hbTahmin
            .BackColor = rTahmin
            .FontBold = False
        Case hbGiriþYanlýþ
            .BackColor = rGiriþYanlýþ
    End Select
End With
End Function

Private Sub UserControl_Initialize()

Dim Sýra As Byte

With txtHücre(0)
    .Left = Boþluk * 2
    .Top = Boþluk * 2
End With

For Sýra = 1 To 8
    Load txtHücre(Sýra)
    With txtHücre(Sýra)
        .Visible = True
        If (Sýra + 1) Mod 3 = 1 Then
            .Left = txtHücre(0).Left
            .Top = txtHücre(Sýra - 3).Top + txtHücre(0).Height + Boþluk
        Else
            .Left = txtHücre(Sýra - 1).Left + txtHücre(0).Width + Boþluk
            .Top = txtHücre(Sýra - 1).Top
        End If
    End With
Next

End Sub

Private Sub UserControl_Resize()

With txtHücre(8)
    Width = .Left + .Width + Boþluk * 2
    Height = .Top + .Height + Boþluk * 2
End With

End Sub


Private Sub txtHücre_Change(Index As Integer)
Const hBoþ As String = " " 'Boþluk tuþuna da basýldýðýnda olay oluþmalý
If txtHücre(Index).Text Like "[1-9]" Or txtHücre(Index).Text = hBoþ Then
    If txtHücre(Index).Text = hBoþ Then txtHücre(Index).Text = ""
    RaiseEvent GiriþYapýldý(CByte(Index))
Else
   txtHücre(Index).Text = ""
End If
End Sub


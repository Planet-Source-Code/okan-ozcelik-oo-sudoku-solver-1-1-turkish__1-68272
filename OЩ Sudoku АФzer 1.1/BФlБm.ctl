VERSION 5.00
Begin VB.UserControl B�l�m 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txtH�cre 
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
Attribute VB_Name = "B�l�m"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   O� Sudoku ��zer 1.1
'   Yap�m Mart 2007

'   Telif Hakk� � Okan �Z�EL�K
'   Geli�tiren Okan �Z�EL�K
'   okan_ozcelik@yahoo.com.tr

'   Bu yaz�l�m kodlar� �zerinde de�i�iklik yap�lmamal�.
'   Bu yaz�l�m kodlar� sadece �rnek olarak incelenmeli.


'3x3 h�cre blo�unu bar�nd�ran b�l�m �retmek i�in.

Option Explicit
Const Bo�luk As Single = 30

Public Enum H�creBi�imiSe�
    hbEtkiYok = 0
    hbGiri� = 1
    hbBulmaca = 2
    hb��z�m = 3
    hbTahmin = 4
    hbGiri�Yanl�� = 5
End Enum

Public Event Giri�Yap�ld�(hIndex As Byte)

Public Sub H�creEtkinle�tir(ByVal Index As Byte)
txtH�cre(Index).SetFocus
End Sub

Public Function H�cre(ByVal Index As Byte, Optional ByVal hRakam As Byte = 10, _
                                            Optional ByVal hBi�im As H�creBi�imiSe� = hbEtkiYok) As Byte
Const rGiri� As Long = &H80000005
Const rBulmaca As Long = &H8000000F
Const r��z�m As Long = &H80000018
Const rTahmin As Long = &HC0FFFF
Const rGiri�Yanl�� As Long = &HFF&
If txtH�cre(Index).Text = "" Then H�cre = 0 Else H�cre = CByte(txtH�cre(Index).Text)
If Not hRakam = 10 Then txtH�cre(Index).Text = CStr(hRakam) 'hRakam=10 etkisiz olacakt�r,
If hBi�im = hbEtkiYok Then Exit Function
With txtH�cre(Index)
    .FontBold = True
    .Locked = True
    Select Case hBi�im
        Case hbGiri�
            .BackColor = rGiri�
            .Locked = False
        Case hbBulmaca
            .BackColor = rBulmaca
        Case hb��z�m
            .BackColor = r��z�m
            .FontBold = False
        Case hbTahmin
            .BackColor = rTahmin
            .FontBold = False
        Case hbGiri�Yanl��
            .BackColor = rGiri�Yanl��
    End Select
End With
End Function

Private Sub UserControl_Initialize()

Dim S�ra As Byte

With txtH�cre(0)
    .Left = Bo�luk * 2
    .Top = Bo�luk * 2
End With

For S�ra = 1 To 8
    Load txtH�cre(S�ra)
    With txtH�cre(S�ra)
        .Visible = True
        If (S�ra + 1) Mod 3 = 1 Then
            .Left = txtH�cre(0).Left
            .Top = txtH�cre(S�ra - 3).Top + txtH�cre(0).Height + Bo�luk
        Else
            .Left = txtH�cre(S�ra - 1).Left + txtH�cre(0).Width + Bo�luk
            .Top = txtH�cre(S�ra - 1).Top
        End If
    End With
Next

End Sub

Private Sub UserControl_Resize()

With txtH�cre(8)
    Width = .Left + .Width + Bo�luk * 2
    Height = .Top + .Height + Bo�luk * 2
End With

End Sub


Private Sub txtH�cre_Change(Index As Integer)
Const hBo� As String = " " 'Bo�luk tu�una da bas�ld���nda olay olu�mal�
If txtH�cre(Index).Text Like "[1-9]" Or txtH�cre(Index).Text = hBo� Then
    If txtH�cre(Index).Text = hBo� Then txtH�cre(Index).Text = ""
    RaiseEvent Giri�Yap�ld�(CByte(Index))
Else
   txtH�cre(Index).Text = ""
End If
End Sub


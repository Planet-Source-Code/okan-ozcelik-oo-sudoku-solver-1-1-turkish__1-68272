VERSION 5.00
Begin VB.Form frm��zer 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frm��zer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame SudAlan 
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1935
      Begin O�Sudoku��zer.B�l�m B�l�m 
         Height          =   1665
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   2937
      End
   End
   Begin VB.CommandButton cmd��k 
      Caption         =   "��k"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdHakk�nda 
      Caption         =   "Hakk�nda"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdYeni 
      Caption         =   "Yeni"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdG�ster 
      Caption         =   "Gizle/G�ster"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdB��z 
      Caption         =   "��z"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "frm��zer"
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

'D���nce tarz�n� daha kolay kavramak i�in "A��klama.rtf" dosyas�na da g�zatabilirsiniz.

Option Explicit

Private H�cre(8, 8) As Byte 'H�credeki rakam� tutar.
Private H�creOlas�(8, 8, 8) As Boolean 'H�credeki olas�l�klar� tutar.

Private bH�cre(8, 8) As Boolean 'Bulmacada verilen h�crelerin yerlerini tutar.
Private tH�cre(8, 8) As Boolean 'Tahmin edilen h�crelerin yerini tutar.

Private Enum hoTaramaY�n�
    tyYatay = 0
    tyD��ey = 1
    tyYatayD��ey = 2
End Enum

Private khSay�s� As Byte 'Kesinle�en h�cre say�s�

Private �Gizle As Boolean

Private Enum S�f�rlaSe�enekleri
    ssT�m�
    ssBulmacaD���nda
    ssTahmin
End Enum

Private Const hRakamBo� As Byte = 0 'H�creye bo� atama yapmak i�in

Public ��lemVar As Byte  '0 bigi penceresi a��lmam��, 1 i�lem iptal edilmi�, 2 i�lem s�r�yor

Dim FormY�klendi As Boolean

Private Sub cmdB��z_Click()
Tu�Kilidi False
Bulmacay�Al
If Not Bulmacay���z Then
    S�f�rla ssBulmacaD���nda
    MsgBox "Bulmaca ��z�lemedi!", vbCritical
Else
    H�creleriYaz
End If
End Sub

Private Sub Bulmacay�Al()
Dim bS�ra As Byte 'b�l�m s�ra
Dim hS�ra As Byte 'h�cre s�ra

Dim oS�ra As Byte 'olas�l�k s�ra

For bS�ra = 0 To 8
    For hS�ra = 0 To 8
        If Not B�l�m(bS�ra).H�cre(hS�ra) = hRakamBo� Then
            For oS�ra = 0 To 8
                H�creOlas�(bS�ra, hS�ra, oS�ra) = False
            Next
            H�creOlas�(bS�ra, hS�ra, B�l�m(bS�ra).H�cre(hS�ra) - 1) = True
            B�l�m(bS�ra).H�cre hS�ra, , hbBulmaca
            bH�cre(bS�ra, hS�ra) = True
        End If
    Next
Next

End Sub

Private Function Bulmacay���z() As Boolean
'Tarama yordamlar�, d�ng�n�n her ad�m�nda, bulmacay� tekrar tekrar ba�tan sona tararlar.

Dim �khSay�s� As Byte '�nceki khSay�s�
Const khSay�s�Tam As Byte = 81
Const bTahminSor As String = "Bu bulmacay� ��zmek i�in tahmin de denemek gerekiyor olabilir." & _
"Tahmin denensin mi?"
Dim TahminOnay As VbMsgBoxResult
Dim Ba�lamaAn� As Single
Const ZamanA��m� As Single = 10


Do
    S�f�rla ssTahmin
    hoKesinle�enleriTara
    Do
        Do
            �khSay�s� = khSay�s�
            hoBTekEksenTara
            hoFarkBirTara
            hoKesinle�enleriTara
            If khSay�s� = �khSay�s� Then
                If Not TahminOnay = vbYes Then _
                TahminOnay = MsgBox(bTahminSor, vbYesNo + vbQuestion)
                Exit Do
            End If
        Loop Until khSay�s� = khSay�s�Tam
        If khSay�s� = khSay�s�Tam Or TahminOnay = vbNo Then GoTo Bitir
    Loop While hoTahmin
    If Ba�lamaAn� = 0 Then
        Ba�lamaAn� = Timer
        frm��lemVar.Show
    End If
    DoEvents
    If ��lemVar = 1 Then GoTo Bitir
Loop While Timer - Ba�lamaAn� < ZamanA��m�

Bitir:
If Not ��lemVar = 0 Then ��lemVar = 0: Unload frm��lemVar
If khSay�s� = khSay�s�Tam Then Bulmacay���z = True

End Function

Private Sub hoKesinle�enleriTara()
'H�creOlas�'da bire inmi� olas�l�klar�n rakam de�erini H�cre'ye yerle�tirir.

Dim bS�ra As Byte
Dim hS�ra As Byte

Dim khRakam� As Byte 'kesinle�en h�cre rakam�

For bS�ra = 0 To 8
    For hS�ra = 0 To 8
        If H�cre(bS�ra, hS�ra) = 0 Then
            khRakam� = hoSaltBir(bS�ra, hS�ra)
            If Not khRakam� = 0 Then
                H�cre(bS�ra, hS�ra) = khRakam�
                hoOlas�l�kEle bS�ra, hS�ra, H�cre(bS�ra, hS�ra) - 1
                khSay�s� = khSay�s� + 1
            End If
        End If
    Next
Next

End Sub

Private Function hoSaltBir(sbB�l�m As Byte, sbH�cre As Byte) As Byte
'Belirtilen H�cre'nin olas�l�klar�n�n sadece bir tane oldu�unu denetler ve bunun hangi rakam oldu�unu.

Dim sbS�ra As Integer

For sbS�ra = 0 To 8
    If Not hoSaltBir = 0 Then
        If H�creOlas�(sbB�l�m, sbH�cre, sbS�ra) Then hoSaltBir = 0: Exit Function
    Else
        If H�creOlas�(sbB�l�m, sbH�cre, sbS�ra) Then hoSaltBir = sbS�ra + 1
    End If
Next

End Function

Private Function sYer��z(y�EksenKonum As Byte, y�teNesne As Byte, y�Y�n As hoTaramaY�n�) As Byte
'S�ra Yer ��z
'Konumland��� ekseni ve eksendeki s�ras� verilen nesnenin Index'ini verir.
If y�Y�n = tyYatay Then sYer��z = y�EksenKonum * 3 + y�teNesne _
Else sYer��z = y�EksenKonum + y�teNesne * 3
End Function

Private Function EksenKonumBul(ebNesne As Byte, ebY�n As hoTaramaY�n�) As Byte
''Index'i verilen nesnenin konumland��� ekseni verir.

Dim teNesne As Byte
If ebY�n = tyYatay Then
    Do While EksenKonumBul * 3 < ebNesne + 1
        EksenKonumBul = EksenKonumBul + 1
    Loop
    EksenKonumBul = EksenKonumBul - 1
Else
    For EksenKonumBul = 0 To 2
        For teNesne = 0 To 2
            If sYer��z(EksenKonumBul, teNesne, tyD��ey) = ebNesne Then Exit Function
        Next
    Next
End If

End Function

Private Function hoOlas�l�kEle(oeB�l�m As Byte, oeH�cre As Byte, oeElenen As Byte, _
                                                    Optional oeB�l�m�Tarama As Boolean = False, _
                                                    Optional oeY�n As hoTaramaY�n� = tyYatayD��ey) _
                                                    As Boolean

'En �ok olas�l��� bire inmi� h�crenin bulundu�u b�l�m ve eksenlerdeki h�crelerde, bu kesinle�en olas�l��� elemek amac�yla _
kullan�l�yor. BTekEksenTara yordam�na bakarsan�z, olas�l���n eksende kesinlik kazand���nda da kullan�ld���n� g�receksiniz.

Dim bS�ra As Byte
Dim hS�ra As Byte

Dim hEksenKonum As Byte
Dim bEksenKonum As Byte
Dim teH�cre As Byte 'Taranan eksendeki H�cre
Dim teB�l�m As Byte

'Ayn� B�l�m Kaynakl� Elemeler
If Not oeB�l�m�Tarama Then
    For hS�ra = 0 To 8
        If Not hS�ra = oeH�cre Then
            If H�creOlas�(oeB�l�m, hS�ra, oeElenen) Then
                If Not hoOlas�l�kEle Then hoOlas�l�kEle = True
                H�creOlas�(oeB�l�m, hS�ra, oeElenen) = False
            End If
        End If
    Next
End If

'Yatay kaynakl� elemeler
If oeY�n = tyYatayD��ey Or oeY�n = tyYatay Then
    bEksenKonum = EksenKonumBul(oeB�l�m, tyYatay)
    hEksenKonum = EksenKonumBul(oeH�cre, tyYatay)
    
    For teB�l�m = 0 To 2
        bS�ra = sYer��z(bEksenKonum, teB�l�m, tyYatay)
        For teH�cre = 0 To 2
            hS�ra = sYer��z(hEksenKonum, teH�cre, tyYatay)
            If Not (oeB�l�m�Tarama And bS�ra = oeB�l�m) Then
                If Not (bS�ra = oeB�l�m And hS�ra = oeH�cre) Then
                    If H�creOlas�(bS�ra, hS�ra, oeElenen) Then
                        If Not hoOlas�l�kEle Then hoOlas�l�kEle = True
                        H�creOlas�(bS�ra, hS�ra, oeElenen) = False
                    End If
                End If
            End If
        Next
    Next
End If

If oeY�n = tyYatayD��ey Or oeY�n = tyD��ey Then
    'D��ey kaynakl� elemeler
    bEksenKonum = EksenKonumBul(oeB�l�m, tyD��ey)
    hEksenKonum = EksenKonumBul(oeH�cre, tyD��ey)
    
    For teB�l�m = 0 To 2
        bS�ra = sYer��z(bEksenKonum, teB�l�m, tyD��ey)
        For teH�cre = 0 To 2
            hS�ra = sYer��z(hEksenKonum, teH�cre, tyD��ey)
            If Not (oeB�l�m�Tarama And bS�ra = oeB�l�m) Then
                If Not (bS�ra = oeB�l�m And hS�ra = oeH�cre) Then
                    If H�creOlas�(bS�ra, hS�ra, oeElenen) Then
                        If Not hoOlas�l�kEle Then hoOlas�l�kEle = True
                        H�creOlas�(bS�ra, hS�ra, oeElenen) = False
                    End If
                End If
            End If
        Next
    Next
End If

End Function

Private Sub hoFarkBirTara()
'Bir h�crenin olas�l�klar�ndan biri, bulundu�u eksen ve ya b�l�mdeki di�er h�crelerde yoksa _
bu h�crenin tek olas�l���n�n bu oldu�u ��kar�labilir.

Dim bS�ra As Byte
Dim hS�ra As Byte

Dim hEksenKonum As Byte
Dim bEksenKonum As Byte
Dim teH�cre As Byte 'Taranan eksendeki H�cre
Dim teB�l�m As Byte

Dim fbY�n As hoTaramaY�n� '0 yatay, 1 dikey
Dim fbOlas� As Byte
Dim bS�raAyn� As Byte
Dim hS�raAyn� As Byte ''Taranan h�crenin bir olas�l���n�n, ayn� _
olas�l��a sahip mi diye kar��la�t�r�ld��� h�crenin yerini belirliyor.
Dim tebS�raAyn� As Byte
Dim tehS�raAyn� As Byte
Dim fbOlas�l�kSil As Byte


'B�l�mlerde farkl� h�creler
For bS�ra = 0 To 8
    For hS�ra = 0 To 8
        If H�cre(bS�ra, hS�ra) = 0 Then
            For fbOlas� = 0 To 8
                If H�creOlas�(bS�ra, hS�ra, fbOlas�) Then
                    For hS�raAyn� = 0 To 8
                        If H�cre(bS�ra, hS�raAyn�) = 0 Then
                            If Not hS�ra = hS�raAyn� Then
                                If H�creOlas�(bS�ra, hS�raAyn�, fbOlas�) Then GoTo fbbDi�erOlas�
                            End If
                        End If
                    Next
                    'H�crenin di�erlerinde olmayan bir olas�l��a sahip oldu�u ortaya ��kt�.
                    For fbOlas�l�kSil = 0 To 8
                        H�creOlas�(bS�ra, hS�ra, fbOlas�l�kSil) = False
                    Next
                    H�creOlas�(bS�ra, hS�ra, fbOlas�) = True
                    GoTo fbbDi�erH�cre
                End If
fbbDi�erOlas�:
            Next
        End If
fbbDi�erH�cre:
    Next
Next


'Yatay ve D��ey'de bir farkl� h�creler
For fbY�n = 0 To 1
    For bEksenKonum = 0 To 2
        For hEksenKonum = 0 To 2
            For teB�l�m = 0 To 2
                bS�ra = sYer��z(bEksenKonum, teB�l�m, fbY�n)
                For teH�cre = 0 To 2
                    hS�ra = sYer��z(hEksenKonum, teH�cre, fbY�n)
                    If H�cre(bS�ra, hS�ra) = 0 Then
                        For fbOlas� = 0 To 8
                            If H�creOlas�(bS�ra, hS�ra, fbOlas�) Then
                                For tebS�raAyn� = 0 To 2
                                    bS�raAyn� = sYer��z(bEksenKonum, tebS�raAyn�, fbY�n)
                                    For tehS�raAyn� = 0 To 2
                                        hS�raAyn� = sYer��z(hEksenKonum, tehS�raAyn�, fbY�n)
                                        If H�cre(bS�raAyn�, hS�raAyn�) = 0 Then
                                            If Not (bS�ra = bS�raAyn� And hS�ra = hS�raAyn�) Then
                                                If H�creOlas�(bS�raAyn�, hS�raAyn�, fbOlas�) Then GoTo fbdDi�erOlas�
                                            End If
                                        End If
                                    Next
                                Next
                                'H�crenin di�erlerinde olmayan bir olas�l��a sahip oldu�u ortaya ��kt�.
                                For fbOlas�l�kSil = 0 To 8
                                    H�creOlas�(bS�ra, hS�ra, fbOlas�l�kSil) = False
                                Next
                                H�creOlas�(bS�ra, hS�ra, fbOlas�) = True
                                GoTo fbdDi�erH�cre
                            End If
fbdDi�erOlas�:
                        Next
                    End If
fbdDi�erH�cre:
                Next
            Next
        Next
    Next
Next

End Sub

Private Sub hoBTekEksenTara()
'Bir b�l�mde sadece tek eksende bulunabilecek olas�l�klar� tarar.

Dim bS�ra As Byte
Dim hS�ra As Byte

Dim hoOlas�l�kElendi As Boolean

Dim hEksenKonum As Byte
Dim bEksenKonum As Byte
Dim teH�cre As Byte 'Taranan eksendeki H�cre
Dim teB�l�m As Byte
Dim bteY�n As hoTaramaY�n�


Dim dhEksenKonum As Byte 'Di�er h�cre eksen konum
Dim bteOlas� As Byte
Dim dteH�cre As Byte
Dim dhS�ra As Byte


Do
hoOlas�l�kElendi = False

'Bir b�l�m�n sadece tek bir h�cre eksenine gelebilecek olas�l�klar varsa ayn� eksendeki di�er b�l�mlerin o h�cre _
eksenlerindeki olas�l�klar elenir. �rne�in;
'B�l�m(1)'in 2. h�cre ekseni d���nda 3 rakam�n�n gelemeyece�i ortaya ��kt�ysa, B�l�m(0 ve 2)'in 2. h�cre eksenlerindeki _
3 rakam� olas�l��� elenir.
For bteY�n = 0 To 1
    For bS�ra = 0 To 8
        For hEksenKonum = 0 To 2
            For teH�cre = 0 To 2
                hS�ra = sYer��z(hEksenKonum, teH�cre, bteY�n)
                If H�cre(bS�ra, hS�ra) = 0 Then
                    For bteOlas� = 0 To 8
                        If H�creOlas�(bS�ra, hS�ra, bteOlas�) Then
                            For dhEksenKonum = 0 To 2
                                If Not hEksenKonum = dhEksenKonum Then
                                    For dteH�cre = 0 To 2
                                        dhS�ra = sYer��z(dhEksenKonum, dteH�cre, bteY�n)
                                        If H�creOlas�(bS�ra, dhS�ra, bteOlas�) Then GoTo bteDi�erOlas�
                                    Next
                                End If
                            Next
                            If hoOlas�l�kEle(bS�ra, hS�ra, bteOlas�, True, bteY�n) Then hoOlas�l�kElendi = True
                        End If
bteDi�erOlas�:
                    Next
                End If
            Next
        Next
    Next
Next

Loop While hoOlas�l�kElendi

End Sub

Private Function hoTahmin() As Boolean
Dim bS�ra As Byte
Dim hS�ra As Byte
Dim oS�ra As Byte
Dim tOlas� As Byte 'Tahmin edilen olas�l�k
Dim dOlas�() As Boolean 'Denenmi� olas�l�klar, h�crede hi� olas�l�k kalmad���n� anlamak i�in.
Dim dOlas�S�ra As Byte
Dim Olas�l�kSil As Byte

For bS�ra = 0 To 8
    For hS�ra = 0 To 8
        If H�cre(bS�ra, hS�ra) = 0 Then
            ReDim dOlas�(8) As Boolean
            Do
                For dOlas�S�ra = 0 To 8
                    If Not dOlas�(dOlas�S�ra) Then Exit For
                Next
                If dOlas�S�ra = 8 + 1 Then Exit Function 'B�t�n olas�l�klar denenmi�.
                oS�ra = Rnd * 8
                dOlas�(oS�ra) = True
            Loop Until H�creOlas�(bS�ra, hS�ra, oS�ra)
            For Olas�l�kSil = 0 To 8
                H�creOlas�(bS�ra, hS�ra, Olas�l�kSil) = False
            Next
            H�creOlas�(bS�ra, hS�ra, oS�ra) = True
            tH�cre(bS�ra, hS�ra) = True
            hoTahmin = True
            Exit Function
        End If
    Next
Next

End Function

Private Sub H�creleriYaz()

Dim bS�ra As Byte
Dim hS�ra As Byte

For bS�ra = 0 To 8
    For hS�ra = 0 To 8
            If Not bH�cre(bS�ra, hS�ra) Then
                If tH�cre(bS�ra, hS�ra) Then B�l�m(bS�ra).H�cre hS�ra, H�cre(bS�ra, hS�ra), hbTahmin _
                Else: B�l�m(bS�ra).H�cre hS�ra, H�cre(bS�ra, hS�ra), hb��z�m
            End If
    Next
Next

End Sub

Private Sub S�f�rla(Optional sSe�im As S�f�rlaSe�enekleri = ssT�m�)
Dim bS�ra As Byte
Dim hS�ra As Byte
Dim oS�ra As Byte

khSay�s� = 0

If Not sSe�im = ssTahmin Then
    �Gizle = True
    Tu�Kilidi True
End If

For bS�ra = 0 To 8
    For hS�ra = 0 To 8
        H�cre(bS�ra, hS�ra) = 0
        If (sSe�im = ssTahmin And Not bH�cre(bS�ra, hS�ra)) Or Not sSe�im = ssTahmin Then
            For oS�ra = 0 To 8
                H�creOlas�(bS�ra, hS�ra, oS�ra) = True
            Next
        End If
        If sSe�im = ssBulmacaD���nda Then
            If bH�cre(bS�ra, hS�ra) Then B�l�m(bS�ra).H�cre hS�ra, , hbGiri�
        ElseIf sSe�im = ssT�m� Then
            B�l�m(bS�ra).H�cre hS�ra, hRakamBo�, hbGiri�
        End If
        If Not sSe�im = ssTahmin Then bH�cre(bS�ra, hS�ra) = False
        tH�cre(bS�ra, hS�ra) = False
    Next
Next

End Sub

Private Sub Tu�Kilidi(tkYeni As Boolean)
If tkYeni Then
    cmdB��z.Enabled = True
    cmdG�ster.Enabled = False
Else
    cmdB��z.Enabled = False
    cmdG�ster.Enabled = True
End If
End Sub

Private Sub Form_Load()
Dim S�ra As Byte
Const yrlNesneAra As Single = 240  'Form �zerindeki nesnelerin birbirine uzakl���
Dim yrlTu�Ara As Single

frmHakk�nda.PAd� = "O� Sudoku ��zer 1.1"
App.Title = frmHakk�nda.PAd�
frm��zer.Caption = frmHakk�nda.PAd�

SudAlan.Left = yrlNesneAra
SudAlan.Top = yrlNesneAra
B�l�m(0).Left = yrlNesneAra
B�l�m(0).Top = yrlNesneAra

For S�ra = 1 To 8
    Load B�l�m(S�ra)
    With B�l�m(S�ra)
        .Visible = True
        If (S�ra + 1) Mod 3 = 1 Then
            .Left = B�l�m(0).Left
            .Top = B�l�m(S�ra - 3).Top + B�l�m(0).Height
        Else
            .Left = B�l�m(S�ra - 1).Left + B�l�m(0).Width
            .Top = B�l�m(S�ra - 1).Top
        End If
    End With
Next

With B�l�m(8)
    SudAlan.Width = .Left + .Width + yrlNesneAra
    SudAlan.Height = .Top + .Height + yrlNesneAra
End With

With cmdYeni
    .Top = SudAlan.Top + SudAlan.Height + yrlNesneAra
    .Left = SudAlan.Left + SudAlan.Width - .Width
    cmdG�ster.Top = .Top
    cmdB��z.Top = .Top
    yrlTu�Ara = .Width + yrlNesneAra
    cmdG�ster.Left = .Left - yrlTu�Ara
    Me.Height = .Top + .Height + 3 * yrlNesneAra
End With
cmdB��z.Left = cmdG�ster.Left - yrlTu�Ara

With cmdHakk�nda
    .Top = SudAlan.Top
    .Left = SudAlan.Left + SudAlan.Width + yrlNesneAra
    cmd��k.Top = .Top + .Height + yrlNesneAra
    cmd��k.Left = .Left
    Me.Width = .Left + yrlTu�Ara
End With


Randomize Timer

End Sub

Private Sub Form_Activate()
If Not FormY�klendi Then
    cmdYeni_Click
    FormY�klendi = True
End If
End Sub

Private Sub cmdYeni_Click()
S�f�rla
B�l�m(0).H�creEtkinle�tir 0
End Sub

Private Sub cmdG�ster_Click()
Dim bS�ra As Byte
Dim hS�ra As Byte


For bS�ra = 0 To 8
    For hS�ra = 0 To 8
        If Not bH�cre(bS�ra, hS�ra) Then B�l�m(bS�ra).H�cre hS�ra, hRakamBo�, hbGiri�
    Next
Next

If Not �Gizle Then H�creleriYaz
�Gizle = Not �Gizle
End Sub

Private Sub cmdHakk�nda_Click()
frmHakk�nda.Show 1
End Sub

Private Sub B�l�m_Giri�Yap�ld�(Index As Integer, hIndex As Byte)
Dim bgY�n As hoTaramaY�n�
Dim bEksenKonum As Byte
Dim hEksenKonum As Byte
Dim teB�l�m As Byte
Dim teH�cre As Byte
Dim bS�ra As Byte
Dim hS�ra As Byte
Const nSon As Byte = 8 'Sonuncu nesne

If B�l�m(Index).H�cre(hIndex) = hRakamBo� Then GoTo SonrakineGe�

For hS�ra = 0 To 8
    If Not hIndex = hS�ra Then
        If B�l�m(Index).H�cre(hIndex) = B�l�m(Index).H�cre(hS�ra) Then
            B�l�m(Index).H�cre hIndex, hRakamBo�
            hGiri�Yanl�� CByte(Index), hS�ra
            Exit Sub
        End If
    End If
Next

For bgY�n = 0 To 1
    bEksenKonum = EksenKonumBul(CByte(Index), bgY�n)
    hEksenKonum = EksenKonumBul(hIndex, bgY�n)
    For teB�l�m = 0 To 2
        bS�ra = sYer��z(bEksenKonum, teB�l�m, bgY�n)
        For teH�cre = 0 To 2
            hS�ra = sYer��z(hEksenKonum, teH�cre, bgY�n)
            If Not (CByte(Index) = bS�ra And hIndex = hS�ra) Then
                If B�l�m(bS�ra).H�cre(hS�ra) = B�l�m(Index).H�cre(hIndex) Then
                    B�l�m(Index).H�cre hIndex, hRakamBo�
                    hGiri�Yanl�� bS�ra, hS�ra
                    Exit Sub
                End If
            End If
        Next
    Next
Next
            
SonrakineGe�:
If Not hIndex = nSon Then
    B�l�m(Index).H�creEtkinle�tir hIndex + 1
ElseIf Not Index = nSon Then
    B�l�m(Index + 1).H�creEtkinle�tir 0
End If


End Sub

Private Sub hGiri�Yanl��(gyB�l�m As Byte, gyH�cre As Byte)
Dim Ba�lamaAn� As Single
Const Durakla As Single = 0.3
B�l�m(gyB�l�m).H�cre gyH�cre, , hbGiri�Yanl��
Ba�lamaAn� = Timer
Beep
Do
    DoEvents
Loop Until Timer - Ba�lamaAn� > Durakla
If bH�cre(gyB�l�m, gyH�cre) Then B�l�m(gyB�l�m).H�cre gyH�cre, , hbBulmaca _
Else B�l�m(gyB�l�m).H�cre gyH�cre, , hbGiri�
End Sub

Private Sub cmd��k_Click()
End
End Sub

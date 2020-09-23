VERSION 5.00
Begin VB.Form frmÇözer 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmÇözer.frx":0000
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
      Begin OÖSudokuÇözer.Bölüm Bölüm 
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
   Begin VB.CommandButton cmdÇýk 
      Caption         =   "Çýk"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdHakkýnda 
      Caption         =   "Hakkýnda"
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
   Begin VB.CommandButton cmdGöster 
      Caption         =   "Gizle/Göster"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdBÇöz 
      Caption         =   "Çöz"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "frmÇözer"
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

'Düþünce tarzýný daha kolay kavramak için "Açýklama.rtf" dosyasýna da gözatabilirsiniz.

Option Explicit

Private Hücre(8, 8) As Byte 'Hücredeki rakamý tutar.
Private HücreOlasý(8, 8, 8) As Boolean 'Hücredeki olasýlýklarý tutar.

Private bHücre(8, 8) As Boolean 'Bulmacada verilen hücrelerin yerlerini tutar.
Private tHücre(8, 8) As Boolean 'Tahmin edilen hücrelerin yerini tutar.

Private Enum hoTaramaYönü
    tyYatay = 0
    tyDüþey = 1
    tyYatayDüþey = 2
End Enum

Private khSayýsý As Byte 'Kesinleþen hücre sayýsý

Private ÇGizle As Boolean

Private Enum SýfýrlaSeçenekleri
    ssTümü
    ssBulmacaDýþýnda
    ssTahmin
End Enum

Private Const hRakamBoþ As Byte = 0 'Hücreye boþ atama yapmak için

Public ÝþlemVar As Byte  '0 bigi penceresi açýlmamýþ, 1 iþlem iptal edilmiþ, 2 iþlem sürüyor

Dim FormYüklendi As Boolean

Private Sub cmdBÇöz_Click()
TuþKilidi False
BulmacayýAl
If Not BulmacayýÇöz Then
    Sýfýrla ssBulmacaDýþýnda
    MsgBox "Bulmaca çözülemedi!", vbCritical
Else
    HücreleriYaz
End If
End Sub

Private Sub BulmacayýAl()
Dim bSýra As Byte 'bölüm sýra
Dim hSýra As Byte 'hücre sýra

Dim oSýra As Byte 'olasýlýk sýra

For bSýra = 0 To 8
    For hSýra = 0 To 8
        If Not Bölüm(bSýra).Hücre(hSýra) = hRakamBoþ Then
            For oSýra = 0 To 8
                HücreOlasý(bSýra, hSýra, oSýra) = False
            Next
            HücreOlasý(bSýra, hSýra, Bölüm(bSýra).Hücre(hSýra) - 1) = True
            Bölüm(bSýra).Hücre hSýra, , hbBulmaca
            bHücre(bSýra, hSýra) = True
        End If
    Next
Next

End Sub

Private Function BulmacayýÇöz() As Boolean
'Tarama yordamlarý, döngünün her adýmýnda, bulmacayý tekrar tekrar baþtan sona tararlar.

Dim ökhSayýsý As Byte 'Önceki khSayýsý
Const khSayýsýTam As Byte = 81
Const bTahminSor As String = "Bu bulmacayý çözmek için tahmin de denemek gerekiyor olabilir." & _
"Tahmin denensin mi?"
Dim TahminOnay As VbMsgBoxResult
Dim BaþlamaAný As Single
Const ZamanAþýmý As Single = 10


Do
    Sýfýrla ssTahmin
    hoKesinleþenleriTara
    Do
        Do
            ökhSayýsý = khSayýsý
            hoBTekEksenTara
            hoFarkBirTara
            hoKesinleþenleriTara
            If khSayýsý = ökhSayýsý Then
                If Not TahminOnay = vbYes Then _
                TahminOnay = MsgBox(bTahminSor, vbYesNo + vbQuestion)
                Exit Do
            End If
        Loop Until khSayýsý = khSayýsýTam
        If khSayýsý = khSayýsýTam Or TahminOnay = vbNo Then GoTo Bitir
    Loop While hoTahmin
    If BaþlamaAný = 0 Then
        BaþlamaAný = Timer
        frmÝþlemVar.Show
    End If
    DoEvents
    If ÝþlemVar = 1 Then GoTo Bitir
Loop While Timer - BaþlamaAný < ZamanAþýmý

Bitir:
If Not ÝþlemVar = 0 Then ÝþlemVar = 0: Unload frmÝþlemVar
If khSayýsý = khSayýsýTam Then BulmacayýÇöz = True

End Function

Private Sub hoKesinleþenleriTara()
'HücreOlasý'da bire inmiþ olasýlýklarýn rakam deðerini Hücre'ye yerleþtirir.

Dim bSýra As Byte
Dim hSýra As Byte

Dim khRakamý As Byte 'kesinleþen hücre rakamý

For bSýra = 0 To 8
    For hSýra = 0 To 8
        If Hücre(bSýra, hSýra) = 0 Then
            khRakamý = hoSaltBir(bSýra, hSýra)
            If Not khRakamý = 0 Then
                Hücre(bSýra, hSýra) = khRakamý
                hoOlasýlýkEle bSýra, hSýra, Hücre(bSýra, hSýra) - 1
                khSayýsý = khSayýsý + 1
            End If
        End If
    Next
Next

End Sub

Private Function hoSaltBir(sbBölüm As Byte, sbHücre As Byte) As Byte
'Belirtilen Hücre'nin olasýlýklarýnýn sadece bir tane olduðunu denetler ve bunun hangi rakam olduðunu.

Dim sbSýra As Integer

For sbSýra = 0 To 8
    If Not hoSaltBir = 0 Then
        If HücreOlasý(sbBölüm, sbHücre, sbSýra) Then hoSaltBir = 0: Exit Function
    Else
        If HücreOlasý(sbBölüm, sbHücre, sbSýra) Then hoSaltBir = sbSýra + 1
    End If
Next

End Function

Private Function sYerÇöz(yçEksenKonum As Byte, yçteNesne As Byte, yçYön As hoTaramaYönü) As Byte
'Sýra Yer Çöz
'Konumlandýðý ekseni ve eksendeki sýrasý verilen nesnenin Index'ini verir.
If yçYön = tyYatay Then sYerÇöz = yçEksenKonum * 3 + yçteNesne _
Else sYerÇöz = yçEksenKonum + yçteNesne * 3
End Function

Private Function EksenKonumBul(ebNesne As Byte, ebYön As hoTaramaYönü) As Byte
''Index'i verilen nesnenin konumlandýðý ekseni verir.

Dim teNesne As Byte
If ebYön = tyYatay Then
    Do While EksenKonumBul * 3 < ebNesne + 1
        EksenKonumBul = EksenKonumBul + 1
    Loop
    EksenKonumBul = EksenKonumBul - 1
Else
    For EksenKonumBul = 0 To 2
        For teNesne = 0 To 2
            If sYerÇöz(EksenKonumBul, teNesne, tyDüþey) = ebNesne Then Exit Function
        Next
    Next
End If

End Function

Private Function hoOlasýlýkEle(oeBölüm As Byte, oeHücre As Byte, oeElenen As Byte, _
                                                    Optional oeBölümüTarama As Boolean = False, _
                                                    Optional oeYön As hoTaramaYönü = tyYatayDüþey) _
                                                    As Boolean

'En çok olasýlýðý bire inmiþ hücrenin bulunduðu bölüm ve eksenlerdeki hücrelerde, bu kesinleþen olasýlýðý elemek amacýyla _
kullanýlýyor. BTekEksenTara yordamýna bakarsanýz, olasýlýðýn eksende kesinlik kazandýðýnda da kullanýldýðýný göreceksiniz.

Dim bSýra As Byte
Dim hSýra As Byte

Dim hEksenKonum As Byte
Dim bEksenKonum As Byte
Dim teHücre As Byte 'Taranan eksendeki Hücre
Dim teBölüm As Byte

'Ayný Bölüm Kaynaklý Elemeler
If Not oeBölümüTarama Then
    For hSýra = 0 To 8
        If Not hSýra = oeHücre Then
            If HücreOlasý(oeBölüm, hSýra, oeElenen) Then
                If Not hoOlasýlýkEle Then hoOlasýlýkEle = True
                HücreOlasý(oeBölüm, hSýra, oeElenen) = False
            End If
        End If
    Next
End If

'Yatay kaynaklý elemeler
If oeYön = tyYatayDüþey Or oeYön = tyYatay Then
    bEksenKonum = EksenKonumBul(oeBölüm, tyYatay)
    hEksenKonum = EksenKonumBul(oeHücre, tyYatay)
    
    For teBölüm = 0 To 2
        bSýra = sYerÇöz(bEksenKonum, teBölüm, tyYatay)
        For teHücre = 0 To 2
            hSýra = sYerÇöz(hEksenKonum, teHücre, tyYatay)
            If Not (oeBölümüTarama And bSýra = oeBölüm) Then
                If Not (bSýra = oeBölüm And hSýra = oeHücre) Then
                    If HücreOlasý(bSýra, hSýra, oeElenen) Then
                        If Not hoOlasýlýkEle Then hoOlasýlýkEle = True
                        HücreOlasý(bSýra, hSýra, oeElenen) = False
                    End If
                End If
            End If
        Next
    Next
End If

If oeYön = tyYatayDüþey Or oeYön = tyDüþey Then
    'Düþey kaynaklý elemeler
    bEksenKonum = EksenKonumBul(oeBölüm, tyDüþey)
    hEksenKonum = EksenKonumBul(oeHücre, tyDüþey)
    
    For teBölüm = 0 To 2
        bSýra = sYerÇöz(bEksenKonum, teBölüm, tyDüþey)
        For teHücre = 0 To 2
            hSýra = sYerÇöz(hEksenKonum, teHücre, tyDüþey)
            If Not (oeBölümüTarama And bSýra = oeBölüm) Then
                If Not (bSýra = oeBölüm And hSýra = oeHücre) Then
                    If HücreOlasý(bSýra, hSýra, oeElenen) Then
                        If Not hoOlasýlýkEle Then hoOlasýlýkEle = True
                        HücreOlasý(bSýra, hSýra, oeElenen) = False
                    End If
                End If
            End If
        Next
    Next
End If

End Function

Private Sub hoFarkBirTara()
'Bir hücrenin olasýlýklarýndan biri, bulunduðu eksen ve ya bölümdeki diðer hücrelerde yoksa _
bu hücrenin tek olasýlýðýnýn bu olduðu çýkarýlabilir.

Dim bSýra As Byte
Dim hSýra As Byte

Dim hEksenKonum As Byte
Dim bEksenKonum As Byte
Dim teHücre As Byte 'Taranan eksendeki Hücre
Dim teBölüm As Byte

Dim fbYön As hoTaramaYönü '0 yatay, 1 dikey
Dim fbOlasý As Byte
Dim bSýraAyný As Byte
Dim hSýraAyný As Byte ''Taranan hücrenin bir olasýlýðýnýn, ayný _
olasýlýða sahip mi diye karþýlaþtýrýldýðý hücrenin yerini belirliyor.
Dim tebSýraAyný As Byte
Dim tehSýraAyný As Byte
Dim fbOlasýlýkSil As Byte


'Bölümlerde farklý hücreler
For bSýra = 0 To 8
    For hSýra = 0 To 8
        If Hücre(bSýra, hSýra) = 0 Then
            For fbOlasý = 0 To 8
                If HücreOlasý(bSýra, hSýra, fbOlasý) Then
                    For hSýraAyný = 0 To 8
                        If Hücre(bSýra, hSýraAyný) = 0 Then
                            If Not hSýra = hSýraAyný Then
                                If HücreOlasý(bSýra, hSýraAyný, fbOlasý) Then GoTo fbbDiðerOlasý
                            End If
                        End If
                    Next
                    'Hücrenin diðerlerinde olmayan bir olasýlýða sahip olduðu ortaya çýktý.
                    For fbOlasýlýkSil = 0 To 8
                        HücreOlasý(bSýra, hSýra, fbOlasýlýkSil) = False
                    Next
                    HücreOlasý(bSýra, hSýra, fbOlasý) = True
                    GoTo fbbDiðerHücre
                End If
fbbDiðerOlasý:
            Next
        End If
fbbDiðerHücre:
    Next
Next


'Yatay ve Düþey'de bir farklý hücreler
For fbYön = 0 To 1
    For bEksenKonum = 0 To 2
        For hEksenKonum = 0 To 2
            For teBölüm = 0 To 2
                bSýra = sYerÇöz(bEksenKonum, teBölüm, fbYön)
                For teHücre = 0 To 2
                    hSýra = sYerÇöz(hEksenKonum, teHücre, fbYön)
                    If Hücre(bSýra, hSýra) = 0 Then
                        For fbOlasý = 0 To 8
                            If HücreOlasý(bSýra, hSýra, fbOlasý) Then
                                For tebSýraAyný = 0 To 2
                                    bSýraAyný = sYerÇöz(bEksenKonum, tebSýraAyný, fbYön)
                                    For tehSýraAyný = 0 To 2
                                        hSýraAyný = sYerÇöz(hEksenKonum, tehSýraAyný, fbYön)
                                        If Hücre(bSýraAyný, hSýraAyný) = 0 Then
                                            If Not (bSýra = bSýraAyný And hSýra = hSýraAyný) Then
                                                If HücreOlasý(bSýraAyný, hSýraAyný, fbOlasý) Then GoTo fbdDiðerOlasý
                                            End If
                                        End If
                                    Next
                                Next
                                'Hücrenin diðerlerinde olmayan bir olasýlýða sahip olduðu ortaya çýktý.
                                For fbOlasýlýkSil = 0 To 8
                                    HücreOlasý(bSýra, hSýra, fbOlasýlýkSil) = False
                                Next
                                HücreOlasý(bSýra, hSýra, fbOlasý) = True
                                GoTo fbdDiðerHücre
                            End If
fbdDiðerOlasý:
                        Next
                    End If
fbdDiðerHücre:
                Next
            Next
        Next
    Next
Next

End Sub

Private Sub hoBTekEksenTara()
'Bir bölümde sadece tek eksende bulunabilecek olasýlýklarý tarar.

Dim bSýra As Byte
Dim hSýra As Byte

Dim hoOlasýlýkElendi As Boolean

Dim hEksenKonum As Byte
Dim bEksenKonum As Byte
Dim teHücre As Byte 'Taranan eksendeki Hücre
Dim teBölüm As Byte
Dim bteYön As hoTaramaYönü


Dim dhEksenKonum As Byte 'Diðer hücre eksen konum
Dim bteOlasý As Byte
Dim dteHücre As Byte
Dim dhSýra As Byte


Do
hoOlasýlýkElendi = False

'Bir bölümün sadece tek bir hücre eksenine gelebilecek olasýlýklar varsa ayný eksendeki diðer bölümlerin o hücre _
eksenlerindeki olasýlýklar elenir. Örneðin;
'Bölüm(1)'in 2. hücre ekseni dýþýnda 3 rakamýnýn gelemeyeceði ortaya çýktýysa, Bölüm(0 ve 2)'in 2. hücre eksenlerindeki _
3 rakamý olasýlýðý elenir.
For bteYön = 0 To 1
    For bSýra = 0 To 8
        For hEksenKonum = 0 To 2
            For teHücre = 0 To 2
                hSýra = sYerÇöz(hEksenKonum, teHücre, bteYön)
                If Hücre(bSýra, hSýra) = 0 Then
                    For bteOlasý = 0 To 8
                        If HücreOlasý(bSýra, hSýra, bteOlasý) Then
                            For dhEksenKonum = 0 To 2
                                If Not hEksenKonum = dhEksenKonum Then
                                    For dteHücre = 0 To 2
                                        dhSýra = sYerÇöz(dhEksenKonum, dteHücre, bteYön)
                                        If HücreOlasý(bSýra, dhSýra, bteOlasý) Then GoTo bteDiðerOlasý
                                    Next
                                End If
                            Next
                            If hoOlasýlýkEle(bSýra, hSýra, bteOlasý, True, bteYön) Then hoOlasýlýkElendi = True
                        End If
bteDiðerOlasý:
                    Next
                End If
            Next
        Next
    Next
Next

Loop While hoOlasýlýkElendi

End Sub

Private Function hoTahmin() As Boolean
Dim bSýra As Byte
Dim hSýra As Byte
Dim oSýra As Byte
Dim tOlasý As Byte 'Tahmin edilen olasýlýk
Dim dOlasý() As Boolean 'Denenmiþ olasýlýklar, hücrede hiç olasýlýk kalmadýðýný anlamak için.
Dim dOlasýSýra As Byte
Dim OlasýlýkSil As Byte

For bSýra = 0 To 8
    For hSýra = 0 To 8
        If Hücre(bSýra, hSýra) = 0 Then
            ReDim dOlasý(8) As Boolean
            Do
                For dOlasýSýra = 0 To 8
                    If Not dOlasý(dOlasýSýra) Then Exit For
                Next
                If dOlasýSýra = 8 + 1 Then Exit Function 'Bütün olasýlýklar denenmiþ.
                oSýra = Rnd * 8
                dOlasý(oSýra) = True
            Loop Until HücreOlasý(bSýra, hSýra, oSýra)
            For OlasýlýkSil = 0 To 8
                HücreOlasý(bSýra, hSýra, OlasýlýkSil) = False
            Next
            HücreOlasý(bSýra, hSýra, oSýra) = True
            tHücre(bSýra, hSýra) = True
            hoTahmin = True
            Exit Function
        End If
    Next
Next

End Function

Private Sub HücreleriYaz()

Dim bSýra As Byte
Dim hSýra As Byte

For bSýra = 0 To 8
    For hSýra = 0 To 8
            If Not bHücre(bSýra, hSýra) Then
                If tHücre(bSýra, hSýra) Then Bölüm(bSýra).Hücre hSýra, Hücre(bSýra, hSýra), hbTahmin _
                Else: Bölüm(bSýra).Hücre hSýra, Hücre(bSýra, hSýra), hbÇözüm
            End If
    Next
Next

End Sub

Private Sub Sýfýrla(Optional sSeçim As SýfýrlaSeçenekleri = ssTümü)
Dim bSýra As Byte
Dim hSýra As Byte
Dim oSýra As Byte

khSayýsý = 0

If Not sSeçim = ssTahmin Then
    ÇGizle = True
    TuþKilidi True
End If

For bSýra = 0 To 8
    For hSýra = 0 To 8
        Hücre(bSýra, hSýra) = 0
        If (sSeçim = ssTahmin And Not bHücre(bSýra, hSýra)) Or Not sSeçim = ssTahmin Then
            For oSýra = 0 To 8
                HücreOlasý(bSýra, hSýra, oSýra) = True
            Next
        End If
        If sSeçim = ssBulmacaDýþýnda Then
            If bHücre(bSýra, hSýra) Then Bölüm(bSýra).Hücre hSýra, , hbGiriþ
        ElseIf sSeçim = ssTümü Then
            Bölüm(bSýra).Hücre hSýra, hRakamBoþ, hbGiriþ
        End If
        If Not sSeçim = ssTahmin Then bHücre(bSýra, hSýra) = False
        tHücre(bSýra, hSýra) = False
    Next
Next

End Sub

Private Sub TuþKilidi(tkYeni As Boolean)
If tkYeni Then
    cmdBÇöz.Enabled = True
    cmdGöster.Enabled = False
Else
    cmdBÇöz.Enabled = False
    cmdGöster.Enabled = True
End If
End Sub

Private Sub Form_Load()
Dim Sýra As Byte
Const yrlNesneAra As Single = 240  'Form üzerindeki nesnelerin birbirine uzaklýðý
Dim yrlTuþAra As Single

frmHakkýnda.PAdý = "OÖ Sudoku Çözer 1.1"
App.Title = frmHakkýnda.PAdý
frmÇözer.Caption = frmHakkýnda.PAdý

SudAlan.Left = yrlNesneAra
SudAlan.Top = yrlNesneAra
Bölüm(0).Left = yrlNesneAra
Bölüm(0).Top = yrlNesneAra

For Sýra = 1 To 8
    Load Bölüm(Sýra)
    With Bölüm(Sýra)
        .Visible = True
        If (Sýra + 1) Mod 3 = 1 Then
            .Left = Bölüm(0).Left
            .Top = Bölüm(Sýra - 3).Top + Bölüm(0).Height
        Else
            .Left = Bölüm(Sýra - 1).Left + Bölüm(0).Width
            .Top = Bölüm(Sýra - 1).Top
        End If
    End With
Next

With Bölüm(8)
    SudAlan.Width = .Left + .Width + yrlNesneAra
    SudAlan.Height = .Top + .Height + yrlNesneAra
End With

With cmdYeni
    .Top = SudAlan.Top + SudAlan.Height + yrlNesneAra
    .Left = SudAlan.Left + SudAlan.Width - .Width
    cmdGöster.Top = .Top
    cmdBÇöz.Top = .Top
    yrlTuþAra = .Width + yrlNesneAra
    cmdGöster.Left = .Left - yrlTuþAra
    Me.Height = .Top + .Height + 3 * yrlNesneAra
End With
cmdBÇöz.Left = cmdGöster.Left - yrlTuþAra

With cmdHakkýnda
    .Top = SudAlan.Top
    .Left = SudAlan.Left + SudAlan.Width + yrlNesneAra
    cmdÇýk.Top = .Top + .Height + yrlNesneAra
    cmdÇýk.Left = .Left
    Me.Width = .Left + yrlTuþAra
End With


Randomize Timer

End Sub

Private Sub Form_Activate()
If Not FormYüklendi Then
    cmdYeni_Click
    FormYüklendi = True
End If
End Sub

Private Sub cmdYeni_Click()
Sýfýrla
Bölüm(0).HücreEtkinleþtir 0
End Sub

Private Sub cmdGöster_Click()
Dim bSýra As Byte
Dim hSýra As Byte


For bSýra = 0 To 8
    For hSýra = 0 To 8
        If Not bHücre(bSýra, hSýra) Then Bölüm(bSýra).Hücre hSýra, hRakamBoþ, hbGiriþ
    Next
Next

If Not ÇGizle Then HücreleriYaz
ÇGizle = Not ÇGizle
End Sub

Private Sub cmdHakkýnda_Click()
frmHakkýnda.Show 1
End Sub

Private Sub Bölüm_GiriþYapýldý(Index As Integer, hIndex As Byte)
Dim bgYön As hoTaramaYönü
Dim bEksenKonum As Byte
Dim hEksenKonum As Byte
Dim teBölüm As Byte
Dim teHücre As Byte
Dim bSýra As Byte
Dim hSýra As Byte
Const nSon As Byte = 8 'Sonuncu nesne

If Bölüm(Index).Hücre(hIndex) = hRakamBoþ Then GoTo SonrakineGeç

For hSýra = 0 To 8
    If Not hIndex = hSýra Then
        If Bölüm(Index).Hücre(hIndex) = Bölüm(Index).Hücre(hSýra) Then
            Bölüm(Index).Hücre hIndex, hRakamBoþ
            hGiriþYanlýþ CByte(Index), hSýra
            Exit Sub
        End If
    End If
Next

For bgYön = 0 To 1
    bEksenKonum = EksenKonumBul(CByte(Index), bgYön)
    hEksenKonum = EksenKonumBul(hIndex, bgYön)
    For teBölüm = 0 To 2
        bSýra = sYerÇöz(bEksenKonum, teBölüm, bgYön)
        For teHücre = 0 To 2
            hSýra = sYerÇöz(hEksenKonum, teHücre, bgYön)
            If Not (CByte(Index) = bSýra And hIndex = hSýra) Then
                If Bölüm(bSýra).Hücre(hSýra) = Bölüm(Index).Hücre(hIndex) Then
                    Bölüm(Index).Hücre hIndex, hRakamBoþ
                    hGiriþYanlýþ bSýra, hSýra
                    Exit Sub
                End If
            End If
        Next
    Next
Next
            
SonrakineGeç:
If Not hIndex = nSon Then
    Bölüm(Index).HücreEtkinleþtir hIndex + 1
ElseIf Not Index = nSon Then
    Bölüm(Index + 1).HücreEtkinleþtir 0
End If


End Sub

Private Sub hGiriþYanlýþ(gyBölüm As Byte, gyHücre As Byte)
Dim BaþlamaAný As Single
Const Durakla As Single = 0.3
Bölüm(gyBölüm).Hücre gyHücre, , hbGiriþYanlýþ
BaþlamaAný = Timer
Beep
Do
    DoEvents
Loop Until Timer - BaþlamaAný > Durakla
If bHücre(gyBölüm, gyHücre) Then Bölüm(gyBölüm).Hücre gyHücre, , hbBulmaca _
Else Bölüm(gyBölüm).Hücre gyHücre, , hbGiriþ
End Sub

Private Sub cmdÇýk_Click()
End
End Sub

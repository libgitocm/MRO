VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_ZAZNAM 
   Caption         =   "DLR / NW"
   ClientHeight    =   10335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14910
   OleObjectBlob   =   "frm_ZAZNAM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_ZAZNAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===== MRO 2015 =====

Dim vyuzito     As String       'pro edit reklamy - �daj, zda byla ji� vyu�it� pro kampa�
Dim cisRadek    As String       '��dek ��seln�ku pro na�ten�
Dim mesicPuvodni As Integer     'pro z�pis zm�n m�s�ce (evidence dokl., slo�ky!)
'Dim dlr         As Long         'index DLRa - pro hled�n� v pol�ch - je PUBLIC!!
Dim cerpaniP    As Single   'pro okno kontroly �erp�n�
Dim cerpaniS    As Single   '(na�tou se �erp�n� z nezam�tnut�ch z�znam�)



Private Sub UserForm_Initialize()

    'poloha a rozm�ry frm:
    '---------------------
    With Me
    .StartUpPosition = 0
    .Width = 730
    .Height = 434
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    If reklama = True Then
        .Caption = Range("A1").Value & " - " & Range("B1").Value & "  (" & Range("O1").Value & ")"
    Else
        .Caption = Range("A1").Value & " - " & Range("B1").Value & "  (" & Range("G1").Value & ")"
    End If
    If admin = True Then .Caption = .Caption & "     [-ADMIN-]"

    soubor = ActiveWorkbook.Name    'jm�no souboru pro pr�ci se spole�nou reklamou
    '(spol. rekl. - Storno = prom�nn� pro aktivaci souboru DLR)
    spolReklama = False     'default
    ciselnikZaznam = True   'zp�tn� vazba z frmCiselnik do frm_ZAZNAM - na�te se editovan� ��seln�k

    Set wsDlrData1 = ActiveWorkbook.Worksheets(1)                   'aktu�ln� objekt wsDlrData1 - list ��dost�
    Set wsDlrData2 = ActiveWorkbook.Worksheets(2)                   'aktu�ln� objekt wsDlrData2 - list reklam

    'loga instalac�:
    .imgSkodaPlus.Visible = skodaPlus

    'ikona TC:
    If cestaTcmd <> "" Then
        .imgTcmd.Visible = True
    Else
        .imgTcmd.Visible = False
    End If

    'podle NW vyhledat DLR a z�skat index: (pro kontroly �erp�n�, vyu�it� bonus� a z�pis do evidence doklad�)
    For i = LBound(dlrNw) To UBound(dlrNw)
        If Range("A1").Value = dlrNw(i) Then
            dlr = i
            GoTo KonecNw
        End If
    Next i
KonecNw:

    're�im REKLAMA:
    If reklama = True Then
        Call ZaznamRezimReklama
    End If

    'nastaven� pro extern� MRO:
    '--------------------------
    .lblHelp1.Visible = Not externi
    .lblCiselnik1.Visible = Not externi
    .lblCiselnik2.Visible = Not externi
    .lblCiselnik3.Visible = Not externi
    .lblCiselnik4.Visible = Not externi
    .lblCiselnik5.Visible = Not externi
    .lblCiselnik6.Visible = Not externi
    .chkBonus.Visible = Not externi

    're�im KAMPA� - detekce p�evzet� reklamy:
    idReklamy = -100    'pokud se nezm�n�, p�i ulo�en� z�znamu se neukl�d� vyu�it� reklamy
    idBonusy = -100

    'v�b�r listu - prov�d� se v frm_MENU.KontrolaListu

    'label Neaktivn�:
    If UCase(Range("G2").Value) = "PLATN�: NE" Or UCase(Range("O2").Value) = "PLATN�: NE" Then
        .lblNeaktivni.Visible = True
        .cmdUlozit.Enabled = False
    Else
        .lblNeaktivni.Visible = False
    End If

    'Bonus:
    .lblAlertBonus.Visible = .chkBonus.Value
    .cmdPrevzitBonus.Visible = .chkBonus.Value

    'p�evzat� bonus:
    typBonus = ""   'default!
    .lblZdrojovyBonus.Visible = False   'zapne se a� v�b�rem IP
    .lblTypBonus.Visible = False
    .lblTypBonus.Caption = typBonus

    'label Spole�n� rekl:
    .lblAlertSpolecna.Visible = .chkSpolReklama.Value

    '�ipka pro kopii reg. ��sla a ��sla ��dosti:
    .lblKopieCisel.Visible = False    'zapne se pouze pro EDIT kampan�

    'option Kampa�:
    .optKamImport.Value = True       'default

    'n�zvy kampan�:
    Call NacistCiselnikKamNazev

    'combo Typ kampan�:
    Call NacistCiselnikKamTyp

    'combo Zam��en� kampan�:
    Call NacistCiselnikKamZamereni

    'combo Medium-typ:
    Call NacistCiselnikKamMedium

    'combo Medium-nazev:
    Call NacistCiselnikKamMediumNazev

    'combo Zdroj:
    With .cmbKamZdroj
    .Clear
    .AddItem "DAS"
    .AddItem "B2B"
    .AddItem "vlastn�"
    End With
'    .chkDas.Visible = False

    'combo Form�t:
    Call NacistCiselnikKamFormat

    'popisky hodnocen�:
    '------------------
    .chkHodn1.Caption = cisHodnoceni(1)
    .chkHodn2.Caption = cisHodnoceni(2)
    .chkHodn3.Caption = cisHodnoceni(3)
    .chkHodn4.Caption = cisHodnoceni(4)
    .chkHodn5.Caption = cisHodnoceni(5)
    .chkHodn6.Caption = cisHodnoceni(6)
    '------------------

    'viditelnost popisk� hodnocen� podle nastaven�:
    '----------------------------------------------
    '(chkbox se skryje a nastav� na True, aby fungoval logick� sou�in pro schv�len�)
    If cisHodnoceni(1) = "X" Then
        With .chkHodn1
        .Value = True
        .Visible = False
        End With
        .txtHodn1.Visible = False
    End If

    If cisHodnoceni(2) = "X" Then
        With .chkHodn2
        .Value = True
        .Visible = False
        End With
        .txtHodn2.Visible = False
    End If

    If cisHodnoceni(3) = "X" Then
        With .chkHodn3
        .Value = True
        .Visible = False
        End With
        .txtHodn3.Visible = False
    End If

    If cisHodnoceni(4) = "X" Then
        With .chkHodn4
        .Value = True
        .Visible = False
        End With
        .txtHodn4.Visible = False
    End If

    If cisHodnoceni(5) = "X" Then
        With .chkHodn5
        .Value = True
        .Visible = False
        End With
        .txtHodn5.Visible = False
    End If

    If cisHodnoceni(6) = "X" Then
        With .chkHodn6
        .Value = True
        .Visible = False
        End With
        .txtHodn6.Visible = False
    End If
    '----------------------------------------------

    'pozn�mka k uz�v�rce:
    .txtPoznamka.Value = ""

    'pozn�mka ke schv�len� reklamy:
    .txtReklPozn.Value = ""

    'opt Schv�lit:
    .optSchvalit.Enabled = False

    're�im ADMIN - mo�nost schv�len� kampan� adminem:
    With .chkKamAdmin
    .Visible = admin
    .Value = False
    .Enabled = False    'zapne se a� p�i schv�len�/zam�tnut� kam.!!
    End With
    .lblKamAdmin.Visible = admin

    'p�ehled �erp�n�:
    '----------------
    'rozpo�ty:
    If dlrCinnost(dlr) = "prodej" Then
'        .lblRozpP.Caption = dlrRozpocetProdej(dlr)
'        .lblRozpS.Caption = dlrRozpocetServis(dlr)
        .lblRozpP.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetProdej(dlr), 2)
        .lblRozpP.Caption = Application.WorksheetFunction.Substitute(.lblRozpP.Caption, ",", " ")
        .lblRozpP.Caption = Application.WorksheetFunction.Substitute(.lblRozpP.Caption, ".", ",")

        .lblRozpS.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetServis(dlr), 2)
        .lblRozpS.Caption = Application.WorksheetFunction.Substitute(.lblRozpS.Caption, ",", " ")
        .lblRozpS.Caption = Application.WorksheetFunction.Substitute(.lblRozpS.Caption, ".", ",")

    Else    'servis a Plus
        .lblRozpP.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetFinal(dlr), 2)
        .lblRozpP.Caption = Application.WorksheetFunction.Substitute(.lblRozpP.Caption, ",", " ")
        .lblRozpP.Caption = Application.WorksheetFunction.Substitute(.lblRozpP.Caption, ".", ",")
        .lblRozpS.Caption = ""
        .lblProdej.Enabled = False
        .lblServis.Enabled = False
    End If

    '�erp�n� p�ed zapisovanou kampan�:
    cerpaniP = 0
    cerpaniS = 0
    poslRadek = Cells(Rows.Count, 2).End(xlUp).Row
    For i = 4 To poslRadek
        If Cells(i, 13).Value = "ne" And Cells(i, 14).Value <> "ne" Then   'pouze NEbonus a NEzam�stnut�
            'na��st podle oblasti �erp�n�:
            If Cells(i, 8).Value = "prodej" Then
                cerpaniP = cerpaniP + Cells(i, 11).Value
            Else
                cerpaniS = cerpaniS + Cells(i, 11).Value
            End If
        End If
    Next i

    Call VypocetCeny    ' spustit pro v�po�et z�statku p�ed kampan�
    If typZaznamu = "edit" Then
        .lblZustP1.Visible = False
        .lblZustS1.Visible = False
        .lblKc2.Visible = False
        .lblKc5.Visible = False
    End If

    '�erp�n� po kampani:
    .lblZustP2.Caption = ""
    .lblZustS2.Caption = ""
    
    '----------------------
    'konec p�ehledu �erp�n�

    End With    'Me

  'nastaven� FRM podle typu z�znamu:
  '---------------------------------
  
  'NOVY:
  '-----
    If typZaznamu = "novy" Then
      '==================
        With Me

        Call GenerovatCisloKampane
'        If reklama = False Then Call GenerovatDatum     'pouze pro kampan�!!

        .txtDatum.SetFocus
'        If reklama = False Then SendKeys ("^a")   'ozna�it datum

        If skodaPlus = False Then
            .txtKamProcento.Value = prispevek
        Else
            Call NajitPenetraci     'podle NW naj�t penetraci a p�i�adit % p��sp�vku
        End If
        .txtKamSchval.Value = 0      'v�choz� p��sp�vky
        .txtKamNeschval.Value = 0
        .txtKamPocet.Value = 1    'v�choz� po�et mat. do�l�ch ke schv�len�

        'zkop�rovat �. ��dosti a reg. �. RF z posledn�ho ��dku:
        '(poslRadek vrac� u� GenerovatCisloKampane)
        .txtRegCislo.Value = Cells(poslRadek, 3).Value    'reg. �. RF
        .txtDlrZadost.Value = Cells(poslRadek, 4).Value      '�. ��dosti

        'nastaven� podle �innosti DLR:
        If reklama = False Then
            If Range("G1").Value = "�innost: servis" Then
                .optServis.Value = True
            ElseIf Range("G1").Value = "�innost: �KODA Plus" Then
                .optProdej.Value = True
            End If
        Else
            If Range("O1").Value = "�innost: servis" Then
                .optServis.Value = True
            ElseIf Range("O1").Value = "�innost: �KODA Plus" Then
                .optProdej.Value = True
            End If
        End If

        'z�pis reklamy - autom. p��znak spole�n� reklama:
        If reklama = True And InStr(ActiveWorkbook.Name, "999SR") > 0 Then
            With .chkSpolReklama
            .Value = True
            .Enabled = False
            End With
        End If
        .lblUser.Caption = userAp
        End With    'Me

        'nastaven� pro extern� MRO:
        '--------------------------
        With Me
        If externi = True Then
            .optServis.Value = True
            .cmbKamNazev.ListIndex = 0
            .cmbKamTyp.ListIndex = 0
            .cmbKamZamereni.ListIndex = 0
            .cmbKamZdroj.Value = "DAS"
        End If
        End With

  'KOPIE:
  '------
    ElseIf typZaznamu = "kopie" Then
          '====================
        With Me

        radek = ActiveCell.Row
        Call GenerovatCisloKampane
        Call GenerovatDatum

        'na��st ��dek do FRM - kampa�:
        '-----------------------------
        Call NacistRadek        'na��st polo�ky spole�n� pro KOPII a EDIT

        If reklama = False Then
            If Cells(radek, 33).Value = "ano" Then          'kombin. �erp�n�
                With .chkKombinace
                .Value = True
                .Locked = True
                End With
            Else
                .chkKombinace.Value = False
            End If
        End If

        .txtKamPocet.Value = 1    'v�choz� po�et mat. do�l�ch ke schv�len�

        'zkop�rovat �. ��dosti a reg. �. RF z posledn�ho ��dku:
        '(poslRadek vrac� u� GenerovatCisloKampane)
        If MsgBox("Chcete zachovat �daje ze zdrojov� kampan�?      " & vbNewLine & vbNewLine _
            & "(Datum,  Reg. ��slo,  �. ��dosti,  �. subdod. faktury)      ", 36, "Kopie z�znamu") = vbYes Then
            .txtDatum.Value = Day(Cells(radek, 1).Value) & "." & Month(Cells(radek, 1).Value) & "." & Year(Cells(radek, 1).Value)
            .txtRegCislo.Value = Cells(radek, 3).Value    'reg. �. RF
            .txtDlrZadost.Value = Cells(radek, 4).Value      '�. ��dosti
            .txtSubFak.Value = Cells(radek, 6).Value      '�. sub. fak.
        End If

        With .txtDatum
'        .Value = datum
        .SetFocus
        End With

        'kopie z�znamu generovan�ho ze schv�len� reklamy:
        If InStr(Cells(radek, 36).Value, "-") > 0 Then  'existuje ��slo zdrojov� reklamy RF
            Dim poslRadekRekl   As Long
            Dim radekRekl       As Long
            Dim cisloRekl       As String
            cisloRekl = wsDlrData1.Cells(radek, 36).Value

            'kopie z vyu�it� spole�n� reklamy - zm�n� se list reklam DLR na 999SR:
            If Left(cisloRekl, 5) = "999SR" Then
                spolReklama = True
                'otev��t soubor SR podle instalace:
                Call OtevritSpolecnouReklamu
                Set wsDlrData2 = ActiveWorkbook.Worksheets(2)   'p�ep�e se p�vodn� prom�nn� - z listu DLR na 999SR!
            End If

            'd�le stejn� p��kazy pro kopii z vyu�it� reklamy DLR i 999SR:
            poslRadekRekl = wsDlrData2.Cells(Rows.Count, 2).End(xlUp).Row   'wsDlrData2 ukazuje podle na DLR nebo SR!!
            'vyhled�n� zdrojov� reklamy:
            For i = 4 To poslRadekRekl
                If cisloRekl = wsDlrData2.Cells(i, 2).Value Then
                    idReklamy = i - 4   'pro z�pis vyu�it� reklamy

                    'info o p�evzet� reklamy do frm:
                    With .lblZdrojovaReklama
                    .Caption = "Zdrojov� reklama: " & wsDlrData2.Cells(i, 2).Value & "."
                    .Visible = True
                    If wsDlrData2.Cells(i, 34).Value = "ano" Then
                        .Caption = .Caption & " Ji� vyu�ito: " & wsDlrData2.Cells(i, 35).Value & "x."
                    Else
                        .Caption = .Caption & " Je�t� nevyu�ito."
                    End If
                    End With    'lblZdrojovaReklama
                    .chkHodn6.Value = True
                    GoTo KonecReklamy
                End If
            Next i

KonecReklamy:
            'kopie z vyu�it� spole�n� reklamy - n�vrat k souboru/listu reklam DLR:
            If spolReklama = True Then
                ActiveWorkbook.Close                            'zav��t 999SR
                Windows(soubor).Activate                        'aktivace souboru DLR
                Set wsDlrData2 = ActiveWorkbook.Worksheets(2)   'n�vrat prom�nn� na list reklam DLR!
            End If

        End If  'konec kopie ze schv�len� reklamy
        .lblUser.Caption = userAp
        End With    'Me

  'EDIT:
  '-----
    ElseIf typZaznamu = "edit" Then
          '===================
        With Me

        radek = ActiveCell.Row

        '�ipka pro kopii reg. ��sla a ��sla ��dosti:
        .lblKopieCisel.Visible = True    'zapne se pouze pro EDIT kampan�

        'na��st ��dek do FRM - kampa�:
        '-----------------------------
        .txtDatum.Value = Day(Cells(radek, 1).Value) & "." & Month(Cells(radek, 1).Value) & "." & Year(Cells(radek, 1).Value)
        'pro sledov�n� zm�n m�s�ce v z�znamu:
        If reklama = False Then
            mesicPuvodni = Month(Cells(radek, 1).Value)
        End If

        .txtKamCislo.Value = Cells(radek, 2).Value   '�. kampan�
        .txtRegCislo.Value = Cells(radek, 3).Value   'reg. ��slo RF
        .txtDlrZadost.Value = Cells(radek, 4).Value  '�. ��dosti dlr
        .txtSubFak.Value = Cells(radek, 6).Value     '�. sub. fa
        Call NacistRadek        'na��st polo�ky spole�n� pro KOPII a EDIT

        If reklama = False Then
            If Cells(radek, 33).Value = "ano" Then .chkKombinace.Value = True   'kombin. �erp�n� pro KAMPAN�
        End If
        End With    'Me
        'na��st ��dek do FRM - hodnocen�:
        '--------------------------------
        On Error Resume Next    'ignorovat chyby �ten� pr�zdn�ch koment���

        'stav:
        With Cells(radek, 22)
        If .Value = "ano" Then
            Me.chkHodn1.Value = True
        Else
            Me.chkHodn1.Value = False
        End If
        'koment��:
        Me.txtHodn1.Value = .Comment.Text
        End With

        'stav:
        With Cells(radek, 23)
        If .Value = "ano" Then
            Me.chkHodn2.Value = True
        Else
            Me.chkHodn2.Value = False
        End If
        'koment��:
        Me.txtHodn2.Value = .Comment.Text
        End With

        'stav:
        With Cells(radek, 24)
        If .Value = "ano" Then
            Me.chkHodn3.Value = True
        Else
            Me.chkHodn3.Value = False
        End If
        'koment��:
        Me.txtHodn3.Value = .Comment.Text
        End With

        'stav:
        With Cells(radek, 25)
        If .Value = "ano" Then
            Me.chkHodn4.Value = True
        Else
            Me.chkHodn4.Value = False
        End If
        'koment��:
        Me.txtHodn4.Value = .Comment.Text
        End With

        'stav:
        With Cells(radek, 26)
        If .Value = "ano" Then
            Me.chkHodn5.Value = True
        Else
            Me.chkHodn5.Value = False
        End If
        'koment��:
        Me.txtHodn5.Value = .Comment.Text
        End With

        'stav:
        With Cells(radek, 27)
        If .Value = "ano" Then
            Me.chkHodn6.Value = True
        Else
            Me.chkHodn6.Value = False
        End If
        'koment��:
        Me.txtHodn6.Value = .Comment.Text
        End With

        With Me

        'na��st pozn�mku k uz�v�rce:
        If reklama = False Then
            .txtPoznamka.Value = Cells(radek, 35).Comment.Text
        End If

        'na��st pozn�mku k reklam�:
        If reklama = True Then
            .txtReklPozn.Value = Cells(radek, 14).Comment.Text
        End If

        'na��st stav SCHV�LENO:
        If Cells(radek, 14).Value = "ano" Then
            If reklama = False Then
                .optSchvalit.Value = True
            Else
                .optSchvalitRekl.Value = True
            End If
            If Cells(radek, 35).Value <> "ano" Then _
                MsgBox "Tento z�znam byl ji� SCHV�LEN�...      ", vbInformation, aplikace
        ElseIf Cells(radek, 14).Value = "ne" Then
            If reklama = False Then
                .optZamitnout.Value = True
            Else
                .optZamitnoutRekl.Value = True
            End If
            If Cells(radek, 35).Value <> "ano" Then _
                MsgBox "Tento z�znam byl ji� ZAM�TNUT�...      ", vbInformation, aplikace
        End If

        'na��st stav SCHV�LIT ADMINISTR�TOREM - pouze pro kampan�
        'nebo stav vyu�it� - pouze pro reklamy:
        If reklama = False Then
            If admin = True Then
                If Cells(radek, 34).Value = "ano" And Cells(radek, 35).Value <> "ano" Then
                    MsgBox "Tento z�znam byl ji� SCHV�LEN� ADMINISTR�TOREM...      ", vbExclamation, aplikace
                    .chkKamAdmin.Value = True
                Else
                    .chkKamAdmin.Value = False
                End If
            End If
        Else
            vyuzito = Cells(radek, 34).Value
        End If
        .lblUser.Caption = Cells(radek, 30).Value
'        .cmdPrevzitReklamu.Visible = False

        cisKampane = .txtKamCislo   '�. kampan� - pro editaci+nov� p�evzet� reklamy!!

        'info o zdrojov� reklam�:
        If reklama = False Then
            If Cells(radek, 36).Value <> "" Then    'existuje ��slo zdroj. reklamy
                With .lblZdrojovaReklama
                .Caption = "P�evzato z reklamy " & Cells(radek, 36).Value
                .Visible = True
                End With
            End If
        End If
        End With    'Me
    End If
    Me.lblZaznamStatus.Caption = UCase(typZaznamu)

    '----------------------------------
    'konec nast. FRM podle typu z�znamu

    With Me
'    'p�ehled �erp�n� - pro EDIT se resetuje ode�ten� aktu�ln� kampan� a� do zm�ny ��stky:
'    '----------------
'    If typZaznamu = "edit" Then
'        .lblZustP1.Caption = Range("N2").Value
'        .lblZustP2.Caption = Range("N2").Value
'        .lblZustS1.Caption = Range("O2").Value
'        .lblZustS2.Caption = Range("O2").Value
'    End If

    'd�lka pozn�mky k uz�v�rce:
    .lblPoznDelka.Caption = Len(.txtPoznamka.Value)

    'n�pov�dn� texty:
    '----------------
    .lblHelp1.Visible = zobrHelp
    .lblHelp2.Visible = zobrHelp
    .lblHelp3.Visible = zobrHelp
    End With    'Me

    Me.txtHodn3.Visible = False 'p��znak ��dost zasl�na - bez textu   od v. 3.0

    'Edit - z�znam po uz�v�rce jen pro n�hled:      od v. 4.3
    '-----------------------------------------
    If typZaznamu = "edit" Then
        '��dek je po uz�v�rce:
        If Cells(ActiveCell.Row, 35).Value = "ano" Then
            'vypnout v�echny controls:
            Dim control As control
            For Each control In frm_ZAZNAM.Controls
                control.Enabled = False
            Next control
            'nastavit vybran� controls:
            With Me
            .fraCerpani.Visible = False
            .cmdStorno.Enabled = True
            .lblZaznamStatus.Caption = "N�HLED"
            .lblZaznamStatus.Enabled = True
            End With
            MsgBox "Kampa� " & Cells(ActiveCell.Row, 2) & " je po uz�v�rce!      " & vbNewLine _
                & "Pouze administr�tor m��e zru�it uz�v�rku m�s�ce a z�znam editovat.      ", vbExclamation, aplikace
            Exit Sub
        End If
    End If

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'zav�en� frm k��kem
    Call cmdStorno_Click
End Sub


Private Sub UserForm_Terminate()
    'ukon�en� frm
    ciselnikZaznam = False
End Sub



'================= NA�TEN� ��SELN�K� ==================

Private Sub NacistCiselnikKamNazev()

    With Me.cmbKamNazev
    .Clear
    cisRadek = ""
    Open cestaMroManager & "\Common_Files\Codelist\k_nazev.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, cisRadek             '�ten� jednoho ��dku a najet� na dal�� a� do EOF=True!!
        .AddItem cisRadek
    Loop
    Close #1
    End With

End Sub


Private Sub NacistCiselnikKamTyp()

    With Me.cmbKamTyp
    .Clear
    cisRadek = ""
    Open cestaMroManager & "\Common_Files\Codelist\k_typ.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, cisRadek
        .AddItem cisRadek
    Loop
    Close #1
    End With

End Sub


Private Sub NacistCiselnikKamZamereni()

    With Me.cmbKamZamereni
    .Clear
    cisRadek = ""
    Open cestaMroManager & "\Common_Files\Codelist\k_zamereni.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, cisRadek
        .AddItem cisRadek
    Loop
    Close #1
    End With

End Sub


Private Sub NacistCiselnikKamMedium()

    With Me.cmbKamMedium
    .Clear
    cisRadek = ""
    Open cestaMroManager & "\Common_Files\Codelist\m_typ.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, cisRadek
        .AddItem cisRadek
    Loop
    Close #1
    End With

End Sub


Private Sub NacistCiselnikKamMediumNazev()
    'tridilny ciselnik!

    With Me.cmbKamMediumNazev
    .Clear

    cisRadek = ""
    Open cestaMroManager & "\Common_Files\Codelist\m_nazev1.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, cisRadek
        .AddItem cisRadek
    Loop
    Close #1

    cisRadek = ""
    Open cestaMroManager & "\Common_Files\Codelist\m_nazev2.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, cisRadek
        .AddItem cisRadek
    Loop
    Close #1

    cisRadek = ""
    Open cestaMroManager & "\Common_Files\Codelist\m_nazev3.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, cisRadek
        .AddItem cisRadek
    Loop
    Close #1

    End With

End Sub


Private Sub NacistCiselnikKamFormat()

    With Me.cmbKamFormat
    .Clear
    cisRadek = ""
    Open cestaMroManager & "\Common_Files\Codelist\k_format.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, cisRadek
        .AddItem cisRadek
    Loop
    Close #1
    End With

End Sub
'================= NA��T�N� ��SELN�K� - KONEC ==================



Private Sub NajitPenetraci()
    'pouze �kodaPlus
    'podle NW naj�t penetraci a p�i�adit % p��sp�vku
    'zobrazit penetraci ve frm
    'vol�no z frm_init

    With Me
    For i = LBound(dlrNw) To UBound(dlrNw)
        If dlrNw(i) = Range("A1").Value Then
            .txtKamProcento.Value = dlrSpPrispevek(i)
            With .lblPenetrace
            .Visible = True
            .Caption = "Penetrace �koFin " & dlrSpPenetrace(i) & "%"
            End With
            Exit Sub
        End If
    Next i
    End With    'Me

End Sub



Private Sub ZaznamRezimReklama()
    'p�epnout do re�imu schvalov�n� reklam

Dim barva       As Long
Dim posuv       As Integer      'posuv objekt� nahoru v re�imu reklamy

    barva = RGB(236, 130, 4)    'RF
    posuv = 72

    '�prava frm:
    With Me
    .StartUpPosition = 0
    .Width = 343
    .Height = 464
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)

    'z�hlav� r�me�ku KAMPA�:
    With .lblTitul1
    .ForeColor = barva
    .Width = 96
    .Caption = "Schv�len� reklamy:"
    End With
    With .lblIntCisloKam
    .ForeColor = barva
    .Caption = "Intern� ��slo reklamy:"
'        .Left = 150
    End With
    With .txtKamCislo
    .ForeColor = barva
    .Left = 268
    End With

    'r�me�ek Kampa�:
    .lblPrecislovatFakturu.Visible = False

    .lblDatum.Caption = "Datum ��dosti:"

    .lblRegCislo.Enabled = False
    .txtRegCislo.Enabled = False

    .lblDlrZadost.Enabled = False
    .txtDlrZadost.Enabled = False

    .lblSubFak.Enabled = False
    .txtSubFak.Enabled = False

    .chkKombinace.Visible = False
    .lblKombinace.Visible = False

    .lblKamNazev.Caption = "N�zev reklamy:"

    .fraCastky.Visible = False
    .txtKamProcento.Value = ""

    .lblKamTyp.Top = .lblKamTyp.Top - posuv
    .cmbKamTyp.Top = .cmbKamTyp.Top - posuv

    .lblKamZamereni.Top = .lblKamZamereni.Top - posuv
    .cmbKamZamereni.Top = .cmbKamZamereni.Top - posuv

    .lblKamMedium.Top = .lblKamMedium.Top - posuv
    .cmbKamMedium.Top = .cmbKamMedium.Top - posuv

    .lblKamMediumNazev.Top = .lblKamMediumNazev.Top - posuv
    .cmbKamMediumNazev.Top = .cmbKamMediumNazev.Top - posuv

    .lblKamZdroj.Top = .lblKamZdroj.Top - posuv
    .cmbKamZdroj.Top = .cmbKamZdroj.Top - posuv
'    .chkDas.Top = .chkDas.Top - posuv

    .lblKamFormat.Top = .lblKamFormat.Top - posuv
    .cmbKamFormat.Top = .cmbKamFormat.Top - posuv

    .lblKamPocet.Top = .lblKamPocet.Top - posuv
    .txtKamPocet.Top = .txtKamPocet.Top - posuv
    .lblKamPocetReklamDok.Top = .lblKamPocetReklamDok.Top - posuv
    .lblCiselnik2.Top = .lblCiselnik2.Top - posuv
    .lblCiselnik3.Top = .lblCiselnik3.Top - posuv
    .lblCiselnik4.Top = .lblCiselnik4.Top - posuv
    .lblCiselnik5.Top = .lblCiselnik5.Top - posuv
    .lblCiselnik6.Top = .lblCiselnik6.Top - posuv

    'objekty jen pro reklamu:
    .lblReklPozn.Top = 264
    .lblReklDelka.Top = 264
    .txtReklPozn.Top = 276
    .lblKamPocetReklamDok.Visible = True    'n�pov�da k po�tu mat.
'    With Me.chkSchvalitReklamu
'        .Top = 456
'        .Left = 12
'    End With
    With .optSchvalitRekl
    .Visible = True
    .Value = False
    End With
    With .optZamitnoutRekl
    .Visible = True
    .Value = False
    End With

    'stavov� hl�ky:
    With .lblZaznamStatus
    .BorderColor = barva
    .ForeColor = barva
    End With
    .lblAlertKombinace.BackColor = barva
    .lblAlertSpolecna.BackColor = barva
    .lblAlertBonus.BackColor = barva

    'tla��tka:
    With .cmdUlozit
        .Top = 406
        .Left = 266
'        .Enabled = False        'VYPNUTO, NE� SE DOPROGRAMUJE Z�PIS!!!!
    End With
    With .cmdStorno
        .Top = 406
        .Left = 196
    End With

    End With    'Me

End Sub



Private Sub NacistRadek()
    'volan� sub - na��st ��dek z tabulky do FRM
    'na�tou se spole�n� polo�ky pro KOPII a EDIT z�znamu
    'vol�no z FRM INIT

    appEvents = False

    With Me
    If Cells(radek, 13).Value = "ano" Then          'bonus akce ano/ne
        .chkBonus.Value = True
        If Cells(radek, 13).Comment.Text = "R" Then
            typBonus = "R"
        ElseIf Cells(radek, 13).Comment.Text = "I" Then
            typBonus = "I"
        Else
            typBonus = ""
        End If
    Else
        .chkBonus.Value = False
        typBonus = ""
    End If
    .lblTypBonus.Caption = typBonus

    If Cells(radek, 32).Value = "ano" Then          'spole�n� reklama ano/ne
        .chkSpolReklama.Value = True
    Else
        .chkSpolReklama.Value = False
    End If

    If Cells(radek, 31).Value = "ano" Then          'kam. vlastn�/import
        .optKamVlastni.Value = True
    Else
        .optKamImport.Value = True
    End If
    End With    'Me

    With Cells(radek, 8)
        If .Value = "prodej" Then                   'oblast p��sp�vku
            Me.optProdej.Value = True
        ElseIf .Value = "servis" Then
            Me.optServis.Value = True
        Else
            MsgBox "Nen� definov�na oblast p��sp�vku prodej/servis...      ", vbExclamation, aplikace
        End If
    End With

    With Me
    .cmbKamNazev.Value = Cells(radek, 7).Value       'n�zev kam.

'    .txtKamCena.Value = Cells(radek, 9).Value        'cena celkem
    .txtKamCena.Value = Application.WorksheetFunction.Substitute(Cells(radek, 9).Value, ".", ",")       'cena celkem (konverze te�ky na ��rku)
    .txtKamProcento.Value = Cells(radek, 10).Value    '% p��sp�vku
    .txtKamSchval.Value = Cells(radek, 11).Value      'schv�len� ��stka
    .txtKamNeschval.Value = Application.WorksheetFunction.Substitute(Cells(radek, 12).Value, ".", ",")   'neschv�len� ��stka

    .cmbKamTyp.Value = Cells(radek, 15).Value        'typ kam.
    .cmbKamZamereni.Value = Cells(radek, 16).Value   'zam��en� kam.
    .cmbKamMedium.Value = Cells(radek, 17).Value     'medium-typ
    .cmbKamMediumNazev.Value = Cells(radek, 18).Value     'medium-n�zev
    .cmbKamZdroj.Value = Cells(radek, 19).Value      'zdroj dat
'    If Cells(radek, 19).Value = "PP" Then                 'PP-DAS
'        If Cells(radek, 19).Comment.Text = "ano" Then
'            .chkDas.Value = True
'        Else
'            .chkDas.Value = False
'        End If
'    End If

    .cmbKamFormat.Value = Cells(radek, 20).Value     'form�t
    .txtKamPocet.Value = Cells(radek, 21).Value      'po�et ks mat.
    End With    'Me

    appEvents = True

End Sub



Private Sub GenerovatCisloKampane()
    'volan� sub - generuje ��slo kampan� pro Nov� z�znam a Kopii z�znamu
    'vol�no z FRM INIT

    poslRadek = Cells(Rows.Count, 2).End(xlUp).Row              'podle sloupce �. kampan� (datum m��e b�t pr�zdn�)
    If poslRadek = 4 And Cells(poslRadek, 1).Value = "" Then    'tabulka je zat�m pr�zdn�
        If reklama = False Then
            cisKampane = Range("A1").Value & "-" & "001"            'prvn� ��slo kampan� pro ��dosti
        Else
            cisKampane = Range("A1").Value & "-" & "9001"           'prvn� ��slo kampan� pro reklamy
        End If
    Else
        'generovat nov� po�. ��slo kampan�:
        If reklama = False Then     'pro ��dosti
        
            cisKampane = CInt(Right(Cells(poslRadek, 2).Value, 3))   'koncov� ��slo p�edchoz� kampan�
            cisKampane = CStr(cisKampane + 1)
            'doplnit na 3 m�stn� ��slo:
            If Len(cisKampane) = 1 Then
                cisKampane = "00" & cisKampane
            ElseIf Len(cisKampane) = 2 Then
                cisKampane = "0" & cisKampane
            End If
            
        Else                        'pro reklamy
        
            cisKampane = CInt(Right(Cells(poslRadek, 2).Value, 4))   'koncov� ��slo p�edchoz� kampan�
            cisKampane = CStr(cisKampane + 1)
            
        End If
        cisKampane = Range("A1").Value & "-" & cisKampane   'sestavit ��slo kampan� DLR
        
        'korekce ��sla rekl. - z ext. instalace jsou ��slovan� xxxxx-8xxx:
        If Mid(cisKampane, 7, 1) = "8" Then _
            cisKampane = Application.WorksheetFunction.Substitute(cisKampane, "-8", "-7")
    End If

    Me.txtKamCislo.Value = cisKampane

End Sub



Private Sub lblDnes_Click()
    'do pole Datum generovat aktu�ln� datum (a sko�it na dal�� prvek)

    Call GenerovatDatum
    With Me
    If reklama = False Then
        .txtRegCislo.SetFocus
    Else
        .optProdej.SetFocus
    End If
    End With

End Sub



Private Sub GenerovatDatum()
    'volan� sub - generovat datum kampan� do FRM

    'zat�m aktu�ln� datum
    Me.txtDatum.Value = Day(Date) & "." & Month(Date) & "." & Year(Date)

End Sub



Private Sub txtDatum_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'odchod z pole datum = generovat ��slo ��dosti
    Call GenerovatCisloZadosti

End Sub



Private Sub GenerovatCisloZadosti()
    'generov�n� �. ��dosti z pole Datum
    'vol�no z txtDatum_Exit a cmdUlozit_Click

    If reklama = False Then
        On Error GoTo Konec
        With Me
        mesicSoub = Month(.txtDatum.Value)
        If Len(mesicSoub) = 1 Then
            mesicSoub = "0" & mesicSoub
        End If
        If skodaPlus = False Then
            .txtDlrZadost.Value = Right(CStr(rok), 2) & mesicSoub & "-" & Range("A1").Value
        Else
            .txtDlrZadost.Value = "9" & Right(CStr(rok), 2) & mesicSoub & "-" & Range("A1").Value
        End If
        End With
    End If
Konec:
End Sub



Private Sub txtSubFak_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'odchod z pole = kontrola ��sla subdod. fa na ji� existuj�c�

Dim faktura As String
Dim cena    As String
Dim prisp   As String
Dim pozn1   As String
Dim pozn2   As String

    If appEvents = True Then    'm��e b�t False p�i ukl�d�n� z�znamu - o�et�ena chyba s kopi� z�znamu (ZATIM VYPNUTO)
    
        faktura = Me.txtSubFak.Value
        poslRadek = Cells(Rows.Count, 2).End(xlUp).Row
        If Not (poslRadek = 4 And Cells(poslRadek, 2).Value = "") Then    'tabulka nen� pr�zdn�
            For i = 4 To poslRadek
                If faktura = Cells(i, 6).Value Then     'nalezeno duplicitn� ��slo fa
                
                    cena = Application.WorksheetFunction.Fixed(Cells(i, 9).Value, 2)
                    cena = Application.WorksheetFunction.Substitute(cena, ",", " ")
                    cena = Application.WorksheetFunction.Substitute(cena, ".", ",")
                    
                    prisp = Application.WorksheetFunction.Fixed(Cells(i, 11).Value, 2)
                    prisp = Application.WorksheetFunction.Substitute(prisp, ",", " ")
                    prisp = Application.WorksheetFunction.Substitute(prisp, ".", ",")
                    
                    On Error Resume Next    'ignorovat chyby �ten� pr�zdn�ch koment���
                    pozn1 = Cells(i, 22).Comment.Text
                    pozn2 = Cells(i, 23).Comment.Text
    
                    MsgBox "Byla nalezena kampa� s ��slem subdod. fa " & faktura & ":" & vbNewLine & vbNewLine _
                        & "�. kampan�:  " & Cells(i, 2).Value & vbNewLine _
                        & "Datum:           " & Cells(i, 1).Value & vbNewLine _
                        & "Kampa�:        " & Cells(i, 7).Value & vbNewLine _
                        & "Cena:              " & cena & vbNewLine _
                        & "P��sp�vek:      " & prisp & vbNewLine _
                        & "Schv�leno:     " & UCase(Cells(i, 14).Value) & vbNewLine _
                        & vbNewLine _
                        & "Pozn. 1:          " & pozn1 & vbNewLine _
                        & "Pozn. 2:          " & pozn2 & vbNewLine _
                        , vbExclamation, "Duplicitn� ��slo fa dodavatele!"
                End If
            Next i
        End If
    
    End If  'appEvents

End Sub



Private Sub chkBonus_Change()
    'zm�na chk Bonus akce

    With Me
    .lblAlertBonus.Visible = .chkBonus.Value
    If reklama = False Then
        .cmdPrevzitBonus.Visible = .chkBonus.Value
        .lblTypBonus.Visible = .chkBonus.Value
    End If
    If .chkBonus.Value = True Then
        .txtKamProcento.Value = prispevek2
    Else
        If skodaPlus = False Then
            .txtKamProcento.Value = prispevek
        Else
            Call NajitPenetraci     'podle NW naj�t penetraci a p�i�adit % p��sp�vku
        End If
        'reset zdroj. bonus:
        idBonusy = -100
        .lblZdrojovyBonus.Visible = False
        .lblZdrojovyBonus.Caption = "P�evzato: "
    End If
    End With
    Call VypocetCeny

End Sub



Private Sub cmdPrevzitBonus_Click()
    'frm p�evz�t bonus

    frmBonusy.Show

End Sub



Private Sub chkSpolReklama_Change()
    'zm�na chk Spole�n� reklama

    Me.lblAlertSpolecna.Visible = Me.chkSpolReklama.Value

End Sub



Private Sub optKamImport_Change()
    'option Kampa� import�ra

'    If optKamImport = True Then
'
'        'nastavit seznam na kampan� import�ra:
'        With cmbKamNazev
'            .Clear
'            For i = LBound(cisKamImport) To UBound(cisKamImport)
'                .AddItem cisKamImport(i)
'            Next i
'            .ShowDropButtonWhen = fmShowDropButtonWhenAlways    'kam. import�ra nab�z� rozbalen� seznamu
'        End With
'
'    Else
'
'        'nastavit seznam na kampan� vlastn�:
'        With cmbKamNazev
'            .Clear
'            .ShowDropButtonWhen = fmShowDropButtonWhenNever     'vlastn� kam. - pouze zapsat
'        End With
'
'    End If

End Sub



Private Sub optKamVlastni_Change()
    Call optKamImport_Change
    '(optiony se p�ep�naj�)
End Sub



Private Sub optProdej_Change()
    'zm�na prodej/servis
    Call VypocetCeny    'pro p�epo�et z�statk�
End Sub



Private Sub optServis_Change()
    'zm�na prodej/servis
    Call VypocetCeny    'pro p�epo�et z�statk�
End Sub



Private Sub chkKombinace_Change()

    With Me
    .lblAlertKombinace.Visible = .chkKombinace.Value

'PRO Z�ZNAM KAMPAN�:
'-------------------
'    If .chkKombinace.Value = True Then
    If .chkKombinace.Value = True And typZaznamu = "kopie" Then
        'zapnout kopii z jedn� kampan� - zkop�rovat vybran� hodnoty z aktu�ln�ho ��dku:
        If radek = poslRadek Then .txtKamCislo.Value = Cells(radek, 2).Value     'int. ��slo kam. - pouze pro kopii z posledn�ho ��dku!!
        .txtDatum.Value = Cells(radek, 1).Value       'datum
        .txtDatum.Value = Day(Cells(radek, 1).Value) & "." & Month(Cells(radek, 1).Value) & "." & Year(Cells(radek, 1).Value)

        .txtDlrZadost.Value = Cells(radek, 4).Value      '�. ��dosti
        .txtRegCislo.Value = Cells(radek, 3).Value    'reg. ��slo RF
        .txtSubFak.Value = Cells(radek, 6).Value          '�. subdod. fa

        If Cells(radek, 8).Value = "prodej" Then             'prvn� ��dek je prodej (zapnout 2. mo�nost)
            .optServis.Value = True
        ElseIf Cells(radek, 8).Value = "servis" Then
            .optProdej.Value = True
        End If

        With .txtKamCena
            .Value = 0
            .SetFocus
        End With

        'smazat zadan� polo�ky - za P/S se li��:
        .cmbKamTyp.Value = ""
        .cmbKamZamereni.Value = ""

    End If
    End With    'Me

End Sub



Private Sub cmbKamNazev_Change()
    'n�zev kampan� m��e generovat typ/zam��en� kampan�

Dim kamNazev As String

    'od v. 5.2.1 - p�i zm�n� n�zvu kampan� z�stavaly p�vodn� parametry z�znamu
    If typZaznamu = "edit" And appEvents = True Then
        With Me
            .cmbKamTyp.Value = ""
            .cmbKamZamereni.Value = ""
            .cmbKamMedium.Value = ""
            .cmbKamMediumNazev.Value = ""
            .cmbKamZdroj.Value = ""
            .cmbKamFormat.Value = ""
            .optProdej.Value = False
            .optServis.Value = False
        End With
    End If

'    If typZaznamu = "novy" Then
        With Me
        kamNazev = .cmbKamNazev.Value
    
        If InStr(kamNazev, "Citigo") > 0 Then
            .cmbKamTyp.Value = "Nov� vozy"
            .cmbKamZamereni.Value = "Citigo"

        ElseIf InStr(kamNazev, "Testovac� j�zdy Fabia 2015") > 0 Then
            .cmbKamTyp.Value = "Nov� vozy"
            .cmbKamZamereni.Value = "Fabia"

        ElseIf InStr(kamNazev, "Operativn� leasing - Bez starost�") > 0 Then
                                '----------------------
            .cmbKamTyp.Value = "Nov� vozy"
            .cmbKamZamereni.Value = "V�ce model� voz�"

        ElseIf InStr(kamNazev, "Centr�ln� region�ln� kampa� 2015") > 0 Or _
            InStr(kamNazev, "Centr�ln� kampa� - Hokejov� extraliga 2014/15") > 0 Then
                                '----------------------
            If typZaznamu <> "edit" Then .chkBonus.Value = True
            .cmbKamTyp.Value = "Nov� vozy"
            .cmbKamZamereni.Value = "V�ce model� voz�"
            .cmbKamMedium.Value = "Akce - prezentace"
            .cmbKamMediumNazev.Value = "Centr�ln� kampa� �A"
            .cmbKamZdroj.Value = "B2B"
            .cmbKamFormat.Value = "Vlastn�"

        ElseIf InStr(kamNazev, "Centr�ln� kampa� - Car Configurator (PCC)") > 0 Then
                                '----------------------
            If typZaznamu <> "edit" Then .chkBonus.Value = True
            .cmbKamTyp.Value = "Nov� vozy"
            .cmbKamZamereni.Value = "V�ce model� voz�"
            .cmbKamMedium.Value = "TV LED"
            .cmbKamMediumNazev.Value = "Centr�ln� kampa� �A"
            .cmbKamZdroj.Value = "B2B"
            .cmbKamFormat.Value = "Data"

        ElseIf InStr(kamNazev, "�KODA Poji�t�n�") > 0 Then
                                '----------------------
            .cmbKamTyp.Value = "Prezentace spole�nosti"
            .cmbKamZamereni.Value = "Prezentace spole�nosti"
            .cmbKamMedium.Value = "Tisk"

        ElseIf InStr(kamNazev, "Fabia NOV� Combi") > 0 Then
            .cmbKamTyp.Value = "Nov� vozy"
            .cmbKamZamereni.Value = "Fabia Combi"

        ElseIf InStr(kamNazev, "Fabia") > 0 Then
            .cmbKamTyp.Value = "Nov� vozy"
            .cmbKamZamereni.Value = "Fabia"

        ElseIf InStr(kamNazev, "Octavia Combi") > 0 Then
            .cmbKamTyp.Value = "Nov� vozy"
            .cmbKamZamereni.Value = "Octavia Combi"

        ElseIf InStr(kamNazev, "Octavia") > 0 Then
            .cmbKamTyp.Value = "Nov� vozy"
            .cmbKamZamereni.Value = "Octavia"

        ElseIf InStr(kamNazev, "Roomster") > 0 Then
            .cmbKamTyp.Value = "Nov� vozy"
            .cmbKamZamereni.Value = "Roomster"

        ElseIf InStr(kamNazev, "Superb Combi") > 0 Then
            .cmbKamTyp.Value = "Nov� vozy"
            .cmbKamZamereni.Value = "Superb Combi"

        ElseIf InStr(kamNazev, "Superb") > 0 Then
            .cmbKamTyp.Value = "Nov� vozy"
            .cmbKamZamereni.Value = "Superb"

        ElseIf InStr(kamNazev, "Rapid Spaceback") > 0 Then
            .cmbKamTyp.Value = "Nov� vozy"
            .cmbKamZamereni.Value = "Rapid Spaceback"

        ElseIf InStr(kamNazev, "Rapid") > 0 Then
            .cmbKamTyp.Value = "Nov� vozy"
            .cmbKamZamereni.Value = "Rapid"

        ElseIf InStr(kamNazev, "Yeti") > 0 Then
            .cmbKamTyp.Value = "Nov� vozy"
            .cmbKamZamereni.Value = "Yeti"

        ElseIf InStr(kamNazev, "servisn� akce") > 0 Then
            .cmbKamTyp.Value = "Servis"
            .cmbKamZamereni.Value = "Sez�nn� servisn� akce"
            .optServis.Value = True

        ElseIf InStr(UCase(kamNazev), "SERVIS") > 0 Then
            .optServis.Value = True
            .cmbKamTyp.Value = "Servis"
            .cmbKamZamereni.Value = "Servis"
            .optServis.Value = True

'        Else
'            .cmbKamZamereni.Value = ""
'            .cmbKamTyp.Value = ""
        End If
        End With
'    End If

End Sub



Private Sub cmbKamTyp_Change()
    'oblast �erp�n� podle typu kampan�

    With Me
    If .cmbKamTyp.Value = "Nov� vozy" Or .cmbKamTyp.Value = "�KODA Plus" Then
        .optProdej.Value = True
    End If
    End With

End Sub



Private Sub cmbKamMediumNazev_Change()
    'pro m�dia "distribuce" nastavit zdroj dat "Vlastn�"

    With Me
    If InStr(.cmbKamMediumNazev.Value, "distribuce") > 0 Then
        .cmbKamZdroj.Value = "Vlastn�"
    End If
    End With

End Sub



Private Sub txtKamCena_Enter()
    'vstup na cenu kampan� - kontrola oblasti P/S (nutn� pro p�epo�et z�statk� ve FRM)

    With Me
    If .optProdej.Value = False And .optServis.Value = False Then
        MsgBox "Nejd��ve vyberte oblast �erp�n� (prodej/servis).      ", vbExclamation, aplikace
        .optProdej.SetFocus
    End If
    End With    'Me

End Sub



Private Sub lblKopieCisel_Click()
    'pouze pro EDIT - kopie reg. ��sla a �. ��dosti z minul�ho ��dku

    With Me
    .txtRegCislo.Value = Cells(radek - 1, 3).Value
    .txtDlrZadost.Value = Cells(radek - 1, 4).Value
    End With    'Me

End Sub



Private Sub lblPrecislovatFakturu_Click()
    'POUZE PRO KAMPAN�
    'zm�na ��sla faktury pro v�echny ��dky ��dosti

Dim stareCislo  As String
Dim noveCislo   As String

    stareCislo = Me.txtDlrZadost
    noveCislo = InputBox("Zam�nit ��slo fa " & stareCislo & " za: ", _
        "Zm�na ��slo fa v ��dosti �. " & stareCislo, stareCislo)
    If noveCislo = "" Then Exit Sub
    poslRadek = Cells(Rows.Count, 2).End(xlUp).Row              'podle sloupce �. kampan� (datum m��e b�t pr�zdn�)

    'proj�t v�echny ��dky tabulky:
    '-----------------------------
    ActiveSheet.Unprotect Password:=heslo
    For i = 4 To poslRadek
        Application.StatusBar = "Kontroluji z�znam " & i - 3 & "/" & poslRadek - 3 & "..."
        If Cells(i, 4).Value = stareCislo Then Cells(i, 4).Value = noveCislo
    Next i
    Application.StatusBar = False
    ActiveSheet.Protect Password:=heslo
    ActiveWorkbook.Save
    Me.txtDlrZadost.Value = noveCislo
    MsgBox "��slo faktury DLR bylo zm�n�no.      ", vbInformation, ""

End Sub



'Private Sub cmbKamZdroj_Change()
'    'v�b�r PP povol� volbu DAS
'
'    With Me
'    If .cmbKamZdroj.Value = "PP" Then
'        .chkDas.Visible = True
'        .chkDas.Value = True
'    Else
'        .chkDas.Visible = False
'        .chkDas.Value = False
'    End If
'    End With
'
'End Sub



'========== V�PO�ET P��SP�VK�: ==========

Private Sub txtKamCena_Change()
    'zm�na celkov� ceny
    Call VypocetCeny
End Sub


Private Sub txtKamProcento_Change()
    'zm�na procenta p��sp�vku
    Call VypocetCeny
End Sub


Private Sub txtKamNeschval_Change()
    'neschv�len� ��stka
    Call VypocetCeny
End Sub



Private Sub VypocetCeny()
    'volan� sub - v�po�et schv�len� ��stky p�i zm�n� ceny nebo procenta
    'p�id�n p�epo�et z�statk� ve FRM
Dim prispevek   As Single

    With Me
    On Error Resume Next

    'v�po�et p��sp�vku a zaokrouhlen�:
    prispevek = Round((.txtKamCena.Value - .txtKamNeschval.Value) * .txtKamProcento.Value / 100, 2)
    .txtKamSchval.Value = prispevek
    'konverze te�ky na ��rku - pouze pro zobrazen�:
    .txtKamSchval.Value = Application.WorksheetFunction.Substitute(.txtKamSchval.Value, ".", ",")


'        'n�vrh do�erp�n� rozpo�tu:
'        MsgBox "Rozpo�et je p�e�erp�n o " & -1 * precerpani & " K�," & vbNewLine _
'            & "zapisuji neschv�lenou ��stku " & -1 * neschvaleno & " K�.      ", _
'            vbExclamation, "P�e�erp�n� rozpo�tu"
'        .txtKamNeschval.Value = -1 * neschvaleno



    'kontrola �erp�n�/z�statku:
    '==========================
    .lblZustP2.Caption = ""     'reset
    .lblZustS2.Caption = ""

    If .chkBonus.Value = False Then

        .lblRozpP.Visible = True
        .lblRozpS.Visible = True
        .lblZustP1.Visible = True
        .lblZustS1.Visible = True
        .lblZustP2.Visible = True
        .lblZustS2.Visible = True

        If typZaznamu = "edit" Then
            .lblZustP1.Visible = False
            .lblZustS1.Visible = False
            .lblKc2.Visible = False
            .lblKc5.Visible = False
        End If

        'v�po�et z�statku p�ed kampan�:
        '------------------------------
        If dlrCinnost(dlr) = "prodej" Then
            If typZaznamu <> "edit" Then   'pro nov� z�znamy prodej
                .lblZustP1.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetProdej(dlr) - cerpaniP, 2)
                .lblZustS1.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetServis(dlr) - cerpaniS, 2)
            Else    'p�i editaci p�i��st aktu�ln� ��dek za prodej/servis!
                If .optProdej.Value = True Then
                    .lblZustP1.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetProdej(dlr) - cerpaniP + Cells(ActiveCell.Row, 11).Value, 2)
                    .lblZustS1.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetServis(dlr) - cerpaniS, 2)
                ElseIf .optServis.Value = True Then
                    .lblZustP1.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetProdej(dlr) - cerpaniP, 2)
                    .lblZustS1.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetServis(dlr) - cerpaniS + Cells(ActiveCell.Row, 11).Value, 2)
                End If
            End If
            .lblZustP1.Caption = Application.WorksheetFunction.Substitute(.lblZustP1.Caption, ",", " ")
            .lblZustP1.Caption = Application.WorksheetFunction.Substitute(.lblZustP1.Caption, ".", ",")
    
            .lblZustS1.Caption = Application.WorksheetFunction.Substitute(.lblZustS1.Caption, ",", " ")
            .lblZustS1.Caption = Application.WorksheetFunction.Substitute(.lblZustS1.Caption, ".", ",")
    
        Else
            If typZaznamu <> "edit" Then   'pro nov� z�znamy prodej
                .lblZustP1.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetFinal(dlr) - cerpaniP - cerpaniS, 2)
            Else
                .lblZustP1.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetFinal(dlr) - cerpaniP - cerpaniS + Cells(ActiveCell.Row, 11).Value, 2)
            End If
            .lblZustP1.Caption = Application.WorksheetFunction.Substitute(.lblZustP1.Caption, ",", " ")
            .lblZustP1.Caption = Application.WorksheetFunction.Substitute(.lblZustP1.Caption, ".", ",")
    
            .lblZustS1.Caption = ""
    
        End If
    
        'v�po�et z�statku po kampani:
        '----------------------------
        If dlrCinnost(dlr) = "prodej" Then
        
            If .optProdej.Value = True Then
            
                If typZaznamu <> "edit" Then
                
                    .lblZustP2.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetProdej(dlr) - cerpaniP - prispevek, 2)
                    If dlrRozpocetProdej(dlr) - cerpaniP - prispevek < 0 Then
                        Beep
                        .lblAlert.Visible = True
'                        .cmdUlozit.Enabled = False
                    Else
                        .lblAlert.Visible = False
                        .cmdUlozit.Enabled = True
                    End If
                    
                Else    'p�i editaci p�i��st aktu�ln� ��dek!
                
                    .lblZustP2.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetProdej(dlr) - cerpaniP - prispevek + Cells(ActiveCell.Row, 11).Value, 2)
                    If dlrRozpocetProdej(dlr) - cerpaniP - prispevek + Cells(ActiveCell.Row, 11).Value < 0 Then
                        Beep
                        .lblAlert.Visible = True
'                        .cmdUlozit.Enabled = False
                    Else
                        .lblAlert.Visible = False
                        .cmdUlozit.Enabled = True
                    End If
                End If
                'konverze form�tu:
                .lblZustP2.Caption = Application.WorksheetFunction.Substitute(.lblZustP2.Caption, ",", " ")
                .lblZustP2.Caption = Application.WorksheetFunction.Substitute(.lblZustP2.Caption, ".", ",")
                
            ElseIf .optServis.Value = True Then
            
                If typZaznamu <> "edit" Then
                
                    .lblZustS2.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetServis(dlr) - cerpaniS - prispevek, 2)
                    If dlrRozpocetServis(dlr) - cerpaniS - prispevek < 0 Then
                        Beep
                        .lblAlert.Visible = True
'                        .cmdUlozit.Enabled = False
                    Else
                        .lblAlert.Visible = False
                        .cmdUlozit.Enabled = True
                    End If
                    
                Else    'p�i editaci p�i��st aktu�ln� ��dek!
                
                    .lblZustS2.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetServis(dlr) - cerpaniS - prispevek + Cells(ActiveCell.Row, 11).Value, 2)
                    If dlrRozpocetServis(dlr) - cerpaniS - prispevek + Cells(ActiveCell.Row, 11).Value < 0 Then
                        Beep
                        .lblAlert.Visible = True
'                        .cmdUlozit.Enabled = False
                    Else
                        .lblAlert.Visible = False
                        .cmdUlozit.Enabled = True
                    End If
                End If
                'konverze form�tu:
                .lblZustS2.Caption = Application.WorksheetFunction.Substitute(.lblZustS2.Caption, ",", " ")
                .lblZustS2.Caption = Application.WorksheetFunction.Substitute(.lblZustS2.Caption, ".", ",")
                
            Else    'nov� z�znam - stav p�ed v�b�rem obl. �erp�n�
                .lblZustP2.Caption = ""
                .lblZustS2.Caption = ""
            End If
            
        Else    '�innost servis/Plus
            
            If typZaznamu <> "edit" Then
                .lblZustP2.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetFinal(dlr) - cerpaniP - cerpaniS - prispevek, 2)
            Else
                .lblZustP2.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetFinal(dlr) - cerpaniP - cerpaniS - prispevek + Cells(ActiveCell.Row, 11).Value, 2)
            End If
            
            .lblZustP2.Caption = Application.WorksheetFunction.Substitute(.lblZustP2.Caption, ",", " ")
            .lblZustP2.Caption = Application.WorksheetFunction.Substitute(.lblZustP2.Caption, ".", ",")
            If dlrRozpocetFinal(dlr) - cerpaniP - cerpaniS - prispevek < 0 Then
                Beep
                .lblAlert.Visible = True
'                .cmdUlozit.Enabled = False
            Else
                .lblAlert.Visible = False
                .cmdUlozit.Enabled = True
            End If
    
            .lblZustS2.Caption = ""
        End If
        
    Else    'BONUS
        .lblRozpP.Visible = False
        .lblRozpS.Visible = False
        .lblZustP1.Visible = False
        .lblZustS1.Visible = False
        .lblZustP2.Visible = False
        .lblZustS2.Visible = False
        
        .lblAlert.Visible = False   '3.2.3 opravena chyba - zapnut� Bonusu nevypnulo p�e�erp�n�
        .cmdUlozit.Enabled = True

    End If

    End With    'Me

End Sub
'========== V�PO�ET P��SP�VK� - KONEC ==========



Private Sub cmdZustKontrola_Click()
    'tl. p�epo�et z�statk� po kampani

Dim precerpani  As Long
Dim neschvaleno As Long

    'refresh z�statk� v info r�me�ku a indikace p�e�erp�n�: - POUZE PRO NEBONUSOV�!!!!!
    '------------------------------------------------------
    With Me
    If .chkBonus = False Then
        .lblZustP2.Caption = .lblZustP1.Caption
        .lblZustS2.Caption = .lblZustS1.Caption

        'pro PRODEJ je kontrola rozd�lena na prodej/servis:
        If Range("G1").Value = "�innost: prodej" Then
            If .optProdej.Value = True Then
    
                precerpani = Application.WorksheetFunction.RoundUp((CSng(.lblZustP1.Caption)) - (CSng(.txtKamSchval.Value)), 0)
                neschvaleno = precerpani / .txtKamProcento.Value * 100
                .lblZustP2.Caption = precerpani
                If precerpani < 0 Then              'p�e�erp�n� prodej
                    Call ZobrazitPrecerpani
                    'n�vrh do�erp�n� rozpo�tu:
                    MsgBox "Prodej je p�e�erp�n o " & -1 * precerpani & " K�," & vbNewLine _
                        & "zapisuji neschv�lenou ��stku " & -1 * neschvaleno & " K�.      ", _
                        vbExclamation, "P�e�erp�n� rozpo�tu"
                    .txtKamNeschval.Value = -1 * neschvaleno
                Else
                    .lblPrecerpani.Visible = False
                End If
    
            ElseIf .optServis.Value = True Then
    
                precerpani = Application.WorksheetFunction.RoundUp((CSng(.lblZustS1.Caption)) - (CSng(.txtKamSchval.Value)), 0)
                neschvaleno = precerpani / .txtKamProcento.Value * 100
                .lblZustS2.Caption = precerpani
                If precerpani < 0 Then              'p�e�erp�n� servis
                    Call ZobrazitPrecerpani
                    'n�vrh do�erp�n� rozpo�tu:
                    MsgBox "Servis je p�e�erp�n o " & -1 * precerpani & " K�," & vbNewLine _
                        & "zapisuji neschv�lenou ��stku " & -1 * neschvaleno & " K�.      ", _
                        vbExclamation, "P�e�erp�n� rozpo�tu"
                    .txtKamNeschval.Value = -1 * neschvaleno
                Else
                    .lblPrecerpani.Visible = False
                End If
    
            End If
            
        'pro SERVIS / PLUS / RETAIL je pou�ito jednoduch� zobrazen� - pouze �. 1:
        Else
            precerpani = Application.WorksheetFunction.RoundUp((CSng(.lblZustP1.Caption)) - (CSng(.txtKamSchval.Value)), 0)
            neschvaleno = precerpani / .txtKamProcento.Value * 100
            .lblZustP2.Caption = precerpani
            If precerpani < 0 Then              'p�e�erp�n� MRO
                Call ZobrazitPrecerpani
                'n�vrh do�erp�n� rozpo�tu:
                MsgBox "Rozpo�et je p�e�erp�n o " & -1 * precerpani & " K�," & vbNewLine _
                    & "zapisuji neschv�lenou ��stku " & -1 * neschvaleno & " K�.      ", _
                    vbExclamation, "P�e�erp�n� rozpo�tu"
                .txtKamNeschval.Value = -1 * neschvaleno
            Else
                .lblPrecerpani.Visible = False
            End If
        End If

    Else    'pro Bonus akce se nep�epo��t� p�ehled �erp�n�:
        .lblZustP1.Caption = Range("L3").Value
        .lblZustP2.Caption = Range("L3").Value
        .lblZustS1.Caption = Range("O3").Value
        .lblZustS2.Caption = Range("O3").Value
        .lblPrecerpani.Visible = False
    End If

    If Range("D2").Value <> "�innost: prodej" Then
        .lblZustP2.Caption = Application.WorksheetFunction.Round((CSng(.lblZustP1.Caption)) - (CSng(.txtKamSchval.Value)), 1)
    End If
    End With    'Me

End Sub



Private Sub ZobrazitPrecerpani()
    'zobrazit alert p�e�erp�n�, vol�no z cmdZustKontrola_Click

    With Me.lblPrecerpani
    If Range("G1").Value <> "�innost: prodej" Then
        .Caption = "POZOR! P�e�erp�n� rozpo�tu MRO!"
    End If
    .Top = 312
'    .Left = 309
    .Visible = True    'alert
    End With

End Sub



'========== HODNOCEN�: ==========

Private Sub chkHodn1_Change()
    Call Hodnoceni
End Sub

Private Sub chkHodn2_Change()
    Call Hodnoceni
End Sub

Private Sub chkHodn3_Change()
    Call Hodnoceni
End Sub

Private Sub chkHodn4_Change()
    Call Hodnoceni
End Sub

Private Sub chkHodn5_Change()
    Call Hodnoceni
End Sub

Private Sub chkHodn6_Change()
'    Call Hodnoceni
End Sub


Private Sub Hodnoceni()

    With Me
    If .chkHodn1.Value = True And .chkHodn2.Value = True _
    And .chkHodn3.Value = True And .chkHodn4.Value = True _
    And .chkHodn5.Value = True Then
        With .optSchvalit
            .Enabled = True
'            .Value = True
        End With
    Else
        With .optSchvalit
            .Value = False
            .Enabled = False
        End With
    End If
    End With    'Me

End Sub


Private Sub lblPozn1_Click()
    'kop�rovat pozn. do pozn�mky k uz�v�rce
    With Me
    If Len(.txtHodn1.Value) > 0 Then .txtPoznamka.Value = .txtPoznamka.Value & " " & .txtHodn1.Value
    End With    'Me
End Sub

Private Sub lblPozn2_Click()
    'kop�rovat pozn. do pozn�mky k uz�v�rce
    With Me
    If Len(.txtHodn2.Value) > 0 Then .txtPoznamka.Value = .txtPoznamka.Value & " " & .txtHodn2.Value
    End With    'Me
End Sub

Private Sub lblPozn3_Click()
    'kop�rovat pozn. do pozn�mky k uz�v�rce
    With Me
    If Len(.txtHodn3.Value) > 0 Then .txtPoznamka.Value = .txtPoznamka.Value & " " & .txtHodn3.Value
    End With    'Me
End Sub

Private Sub lblPozn4_Click()
    'kop�rovat pozn. do pozn�mky k uz�v�rce
    With Me
    If Len(.txtHodn4.Value) > 0 Then .txtPoznamka.Value = .txtPoznamka.Value & " " & .txtHodn4.Value
    End With    'Me
End Sub

Private Sub lblPozn5_Click()
    'kop�rovat pozn. do pozn�mky k uz�v�rce
    With Me
    If Len(.txtHodn5.Value) > 0 Then .txtPoznamka.Value = .txtPoznamka.Value & " " & .txtHodn5.Value
    End With    'Me
End Sub

Private Sub lblPozn6_Click()
    'kop�rovat pozn. do pozn�mky k uz�v�rce
    With Me
    If Len(.txtHodn6.Value) > 0 Then .txtPoznamka.Value = .txtPoznamka.Value & " " & .txtHodn6.Value
    End With    'Me
End Sub



Private Sub txtPoznamka_Change()
    'zm�na d�lky pozn�mky k uz�v�rce

Dim limit       As Integer

    limit = 55

    With Me
    .lblPoznDelka.Caption = limit - Len(.txtPoznamka.Value)
    'omezen� d�lky:
    If Len(.txtPoznamka.Value) > limit Then .txtPoznamka.Value = Left(.txtPoznamka.Value, limit)
    End With    'Me

End Sub



Private Sub lblCenaRovno_Click()
    'autom. text do pozn. k uz�v�rce
    With Me
    .txtPoznamka.Value = "cena dle cen�ku vydavatele" & .txtPoznamka.Value
    End With    'Me
End Sub

Private Sub lblCenaNizsi_Click()
    'autom. text do pozn. k uz�v�rce
    With Me
    .txtPoznamka.Value = "cena ni��� ne� cen�kov�" & .txtPoznamka.Value
    End With    'Me
End Sub



Private Sub txtReklPozn_Change()
    'zm�na d�lky pozn�mky k reklam�

Dim limit       As Integer

    limit = 255

    With Me
    .lblReklDelka.Caption = limit - Len(.txtReklPozn.Value)
    'omezen� d�lky:
    If Len(.txtReklPozn.Value) > limit Then .txtReklPozn.Value = Left(.txtReklPozn.Value, limit)
    End With    'Me

End Sub



Private Sub optSchvalit_Change()
    Call ZapnoutAdmin
End Sub


Private Sub optZamitnout_Change()
    With Me
    If .optZamitnout.Value = True Then
        .txtKamNeschval.Value = .txtKamCena.Value
    End If
    End With
    Call ZapnoutAdmin
End Sub


Private Sub ZapnoutAdmin()
    'schv�len�/zam�tnut� teprve zapne schv�len� adminem!
    With Me
    If .optSchvalit.Value = False And .optZamitnout.Value = False Then
        With .chkKamAdmin
            .Enabled = False
            .Value = False
        End With
    Else
        .chkKamAdmin.Enabled = True
    End If
    End With    'Me

End Sub



Private Sub chkKamAdmin_Change()
    On Error Resume Next    'p�i p�e�erp�n� se vypne tla��tko a nejde focus!!
    With Me
    If .chkKamAdmin.Value = True Then .cmdUlozit.SetFocus
    End With    'Me
End Sub



Private Sub optSchvalitRekl_Change()
    Call DokoncitZaznam
End Sub

Private Sub optZamitnoutRekl_Change()
    Call DokoncitZaznam
End Sub

Private Sub DokoncitZaznam()
    Me.cmdUlozit.SetFocus
End Sub
'========== HODNOCEN� - KONEC ==========



Private Sub cmdUlozit_Click()
    'tl. Ulo�it (z�znam)

Dim mesicDokl As Integer    '�. m�s�ce pro evidenci doklad�

    On Error GoTo Chyba

    'logick� kontrola dat:
    '=====================

    'kontroly pouze pro ��dosti:
    '---------------------------
    With Me
    If reklama = False Then

'        'reg. ��slo RF:
'        If .txtRegCislo.Value = "" Then
'            MsgBox "Nen� vypln�no reg. ��slo RF!      ", vbExclamation, "Chyb� povinn� �daj"
'            .txtRegCislo.SetFocus
'            Exit Sub
'        End If

        '��slo ��dosti:
        If .txtDlrZadost.Value = "" Then
            MsgBox "Nen� vypln�no ��slo ��dosti!      ", vbExclamation, "Chyb� povinn� �daj"
            .txtDlrZadost.SetFocus
            Exit Sub
        End If

        'celkem K�:
        If .txtKamCena.Value = "" Then
            MsgBox "Nen� vypln�na cena kampan�!      ", vbExclamation, "Chyb� povinn� �daj"
            .txtKamCena.SetFocus
            Exit Sub
        End If

        'p�evzat� bonus:
        If .chkBonus.Value = True And typBonus = "" Then
            If MsgBox("Bonus nen� p�evzat� ze schv�len�ch.     " _
            & vbNewLine & vbNewLine & "Je to v po��dku?", vbYesNo + vbExclamation, "") = vbNo Then
                .cmdPrevzitBonus.SetFocus
                Exit Sub
            End If
        End If

    End If

    'kontroly spole�n� pro reklamy/��dosti:
    '--------------------------------------
    'datum - rok:
    If Year(CDate(.txtDatum.Value)) <> rok Then
        If MsgBox("Datum z�znamu nepat�� do roku " & rok & "!      " _
            & vbNewLine & vbNewLine & "Je to v po��dku?", vbYesNo + vbExclamation, "") = vbNo Then
            .txtDatum.SetFocus
            Exit Sub
        End If
    End If

    'intern� �. kampan�:
    If .txtKamCislo.Value = "" Then
        MsgBox "Chyb� intern� ��slo kampan�!      ", vbExclamation, "Chyb� povinn� �daj"
        .txtKamCislo.SetFocus
        Exit Sub
    End If

    'oblast prodej/servis:
    If .optProdej.Value = False And .optServis.Value = False Then
        MsgBox "Zadejte oblast p��sp�vku prodej/servis!      ", vbExclamation, "Chyb� povinn� �daj"
        Exit Sub
    End If

    'n�zev kampan�:
    If .cmbKamNazev.Value = "" Then
        MsgBox "Nen� vypln�n n�zev kampan�!      ", vbExclamation, "Chyb� povinn� �daj"
        .cmbKamNazev.SetFocus
        Exit Sub
    End If

    'typ kampan�:
    If .cmbKamTyp.Value = "" Then
        MsgBox "Nen� vypln�n typ kampan�!      ", vbExclamation, "Chyb� povinn� �daj"
        .cmbKamTyp.SetFocus
        Exit Sub
    End If

    'zam��en� kampan�:
    If .cmbKamZamereni.Value = "" Then
        MsgBox "Nen� vypln�no zam��en� kampan�!      ", vbExclamation, "Chyb� povinn� �daj"
        .cmbKamZamereni.SetFocus
        Exit Sub
    End If

    'zdroj dat:
    If .cmbKamZdroj.Value = "" Then
        MsgBox "Nen� vypln�n zdroj dat!      ", vbExclamation, "Chyb� povinn� �daj"
        .cmbKamZdroj.SetFocus
        Exit Sub
    End If

    'form�t:
'    If .cmbKamFormat.Value = "" And .cmbKamZdroj.Value = "DAS" Then
'        MsgBox "Nen� vypln�n form�t!      ", vbExclamation, "Chyb� povinn� �daj"
'        .cmbKamFormat.SetFocus
'        Exit Sub
'    End If
    If .cmbKamFormat.Value = "" Then
        MsgBox "Nen� vypln�n form�t!      ", vbExclamation, "Chyb� povinn� �daj"
        .cmbKamFormat.SetFocus
        Exit Sub
    End If

    'kontrola oblasti �erp�n�:
    If (InStr(UCase(.cmbKamTyp.Value), "SERVIS") > 0 Or InStr(UCase(.cmbKamZamereni.Value), "SERVIS") > 0) And .optProdej.Value = True Then
        If MsgBox("Zkontrolujte Oblast �erp�n� a Typ/Zam��en� kampan�!   " & vbNewLine & vbNewLine & "Je to v po��dku?", _
        vbExclamation + vbYesNo + vbDefaultButton2, "Logick� kontrola dat") = vbNo Then Exit Sub
    End If

    'kontroly pouze pro reklamy:
    '---------------------------
    
'    If reklama = True Then
'        If .txtKamPocet.Value > 10 Then
'        If MsgBox("Po�et schv�len�ch rekl. dokument� je p��li� vysok�.      " _
'            & vbNewLine & vbNewLine _
'            & "Je to v po��dku?" _
'            & vbNewLine & vbNewLine _
'            & "(Zapisuje se po�et schv�len�ch dokument�, ne po�et vyroben�ch!)", vbYesNo + vbExclamation, "Ukl�d�n� rekl. dokumentu") = vbNo Then Exit Sub
'        End If
'    End If

    If reklama = True Then
        If .txtKamPocet.Value > 10 Then
        If MsgBox("Po�et schv�len�ch rekl. dokument� je p��li� vysok�.      " _
            & vbNewLine & vbNewLine _
            & "Zapi�te po�et schv�len�ch dokument�, ne po�et vyroben�ch!", vbOKOnly + vbCritical, "Ukl�d�n� rekl. dokumentu") = vbNo Then Exit Sub
        End If
    End If

    End With    'Me
    '=======================
    'konec kontroly dat

    'podle NW vyhledat DLR a z�skat index: (pro vyu�it� bonus� a z�pis do evidence doklad�)
    For i = LBound(dlrNw) To UBound(dlrNw)
        If Range("A1").Value = dlrNw(i) Then
            dlr = i
            GoTo KonecNw
        End If
    Next i
KonecNw:

    'odemknout list pro z�pis:
    ActiveSheet.Unprotect Password:=heslo

    'v�b�r ��dku pro z�pis kampan�/reklamy:
    If typZaznamu <> "edit" Then
        'pro nov� z�znam a kopii:
        poslRadek = Cells(Rows.Count, 2).End(xlUp).Row
        If poslRadek = 4 And Cells(poslRadek, 2).Value = "" Then    'tabulka je zat�m pr�zdn�
            radek = 4
        Else
            radek = poslRadek + 1
        End If
    Else
        'pro EDIT je ��dek z ActiveCell
        radek = ActiveCell.Row
    End If

    'z�pis polo�ek pouze pro kampan� (ostatn� jsou spole�n� pro kampan� i reklamy):
    '===============================
    If reklama = False Then
        With Cells(radek, 3)                                'reg. �. RF
            .NumberFormat = "@"
            .Value = CStr(txtRegCislo.Value)
        End With

        'kontrola �. ��dosti p�ed ulo�en�m - od v. 5.0
        '(zm�na data v editaci se neprojevila ve zm�n� ��sla, pokud se nezm�nil focus)
        Call GenerovatCisloZadosti

        With Cells(radek, 4)                                '�. ��dosti/fa
            .NumberFormat = "@"
            .Value = CStr(txtDlrZadost.Value)
            .HorizontalAlignment = xlCenter
        End With

        With Cells(radek, 6)                                '�. subdod. fa
            .NumberFormat = "@"
            .Value = CStr(txtSubFak.Value)
        End With

        'cena celkem - konverze ��rky na te�ku:
        With Cells(radek, 9)                                'cena celkem
    '        .Value = txtKamCena.Value
            .Value = Application.WorksheetFunction.Substitute(txtKamCena.Value, ",", ".")
            .NumberFormat = "#,##0.00 $"
        End With

        Cells(radek, 10).Value = txtKamProcento.Value        'procento p��sp�vku

        'schv�len� p��sp�vek - konverze ��rky na te�ku:
        With Cells(radek, 11)
            .Value = Application.WorksheetFunction.Substitute(txtKamSchval.Value, ",", ".")
            .NumberFormat = "#,##0.00 $"
        End With

    '    'neschv�len� ��stka - konverze ��rky na te�ku:
        With Cells(radek, 12)
            .Value = Application.WorksheetFunction.Substitute(txtKamNeschval.Value, ",", ".")
            .NumberFormat = "#,##0.00 $"
        End With

        'BONUS - kv�li p�episov�n� s kombinovan�m �erp�n�m p�esunut na konec bloku

        'z�pis hodnocen�:
        '----------------
        On Error Resume Next

        'hodnocen� 1:
        With Cells(radek, 22)
            If chkHodn1.Value = True Then
                .Value = "ano"
            Else
                .Value = "ne"
            End If
        'koment��:
            If txtHodn1.Value <> "" Then
                .AddComment
                .Comment.Text Text:=txtHodn1.Value
            Else
                .Comment.Delete
            End If
        End With

        'hodnocen� 2:
        With Cells(radek, 23)
            If chkHodn2.Value = True Then
                .Value = "ano"
            Else
                .Value = "ne"
            End If
        'koment��:
            If txtHodn2.Value <> "" Then
                .AddComment
                .Comment.Text Text:=txtHodn2.Value
            Else
                .Comment.Delete
            End If
        End With

        'hodnocen� 3:
        With Cells(radek, 24)
            If chkHodn3.Value = True Then
                .Value = "ano"
            Else
                .Value = "ne"
            End If
        'koment��:
            If txtHodn3.Value <> "" Then
                .AddComment
                .Comment.Text Text:=txtHodn3.Value
            Else
                .Comment.Delete
            End If
        End With

        'hodnocen� 4:
        With Cells(radek, 25)
            If chkHodn4.Value = True Then
                .Value = "ano"
            Else
                .Value = "ne"
            End If
        'koment��:
            If txtHodn4.Value <> "" Then
                .AddComment
                .Comment.Text Text:=txtHodn4.Value
            Else
                .Comment.Delete
            End If
        End With

        'hodnocen� 5:
        With Cells(radek, 26)
            If chkHodn5.Value = True Then
                .Value = "ano"
            Else
                .Value = "ne"
            End If
        'koment��:
            If txtHodn5.Value <> "" Then
                .AddComment
                .Comment.Text Text:=txtHodn5.Value
            Else
                .Comment.Delete
            End If
        End With

        'hodnocen� 6: - POZN�MKY K REKLAMN�MU DOKUMENTU!!!
        With Cells(radek, 27)
            If chkHodn6.Value = True Then
                .Value = "ano"
            Else
                .Value = "ne"
            End If
        'koment��:
            If txtHodn6.Value <> "" Then
                .AddComment
                .Comment.Text Text:=txtHodn6.Value
            Else
                .Comment.Delete
            End If
        End With
        'konec hodnocen�

        With Cells(radek, 33)                                       'kombinovan� �erp�n�
            If Me.chkKombinace.Value = True Then
                .Value = "ano"
                Cells(radek, 1).Interior.Color = RGB(230, 185, 184)
            Else
                .Value = "ne"
                Cells(radek, 1).Interior.Color = xlNone
            End If
        End With
    
        With Cells(radek, 13)               'BONUS - s typem bonusu do koment��e
            If chkBonus.Value = True Then
                .Value = "ano"
'                If typBonus <> "" Then
                    On Error Resume Next
                    .AddComment
                    .Comment.Text Text:=typBonus
                Cells(radek, 1).Interior.Color = RGB(215, 228, 188)
            Else
                .Value = "ne"
                On Error Resume Next
                .Comment.Delete
                Cells(radek, 1).Interior.Color = xlNone
            End If
            .HorizontalAlignment = xlCenter
        End With
    
        'schv�len�/zam�tnut� kampan� - z�pis do sl. 14 a barva p�sma ��dku:
        '---------------------------
        With Cells(radek, 14)
            If optSchvalit.Value = True Then
                'schv�len� z�znam:
                .Value = "ano"
                Range("A" & radek & ":AZ" & radek).Font.Color = RGB(0, 176, 80)
            ElseIf optZamitnout.Value = True Then
                'zam�tnut� z�znam:
                .Value = "ne"
                Range("A" & radek & ":AZ" & radek).Font.Color = RGB(255, 0, 0)
            Else
                'rozpracovan� z�znam:
                .Value = ""
                Range("A" & radek & ":AZ" & radek).Font.Color = RGB(0, 0, 0)
            End If
        End With
    
        'schv�len� administr�torem - z�pis do sl. 34 a v�pl� ��sla kampan�:
        '---------------------------
        With Cells(radek, 34)
            If Me.chkKamAdmin.Value = True Then
                'schv�leno adminem:
                .Value = "ano"
                Range("B" & radek).Interior.Color = vbYellow
            Else
                'zat�m neschv�leno adminem:
                .Value = "ne"
                With Range("B" & radek).Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        End With
    
        'stav uz�v�rky - po z�pisu je NE:
        '--------------------------------
        With Cells(radek, 35)
            .Value = "ne"
            If Len(txtPoznamka.Value) > 0 Then
                .AddComment
                .Comment.Text Text:=txtPoznamka.Value
            Else
                .Comment.Delete
            End If
        End With

    End If  'konec z�pisu polo�ek pouze pro kampan�


    'z�pis polo�ek pro kampan� i reklamy:
    '====================================

    Cells(radek, 1).Value = CDate(txtDatum.Value)       'datum
    Cells(radek, 1).NumberFormat = "d/m/yyyy;@"         'oprava pro Exc2013 (od v. 5.3.1)
    Cells(radek, 2).Value = txtKamCislo.Value           'intern� �. kampan�/rekl.
    Cells(radek, 7).Value = cmbKamNazev.Value           'n�zev kam.

    With Cells(radek, 8)                                'oblast p��sp�vku
        If Me.optProdej.Value = True Then
            .Value = "prodej"
        ElseIf Me.optServis.Value = True Then
            .Value = "servis"
        End If
    End With

    With Me
    Cells(radek, 15).Value = .cmbKamTyp.Value           'typ kam.
    Cells(radek, 16).Value = .cmbKamZamereni.Value      'zam��en� kam.
    Cells(radek, 17).Value = .cmbKamMedium.Value        'm�dium-typ
    Cells(radek, 18).Value = .cmbKamMediumNazev.Value   'm�dium-n�zev
    Cells(radek, 19).Value = .cmbKamZdroj.Value         'zdroj
'    If .cmbKamZdroj.Value = "PP" Then
'        On Error Resume Next
'        Cells(radek, 19).AddComment
'        If .chkDas.Value = True Then        'data z DASu
'            Cells(radek, 19).Comment.Text Text:="ano"
'        Else
'            Cells(radek, 19).Comment.Text Text:="ne"
'        End If
'    Else
'        On Error Resume Next
'        Cells(radek, 19).Comment.Delete
'    End If
    Cells(radek, 20).Value = .cmbKamFormat.Value        'form�t
    Cells(radek, 21).Value = .txtKamPocet.Value         'po�et ks mat.
    End With    'Me

    mesicDokl = Month(Cells(radek, 1).Value)
    Cells(radek, 28).Value = mesicDokl                      'm�s�c
    Cells(radek, 29).Value = Year(Cells(radek, 1).Value)    'rok
    If typZaznamu <> "edit" Then
        With Cells(radek, 30)
            .Value = userAp                         'kod u�ivatele - JEN PRO NOV� A KOPIE
'            .AddComment
'            .Comment.Text Text:=Day(Date) & "." & Month(Date) & "." & Year(Date)    'datum z�pisu
        End With
    End If

    With Cells(radek, 31)
        If Me.optKamVlastni.Value = True Then
            .Value = "ano"                              'vlastn� kampa� DLR
        Else
            .Value = "ne"                               'kampa� import�ra
        End If
    End With

    With Cells(radek, 32)
        If Me.chkSpolReklama.Value = True Then
            .Value = "ano"                              'spole�n� reklama v�ce DLR
        Else
            .Value = "ne"
        End If
    End With

    'datum z�pisu (nov� a kopie):
    If typZaznamu <> "edit" Then
        Cells(radek, 37).Value = CDate(Day(Date) & "." & Month(Date) & "." & Year(Date))
        Cells(radek, 37).NumberFormat = "d/m/yyyy;@"
    End If

    'LOG - sledov�n� zm�n m�s�ce v z�znamu:
    If reklama = False And typZaznamu = "edit" And mesicPuvodni <> mesicDokl Then
        Call frm_MENU.SysEvent _
            (sysZprava:=txtKamCislo.Value & " - zm�na m�s�ce z " & mesicPuvodni & " na " & mesicDokl)
    End If

    'z�pis polo�ek pouze pro reklamy:
    '================================
    If reklama = True Then

        With Cells(radek, 13)               'BONUS - reklama bez typu bonusu do koment��e
            If chkBonus.Value = True Then
                .Value = "ano"
            Else
                .Value = "ne"
            End If
            .HorizontalAlignment = xlCenter
        End With

        'schv�len�/zam�tnut� reklamy - z�pis do sl. 14 a barva p�sma ��dku:
        '---------------------------
        With Cells(radek, 14)
            If optSchvalitRekl.Value = True Then
                'schv�len� z�znam:
                .Value = "ano"
                Range("A" & radek & ":AZ" & radek).Font.Color = RGB(0, 176, 80)
            ElseIf optZamitnoutRekl.Value = True Then
                'zam�tnut� z�znam:
                .Value = "ne"
                Range("A" & radek & ":AZ" & radek).Font.Color = RGB(255, 0, 0)
            Else
                'rozpracovan� z�znam:
                .Value = ""
                Range("A" & radek & ":AZ" & radek).Font.Color = RGB(0, 0, 0)
            End If

On Error Resume Next    'kvuli .Comment.Delete
            'zapsat koment�� ke schv�len� kampan�:
            If Me.txtReklPozn.Value <> "" Then
                .AddComment
                .Comment.Text Text:=Me.txtReklPozn.Value
            Else
                .Comment.Delete
            End If
        End With    'konec pr�ce s Cells(radek, 14)

On Error GoTo Chyba
        
        If typZaznamu = "edit" Then
            Cells(radek, 34).Value = vyuzito                    'vyu�it� na�ten� p�i editaci
        Else
            Cells(radek, 34).Value = "ne"                       'default vyu�it� nov� rekl.
        End If

    End If      'konec z�pisu polo�ek pouze pro reklamy
    '--------------------------------------------------

    'p�smo ��dku standard:
    '---------------------
    Range("A" & radek & ":AZ" & radek).Font.Italic = False

    'p�epo�et rozpo�tu DLR:
    '----------------------
    If reklama = False Then Call PrepocetDlr

'    DEBUG
'    ActiveWindow.DisplayHeadings = False    'vypnout ��sla ��dk� a sloupc�

    'zapsat vyu�it� reklamy do kampan�:
    '----------------------------------
    If idReklamy <> -100 Then       'default - nastaveno v INIT

        If spolReklama = True Then
            Call OtevritSpolecnouReklamu
            Set wsDlrData2 = ActiveWorkbook.Worksheets(2)   'p�ep�e se p�vodn� prom�nn� - z listu DLR na 999SR!
        End If

        'd�le stejn� p��kazy pro z�pis vyu�it� rekl. DLR/999SR/Retail:
        wsDlrData1.Cells(radek, 36).Value = wsDlrData2.Cells(idReklamy + 4, 2).Value    'do kampan� zapsat �. zdrojov� rekl.
        With wsDlrData2
            .Unprotect Password:=heslo
            .Cells(idReklamy + 4, 34).Value = "ano"
            .Cells(idReklamy + 4, 35).Value = .Cells(idReklamy + 4, 35).Value + 1
            .Protect Password:=heslo
        End With

        'z�pis vyu�it� SR/Retail - n�vrat k souboru DLR:
        If spolReklama = True Then
            With ActiveWorkbook
                .Save
                .Close                                      'zav��t 999SR
            End With
            Windows(soubor).Activate                        'aktivace souboru DLR
            Set wsDlrData2 = ActiveWorkbook.Worksheets(2)   'n�vrat prom�nn� na list reklam DLR!
        End If

    End If

    'zapsat vyu�it� Bonusu R/I do rozpo�t�:
    '--------------------------------------
    If idBonusy <> -100 Then       'default - nastaveno v INIT

        'otev��t rozpo�ty:
        With Application
        .ScreenUpdating = False
        .StatusBar = "Zapisuji vyu�it� IP..."
        End With
        Set wbRozpocty = GetObject(cesta & "\Data\Rozpocty\" & souborRozpocty)
        Set wsRozpocty = wbRozpocty.Worksheets(1)
    
        'zapsat vyu�it� do rozpo�t�:
        If typBonus = "R" Then          'Region bonus
            If idBonusy = 0 Then        'jarn�
                wsRozpocty.Cells(dlr + 1, 109).Value = "ano"
            ElseIf idBonusy = 1 Then    'podzimn�
                wsRozpocty.Cells(dlr + 1, 112).Value = "ano"
            Else
                MsgBox "idBonusy nen� v po��dku", vbCritical, "Rozpo�ty - vyu�it� Bonusu"
            End If
        ElseIf typBonus = "I" Then      'Indiv. bonusy v�. centr�ln�ch
            wsRozpocty.Cells(dlr + 1, idBonusy + 154).Value = "ano"
        Else
            MsgBox "Typ bonusu nen� v po��dku", vbCritical, "Rozpo�ty - vyu�it� Bonusu"
        End If

        'zav��t rozpo�ty:
        Set wsRozpocty = Nothing
        Windows(souborRozpocty).Visible = True     'p�i ukl�d�n� rozpo�tu po otev�en� metodou GetObject (se�it z�st�val skryt�)
        Application.ReferenceStyle = xlA1
        With wbRozpocty
        .Save
        .Close
        End With
        Set wbRozpocty = Nothing
        With Application
        .ScreenUpdating = True
        .StatusBar = False
        End With

    End If
    '------------------
    'konec z�pisu ��dku

    Cells(radek, 1).Select      'nastavit zapsan� ��dek, hlavn� pro tisk ko�ilky

    'zamknout list po z�pisu ��dku:
    '------------------------------
    ActiveSheet.Protect Password:=heslo

    'ulo�en� souboru DLR:
    '--------------------
    With Application
        .StatusBar = "UKL�D�N� SOUBORU:  " & ActiveWorkbook.Name & "..."
        .ReferenceStyle = xlA1
    ActiveWorkbook.Save
        .StatusBar = False
    End With

    'kontrola �. ��dosti, p�i zm�n� upozorn�n� na ko�ilku pro novou ��dost DLR:
    '-------------------------------------------------------------------------
    If reklama = False Then
        If Cells(radek, 4).Value <> Cells(radek - 1, 4).Value Then
            MsgBox "Zm�nilo se ��slo ��dosti, mo�n� bude dobr� vygenerovat ko�ilku...      ", vbInformation, aplikace
        End If
    End If

    'admin editace - skok o ��dek dol�:
    If reklama = False And admin = True Then _
        Cells(radek + 1, 1).Select

    'reset p��znaku spole�n� reklama:
    spolReklama = False

    'zapsat do syslogu:
    If typZaznamu <> "edit" Then    'nov� z�znam / kopie
        If reklama = False Then
            Call frm_MENU.SysEvent(sysZprava:=dlrNw(dlr) & " - z�pis kampan� " & txtKamCislo.Value & " (" & typZaznamu & ")")
        Else
            Call frm_MENU.SysEvent(sysZprava:=dlrNw(dlr) & " - z�pis rekl. dokumentu " & txtKamCislo.Value & " (" & typZaznamu & ")")
        End If
    Else                            'editace
        If reklama = False Then
            Call frm_MENU.SysEvent(sysZprava:=dlrNw(dlr) & " - editace kampan� " & txtKamCislo.Value)
        Else
            Call frm_MENU.SysEvent(sysZprava:=dlrNw(dlr) & " - editace rekl. dokumentu " & txtKamCislo.Value)
        End If
    End If
    If idBonusy <> -100 Then
        If typBonus = "I" Then
            Call frm_MENU.SysEvent(sysZprava:=dlrNw(dlr) & " - vyu�it� IP " & idBonusy + 1)
        ElseIf typBonus = "R" Then
            Call frm_MENU.SysEvent(sysZprava:=dlrNw(dlr) & " - vyu�it� RB " & idBonusy + 1)
        End If
    End If

    'zapsat status do evidence doklad�:     MRO 2013
    '----------------------------------
    If reklama = False Then
        With Application
        .StatusBar = "Zapisuji do Evidence doklad�..."
        .ScreenUpdating = False
        End With
        Workbooks.Open (cesta & "\Data\Doklady\DLR_doklady_" & typMro & ".xlsm")
        ActiveSheet.Unprotect Password:=heslo
        Cells(dlr + 1, mesicDokl + 2).Value = CDate(Day(Date) & "." & Month(Date) & "." & Year(Date))
        'status = form�t:
        Select Case Me.chkHodn3.Value   'kontrola p��znaku ��dost zasl�na
        Case True
            Cells(dlr + 1, mesicDokl + 2).Font.Color = RGB(0, 176, 80)
            Call frm_MENU.DocEvent _
                (sysZprava:=Cells(dlr + 1, 1).Value & " - do�la ��dost za m�s�c " & mesicSoub & "(z�pis kampan�)")
        Case False
            Cells(dlr + 1, mesicDokl + 2).Font.Color = RGB(0, 0, 0)
            Call frm_MENU.DocEvent _
                (sysZprava:=Cells(dlr + 1, 1).Value & " - do�ly podklady za m�s�c " & mesicSoub & "(z�pis kampan�)")
        End Select
    '    ActiveSheet.Protect Password:=heslo
        ActiveSheet.Protect
        With ActiveWorkbook
        .Save
        .Close
        End With
        With Application
        .ReferenceStyle = xlA1
        .StatusBar = False
        .ScreenUpdating = True
        End With
    End If

    'vytvo�it slo�ku v dokumentech NW:      MRO 2013
    '---------------------------------
    '(kontrola, pokud se nevytvo�� p�es DLR info)
    If reklama = False Then
        mesicSoub = CStr(mesicDokl)
        If Len(mesicSoub) = 1 Then
            mesicSoub = "0" & mesicSoub
        End If
        If frm_MENU.PathExists(cestaDokNw & "\" & dlrNw(dlr)) = False Then _
            MkDir (cestaDokNw & "\" & dlrNw(dlr))
        If frm_MENU.PathExists(cestaDokNw & "\" & dlrNw(dlr) & "\" & rok) = False Then _
            MkDir (cestaDokNw & "\" & dlrNw(dlr) & "\" & rok)
        If frm_MENU.PathExists(cestaDokNw & "\" & dlrNw(dlr) & "\" & rok & "\" & mesicSoub & "_" & typMro) = False Then
            MkDir (cestaDokNw & "\" & dlrNw(dlr) & "\" & rok & "\" & mesicSoub & "_" & typMro)
            Call frm_MENU.DocEvent(sysZprava:=dlrNw(dlr) & " - vytvo�ena slo�ka " & mesicSoub & "_" & typMro & "(z�pis kampan�)")
        End If
    End If

'    appEvents = False
    Unload Me
'    appEvents = True
    Exit Sub

Chyba:  'p�i chyb� programu zamknout list
'    MsgBox Err.Number & " " & Err.Description, vbCritical, "===  CHYBA  ==="
'    ActiveSheet.Protect Password:=heslo         'zamknout list

End Sub



Private Sub PrepocetDlr()   'pouze pro kampan�
    'p�epo�et rozpo�tu DLRa
    'vol�no z cmdUlozit_Click

    'p�i p�epo�tu kontrolovat bonus akce, schv�len� kampan�, oblast p��sp�vku!!!

Dim cerpProdej      As Single   '�erp�n� p��sp. na prodej
Dim cerpServis      As Single   '�erp�n� p��sp. na servis
Dim cerpCelkem      As Single   '�erp�n� celkem

    cerpProdej = 0
    cerpServis = 0
    cerpCelkem = 0

    'na��st �erp�n� ze zapsan�ch kampan�:
    '------------------------------------
    poslRadek = Cells(Rows.Count, 2).End(xlUp).Row

    For i = 4 To poslRadek
        If Cells(i, 13).Value = "ne" And Cells(i, 14).Value = "ano" Then    '��dek nen� bonus akce a je schv�len�!!
            cerpCelkem = cerpCelkem + Cells(i, 11).Value                    'p�i��st �erp. celkem
            If Cells(i, 8).Value = "prodej" Then
                cerpProdej = cerpProdej + Cells(i, 11).Value                'p�i��st �erp�n� prodej
            ElseIf Cells(i, 8).Value = "servis" Then
                cerpServis = cerpServis + Cells(i, 11).Value                'p�i��st �erp�n� servis
            End If
        End If
    Next i

End Sub



Private Sub cmdStorno_Click()
    'tl. Storno

    If reklama = False And InStr(ActiveWorkbook.Name, "999SR") > 0 Then   'pouze pro kampan�!
        ActiveWorkbook.Close
        Windows(soubor).Activate    'aktivace souboru DLR
    End If
    Set wsDlrData1 = Nothing
    Set wsDlrData2 = Nothing
    Unload frm_ZAZNAM
    
End Sub



Private Function FileExists(fname) As Boolean
    ' vrac� TRUE, pokud soubor existuje

    FileExists = (Dir(fname) <> "")

End Function     '----- End of Function FileExists -----



Private Function PathExists(pname) As Boolean
    ' vrac� TRUE, pokud cesta existuje

    If Dir(pname, vbDirectory) = "" Then
        PathExists = False
    Else
        PathExists = (GetAttr(pname) And vbDirectory) = vbDirectory
    End If

End Function     '----- End of Function PathExists -----



'============== P�EVZET� REKLAMY DO KAMPAN�: ===================

Private Sub cmdPrevzitReklamu_Click()
    'p�evz�t z�znam ze schv�len�ch reklam DLR

    frmReklamy.Show

End Sub



Private Sub cmdPrevzitReklamu2_Click()
    'p�evz�t z�znam ze spole�n�ch reklam (DLR 999SR)

    'otev��t soubor spol. reklam podle instalace:
    Call OtevritSpolecnouReklamu
    frmReklamy.Show

End Sub



Private Sub OtevritSpolecnouReklamu()
    'otev��t soubor spol. reklam podle instalace:

    Workbooks.Open (cesta & "\Data\DLR_" & typMro & "\999SR - SPOLE�N� REKLAMA [" & typMro & "].xlsm")

End Sub



Sub NacistReklamu()
    'na��st data vybran� reklamy do frm
    'vol�no z frmReklamy.cmdPrevzit

Dim radek2      As Long

    appEvents = False

    On Error GoTo Chyba
    radek2 = idReklamy + 4   'listindex od 0, z�hlav� 3 ��dky

    'info o p�evzet� reklamy:
    With Me.lblZdrojovaReklama
        .Caption = "Zdrojov� reklama: " & wsDlrData2.Cells(radek2, 2).Value & "."
        .Visible = True
    If wsDlrData2.Cells(radek2, 34).Value = "ano" Then
        .Caption = .Caption & " Ji� vyu�ito: " & wsDlrData2.Cells(radek2, 35).Value & "x."
    Else
        .Caption = .Caption & " Je�t� nevyu�ito."
    End If
    End With    'lblZdrojovaReklama

    'z�pis p�evzet� do logu:
    Call frm_MENU.SysEvent(sysZprava:=dlrNw(dlr) & " - vyu�it� rekl. dokumentu (" & wsDlrData2.Cells(radek2, 2).Value & " / " & Me.txtKamCislo.Value & ")")
    
    'data reklamy do kampan�:
    With wsDlrData2

    If .Cells(radek2, 13).Value = "ano" Then             'bonus akce
        Me.chkBonus.Value = True
    Else
        Me.chkBonus.Value = False
    End If

    If .Cells(radek2, 32).Value = "ano" Then             'spole�n� rekl.
        Me.chkSpolReklama.Value = True
    Else
        Me.chkSpolReklama.Value = False
    End If

    If .Cells(radek2, 31).Value = "ne" Then              'vlastn� kam.
        Me.optKamImport.Value = True
    Else
        Me.optKamVlastni.Value = True
    End If

    If .Cells(radek2, 8).Value = "prodej" Then                'oblast �erp�n�
        Me.optProdej.Value = True
    Else
        Me.optServis.Value = True
    End If

    Me.cmbKamNazev.Value = .Cells(radek2, 7).Value               'n�zev kam.
    Me.cmbKamTyp.Value = .Cells(radek2, 15).Value                'typ kam.
    Me.cmbKamZamereni.Value = .Cells(radek2, 16).Value           'zam��en� kam.
    Me.cmbKamMedium.Value = .Cells(radek2, 17).Value             'typ m�dia
    Me.cmbKamMediumNazev.Value = .Cells(radek2, 18).Value        'n�zev m�dia
    Me.cmbKamZdroj.Value = .Cells(radek2, 19).Value              'zdroj
'    If .Cells(radek2, 19).Value = "PP" Then                 'PP-DAS
'        Me.chkDas.Visible = True
'        If .Cells(radek2, 19).Comment.Text = "ano" Then
'            Me.chkDas.Value = True
'        Else
'            Me.chkDas.Value = False
'        End If
'    Else
'        With Me.chkDas
'        .Visible = False
'        .Value = False
'        End With
'    End If
    Me.cmbKamFormat.Value = .Cells(radek2, 20).Value             'form�t
    If .Cells(radek2, 14).Value = "ano" Then Me.chkHodn6.Value = True       'schv�len� reklamy
    On Error Resume Next    'pr�zdn� koment��:
    Me.txtHodn6.Value = .Cells(radek2, 14).Comment.Text          'pozn�mka k rekl. dokumentu
    On Error GoTo Chyba

    End With    'wsDlrData2
    Me.txtDatum.SetFocus
    Exit Sub

    appEvents = True

Chyba:
'    MsgBox Err.Number & " " & Err.Description, vbCritical, aplikace

End Sub



Sub NacistBonus()
    'na��st data vybran�ho bonusu do frm
    'vol�no z frmBonusy.cmdPrevzit

    'zapsat data do frm:
    With Me
    If typBonus = "I" Then
        .txtHodn6.Value = frmBonusy.lstBonusInd.List(idBonusy, 1)
    ElseIf typBonus = "R" Then
        .txtHodn6.Value = frmBonusy.lstBonusReg.List(idBonusy, 1)
    Else
        MsgBox "Typ bonusu nen� v po��dku, kontaktujte podporu.", vbCritical, "Rozpo�ty - vyu�it� Bonusu"
    End If
    '    .txtKamCena.Value = frmBonusy.lstBonusInd.List(idBonusy, 0)

    'zapsat indik�tor do frm:
    With .lblZdrojovyBonus
    .Visible = True
    .Caption = "P�evzato: " & idBonusy + 1 & " / " & typBonus
    End With
    .lblTypBonus.Caption = typBonus

    End With    'Me

End Sub



Private Sub imgDokumenty_Click()
    
    'podle NW vyhledat DLR a z�skat index: (pro vyu�it� bonus� a z�pis do evidence doklad�)
    For i = LBound(dlrNw) To UBound(dlrNw)
        If Range("A1").Value = dlrNw(i) Then
            dlr = i
            GoTo KonecNw
        End If
    Next i
KonecNw:
    Call Shell("explorer.exe " & cestaDokNw & "\" & dlrNw(dlr), vbNormalFocus)
End Sub



Private Sub imgTcmd_Click()

    'podle NW vyhledat DLR a z�skat index: (pro vyu�it� bonus� a z�pis do evidence doklad�)
    For i = LBound(dlrNw) To UBound(dlrNw)
        If Range("A1").Value = dlrNw(i) Then
            dlr = i
            GoTo KonecNw
        End If
    Next i
KonecNw:
    Call Shell(cestaTcmd & "\totalcmd.exe /O /T /R= " & cestaDokNw & "\" & dlrNw(dlr), vbNormalFocus)

End Sub



Private Sub imgCalc_Click()
    'spustit ext. kalk.
    Set myShell = CreateObject("Wscript.Shell")
    myShell.Run cestaMroManager & "\Common_Files\Calc\MoffFreeCalc.exe", 1, True
    Set myShell = Nothing
End Sub



'-------------------- otev�en� ��seln�k� --------------------

Private Sub lblCiselnik1_Click()
    ciselnik = 1
    frmCiselniky.Show
End Sub


Private Sub lblCiselnik2_Click()
    ciselnik = 2
    frmCiselniky.Show
End Sub


Private Sub lblCiselnik3_Click()
    ciselnik = 3
    frmCiselniky.Show
End Sub


Private Sub lblCiselnik4_Click()
    ciselnik = 4
    frmCiselniky.Show
End Sub


Private Sub lblCiselnik5_Click()
    ciselnik = 5
    frmCiselniky.Show
End Sub


Private Sub lblCiselnik6_Click()
    ciselnik = 8
    frmCiselniky.Show
End Sub




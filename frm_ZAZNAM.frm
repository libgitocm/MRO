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

Dim vyuzito     As String       'pro edit reklamy - údaj, zda byla již využitá pro kampaò
Dim cisRadek    As String       'øádek èíselníku pro naètení
Dim mesicPuvodni As Integer     'pro zápis zmìn mìsíce (evidence dokl., složky!)
'Dim dlr         As Long         'index DLRa - pro hledání v polích - je PUBLIC!!
Dim cerpaniP    As Single   'pro okno kontroly èerpání
Dim cerpaniS    As Single   '(naètou se èerpání z nezamítnutých záznamù)



Private Sub UserForm_Initialize()

    'poloha a rozmìry frm:
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

    soubor = ActiveWorkbook.Name    'jméno souboru pro práci se spoleènou reklamou
    '(spol. rekl. - Storno = promìnná pro aktivaci souboru DLR)
    spolReklama = False     'default
    ciselnikZaznam = True   'zpìtná vazba z frmCiselnik do frm_ZAZNAM - naète se editovaný èíselník

    Set wsDlrData1 = ActiveWorkbook.Worksheets(1)                   'aktuální objekt wsDlrData1 - list žádostí
    Set wsDlrData2 = ActiveWorkbook.Worksheets(2)                   'aktuální objekt wsDlrData2 - list reklam

    'loga instalací:
    .imgSkodaPlus.Visible = skodaPlus

    'ikona TC:
    If cestaTcmd <> "" Then
        .imgTcmd.Visible = True
    Else
        .imgTcmd.Visible = False
    End If

    'podle NW vyhledat DLR a získat index: (pro kontroly èerpání, využití bonusù a zápis do evidence dokladù)
    For i = LBound(dlrNw) To UBound(dlrNw)
        If Range("A1").Value = dlrNw(i) Then
            dlr = i
            GoTo KonecNw
        End If
    Next i
KonecNw:

    'režim REKLAMA:
    If reklama = True Then
        Call ZaznamRezimReklama
    End If

    'nastavení pro externí MRO:
    '--------------------------
    .lblHelp1.Visible = Not externi
    .lblCiselnik1.Visible = Not externi
    .lblCiselnik2.Visible = Not externi
    .lblCiselnik3.Visible = Not externi
    .lblCiselnik4.Visible = Not externi
    .lblCiselnik5.Visible = Not externi
    .lblCiselnik6.Visible = Not externi
    .chkBonus.Visible = Not externi

    'režim KAMPAÒ - detekce pøevzetí reklamy:
    idReklamy = -100    'pokud se nezmìní, pøi uložení záznamu se neukládá využití reklamy
    idBonusy = -100

    'výbìr listu - provádí se v frm_MENU.KontrolaListu

    'label Neaktivní:
    If UCase(Range("G2").Value) = "PLATNÝ: NE" Or UCase(Range("O2").Value) = "PLATNÝ: NE" Then
        .lblNeaktivni.Visible = True
        .cmdUlozit.Enabled = False
    Else
        .lblNeaktivni.Visible = False
    End If

    'Bonus:
    .lblAlertBonus.Visible = .chkBonus.Value
    .cmdPrevzitBonus.Visible = .chkBonus.Value

    'pøevzatý bonus:
    typBonus = ""   'default!
    .lblZdrojovyBonus.Visible = False   'zapne se až výbìrem IP
    .lblTypBonus.Visible = False
    .lblTypBonus.Caption = typBonus

    'label Spoleèná rekl:
    .lblAlertSpolecna.Visible = .chkSpolReklama.Value

    'šipka pro kopii reg. èísla a èísla žádosti:
    .lblKopieCisel.Visible = False    'zapne se pouze pro EDIT kampanì

    'option Kampaò:
    .optKamImport.Value = True       'default

    'názvy kampaní:
    Call NacistCiselnikKamNazev

    'combo Typ kampanì:
    Call NacistCiselnikKamTyp

    'combo Zamìøení kampanì:
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
    .AddItem "vlastní"
    End With
'    .chkDas.Visible = False

    'combo Formát:
    Call NacistCiselnikKamFormat

    'popisky hodnocení:
    '------------------
    .chkHodn1.Caption = cisHodnoceni(1)
    .chkHodn2.Caption = cisHodnoceni(2)
    .chkHodn3.Caption = cisHodnoceni(3)
    .chkHodn4.Caption = cisHodnoceni(4)
    .chkHodn5.Caption = cisHodnoceni(5)
    .chkHodn6.Caption = cisHodnoceni(6)
    '------------------

    'viditelnost popiskù hodnocení podle nastavení:
    '----------------------------------------------
    '(chkbox se skryje a nastaví na True, aby fungoval logický souèin pro schválení)
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

    'poznámka k uzávìrce:
    .txtPoznamka.Value = ""

    'poznámka ke schválení reklamy:
    .txtReklPozn.Value = ""

    'opt Schválit:
    .optSchvalit.Enabled = False

    'režim ADMIN - možnost schválení kampanì adminem:
    With .chkKamAdmin
    .Visible = admin
    .Value = False
    .Enabled = False    'zapne se až pøi schválení/zamítnutí kam.!!
    End With
    .lblKamAdmin.Visible = admin

    'pøehled èerpání:
    '----------------
    'rozpoèty:
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

    'èerpání pøed zapisovanou kampaní:
    cerpaniP = 0
    cerpaniS = 0
    poslRadek = Cells(Rows.Count, 2).End(xlUp).Row
    For i = 4 To poslRadek
        If Cells(i, 13).Value = "ne" And Cells(i, 14).Value <> "ne" Then   'pouze NEbonus a NEzamístnuté
            'naèíst podle oblasti èerpání:
            If Cells(i, 8).Value = "prodej" Then
                cerpaniP = cerpaniP + Cells(i, 11).Value
            Else
                cerpaniS = cerpaniS + Cells(i, 11).Value
            End If
        End If
    Next i

    Call VypocetCeny    ' spustit pro výpoèet zùstatku pøed kampaní
    If typZaznamu = "edit" Then
        .lblZustP1.Visible = False
        .lblZustS1.Visible = False
        .lblKc2.Visible = False
        .lblKc5.Visible = False
    End If

    'èerpání po kampani:
    .lblZustP2.Caption = ""
    .lblZustS2.Caption = ""
    
    '----------------------
    'konec pøehledu èerpání

    End With    'Me

  'nastavení FRM podle typu záznamu:
  '---------------------------------
  
  'NOVY:
  '-----
    If typZaznamu = "novy" Then
      '==================
        With Me

        Call GenerovatCisloKampane
'        If reklama = False Then Call GenerovatDatum     'pouze pro kampanì!!

        .txtDatum.SetFocus
'        If reklama = False Then SendKeys ("^a")   'oznaèit datum

        If skodaPlus = False Then
            .txtKamProcento.Value = prispevek
        Else
            Call NajitPenetraci     'podle NW najít penetraci a pøiøadit % pøíspìvku
        End If
        .txtKamSchval.Value = 0      'výchozí pøíspìvky
        .txtKamNeschval.Value = 0
        .txtKamPocet.Value = 1    'výchozí poèet mat. došlých ke schválení

        'zkopírovat è. žádosti a reg. è. RF z posledního øádku:
        '(poslRadek vrací už GenerovatCisloKampane)
        .txtRegCislo.Value = Cells(poslRadek, 3).Value    'reg. è. RF
        .txtDlrZadost.Value = Cells(poslRadek, 4).Value      'è. žádosti

        'nastavení podle èinnosti DLR:
        If reklama = False Then
            If Range("G1").Value = "Èinnost: servis" Then
                .optServis.Value = True
            ElseIf Range("G1").Value = "Èinnost: ŠKODA Plus" Then
                .optProdej.Value = True
            End If
        Else
            If Range("O1").Value = "Èinnost: servis" Then
                .optServis.Value = True
            ElseIf Range("O1").Value = "Èinnost: ŠKODA Plus" Then
                .optProdej.Value = True
            End If
        End If

        'zápis reklamy - autom. pøíznak spoleèná reklama:
        If reklama = True And InStr(ActiveWorkbook.Name, "999SR") > 0 Then
            With .chkSpolReklama
            .Value = True
            .Enabled = False
            End With
        End If
        .lblUser.Caption = userAp
        End With    'Me

        'nastavení pro externí MRO:
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

        'naèíst øádek do FRM - kampaò:
        '-----------------------------
        Call NacistRadek        'naèíst položky spoleèné pro KOPII a EDIT

        If reklama = False Then
            If Cells(radek, 33).Value = "ano" Then          'kombin. èerpání
                With .chkKombinace
                .Value = True
                .Locked = True
                End With
            Else
                .chkKombinace.Value = False
            End If
        End If

        .txtKamPocet.Value = 1    'výchozí poèet mat. došlých ke schválení

        'zkopírovat è. žádosti a reg. è. RF z posledního øádku:
        '(poslRadek vrací už GenerovatCisloKampane)
        If MsgBox("Chcete zachovat údaje ze zdrojové kampanì?      " & vbNewLine & vbNewLine _
            & "(Datum,  Reg. èíslo,  È. žádosti,  È. subdod. faktury)      ", 36, "Kopie záznamu") = vbYes Then
            .txtDatum.Value = Day(Cells(radek, 1).Value) & "." & Month(Cells(radek, 1).Value) & "." & Year(Cells(radek, 1).Value)
            .txtRegCislo.Value = Cells(radek, 3).Value    'reg. è. RF
            .txtDlrZadost.Value = Cells(radek, 4).Value      'è. žádosti
            .txtSubFak.Value = Cells(radek, 6).Value      'è. sub. fak.
        End If

        With .txtDatum
'        .Value = datum
        .SetFocus
        End With

        'kopie záznamu generovaného ze schválené reklamy:
        If InStr(Cells(radek, 36).Value, "-") > 0 Then  'existuje èíslo zdrojové reklamy RF
            Dim poslRadekRekl   As Long
            Dim radekRekl       As Long
            Dim cisloRekl       As String
            cisloRekl = wsDlrData1.Cells(radek, 36).Value

            'kopie z využité spoleèné reklamy - zmìní se list reklam DLR na 999SR:
            If Left(cisloRekl, 5) = "999SR" Then
                spolReklama = True
                'otevøít soubor SR podle instalace:
                Call OtevritSpolecnouReklamu
                Set wsDlrData2 = ActiveWorkbook.Worksheets(2)   'pøepíše se pùvodní promìnná - z listu DLR na 999SR!
            End If

            'dále stejné pøíkazy pro kopii z využité reklamy DLR i 999SR:
            poslRadekRekl = wsDlrData2.Cells(Rows.Count, 2).End(xlUp).Row   'wsDlrData2 ukazuje podle na DLR nebo SR!!
            'vyhledání zdrojové reklamy:
            For i = 4 To poslRadekRekl
                If cisloRekl = wsDlrData2.Cells(i, 2).Value Then
                    idReklamy = i - 4   'pro zápis využití reklamy

                    'info o pøevzetí reklamy do frm:
                    With .lblZdrojovaReklama
                    .Caption = "Zdrojová reklama: " & wsDlrData2.Cells(i, 2).Value & "."
                    .Visible = True
                    If wsDlrData2.Cells(i, 34).Value = "ano" Then
                        .Caption = .Caption & " Již využito: " & wsDlrData2.Cells(i, 35).Value & "x."
                    Else
                        .Caption = .Caption & " Ještì nevyužito."
                    End If
                    End With    'lblZdrojovaReklama
                    .chkHodn6.Value = True
                    GoTo KonecReklamy
                End If
            Next i

KonecReklamy:
            'kopie z využité spoleèné reklamy - návrat k souboru/listu reklam DLR:
            If spolReklama = True Then
                ActiveWorkbook.Close                            'zavøít 999SR
                Windows(soubor).Activate                        'aktivace souboru DLR
                Set wsDlrData2 = ActiveWorkbook.Worksheets(2)   'návrat promìnné na list reklam DLR!
            End If

        End If  'konec kopie ze schválené reklamy
        .lblUser.Caption = userAp
        End With    'Me

  'EDIT:
  '-----
    ElseIf typZaznamu = "edit" Then
          '===================
        With Me

        radek = ActiveCell.Row

        'šipka pro kopii reg. èísla a èísla žádosti:
        .lblKopieCisel.Visible = True    'zapne se pouze pro EDIT kampanì

        'naèíst øádek do FRM - kampaò:
        '-----------------------------
        .txtDatum.Value = Day(Cells(radek, 1).Value) & "." & Month(Cells(radek, 1).Value) & "." & Year(Cells(radek, 1).Value)
        'pro sledování zmìn mìsíce v záznamu:
        If reklama = False Then
            mesicPuvodni = Month(Cells(radek, 1).Value)
        End If

        .txtKamCislo.Value = Cells(radek, 2).Value   'è. kampanì
        .txtRegCislo.Value = Cells(radek, 3).Value   'reg. èíslo RF
        .txtDlrZadost.Value = Cells(radek, 4).Value  'è. žádosti dlr
        .txtSubFak.Value = Cells(radek, 6).Value     'è. sub. fa
        Call NacistRadek        'naèíst položky spoleèné pro KOPII a EDIT

        If reklama = False Then
            If Cells(radek, 33).Value = "ano" Then .chkKombinace.Value = True   'kombin. èerpání pro KAMPANÌ
        End If
        End With    'Me
        'naèíst øádek do FRM - hodnocení:
        '--------------------------------
        On Error Resume Next    'ignorovat chyby ètení prázdných komentáøù

        'stav:
        With Cells(radek, 22)
        If .Value = "ano" Then
            Me.chkHodn1.Value = True
        Else
            Me.chkHodn1.Value = False
        End If
        'komentáø:
        Me.txtHodn1.Value = .Comment.Text
        End With

        'stav:
        With Cells(radek, 23)
        If .Value = "ano" Then
            Me.chkHodn2.Value = True
        Else
            Me.chkHodn2.Value = False
        End If
        'komentáø:
        Me.txtHodn2.Value = .Comment.Text
        End With

        'stav:
        With Cells(radek, 24)
        If .Value = "ano" Then
            Me.chkHodn3.Value = True
        Else
            Me.chkHodn3.Value = False
        End If
        'komentáø:
        Me.txtHodn3.Value = .Comment.Text
        End With

        'stav:
        With Cells(radek, 25)
        If .Value = "ano" Then
            Me.chkHodn4.Value = True
        Else
            Me.chkHodn4.Value = False
        End If
        'komentáø:
        Me.txtHodn4.Value = .Comment.Text
        End With

        'stav:
        With Cells(radek, 26)
        If .Value = "ano" Then
            Me.chkHodn5.Value = True
        Else
            Me.chkHodn5.Value = False
        End If
        'komentáø:
        Me.txtHodn5.Value = .Comment.Text
        End With

        'stav:
        With Cells(radek, 27)
        If .Value = "ano" Then
            Me.chkHodn6.Value = True
        Else
            Me.chkHodn6.Value = False
        End If
        'komentáø:
        Me.txtHodn6.Value = .Comment.Text
        End With

        With Me

        'naèíst poznámku k uzávìrce:
        If reklama = False Then
            .txtPoznamka.Value = Cells(radek, 35).Comment.Text
        End If

        'naèíst poznámku k reklamì:
        If reklama = True Then
            .txtReklPozn.Value = Cells(radek, 14).Comment.Text
        End If

        'naèíst stav SCHVÁLENO:
        If Cells(radek, 14).Value = "ano" Then
            If reklama = False Then
                .optSchvalit.Value = True
            Else
                .optSchvalitRekl.Value = True
            End If
            If Cells(radek, 35).Value <> "ano" Then _
                MsgBox "Tento záznam byl již SCHVÁLENÝ...      ", vbInformation, aplikace
        ElseIf Cells(radek, 14).Value = "ne" Then
            If reklama = False Then
                .optZamitnout.Value = True
            Else
                .optZamitnoutRekl.Value = True
            End If
            If Cells(radek, 35).Value <> "ano" Then _
                MsgBox "Tento záznam byl již ZAMÍTNUTÝ...      ", vbInformation, aplikace
        End If

        'naèíst stav SCHVÁLIT ADMINISTRÁTOREM - pouze pro kampanì
        'nebo stav využití - pouze pro reklamy:
        If reklama = False Then
            If admin = True Then
                If Cells(radek, 34).Value = "ano" And Cells(radek, 35).Value <> "ano" Then
                    MsgBox "Tento záznam byl již SCHVÁLENÝ ADMINISTRÁTOREM...      ", vbExclamation, aplikace
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

        cisKampane = .txtKamCislo   'è. kampanì - pro editaci+nové pøevzetí reklamy!!

        'info o zdrojové reklamì:
        If reklama = False Then
            If Cells(radek, 36).Value <> "" Then    'existuje èíslo zdroj. reklamy
                With .lblZdrojovaReklama
                .Caption = "Pøevzato z reklamy " & Cells(radek, 36).Value
                .Visible = True
                End With
            End If
        End If
        End With    'Me
    End If
    Me.lblZaznamStatus.Caption = UCase(typZaznamu)

    '----------------------------------
    'konec nast. FRM podle typu záznamu

    With Me
'    'pøehled èerpání - pro EDIT se resetuje odeètení aktuální kampanì až do zmìny èástky:
'    '----------------
'    If typZaznamu = "edit" Then
'        .lblZustP1.Caption = Range("N2").Value
'        .lblZustP2.Caption = Range("N2").Value
'        .lblZustS1.Caption = Range("O2").Value
'        .lblZustS2.Caption = Range("O2").Value
'    End If

    'délka poznámky k uzávìrce:
    .lblPoznDelka.Caption = Len(.txtPoznamka.Value)

    'nápovìdné texty:
    '----------------
    .lblHelp1.Visible = zobrHelp
    .lblHelp2.Visible = zobrHelp
    .lblHelp3.Visible = zobrHelp
    End With    'Me

    Me.txtHodn3.Visible = False 'pøíznak Žádost zaslána - bez textu   od v. 3.0

    'Edit - záznam po uzávìrce jen pro náhled:      od v. 4.3
    '-----------------------------------------
    If typZaznamu = "edit" Then
        'øádek je po uzávìrce:
        If Cells(ActiveCell.Row, 35).Value = "ano" Then
            'vypnout všechny controls:
            Dim control As control
            For Each control In frm_ZAZNAM.Controls
                control.Enabled = False
            Next control
            'nastavit vybrané controls:
            With Me
            .fraCerpani.Visible = False
            .cmdStorno.Enabled = True
            .lblZaznamStatus.Caption = "NÁHLED"
            .lblZaznamStatus.Enabled = True
            End With
            MsgBox "Kampaò " & Cells(ActiveCell.Row, 2) & " je po uzávìrce!      " & vbNewLine _
                & "Pouze administrátor mùže zrušit uzávìrku mìsíce a záznam editovat.      ", vbExclamation, aplikace
            Exit Sub
        End If
    End If

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'zavøení frm køížkem
    Call cmdStorno_Click
End Sub


Private Sub UserForm_Terminate()
    'ukonèení frm
    ciselnikZaznam = False
End Sub



'================= NAÈTENÍ ÈÍSELNÍKÙ ==================

Private Sub NacistCiselnikKamNazev()

    With Me.cmbKamNazev
    .Clear
    cisRadek = ""
    Open cestaMroManager & "\Common_Files\Codelist\k_nazev.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, cisRadek             'ètení jednoho øádku a najetí na další až do EOF=True!!
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
'================= NAÈÍTÁNÍ ÈÍSELNÍKÙ - KONEC ==================



Private Sub NajitPenetraci()
    'pouze ŠkodaPlus
    'podle NW najít penetraci a pøiøadit % pøíspìvku
    'zobrazit penetraci ve frm
    'voláno z frm_init

    With Me
    For i = LBound(dlrNw) To UBound(dlrNw)
        If dlrNw(i) = Range("A1").Value Then
            .txtKamProcento.Value = dlrSpPrispevek(i)
            With .lblPenetrace
            .Visible = True
            .Caption = "Penetrace ŠkoFin " & dlrSpPenetrace(i) & "%"
            End With
            Exit Sub
        End If
    Next i
    End With    'Me

End Sub



Private Sub ZaznamRezimReklama()
    'pøepnout do režimu schvalování reklam

Dim barva       As Long
Dim posuv       As Integer      'posuv objektù nahoru v režimu reklamy

    barva = RGB(236, 130, 4)    'RF
    posuv = 72

    'úprava frm:
    With Me
    .StartUpPosition = 0
    .Width = 343
    .Height = 464
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)

    'záhlaví rámeèku KAMPAÒ:
    With .lblTitul1
    .ForeColor = barva
    .Width = 96
    .Caption = "Schválení reklamy:"
    End With
    With .lblIntCisloKam
    .ForeColor = barva
    .Caption = "Interní èíslo reklamy:"
'        .Left = 150
    End With
    With .txtKamCislo
    .ForeColor = barva
    .Left = 268
    End With

    'rámeèek Kampaò:
    .lblPrecislovatFakturu.Visible = False

    .lblDatum.Caption = "Datum žádosti:"

    .lblRegCislo.Enabled = False
    .txtRegCislo.Enabled = False

    .lblDlrZadost.Enabled = False
    .txtDlrZadost.Enabled = False

    .lblSubFak.Enabled = False
    .txtSubFak.Enabled = False

    .chkKombinace.Visible = False
    .lblKombinace.Visible = False

    .lblKamNazev.Caption = "Název reklamy:"

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
    .lblKamPocetReklamDok.Visible = True    'nápovìda k poètu mat.
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

    'stavové hlášky:
    With .lblZaznamStatus
    .BorderColor = barva
    .ForeColor = barva
    End With
    .lblAlertKombinace.BackColor = barva
    .lblAlertSpolecna.BackColor = barva
    .lblAlertBonus.BackColor = barva

    'tlaèítka:
    With .cmdUlozit
        .Top = 406
        .Left = 266
'        .Enabled = False        'VYPNUTO, NEŽ SE DOPROGRAMUJE ZÁPIS!!!!
    End With
    With .cmdStorno
        .Top = 406
        .Left = 196
    End With

    End With    'Me

End Sub



Private Sub NacistRadek()
    'volaná sub - naèíst øádek z tabulky do FRM
    'naètou se spoleèné položky pro KOPII a EDIT záznamu
    'voláno z FRM INIT

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

    If Cells(radek, 32).Value = "ano" Then          'spoleèná reklama ano/ne
        .chkSpolReklama.Value = True
    Else
        .chkSpolReklama.Value = False
    End If

    If Cells(radek, 31).Value = "ano" Then          'kam. vlastní/import
        .optKamVlastni.Value = True
    Else
        .optKamImport.Value = True
    End If
    End With    'Me

    With Cells(radek, 8)
        If .Value = "prodej" Then                   'oblast pøíspìvku
            Me.optProdej.Value = True
        ElseIf .Value = "servis" Then
            Me.optServis.Value = True
        Else
            MsgBox "Není definována oblast pøíspìvku prodej/servis...      ", vbExclamation, aplikace
        End If
    End With

    With Me
    .cmbKamNazev.Value = Cells(radek, 7).Value       'název kam.

'    .txtKamCena.Value = Cells(radek, 9).Value        'cena celkem
    .txtKamCena.Value = Application.WorksheetFunction.Substitute(Cells(radek, 9).Value, ".", ",")       'cena celkem (konverze teèky na èárku)
    .txtKamProcento.Value = Cells(radek, 10).Value    '% pøíspìvku
    .txtKamSchval.Value = Cells(radek, 11).Value      'schválená èástka
    .txtKamNeschval.Value = Application.WorksheetFunction.Substitute(Cells(radek, 12).Value, ".", ",")   'neschválená èástka

    .cmbKamTyp.Value = Cells(radek, 15).Value        'typ kam.
    .cmbKamZamereni.Value = Cells(radek, 16).Value   'zamìøení kam.
    .cmbKamMedium.Value = Cells(radek, 17).Value     'medium-typ
    .cmbKamMediumNazev.Value = Cells(radek, 18).Value     'medium-název
    .cmbKamZdroj.Value = Cells(radek, 19).Value      'zdroj dat
'    If Cells(radek, 19).Value = "PP" Then                 'PP-DAS
'        If Cells(radek, 19).Comment.Text = "ano" Then
'            .chkDas.Value = True
'        Else
'            .chkDas.Value = False
'        End If
'    End If

    .cmbKamFormat.Value = Cells(radek, 20).Value     'formát
    .txtKamPocet.Value = Cells(radek, 21).Value      'poèet ks mat.
    End With    'Me

    appEvents = True

End Sub



Private Sub GenerovatCisloKampane()
    'volaná sub - generuje èíslo kampanì pro Nový záznam a Kopii záznamu
    'voláno z FRM INIT

    poslRadek = Cells(Rows.Count, 2).End(xlUp).Row              'podle sloupce è. kampanì (datum mùže být prázdné)
    If poslRadek = 4 And Cells(poslRadek, 1).Value = "" Then    'tabulka je zatím prázdná
        If reklama = False Then
            cisKampane = Range("A1").Value & "-" & "001"            'první èíslo kampanì pro žádosti
        Else
            cisKampane = Range("A1").Value & "-" & "9001"           'první èíslo kampanì pro reklamy
        End If
    Else
        'generovat nové poø. èíslo kampanì:
        If reklama = False Then     'pro žádosti
        
            cisKampane = CInt(Right(Cells(poslRadek, 2).Value, 3))   'koncové èíslo pøedchozí kampanì
            cisKampane = CStr(cisKampane + 1)
            'doplnit na 3 místné èíslo:
            If Len(cisKampane) = 1 Then
                cisKampane = "00" & cisKampane
            ElseIf Len(cisKampane) = 2 Then
                cisKampane = "0" & cisKampane
            End If
            
        Else                        'pro reklamy
        
            cisKampane = CInt(Right(Cells(poslRadek, 2).Value, 4))   'koncové èíslo pøedchozí kampanì
            cisKampane = CStr(cisKampane + 1)
            
        End If
        cisKampane = Range("A1").Value & "-" & cisKampane   'sestavit èíslo kampanì DLR
        
        'korekce èísla rekl. - z ext. instalace jsou èíslované xxxxx-8xxx:
        If Mid(cisKampane, 7, 1) = "8" Then _
            cisKampane = Application.WorksheetFunction.Substitute(cisKampane, "-8", "-7")
    End If

    Me.txtKamCislo.Value = cisKampane

End Sub



Private Sub lblDnes_Click()
    'do pole Datum generovat aktuální datum (a skoèit na další prvek)

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
    'volaná sub - generovat datum kampanì do FRM

    'zatím aktuální datum
    Me.txtDatum.Value = Day(Date) & "." & Month(Date) & "." & Year(Date)

End Sub



Private Sub txtDatum_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'odchod z pole datum = generovat èíslo žádosti
    Call GenerovatCisloZadosti

End Sub



Private Sub GenerovatCisloZadosti()
    'generování è. žádosti z pole Datum
    'voláno z txtDatum_Exit a cmdUlozit_Click

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
    'odchod z pole = kontrola èísla subdod. fa na již existující

Dim faktura As String
Dim cena    As String
Dim prisp   As String
Dim pozn1   As String
Dim pozn2   As String

    If appEvents = True Then    'mùže být False pøi ukládání záznamu - ošetøena chyba s kopií záznamu (ZATIM VYPNUTO)
    
        faktura = Me.txtSubFak.Value
        poslRadek = Cells(Rows.Count, 2).End(xlUp).Row
        If Not (poslRadek = 4 And Cells(poslRadek, 2).Value = "") Then    'tabulka není prázdná
            For i = 4 To poslRadek
                If faktura = Cells(i, 6).Value Then     'nalezeno duplicitní èíslo fa
                
                    cena = Application.WorksheetFunction.Fixed(Cells(i, 9).Value, 2)
                    cena = Application.WorksheetFunction.Substitute(cena, ",", " ")
                    cena = Application.WorksheetFunction.Substitute(cena, ".", ",")
                    
                    prisp = Application.WorksheetFunction.Fixed(Cells(i, 11).Value, 2)
                    prisp = Application.WorksheetFunction.Substitute(prisp, ",", " ")
                    prisp = Application.WorksheetFunction.Substitute(prisp, ".", ",")
                    
                    On Error Resume Next    'ignorovat chyby ètení prázdných komentáøù
                    pozn1 = Cells(i, 22).Comment.Text
                    pozn2 = Cells(i, 23).Comment.Text
    
                    MsgBox "Byla nalezena kampaò s èíslem subdod. fa " & faktura & ":" & vbNewLine & vbNewLine _
                        & "È. kampanì:  " & Cells(i, 2).Value & vbNewLine _
                        & "Datum:           " & Cells(i, 1).Value & vbNewLine _
                        & "Kampaò:        " & Cells(i, 7).Value & vbNewLine _
                        & "Cena:              " & cena & vbNewLine _
                        & "Pøíspìvek:      " & prisp & vbNewLine _
                        & "Schváleno:     " & UCase(Cells(i, 14).Value) & vbNewLine _
                        & vbNewLine _
                        & "Pozn. 1:          " & pozn1 & vbNewLine _
                        & "Pozn. 2:          " & pozn2 & vbNewLine _
                        , vbExclamation, "Duplicitní èíslo fa dodavatele!"
                End If
            Next i
        End If
    
    End If  'appEvents

End Sub



Private Sub chkBonus_Change()
    'zmìna chk Bonus akce

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
            Call NajitPenetraci     'podle NW najít penetraci a pøiøadit % pøíspìvku
        End If
        'reset zdroj. bonus:
        idBonusy = -100
        .lblZdrojovyBonus.Visible = False
        .lblZdrojovyBonus.Caption = "Pøevzato: "
    End If
    End With
    Call VypocetCeny

End Sub



Private Sub cmdPrevzitBonus_Click()
    'frm pøevzít bonus

    frmBonusy.Show

End Sub



Private Sub chkSpolReklama_Change()
    'zmìna chk Spoleèná reklama

    Me.lblAlertSpolecna.Visible = Me.chkSpolReklama.Value

End Sub



Private Sub optKamImport_Change()
    'option Kampaò importéra

'    If optKamImport = True Then
'
'        'nastavit seznam na kampanì importéra:
'        With cmbKamNazev
'            .Clear
'            For i = LBound(cisKamImport) To UBound(cisKamImport)
'                .AddItem cisKamImport(i)
'            Next i
'            .ShowDropButtonWhen = fmShowDropButtonWhenAlways    'kam. importéra nabízí rozbalení seznamu
'        End With
'
'    Else
'
'        'nastavit seznam na kampanì vlastní:
'        With cmbKamNazev
'            .Clear
'            .ShowDropButtonWhen = fmShowDropButtonWhenNever     'vlastní kam. - pouze zapsat
'        End With
'
'    End If

End Sub



Private Sub optKamVlastni_Change()
    Call optKamImport_Change
    '(optiony se pøepínají)
End Sub



Private Sub optProdej_Change()
    'zmìna prodej/servis
    Call VypocetCeny    'pro pøepoèet zùstatkù
End Sub



Private Sub optServis_Change()
    'zmìna prodej/servis
    Call VypocetCeny    'pro pøepoèet zùstatkù
End Sub



Private Sub chkKombinace_Change()

    With Me
    .lblAlertKombinace.Visible = .chkKombinace.Value

'PRO ZÁZNAM KAMPANÍ:
'-------------------
'    If .chkKombinace.Value = True Then
    If .chkKombinace.Value = True And typZaznamu = "kopie" Then
        'zapnout kopii z jedné kampanì - zkopírovat vybrané hodnoty z aktuálního øádku:
        If radek = poslRadek Then .txtKamCislo.Value = Cells(radek, 2).Value     'int. èíslo kam. - pouze pro kopii z posledního øádku!!
        .txtDatum.Value = Cells(radek, 1).Value       'datum
        .txtDatum.Value = Day(Cells(radek, 1).Value) & "." & Month(Cells(radek, 1).Value) & "." & Year(Cells(radek, 1).Value)

        .txtDlrZadost.Value = Cells(radek, 4).Value      'è. žádosti
        .txtRegCislo.Value = Cells(radek, 3).Value    'reg. èíslo RF
        .txtSubFak.Value = Cells(radek, 6).Value          'è. subdod. fa

        If Cells(radek, 8).Value = "prodej" Then             'první øádek je prodej (zapnout 2. možnost)
            .optServis.Value = True
        ElseIf Cells(radek, 8).Value = "servis" Then
            .optProdej.Value = True
        End If

        With .txtKamCena
            .Value = 0
            .SetFocus
        End With

        'smazat zadané položky - za P/S se liší:
        .cmbKamTyp.Value = ""
        .cmbKamZamereni.Value = ""

    End If
    End With    'Me

End Sub



Private Sub cmbKamNazev_Change()
    'název kampanì mùže generovat typ/zamìøení kampanì

Dim kamNazev As String

    'od v. 5.2.1 - pøi zmìnì názvu kampanì zùstavaly pùvodní parametry záznamu
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
            .cmbKamTyp.Value = "Nové vozy"
            .cmbKamZamereni.Value = "Citigo"

        ElseIf InStr(kamNazev, "Testovací jízdy Fabia 2015") > 0 Then
            .cmbKamTyp.Value = "Nové vozy"
            .cmbKamZamereni.Value = "Fabia"

        ElseIf InStr(kamNazev, "Operativní leasing - Bez starostí") > 0 Then
                                '----------------------
            .cmbKamTyp.Value = "Nové vozy"
            .cmbKamZamereni.Value = "Více modelù vozù"

        ElseIf InStr(kamNazev, "Centrální regionální kampaò 2015") > 0 Or _
            InStr(kamNazev, "Centrální kampaò - Hokejová extraliga 2014/15") > 0 Then
                                '----------------------
            If typZaznamu <> "edit" Then .chkBonus.Value = True
            .cmbKamTyp.Value = "Nové vozy"
            .cmbKamZamereni.Value = "Více modelù vozù"
            .cmbKamMedium.Value = "Akce - prezentace"
            .cmbKamMediumNazev.Value = "Centrální kampaò ŠA"
            .cmbKamZdroj.Value = "B2B"
            .cmbKamFormat.Value = "Vlastní"

        ElseIf InStr(kamNazev, "Centrální kampaò - Car Configurator (PCC)") > 0 Then
                                '----------------------
            If typZaznamu <> "edit" Then .chkBonus.Value = True
            .cmbKamTyp.Value = "Nové vozy"
            .cmbKamZamereni.Value = "Více modelù vozù"
            .cmbKamMedium.Value = "TV LED"
            .cmbKamMediumNazev.Value = "Centrální kampaò ŠA"
            .cmbKamZdroj.Value = "B2B"
            .cmbKamFormat.Value = "Data"

        ElseIf InStr(kamNazev, "ŠKODA Pojištìní") > 0 Then
                                '----------------------
            .cmbKamTyp.Value = "Prezentace spoleènosti"
            .cmbKamZamereni.Value = "Prezentace spoleènosti"
            .cmbKamMedium.Value = "Tisk"

        ElseIf InStr(kamNazev, "Fabia NOVÁ Combi") > 0 Then
            .cmbKamTyp.Value = "Nové vozy"
            .cmbKamZamereni.Value = "Fabia Combi"

        ElseIf InStr(kamNazev, "Fabia") > 0 Then
            .cmbKamTyp.Value = "Nové vozy"
            .cmbKamZamereni.Value = "Fabia"

        ElseIf InStr(kamNazev, "Octavia Combi") > 0 Then
            .cmbKamTyp.Value = "Nové vozy"
            .cmbKamZamereni.Value = "Octavia Combi"

        ElseIf InStr(kamNazev, "Octavia") > 0 Then
            .cmbKamTyp.Value = "Nové vozy"
            .cmbKamZamereni.Value = "Octavia"

        ElseIf InStr(kamNazev, "Roomster") > 0 Then
            .cmbKamTyp.Value = "Nové vozy"
            .cmbKamZamereni.Value = "Roomster"

        ElseIf InStr(kamNazev, "Superb Combi") > 0 Then
            .cmbKamTyp.Value = "Nové vozy"
            .cmbKamZamereni.Value = "Superb Combi"

        ElseIf InStr(kamNazev, "Superb") > 0 Then
            .cmbKamTyp.Value = "Nové vozy"
            .cmbKamZamereni.Value = "Superb"

        ElseIf InStr(kamNazev, "Rapid Spaceback") > 0 Then
            .cmbKamTyp.Value = "Nové vozy"
            .cmbKamZamereni.Value = "Rapid Spaceback"

        ElseIf InStr(kamNazev, "Rapid") > 0 Then
            .cmbKamTyp.Value = "Nové vozy"
            .cmbKamZamereni.Value = "Rapid"

        ElseIf InStr(kamNazev, "Yeti") > 0 Then
            .cmbKamTyp.Value = "Nové vozy"
            .cmbKamZamereni.Value = "Yeti"

        ElseIf InStr(kamNazev, "servisní akce") > 0 Then
            .cmbKamTyp.Value = "Servis"
            .cmbKamZamereni.Value = "Sezónní servisní akce"
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
    'oblast èerpání podle typu kampanì

    With Me
    If .cmbKamTyp.Value = "Nové vozy" Or .cmbKamTyp.Value = "ŠKODA Plus" Then
        .optProdej.Value = True
    End If
    End With

End Sub



Private Sub cmbKamMediumNazev_Change()
    'pro média "distribuce" nastavit zdroj dat "Vlastní"

    With Me
    If InStr(.cmbKamMediumNazev.Value, "distribuce") > 0 Then
        .cmbKamZdroj.Value = "Vlastní"
    End If
    End With

End Sub



Private Sub txtKamCena_Enter()
    'vstup na cenu kampanì - kontrola oblasti P/S (nutné pro pøepoèet zùstatkù ve FRM)

    With Me
    If .optProdej.Value = False And .optServis.Value = False Then
        MsgBox "Nejdøíve vyberte oblast èerpání (prodej/servis).      ", vbExclamation, aplikace
        .optProdej.SetFocus
    End If
    End With    'Me

End Sub



Private Sub lblKopieCisel_Click()
    'pouze pro EDIT - kopie reg. èísla a è. žádosti z minulého øádku

    With Me
    .txtRegCislo.Value = Cells(radek - 1, 3).Value
    .txtDlrZadost.Value = Cells(radek - 1, 4).Value
    End With    'Me

End Sub



Private Sub lblPrecislovatFakturu_Click()
    'POUZE PRO KAMPANÌ
    'zmìna èísla faktury pro všechny øádky žádosti

Dim stareCislo  As String
Dim noveCislo   As String

    stareCislo = Me.txtDlrZadost
    noveCislo = InputBox("Zamìnit èíslo fa " & stareCislo & " za: ", _
        "Zmìna èíslo fa v žádosti è. " & stareCislo, stareCislo)
    If noveCislo = "" Then Exit Sub
    poslRadek = Cells(Rows.Count, 2).End(xlUp).Row              'podle sloupce è. kampanì (datum mùže být prázdné)

    'projít všechny øádky tabulky:
    '-----------------------------
    ActiveSheet.Unprotect Password:=heslo
    For i = 4 To poslRadek
        Application.StatusBar = "Kontroluji záznam " & i - 3 & "/" & poslRadek - 3 & "..."
        If Cells(i, 4).Value = stareCislo Then Cells(i, 4).Value = noveCislo
    Next i
    Application.StatusBar = False
    ActiveSheet.Protect Password:=heslo
    ActiveWorkbook.Save
    Me.txtDlrZadost.Value = noveCislo
    MsgBox "Èíslo faktury DLR bylo zmìnìno.      ", vbInformation, ""

End Sub



'Private Sub cmbKamZdroj_Change()
'    'výbìr PP povolí volbu DAS
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



'========== VÝPOÈET PØÍSPÌVKÙ: ==========

Private Sub txtKamCena_Change()
    'zmìna celkové ceny
    Call VypocetCeny
End Sub


Private Sub txtKamProcento_Change()
    'zmìna procenta pøíspìvku
    Call VypocetCeny
End Sub


Private Sub txtKamNeschval_Change()
    'neschválená èástka
    Call VypocetCeny
End Sub



Private Sub VypocetCeny()
    'volaná sub - výpoèet schválené èástky pøi zmìnì ceny nebo procenta
    'pøidán pøepoèet zùstatkù ve FRM
Dim prispevek   As Single

    With Me
    On Error Resume Next

    'výpoèet pøíspìvku a zaokrouhlení:
    prispevek = Round((.txtKamCena.Value - .txtKamNeschval.Value) * .txtKamProcento.Value / 100, 2)
    .txtKamSchval.Value = prispevek
    'konverze teèky na èárku - pouze pro zobrazení:
    .txtKamSchval.Value = Application.WorksheetFunction.Substitute(.txtKamSchval.Value, ".", ",")


'        'návrh doèerpání rozpoètu:
'        MsgBox "Rozpoèet je pøeèerpán o " & -1 * precerpani & " Kè," & vbNewLine _
'            & "zapisuji neschválenou èástku " & -1 * neschvaleno & " Kè.      ", _
'            vbExclamation, "Pøeèerpání rozpoètu"
'        .txtKamNeschval.Value = -1 * neschvaleno



    'kontrola èerpání/zùstatku:
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

        'výpoèet zùstatku pøed kampaní:
        '------------------------------
        If dlrCinnost(dlr) = "prodej" Then
            If typZaznamu <> "edit" Then   'pro nové záznamy prodej
                .lblZustP1.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetProdej(dlr) - cerpaniP, 2)
                .lblZustS1.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetServis(dlr) - cerpaniS, 2)
            Else    'pøi editaci pøièíst aktuální øádek za prodej/servis!
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
            If typZaznamu <> "edit" Then   'pro nové záznamy prodej
                .lblZustP1.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetFinal(dlr) - cerpaniP - cerpaniS, 2)
            Else
                .lblZustP1.Caption = Application.WorksheetFunction.Fixed(dlrRozpocetFinal(dlr) - cerpaniP - cerpaniS + Cells(ActiveCell.Row, 11).Value, 2)
            End If
            .lblZustP1.Caption = Application.WorksheetFunction.Substitute(.lblZustP1.Caption, ",", " ")
            .lblZustP1.Caption = Application.WorksheetFunction.Substitute(.lblZustP1.Caption, ".", ",")
    
            .lblZustS1.Caption = ""
    
        End If
    
        'výpoèet zùstatku po kampani:
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
                    
                Else    'pøi editaci pøièíst aktuální øádek!
                
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
                'konverze formátu:
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
                    
                Else    'pøi editaci pøièíst aktuální øádek!
                
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
                'konverze formátu:
                .lblZustS2.Caption = Application.WorksheetFunction.Substitute(.lblZustS2.Caption, ",", " ")
                .lblZustS2.Caption = Application.WorksheetFunction.Substitute(.lblZustS2.Caption, ".", ",")
                
            Else    'nový záznam - stav pøed výbìrem obl. èerpání
                .lblZustP2.Caption = ""
                .lblZustS2.Caption = ""
            End If
            
        Else    'èinnost servis/Plus
            
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
        
        .lblAlert.Visible = False   '3.2.3 opravena chyba - zapnutí Bonusu nevypnulo pøeèerpání
        .cmdUlozit.Enabled = True

    End If

    End With    'Me

End Sub
'========== VÝPOÈET PØÍSPÌVKÙ - KONEC ==========



Private Sub cmdZustKontrola_Click()
    'tl. pøepoèet zùstatkù po kampani

Dim precerpani  As Long
Dim neschvaleno As Long

    'refresh zùstatkù v info rámeèku a indikace pøeèerpání: - POUZE PRO NEBONUSOVÉ!!!!!
    '------------------------------------------------------
    With Me
    If .chkBonus = False Then
        .lblZustP2.Caption = .lblZustP1.Caption
        .lblZustS2.Caption = .lblZustS1.Caption

        'pro PRODEJ je kontrola rozdìlena na prodej/servis:
        If Range("G1").Value = "Èinnost: prodej" Then
            If .optProdej.Value = True Then
    
                precerpani = Application.WorksheetFunction.RoundUp((CSng(.lblZustP1.Caption)) - (CSng(.txtKamSchval.Value)), 0)
                neschvaleno = precerpani / .txtKamProcento.Value * 100
                .lblZustP2.Caption = precerpani
                If precerpani < 0 Then              'pøeèerpání prodej
                    Call ZobrazitPrecerpani
                    'návrh doèerpání rozpoètu:
                    MsgBox "Prodej je pøeèerpán o " & -1 * precerpani & " Kè," & vbNewLine _
                        & "zapisuji neschválenou èástku " & -1 * neschvaleno & " Kè.      ", _
                        vbExclamation, "Pøeèerpání rozpoètu"
                    .txtKamNeschval.Value = -1 * neschvaleno
                Else
                    .lblPrecerpani.Visible = False
                End If
    
            ElseIf .optServis.Value = True Then
    
                precerpani = Application.WorksheetFunction.RoundUp((CSng(.lblZustS1.Caption)) - (CSng(.txtKamSchval.Value)), 0)
                neschvaleno = precerpani / .txtKamProcento.Value * 100
                .lblZustS2.Caption = precerpani
                If precerpani < 0 Then              'pøeèerpání servis
                    Call ZobrazitPrecerpani
                    'návrh doèerpání rozpoètu:
                    MsgBox "Servis je pøeèerpán o " & -1 * precerpani & " Kè," & vbNewLine _
                        & "zapisuji neschválenou èástku " & -1 * neschvaleno & " Kè.      ", _
                        vbExclamation, "Pøeèerpání rozpoètu"
                    .txtKamNeschval.Value = -1 * neschvaleno
                Else
                    .lblPrecerpani.Visible = False
                End If
    
            End If
            
        'pro SERVIS / PLUS / RETAIL je použito jednoduché zobrazení - pouze ø. 1:
        Else
            precerpani = Application.WorksheetFunction.RoundUp((CSng(.lblZustP1.Caption)) - (CSng(.txtKamSchval.Value)), 0)
            neschvaleno = precerpani / .txtKamProcento.Value * 100
            .lblZustP2.Caption = precerpani
            If precerpani < 0 Then              'pøeèerpání MRO
                Call ZobrazitPrecerpani
                'návrh doèerpání rozpoètu:
                MsgBox "Rozpoèet je pøeèerpán o " & -1 * precerpani & " Kè," & vbNewLine _
                    & "zapisuji neschválenou èástku " & -1 * neschvaleno & " Kè.      ", _
                    vbExclamation, "Pøeèerpání rozpoètu"
                .txtKamNeschval.Value = -1 * neschvaleno
            Else
                .lblPrecerpani.Visible = False
            End If
        End If

    Else    'pro Bonus akce se nepøepoèítá pøehled èerpání:
        .lblZustP1.Caption = Range("L3").Value
        .lblZustP2.Caption = Range("L3").Value
        .lblZustS1.Caption = Range("O3").Value
        .lblZustS2.Caption = Range("O3").Value
        .lblPrecerpani.Visible = False
    End If

    If Range("D2").Value <> "Èinnost: prodej" Then
        .lblZustP2.Caption = Application.WorksheetFunction.Round((CSng(.lblZustP1.Caption)) - (CSng(.txtKamSchval.Value)), 1)
    End If
    End With    'Me

End Sub



Private Sub ZobrazitPrecerpani()
    'zobrazit alert pøeèerpání, voláno z cmdZustKontrola_Click

    With Me.lblPrecerpani
    If Range("G1").Value <> "Èinnost: prodej" Then
        .Caption = "POZOR! Pøeèerpání rozpoètu MRO!"
    End If
    .Top = 312
'    .Left = 309
    .Visible = True    'alert
    End With

End Sub



'========== HODNOCENÍ: ==========

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
    'kopírovat pozn. do poznámky k uzávìrce
    With Me
    If Len(.txtHodn1.Value) > 0 Then .txtPoznamka.Value = .txtPoznamka.Value & " " & .txtHodn1.Value
    End With    'Me
End Sub

Private Sub lblPozn2_Click()
    'kopírovat pozn. do poznámky k uzávìrce
    With Me
    If Len(.txtHodn2.Value) > 0 Then .txtPoznamka.Value = .txtPoznamka.Value & " " & .txtHodn2.Value
    End With    'Me
End Sub

Private Sub lblPozn3_Click()
    'kopírovat pozn. do poznámky k uzávìrce
    With Me
    If Len(.txtHodn3.Value) > 0 Then .txtPoznamka.Value = .txtPoznamka.Value & " " & .txtHodn3.Value
    End With    'Me
End Sub

Private Sub lblPozn4_Click()
    'kopírovat pozn. do poznámky k uzávìrce
    With Me
    If Len(.txtHodn4.Value) > 0 Then .txtPoznamka.Value = .txtPoznamka.Value & " " & .txtHodn4.Value
    End With    'Me
End Sub

Private Sub lblPozn5_Click()
    'kopírovat pozn. do poznámky k uzávìrce
    With Me
    If Len(.txtHodn5.Value) > 0 Then .txtPoznamka.Value = .txtPoznamka.Value & " " & .txtHodn5.Value
    End With    'Me
End Sub

Private Sub lblPozn6_Click()
    'kopírovat pozn. do poznámky k uzávìrce
    With Me
    If Len(.txtHodn6.Value) > 0 Then .txtPoznamka.Value = .txtPoznamka.Value & " " & .txtHodn6.Value
    End With    'Me
End Sub



Private Sub txtPoznamka_Change()
    'zmìna délky poznámky k uzávìrce

Dim limit       As Integer

    limit = 55

    With Me
    .lblPoznDelka.Caption = limit - Len(.txtPoznamka.Value)
    'omezení délky:
    If Len(.txtPoznamka.Value) > limit Then .txtPoznamka.Value = Left(.txtPoznamka.Value, limit)
    End With    'Me

End Sub



Private Sub lblCenaRovno_Click()
    'autom. text do pozn. k uzávìrce
    With Me
    .txtPoznamka.Value = "cena dle ceníku vydavatele" & .txtPoznamka.Value
    End With    'Me
End Sub

Private Sub lblCenaNizsi_Click()
    'autom. text do pozn. k uzávìrce
    With Me
    .txtPoznamka.Value = "cena nižší než ceníková" & .txtPoznamka.Value
    End With    'Me
End Sub



Private Sub txtReklPozn_Change()
    'zmìna délky poznámky k reklamì

Dim limit       As Integer

    limit = 255

    With Me
    .lblReklDelka.Caption = limit - Len(.txtReklPozn.Value)
    'omezení délky:
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
    'schválení/zamítnutí teprve zapne schválení adminem!
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
    On Error Resume Next    'pøi pøeèerpání se vypne tlaèítko a nejde focus!!
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
'========== HODNOCENÍ - KONEC ==========



Private Sub cmdUlozit_Click()
    'tl. Uložit (záznam)

Dim mesicDokl As Integer    'è. mìsíce pro evidenci dokladù

    On Error GoTo Chyba

    'logická kontrola dat:
    '=====================

    'kontroly pouze pro žádosti:
    '---------------------------
    With Me
    If reklama = False Then

'        'reg. èíslo RF:
'        If .txtRegCislo.Value = "" Then
'            MsgBox "Není vyplnìno reg. èíslo RF!      ", vbExclamation, "Chybí povinný údaj"
'            .txtRegCislo.SetFocus
'            Exit Sub
'        End If

        'èíslo žádosti:
        If .txtDlrZadost.Value = "" Then
            MsgBox "Není vyplnìno èíslo žádosti!      ", vbExclamation, "Chybí povinný údaj"
            .txtDlrZadost.SetFocus
            Exit Sub
        End If

        'celkem Kè:
        If .txtKamCena.Value = "" Then
            MsgBox "Není vyplnìna cena kampanì!      ", vbExclamation, "Chybí povinný údaj"
            .txtKamCena.SetFocus
            Exit Sub
        End If

        'pøevzatý bonus:
        If .chkBonus.Value = True And typBonus = "" Then
            If MsgBox("Bonus není pøevzatý ze schválených.     " _
            & vbNewLine & vbNewLine & "Je to v poøádku?", vbYesNo + vbExclamation, "") = vbNo Then
                .cmdPrevzitBonus.SetFocus
                Exit Sub
            End If
        End If

    End If

    'kontroly spoleèné pro reklamy/žádosti:
    '--------------------------------------
    'datum - rok:
    If Year(CDate(.txtDatum.Value)) <> rok Then
        If MsgBox("Datum záznamu nepatøí do roku " & rok & "!      " _
            & vbNewLine & vbNewLine & "Je to v poøádku?", vbYesNo + vbExclamation, "") = vbNo Then
            .txtDatum.SetFocus
            Exit Sub
        End If
    End If

    'interní è. kampanì:
    If .txtKamCislo.Value = "" Then
        MsgBox "Chybí interní èíslo kampanì!      ", vbExclamation, "Chybí povinný údaj"
        .txtKamCislo.SetFocus
        Exit Sub
    End If

    'oblast prodej/servis:
    If .optProdej.Value = False And .optServis.Value = False Then
        MsgBox "Zadejte oblast pøíspìvku prodej/servis!      ", vbExclamation, "Chybí povinný údaj"
        Exit Sub
    End If

    'název kampanì:
    If .cmbKamNazev.Value = "" Then
        MsgBox "Není vyplnìn název kampanì!      ", vbExclamation, "Chybí povinný údaj"
        .cmbKamNazev.SetFocus
        Exit Sub
    End If

    'typ kampanì:
    If .cmbKamTyp.Value = "" Then
        MsgBox "Není vyplnìn typ kampanì!      ", vbExclamation, "Chybí povinný údaj"
        .cmbKamTyp.SetFocus
        Exit Sub
    End If

    'zamìøení kampanì:
    If .cmbKamZamereni.Value = "" Then
        MsgBox "Není vyplnìno zamìøení kampanì!      ", vbExclamation, "Chybí povinný údaj"
        .cmbKamZamereni.SetFocus
        Exit Sub
    End If

    'zdroj dat:
    If .cmbKamZdroj.Value = "" Then
        MsgBox "Není vyplnìn zdroj dat!      ", vbExclamation, "Chybí povinný údaj"
        .cmbKamZdroj.SetFocus
        Exit Sub
    End If

    'formát:
'    If .cmbKamFormat.Value = "" And .cmbKamZdroj.Value = "DAS" Then
'        MsgBox "Není vyplnìn formát!      ", vbExclamation, "Chybí povinný údaj"
'        .cmbKamFormat.SetFocus
'        Exit Sub
'    End If
    If .cmbKamFormat.Value = "" Then
        MsgBox "Není vyplnìn formát!      ", vbExclamation, "Chybí povinný údaj"
        .cmbKamFormat.SetFocus
        Exit Sub
    End If

    'kontrola oblasti èerpání:
    If (InStr(UCase(.cmbKamTyp.Value), "SERVIS") > 0 Or InStr(UCase(.cmbKamZamereni.Value), "SERVIS") > 0) And .optProdej.Value = True Then
        If MsgBox("Zkontrolujte Oblast èerpání a Typ/Zamìøení kampanì!   " & vbNewLine & vbNewLine & "Je to v poøádku?", _
        vbExclamation + vbYesNo + vbDefaultButton2, "Logická kontrola dat") = vbNo Then Exit Sub
    End If

    'kontroly pouze pro reklamy:
    '---------------------------
    
'    If reklama = True Then
'        If .txtKamPocet.Value > 10 Then
'        If MsgBox("Poèet schválených rekl. dokumentù je pøíliš vysoký.      " _
'            & vbNewLine & vbNewLine _
'            & "Je to v poøádku?" _
'            & vbNewLine & vbNewLine _
'            & "(Zapisuje se poèet schválených dokumentù, ne poèet vyrobených!)", vbYesNo + vbExclamation, "Ukládání rekl. dokumentu") = vbNo Then Exit Sub
'        End If
'    End If

    If reklama = True Then
        If .txtKamPocet.Value > 10 Then
        If MsgBox("Poèet schválených rekl. dokumentù je pøíliš vysoký.      " _
            & vbNewLine & vbNewLine _
            & "Zapište poèet schválených dokumentù, ne poèet vyrobených!", vbOKOnly + vbCritical, "Ukládání rekl. dokumentu") = vbNo Then Exit Sub
        End If
    End If

    End With    'Me
    '=======================
    'konec kontroly dat

    'podle NW vyhledat DLR a získat index: (pro využití bonusù a zápis do evidence dokladù)
    For i = LBound(dlrNw) To UBound(dlrNw)
        If Range("A1").Value = dlrNw(i) Then
            dlr = i
            GoTo KonecNw
        End If
    Next i
KonecNw:

    'odemknout list pro zápis:
    ActiveSheet.Unprotect Password:=heslo

    'výbìr øádku pro zápis kampanì/reklamy:
    If typZaznamu <> "edit" Then
        'pro nový záznam a kopii:
        poslRadek = Cells(Rows.Count, 2).End(xlUp).Row
        If poslRadek = 4 And Cells(poslRadek, 2).Value = "" Then    'tabulka je zatím prázdná
            radek = 4
        Else
            radek = poslRadek + 1
        End If
    Else
        'pro EDIT je øádek z ActiveCell
        radek = ActiveCell.Row
    End If

    'zápis položek pouze pro kampanì (ostatní jsou spoleèné pro kampanì i reklamy):
    '===============================
    If reklama = False Then
        With Cells(radek, 3)                                'reg. è. RF
            .NumberFormat = "@"
            .Value = CStr(txtRegCislo.Value)
        End With

        'kontrola è. žádosti pøed uložením - od v. 5.0
        '(zmìna data v editaci se neprojevila ve zmìnì èísla, pokud se nezmìnil focus)
        Call GenerovatCisloZadosti

        With Cells(radek, 4)                                'è. žádosti/fa
            .NumberFormat = "@"
            .Value = CStr(txtDlrZadost.Value)
            .HorizontalAlignment = xlCenter
        End With

        With Cells(radek, 6)                                'è. subdod. fa
            .NumberFormat = "@"
            .Value = CStr(txtSubFak.Value)
        End With

        'cena celkem - konverze èárky na teèku:
        With Cells(radek, 9)                                'cena celkem
    '        .Value = txtKamCena.Value
            .Value = Application.WorksheetFunction.Substitute(txtKamCena.Value, ",", ".")
            .NumberFormat = "#,##0.00 $"
        End With

        Cells(radek, 10).Value = txtKamProcento.Value        'procento pøíspìvku

        'schválený pøíspìvek - konverze èárky na teèku:
        With Cells(radek, 11)
            .Value = Application.WorksheetFunction.Substitute(txtKamSchval.Value, ",", ".")
            .NumberFormat = "#,##0.00 $"
        End With

    '    'neschválená èástka - konverze èárky na teèku:
        With Cells(radek, 12)
            .Value = Application.WorksheetFunction.Substitute(txtKamNeschval.Value, ",", ".")
            .NumberFormat = "#,##0.00 $"
        End With

        'BONUS - kvùli pøepisování s kombinovaným èerpáním pøesunut na konec bloku

        'zápis hodnocení:
        '----------------
        On Error Resume Next

        'hodnocení 1:
        With Cells(radek, 22)
            If chkHodn1.Value = True Then
                .Value = "ano"
            Else
                .Value = "ne"
            End If
        'komentáø:
            If txtHodn1.Value <> "" Then
                .AddComment
                .Comment.Text Text:=txtHodn1.Value
            Else
                .Comment.Delete
            End If
        End With

        'hodnocení 2:
        With Cells(radek, 23)
            If chkHodn2.Value = True Then
                .Value = "ano"
            Else
                .Value = "ne"
            End If
        'komentáø:
            If txtHodn2.Value <> "" Then
                .AddComment
                .Comment.Text Text:=txtHodn2.Value
            Else
                .Comment.Delete
            End If
        End With

        'hodnocení 3:
        With Cells(radek, 24)
            If chkHodn3.Value = True Then
                .Value = "ano"
            Else
                .Value = "ne"
            End If
        'komentáø:
            If txtHodn3.Value <> "" Then
                .AddComment
                .Comment.Text Text:=txtHodn3.Value
            Else
                .Comment.Delete
            End If
        End With

        'hodnocení 4:
        With Cells(radek, 25)
            If chkHodn4.Value = True Then
                .Value = "ano"
            Else
                .Value = "ne"
            End If
        'komentáø:
            If txtHodn4.Value <> "" Then
                .AddComment
                .Comment.Text Text:=txtHodn4.Value
            Else
                .Comment.Delete
            End If
        End With

        'hodnocení 5:
        With Cells(radek, 26)
            If chkHodn5.Value = True Then
                .Value = "ano"
            Else
                .Value = "ne"
            End If
        'komentáø:
            If txtHodn5.Value <> "" Then
                .AddComment
                .Comment.Text Text:=txtHodn5.Value
            Else
                .Comment.Delete
            End If
        End With

        'hodnocení 6: - POZNÁMKY K REKLAMNÍMU DOKUMENTU!!!
        With Cells(radek, 27)
            If chkHodn6.Value = True Then
                .Value = "ano"
            Else
                .Value = "ne"
            End If
        'komentáø:
            If txtHodn6.Value <> "" Then
                .AddComment
                .Comment.Text Text:=txtHodn6.Value
            Else
                .Comment.Delete
            End If
        End With
        'konec hodnocení

        With Cells(radek, 33)                                       'kombinované èerpání
            If Me.chkKombinace.Value = True Then
                .Value = "ano"
                Cells(radek, 1).Interior.Color = RGB(230, 185, 184)
            Else
                .Value = "ne"
                Cells(radek, 1).Interior.Color = xlNone
            End If
        End With
    
        With Cells(radek, 13)               'BONUS - s typem bonusu do komentáøe
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
    
        'schválení/zamítnutí kampanì - zápis do sl. 14 a barva písma øádku:
        '---------------------------
        With Cells(radek, 14)
            If optSchvalit.Value = True Then
                'schválený záznam:
                .Value = "ano"
                Range("A" & radek & ":AZ" & radek).Font.Color = RGB(0, 176, 80)
            ElseIf optZamitnout.Value = True Then
                'zamítnutý záznam:
                .Value = "ne"
                Range("A" & radek & ":AZ" & radek).Font.Color = RGB(255, 0, 0)
            Else
                'rozpracovaný záznam:
                .Value = ""
                Range("A" & radek & ":AZ" & radek).Font.Color = RGB(0, 0, 0)
            End If
        End With
    
        'schválení administrátorem - zápis do sl. 34 a výplò èísla kampanì:
        '---------------------------
        With Cells(radek, 34)
            If Me.chkKamAdmin.Value = True Then
                'schváleno adminem:
                .Value = "ano"
                Range("B" & radek).Interior.Color = vbYellow
            Else
                'zatím neschváleno adminem:
                .Value = "ne"
                With Range("B" & radek).Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        End With
    
        'stav uzávìrky - po zápisu je NE:
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

    End If  'konec zápisu položek pouze pro kampanì


    'zápis položek pro kampanì i reklamy:
    '====================================

    Cells(radek, 1).Value = CDate(txtDatum.Value)       'datum
    Cells(radek, 1).NumberFormat = "d/m/yyyy;@"         'oprava pro Exc2013 (od v. 5.3.1)
    Cells(radek, 2).Value = txtKamCislo.Value           'interní è. kampanì/rekl.
    Cells(radek, 7).Value = cmbKamNazev.Value           'název kam.

    With Cells(radek, 8)                                'oblast pøíspìvku
        If Me.optProdej.Value = True Then
            .Value = "prodej"
        ElseIf Me.optServis.Value = True Then
            .Value = "servis"
        End If
    End With

    With Me
    Cells(radek, 15).Value = .cmbKamTyp.Value           'typ kam.
    Cells(radek, 16).Value = .cmbKamZamereni.Value      'zamìøení kam.
    Cells(radek, 17).Value = .cmbKamMedium.Value        'médium-typ
    Cells(radek, 18).Value = .cmbKamMediumNazev.Value   'médium-název
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
    Cells(radek, 20).Value = .cmbKamFormat.Value        'formát
    Cells(radek, 21).Value = .txtKamPocet.Value         'poèet ks mat.
    End With    'Me

    mesicDokl = Month(Cells(radek, 1).Value)
    Cells(radek, 28).Value = mesicDokl                      'mìsíc
    Cells(radek, 29).Value = Year(Cells(radek, 1).Value)    'rok
    If typZaznamu <> "edit" Then
        With Cells(radek, 30)
            .Value = userAp                         'kod uživatele - JEN PRO NOVÉ A KOPIE
'            .AddComment
'            .Comment.Text Text:=Day(Date) & "." & Month(Date) & "." & Year(Date)    'datum zápisu
        End With
    End If

    With Cells(radek, 31)
        If Me.optKamVlastni.Value = True Then
            .Value = "ano"                              'vlastní kampaò DLR
        Else
            .Value = "ne"                               'kampaò importéra
        End If
    End With

    With Cells(radek, 32)
        If Me.chkSpolReklama.Value = True Then
            .Value = "ano"                              'spoleèná reklama více DLR
        Else
            .Value = "ne"
        End If
    End With

    'datum zápisu (nové a kopie):
    If typZaznamu <> "edit" Then
        Cells(radek, 37).Value = CDate(Day(Date) & "." & Month(Date) & "." & Year(Date))
        Cells(radek, 37).NumberFormat = "d/m/yyyy;@"
    End If

    'LOG - sledování zmìn mìsíce v záznamu:
    If reklama = False And typZaznamu = "edit" And mesicPuvodni <> mesicDokl Then
        Call frm_MENU.SysEvent _
            (sysZprava:=txtKamCislo.Value & " - zmìna mìsíce z " & mesicPuvodni & " na " & mesicDokl)
    End If

    'zápis položek pouze pro reklamy:
    '================================
    If reklama = True Then

        With Cells(radek, 13)               'BONUS - reklama bez typu bonusu do komentáøe
            If chkBonus.Value = True Then
                .Value = "ano"
            Else
                .Value = "ne"
            End If
            .HorizontalAlignment = xlCenter
        End With

        'schválení/zamítnutí reklamy - zápis do sl. 14 a barva písma øádku:
        '---------------------------
        With Cells(radek, 14)
            If optSchvalitRekl.Value = True Then
                'schválený záznam:
                .Value = "ano"
                Range("A" & radek & ":AZ" & radek).Font.Color = RGB(0, 176, 80)
            ElseIf optZamitnoutRekl.Value = True Then
                'zamítnutý záznam:
                .Value = "ne"
                Range("A" & radek & ":AZ" & radek).Font.Color = RGB(255, 0, 0)
            Else
                'rozpracovaný záznam:
                .Value = ""
                Range("A" & radek & ":AZ" & radek).Font.Color = RGB(0, 0, 0)
            End If

On Error Resume Next    'kvuli .Comment.Delete
            'zapsat komentáø ke schválení kampanì:
            If Me.txtReklPozn.Value <> "" Then
                .AddComment
                .Comment.Text Text:=Me.txtReklPozn.Value
            Else
                .Comment.Delete
            End If
        End With    'konec práce s Cells(radek, 14)

On Error GoTo Chyba
        
        If typZaznamu = "edit" Then
            Cells(radek, 34).Value = vyuzito                    'využití naètené pøi editaci
        Else
            Cells(radek, 34).Value = "ne"                       'default využití nové rekl.
        End If

    End If      'konec zápisu položek pouze pro reklamy
    '--------------------------------------------------

    'písmo øádku standard:
    '---------------------
    Range("A" & radek & ":AZ" & radek).Font.Italic = False

    'pøepoèet rozpoètu DLR:
    '----------------------
    If reklama = False Then Call PrepocetDlr

'    DEBUG
'    ActiveWindow.DisplayHeadings = False    'vypnout èísla øádkù a sloupcù

    'zapsat využití reklamy do kampanì:
    '----------------------------------
    If idReklamy <> -100 Then       'default - nastaveno v INIT

        If spolReklama = True Then
            Call OtevritSpolecnouReklamu
            Set wsDlrData2 = ActiveWorkbook.Worksheets(2)   'pøepíše se pùvodní promìnná - z listu DLR na 999SR!
        End If

        'dále stejné pøíkazy pro zápis využité rekl. DLR/999SR/Retail:
        wsDlrData1.Cells(radek, 36).Value = wsDlrData2.Cells(idReklamy + 4, 2).Value    'do kampanì zapsat è. zdrojové rekl.
        With wsDlrData2
            .Unprotect Password:=heslo
            .Cells(idReklamy + 4, 34).Value = "ano"
            .Cells(idReklamy + 4, 35).Value = .Cells(idReklamy + 4, 35).Value + 1
            .Protect Password:=heslo
        End With

        'zápis využité SR/Retail - návrat k souboru DLR:
        If spolReklama = True Then
            With ActiveWorkbook
                .Save
                .Close                                      'zavøít 999SR
            End With
            Windows(soubor).Activate                        'aktivace souboru DLR
            Set wsDlrData2 = ActiveWorkbook.Worksheets(2)   'návrat promìnné na list reklam DLR!
        End If

    End If

    'zapsat využití Bonusu R/I do rozpoètù:
    '--------------------------------------
    If idBonusy <> -100 Then       'default - nastaveno v INIT

        'otevøít rozpoèty:
        With Application
        .ScreenUpdating = False
        .StatusBar = "Zapisuji využití IP..."
        End With
        Set wbRozpocty = GetObject(cesta & "\Data\Rozpocty\" & souborRozpocty)
        Set wsRozpocty = wbRozpocty.Worksheets(1)
    
        'zapsat využití do rozpoètù:
        If typBonus = "R" Then          'Region bonus
            If idBonusy = 0 Then        'jarní
                wsRozpocty.Cells(dlr + 1, 109).Value = "ano"
            ElseIf idBonusy = 1 Then    'podzimní
                wsRozpocty.Cells(dlr + 1, 112).Value = "ano"
            Else
                MsgBox "idBonusy není v poøádku", vbCritical, "Rozpoèty - využití Bonusu"
            End If
        ElseIf typBonus = "I" Then      'Indiv. bonusy vè. centrálních
            wsRozpocty.Cells(dlr + 1, idBonusy + 154).Value = "ano"
        Else
            MsgBox "Typ bonusu není v poøádku", vbCritical, "Rozpoèty - využití Bonusu"
        End If

        'zavøít rozpoèty:
        Set wsRozpocty = Nothing
        Windows(souborRozpocty).Visible = True     'pøi ukládání rozpoètu po otevøení metodou GetObject (sešit zùstával skrytý)
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
    'konec zápisu øádku

    Cells(radek, 1).Select      'nastavit zapsaný øádek, hlavnì pro tisk košilky

    'zamknout list po zápisu øádku:
    '------------------------------
    ActiveSheet.Protect Password:=heslo

    'uložení souboru DLR:
    '--------------------
    With Application
        .StatusBar = "UKLÁDÁNÍ SOUBORU:  " & ActiveWorkbook.Name & "..."
        .ReferenceStyle = xlA1
    ActiveWorkbook.Save
        .StatusBar = False
    End With

    'kontrola è. žádosti, pøi zmìnì upozornìní na košilku pro novou žádost DLR:
    '-------------------------------------------------------------------------
    If reklama = False Then
        If Cells(radek, 4).Value <> Cells(radek - 1, 4).Value Then
            MsgBox "Zmìnilo se èíslo žádosti, možná bude dobré vygenerovat košilku...      ", vbInformation, aplikace
        End If
    End If

    'admin editace - skok o øádek dolù:
    If reklama = False And admin = True Then _
        Cells(radek + 1, 1).Select

    'reset pøíznaku spoleèná reklama:
    spolReklama = False

    'zapsat do syslogu:
    If typZaznamu <> "edit" Then    'nový záznam / kopie
        If reklama = False Then
            Call frm_MENU.SysEvent(sysZprava:=dlrNw(dlr) & " - zápis kampanì " & txtKamCislo.Value & " (" & typZaznamu & ")")
        Else
            Call frm_MENU.SysEvent(sysZprava:=dlrNw(dlr) & " - zápis rekl. dokumentu " & txtKamCislo.Value & " (" & typZaznamu & ")")
        End If
    Else                            'editace
        If reklama = False Then
            Call frm_MENU.SysEvent(sysZprava:=dlrNw(dlr) & " - editace kampanì " & txtKamCislo.Value)
        Else
            Call frm_MENU.SysEvent(sysZprava:=dlrNw(dlr) & " - editace rekl. dokumentu " & txtKamCislo.Value)
        End If
    End If
    If idBonusy <> -100 Then
        If typBonus = "I" Then
            Call frm_MENU.SysEvent(sysZprava:=dlrNw(dlr) & " - využití IP " & idBonusy + 1)
        ElseIf typBonus = "R" Then
            Call frm_MENU.SysEvent(sysZprava:=dlrNw(dlr) & " - využití RB " & idBonusy + 1)
        End If
    End If

    'zapsat status do evidence dokladù:     MRO 2013
    '----------------------------------
    If reklama = False Then
        With Application
        .StatusBar = "Zapisuji do Evidence dokladù..."
        .ScreenUpdating = False
        End With
        Workbooks.Open (cesta & "\Data\Doklady\DLR_doklady_" & typMro & ".xlsm")
        ActiveSheet.Unprotect Password:=heslo
        Cells(dlr + 1, mesicDokl + 2).Value = CDate(Day(Date) & "." & Month(Date) & "." & Year(Date))
        'status = formát:
        Select Case Me.chkHodn3.Value   'kontrola pøíznaku Žádost zaslána
        Case True
            Cells(dlr + 1, mesicDokl + 2).Font.Color = RGB(0, 176, 80)
            Call frm_MENU.DocEvent _
                (sysZprava:=Cells(dlr + 1, 1).Value & " - došla žádost za mìsíc " & mesicSoub & "(zápis kampanì)")
        Case False
            Cells(dlr + 1, mesicDokl + 2).Font.Color = RGB(0, 0, 0)
            Call frm_MENU.DocEvent _
                (sysZprava:=Cells(dlr + 1, 1).Value & " - došly podklady za mìsíc " & mesicSoub & "(zápis kampanì)")
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

    'vytvoøit složku v dokumentech NW:      MRO 2013
    '---------------------------------
    '(kontrola, pokud se nevytvoøí pøes DLR info)
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
            Call frm_MENU.DocEvent(sysZprava:=dlrNw(dlr) & " - vytvoøena složka " & mesicSoub & "_" & typMro & "(zápis kampanì)")
        End If
    End If

'    appEvents = False
    Unload Me
'    appEvents = True
    Exit Sub

Chyba:  'pøi chybì programu zamknout list
'    MsgBox Err.Number & " " & Err.Description, vbCritical, "===  CHYBA  ==="
'    ActiveSheet.Protect Password:=heslo         'zamknout list

End Sub



Private Sub PrepocetDlr()   'pouze pro kampanì
    'pøepoèet rozpoètu DLRa
    'voláno z cmdUlozit_Click

    'pøi pøepoètu kontrolovat bonus akce, schválené kampanì, oblast pøíspìvku!!!

Dim cerpProdej      As Single   'èerpání pøísp. na prodej
Dim cerpServis      As Single   'èerpání pøísp. na servis
Dim cerpCelkem      As Single   'èerpání celkem

    cerpProdej = 0
    cerpServis = 0
    cerpCelkem = 0

    'naèíst èerpání ze zapsaných kampaní:
    '------------------------------------
    poslRadek = Cells(Rows.Count, 2).End(xlUp).Row

    For i = 4 To poslRadek
        If Cells(i, 13).Value = "ne" And Cells(i, 14).Value = "ano" Then    'øádek není bonus akce a je schválený!!
            cerpCelkem = cerpCelkem + Cells(i, 11).Value                    'pøièíst èerp. celkem
            If Cells(i, 8).Value = "prodej" Then
                cerpProdej = cerpProdej + Cells(i, 11).Value                'pøièíst èerpání prodej
            ElseIf Cells(i, 8).Value = "servis" Then
                cerpServis = cerpServis + Cells(i, 11).Value                'pøièíst èerpání servis
            End If
        End If
    Next i

End Sub



Private Sub cmdStorno_Click()
    'tl. Storno

    If reklama = False And InStr(ActiveWorkbook.Name, "999SR") > 0 Then   'pouze pro kampanì!
        ActiveWorkbook.Close
        Windows(soubor).Activate    'aktivace souboru DLR
    End If
    Set wsDlrData1 = Nothing
    Set wsDlrData2 = Nothing
    Unload frm_ZAZNAM
    
End Sub



Private Function FileExists(fname) As Boolean
    ' vrací TRUE, pokud soubor existuje

    FileExists = (Dir(fname) <> "")

End Function     '----- End of Function FileExists -----



Private Function PathExists(pname) As Boolean
    ' vrací TRUE, pokud cesta existuje

    If Dir(pname, vbDirectory) = "" Then
        PathExists = False
    Else
        PathExists = (GetAttr(pname) And vbDirectory) = vbDirectory
    End If

End Function     '----- End of Function PathExists -----



'============== PØEVZETÍ REKLAMY DO KAMPANÌ: ===================

Private Sub cmdPrevzitReklamu_Click()
    'pøevzít záznam ze schválených reklam DLR

    frmReklamy.Show

End Sub



Private Sub cmdPrevzitReklamu2_Click()
    'pøevzít záznam ze spoleèných reklam (DLR 999SR)

    'otevøít soubor spol. reklam podle instalace:
    Call OtevritSpolecnouReklamu
    frmReklamy.Show

End Sub



Private Sub OtevritSpolecnouReklamu()
    'otevøít soubor spol. reklam podle instalace:

    Workbooks.Open (cesta & "\Data\DLR_" & typMro & "\999SR - SPOLEÈNÁ REKLAMA [" & typMro & "].xlsm")

End Sub



Sub NacistReklamu()
    'naèíst data vybrané reklamy do frm
    'voláno z frmReklamy.cmdPrevzit

Dim radek2      As Long

    appEvents = False

    On Error GoTo Chyba
    radek2 = idReklamy + 4   'listindex od 0, záhlaví 3 øádky

    'info o pøevzetí reklamy:
    With Me.lblZdrojovaReklama
        .Caption = "Zdrojová reklama: " & wsDlrData2.Cells(radek2, 2).Value & "."
        .Visible = True
    If wsDlrData2.Cells(radek2, 34).Value = "ano" Then
        .Caption = .Caption & " Již využito: " & wsDlrData2.Cells(radek2, 35).Value & "x."
    Else
        .Caption = .Caption & " Ještì nevyužito."
    End If
    End With    'lblZdrojovaReklama

    'zápis pøevzetí do logu:
    Call frm_MENU.SysEvent(sysZprava:=dlrNw(dlr) & " - využití rekl. dokumentu (" & wsDlrData2.Cells(radek2, 2).Value & " / " & Me.txtKamCislo.Value & ")")
    
    'data reklamy do kampanì:
    With wsDlrData2

    If .Cells(radek2, 13).Value = "ano" Then             'bonus akce
        Me.chkBonus.Value = True
    Else
        Me.chkBonus.Value = False
    End If

    If .Cells(radek2, 32).Value = "ano" Then             'spoleèná rekl.
        Me.chkSpolReklama.Value = True
    Else
        Me.chkSpolReklama.Value = False
    End If

    If .Cells(radek2, 31).Value = "ne" Then              'vlastní kam.
        Me.optKamImport.Value = True
    Else
        Me.optKamVlastni.Value = True
    End If

    If .Cells(radek2, 8).Value = "prodej" Then                'oblast èerpání
        Me.optProdej.Value = True
    Else
        Me.optServis.Value = True
    End If

    Me.cmbKamNazev.Value = .Cells(radek2, 7).Value               'název kam.
    Me.cmbKamTyp.Value = .Cells(radek2, 15).Value                'typ kam.
    Me.cmbKamZamereni.Value = .Cells(radek2, 16).Value           'zamìøení kam.
    Me.cmbKamMedium.Value = .Cells(radek2, 17).Value             'typ média
    Me.cmbKamMediumNazev.Value = .Cells(radek2, 18).Value        'název média
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
    Me.cmbKamFormat.Value = .Cells(radek2, 20).Value             'formát
    If .Cells(radek2, 14).Value = "ano" Then Me.chkHodn6.Value = True       'schválení reklamy
    On Error Resume Next    'prázdný komentáø:
    Me.txtHodn6.Value = .Cells(radek2, 14).Comment.Text          'poznámka k rekl. dokumentu
    On Error GoTo Chyba

    End With    'wsDlrData2
    Me.txtDatum.SetFocus
    Exit Sub

    appEvents = True

Chyba:
'    MsgBox Err.Number & " " & Err.Description, vbCritical, aplikace

End Sub



Sub NacistBonus()
    'naèíst data vybraného bonusu do frm
    'voláno z frmBonusy.cmdPrevzit

    'zapsat data do frm:
    With Me
    If typBonus = "I" Then
        .txtHodn6.Value = frmBonusy.lstBonusInd.List(idBonusy, 1)
    ElseIf typBonus = "R" Then
        .txtHodn6.Value = frmBonusy.lstBonusReg.List(idBonusy, 1)
    Else
        MsgBox "Typ bonusu není v poøádku, kontaktujte podporu.", vbCritical, "Rozpoèty - využití Bonusu"
    End If
    '    .txtKamCena.Value = frmBonusy.lstBonusInd.List(idBonusy, 0)

    'zapsat indikátor do frm:
    With .lblZdrojovyBonus
    .Visible = True
    .Caption = "Pøevzato: " & idBonusy + 1 & " / " & typBonus
    End With
    .lblTypBonus.Caption = typBonus

    End With    'Me

End Sub



Private Sub imgDokumenty_Click()
    
    'podle NW vyhledat DLR a získat index: (pro využití bonusù a zápis do evidence dokladù)
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

    'podle NW vyhledat DLR a získat index: (pro využití bonusù a zápis do evidence dokladù)
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



'-------------------- otevøení èíselníkù --------------------

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




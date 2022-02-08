# VB


    Private Sub CommandButton1_Click()
        Call MsgBox("Nacisnieto Przycisk")
        Call informacja
    End Sub

    Private Sub CommandButton2_Click()
        Dim PromienR As Double

        PromienR = InputBox("Podaj promien kola")
        If PromienR > 0 Then
        MsgBox Geometria.PoleKola(PromienR)
        Else
        MsgBox ("Bledny Promien")
        End If
    End Sub

    Private Sub CommandButton3_Click()
        Dim lancuch1 As String
        Dim Lancuch2 As String

        lancuch1 = InputBox("Podaj Lancuch1")
        Lancuch2 = InputBox("Podaj Lancuch2")

        If InStr(Lancuch2, lancuch1) > 0 Then
        MsgBox "Z Lancucha 2 mozna uzyskac Lancuch 1 poprzez usuniecie znakow z poczatku i/lub konca"
        Else
        MsgBox "Z Lancucha 2 nie mozna uzyskac lancucha 1 poprzez usuniecie znakow z poczatku i/lub konca"
        End If
    End Sub

    Private Sub CommandButton4_Click()
        MsgBox (Now)
    End Sub

    Private Sub CommandButton5_Click()
        Dim dzien As Integer
        Dim miesiac As Integer
        Dim rok As Integer
        Dim Data As Date
        dzien = InputBox("Podaj dzien")
        miesiac = InputBox("Podaj miesiac")
        rok = InputBox("Podaj rok")
        Data = DateSerial(rok, miesiac, dzien)
        MsgBox (WeekdayName(Weekday(Data, vbMonday)))
    End Sub
    
    Option Explicit

    Type Student
        Imie As String
        DataUrodzenia As Date
        Wiek As Byte
        Wzrost As Single
        KolorLegitymacji As KoloryTeczy
        CzyKobieta As Boolean
    End Type
    Enum KoloryTeczy
        Czerwony
        Pomaranczowy
        Zolty
        Zielony
        Niebieski
    End Enum
        Function SumaKwadratow(LiczbaA As Integer, LiczbaB As Integer)
        SumaKwadratow = LiczbaA ^ 2 + LiczbaB ^ 2
        End Function

    Sub informacja()
        Call MsgBox("Dwa", vbExclamation + vbOKOnly, "cyfra")
        Call MsgBox("komputer", vbCritical + vbOKCancel, "tak")
        Call MsgBox("myszka", vbInformation + vbYesNo, "Nie")
        Call MsgBox("Klawiatura", vbQuestion + vbRetryCancel, "placki")
    End Sub
    
    Option Explicit

    Function PoleProstokata(BokA As Double, BokB As Double) As Double
        PoleProstokata = BokA * BokB
    End Function
    Function ObwodProstokata(BokA As Double, BokB As Double) As Double
        ObwodProstokata = BokA * 2 + BokB * 2
    End Function
    Function PoleTrapezu(BokA As Double, BokB As Double, WysH As Double) As Double
        PoleTrapezu = (BokA + BokB) / 2 * WysH
    End Function
    Function PoleTrojkata(PodsA As Double, WysH As Double) As Double
        PoleTrojkata = PodsA * WysH / 2
    End Function
    Function PoleKola(PromienR As Double) As Double
        PoleKola = PromienR ^ 2 * WorksheetFunction.Pi
    End Function
    Function ObwodKola(PromienR As Double) As Double
        ObwodKola = 2 * PromienR * WorksheetFunction.Pi
    End Function
    Function ObjetoscKuli(PromienR As Double) As Double
        ObjetoscKuli = 4 / 3 * WorksheetFunction.Pi * PromienR ^ 3
    End Function
    Function PolePowierzchniKuli(PromienR As Double) As Double
        PolePowierzchniKuli = 4 * WorksheetFunction.Pi * PromienR ^ 2
    End Function
    Function PoleKwadratu(BokA As Double) As Double
        PoleKwadratu = PoleProstokata(BokA, BokA)
    End Function
    Function PolePowCalkowProstopadloscianu(BokA As Double, BokB As Double, WysH As Double) As Double
        PolePowCalkowProstopadloscianu = PoleProstokata(BokA, WysH) * 2 + PoleProstokata(BokB, WysH) * 2 + PoleProstokata(BokA, BokB)
    End Function
    Function ObjetoscWalca(PromienR As Double, WysH As Double) As Double
        ObjetoscWalca = PoleKola(PromienR) * WysH
    End Function
    Function PolePowBocznejWalca(PromienR As Double, WysH As Double) As Double
        PolePowBocznejWalca = ObwodKola(PromienR) * WysH
    End Function
    
    Option Explicit

    Function Zadanie59(Lancuch As String, Lancuch2 As String) As Boolean
        If Len(Lancuch) = Len(Lancuch2) Then
            Zadanie59 = True
        Else
        End If
    End Function
    Function IleZnakowWLancuchu(Lancuch As String, Znak As String) As Integer
        Dim LiczbaWystapien As Integer
        Dim a As Integer

        LiczbaWystapien = 0
            For a = 1 To Len(Lancuch)
            If Znak = Cwiczenie53(Lancuch, a) Then
            LiczbaWystapien = LiczbaWystapien + 1
            End If
            Next
        IleZnakowWLancuchu = LiczbaWystapien
    End Function
    Function CzyJestZapisemLiczbyCalkowitej(Lancuch As String) As Boolean
    Dim Lancuch2 As String
    Lancuch2 = ","
    If IsNumeric(Lancuch) And InStr(Lancuch, Lancuch2) = 0 Then
    CzyJestZapisemLiczbyCalkowitej = True
    Else
    End If
    End Function
    Function CzyZaczynaSieZWielkiejLitery(Lancuch As String) As Boolean
    Dim PierwszaLitera As String
    PierwszaLitera = StrConv(Left(Lancuch, 1), vbUpperCase)

    If Left(Lancuch, 1) = PierwszaLitera Then
    CzyZaczynaSieZWielkiejLitery = True
    Else
    CzyZaczynaSieZWielkiejLitery = False
    End If
    End Function
    Function SkrotOrganizacji(NazwaOrganizacji As String) As String
        Dim PoprawnaNazwa As String

        PoprawnaNazwa = StrConv(NazwaOrganizacji, vbProperCase)


    End Function
    
    Option Explicit

    Function Zadanie59(Lancuch As String, Lancuch2 As String) As Boolean
        If Len(Lancuch) = Len(Lancuch2) Then
            Zadanie59 = True
        Else
        End If
    End Function
    Function IleZnakowWLancuchu(Lancuch As String, Znak As String) As Integer
        Dim LiczbaWystapien As Integer
        Dim a As Integer

        LiczbaWystapien = 0
            For a = 1 To Len(Lancuch)
            If Znak = Cwiczenie53(Lancuch, a) Then
            LiczbaWystapien = LiczbaWystapien + 1
            End If
            Next
        IleZnakowWLancuchu = LiczbaWystapien
    End Function
    Function CzyJestZapisemLiczbyCalkowitej(Lancuch As String) As Boolean
        Dim Lancuch2 As String
        Lancuch2 = ","
        If IsNumeric(Lancuch) And InStr(Lancuch, Lancuch2) = 0 Then
        CzyJestZapisemLiczbyCalkowitej = True
        Else
        End If
        End Function
        Function CzyZaczynaSieZWielkiejLitery(Lancuch As String) As Boolean
        Dim PierwszaLitera As String
        PierwszaLitera = StrConv(Left(Lancuch, 1), vbUpperCase)

        If Left(Lancuch, 1) = PierwszaLitera Then
            CzyZaczynaSieZWielkiejLitery = True
        Else
            CzyZaczynaSieZWielkiejLitery = False
        End If
    End Function
    Function SkrotOrganizacji(NazwaOrganizacji As String) As String
        Dim PoprawnaNazwa As String

        PoprawnaNazwa = StrConv(NazwaOrganizacji, vbProperCase)


    End Function
    
    Option Explicit


    Function Cwiczenie41(Pod As Integer, Wyk As Integer) As Variant
    Dim Potega As Double
    Dim Licznik As Integer

    If Pod = 0 And Wyk = 0 Then
        Cwiczenie41 = ("Niepoprawne dane")
    Else
        Potega = 1
        For Licznik = 1 To Wyk
            Potega = Pod * Potega
        Next
        Cwiczenie41 = Potega
    End If
    End Function
    Function Cwiczenie42(a As Integer) As Integer
    Dim LiczbaDziel As Integer
    Dim I As Integer
    I = 1
    LiczbaDziel = 0
    If a <= 0 Then
        Cwiczenie42 = 0
    Else
       For I = 1 To a
        If a Mod I = 0 Then
            LiczbaDziel = LiczbaDziel + 1
        End If
       Next
    End If
    Cwiczenie42 = LiczbaDziel
    End Function

    Function Cwiczenie43(LiczbaRzymska As String) As Integer
    Select Case LiczbaRzymska
    Case "I"
        Cwiczenie43 = 1
    Case "V"
        Cwiczenie43 = 5
    Case "X"
        Cwiczenie43 = 10
    Case "L"
        Cwiczenie43 = 50
    Case "C"
        Cwiczenie43 = 100
    Case "D"
        Cwiczenie43 = 500
    Case "M"
        Cwiczenie43 = 1000
    End Select
    End Function
    Function CzyPrzestępny(rok As Integer) As Boolean
    If rok Mod 400 = 0 Then
        CzyPrzestępny = True
    ElseIf rok Mod 100 = 0 Then
        CzyPrzestępny = False
    ElseIf rok Mod 4 = 0 Then
        CzyPrzestępny = True
    Else
        CzyPrzestępny = False
    End If
    End Function
    Function IleDniWMies(miesiac As Byte, rok As Integer) As Byte
    Select Case CzyPrzestępny(rok)
        Case False
            Select Case miesiac
            Case 1
                IleDniWMies = 31
            Case 2
                IleDniWMies = 28
            Case 3
                IleDniWMies = 31
            Case 4
                IleDniWMies = 30
            Case 5
                IleDniWMies = 31
            Case 6
                IleDniWMies = 30
            Case 7
                IleDniWMies = 31
            Case 8
                IleDniWMies = 31
            Case 9
                IleDniWMies = 30
            Case 10
                IleDniWMies = 31
            Case 11
                IleDniWMies = 30
            Case 12
                IleDniWMies = 31
            End Select
        Case True
            Select Case miesiac
            Case 1
                IleDniWMies = 31
            Case 2
                IleDniWMies = 29
            Case 3
                IleDniWMies = 31
            Case 4
                IleDniWMies = 30
            Case 5
                IleDniWMies = 31
            Case 6
                IleDniWMies = 30
            Case 7
                IleDniWMies = 31
            Case 8
                IleDniWMies = 31
            Case 9
                IleDniWMies = 30
            Case 10
                IleDniWMies = 31
            Case 11
                IleDniWMies = 30
            Case 12
                IleDniWMies = 31
            End Select
    End Select
    End Function
    Function Cwiczenie44(x As Double) As Double

    Cwiczenie44 = IIf(x = 0, 777, Tan(x * Abs(x)) + Sin(Abs(x)) + Cos(x ^ (2) - 1))

    End Function
    Function Cwiczenie45(Liczba As String) As Double
    Cwiczenie45 = Switch(Liczba = "I", 1, Liczba = "V", 5, Liczba = "X", 10, Liczba = "L", 50, Liczba = "C", 100, Liczba = "D", 500, Liczba = "M", 1000)

    End Function
    Function Cwiczenie46(Liczba As Double) As String
    Cwiczenie46 = IIf(Liczba Mod 5 = 0, "Reszta z dzielenia wynosi zero", Choose(Liczba Mod 5, "jeden", "dwa", "trzy", "cztery"))
    End Function
    Function Cwiczenie47(LiczbaA As Double, LiczbaB As Double) As Double
    Cwiczenie47 = IIf(LiczbaA >= LiczbaB, LiczbaA, LiczbaB)
    End Function
    Function Cwiczenie48(a As Double, B As Double, C As Double) As Double
    Cwiczenie48 = Switch(a >= B And a >= C, a, B >= a And B >= C, B, C >= a And C >= B, C)
    End Function
    
    Option Explicit
    Function MojaSilnia(N As Integer) As Double
    Dim Iloczyn As Double
    Dim Liczba As Integer

    Iloczyn = 1
    Liczba = 1
        Do
            Iloczyn = Iloczyn * Liczba
            Liczba = Liczba + 1
        Loop While Liczba <= N
    MojaSilnia = Iloczyn

    End Function
    Function NajwWspolDziel(a As Integer, B As Integer) As Double
        Do Until a = B
            If a > B Then
                a = a - B
            Else
                B = B - a
            End If
        Loop
        NajwWspolDziel = a

    End Function

    Function SumaCyfrLiczbyCalkow(a As Long) As Byte
        Dim suma As Byte

        suma = 0
        Do Until a < 10
            suma = suma + (a Mod 10)
            a = a \ 10
        Loop
        SumaCyfrLiczbyCalkow = a + suma
     End Function
    Function OdwrotnaKolejnosc(a As Integer) As Double

        While a > 0
            OdwrotnaKolejnosc = OdwrotnaKolejnosc * 10 + a Mod 10
            a = a \ 10
        Wend


    End Function
    Sub SumaPodanychLiczb()
        Dim IleLiczb As Byte
        Dim SumaLiczb As Long
        Dim a As Long


        IleLiczb = InputBox("Podaj ile liczb chcesz zsumować")
        For a = 1 To IleLiczb
            SumaLiczb = SumaLiczb + InputBox("podaj liczbe")
        Next a
        MsgBox "Suma liczb wynosi " & SumaLiczb
    End Sub

    Sub LiczbaNajmniejsza()
        Dim IleLiczb As Byte
        Dim Minimum As Long
        Dim a As Byte

        IleLiczb = InputBox("Podaj ile liczb  zamierzasz podac")
        For a = 1 To IleLiczb

        Next a
    End Sub
    
    Option Explicit
    Function MaxZDwochLiczb(LiczbaA, LiczbaB) As Double
    If LiczbaA >= LiczbaB Then
    MaxZDwochLiczb = LiczbaA
    Else
    MaxZDwochLiczb = LiczbaB
    End If
    End Function

    Function MaxZTrzechLiczb(LiczbaA, LiczbaB, LiczbaC) As Double
        If LiczbaA >= LiczbaB And LiczbaA >= LiczbaC Then
        MaxZTrzechLiczb = LiczbaA
        ElseIf LiczbaB >= LiczbaA And LiczbaB >= LiczbaC Then
        MaxZTrzechLiczb = LiczbaB
        Else
        MaxZTrzechLiczb = LiczbaC
        End If
    End Function

    Function CzyWiekszyOdZera(LiczbaA)
        If LiczbaA > 0 Then
        CzyWiekszyOdZera = True
        Else
        CzyWiekszyOdZera = False
        End If
    End Function
    Function CzyJednocyfrowa(LiczbaA)
        If LiczbaA < 10 And LiczbaA >= 1 And LiczbaA \ 1 = LiczbaA Then
        CzyJednocyfrowa = True
        Else
        CzyJednocyfrowa = False
        End If
    End Function
    Function CzyParzysta(LiczbaA)
        If LiczbaA \ 2 = LiczbaA / 2 Then
        CzyParzysta = True
        Else
        CzyParzysta = False
        End If
    End Function
    Function CzyPodzielnaPrzez17(LiczbaA)
        If LiczbaA \ 17 = LiczbaA / 17 Then
        CzyPodzielnaPrzez17 = True
        Else
        CzyPodzielnaPrzez17 = False
        End If
    End Function
    Function CzyCalkowita(LiczbaA)
        If LiczbaA \ 1 = LiczbaA Then
        CzyCalkowita = True
        Else
        CzyCalkowita = False
        End If
    End Function
    Sub CoPodano()
        Dim ZmiennaA As Variant

        ZmiennaA = InputBox("Podaj jedna z zmiennych:Liczba,Data,Tekst")
        If IsDate(ZmiennaA) Then
        MsgBox ("Podano date")
        ElseIf IsNumeric(ZmiennaA) Then
        MsgBox ("Podano Liczbe")
        Else
        MsgBox ("Podano tekst")
        End If


    End Sub

    Function ZlotyCzyZlotych(LiczbaA As Double)
        If LiczbaA = 1 Then
        ZlotyCzyZlotych = LiczbaA & " zloty"
        ElseIf LiczbaA Mod 100 = 12 Or 13 Or 14 Then
        ZlotyCzyZlotych = LiczbaA & " zlotych"
        ElseIf LiczbaA Mod 10 = 2 Or 3 Or 4 Then
        ZlotyCzyZlotych = LiczbaA & " zloty"
        Else
        ZlotyCzyZlotych = LiczbaA & " zlotych"
        End If
    End Function

    
        Option Explicit
    Type Miejscowosc
        Nazwa As String
        DataZalozenia As Date
        LiczbaMieszkancow As Integer
    End Type
    Sub ObliczPolewadratu()
        Dim BokKwadratu As Double
        Dim PoleKwadratu As Double

        BokKwadratu = InputBox("Podaj długość boku kwadratu")
        PoleKwadratu = Geometria.PoleKwadratu(BokKwadratu)
        MsgBox ("Pole kwadratu o danym boku wynosi: " & PoleKwadratu)
    End Sub
    Sub Wyswietlenie()
        Dim LiczbaA As Byte
        Dim LiczbaB As Integer
        Dim LiczbaC As Long
        Dim LiczbaD As Boolean
        Dim LiczbaE As Single
        Dim LiczbaF As Double
        Dim Waluta As Currency
        Dim Data As Date
        Dim slowo As String
        Dim Cos As Variant


        MsgBox (LiczbaA)
        MsgBox (LiczbaB)
        MsgBox (LiczbaC)

        MsgBox (LiczbaD)
        MsgBox (LiczbaE)
        MsgBox (LiczbaF)
        MsgBox (Waluta)
        MsgBox (Data)
        MsgBox (LiczbaA & slowo & LiczbaB)
        MsgBox (Cos)
    End Sub
    Sub Zadanie()
        Dim CosTam As Variant

        CosTam = (23)
        MsgBox TypeName(CosTam)

    End Sub


    Sub DaneMiejscowosc()
        Dim DaneMiejscowosc As Miejscowosc

        With DaneMiejscowosc

        .Nazwa = InputBox("Podaj nazwe miejscowosci")
        .LiczbaMieszkancow = InputBox("Podaj liczbe mieszkancow")
        .DataZalozenia = InputBox("Podaj date zalozenia miejscowosci")
        MsgBox ("Miejscowosc " & .Nazwa & " zalozona " & .DataZalozenia & " obecnie ma " & .LiczbaMieszkancow & " mieszkancow")
        End With
    End Sub

    Sub WartDomZmien()
        Dim DaneMiejscowosc As Miejscowosc


        With DaneMiejscowosc
            MsgBox (.Nazwa)
            MsgBox (.LiczbaMieszkancow)
            MsgBox (.DataZalozenia)
        End With
    End Sub

    Sub wartos()
    MsgBox TypeName(InputBox("wprowadź"))
    End Sub
    
    Function CelsjuszNaFahrenheita(CelsjuszA As Double) As Double
        CelsjuszNaFahrenheita = 9 / 5 * CelsjuszA + 32
    End Function
    
        Option Explicit
    Function cwiczenie75(Nazwa As String) As String
    Dim skrot() As String
    Dim a As Byte
    Dim Wyraz As String
    Wyraz = ""

    skrot = Split(Nazwa)
        If UBound(skrot) < 5 Then
            For a = LBound(skrot) To UBound(skrot)
                Wyraz = Wyraz + Left(skrot(a), 1)
            Next a
        Else
            For a = LBound(skrot) To 4
                Wyraz = Wyraz + Left(skrot(a), 1)
            Next a
        End If
    cwiczenie75 = UCase(Wyraz)
    End Function
    Function Cwiczenie76() As Double

    End Function
    Sub Cwiczenie77()

    End Sub
    Function Cwiczenie78a(Cos As String) As Double
    Dim tablica() As String
    Dim a As Byte

    tablica = Split(Cos)

    End Function
    Function Cwiczenie78b(Cos As String) As Double
    Dim tablica() As String
    Dim a As Integer
    Dim suma As Double
    tablica = Split(Cos)
    suma = 0
        For a = LBound(tablica) To UBound(tablica)
            suma = suma + tablica(a)
        Next a
    Cwiczenie78b = suma
    End Function
    Function Cwiczenie78c(Cos As String) As Double
    Dim tablica() As String
    tablica = Split(Cos)

    Cwiczenie78c = Cwiczenie78b(Cos) / (UBound(tablica) + 1)
    End Function
    Sub Cwiczenie79()

    End Sub

    Function Cwiczenie80(Lancuch As String, Litera As String) As Integer
    Dim tablica() As String
    Dim tablica2() As String
    tablica = Split(Lancuch)
    tablica2 = Filter(tablica, Litera)
    Cwiczenie80 = UBound(tablica2) + 1
    End Function

    
        Option Explicit
    Function Cwiczenie67(Lancuch As String, slowo As String) As String
    Cwiczenie67 = Replace(Lancuch, slowo, "")
    End Function
    Function MTrim(Lancuch As String) As String
    Dim lancuch1 As String

    lancuch1 = Trim(Lancuch)
    Do
    lancuch1 = Replace(lancuch1, "  ", " ")
    Loop Until InStr(lancuch1, "  ") = 0
    MTrim = lancuch1
    End Function
    Function Cwiczenie69(Lancuch As String) As Boolean
    If Lancuch = StrReverse(Lancuch) Then
    Cwiczenie69 = True
    Else
    End If

    End Function
    Sub Cwiczenie72()
        Dim Start, Koniec, Wynik As Double

        If MsgBox("Wcisnij tak") = vbOK Then
        Start = Timer
        MsgBox ("Wcisnij ok")
        Else
        End If
        Koniec = Timer
        Wynik = Koniec - Start
        MsgBox Wynik
    End Sub

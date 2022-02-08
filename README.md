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


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

Attribute VB_Name = "modSplines"
'==========================================================
' Descrizione.....: Routines di interpolazione con Splines.
' Nome dei Files..: Splines_01.bas, Splines_01.frm
'                   InfoCr.frm [, frmStampaFPB.frm]
' Data............: 27/11/1999
' Versione........: 1.0 a 32 bits
' Sistema.........: Visual Basic 6.0 sotto Windows NT.
' Scritto da......: F. Languasco ®
' E-Mail..........: MC7061@mclink.it
'==========================================================
'
'   Gli algoritmi usati sono di:
'   P. Bourke     -  Aprile 1989
'   D. Cholaseuk  -  8/Dic./1999
'
'   Le curve Splines vengono calcolate in modo parametrico
'   e quindi, con opportuni adattamenti, possono essere
'   usate per interpolare punti a n dimensioni.
'
Option Explicit
'
'Private Type CadPoint
'    x As Single
'    y As Single
'End Type

Public Sub Bezier_C(Pi() As CadPoint, Pc() As CadPoint)
'
'   Ritorna, nel vettore Pc(), i valori della curva di Bezier calcolata
'   al valore u (0 <= u <= 1). La curva e' calcolata in modo
'   parametrico con il valore 0 di u corrispondente a Pc(0)
'   ed il valore 1 corrispondente a Pc(NPC_1).
'   Questo algoritmo ricalca la forma classica del polinomio
'   di Bernstein.
'
   Dim I&, K&, NPI_1&, NPC_1&, NF&, u#, BF#
'
    NPI_1 = UBound(Pi)
    NPC_1 = UBound(Pc)
    'NF = Prodotto(NPI_1)
'
    For I = 0 To NPC_1
        u = CDbl(I) / CDbl(NPC_1)
        Pc(I).x = 0#
        Pc(I).y = 0#
        'Pc(I).z = 0#
        For K = 0 To NPI_1
            'BF = NF * (u ^ K) * ((1 - u) ^ (NPI_1 - K)) / (Prodotto(K) * Prodotto(NPI_1 - K))
            BF = Prodotto(NPI_1, K + 1) * (u ^ K) * ((1 - u) ^ (NPI_1 - K)) / Prodotto(NPI_1 - K)
            Pc(I).x = Pc(I).x + Pi(K).x * BF
            Pc(I).y = Pc(I).y + Pi(K).y * BF
            'Zu = Zu + Pi(K).z * BF
        Next K
    Next I
'
'
'
End Sub
Public Sub Bezier(Pi() As CadPoint, Pc() As CadPoint)
'
'   Ritorna, nel vettore Pc(), i valori della curva di Bezier.
'   La curva e' calcolata in modo parametrico (0 <= u < 1)
'   con il valore 0 di u corrispondente a Pc(0);
'   Attenzione: il punto Pc(NPC_1), corrispondente al valore u = 1,
'               non puo' essere calcolato.
'
'   Parametri:
'       Pi(0 to NPI - 1):   Vettore dei punti, dati, da
'                           approssimare.
'       Pc(0 to NPC - 1):   Vettore dei punti, calcolati,
'                           della curva approssimante.
'
   Dim I&, K&, NPI_1&, NPC_1&
   Dim u#, u_1#, ue#, u_1e#, BF#
'
    NPI_1 = UBound(Pi) ' N. di punti da approssimare - 1.
    NPC_1 = UBound(Pc) ' N. di punti sulla curva - 1.
'
    ' La curva inizia sempre da Pi(0) -> u = 0:
    Pc(0).x = Pi(0).x
    Pc(0).y = Pi(0).y
    'Pc(0).z = Pi(0).z
'
    For I = 1 To NPC_1 - 1
        u = CDbl(I) / CDbl(NPC_1)
        ue = 1#
        u_1 = 1# - u
        u_1e = u_1 ^ NPI_1
'
        Pc(I).x = 0#
        Pc(I).y = 0#
        'Pc(I).z = 0#
        For K = 0 To NPI_1
            BF = Prodotto(NPI_1, K + 1) * ue * u_1e / Prodotto(NPI_1 - K)
            Pc(I).x = Pc(I).x + Pi(K).x * BF
            Pc(I).y = Pc(I).y + Pi(K).y * BF
            'Pc(I).z = Pc(I).z + Pi(K).z * BF
'
            ue = ue * u
            u_1e = u_1e / u_1
        Next K
    Next I
'
    ' La curva finisce sempre su Pi(NPI_1) -> u = 1:
    Pc(NPC_1).x = Pi(NPI_1).x
    Pc(NPC_1).y = Pi(NPI_1).y
    'Pc(NPC_1).z = Pi(NPI_1).z
'
'
'
End Sub

Public Sub Bezier_P(Pi() As CadPoint, Pc() As CadPoint)
'
'   Ritorna, nel vettore Pc(), i valori della curva di Bezier calcolata
'   al valore u (0 <= u < 1). La curva e' calcolata in modo
'   parametrico con il valore 0 di u corrispondente a Pc(0);
'   Attenzione: il punto Pc(NPC_1), corrispondente al valore u = 1,
'               non puo' essere calcolato.
'
'   Questo algoritmo (tratto da una pubblicazione di P. Bourke
'   e tradotto dal C) e' particolarmente interessante, in quanto
'   evita l' uso dei fattoriali della forma normale.
'
    Dim K&, I&, KN&, NPI_1&, NPC_1&, NN&, NKN&
    Dim u#, uk#, unk#, Blend#
'
    NPI_1 = UBound(Pi)
    NPC_1 = UBound(Pc)
'
    For I = 0 To NPC_1 - 1
        u = CDbl(I) / CDbl(NPC_1)
        uk = 1#
        unk = (1# - u) ^ NPI_1
'
        Pc(I).x = 0#
        Pc(I).y = 0#
        'Pc(I).z = 0#
'
        For K = 0 To NPI_1
            NN = NPI_1
            KN = K
            NKN = NPI_1 - K
            Blend = uk * unk
            uk = uk * u
            unk = unk / (1# - u)
            Do While NN >= 1
                Blend = Blend * CDbl(NN)
                NN = NN - 1
                If KN > 1 Then
                    Blend = Blend / CDbl(KN)
                    KN = KN - 1
                End If
                If NKN > 1 Then
                    Blend = Blend / CDbl(NKN)
                    NKN = NKN - 1
                End If
            Loop
'
            Pc(I).x = Pc(I).x + Pi(K).x * Blend
            Pc(I).y = Pc(I).y + Pi(K).y * Blend
            'Pc(I).z = Pc(I).z + Pi(K).z * Blend
        Next K
    Next I
'
    ' La curva finisce sempre su Pi(NPI_1) -> u = 1:
    Pc(NPC_1).x = Pi(NPI_1).x
    Pc(NPC_1).y = Pi(NPI_1).y
    'Pc(NPC_1).z = Pi(NPI_1).z
'
'
'
End Sub

Private Function Prodotto(ByVal N2&, Optional ByVal N1& = 2) As Double
'
'   Ritorna il prodotto dei numeri da N1& a N2& (N2& >= N1&).
'   Se N1& manca, ritorna il Fattoriale di N2&:
'
    Dim f#, I&
'
    f = 1#
    For I = N1 To N2
        f = f * CDbl(I)
    Next I
'
    Prodotto = f
'
'
'
End Function


Public Sub B_Spline(Pi() As CadPoint, ByVal NK&, Pc() As CadPoint)
'
'   Ritorna, nel vettore Pc(), i valori della curva B-Spline.
'   La curva e' calcolata in modo parametrico (0 <= u <= 1)
'   con il valore 0 di u corrispondente a Pc(0) ed il valore
'   1 corrispondente a Pc(NPC_1).
'
'   Parametri:
'       Pi(0 to NPI - 1):   Vettore dei punti, dati, da
'                           approssimare.
'       Pc(0 to NPC - 1):   Vettore dei punti, calcolati,
'                           della curva approssimante.
'       NK:                 Numero di nodi della curva
'                           approssimante:
'                           NK = 2    -> segmenti di retta.
'                           NK = 3    -> curve quadratiche.
'                           ..   .       ..................
'                           NK = NPI  -> splines di Bezier.

    Dim NPI_1&, NPC_1&, I&, J&, tmax#, u#, ut#, Eps#, bn#()
'
    NPI_1 = UBound(Pi)  ' N. di punti da approssimare - 1.
    NPC_1 = UBound(Pc)  ' N. di punti sulla curva - 1.
    Eps = 0.0000001
    tmax = NPI_1 - NK + 2
'
    ' La curva inizia sempre da Pi(0) -> u = 0:
    Pc(0).x = Pi(0).x
    Pc(0).y = Pi(0).y
'
    For I = 1 To NPC_1 - 1
        u = CDbl(I) / CDbl(NPC_1)
        ut = u * tmax
        If Abs(ut - CDbl(NPI_1 + NK - 2)) <= Eps Then
            Pc(I).x = Pi(NPI_1).x
            Pc(I).y = Pi(NPI_1).y
        Else
            Call B_Basis(NPI_1, ut, NK, bn())
            Pc(I).x = 0#
            Pc(I).y = 0#
            For J = 0 To NPI_1
                Pc(I).x = Pc(I).x + bn(J) * Pi(J).x
                Pc(I).y = Pc(I).y + bn(J) * Pi(J).y
            Next J
        End If
    Next I
'
    ' La curva finisce sempre su Pi(NPI_1) -> u = 1:
    Pc(NPC_1).x = Pi(NPI_1).x
    Pc(NPC_1).y = Pi(NPI_1).y
'
'
'
End Sub

Private Sub B_Basis(ByVal NPI_1&, ByVal ut#, ByVal K&, bn#())
'
'   Compute the basis function (also called weight)
'   for the B-Spline approximation curve:
'
    Dim NT&, I&, J&
    Dim b0#, b1#, bl0#, bl1#, bu0#, bu1#
    ReDim bn#(0 To NPI_1 + 1), bn0#(0 To NPI_1 + 1), t#(0 To NPI_1 + K + 1)
'
    NT = NPI_1 + K + 1
    For I = 0 To NT
        If (I < K) Then t(I) = 0#
        If ((I >= K) And (I <= NPI_1)) Then t(I) = CDbl(I - K + 1)
        If (I > NPI_1) Then t(I) = CDbl(NPI_1 - K + 2)
    Next I
    For I = 0 To NPI_1
        bn0(I) = 0#
        If ((ut >= t(I)) And (ut < t(I + 1))) Then bn0(I) = 1#
        If ((t(I) = 0#) And (t(I + 1) = 0#)) Then bn0(I) = 0#
    Next I
'
    For J = 2 To K
        For I = 0 To NPI_1
            bu0 = (ut - t(I)) * bn0(I)
            bl0 = t(I + J - 1) - t(I)
            If (bl0 = 0#) Then
                b0 = 0#
            Else
                b0 = bu0 / bl0
            End If
            bu1 = (t(I + J) - ut) * bn0(I + 1)
            bl1 = t(I + J) - t(I + 1)
            If (bl1 = 0#) Then
                b1 = 0#
            Else
                b1 = bu1 / bl1
            End If
            bn(I) = b0 + b1
        Next I
        For I = 0 To NPI_1
            bn0(I) = bn(I)
        Next I
    Next J
'
'
'
End Sub

Public Sub C_Spline(Pi() As CadPoint, Pc() As CadPoint)
'
'   Ritorna, nel vettore Pc(), i valori della curva C-Spline.
'   La curva e' calcolata in modo parametrico (0 <= u <= 1)
'   con il valore 0 di u corrispondente a Pc(0) ed il valore
'   1 corrispondente a Pc(NPC_1).
'
'   Parametri:
'       Pi(0 to NPI - 1):   Vettore dei punti, dati, da
'                           interpolare.
'       Pc(0 to NPC - 1):   Vettore dei punti, calcolati,
'                           della curva interpolante.
'
    Dim NPI_1&, NPC_1&, I&, J&
    Dim u#, ui#, uui#
    Dim cof() As CadPoint
'
    NPI_1 = UBound(Pi)      ' N. di punti da interpolare - 1.
    NPC_1 = UBound(Pc)      ' N. di punti sulla curva - 1.
'
    Call Find_CCof(Pi(), NPI_1 + 1, cof())
'
    ' La curva inizia sempre da Pi(0) -> u = 0:
    Pc(0).x = Pi(0).x
    Pc(0).y = Pi(0).y
'
    For I = 1 To NPC_1 - 1
        u = CDbl(I) / CDbl(NPC_1)
        J = Int(u * CDbl(NPI_1)) + 1
        If (J > (NPI_1)) Then J = NPI_1
'
        ui = CDbl(J - 1) / CDbl(NPI_1)
        uui = u - ui
'
        Pc(I).x = cof(4, J).x * uui ^ 3 + cof(3, J).x * uui ^ 2 + cof(2, J).x * uui + cof(1, J).x
        Pc(I).y = cof(4, J).y * uui ^ 3 + cof(3, J).y * uui ^ 2 + cof(2, J).y * uui + cof(1, J).y
    Next I
'
    ' La curva finisce sempre su Pi(NPI_1) -> u = 1:
    Pc(NPC_1).x = Pi(NPI_1).x
    Pc(NPC_1).y = Pi(NPI_1).y
'
'
'
End Sub
Public Sub T_Spline(Pi() As CadPoint, ByVal VZ&, Pc() As CadPoint)
'
'   Ritorna, nel vettore Pc(), i valori della curva T-Spline.
'   La curva e' calcolata in modo parametrico (0 <= u <= 1)
'   con il valore 0 di u corrispondente a Pc(0) ed il valore
'   1 corrispondente a Pc(NPC_1).
'
'   Parametri:
'       Pi(0 to NPI - 1):   Vettore dei punti, dati, da
'                           interpolare.
'       Pc(0 to NPC - 1):   Vettore dei punti, calcolati,
'                           della curva interpolante.
'       VZ:                 Valore di tensione della curva
'                           (1 <= VZ <= 100): valori grandi
'                           di VZ appiattiscono la curva.
'
    Dim NPI_1&, NPC_1&, I&, J&
    Dim h#, z#, z2i#, szh#, u#, u0#, u1#, du1#, du0#
    Dim s() As CadPoint
'
    NPI_1 = UBound(Pi)      ' N. di punti da interpolare - 1.
    NPC_1 = UBound(Pc)      ' N. di punti sulla curva - 1.
    z = CDbl(VZ)
'
    Call Find_TCof(Pi(), NPI_1 + 1, s(), z)
'
    ' La curva inizia sempre da Pi(0) -> u = 0:
    Pc(0).x = Pi(0).x
    Pc(0).y = Pi(0).y
'
    h = 1# / CDbl(NPI_1)
    szh = Sinh(z * h)
    z2i = 1# / z / z
    For I = 1 To NPC_1 - 1
        u = CDbl(I) / CDbl(NPC_1)
        J = Int(u * CDbl(NPI_1)) + 1
        If (J > (NPI_1)) Then J = NPI_1
'
        u0 = CDbl(J - 1) / CDbl(NPI_1)
        u1 = CDbl(J) / CDbl(NPI_1)
        du1 = u1 - u
        du0 = u - u0
'
        Pc(I).x = s(J).x * z2i * Sinh(z * du1) / szh + (Pi(J - 1).x - s(J).x * z2i) * du1 / h
        Pc(I).x = Pc(I).x + s(J + 1).x * z2i * Sinh(z * du0) / szh + (Pi(J).x - s(J + 1).x * z2i) * du0 / h
    
        Pc(I).y = s(J).y * z2i * Sinh(z * du1) / szh + (Pi(J - 1).y - s(J).y * z2i) * du1 / h
        Pc(I).y = Pc(I).y + s(J + 1).y * z2i * Sinh(z * du0) / szh + (Pi(J).y - s(J + 1).y * z2i) * du0 / h
    Next I
'
    ' La curva finisce sempre su Pi(NPI_1) -> u = 1:
    Pc(NPC_1).x = Pi(NPI_1).x
    Pc(NPC_1).y = Pi(NPI_1).y
'
'
'
End Sub

Private Sub Find_TCof(Pi() As CadPoint, ByVal NPI&, s() As CadPoint, ByVal z#)
'
'   Find the coefficients of the T-Spline
'   using constant interval:
'
    Dim I&, h#, a0#, b0#, zh#, z2#
'
    ReDim s(1 To NPI) As CadPoint, f(1 To NPI) As CadPoint
    ReDim a(1 To NPI) As Double, b(1 To NPI) As Double, c(1 To NPI) As Double
'
    h = 1# / CDbl(NPI - 1)
    zh = z * h
    a0 = 1# / h - z / Sinh(zh)
    b0 = z * 2# * Cosh(zh) / Sinh(zh) - 2# / h
    For I = 1 To NPI - 2
        a(I) = a0
        b(I) = b0
        c(I) = a0
    Next I
'
    z2 = z * z / h
    For I = 1 To NPI - 2
        f(I).x = (Pi(I + 1).x - 2# * Pi(I).x + Pi(I - 1).x) * z2
        f(I).y = (Pi(I + 1).y - 2# * Pi(I).y + Pi(I - 1).y) * z2
    Next I
'
    Call TRIDAG(a(), b(), c(), f(), s(), NPI - 2)
    For I = 1 To NPI - 2
        s(NPI - I).x = s(NPI - 1 - I).x
        s(NPI - I).y = s(NPI - 1 - I).y
    Next I
'
    s(1).x = 0#
    s(NPI).x = 0#
    s(1).y = 0#
    s(NPI).y = 0#
'
'
'
End Sub
Private Sub Find_CCof(Pi() As CadPoint, ByVal NPI&, cof() As CadPoint)
'
'   Find the coefficients of the cubic spline
'   using constant interval parameterization:
'
    Dim I&, h#
'
    ReDim s(1 To NPI) As CadPoint, f(1 To NPI) As CadPoint, cof(1 To 4, 1 To NPI) As CadPoint
    ReDim a(1 To NPI) As Double, b(1 To NPI) As Double, c(1 To NPI) As Double
'
    h = 1# / CDbl(NPI - 1)
    For I = 1 To NPI - 2
        a(I) = 1#
        b(I) = 4#
        c(I) = 1#
    Next I
'
    For I = 1 To NPI - 2
        f(I).x = 6# * (Pi(I + 1).x - 2# * Pi(I).x + Pi(I - 1).x) / h / h
        f(I).y = 6# * (Pi(I + 1).y - 2# * Pi(I).y + Pi(I - 1).y) / h / h
    Next I
'
    Call TRIDAG(a(), b(), c(), f(), s(), NPI - 2)
    For I = 1 To NPI - 2
        s(NPI - I).x = s(NPI - 1 - I).x
        s(NPI - I).y = s(NPI - 1 - I).y
    Next I
'
    s(1).x = 0#
    s(NPI).x = 0#
    s(1).y = 0#
    s(NPI).y = 0#
    For I = 1 To NPI - 1
        cof(4, I).x = (s(I + 1).x - s(I).x) / 6# / h
        cof(4, I).y = (s(I + 1).y - s(I).y) / 6# / h
        cof(3, I).x = s(I).x / 2#
        cof(3, I).y = s(I).y / 2#
        cof(2, I).x = (Pi(I).x - Pi(I - 1).x) / h - (2# * s(I).x + s(I + 1).x) * h / 6#
        cof(2, I).y = (Pi(I).y - Pi(I - 1).y) / h - (2# * s(I).y + s(I + 1).y) * h / 6#
        cof(1, I).x = Pi(I - 1).x
        cof(1, I).y = Pi(I - 1).y
    Next I
'
'
'
End Sub

Private Sub TRIDAG(a#(), b#(), c#(), f() As CadPoint, s() As CadPoint, ByVal NPI_2&)
'
'   Solves the tridiagonal linear system of equations:
'
    Dim J&, bet#
    ReDim gam#(1 To NPI_2)
'
    If b(1) = 0 Then Exit Sub
'
    bet = b(1)
    s(1).x = f(1).x / bet
    s(1).y = f(1).y / bet
    For J = 2 To NPI_2
        gam(J) = c(J - 1) / bet
        bet = b(J) - a(J) * gam(J)
        If (bet = 0) Then Exit Sub
        s(J).x = (f(J).x - a(J) * s(J - 1).x) / bet
        s(J).y = (f(J).y - a(J) * s(J - 1).y) / bet
    Next J
'
    For J = NPI_2 - 1 To 1 Step -1
        s(J).x = s(J).x - gam(J + 1) * s(J + 1).x
        s(J).y = s(J).y - gam(J + 1) * s(J + 1).y
    Next J
'
'
'
End Sub

Private Function Cosh(ByVal z As Double) As Double
'
'   Ritorna il coseno iperbolico di z#:
'
    Cosh = (Exp(z) + Exp(-z)) / 2#
'
'
'
End Function
Private Function Sinh(ByVal z As Double) As Double
'
'   Ritorna il seno iperbolico di z#:
'
    Sinh = (Exp(z) - Exp(-z)) / 2#
'
'
'
End Function


Function INTERPOLA(matriz, coldatos, plazobus)

r = matriz.Rows.Count

For i = 1 To r - 1
    If plazobus >= matriz(i, 1) And plazobus <= matriz(i + 1, 1) Then
        m = (matriz(i + 1, coldatos) - matriz(i, coldatos)) / (matriz(i + 1, 1) - matriz(i, 1))
        temp = matriz(i, coldatos) + m * (plazobus - matriz(i, 1))
    End If
Next i

If plazobus < matriz(1, 1) Then
    temp = matriz(1, coldatos)
ElseIf plazobus > matriz(r, 1) Then
    temp = matriz(r, coldatos)
End If

INTERPOLA = temp

End Function

'#####################################################################################################################################

Function INTERPOLA2(matriz, coldatos, plazobus)

r = matriz.Rows.Count

For i = 1 To r - 1
    If plazobus >= matriz(i, 1) And plazobus <= matriz(i + 1, 1) Then
        m = (matriz(i + 1, coldatos) - matriz(i, coldatos)) / (matriz(i + 1, 1) - matriz(i, 1))
        temp = matriz(i, coldatos) + m * (plazobus - matriz(i, 1))
    End If
Next i

If plazobus < matriz(1, 1) Then
    m = (matriz(2, coldatos) - matriz(1, coldatos)) / (matriz(2, 1) - matriz(1, 1))
    temp = matriz(1, coldatos) + m * (plazobus - matriz(1, 1))
ElseIf plazobus > matriz(r, 1) Then
    m = (matriz(r, coldatos) - matriz(r - 1, coldatos)) / (matriz(r, 1) - matriz(r - 1, 1))
        temp = matriz(r - 1, coldatos) + m * (plazobus - matriz(r - 1, 1))
    'temp = matriz(r, coldatos)
End If

INTERPOLA2 = temp

End Function

'#####################################################################################################################################

Function OPT_BT_Ame_EQUITY(Tipo, Subyacente, Strike, Plazo, Volatilidad, Tasa, Dividendo, Tset, Optional Nodos As Integer, Optional Base = 360, Optional Opcion = 1) As Double

Dim Prob, Up, Down, Cc  As Integer, fd As Double, dt As Double, n As Integer, i As Integer
Dim j As Integer, V1() As Double, V0() As Double, V2() As Double

Plazo = Plazo / Base
Tset = Tset / Base


If Nodos > 0 Then
    n = Nodos
Else
    n = 500
End If

Cc = 1

If Tipo = "call" Or Tipo = "Call" Or Tipo = "C" Or Tipo = "c" Then
    Cc = -1
End If


ReDim V1(0 To n) As Double
ReDim V0(0 To n) As Double
ReDim V2(0 To n) As Double

dt = Plazo / n
Up = Exp(Volatilidad * Sqr(dt))
Down = 1 / Up
Prob = (Exp((Tasa - Dividendo) * Tset / n) - Down) / (Up - Down)
fd = Exp(-Tasa * Tset / n)

'Tree Subyacente
For j = 0 To n
    V1(j) = Application.WorksheetFunction.Max(-Cc * Subyacente * (Up ^ j) * (Down ^ (n - j)) + Cc * Strike, 0)
Next j

'Tree Prima
For i = n - 1 To 0 Step -1
    For j = 0 To i
        V0(j) = Application.WorksheetFunction.Max((Prob * V1(j + 1) + (1 - Prob) * V1(j)) * fd, -Cc * Subyacente * (Up ^ j) * (Down ^ (i - j)) + Cc * Strike)
        V2(j) = V1(j)
    Next j
    For j = 0 To i
        V1(j) = V0(j)
    Next j
Next i

'Si
Select Case Opcion
 Case Is = 1
    OPT_BT_Ame_EQUITY = V0(0)
Case Is = 2
    OPT_BT_Ame_EQUITY = V1(1) 'prima superior
Case Is = 3
    OPT_BT_Ame_EQUITY = V2(0) 'prima inferior
Case Else
    OPT_BT_Ame_EQUITY = Up
 End Select
 
 
Erase V1()
Erase V2()
Erase V0()
 
 
 
End Function

'#####################################################################################################################################

Function Normal_Acum(X)
Normal_Acum = Application.WorksheetFunction.NormSDist(X)
End Function
Function Normal_Densidad(X) As Double
    Normal_Densidad = Exp(-X ^ 2 / 2) / Sqr(2 * Application.WorksheetFunction.Pi())
End Function
Function Tnom_Tcont(Tasa, Plazo, Optional Base = 360)

Tnom_Tcont = Log(1 + Tasa * Plazo / Base) * Base / Plazo

End Function

'#####################################################################################################################################

Function Opc_FWDDelta(Tipo As String, Subyacente As Double, Strike As Double, Tasa1 As Double, _
Tasa2 As Double, Plazo_mat As Double, Plazo_set As Double, Volatilidad As Double) As Double

Dim d1, d2, Nd1, Nd2, Nmd1, Nmd2, N1d1, N2d2
t_mat = Plazo_mat / 365
t_set = Plazo_set / 365
If Plazo_mat > 0 And Volatilidad > 0 Then
    d1 = (Log(Subyacente / Strike) + (Tasa1 - Tasa2) * t_set + ((Volatilidad ^ 2) / 2) * t_mat) / (Volatilidad * Sqr(t_mat))
    d2 = d1 - Volatilidad * Sqr(t_mat)
    Nd1 = Normal_Acum(d1)
    N1d1 = Normal_Densidad(d1)
    Nd2 = Normal_Acum(d2)
    Nmd1 = Normal_Acum(-1 * d1)
    Nmd2 = Normal_Acum(-1 * d2)
        If Tipo = "C" Or Tipo = "c" Or Tipo = "Call" Or Tipo = "call" Then
            Opc_FWDDelta = Nd1
        Else
            Opc_FWDDelta = (Nd1 - 1)
        End If
Else
    Opc_FWDDelta = 0
End If

End Function

'#####################################################################################################################################

Function interpola_Vol(Tipo As String, Subyacente As Double, Strike As Double, Tasa1 As Double, _
Tasa2 As Double, Plazo_mat As Double, Plazo_set As Double, VolATM As Double, matriz, coldatos, Result, Basis) As Double

Dim vol As Double
Dim conBasis As Double

If Basis = 365 Then
    conBasis = 1
Else
    conBasis = Sqr(365) / Sqr(Basis)
End If

Epsilon = 1

If Tipo = "C" Or Tipo = "c" Or Tipo = "Call" Or Tipo = "call" Then
    Delta = Opc_FWDDelta(Tipo, Subyacente, Strike, Tasa1, Tasa2, Plazo_mat, Plazo_set, VolATM / 100)
Else
    Delta = -1 * Opc_FWDDelta(Tipo, Subyacente, Strike, Tasa1, Tasa2, Plazo_mat, Plazo_set, VolATM / 100)
End If

Do

If Epsilon > 2 Then
    vol = INTERPOLA2(matriz, coldatos, Delta) * conBasis
Else
    vol = INTERPOLA(matriz, coldatos, Delta) * conBasis
End If
If Tipo = "C" Or Tipo = "c" Or Tipo = "Call" Or Tipo = "call" Then
    deltaE = Opc_FWDDelta(Tipo, Subyacente, Strike, Tasa1, Tasa2, Plazo_mat, Plazo_set, vol / 100)
Else
    deltaE = -1 * Opc_FWDDelta(Tipo, Subyacente, Strike, Tasa1, Tasa2, Plazo_mat, Plazo_set, vol / 100)
End If

Delta = deltaE
Epsilon = Epsilon + 1
Loop While Epsilon < 30
    
Select Case Result
Case 1 'Volatilidad
    interpola_Vol = vol
Case 2 'Delta
    interpola_Vol = Delta
    End Select
End Function

'#####################################################################################################################################
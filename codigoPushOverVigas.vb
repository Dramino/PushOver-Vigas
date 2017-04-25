Option Base 1
Sub MatrizRigidez()
Dim numVigas As Integer
Dim i, j, barra, numGDL, nodos As Integer
Dim gdl
Dim L, EI
Dim rigidez, pocisionVigas
Dim matrizGlobal
Dim apoyo
Dim contador As Integer
Dim mFuerzas, celdas(2)
dim Fn
dim deformaciones
Sheets("Datos").Activate
'cuenta el número de nodos existente
[Nodo].Select
Do Until ActiveCell.Offset(nodos, 0) = ""
    nodos = nodos + 1
Loop
nodos = nodos - 1
'crea las matricez del tamaño necesario
numVigas = nodos - 1
ReDim EI(numVigas), L(numVigas)
ReDim rigidezVigas(numVigas, 4, 4)
ReDim pocisionVigas(numVigas, 4)
ReDim apoyo(numVigas, 2)
[elemento].Select
' obtiene los grados de libertad de cada barra
contador = 1
For i = 1 To numVigas
    EI(i) = ActiveCell.Offset(i, 2)
    If EI(i) = "" Then EI(i) = 1
    L(i) = ActiveCell.Offset(i, 1)
    If i = 1 Then
    
        Select Case ActiveCell.Offset(i, 6)
        Case ""
            apoyo(i, 1) = 0
            pocisionVigas(i, 1) = contador
            pocisionVigas(i, 2) = contador + 1
            contador = contador + 2
        Case "Empotrado"
            apoyo(i, 1) = 1
            pocisionVigas(i, 1) = 0
            pocisionVigas(i, 2) = 0
    
        Case "Articulado"
            apoyo(i, 1) = 2
            pocisionVigas(i, 1) = 0
            pocisionVigas(i, 2) = contador
            contador = contador + 1
    
        End Select
    End If
    
Select Case ActiveCell.Offset(i + 1, 6)
    Case ""
        apoyo(i, 1) = 0
        pocisionVigas(i, 3) = contador
        pocisionVigas(i, 4) = contador + 1
        contador = contador + 2
    Case "Empotrado"
        apoyo(i, 1) = 1
        pocisionVigas(i, 3) = 0
        pocisionVigas(i, 4) = 0

    Case "Articulado"
        apoyo(i, 1) = 2
        pocisionVigas(i, 3) = 0
        pocisionVigas(i, 4) = contador
        contador = contador + 1

    End Select
    If i < numVigas Then
        pocisionVigas(i + 1, 1) = pocisionVigas(i, 3)
        pocisionVigas(i + 1, 2) = pocisionVigas(i, 4)
    End If

Next
For i = 1 To numVigas
    rigidezVigas(i, 1, 1) = 12 * EI(i) / L(i) ^ 3
    rigidezVigas(i, 1, 2) = 6 * EI(i) / L(i) ^ 2
    rigidezVigas(i, 1, 3) = -12 * EI(i) / L(i) ^ 3
    rigidezVigas(i, 1, 4) = 6 * EI(i) / L(i) ^ 2
    
    rigidezVigas(i, 2, 1) = 6 * EI(i) / L(i) ^ 2
    rigidezVigas(i, 2, 2) = 4 * EI(i) / L(i)
    rigidezVigas(i, 2, 3) = -6 * EI(i) / L(i) ^ 2
    rigidezVigas(i, 2, 4) = 2 * EI(i) / L(i)

    rigidezVigas(i, 3, 1) = -12 * EI(i) / L(i) ^ 3
    rigidezVigas(i, 3, 2) = -6 * EI(i) / L(i) ^ 2
    rigidezVigas(i, 3, 3) = 12 * EI(i) / L(i) ^ 3
    rigidezVigas(i, 3, 4) = -6 * EI(i) / L(i) ^ 2
    
    rigidezVigas(i, 4, 1) = 6 * EI(i) / L(i) ^ 2
    rigidezVigas(i, 4, 2) = 2 * EI(i) / L(i)
    rigidezVigas(i, 4, 3) = -6 * EI(i) / L(i) ^ 2
    rigidezVigas(i, 4, 4) = 4 * EI(i) / L(i)
Next
celdas(1) = 8
celdas(2) = 3
gdl = gradosLibertad(celdas, nodos)
numGDL = UBound(gdl)
ReDim matrizGlobal(numGDL, numGDL)
For barra = 1 To numVigas
    For i = 1 To 4
        For j = 1 To 4
            If pocisionVigas(barra, j) <> 0 And pocisionVigas(barra, i) <> 0 Then
                matrizGlobal(pocisionVigas(barra, i), pocisionVigas(barra, j)) = matrizGlobal(pocisionVigas(barra, i), pocisionVigas(barra, j)) + rigidezVigas(barra, i, j)
            End If
        Next
        
    Next

Next
celdas(1) = 8
celdas(2) = 3

mFuerzas = fuerzas(celdas, nodos, numGDL)
Set Fn = Application.WorksheetFunction
'Deformaciones
deformaciones=Fn.Mmult(Fn.Minverse(matrizGlobal),mFuerzas)
End Sub

Function gradosLibertad(valorCeldas, valorNodos)
    Dim nodos, i, gdl As Integer
    Dim X, Y As Integer
    Dim condicion As String
    Dim resultado
    Dim contador As Integer
    contador = 0
    nodos = valorNodos
    X = valorCeldas(1)
    Y = valorCeldas(2)
    For i = 1 To nodos
        condicion = Cells(Y + i, X - 1)
        Select Case condicion
            Case ""
                contador = contador + 2
            Case "Articulado"
                contador = contador + 1
            Case "Empotrado"
        
        End Select
    Next
    ReDim resultado(contador)
    contador = 1
    For i = 1 To nodos
        condicion = Cells(Y + i, X - 1)
        Select Case condicion
            Case ""
                resultado(contador) = 2 * i - 1
                resultado(contador + 1) = 2 * i
                contador = contador + 2
            Case "Articulado"
                contador = contador + 1
                resultado(contador + 1) = 2 * i
            Case "Empotrado"
        
        End Select
    Next
    gradosLibertad = resultado
End Function

Function fuerzas(valorCeldas, valorNodos, valorGDL)
    Dim nodos, i, gdl As Integer
    Dim X, Y As Integer
    Dim condicion As String
    Dim resultado
    Dim contador As Integer
    gdl = valorGDL
    ReDim resultado(gdl, 1)
    contador = 1
    nodos = valorNodos
    X = valorCeldas(1)
    Y = valorCeldas(2)
    For i = 1 To nodos
        condicion = Cells(Y + i, X - 1)
        Select Case condicion
            Case ""
                resultado(2 * contador - 1, 1) = Cells(Y + i, X)
                resultado(2 * contador, 1) = Cells(Y + i, X + 1)
                contador = contador + 1
            Case "Articulado"
                resultado(2 * contador, 1) = Cells(Y + i, X + 1)
                contador = contador + 1
            Case "Empotrado"
        
        End Select
    Next
    fuerzas = resultado
End Function


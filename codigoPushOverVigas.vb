Dim numVigas As Integer
Dim i, j, barra, gdl As Integer
Dim L, EI
Dim rigidez, pocisionVigas
Dim matrizGlobal
Dim apoyo
Dim contador As Integer
Sheets("Datos").Activate

[Nodo].Select
Do Until ActiveCell.Offset(gdl, 0) = ""
    gdl = gdl + 1
Loop
gdl = gdl - 1

numVigas = gdl - 1
ReDim EI(numVigas), L(numVigas)
ReDim rigidezVigas(numVigas, 4, 4)
ReDim pocisionVigas(numVigas, 4)
ReDim apoyo(numVigas, 2)
[elemento].Select

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
ReDim matrizGlobal(gdl, gdl)

For barra = 1 To numVigas
    For i = 1 To 4
        For j = 1 To 4
            If pocisionVigas(barra, j) <> 0 And pocisionVigas(barra, i) <> 0 Then
                matrizGlobal(pocisionVigas(barra, i), pocisionVigas(barra, j)) = matrizGlobal(pocisionVigas(barra, i), pocisionVigas(barra, j)) + rigidezVigas(barra, i, j)
            End If
        Next
        
    Next

Next


End Sub
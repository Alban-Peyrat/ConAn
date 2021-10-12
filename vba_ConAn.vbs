Option Explicit

Dim CAX As String
Dim bib As String
Dim colIndex As String
Dim rowIndex As Integer
Dim inputPPN As String
Dim folderPath As String
Dim fileName As String
Dim mainWorkBook As Workbook
Dim exportAlma As Workbook
Sub read_Alma_Data()
'Originaux : https://www.mrexcel.com/board/threads/vba-selecting-only-rows-with-data-in-a-range.752110/
'et : https://stackoverflow.com/questions/41725730/how-to-paste-values-and-keep-source-formatting
'et : https://www.ozgrid.com/forum/index.php?thread/143595-deselect-cells-that-have-been-copied/

'heu va falloir couper en deux la partie propre à mn script et la partie CA2
'Cf read_Sudoc_Data dans ConStance


'Ouvre le ficheir export_alma et chope le nb de lignes
    Dim nbRow As Integer
    
    Workbooks.Open fileName:=folderPath & "\export_alma_ConAn.xlsx"
    Set exportAlma = Workbooks("export_alma_ConAn.xlsx")
    
    nbRow = Cells(Rows.count, "K").End(xlUp).Row

    'Récupère les données
    For rowIndex = 2 To nbRow
    
            'Lance le bon script
            Select Case CAX
                Case "[CA2]"
                    getUBAGE
                Case "[CA3]"
                    getEdition
                Case Else
                    MsgBox "La valeur entrée en H2 ne devrait pas être possible"
            End Select
       
    Next
    Workbooks("export_alma_ConAn.xlsx").Close
'Permet de remettre rowIndex au niveau de nbRow
    rowIndex = rowIndex - 1
    
End Sub
Sub getUBAGE()

    Dim PPN As String
    Dim expEx As Variant, nbEx As Integer, holding As Variant, posPar As Integer, posEx As Integer, ll As Integer, nbExReW As String
    Dim expEd As Variant, expEdOcc As Variant, posCol As Integer, posSemi As Integer, annee As String, anneeReW As String, anneeOutput As String
    Dim antiWhile As Integer

    PPN = exportAlma.Worksheets("Results").Cells(rowIndex, 10).Value
    expEx = exportAlma.Worksheets("Results").Cells(rowIndex, 11).Value
    expEd = exportAlma.Worksheets("Results").Cells(rowIndex, 1).Value

'Clean le PPN
    PPN = Mid(PPN, InStr(PPN, "(PPN)"), 14)
    
'Récupère l'année
    'PosePB avec les 21X qui ont entre paranthèse le pays
    'expEd = Split(expEd, "(")
    'expEd = expEd(UBound(expEd))
    expEd = Mid(expEd, InStr(expEd, "("), Len(expEd))
    
    expEd = Split(expEd, ";")
    anneeOutput = ""
    For Each expEdOcc In expEd
        posCol = InStrRev(expEdOcc, ",")
'Le > 4 sert à empêcher les virgules mal cataloguées en fin $d
        If posCol > 0 And Len(expEdOcc) - posCol > 4 Then
            expEdOcc = Mid(expEdOcc, posCol, Len(expEdOcc))
        End If
        
'ConStance CS 4
        annee = Replace(expEdOcc, " ", "")
        annee = Replace(annee, "DL", "")
        annee = Replace(annee, "C", "")
        annee = Replace(annee, "cop", "")
        annee = Replace(annee, "P", "")
        annee = Replace(annee, ".", "")
        annee = Replace(annee, ",", "")
        annee = Replace(annee, "-", "")
        annee = Replace(annee, "?", "")
        annee = Replace(annee, "(", "")
        annee = Replace(annee, ")", "")
        'Si après cette première vague, du texte est encore présent, une détection caratère à caractère est effectuée
        If IsNumeric(annee) = False Then
            For ll = 1 To Len(annee)
                If IsNumeric(Mid(annee, ll, 1)) = True Then
                    anneeReW = anneeReW & Mid(annee, ll, 1)
                    If Len(anneeReW) = 8 Then
                        If Left(anneeReW, 4) < Right(anneeReW, 4) Then
                           anneeReW = Right(anneeReW, 4)
                        Else
                            anneeReW = Left(anneeReW, 4)
                        End If
                    End If
                End If
            Next
            'Donne 0 à la valeur annnee pour éviter de poser problèmes plus tard
            If anneeReW <> "" Then
                annee = anneeReW
            Else
                annee = "0"
            End If
        End If
        
        'Vérifie si l'année fait 4 chiffres (0 exclus), sinon essaye de voir si les 4 premiers chiffres font une année entre 1000 et 9999, sinon laisse vide
        If CDbl(annee) > 2030 Or CDbl(annee) < 1900 Then
            If CLng(Right(annee, 4)) < 2030 And CLng(Right(annee, 4)) > 1900 Then
                annee = Right(annee, 4)
            ElseIf CLng(Left(annee, 4)) < 2030 And CLng(Left(annee, 4)) > 1900 Then
                annee = Left(annee, 4)
            ElseIf CLng(annee) < 2030 And CLng(annee) > 1900 Then
                annee = annee
            Else
                annee = "0"
            End If
        End If
    
        'Regarde si une valeur est déjà présente. Si oui, regarde laquelle est la plus grande
        If anneeOutput <> "" Then
            If CLng(annee) > CLng(anneeOutput) Then
                anneeOutput = annee
            End If
        Else
            anneeOutput = annee
        End If
'Fin de ConStance CS4
        
    Next
    
'Récupère le nombre d'exemplaires
    nbEx = 0
    expEx = Split(expEx, Chr(10))
    For Each holding In expEx
    'MsgBox holding
'Permet d'ignorer les holdings pas de la BU et les holdigns liées
        If InStr(holding, bib) > 0 And InStr(holding, "notices liées") = 0 Then
            posEx = InStr(holding, "exemplaire")
            If posEx <> 0 Then
                posPar = InStr(holding, "(")
                antiWhile = 0
                While posEx - posPar > 7
                    holding = Mid(holding, posPar, Len(holding))
                    antiWhile = antiWhile + 1
                    If antiWhile > 10 Then
    'Compter le nombre de antiproc ?
                        nbEx = nbEx + 0
                        Exit For
                    End If
                Wend
                holding = Mid(holding, posPar + 1, posEx - posPar)
                holding = Replace(holding, "e", "")
                holding = Replace(holding, " ", "")
                If IsNumeric(holding) = True Then
                    nbEx = nbEx + CInt(holding)
                Else
                    For ll = 1 To Len(holding)
                        If IsNumeric(Mid(holding, ll, 1)) = True Then
                            nbExReW = nbExReW & Mid(holding, ll, 1)
                        End If
                    Next
                    If nbExReW = "" Then
                        nbExReW = "0"
                    End If
                    nbEx = nbEx + CInt(nbExReW)
                End If
            End If
        End If
        
    Next
    
'output les données
    mainWorkBook.Worksheets("Résultats").Cells(rowIndex, 1).Value = PPN
    If anneeOutput = "" Or anneeOutput = "0" Then
        mainWorkBook.Sheets("Résultats").Range("A" & rowIndex & ":C" & rowIndex).Interior.Color = RGB(255, 0, 0)
    Else
        mainWorkBook.Sheets("Résultats").Range("B" & rowIndex).Value = CLng(anneeOutput)
    End If
    mainWorkBook.Worksheets("Résultats").Cells(rowIndex, 3).Value = nbEx
    'mainWorkBook.Worksheets("Résultats").Cells(rowIndex, 7).Value = mainWorkBook.Worksheets("Résultats").Cells(rowIndex, 7).Value + nbEx

End Sub
Sub getStatsUBAge()
    Dim nonEmptyRows As Integer
    Dim mediane As Long, cumul As Long, zz As Integer, odd As Boolean
    
    mainWorkBook.Sheets("Résultats").Activate
    Range("A:C").Sort key1:=Cells(2, 2), order1:=xlAscending, Header:=xlYes
    nonEmptyRows = Application.WorksheetFunction.CountA(Range("B:B"))
    
    Range("E2") = Application.WorksheetFunction.SumProduct(Range("B2:B" & nonEmptyRows), Range("C2:C" & nonEmptyRows)) / Application.WorksheetFunction.Sum(Range("C2:C" & nonEmptyRows))
'Calcul de la médiane
    mediane = Application.WorksheetFunction.Sum(Range("C2:C" & nonEmptyRows)) / 2
    If mediane = Int(mediane) Then
        odd = False
    Else
        odd = True
    End If
    cumul = 0
    zz = 1
'Donx la boucle s'arrête AU MOMENT où cumul = mediane
    While cumul < mediane
        zz = zz + 1
        cumul = cumul + Range("C" & zz)
    Wend
'Si cumul est strictement supérieur à Int(med)+1 ça veut dire que la mediane est forcément dans ce paquet d'exemplaire
    If cumul > Int(mediane) + 1 Then
        mediane = Range("B" & zz).Value
'Maintenant on traite le cas de cumul = mediane
'Dans ce cas, c'est forcément cumul puisque la boucle s'arrête quand elle rencontre mediane
    ElseIf odd = True Then
        mediane = Range("B" & zz).Value
    Else
'là la médiane est forcément PAIRE et cumul est forcément soit = à la médiane soit égale à mediane+1
        If cumul = mediane Then
            mediane = (Range("B" & zz).Value + Range("B" & zz + 1).Value) / 2
        Else
            mediane = (Range("B" & zz - 1).Value + Range("B" & zz).Value) / 2
        End If
        
    End If
    
    Range("F2") = mediane
    
    Range("H2") = rowIndex - nonEmptyRows
    If Range("H2").Value > 0 Then
        Range("I2") = Application.WorksheetFunction.Sum(Range("C" & nonEmptyRows + 1 & ":C" & rowIndex))
    Else
        Range("I2") = 0
    End If
    
End Sub
Sub getEdition()

    Dim PPN As String
    Dim expTitre As String, clefTitre As String, clefTitreCount As Integer, word As Variant
    Dim expISBN13 As String, expISBN10 As String
    Dim expEd As String
    Dim resultat As String
    Dim antiWhile As Integer, temp As Variant

    PPN = exportAlma.Worksheets("Results").Cells(rowIndex, 10).Value
    expTitre = exportAlma.Worksheets("Results").Cells(rowIndex, 4).Value
    expEd = exportAlma.Worksheets("Results").Cells(rowIndex, 9).Value
    expISBN13 = exportAlma.Worksheets("Results").Cells(rowIndex, 17).Value
    expISBN10 = exportAlma.Worksheets("Results").Cells(rowIndex, 16).Value

'Clean le PPN
    PPN = Mid(PPN, InStr(PPN, "(PPN)"), 14)
    
'Génère la clef du titre
    temp = UCase(Mid(expTitre, 1, InStr(expTitre, " / ")))
    If temp = "" Then
        temp = UCase(expTitre)
    End If
    temp = Replace(temp, " [TEXTE IMPRIMÉ]", "")
    temp = Replace(temp, "LE ", "")
    temp = Replace(temp, "LES ", "")
    temp = Replace(temp, "LA ", "")
    temp = Replace(temp, "L'", "")
    temp = Replace(temp, "UN ", "")
    temp = Replace(temp, "UNE ", "")
    temp = Replace(temp, "DES ", "")
    temp = Replace(temp, ",", "")
    temp = Replace(temp, ": ", "")
    temp = Replace(temp, ";", "")
    temp = Replace(temp, ".", "")
    temp = Replace(temp, "(", "")
    temp = Replace(temp, ")", "")
    temp = Replace(temp, "[", "")
    temp = Replace(temp, "]", "")
    temp = Replace(temp, """", "")
    temp = Replace(temp, " &", "")
    temp = Replace(temp, " =", "")
    temp = Replace(temp, Chr(171), "")
    temp = Replace(temp, Chr(187), "")
    temp = Replace(temp, "'", "")
    temp = Replace(temp, "-", " ")
    temp = Trim(temp)
    temp = Split(temp, " ")
    clefTitreCount = 0
'La clef prend 4 caractères du premier mot, puis 2 caractère des 3 prochains mots
    For Each word In temp
        If clefTitreCount = 0 Then
            clefTitre = Left(word, 4)
        Else
            clefTitre = clefTitre & "_" & Left(word, 2)
        End If
        clefTitreCount = clefTitreCount + 1
        If clefTitreCount = 4 Then
            Exit For
        End If
    Next
    
'output les données
    mainWorkBook.Worksheets("Résultats").Cells(rowIndex, 2).Value = PPN
    mainWorkBook.Worksheets("Résultats").Cells(rowIndex, 3).Value = expTitre
    mainWorkBook.Worksheets("Résultats").Cells(rowIndex, 1).Value = clefTitre
    mainWorkBook.Worksheets("Résultats").Cells(rowIndex, 4).Value = expEd
    mainWorkBook.Worksheets("Résultats").Cells(rowIndex, 5).Value = expISBN13
    mainWorkBook.Worksheets("Résultats").Cells(rowIndex, 6).Value = expISBN10

End Sub
Sub getEditionFind()

    Dim zz As Integer, yy As Integer
    Dim tableMatch(99, 2), tableCount As Integer, output As String
    Dim clefISBN_or As String, clefISBN_dup As String
    
    
    mainWorkBook.Sheets("Résultats").Activate
    
    For zz = 2 To rowIndex
        If Range("D" & zz).Value <> "" Then
'Regarde similarité dans la clef de titre
            tableCount = 0
            For yy = 2 To rowIndex
                If Range("A" & zz).Value = Range("A" & yy).Value And _
                Range("B" & zz).Value <> Range("B" & yy).Value Then
                    tableMatch(tableCount, 0) = Range("B" & yy).Value
                    tableMatch(tableCount, 1) = Range("E" & yy).Value
                    tableMatch(tableCount, 2) = Range("F" & yy).Value
                    tableCount = tableCount + 1
                End If
            Next
'Regarde si les débuts de ISBN coincide
'check si ISBN 13 à des tirets et left(ISBN13, 1, 17) = Left(Replace(ISBN13, "-", ""),1 , 13)
'Sinon prendre le ISBN 10
'Si aucun des deux, move on

            If tableCount > 0 Then
            output = "Double éd. possible :"
'Vérification ISBN
'Su les IsNumeric(Left 16 et left 12) -> on exlcut la clef de contrôle qui peut prendre la valeur X

'Récupération clef ISBN_or
                clefISBN_or = ""
'Si ISBN n'est pas vide ET a le format avec tiret
'Le premier prend l'ISBN 13
                If Range("E" & zz).Value <> "" And IsNumeric(Replace(Left(Range("E" & zz).Value, 16), "-", "")) = True And InStr(Left(Range("E" & zz).Value, 17), "-") > 0 Then
                    clefISBN_or = Left(Range("E" & zz).Value, InStrRev(Left(Range("E" & zz).Value, 15), "-"))
                    clefISBN_or = Left(clefISBN_or, 2) & "8" & Mid(clefISBN_or, 4, Len(clefISBN_or))
'Le second prend l'ISBN 10
                ElseIf Range("F" & zz).Value <> "" And IsNumeric(Replace(Left(Range("F" & zz).Value, 12), "-", "")) = True And InStr(Left(Range("F" & zz).Value, 13), "-") > 0 Then
                    clefISBN_or = "978-" & Left(Range("F" & zz).Value, InStrRev(Left(Range("F" & zz).Value, 11), "-"))
'Si aucun ISBN valide a pu être récupérer, la valeur reste ""
                End If
                
'Boucle pour doubles potentiels
                For yy = 0 To tableCount - 1
                    clefISBN_dup = ""
'clef ISBN pour le qui a enregistré des doubls potentiels
'Si ISBN n'est pas vide ET a le format avec tiret
'Le premier prend l'ISBN 13
                        If CStr(tableMatch(yy, 1)) <> "" And IsNumeric(Replace(Left(CStr(tableMatch(yy, 1)), 16), "-", "")) = True And InStr(Left(CStr(tableMatch(yy, 1)), 17), "-") > 0 Then
                            clefISBN_dup = Left(CStr(tableMatch(yy, 1)), InStrRev(Left(CStr(tableMatch(yy, 1)), 15), "-"))
                            clefISBN_dup = Left(clefISBN_dup, 2) & "8" & Mid(clefISBN_dup, 4, Len(clefISBN_dup))
'Le second prend l'ISBN 10
                        ElseIf CStr(tableMatch(yy, 2)) <> "" And IsNumeric(Replace(Left(CStr(tableMatch(yy, 2)), 12), "-", "")) = True And InStr(Left(CStr(tableMatch(yy, 2)), 13), "-") > 0 Then
                            clefISBN_dup = "978-" & Left(CStr(tableMatch(yy, 2)), InStrRev(Left(CStr(tableMatch(yy, 2)), 11), "-"))
'Si aucun ISBN valide a pu être récupérer, la valeur reste ""
                        End If
                        
'Comapre les clefs ISBN
                        If clefISBN_or = "" Or clefISBN_dup = "" Then
                            output = appendNote(output, "(imp. ISBN) " & tableMatch(yy, 0))
                        ElseIf clefISBN_or = clefISBN_dup Then
                            output = appendNote(output, "(corr. ISBN) " & tableMatch(yy, 0))
                        Else
                            output = appendNote(output, "(NO corr. ISBN) " & tableMatch(yy, 0))
                        End If
                    
                Next
                
                If InStr(output, "(corr. ISBN)") > 0 Then
                    mainWorkBook.Sheets("Résultats").Range("A" & zz & ":G" & zz).Interior.Color = RGB(255, 0, 0)
                ElseIf InStr(output, "(imp. ISBN)") > 0 Then
                    mainWorkBook.Sheets("Résultats").Range("A" & zz & ":G" & zz).Interior.Color = RGB(255, 192, 0)
                ElseIf InStr(output, "(NO corr. ISBN)") > 0 Then
                    mainWorkBook.Sheets("Résultats").Range("A" & zz & ":G" & zz).Interior.Color = RGB(0, 176, 240)
                End If
            Else
                output = "Aucune détection automatique"
                mainWorkBook.Sheets("Résultats").Range("A" & zz & ":G" & zz).Interior.Color = RGB(146, 208, 80)
            End If
            Range("G" & zz) = output
        End If
    Next
    
End Sub
Sub formatEnTetes()
    'https://www.automateexcel.com/vba/format-cells/
    'Crée les en-têtes pour la feuille "Résultats"
    
    mainWorkBook.Worksheets("Résultats").Activate
    Select Case CAX
        Case "[CA2]"
            Range("A1").Value = "PPN"
            Range("B1").Value = "Année"
            Range("C1").Value = "Nb d'ex"
            Range("E1").Value = "Année moyenne"
            Range("F1").Value = "Année médianne"
            Range("H1").Value = "Nb titres exclus"
            Range("I1").Value = "Nb exemplaires exclus"
            colIndex = "C"
        Case "[CA1]"
        
        Case "[CA3]"
            Range("A1").Value = "Clef titre"
            Range("B1").Value = "PPN"
            Range("C1").Value = "Zone Titre"
            Range("D1").Value = "Édition"
            Range("E1").Value = "ISBN 13"
            Range("F1").Value = "ISBN 10"
            Range("G1").Value = "Résultats"
            colIndex = "G"
        Case Else
            MsgBox "La valeur entrée en H2 ne devrait pas être possible"
    End Select
    With mainWorkBook.Worksheets("Résultats").Range("A1:" & colIndex & "1")
        .Interior.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(255, 255, 255)
    End With
    
    'Pour éviter que les PPN deviennent des nombres
    Range("A:" & colIndex).NumberFormat = "@"
    
End Sub


Function appendNote(var As String, text As String)
    If var = "" Then
        var = text
    Else
        var = var & Chr(10) & text
    End If
    appendNote = var
End Function
Sub cleanData()
    Worksheets("Résultats").Activate
    Range("A:ZZ").Delete
    Worksheets("Introduction").Activate
    Range("H2").Select
End Sub
Sub Main()
'Timer : https://www.thespreadsheetguru.com/the-code-vault/2015/1/28/vba-calculate-macro-run-time

'Timer : début
Dim StartTime As Double
Dim MinutesElapsed As String
StartTime = Timer

Set mainWorkBook = ActiveWorkbook
folderPath = Application.ActiveWorkbook.Path

mainWorkBook.Worksheets("Introduction").Activate
CAX = Right(Range("H2").Value, 5)
bib = Range("H4").Value

formatEnTetes

read_Alma_Data

'Lance un script additionnel si nécessaire
Select Case CAX
    Case "[CA2]"
        getStatsUBAge
    Case "[CA3]"
        getEditionFind
End Select

'Formattage cellules
With mainWorkBook.Sheets("Résultats").Range("A2:" & colIndex & rowIndex)
    .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    .Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
    .Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

mainWorkBook.Worksheets("Résultats").Activate
Columns("A:" & colIndex).AutoFit
Rows("1:" & rowIndex).AutoFit

'Formattage spéciaux pour un script
Select Case CAX
    Case "[CA2]"
        With mainWorkBook.Worksheets("Résultats").Range("E1:I1")
            .Interior.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlCenter
            .Font.Color = RGB(255, 255, 255)
        End With
        With mainWorkBook.Sheets("Résultats").Range("E2:I2")
            .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
            .Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        Columns("E:I").AutoFit
        mainWorkBook.Sheets("Résultats").Range("G1:G2").ClearFormats
    Case "[CA3]"
        mainWorkBook.Sheets("Résultats").Range("C:D").HorizontalAlignment = xlLeft
        mainWorkBook.Sheets("Résultats").Range("C:C").ColumnWidth = 65
        mainWorkBook.Sheets("Résultats").Range("D:D").ColumnWidth = 30
        mainWorkBook.Sheets("Résultats").Range("C1:D1").HorizontalAlignment = xlCenter
'faire le format pour les colonnes titre et éd
End Select
Range("A1").Select


'Timer suite & fin
MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
MsgBox "Exécution terminée en " & MinutesElapsed & "."

End Sub

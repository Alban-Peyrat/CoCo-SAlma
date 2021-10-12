Option Explicit

Dim bib As String
Dim RCR As String
Dim inputPPN As String
Dim rowIndex As Integer
Dim folderPath As String
Dim fileName As String
Dim mainWorkBook As Workbook
Sub read_Alma_Data()
    'Originaux : https://www.mrexcel.com/board/threads/vba-selecting-only-rows-with-data-in-a-range.752110/
    'et : https://stackoverflow.com/questions/41725730/how-to-paste-values-and-keep-source-formatting
    'et : https://www.ozgrid.com/forum/index.php?thread/143595-deselect-cells-that-have-been-copied/
    
    'Ouvre le ficheir export_alma et chope le nb de lignes
    Dim nbRow As Integer
    Dim exportAlma As Workbook
    Dim expCote As String, PPN As String, bibList As Variant, cote As String, holding As Variant
    
    Workbooks.Open fileName:=folderPath & "\export_alma.xlsx"
    Set exportAlma = Workbooks("export_alma.xlsx")
    
    nbRow = Cells(Rows.count, "K").End(xlUp).Row

    'Récupère les données
    For rowIndex = 2 To nbRow
        PPN = exportAlma.Worksheets("Results").Cells(rowIndex, 10).Value
        expCote = exportAlma.Worksheets("Results").Cells(rowIndex, 11).Value

        cote = ""
        PPN = Mid(PPN, InStr(PPN, "(PPN)"), 14)
                    
        bibList = Split(expCote, Chr(10))
        For Each holding In bibList
            If InStr(holding, bib) > 0 And InStr(holding, "notices liées") = 0 Then
                If InStr(holding, ";") > 0 Then
                    If cote <> "" Then
                        cote = cote & ";_;"
                        mainWorkBook.Sheets("Résultats").Range("D" & rowIndex).Interior.Color = RGB(146, 208, 80)
                    End If
                    If InStr(InStr(holding, ";") + 1, holding, ";") > InStr(InStr(holding, ";"), holding, " (") Or InStr(InStr(holding, ";") + 1, holding, ";") = 0 Then
                        cote = cote & Mid(holding, InStr(holding, ";") + 2, InStr(holding, "  ") - InStr(holding, ";") - 2)
                    Else
                        cote = cote & Mid(holding, InStr(holding, ";") + 2, InStr(InStr(holding, ";") + 1, holding, ";") - InStr(holding, ";") - 2)
                    End If
                End If
            End If
            MsgBox holding
        Next
        If cote = "" Then
            cote = "PAS DE COTE"
            mainWorkBook.Sheets("Résultats").Range("D" & rowIndex).Interior.Color = RGB(255, 0, 0)
        End If
    
        mainWorkBook.Worksheets("Résultats").Cells(rowIndex, 3).Value = PPN
        mainWorkBook.Worksheets("Résultats").Cells(rowIndex, 4).Value = cote
       
    Next
    Workbooks("export_alma.xlsx").Close
    
End Sub
Sub purifier_PPN_List_From_Alma()
    
    Dim nbRow As Integer
    Dim exportAlma As Workbook
    Dim PPN As String
    
    Set mainWorkBook = ActiveWorkbook
    folderPath = Application.ActiveWorkbook.Path
    Workbooks.Open fileName:=folderPath & "\export_alma.xlsx"
    Set exportAlma = Workbooks("export_alma.xlsx")
    
    nbRow = Cells(Rows.count, "K").End(xlUp).Row

    'Récupère les données
    For rowIndex = 2 To nbRow
        PPN = exportAlma.Worksheets("Results").Cells(rowIndex, 10).Value
        PPN = Mid(PPN, InStr(PPN, "(PPN)"), 14)
        mainWorkBook.Worksheets("Données").Cells(rowIndex, 1).Value = PPN
    Next
    Workbooks("export_alma.xlsx").Close
    mainWorkBook.Worksheets("Données").Activate
    Range("A2").Select
    
End Sub
Sub clean_Result_Page()
    Worksheets("Résultats").Range("1:999999").Delete
End Sub
Sub compare_Cotes()
    
    Dim output As String, i As Integer
    
    For i = 2 To rowIndex
        If Cells(i, 2).Value = "" And Cells(i, 4).Value = "" Then
            Exit For
        Else
            If Cells(i, 2).Value = Cells(i, 4).Value Then
                output = "Oui"
            Else
                output = "NON"
                mainWorkBook.Sheets("Résultats").Range("E" & i).Interior.Color = RGB(255, 192, 0)
            End If
            If InStr(Cells(i, 2).Value, ";_;") > 0 Then
                output = output & Chr(10) & "2 cotes Sudoc"
                mainWorkBook.Sheets("Résultats").Range("E" & i).Interior.Color = RGB(255, 0, 0)
            End If
            If InStr(Cells(i, 4).Value, ";_;") > 0 Then
                output = output & Chr(10) & "2 cotes Alma"
                mainWorkBook.Sheets("Résultats").Range("E" & i).Interior.Color = RGB(0, 176, 240)
            End If
            If Cells(i, 2).Value = "" Or Cells(i, 2).Value = "PAS DE COTE" Then
                output = output & Chr(10) & "Pas de cote Sudoc"
                mainWorkBook.Sheets("Résultats").Range("E" & i).Interior.Color = RGB(255, 0, 0)
            End If
            If Cells(i, 4).Value = "" Or Cells(i, 4).Value = "PAS DE COTE" Then
                output = output & Chr(10) & "Pas de cote Alma"
                mainWorkBook.Sheets("Résultats").Range("E" & i).Interior.Color = RGB(0, 176, 240)
            End If
            If InStr(Cells(i, 1).Value, "PPN INCORRECT") > 0 Then
                output = "ERREUR PPN"
                mainWorkBook.Sheets("Résultats").Range("E" & i).Interior.Color = RGB(0, 0, 0)
                mainWorkBook.Sheets("Résultats").Range("E" & i).Font.Color = RGB(255, 255, 255)
            End If
            Cells(i, 5).Value = output
        End If
    Next
    
    rowIndex = i
    
End Sub
Sub tri_Data()

    Dim pushNb As Integer, isSudoc As Boolean, ii As Integer, jj As Integer
    
    Range("A:B").Sort key1:=Cells(2, 1), order1:=xlAscending, Header:=xlYes
    Range("C:D").Sort key1:=Cells(2, 3), order1:=xlAscending, Header:=xlYes
    
    With Application.WorksheetFunction
        If .CountA(Range("A:A")) >= .CountA(Range("C:C")) Then
            rowIndex = .CountA(Range("A:A"))
        Else
            rowIndex = .CountA(Range("C:C"))
        End If
    End With
    
    pushNb = 0
    
    For ii = 2 To rowIndex
    
    'Le check de Cells vide sert à savoir qu'il n'y a plus de valeurs dans l'une des deux colonnes
    'Parce que vu que les données sont triées et pousées, il ne devrait jamais y avoir une entrée vide analysée s'il n'y a plus d'entrée après
        If Right(Cells(ii + pushNb, 1).Value, 9) <> Right(Cells(ii + pushNb, 3).Value, 9) And _
        (Cells(ii + pushNb, 1).Value <> "") And (Cells(ii + pushNb, 3).Value <> "") Then
            isSudoc = False
            For jj = 1 To 9
                If Mid(Right(Cells(ii + pushNb, 1).Value, 9), jj, 1) <> Mid(Right(Cells(ii + pushNb, 3).Value, 9), jj, 1) Then
                    If Mid(Right(Cells(ii + pushNb, 1).Value, 9), jj, 1) = "X" Or Mid(Right(Cells(ii + pushNb, 1).Value, 9), jj, 1) > Mid(Right(Cells(ii + pushNb, 3).Value, 9), jj, 1) Then
                        isSudoc = True
                    ElseIf Mid(Right(Cells(ii + pushNb, 3).Value, 9), jj, 1) = "X" Or Mid(Right(Cells(ii + pushNb, 1).Value, 9), jj, 1) < Mid(Right(Cells(ii + pushNb, 3).Value, 9), jj, 1) Then
                        isSudoc = False
                    End If
                    jj = 10
                End If
            Next
            If isSudoc = False Then
                Range(Cells(ii + pushNb, 3), Cells(ii + pushNb, 4)).Insert xlShiftDown
            Else
                Range(Cells(ii + pushNb, 1), Cells(ii + pushNb, 2)).Insert xlShiftDown
            End If
            pushNb = pushNb + 1
            ii = ii - 1
        End If
    
    Next
    rowIndex = rowIndex + pushNb
End Sub
Sub reset_Donnees()

    clean_Result_Page

    Worksheets("Données").Activate
    Range("3:999999").ClearContents
    Range("A2").ClearContents
    Range("C2:E2").ClearContents
    Range("F2:F10").Value = " "
    Range("G2").Value = "Fonds1"
    Cells(2, 9).Value = Worksheets("Introduction").Range("S3")
    Range("A2").Select
    
End Sub
Sub extend_035Field()

    Dim count As Integer, i As Integer, nbPPN As Integer
    
    generate_Folder
    
    Worksheets("Données").Activate
    Range("A:B").Sort key1:=Cells(2, 1), order1:=xlAscending, Header:=xlYes
    nbPPN = Application.WorksheetFunction.CountA(Range("A:A"))
    count = 0
    For i = 2 To nbPPN
        If Cells(i + count, 1).Value <> "" Then
            Range("L2").Copy Range("B" & i + count)
            
        Else
            count = count + 1
            i = i - 1
        End If
    Next
    Range("A1").Select
    generate_Alma_Import_File
    
End Sub
Sub read_Sudoc_Data()
    Dim count As Integer, nbPPN As Integer
    Worksheets("Données").Activate
    nbPPN = Application.WorksheetFunction.CountA(Range("B:B"))
    
    For rowIndex = 2 To nbPPN
        inputPPN = Right(Cells(rowIndex, 2).Value, 9)
        GetSudocXMLData
    Next
    
    'C'est de la magie noire que je ne comprends pas, fait avant que je trie les données en début de script comme ça aurait dû être fait depuis le début
    'Tout ce qui touche au nombre de lignes c'est de la magie noire
    'Comment quelqu'un a pu possiblement créer quelque chose d'aussi horrible
    'Je suis vraiment stupide parfois c'est assez fou
    'faut quand même le faire pour aller mettre des nbPPN +1 juste parce que j'ai décidé d'exclure l'en-tête alors que ya 100 fois plus simple en incluant l'en-tête
    'For rowIndex = 2 To nbPPN + 1
    '    If Cells(rowIndex + count, 2).Value <> "" Then
    '        inputPPN = Right(Cells(rowIndex + count, 2).Value, 9)
    '        GetSudocXMLData
    '    Else
    '        count = count + 1
    '        rowIndex = rowIndex - 1
    '    End If
    'Next
    
    Worksheets("Résultats").Activate
    Range("A:B").RemoveDuplicates Columns:=Array(1), Header:=xlYes
    
End Sub
Sub GetSudocXMLData()
    'http://documentation.abes.fr/sudoc/manuels/administration/aidewebservices/#SudocMarcXML
    'https://excel-macro.tutorialhorizon.com/vba-excel-read-data-from-xml-file/
    'https://www.mrexcel.com/board/threads/reading-xml-into-excel-with-vba.822719/
    'https://software-solutions-online.com/excel-vba-get-data-from-web-using-msxml/
    Dim mainWorkBook As Workbook
    Set mainWorkBook = ActiveWorkbook
        
    Dim URL As String, PPN As Variant, i As Integer, kk As Integer, cote As String, oXMLFile, XMLFileName As String
    Dim RCRNodes, ParentNode, ChildssParentNode
    URL = "https://www.sudoc.fr/" & inputPPN & ".xml"
        
    Set oXMLFile = CreateObject("Microsoft.XMLDOM")
    XMLFileName = URL
    oXMLFile.async = False
    oXMLFile.Load (XMLFileName)
    'Récupère PPN et quitte si le PPN n'est pas bon
    Set PPN = oXMLFile.SelectSingleNode("/record/controlfield[@tag='001']/text()")
    If Not PPN Is Nothing Then
        PPN = PPN.NodeValue
    Else
        PPN = "PPN INCORRECT [entrée n°" & rowIndex & " : " & inputPPN & "]"
        mainWorkBook.Sheets("Résultats").Range("A" & rowIndex & ":B" & rowIndex).Interior.Color = RGB(0, 0, 0)
        mainWorkBook.Sheets("Résultats").Range("A" & rowIndex & ":B" & rowIndex).Font.Color = RGB(255, 255, 255)
        mainWorkBook.Sheets("Résultats").Range("A" & rowIndex & ":B" & rowIndex).Value = PPN
        Exit Sub
    End If
    Set RCRNodes = oXMLFile.SelectNodes("/record/datafield[@tag='930']/subfield[@code='b']/text()")
    
    cote = ""
    For i = 0 To (RCRNodes.Length - 1)
        If RCRNodes(i).NodeValue = RCR Then
            If cote <> "" Then
                cote = cote & ";_;"
                mainWorkBook.Sheets("Résultats").Range("B" & rowIndex).Interior.Color = RGB(146, 208, 80)
            End If
            
            Set ParentNode = RCRNodes(i).ParentNode
            Set ParentNode = ParentNode.ParentNode
            Set ChildssParentNode = ParentNode.ChildNodes
            For kk = 0 To (ChildssParentNode.Length - 1)
                If ChildssParentNode(kk).getAttribute("code") = "a" Then
                cote = cote & ChildssParentNode(kk).Text
                End If
            Next
        End If
    Next
    If cote = "" Then
        cote = "PAS DE COTE"
        mainWorkBook.Sheets("Résultats").Range("B" & rowIndex).Interior.Color = RGB(255, 0, 0)
    End If
    mainWorkBook.Sheets("Résultats").Range("B" & rowIndex).Value = cote
    mainWorkBook.Sheets("Résultats").Range("A" & rowIndex).Value = PPN
    
    'Unload oXMLFile(XMLFileName)
    
End Sub
Sub generate_Alma_Import_File()
    'Original: https://docs.microsoft.com/fr-fr/office/vba/api/excel.worksheet.copy
    Worksheets("Données").Copy
    With ActiveWorkbook
        Application.DisplayAlerts = False
         .SaveAs fileName:=folderPath & "\import_alma"
         .Close SaveChanges:=False
    End With
        
End Sub
Function generate_Folder()
        
    Worksheets("Données").Activate
    folderPath = Application.ActiveWorkbook.Path & "\CoCo-SAlma_" & ReplaceIllegalCharacters(Cells(2, 7).Value, "_")
        
    MkDir folderPath
          
    With ActiveWorkbook
        .SaveAs fileName:=folderPath & "\Fichier_ppal"
    End With
    
End Function
Function ReplaceIllegalCharacters(strIn As String, strChar As String) As String
'Original : https://stackoverflow.com/questions/50846340/remove-illegal-characters-while-saving-workbook-excel-vba#answer-50848245
    Dim strSpecialChars As String
    Dim i As Long
    strSpecialChars = "~""#%&*:<>?{|}/\[]" & Chr(10) & Chr(13)

    For i = 1 To Len(strSpecialChars)
        strIn = Replace(strIn, Mid$(strSpecialChars, i, 1), strChar)
    Next

    ReplaceIllegalCharacters = strIn
End Function
Sub formatEnTetes()
    'Cf ConStance
    'Crée les en-têtes pour la feuille "Résultats"
    
mainWorkBook.Worksheets("Résultats").Activate
Range("A1").Value = "PPN Sudoc"
Range("B1").Value = "Cote Sudoc"
Range("C1").Value = "PPN Alma"
Range("D1").Value = "Cote Alma"
Range("E1").Value = "Correspondance ?"


With Worksheets("Résultats").Range("A1:E1")
    .Interior.Color = RGB(0, 0, 0)
    .HorizontalAlignment = xlCenter
    .Font.Color = RGB(255, 255, 255)
End With
    
    'Pour éviter que les PPN du Sudoc deviennent des nombres
    Range("A:E").NumberFormat = "@"
    
End Sub
Sub Main()

'Timer : début (cf ConStance)
Dim StartTime As Double
Dim MinutesElapsed As String
StartTime = Timer

'Set-up les variables globales
folderPath = Application.ActiveWorkbook.Path
fileName = Application.ActiveWorkbook.Name
Set mainWorkBook = ActiveWorkbook
bib = mainWorkBook.Worksheets("Données").Cells(2, 10).Value
RCR = mainWorkBook.Worksheets("Données").Cells(2, 9).Value

clean_Result_Page
formatEnTetes

read_Sudoc_Data
read_Alma_Data

mainWorkBook.Worksheets("Résultats").Activate
tri_Data
'/!\ rowIndex doit avoir la valeur laissée par tri_data
compare_Cotes

mainWorkBook.Worksheets("Résultats").Activate
With mainWorkBook.Sheets("Résultats").Range("A2:E" & rowIndex)
    .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    .Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
    .Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With
Columns("A:E").AutoFit
Rows("1:" & rowIndex).AutoFit
mainWorkBook.Sheets("Résultats").Range(rowIndex & ":999999").Delete
Range("A2").Select
Application.DisplayAlerts = False
ActiveWorkbook.Save

'Timer suite & fin
MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
MsgBox "Exécution terminée en " & MinutesElapsed & "."
End Sub

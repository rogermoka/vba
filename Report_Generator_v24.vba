Public LastRow As Long
Public LastColumn As Long
Public Projeto As String
Public SetRow As Long
Public SetFormula As String
Public ExportName As String
Public Const Filename As String = "REPORTE DIÁRIO CCO_v25.xlsm"
Public Parte As String
Public ErrorMsg As String
Public ErrorCheck As Boolean
Public SetDate As String
Public SetPath As String
Public QTS As String
Public SetWord As String
Public SetMENU As String
Public SetAlarmes As String
Public SetParada As String

Option Compare Text

Sub sbSetHeader()
     
   'Configurando o Header
    SetRow = Split((ActiveSheet.Cells.Find(Projeto, LookAt:=xlWhole).Address), "$")(2) + 2
    SetMENU = Split((ActiveSheet.Cells.Find(Projeto & "MENU", LookAt:=xlWhole).Address), "$")(2) + 1
    SetAlarmes = Split((ActiveSheet.Cells.Find(Projeto & "ALARMES", LookAt:=xlWhole).Address), "$")(2) + 1
    SetParada = Split((ActiveSheet.Cells.Find(Projeto & "PARADA", LookAt:=xlWhole).Address), "$")(2) - 1
    
End Sub

Sub sbResumo()

    Dim aStrings(1 To 5) As String
    
    aStrings(1) = "dmae": aStrings(2) = "caesb": aStrings(3) = "arespcj": aStrings(4) = "guariroba": aStrings(5) = "votorantim"
    
    For Each vItm In aStrings
    
        Projeto = CStr(vItm)
    
        Sheets("Relatório Diário").Select
         
        sbSetHeader
         
        Total = Range("C" & SetMENU)
        Alarmes = Range("C" & SetMENU + 1)
        Comunicaram = Range("C" & SetMENU + 2)
        Ncomunicaram = Range("C" & SetMENU + 3)
        Ncomunicaram3 = Range("C" & SetMENU + 4)
        
        Sheets("Resumo").Select
    
        BuscaProjeto = Split((ActiveSheet.Cells.Find(Projeto, LookAt:=xlWhole).Address), "$")(2)
    
        Range("E" & BuscaProjeto) = Total
        Range("F" & BuscaProjeto) = Alarmes
        Range("G" & BuscaProjeto) = Comunicaram
        Range("H" & BuscaProjeto) = Ncomunicaram
        Range("I" & BuscaProjeto) = Ncomunicaram3
    
    Next vItm
    
End Sub


Sub sbLastRow()
    'Dim LastRow As Long
    With ActiveSheet.UsedRange
        LastRow = .Rows(.Rows.Count).Row
    End With
    'MsgBox LastRow
End Sub

Sub sbLastColumn()
    'Dim LastColumn As Long
    With ActiveSheet.UsedRange
        LastColumn = .Columns(.Columns.Count).Column
    End With
    'MsgBox LastColumn
End Sub
    

Sub Main()

    'Optimize Macro Speed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False


    Dim myFile As String
    Dim vItm As Variant
    Dim aStrings(1 To 5) As String
    
    aStrings(1) = "dmae": aStrings(2) = "caesb": aStrings(3) = "arespcj": aStrings(4) = "guariroba": aStrings(5) = "votorantim"
    
'MODIFICAR PARA USAR
    SetDate = "07112016"
    SetWord = "export"


For Each vItm In aStrings
    
    Projeto = CStr(vItm)
    
    If Projeto = "dmae" Or Projeto = "caesb" Then
    
            If SetWord = "alarmes" Then
            
                myFile = "C:\Users\rmokarze\OneDrive - \" & Projeto & "\" & SetWord & "_" & Projeto & "_" & SetDate & ".xml"
                
                Else
                
                    myFile = "C:\Users\rmokarze\OneDrive - \" & Projeto & "\" & SetWord & "_" & Projeto & "_" & SetDate & ".csv"
            
            End If
            
        Else
    
        myFile = "C:\Users\rmokarze\OneDrive - \" & Projeto & "\" & SetWord & "_" & Projeto & "_" & SetDate & ".xml"
    
    End If
    
    MyInput = Right(myFile, Len(myFile) - InStrRev(myFile, "."))
     
     Select Case True
       
        Case myFile Like "*alarmes*"
            
            Parte = "alarmes"
        
        Case myFile Like "*export*"
            
            Parte = "export"
        
        Case Else
        
            ErrorCheck = True
            ErrorMsg = "Keyword missing! Please, add alarmes or export to file name."
            GoTo Error
            
    End Select
    
    Select Case True
       
        Case myFile Like "*dmae*"
            
            Projeto = "dmae"
        
        Case myFile Like "*caesb*"
            
            Projeto = "caesb"
        
        Case myFile Like "*arespcj*"
            
            Projeto = "arespcj"
 
        Case myFile Like "*guariroba*"
            
            Projeto = "guariroba"
            
        Case myFile Like "*niteroi*"
            
            Projeto = "niteroi"
            
        Case myFile Like "*votorantim*"
            
            Projeto = "votorantim"
           
        Case Else
        
            ErrorCheck = True
            ErrorMsg = "Keyword missing!Por favor, adicione uma keyword com o nome do projeto: export_NOMEDOPROJETO.csv"
            GoTo Error
            
    End Select


If Parte = "export" Then

    If MyInput = "csv" Then
    
         Workbooks.Open (myFile)

    'Text to Columns
        Columns("A:A").Select
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
            Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
            :=Array(Array(5, xlDMYFormat), Array(6, xlDMYFormat))
    Else
    
            If MyInput = "xml" Then
            
                Workbooks.Open (myFile)
            
            Else
                    
                ErrorCheck = True
                ErrorMsg = "Formato de arquivo não suportado"
                GoTo Error
                    
            End If
    
    End If
    
    ExportName = ActiveWorkbook.Name
    
    Windows(Filename).Activate
    Sheets("Relatório Diário").Select
    
        sbSetHeader
    
        Sheets("export_" & Projeto).Select
        Cells.Select
        Selection.ClearContents
        
        Windows(ExportName).Activate
        
        Range("A1").Select
        'revisar trocar por um lógica melhor
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Windows(Filename).Activate
        Sheets("export_" & Projeto).Select
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
        Windows(ExportName).Activate
        ActiveWorkbook.Close SaveChanges:=False
        
    If MyInput = "csv" Then
        
        'Corrigir datas
            'Última leitura
            Columns("E:E").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("E1") = "Data da última leitura"
            Range("E2") = "=IF(DAY(F2)<12,TEXT(LEFT(F2,10),""mm/dd/aaaa""),TEXT(LEFT(F2,10),""dd/mm/aaaa""))"
            sbLastRow
            Range("E2").Select
            Selection.AutoFill Destination:=Range("E2:E" & LastRow)
            Columns("E:E").EntireColumn.Select
            Selection.Copy
            Range("E1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            Columns("F:F").Select
            Selection.Delete Shift:=xlToLeft
            
            'Última atualização
            Columns("F:F").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("F1") = "Última Atualização de Status"
            Range("F2") = "=IF(DAY(G2)<12,TEXT(LEFT(G2,10),""mm/dd/aaaa""),TEXT(LEFT(G2,10),""dd/mm/aaaa""))"
            sbLastRow
            Range("F2").Select
            Selection.AutoFill Destination:=Range("F2:F" & LastRow)
            Columns("F:F").EntireColumn.Select
            Selection.Copy
            Range("F1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            Columns("G:G").Select
            Selection.Delete Shift:=xlToLeft
        
    End If
    
     
    Sheets("export_" & Projeto).Select
    sbLastRow

    'CHECK
            Range("T1") = "CHECK"
            Range("T2") = "=IF(AND( E2 <>"""",P2<>""Generic Water Meter"",(OR(ISNUMBER(SEARCH(""METERFARM"",G2)),ISNUMBER(SEARCH(""LAB DMAE"",G2))))=FALSE),IF(DATEVALUE(E2)=TODAY(),""OK"",IF(DATEVALUE(E2)<(TODAY()-2),TODAY()-DATEVALUE(E2),""NOK"")),""N/A"")"
            Range("T2").Select
            Selection.AutoFill Destination:=Range("T2:T" & LastRow)
    'CHECK ISNUMBER
            Range("U1") = "ISNUMBER"
            Range("U2") = "=ISNUMBER(T2)"
            Range("U2").Select
            Selection.AutoFill Destination:=Range("U2:U" & LastRow)
    'FILTRO
            ActiveSheet.Range("A:U").AutoFilter Field:=21, Criteria1:=Array( _
            "TRUE")
                
            'CLEAN UP / ADJUSTMENTS
            
            Sheets("Relatório Diário").Select
            'MEDIDORES COM ALARME
            Range("C" & SetMENU + 1) = ""
            Range("C" & SetMENU + 2) = "=COUNTIF(export_" & Projeto & "!T:T,""OK"")"
            Range("C" & SetMENU + 3) = "=COUNTIF(export_" & Projeto & "!T:T,""NOK"") + COUNTIF(export_" & Projeto & "!T:T,"">1"")"
            Range("C" & SetMENU + 4) = "=COUNTIF(export_" & Projeto & "!T:T,"">1"")"
            QTS = Range("C" & SetMENU + 4)
            
            'MsgBox "Project: " & Projeto & " Rows to add: " & QTS & " SetRow: " & SetRow & " SetParada: " & SetParada
            
            'Rows(SetRow & ":" & SetParada).ClearContents
            
             If ((SetRow + 1) <> (SetParada - 1)) Then
            
                Rows(SetRow + 1 & ":" & SetParada - 2).EntireRow.Delete
                
            End If
            
            
            If QTS > 2 Then
            
                Rows(SetRow + 1 & ":" & (SetRow + (QTS - 1))).EntireRow.Insert Shift:=xlDown, _
                    CopyOrigin:=xlFormatFromLeftOrAbove
                    
                Else
                
                    Rows(SetRow + 1 & ":" & SetRow + 1).EntireRow.Insert
  
            End If
            
            
    'v24
    
        'Range("A" & SetRow & ":G" & SetParada - 1).WrapText = True
        
        With Range("A" & SetRow & ":G" & SetParada - 1)
            .Sort Key1:=Range("F" & SetRow), Order1:=xlDescending
            .WrapText = True
        End With
     
        Rows(SetRow & ":" & SetParada).EntireRow.AutoFit
            
         
    'MEDIDOR
            Sheets("export_" & Projeto).Select
            Range("A1:A" & LastRow).Select
            Selection.Copy
            Sheets("Relatório Diário").Select
            Range("B" & SetRow).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
    'N/S MEDIDOR
            Sheets("export_" & Projeto).Select
            Range("B1:B" & LastRow).Select
            Selection.Copy
            Sheets("Relatório Diário").Select
            Range("C" & SetRow).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
    'CLIENTE
            Sheets("export_" & Projeto).Select
            Range("G1:G" & LastRow).Select
            Selection.Copy
            Sheets("Relatório Diário").Select
            Range("A" & SetRow).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
    
    'DATA LEITURA
            Sheets("export_" & Projeto).Select
            Range("E1:E" & LastRow).Select
            Selection.Copy
            Sheets("Relatório Diário").Select
            Range("D" & SetRow).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
    'DATA ATUALIZACAO
            Sheets("export_" & Projeto).Select
            Range("F1:F" & LastRow).Select
            Selection.Copy
            Sheets("Relatório Diário").Select
            Range("E" & SetRow).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
    'CHECK
            Sheets("export_" & Projeto).Select
            Range("T1:T" & LastRow).Select
            Selection.Copy
            Sheets("Relatório Diário").Select
            Range("F" & SetRow).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False

'Ajustes MAIN

    'DELETE HEADER
   
    'Rows(SetRow & ":" & SetRow).ClearContents
     Rows(SetRow & ":" & SetRow).Delete
    
        
Else

    If Parte = "alarmes" Then
    
        If MyInput = "xml" Then
        
            
            'MsgBox "It's working..."
            
            Windows(Filename).Activate
            Sheets("Relatório Diário").Select
        
            sbSetHeader
            
            Workbooks.Open (myFile)
            ExportName = ActiveWorkbook.Name
            
            sbLastRow
            
            Range("B3:C" & LastRow).Select
            Selection.Copy
            Windows(Filename).Activate
            Sheets("Relatório Diário").Select
            Range("E" & SetAlarmes).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
                
            Windows(ExportName).Activate
            ActiveWorkbook.Close SaveChanges:=False

        Else
        
            ErrorCheck = True
            ErrorMsg = "File Format Error! Please always use .xml for alarmes reports."
            GoTo Error
        
        
        End If
        
           
    Else
        
    ErrorCheck = True
    ErrorMsg = "Oops! It shouldn't happen."
    GoTo Error
    
    End If
    
End If


Error:
    
    If ErrorCheck = True Then
    
        MsgBox ErrorMsg
        ErrorCheck = False
        
    End If
    
    
Next vItm

'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub


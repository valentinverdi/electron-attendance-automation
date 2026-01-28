Attribute VB_Name = "Módulo1"
Function NoEsta(valor As Variant, largo As Variant) As Boolean
    Dim i As Integer
    NoEsta = True
    
    If largo <> 0 Then
        For i = 1 To largo
            If Range("O" & i).Value = valor Then
                NoEsta = False
                Exit Function
            End If
        Next i
    Else
        NoEsta = True
        Exit Function
    End If
End Function

Function ObtenerMes(fecha As Variant) As String
    ObtenerMes = Format(fecha, "mmmm")
End Function

Function LunAVie(fecha As Variant, mes As Variant) As String
    If fecha = "" Then
        LunAVie = ""
    ElseIf (Weekday(fecha, vbSunday) >= 2 And Weekday(fecha, vbSunday) <= 6) And (UCase(ObtenerMes(fecha)) = UCase(mes)) And (NoEsta(fecha, Range("N2").Value)) Then
        LunAVie = "x"
    Else
        LunAVie = ""
    End If
End Function

Sub CompletarFechas()
    Dim i As Integer
    Dim j As Integer
    Dim fecha As Date
    
    
    For j = 11 To 33
        For Each col In Array("B", "C", "D", "E", "F", "G", "H")
            Range(col & j).Font.Color = RGB(0, 0, 0)
        Next col
    Next j
    
    j = 11
    
    
    For i = 2 To 32
        If Range("L" & i).Value = "x" Then
            fecha = DateValue(Range("K" & i).Value)
            Range("B" & j).Value = fecha
            j = j + 1
        End If
    Next i
    
    For j = j To 33
        For Each col In Array("B", "C", "D", "E", "F", "G", "H")
            Range(col & j).Font.Color = RGB(255, 255, 255)
        Next col
    Next j
End Sub



Sub ConvertToPDF()
    Dim filePath As String
    Dim exportPath As String
    Dim mes As String
    Dim ws As Worksheet
    Dim printArea As Range

    filePath = ThisWorkbook.FullName
    
    Set ws = ActiveSheet
    
    mes = Range("I6").Value

    exportPath = "rutadeexportacion"
    
    Set printArea = ws.Range("A1:J35")
    
    ws.PageSetup.printArea = printArea.Address
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=exportPath, Quality:=xlQualityStandard
    
    
End Sub

Sub Imprimir()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("hoja1")
    
    With ws.PageSetup
        .printArea = "A1:J35"
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PaperSize = xlPaperA4
    End With
    
    ws.PrintOut
    
    
End Sub

Sub establecerNulos()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("hoja1")
    
    ws.Range("P1").Value = ""
    ws.Range("Q1").Value = ""
End Sub


Private Sub Workbook_Open()
    If Worksheets("hoja1").Range("P1").Value = "pdf" Or Worksheets("hoja1").Range("Q1").Value = "imp" Then
        CompletarFechas
    End If
    If Worksheets("hoja1").Range("P1").Value = "pdf" Then
        ConvertToPDF
    End If
    If Worksheets("hoja1").Range("Q1").Value = "imp" Then
        Imprimir
    End If
    If Worksheets("hoja1").Range("P1").Value = "pdf" Or Worksheets("hoja1").Range("Q1").Value = "imp" Then
        establecerNulos
        ThisWorkbook.Save
        Application.Quit
    End If
End Sub

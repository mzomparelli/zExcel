Imports Microsoft.Office.Interop
Imports System.Data
Imports System.Drawing
Imports System.Runtime.InteropServices.Marshal
Imports System.Windows.Forms

Public Class ExcelWorkbook

    Public XL As Excel.Application
    Public WB As Excel.Workbook

#Region "Initialize"

    Public Sub New()
        CreateExcelObject()

    End Sub

    Private Function CreateExcelObject() As String
        Try
            XL = New Excel.Application()
            XL.DisplayAlerts = False
            XL.Visible = False
            Return ""
        Catch ex As Exception
            'ReleaseComObject(XL)
            MessageBox.Show("CreateExcelObject() " & Environment.NewLine & Environment.NewLine & ex.Message)
        End Try
    End Function

    Public Function LoadExcelFile(Optional Filename As String = "") As String

        If Filename = "" Then
            Dim errors As String = CreateBlankWorkbook()
            If errors <> "" Then
                Return errors
            Else
                DeleteSheet("Sheet2")
                DeleteSheet("Sheet3")
                Return ""
            End If
        End If

        Try
            If System.IO.File.Exists(Filename) Then
                WB = XL.Workbooks.Open(Filename)
                Return ""
            Else
                Return "'" & Filename & "' does not exist."
            End If
        Catch ex As Exception
            Return "LoadExcelFile('" & Filename & "')" & Environment.NewLine & Environment.NewLine & ex.Message
        Finally

        End Try

    End Function

#End Region

#Region "Clean-up"

    Public Sub Dispose()
        Close()
        XL.Quit()

        ReleaseComObject(WB)
        ReleaseComObject(XL)

        GC.Collect()

    End Sub

#End Region

#Region "Public Members"

    Public Sub Show()
        XL.Visible = True
    End Sub

    Public Sub Hide()
        XL.Visible = False
    End Sub

    Public Function GetSheetNames() As List(Of String)

        Dim sheets As List(Of String) = New List(Of String)()

        Try
            For Each ws As Excel.Worksheet In WB.Sheets
                sheets.Add(ws.Name)
            Next
        Catch ex As Exception

        End Try

        Return sheets

    End Function

    Public Function GetRowCount(SheetName As String) As Integer
        Try
            Return WB.Worksheets(SheetName).UsedRange.Rows.Count
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return -1
        End Try

    End Function

    Public Function GetRowCount(Sheet As Integer) As Integer
        Try
            Return WB.Worksheets(Sheet).UsedRange.Rows.Count
        Catch ex As Exception
            Return -1
        End Try

    End Function

    Public Function GetRowCount(SheetName As String, iColumn As Integer) As Integer
        Return GetRowCount(GetSheetNumber(SheetName), iColumn)
    End Function

    Public Function GetRowCount(Sheet As Integer, iColumn As Integer) As Integer
        Try
            Dim data As Object = GetArray(Sheet)
            If IsArray(data) Then
                Dim a(,) As Object = data
                For r As Integer = 0 To a.GetLength(0) Step 1
                    If a(r, iColumn - 1) Is Nothing Then
                        Return r
                    End If
                Next
            Else
                Return 1
            End If



        Catch ex As Exception
            Return -1
        End Try

    End Function

    Public Function GetColumnCount(SheetName As String) As Integer
        Try
            Return WB.Worksheets(SheetName).UsedRange.Columns.Count
        Catch ex As Exception
            Return -1
        End Try

    End Function

    Public Function GetColumnCount(Sheet As Integer) As Integer
        Try
            Return WB.Worksheets(Sheet).UsedRange.Columns.Count
        Catch ex As Exception
            Return -1
        End Try

    End Function

    Public Function GetArray(SheetName As String) As Object
        Try
            Return GetArray(GetSheetNumber(SheetName), New Cell(1, 1), New Cell(GetRowCount(SheetName), GetColumnCount(SheetName)))
        Catch ex As Exception
            Return Nothing
        Finally

        End Try
    End Function

    Public Function GetArray(SheetName As String, StartCell As Cell, EndCell As Cell) As Object
        Return GetArray(GetSheetNumber(SheetName), StartCell, EndCell)
    End Function

    Public Function GetArray(Sheet As Integer) As Object
        Try
            Return GetArray(Sheet, New Cell(1, 1), New Cell(GetRowCount(Sheet), GetColumnCount(Sheet)))
        Catch ex As Exception
            Return Nothing
        Finally

        End Try
    End Function

    Public Function GetArray(Sheet As Integer, StartCell As Cell, EndCell As Cell) As Object
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(Sheet)
            Dim cell1 As String = ExcelColumnLetter(StartCell.Column) & StartCell.Row.ToString()
            Dim cell2 As String = ExcelColumnLetter(EndCell.Column) & EndCell.Row.ToString()

            Dim range As Excel.Range = ws.Range(cell1, cell2)
            Dim aRange As Object = range.Value

            ReleaseComObject(ws)
            ReleaseComObject(range)


            If IsArray(aRange) Then
                'Because this is an Excel range we subtract 1 to convert to zero-based
                Dim o(aRange.GetLength(0) - 1, aRange.GetLength(1) - 1) As Object

                For r As Integer = 0 To aRange.GetLength(0) - 1 Step 1
                    For c As Integer = 0 To aRange.GetLength(1) - 1 Step 1
                        If aRange(r + 1, c + 1) Is Nothing Then
                            o(r, c) = New String("")
                        Else
                            o(r, c) = aRange(r + 1, c + 1)
                        End If
                    Next
                Next

                Return o

            Else
                Return aRange
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return Nothing
        Finally

        End Try
    End Function

    Public Function AutoSizeColumns(SheetName As String) As String
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(SheetName)
            ws.Columns.AutoFit()
            ReleaseComObject(ws)
            Return ""
        Catch ex As Exception
            Return "AutoSizeColumns()" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try

    End Function


    Public Function AutoSizeColumns(Sheet As Integer) As String
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(Sheet)
            ws.Columns.AutoFit()
            ReleaseComObject(ws)
            Return ""
        Catch ex As Exception
            Return "AutoSizeColumns()" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try

    End Function

    Public Function FormatRange(SheetName As String, Format As CellFormat) As String
        Dim errors As String = FormatRange(SheetName, New Cell(1, 1), New Cell(GetRowCount(SheetName), GetColumnCount(SheetName)), Format)
        If errors <> "" Then
            Return errors
        End If

        Return ""

    End Function

    Public Function FormatRange(Sheet As Integer, Format As CellFormat) As String
        Dim errors As String = FormatRange(Sheet, New Cell(1, 1), New Cell(GetRowCount(Sheet), GetColumnCount(Sheet)), Format)
        If errors <> "" Then
            Return errors
        End If

        Return ""

    End Function

    Public Function FormatRange(SheetName As String, StartCell As Cell, EndCell As Cell, Format As CellFormat) As String

        Dim SheetNumber As Integer = GetSheetNumber(SheetName)
        Return FormatRange(SheetNumber, StartCell, EndCell, Format)

    End Function

    Public Function FormatRange(Sheet As Integer, StartCell As Cell, EndCell As Cell, Format As CellFormat) As String
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(Sheet)
            Dim cell1 As String = ExcelColumnLetter(StartCell.Column) & StartCell.Row.ToString()
            Dim cell2 As String = ExcelColumnLetter(EndCell.Column) & EndCell.Row.ToString()

            Dim r As Excel.Range = ws.Range(cell1, cell2)

            'Cell BackColor
            If Format.BackColor.HasValue Then
                If Format.BackColor.Value = Color.Transparent Then
                    r.Interior.Pattern = Excel.Constants.xlNone
                    r.Interior.TintAndShade = 0
                    r.Interior.PatternTintAndShade = 0
                Else
                    r.Interior.Color = RGB(Format.BackColor.Value.R, Format.BackColor.Value.G, Format.BackColor.Value.B)
                End If

            End If

            'Cell Font Color
            If Format.FontColor.HasValue Then
                If Format.FontColor.Value = Color.Transparent Then
                    r.Font.ColorIndex = Excel.Constants.xlAutomatic
                    r.Font.TintAndShade = 0
                Else
                    r.Font.Color = RGB(Format.FontColor.Value.R, Format.FontColor.Value.G, Format.FontColor.Value.B)
                End If

            End If

            'Cell Font Type
            If Format.FontFamily <> "" Then
                r.Font.Name = Format.FontFamily
            End If

            'Cell Font Bold
            If Format.FontBold.HasValue Then
                r.Font.Bold = Format.FontBold.Value
            End If

            'Cell Font Italic
            If Format.FontItalic.HasValue Then
                r.Font.Italic = Format.FontItalic.Value
            End If

            'Cell Font Size
            If Format.FontSize.HasValue Then
                r.Font.Size = Format.FontSize.Value
            End If

            'Cell Number Format
            If Format.NumberFormat.HasValue Then
                r.NumberFormat = ExcelNumberFormat.ConvertNumberFormat(Format.NumberFormat)
            End If

            'Cell Wrap Text
            If Format.WrapText.HasValue Then
                r.WrapText = Format.WrapText.Value
            End If

            'Cell Horizontal Alignment
            If Format.FontHAlign.HasValue Then
                Select Case Format.FontHAlign.Value
                    Case CellFormat.HAlignment.Center
                        r.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    Case CellFormat.HAlignment.Left
                        r.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    Case CellFormat.HAlignment.Right
                        r.HorizontalHAlignment = Excel.XlHAlign.xlHAlignRight
                End Select
            End If

            'Cell Vertical Alignment
            If Format.FontVAlign.HasValue Then
                Select Case Format.FontVAlign.Value
                    Case CellFormat.VAlignment.Center
                        r.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    Case CellFormat.VAlignment.Top
                        r.VerticalAlignment = Excel.XlVAlign.xlVAlignTop
                    Case CellFormat.VAlignment.Bottom
                        r.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                End Select
            End If

            Dim b As Integer

            'Cell Borders Bottom
            b = Excel.XlBordersIndex.xlEdgeBottom
            If Format.BorderBottom.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderBottom.LineStyle)
                If Format.BorderBottom.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderBottom.Color.R, Format.BorderBottom.Color.G, Format.BorderBottom.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderBottom.LineWeight)
            End If

            'Cell Borders DiagnalDown
            b = Excel.XlBordersIndex.xlDiagonalDown
            If Format.BorderDiagonalDown.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderDiagonalDown.LineStyle)
                If Format.BorderDiagonalDown.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderDiagonalDown.Color.R, Format.BorderDiagonalDown.Color.G, Format.BorderDiagonalDown.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderDiagonalDown.LineWeight)
            End If

            'Cell Borders DiagnalUp
            b = Excel.XlBordersIndex.xlDiagonalUp
            If Format.BorderDiagonalUp.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderDiagonalUp.LineStyle)
                If Format.BorderDiagonalUp.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderDiagonalUp.Color.R, Format.BorderDiagonalUp.Color.G, Format.BorderDiagonalUp.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderDiagonalUp.LineWeight)
            End If

            'Cell Borders InsideHorizontal
            b = Excel.XlBordersIndex.xlInsideHorizontal
            If Format.BorderInsideHorizontal.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderInsideHorizontal.LineStyle)
                If Format.BorderInsideHorizontal.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderInsideHorizontal.Color.R, Format.BorderInsideHorizontal.Color.G, Format.BorderInsideHorizontal.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderInsideHorizontal.LineWeight)
            End If

            'Cell Borders InsideVertical
            b = Excel.XlBordersIndex.xlInsideVertical
            If Format.BorderInsideVertical.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderInsideVertical.LineStyle)
                If Format.BorderInsideVertical.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderInsideVertical.Color.R, Format.BorderInsideVertical.Color.G, Format.BorderInsideVertical.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderInsideVertical.LineWeight)
            End If

            'Cell Borders Left
            b = Excel.XlBordersIndex.xlEdgeLeft
            If Format.BorderLeft.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderLeft.LineStyle)
                If Format.BorderLeft.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderLeft.Color.R, Format.BorderLeft.Color.G, Format.BorderLeft.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderLeft.LineWeight)
            End If

            'Cell Borders Right
            b = Excel.XlBordersIndex.xlEdgeRight
            If Format.BorderRight.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderRight.LineStyle)
                If Format.BorderRight.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderRight.Color.R, Format.BorderRight.Color.G, Format.BorderRight.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderRight.LineWeight)
            End If

            'Cell Borders Top
            b = Excel.XlBordersIndex.xlEdgeTop
            If Format.BorderTop.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderTop.LineStyle)
                If Format.BorderTop.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderTop.Color.R, Format.BorderTop.Color.G, Format.BorderTop.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderTop.LineWeight)
            End If




            'Cleanup
            ReleaseComObject(ws)
            ReleaseComObject(r)

            Return ""
        Catch ex As Exception
            Return "FormatRange()" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try

    End Function

    Public Function FormatRange(SheetName As String, iRow As Integer, iCol As Integer, Format As CellFormat) As String

        Dim SheetNumber As Integer = GetSheetNumber(SheetName)
        Return FormatRange(SheetNumber, iRow, iCol, Format)

    End Function

    Public Function FormatRange(Sheet As Integer, iRow As Integer, iCol As Integer, Format As CellFormat) As String
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(Sheet)
            Dim cell1 As String = ExcelColumnLetter(iCol) & iRow.ToString()
            Dim cell2 As String = ExcelColumnLetter(iCol) & iRow.ToString()

            Dim r As Excel.Range = ws.Range(cell1, cell2)

            'Cell BackColor
            If Format.BackColor.HasValue Then
                If Format.BackColor.Value = Color.Transparent Then
                    r.Interior.Pattern = Excel.Constants.xlNone
                    r.Interior.TintAndShade = 0
                    r.Interior.PatternTintAndShade = 0
                Else
                    r.Interior.Color = RGB(Format.BackColor.Value.R, Format.BackColor.Value.G, Format.BackColor.Value.B)
                End If

            End If

            'Cell Font Color
            If Format.FontColor.HasValue Then
                If Format.FontColor.Value = Color.Transparent Then
                    r.Font.ColorIndex = Excel.Constants.xlAutomatic
                    r.Font.TintAndShade = 0
                Else
                    r.Font.Color = RGB(Format.FontColor.Value.R, Format.FontColor.Value.G, Format.FontColor.Value.B)
                End If

            End If

            'Cell Font Type
            If Format.FontFamily <> "" Then
                r.Font.Name = Format.FontFamily
            End If

            'Cell Font Bold
            If Format.FontBold.HasValue Then
                r.Font.Bold = Format.FontBold.Value
            End If

            'Cell Font Italic
            If Format.FontItalic.HasValue Then
                r.Font.Italic = Format.FontItalic.Value
            End If

            'Cell Font Size
            If Format.FontSize.HasValue Then
                r.Font.Size = Format.FontSize.Value
            End If

            'Cell Number Format
            If Format.NumberFormat.HasValue Then
                r.NumberFormat = ExcelNumberFormat.ConvertNumberFormat(Format.NumberFormat)
            End If

            'Cell Wrap Text
            If Format.WrapText.HasValue Then
                r.WrapText = Format.WrapText.Value
            End If

            'Cell Horizontal Alignment
            If Format.FontHAlign.HasValue Then
                Select Case Format.FontHAlign.Value
                    Case CellFormat.HAlignment.Center
                        r.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    Case CellFormat.HAlignment.Left
                        r.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    Case CellFormat.HAlignment.Right
                        r.HorizontalHAlignment = Excel.XlHAlign.xlHAlignRight
                End Select
            End If

            'Cell Vertical Alignment
            If Format.FontVAlign.HasValue Then
                Select Case Format.FontVAlign.Value
                    Case CellFormat.VAlignment.Center
                        r.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    Case CellFormat.VAlignment.Top
                        r.VerticalAlignment = Excel.XlVAlign.xlVAlignTop
                    Case CellFormat.VAlignment.Bottom
                        r.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                End Select
            End If

            Dim b As Integer

            'Cell Borders Bottom
            b = Excel.XlBordersIndex.xlEdgeBottom
            If Format.BorderBottom.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderBottom.LineStyle)
                If Format.BorderBottom.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderBottom.Color.R, Format.BorderBottom.Color.G, Format.BorderBottom.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderBottom.LineWeight)
            End If

            'Cell Borders DiagnalDown
            b = Excel.XlBordersIndex.xlDiagonalDown
            If Format.BorderDiagonalDown.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderDiagonalDown.LineStyle)
                If Format.BorderDiagonalDown.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderDiagonalDown.Color.R, Format.BorderDiagonalDown.Color.G, Format.BorderDiagonalDown.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderDiagonalDown.LineWeight)
            End If

            'Cell Borders DiagnalUp
            b = Excel.XlBordersIndex.xlDiagonalUp
            If Format.BorderDiagonalUp.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderDiagonalUp.LineStyle)
                If Format.BorderDiagonalUp.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderDiagonalUp.Color.R, Format.BorderDiagonalUp.Color.G, Format.BorderDiagonalUp.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderDiagonalUp.LineWeight)
            End If

            'Cell Borders InsideHorizontal
            b = Excel.XlBordersIndex.xlInsideHorizontal
            If Format.BorderInsideHorizontal.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderInsideHorizontal.LineStyle)
                If Format.BorderInsideHorizontal.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderInsideHorizontal.Color.R, Format.BorderInsideHorizontal.Color.G, Format.BorderInsideHorizontal.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderInsideHorizontal.LineWeight)
            End If

            'Cell Borders InsideVertical
            b = Excel.XlBordersIndex.xlInsideVertical
            If Format.BorderInsideVertical.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderInsideVertical.LineStyle)
                If Format.BorderInsideVertical.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderInsideVertical.Color.R, Format.BorderInsideVertical.Color.G, Format.BorderInsideVertical.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderInsideVertical.LineWeight)
            End If

            'Cell Borders Left
            b = Excel.XlBordersIndex.xlEdgeLeft
            If Format.BorderLeft.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderLeft.LineStyle)
                If Format.BorderLeft.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderLeft.Color.R, Format.BorderLeft.Color.G, Format.BorderLeft.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderLeft.LineWeight)
            End If

            'Cell Borders Right
            b = Excel.XlBordersIndex.xlEdgeRight
            If Format.BorderRight.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderRight.LineStyle)
                If Format.BorderRight.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderRight.Color.R, Format.BorderRight.Color.G, Format.BorderRight.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderRight.LineWeight)
            End If

            'Cell Borders Top
            b = Excel.XlBordersIndex.xlEdgeTop
            If Format.BorderTop.ChangeBorder Then
                r.Borders(b).LineStyle = ConvertBorderLineStyle(Format.BorderTop.LineStyle)
                If Format.BorderTop.Color = Color.Transparent Then
                    r.Borders(b).ColorIndex = Excel.Constants.xlAutomatic
                    r.Borders(b).TintAndShade = 0
                Else
                    r.Borders(b).Color = RGB(Format.BorderTop.Color.R, Format.BorderTop.Color.G, Format.BorderTop.Color.B)
                End If
                r.Borders(b).Weight = ConvertBorderLineWeight(Format.BorderTop.LineWeight)
            End If

            'Cleanup
            ReleaseComObject(ws)
            ReleaseComObject(r)

            Return ""
        Catch ex As Exception
            Return "FormatRange()" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try

    End Function

    Public Function FormatColumn(SheetName As String, iCol As Integer, Format As CellFormat, Optional IncludeHeaderRow As Boolean = False) As String
        Dim errors As String = ""
        If IncludeHeaderRow Then
            errors = FormatRange(SheetName, New Cell(1, iCol), New Cell(GetRowCount(SheetName), iCol), Format)
        Else
            errors = FormatRange(SheetName, New Cell(2, iCol), New Cell(GetRowCount(SheetName), iCol), Format)
        End If

        If errors <> "" Then
            Return errors
        End If

        Return ""

    End Function

    Public Function FormatColumn(Sheet As Integer, iCol As Integer, Format As CellFormat, Optional IncludeHeaderRow As Boolean = False) As String
        Dim errors As String = ""
        If IncludeHeaderRow Then
            errors = FormatRange(Sheet, New Cell(1, iCol), New Cell(GetRowCount(Sheet), iCol), Format)
        Else
            errors = FormatRange(Sheet, New Cell(2, iCol), New Cell(GetRowCount(Sheet), iCol), Format)
        End If
        If errors <> "" Then
            Return errors
        End If

        Return ""

    End Function

    Public Function FormatRow(SheetName As String, iRow As Integer, Format As CellFormat) As String
        Dim errors As String = ""
        errors = FormatRange(SheetName, New Cell(iRow, 1), New Cell(iRow, GetColumnCount(SheetName)), Format)

        If errors <> "" Then
            Return errors
        End If

        Return ""

    End Function

    Public Function FormatRow(Sheet As Integer, iRow As Integer, Format As CellFormat) As String
        Dim errors As String = ""
        errors = FormatRange(Sheet, New Cell(iRow, 1), New Cell(iRow, GetColumnCount(Sheet)), Format)
        If errors <> "" Then
            Return errors
        End If

        Return ""

    End Function

    Public Function GetSheetNumber(SheetName As String) As Integer
        Dim i As Integer = 0
        For Each s As Excel.Worksheet In WB.Sheets
            i += 1
            If s.Name = SheetName Then
                Exit For
            End If
        Next

        Return i
    End Function

    Public Function Close() As String
        Try
            WB.Close()
            Return ""
        Catch ex As Exception
            Return "Close()" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try
    End Function

    Public Function AddSheet(SheetName As String) As String
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets.Add()
            ws.Name = SheetName
            ReleaseComObject(ws)
            Return ""
        Catch ex As Exception
            Return "AddSheet('" & SheetName & "')" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try

    End Function

    Public Function AddSheet(SheetName As String, ByRef dt As DataTable) As String

        If SheetExists(SheetName) = False Then
            Dim errors As String = AddSheet(SheetName)
            If errors <> "" Then
                Return "AddSheet()" & Environment.NewLine & Environment.NewLine & errors
            End If
        Else
            Return "Sheet '" & SheetName & "' already exists."
        End If

        Try
            Dim errors As String = WriteDataTable(SheetName, dt)
            If errors <> "" Then
                Return errors
            End If

            Return ""
        Catch ex As Exception
            Return "AddSheet()" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try

    End Function

    Public Function HideSheet(SheetName As String) As String
        Return HideSheet(GetSheetNumber(SheetName))
    End Function

    Public Function HideSheet(Sheet As Integer) As String
        Try
            XL.Sheets(Sheet).Visible = False
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function ShowSheet(SheetName As String) As String
        Return ShowSheet(GetSheetNumber(SheetName))
    End Function

    Public Function ShowSheet(Sheet As Integer) As String
        Try
            XL.Sheets(Sheet).Visible = True
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function InsertDataTable(SheetName As String, ByRef dt As DataTable) As String
        If SheetExists(SheetName) = False Then
            Dim errors As String = AddSheet(SheetName)
            If errors <> "" Then
                Return "AddSheet()" & Environment.NewLine & Environment.NewLine & errors
            End If
        End If

        Try
            Dim errors As String = WriteDataTable(SheetName, dt)
            If errors <> "" Then
                Return errors
            End If

            Return ""
        Catch ex As Exception
            Return "AddSheet()" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try
    End Function

    Public Function WriteArray(SheetName As String, ByRef ObjectArray As Object(,)) As String
        Return WriteArray(GetSheetNumber(SheetName), ObjectArray)
    End Function

    Public Function WriteArray(Sheet As Integer, ByRef ObjectArray As Object(,)) As String
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(Sheet)
            Dim range As Excel.Range = ws.Range("A1").Resize(ObjectArray.GetLength(0), ObjectArray.GetLength(1))
            range.Value = ObjectArray
            AutoSizeColumns(Sheet)

            range = ws.Range("A1").Resize(1, ObjectArray.GetLength(1))

            range.Interior.Color = RGB(0, 70, 132)  'Con-way Blue
            range.Font.Color = RGB(Drawing.Color.White.R, Drawing.Color.White.G, Drawing.Color.White.B)
            range.Font.Bold = True
            range.WrapText = True

            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

            range.Application.ActiveWindow.SplitColumn = 0
            range.Application.ActiveWindow.SplitRow = 1
            range.Application.ActiveWindow.FreezePanes = True


            ReleaseComObject(ws)
            ReleaseComObject(range)

            Return ""
        Catch ex As Exception
            Return "WriteArray()" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try
    End Function

    Public Function WriteArray(SheetName As String, StartingRow As Integer, ByRef ObjectArray As Object(,)) As String
        Return WriteArray(GetSheetNumber(SheetName), StartingRow, ObjectArray)
    End Function

    Public Function WriteArray(Sheet As Integer, StartingRow As Integer, ByRef ObjectArray As Object(,)) As String
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(Sheet)
            Dim range As Excel.Range = ws.Range("A1").Resize(ObjectArray.GetLength(0), ObjectArray.GetLength(1))
            range.Value = ObjectArray
            AutoSizeColumns(Sheet)

            range = ws.Range("A" & StartingRow).Resize(1, ObjectArray.GetLength(1))

            ReleaseComObject(ws)
            ReleaseComObject(range)

            Return ""
        Catch ex As Exception
            Return "WriteArray()" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try
    End Function

    Public Function WriteDataTable(SheetName As String, ByRef dt As DataTable) As String
        Try
            Dim l(dt.Rows.Count + 1, dt.Columns.Count) As Object
            For c As Integer = 0 To dt.Columns.Count - 1
                l(0, c) = dt.Columns(c).ColumnName
            Next

            For r As Integer = 1 To dt.Rows.Count
                For c As Integer = 0 To dt.Columns.Count - 1
                    l(r, c) = dt.Rows(r - 1).Item(c)
                Next
            Next

            Dim errors As String = WriteArray(SheetName, l)
            If errors <> "" Then
                Return errors
            End If

            Return ""
        Catch ex As Exception
            Return "WriteDataTable()" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try
    End Function

    Public Function SheetExists(SheetName As String) As Boolean
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(SheetName)
            ReleaseComObject(ws)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function SetColumnWidth(SheetName As String, iCol As Integer, Width As Decimal) As String
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(SheetName)
            ws.Columns(iCol).ColumnWidth = Width
            ReleaseComObject(ws)
            Return ""
        Catch ex As Exception
            Return "SetColumnWidth()" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try

    End Function

    Public Function DeleteSheet(Sheet As String) As String
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(Sheet)
            ws.Delete()
            ReleaseComObject(ws)
            Return ""
        Catch ex As Exception
            Return "DeleteSheet('" & Sheet & "')" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try

    End Function

    Public Function SetCellValue(iSheet As Integer, iRow As Integer, iCol As Integer, value As Object) As String
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(iSheet)
            Dim r As Excel.Range = ws.Cells(iRow, iCol)
            r.Value = value
            ReleaseComObject(ws)
            ReleaseComObject(r)
            Return ""
        Catch ex As Exception
            Return "SetCellValue('" & iSheet & "', '" & iRow & "', '" & iCol & "', '" & value & "')" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try

    End Function

    Public Function SetCellValue(Sheet As String, iRow As Integer, iCol As Integer, value As Object) As String
        Return SetCellValue(GetSheetNumber(Sheet), iRow, iCol, value)
    End Function

    Public Function SetCellValue(iSheet As Integer, iRow As Integer, iCol As Integer, value As Object, format As CellFormat) As String
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(iSheet)
            Dim r As Excel.Range = ws.Cells(iRow, iCol)
            r.Value = value
            FormatRange(iSheet, iRow, iCol, format)
            ReleaseComObject(ws)
            ReleaseComObject(r)
            Return ""
        Catch ex As Exception
            Return "SetCellValue('" & iSheet & "', '" & iRow & "', '" & iCol & "', '" & value & "')" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try

    End Function

    Public Function SetCellValue(Sheet As String, iRow As Integer, iCol As Integer, value As Object, format As CellFormat) As String
        Return SetCellValue(GetSheetNumber(Sheet), iRow, iCol, value, format)
    End Function

    Public Function SetCellFormula(iSheet As Integer, iRow As Integer, iCol As Integer, value As String) As String
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(iSheet)
            Dim r As Excel.Range = ws.Cells(iRow, iCol)
            r.Formula = value
            ReleaseComObject(ws)
            ReleaseComObject(r)
            Return ""
        Catch ex As Exception
            Return "SetCellValue('" & iSheet & "', '" & iRow & "', '" & iCol & "', '" & value & "')" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try

    End Function

    Public Function SetCellFormula(Sheet As String, iRow As Integer, iCol As Integer, value As String) As String
        Return SetCellFormula(GetSheetNumber(Sheet), iRow, iCol, value)
    End Function

    Public Function SetCellFormula(iSheet As Integer, iRow As Integer, iCol As Integer, value As String, format As CellFormat) As String
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(iSheet)
            Dim r As Excel.Range = ws.Cells(iRow, iCol)
            r.Formula = value
            FormatRange(iSheet, iRow, iCol, format)
            ReleaseComObject(ws)
            ReleaseComObject(r)
            Return ""
        Catch ex As Exception
            Return "SetCellValue('" & iSheet & "', '" & iRow & "', '" & iCol & "', '" & value & "')" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try

    End Function

    Public Function SetCellFormula(Sheet As String, iRow As Integer, iCol As Integer, value As String, format As CellFormat) As String
        Return SetCellFormula(GetSheetNumber(Sheet), iRow, iCol, value, format)
    End Function

    Public Function GetCellValue(iSheet As Integer, iRow As Integer, iCol As Integer) As Object
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(iSheet)
            Dim r As Excel.Range = ws.Cells(iRow, iCol)
            Return r.Value
            ReleaseComObject(ws)
            ReleaseComObject(r)
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Function GetCellValue(Sheet As String, iRow As Integer, iCol As Integer) As Object
        Return GetCellValue(GetSheetNumber(Sheet), iRow, iCol)
    End Function

    Public Function GetCellFormula(iSheet As Integer, iRow As Integer, iCol As Integer) As String
        Try
            Dim ws As Excel.Worksheet = WB.Worksheets(iSheet)
            Dim r As Excel.Range = ws.Cells(iRow, iCol)
            Return r.Formula
            ReleaseComObject(ws)
            ReleaseComObject(r)
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Function GetCellFormula(Sheet As String, iRow As Integer, iCol As Integer) As String
        Return GetCellFormula(GetSheetNumber(Sheet), iRow, iCol)
    End Function

    Public Function SaveAs(Filename As String, Optional format As FileFormat = FileFormat.XLSX) As String
        Try
            WB.SaveAs(Filename, format)
            Return ""
        Catch ex As Exception
            Return "SaveAs('" & Filename & "')" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try
    End Function

    Public Function MergeAllSheets() As String
        Try
            Dim wsMain As Excel.Worksheet = WB.Sheets.Item(1)
            Dim sheetCount = WB.Sheets.Count
            Dim iCol As Integer = 0

            For sheet As Integer = 2 To sheetCount Step 1
                Dim ws As Excel.Worksheet = WB.Sheets.Item(sheet)
                ws.Select()
                ws.Range("A1").Select()
                ws.Range(ws.Range("A1"), ws.Range("A1").SpecialCells(Excel.Constants.xlLastCell)).Copy()
                wsMain.Select()
                iCol = wsMain.Range("A1").SpecialCells(Excel.Constants.xlLastCell).Column()
                wsMain.Range("A1").SpecialCells(Excel.Constants.xlLastCell).Offset(1, iCol - (iCol * 2) + 1).Select()
                wsMain.Paste()
            Next

            For sheet As Integer = sheetCount To 2 Step -1
                XL.SendKeys("{ENTER}", 300)
                WB.Sheets(sheet).Delete()
            Next

            wsMain.Select()
            wsMain.Range("A1").Select()

            ReleaseComObject(wsMain)

            Return ""
        Catch ex As Exception
            Return "MargeAllSheets()" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try

    End Function

    Public Function ExcelColumnLetter(ByVal ColumnNumber As Integer) As String
        Dim CumulativeSum As Integer
        Dim StringPosition As Integer
        Dim i As Integer, Modulus As Integer
        Dim TempString As String, PartialValue As Integer

        Try
            If ColumnNumber < 1 Then
                Return ""
                Exit Function
            Else
                StringPosition = 0
                CumulativeSum = CDec(0)
                TempString = ""
                Do
                    PartialValue = Int(CDec((ColumnNumber - CumulativeSum - 1) / (26 ^ StringPosition)))
                    Modulus = PartialValue - Int(CDec(PartialValue / 26)) * 26
                    TempString = Chr(Modulus + 65) & TempString
                    StringPosition = StringPosition + 1
                    CumulativeSum = CDec(0)
                    For i = 1 To StringPosition
                        CumulativeSum = CDec((CumulativeSum + 1) * 26)
                    Next i
                Loop While ColumnNumber > CumulativeSum
                Return TempString
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function ExcelColumnNumber(ByVal ColumnLetter As String) As Long

        Dim intCount As Integer
        Dim intColumnLetterLength As Integer

        Dim i As Integer = 0 : intColumnLetterLength = Len(ColumnLetter)
        For intCount = 1 To intColumnLetterLength
            i = i * 26 + (Asc(UCase(Mid(ColumnLetter, intCount, 1))) - 64)
        Next intCount
        Return i

    End Function





#End Region

#Region "Private Members"

    Private Function ConvertBorderLineStyle(style As Border.BorderLineStyle) As Integer
        Select Case style
            Case Border.BorderLineStyle.Continuous
                Return Excel.XlLineStyle.xlContinuous
            Case Border.BorderLineStyle.Dash
                Return Excel.XlLineStyle.xlDash
            Case Border.BorderLineStyle.DashDot
                Return Excel.XlLineStyle.xlDashDot
            Case Border.BorderLineStyle.DashDotDot
                Return Excel.XlLineStyle.xlDashDotDot
            Case Border.BorderLineStyle.Dot
                Return Excel.XlLineStyle.xlDot
            Case Border.BorderLineStyle.DoubleLine
                Return Excel.XlLineStyle.xlDouble
            Case Border.BorderLineStyle.LineStyleNone
                Return Excel.XlLineStyle.xlLineStyleNone
            Case Border.BorderLineStyle.None
                Return Excel.Constants.xlNone
            Case Border.BorderLineStyle.SlantDashDot
                Return Excel.XlLineStyle.xlSlantDashDot
            Case Else
                Return Nothing
        End Select

    End Function

    Private Function ConvertBorderLineWeight(weight As Border.BorderWeight) As Integer
        Select Case weight
            Case Border.BorderWeight.Hairline
                Return Excel.XlBorderWeight.xlHairline
            Case Border.BorderWeight.Medium
                Return Excel.XlBorderWeight.xlMedium
            Case Border.BorderWeight.Thick
                Return Excel.XlBorderWeight.xlThick
            Case Border.BorderWeight.Thin
                Return Excel.XlBorderWeight.xlThin
            Case Else
                Return Nothing
        End Select
    End Function

    Private Function CreateBlankWorkbook() As String
        Try
            WB = XL.Workbooks.Add()
            Return ""
        Catch ex As Exception
            Return "CreateBlankWorkbook()" & Environment.NewLine & Environment.NewLine & ex.Message
        End Try

    End Function

    Private Function SplitDataTable(ByRef dt As DataTable, MaxRows As Integer) As List(Of DataTable)
        Dim list As List(Of DataTable) = New List(Of DataTable)()

        Dim iRows As Integer = dt.Rows.Count
        If iRows <= MaxRows Then
            list.Add(dt.Copy())
            Return list
        End If

        Dim iTables As Integer = Math.Ceiling((iRows / MaxRows))

        For i As Integer = 1 To iTables Step 1
            Dim t As DataTable = dt.Copy()
            t.Clear()

            Dim iRow As Integer = 1
            For Each dr As DataRow In dt.Rows
                If iRow > MaxRows Then
                    iRow = 1
                    Exit For
                End If
                Dim newRow As DataRow = t.NewRow()
                For x As Integer = 0 To dt.Columns.Count - 1 Step 1
                    newRow(x) = dr(x)
                Next
                t.Rows.Add(newRow)
                dr.Delete()
                iRow += 1
            Next

            dt.AcceptChanges()
            list.Add(t.Copy())
            t.Dispose()
        Next

        GC.Collect()

        Return list

    End Function

#End Region


    Public Enum FileFormat
        XLSX = 51
        XLSB = 50
        XLSM = 52
        XLS = 56
        CSV = 6
        DBF = 11
    End Enum

    Public Structure Cell
        Public Row As Integer
        Public Column As Integer

        Public Sub New(Row As Integer, Column As Integer)
            Me.Row = Row
            Me.Column = Column
        End Sub
    End Structure

End Class



Public Class CellFormat

    Public FontBold? As Boolean
    Public FontFamily As String = ""
    Public FontItalic? As Boolean
    Public FontSize? As Integer
    Public FontColor? As Color
    Public BackColor? As Color
    Public WrapText? As Boolean
    Public NumberFormat? As ExcelNumberFormat.NumberFormat
    Public FontVAlign? As VAlignment
    Public FontHAlign? As HAlignment

    Public BorderBottom As Border = New Border()
    Public BorderTop As Border = New Border()
    Public BorderLeft As Border = New Border()
    Public BorderRight As Border = New Border()
    Public BorderDiagonalDown As Border = New Border()
    Public BorderDiagonalUp As Border = New Border()
    Public BorderInsideVertical As Border = New Border()
    Public BorderInsideHorizontal As Border = New Border()



    Public Sub SetBorderOutline(color As Color, Optional Style As Border.BorderLineStyle = Border.BorderLineStyle.Continuous, Optional Weight As Border.BorderWeight = Border.BorderWeight.Thin)
        BorderBottom.ChangeBorder = True
        BorderBottom.LineStyle = Style
        BorderBottom.Color = color
        BorderBottom.LineWeight = Weight

        BorderTop.ChangeBorder = True
        BorderTop.LineStyle = Style
        BorderTop.Color = color
        BorderTop.LineWeight = Weight

        BorderLeft.ChangeBorder = True
        BorderLeft.LineStyle = Style
        BorderLeft.Color = color
        BorderLeft.LineWeight = Weight

        BorderRight.ChangeBorder = True
        BorderRight.LineStyle = Style
        BorderRight.Color = color
        BorderRight.LineWeight = Weight

    End Sub

    Public Sub SetBorderOutline(Optional Style As Border.BorderLineStyle = Border.BorderLineStyle.Continuous, Optional Weight As Border.BorderWeight = Border.BorderWeight.Thin)
        BorderBottom.ChangeBorder = True
        BorderBottom.LineStyle = Style
        BorderBottom.Color = Color.Transparent
        BorderBottom.LineWeight = Weight

        BorderTop.ChangeBorder = True
        BorderTop.LineStyle = Style
        BorderTop.Color = Color.Transparent
        BorderTop.LineWeight = Weight

        BorderLeft.ChangeBorder = True
        BorderLeft.LineStyle = Style
        BorderLeft.Color = Color.Transparent
        BorderLeft.LineWeight = Weight

        BorderRight.ChangeBorder = True
        BorderRight.LineStyle = Style
        BorderRight.Color = Color.Transparent
        BorderRight.LineWeight = Weight

    End Sub

    Public Sub SetBorderInside(color As Color, Optional Style As Border.BorderLineStyle = Border.BorderLineStyle.Continuous, Optional Weight As Border.BorderWeight = Border.BorderWeight.Thin)
        BorderInsideVertical.ChangeBorder = True
        BorderInsideVertical.LineStyle = Style
        BorderInsideVertical.Color = color
        BorderInsideVertical.LineWeight = Weight

        BorderInsideHorizontal.ChangeBorder = True
        BorderInsideHorizontal.LineStyle = Style
        BorderInsideHorizontal.Color = color
        BorderInsideHorizontal.LineWeight = Weight
    End Sub

    Public Sub SetBorderInside(Optional Style As Border.BorderLineStyle = Border.BorderLineStyle.Continuous, Optional Weight As Border.BorderWeight = Border.BorderWeight.Thin)
        BorderInsideVertical.ChangeBorder = True
        BorderInsideVertical.LineStyle = Style
        BorderInsideVertical.Color = Color.Transparent
        BorderInsideVertical.LineWeight = Weight

        BorderInsideHorizontal.ChangeBorder = True
        BorderInsideHorizontal.LineStyle = Style
        BorderInsideHorizontal.Color = Color.Transparent
        BorderInsideHorizontal.LineWeight = Weight
    End Sub

    Public Sub ClearAllBorders()
        BorderBottom.ChangeBorder = True
        BorderBottom.LineStyle = Border.BorderLineStyle.None

        BorderTop.ChangeBorder = True
        BorderTop.LineStyle = Border.BorderLineStyle.None

        BorderLeft.ChangeBorder = True
        BorderLeft.LineStyle = Border.BorderLineStyle.None

        BorderRight.ChangeBorder = True
        BorderRight.LineStyle = Border.BorderLineStyle.None

        BorderDiagonalDown.ChangeBorder = True
        BorderDiagonalDown.LineStyle = Border.BorderLineStyle.None

        BorderDiagonalUp.ChangeBorder = True
        BorderDiagonalUp.LineStyle = Border.BorderLineStyle.None

        BorderInsideVertical.ChangeBorder = True
        BorderInsideVertical.LineStyle = Border.BorderLineStyle.None

        BorderInsideHorizontal.ChangeBorder = True
        BorderInsideHorizontal.LineStyle = Border.BorderLineStyle.None

    End Sub


    Public Colors As ColorConst = New ColorConst()

    Public Enum HAlignment
        Left
        Center
        Right
    End Enum

    Public Enum VAlignment
        Top
        Center
        Bottom
    End Enum

    Public Sub New()

    End Sub

End Class

Public Class Border

    Public ChangeBorder As Boolean = False
    Public LineStyle As BorderLineStyle = BorderLineStyle.Continuous
    Public LineWeight As BorderWeight = BorderWeight.Thin
    Public Color As Color = Drawing.Color.Transparent

    Public Sub Remove()
        ChangeBorder = True
        LineStyle = BorderLineStyle.None
        'LineWeight = BorderWeight.Thin
        'Color = Drawing.Color.Transparent
    End Sub

    Public Enum BorderLineStyle
        None
        Continuous
        Dash
        DashDot
        DashDotDot
        Dot
        DoubleLine
        LineStyleNone
        SlantDashDot
    End Enum

    Public Enum BorderWeight
        Hairline
        Medium
        Thick
        Thin
    End Enum

    Public Sub New()

    End Sub


End Class

Public NotInheritable Class ExcelNumberFormat

    Private Const DateShort As String = "m/d/yyyy"
    Private Const DateTime As String = "[$-409]m/d/yy h:mm AM/PM;@"
    Private Const Percent As String = "0.00%"
    Private Const Currency As String = "$#,##0.00"
    Private Const Accounting As String = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Private Const DateLong As String = "[$-F800]dddd, mmmm dd, yyyy"
    Private Const Time As String = "[$-F400]h:mm:ss AM/PM"
    Private Const Fraction As String = "# ?/?"
    Private Const Scientific As String = "0.00E+00"
    Private Const Number_NoDecimal As String = "#,###"
    Private Const Number_NoDecimalNoComma As String = "0"
    Private Const Number_1_Decimal As String = "0.0"
    Private Const Number_2_Decimal As String = "0.00"
    Private Const Number_3_Decimal As String = "0.000"
    Private Const Number_4_Decimal As String = "0.0000"
    Private Const General As String = ""


    Public Shared Function ConvertNumberFormat(format As NumberFormat) As String
        Select Case format
            Case NumberFormat.Accounting
                Return ExcelNumberFormat.Accounting
            Case NumberFormat.Currency
                Return ExcelNumberFormat.Currency
            Case NumberFormat.DateLong
                Return ExcelNumberFormat.DateLong
            Case NumberFormat.DateShort
                Return ExcelNumberFormat.DateShort
            Case NumberFormat.DateTime
                Return ExcelNumberFormat.DateTime
            Case NumberFormat.Fraction
                Return ExcelNumberFormat.Fraction
            Case NumberFormat.General
                Return ExcelNumberFormat.General
            Case NumberFormat.Percent
                Return ExcelNumberFormat.Percent
            Case NumberFormat.Scientific
                Return ExcelNumberFormat.Scientific
            Case NumberFormat.Time
                Return ExcelNumberFormat.Time
            Case NumberFormat.Number_1_Decimal
                Return ExcelNumberFormat.Number_1_Decimal
            Case NumberFormat.Number_2_Decimal
                Return ExcelNumberFormat.Number_2_Decimal
            Case NumberFormat.Number_3_Decimal
                Return ExcelNumberFormat.Number_3_Decimal
            Case NumberFormat.Number_4_Decimal
                Return ExcelNumberFormat.Number_4_Decimal
            Case NumberFormat.Number_NoDecimal
                Return ExcelNumberFormat.Number_NoDecimal
            Case NumberFormat.Number_NoDecimalNoComma
                Return ExcelNumberFormat.Number_NoDecimalNoComma
            Case Else
                Return ""
        End Select

    End Function


    Public Enum NumberFormat
        DateShort
        DateTime
        Percent
        Currency
        Accounting
        DateLong
        Time
        Fraction
        Scientific
        General
        Number_NoDecimal
        Number_NoDecimalNoComma
        Number_1_Decimal
        Number_2_Decimal
        Number_3_Decimal
        Number_4_Decimal
    End Enum

    Private Sub New()

    End Sub

End Class

Public Class ColorConst
    Public ExcelGreen As Color = Color.FromArgb(63, 150, 56)
    Public ConwayBlue As Color = Color.FromArgb(0, 70, 132)
    Public ConwaySapphire As Color = Color.FromArgb(0, 143, 204)
    Public ConwayCobalt As Color = Color.FromArgb(0, 110, 173)
    Public ConwayAqua As Color = Color.FromArgb(71, 175, 226)
    Public ConwaySky As Color = Color.FromArgb(146, 191, 229)

End Class

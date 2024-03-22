Imports Word = Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.Threading
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.InteropServices
Imports System.IO




Public Class WordFunctions

#Region "Members"
    'These members are for sleep timing in 'ms' If sleep is not there, 'Call was rejected by callee.' error occurs.
    'Increase the timing if needed.
    'Process.WaitForInputIdle() is an alternative for this.
    '###################################################################
    '                           DISCLAIMER
    'Increasing of Thread.Sleep() timing will slow down the program
    'Execution
    '###################################################################
    Private Const min As Integer = 30
    Private Const med As Integer = 60
    Private Const max As Integer = 90
    Public Shared WApp As Microsoft.Office.Interop.Word.Application
    Public Shared WDoc As Microsoft.Office.Interop.Word.Document
    Public Shared errorList As New List(Of Tuple(Of String, String, String))
#End Region

#Region "Functions"

    ''' <summary>
    ''' This function is to invoke the Word Application (also kills every inst. before invoking).
    ''' </summary>
    ''' <param name="Basic_details"></param>
    ''' <param name="Component_Name"></param>
    ''' <returns> WDoc </returns>
    Public Shared Function PageSetup(Basic_details As List(Of List(Of String)), Component_Name As String) As Word.Document

        Try
            'System.Diagnostics.Process.WaitForInputIdle()
            Thread.Sleep(med)
            'Deserialize (optional)
            Dim serializedString As String = My.Resources.wordTemplate
            Dim _Bytes As Byte() = Convert.FromBase64String(serializedString)
            Dim formatter As New BinaryFormatter()
            Using stream As New MemoryStream(_Bytes)
                Dim deserializedBytes As Byte() = DirectCast(formatter.Deserialize(stream), Byte())
                File.WriteAllBytes("D:" & "\1.docx", deserializedBytes)
            End Using

            killWordproc()

            WApp = New Word.Application()
            WApp.WindowState = Word.WdWindowState.wdWindowStateMaximize
            WDoc = WApp.Documents.Add("D:" & "\1.docx")
            Thread.Sleep(min)
            WDoc.ShowGrammaticalErrors = False
            Thread.Sleep(min)
            WDoc.ShowRevisions = False
            Thread.Sleep(min)
            WDoc.ShowSpellingErrors = False
            Thread.Sleep(min)
            WApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone

            Thread.Sleep(max)
            Dim _Range As Word.Range
            _Range = WDoc.Range(WDoc.Content.Start, WDoc.Content.End)
            _Range.WholeStory()
            _Range.Font.Name = "Segoe UI"
            _Range.Font.Size = 11
            Marshal.ReleaseComObject(_Range)

            'Change as per requirement, if the file is not from deserilization
            With WDoc.PageSetup
                .PaperSize = Word.WdPaperSize.wdPaperA4
                .Orientation = Word.WdOrientation.wdOrientLandscape
            End With

            WApp.Visible = True

        Catch ex As Exception
            MsgBox("Page Setup Error!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try
        Return WDoc

    End Function

    ''' <summary>
    ''' This particular block will forcefull end up all word instances without any time to shut down or cleanup
    ''' This is to avoid 'Call rejected by Callee'
    ''' </summary>
    Private Shared Sub killWordproc()

        Try
            Dim wordInstances() As Process = Process.GetProcessesByName("WINWORD")
            For Each instance As Process In wordInstances
                instance.Kill()
            Next
        Catch ex As Exception
            MsgBox("Failed to close all active Word Instances!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try

    End Sub
    ''' <summary>
    ''' To serialize a document
    ''' </summary>
    Public Shared Sub Serialize()

            Dim filePath As String = "D:\Template.docx"

            Dim fileBytes As Byte() = File.ReadAllBytes(filePath)

            Dim formatter As New BinaryFormatter()
            Using stream As New MemoryStream()
                formatter.Serialize(stream, fileBytes)
                Dim base64String As String = Convert.ToBase64String(stream.ToArray())
                Console.WriteLine("Serialized Word Document as Base64 String: " & base64String)
            End Using

    End Sub


    ''' <summary>
    ''' To convert the .NET data grid view control to MS Word Table
    ''' </summary>
    ''' <param name="DGV"> Data grid view control to be converted </param>
    ''' <param name="columnstoOmit"> Indexes of Columns that needs to be neglected. </param>
    ''' <param name="remarkIndex"> Remark Notes index, includes highlighting too. </param>
    ''' <returns> table </returns>
    Public Shared Function DGVtoTable(DGV As DataGridView, Optional columnstoOmit() As Integer = Nothing, Optional remarkIndex As Integer = Nothing)
        
        Try
            Thread.Sleep(max)
            Dim range As Word.Range = WDoc.Range
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            Dim xg As Integer = 0
            Dim table As Word.Table
            If columnstoOmit Is Nothing Then
                table = WDoc.Tables.Add(range, DGV.Rows.Count + 1, DGV.Columns.GetColumnCount(DataGridViewElementStates.Displayed))
            Else
                table = WDoc.Tables.Add(range, DGV.Rows.Count + 1, DGV.Columns.GetColumnCount(DataGridViewElementStates.Displayed) - columnstoOmit.Count)
            End If

            With table
                Thread.Sleep(min)
                .Columns.Borders.Enable = True
                Thread.Sleep(min)
                .Range.Font.Bold = False
                Thread.Sleep(min)
                .Range.Font.Size = 10
                Thread.Sleep(min)
                .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                Thread.Sleep(min)
                .Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
            End With

            Thread.Sleep(min)
            WApp.ActiveWindow.ScrollIntoView(range)

            'This code block is to centerize the table in an easy (visual) way. Uncomment the next commented code for Conventional method
            Thread.Sleep(max)
            table.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter


            Thread.Sleep(max)
            table.Rows(1).Range.Font.Bold = True

            If columnstoOmit Is Nothing Then
                For f% = 0 To DGV.Columns.GetColumnCount(DataGridViewElementStates.Displayed) - 1
                    Thread.Sleep(min)
                    table.Cell(1, f + 1).Range.Text = DGV.Columns(f).HeaderText
                Next
            Else
                Dim colIndex = 0
                For f% = 0 To DGV.Columns.GetColumnCount(DataGridViewElementStates.Displayed) - 1
                    Thread.Sleep(min)
                    If Not columnstoOmit.Contains(f) Then
                        Thread.Sleep(min)
                        table.Cell(1, colIndex + 1).Range.Text = DGV.Columns(f).HeaderText
                        colIndex += 1
                    End If
                Next
            End If

            'For row values
            Thread.Sleep(max)
            If columnstoOmit Is Nothing Then
                Dim colIndex = 0
                For x% = 0 To DGV.Rows.Count - 1
                    For y% = 0 To DGV.Columns.GetColumnCount(DataGridViewElementStates.Displayed) - 1
                        Thread.Sleep(min)
                        table.Cell(x + 2, y + 1).Range.Text = DGV.Rows(x).Cells(y).Value.ToString()
                        If remarkIndex <> Nothing Then
                            If y = remarkIndex Then
                                Thread.Sleep(min)
                                If Not DGV.Rows(x).Cells(y).Value.ToString().StartsWith("Pass", True, Globalization.CultureInfo.CurrentCulture) Then
                                    errorList.Add(New Tuple(Of String, String, String)((x + 1).ToString, DGV.Rows(x).Cells(1).Value.ToString(), DGV.Rows(x).Cells(y).Value.ToString()))
                                    table.Cell(x + 2, colIndex + 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorYellow
                                    table.Cell(x + 2, colIndex + 1).Range.Font.Bold = True
                                    table.Cell(x + 2, colIndex + 1).Range.Font.Color = Word.WdColor.wdColorRed
                                End If
                            End If
                            colIndex += 1
                        End If
                    Next
                    colIndex = 0
                Next
            Else

                Dim colIndex = 0

                For x% = 0 To DGV.Rows.Count - 1
                    For y% = 0 To DGV.Columns.GetColumnCount(DataGridViewElementStates.Displayed) - 1
                        If Not columnstoOmit.Contains(y) Then
                            Thread.Sleep(max)
                            table.Cell(x + 2, colIndex + 1).Range.Text = DGV.Rows(x).Cells(y).Value.ToString()
                            If remarkIndex <> Nothing Then
                                If y = remarkIndex Then
                                    Thread.Sleep(min)
                                    If Not DGV.Rows(x).Cells(y).Value.ToString().StartsWith("Pass", True, Globalization.CultureInfo.CurrentCulture) Then
                                        errorList.Add(New Tuple(Of String, String, String)((x + 1).ToString, DGV.Rows(x).Cells(1).Value.ToString(), DGV.Rows(x).Cells(y).Value.ToString()))
                                        table.Cell(x + 2, colIndex + 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorYellow
                                        table.Cell(x + 2, colIndex + 1).Range.Font.Bold = True
                                        table.Cell(x + 2, colIndex + 1).Range.Font.Color = Word.WdColor.wdColorRed
                                    End If
                                End If
                            End If
                            colIndex += 1
                        End If
                    Next
                    colIndex = 0
                Next

            End If

            table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)
            table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
            Marshal.ReleaseComObject(range)

            Return table

        Catch ex As Exception
            MsgBox("Data Grid View to Word Table Error!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            Return Nothing
        End Try

    End Function


    ''' <summary>
    ''' To add a new line to the end of the Paragraph
    ''' </summary>
    ''' <param name="content"> String to be added </param>
    ''' <param name="addAnotherLine"> Need to add another line </param>
    ''' <param name="bold"> To make the letters bold </param>
    ''' <param name="addNewPage"> Need to add blank page </param>
    ''' <param name="italic"> To make the letters italic </param>
    ''' <param name="underLine"> To make the letters underlined </param>
    ''' <param name="addTab"> Need to add a tab </param>
    ''' <returns> WDoc </returns>
    Public Shared Function addLine(content As String, Optional addAnotherLine As Boolean = False, Optional bold As Boolean = False, Optional addNewPage As Boolean = False, Optional italic As Boolean = False, Optional underLine As Boolean = False, Optional addTab As Boolean = False)
        Try
            Thread.Sleep(min)
            Dim para1 As Word.Paragraph
            If addNewPage Then
                WDoc.Range(WDoc.Range.StoryLength - 1).InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)
            End If
            WDoc.Paragraphs.Add()
            para1 = WDoc.Paragraphs.Add(WDoc.Bookmarks.Item("\endofdoc").Range)
            Thread.Sleep(min)
            WApp.ActiveWindow.ScrollIntoView(para1.Range)
            If addAnotherLine Then
                para1 = WDoc.Paragraphs.Add(WDoc.Bookmarks.Item("\endofdoc").Range)
            End If
            With para1
                Thread.Sleep(min)
                If addTab Then
                    .Range.Text = vbTab & content
                Else
                    .Range.Text = content
                End If
                Thread.Sleep(min)
                .Range.Font.Bold = bold
                Thread.Sleep(min)
                .Range.Font.Italic = italic
                Thread.Sleep(min)
                .Range.Font.Underline = underLine
            End With
        Catch ex As Exception
            MsgBox("Failed to add line!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try
        Return WDoc
    End Function


    ''' <summary>
    ''' To append text to the last paragraph of the document.
    ''' </summary>
    ''' <param name="content"> Text to append </param>
    ''' <returns> WDoc </returns>
    Public Shared Function appendText(content As String)

        Try
            Thread.Sleep(min)
            Dim para1 As Word.Paragraph
            para1 = WDoc.Paragraphs(WDoc.Paragraphs.Count)
            Dim txt$ = para1.Range.Text

            'To remove the para break at the end
            txt = Left(txt, Len(txt) - 1)
            txt = txt & content
            para1.Range.Text = txt
        Catch ex As Exception
            MsgBox("Failed to append text!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try

        Return WDoc
    End Function


    ''' <summary>
    ''' To find the specified word and modify it.
    ''' </summary>
    ''' <param name="stringtoFind"> Text to look out </param>
    ''' <param name="bold"> To make the letters bold </param>
    ''' <param name="italic"> To make the letters italic </param>
    ''' <param name="underLine"> To make the letters underlined </param>
    ''' <param name="highlight"> Need to add a tab </param>
    ''' <returns> WDoc </returns>
    Public Shared Function findnModify(stringtoFind As String, Optional bold As Boolean = False, Optional italic As Boolean = False, Optional underLine As Boolean = False, Optional highlight As Boolean = False)
        Try
            Thread.Sleep(min)
            Dim Range As Word.Range = WDoc.Content
            With Range.Find
                .Text = stringtoFind
                .Forward = True
                .Wrap = Word.WdFindWrap.wdFindStop
                .Format = False
                .MatchCase = True
                .MatchWholeWord = True
                .Execute()
                Do While .Found
                    Range.Bold = bold
                    Range.Italic = italic
                    Range.Underline = underLine
                    If highlight Then
                        Range.HighlightColorIndex = Word.WdColorIndex.wdYellow
                    End If
                    .Execute()
                Loop
            End With
        Catch ex As Exception
            MsgBox("Failed to Find and Modify!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try
        Return WDoc
    End Function

    ''' <summary>
    ''' To add formula at end of doc
    ''' </summary>
    ''' <param name="formula"> Mathematical expression in regular format </param>
    ''' <param name="bold"> To make the letter bold </param>
    ''' <param name="Text"> The text need to be in front of the Formula </param>
    ''' <param name="centertoPage"> To centerize the formula </param>
    ''' <returns></returns>
    Public Shared Function addFormula(formula As String, Optional bold As Boolean = False, Optional Text As String = "", Optional centertoPage As Boolean = False)
        Try
            Thread.Sleep(min)
            Dim range As Word.Range
            range = WDoc.Bookmarks.Item("\endofdoc").Range
            Thread.Sleep(min)
            range.InsertParagraphAfter()
            Thread.Sleep(min)
            WApp.ActiveWindow.ScrollIntoView(range)
            Thread.Sleep(min)
            range.InsertAfter(formula)
            Thread.Sleep(min)
            range = WDoc.OMaths.Add(range)
            Thread.Sleep(min)
            range.OMaths.BuildUp()
            Thread.Sleep(min)
            range.Text = range.Text & Text
            Thread.Sleep(min)
            Dim para1 As Word.Paragraph
            Thread.Sleep(min)
            para1 = range.Paragraphs(1)
            Thread.Sleep(min)
            para1.Range.Font.Bold = bold
        Catch ex As Exception
            MsgBox("Failed to add formula!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try
        Return WDoc
    End Function


    ''' <summary>
    ''' To add a blank page
    ''' </summary>
    Public Shared Sub nextPage()
        Try
            Thread.Sleep(min)
            WDoc.Range(WDoc.Range.StoryLength - 1).InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)
        Catch ex As Exception
            MsgBox("Failed to add Blank Page!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try
    End Sub


    ''' <summary>
    ''' To add a tab at End of the Document
    ''' </summary>
    Public Shared Sub addTabatEOD()
        Try
            Dim range As Word.Range
            Thread.Sleep(min)
            range = WDoc.Bookmarks.Item("\endofdoc").Range
            Thread.Sleep(min)
            range.InsertBefore(vbTab)
            Thread.Sleep(min)
            WApp.ActiveWindow.ScrollIntoView(range)
        Catch ex As Exception
            MsgBox("Failed to add Tab to End of Document added!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try
    End Sub


    ''' <summary>
    ''' To Insert a text watermark
    ''' </summary>
    ''' <param name="watermarkText"> Text </param>
    Public Shared Sub InsertWatermark(watermarkText As String)

Genesis:
        Try

            Dim headerRange As Word.HeaderFooter
            Dim section As Word.Section
            Thread.Sleep(min)
            section = WDoc.Sections(1)
            Thread.Sleep(min)
            headerRange = section.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary)
            'For Each shape As Word.Shape In headerRange.Shapes
            '    shape.Delete()
            'Next
            Dim shapeRange As Word.Shape
            Thread.Sleep(min)
            shapeRange = headerRange.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect1, watermarkText, "Arial", 45, False, False, 0, 0)
            With shapeRange
                Thread.Sleep(min)
                .Line.Visible = MsoTriState.msoTrue
                Thread.Sleep(min)
                .Rotation = -45
                Thread.Sleep(min)
                .Rotation = -45
                Thread.Sleep(min)
                .Top = 220
                .Left = 75
            End With
        Catch ex As Exception
            MsgBox("Failed to add watermark!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try

    End Sub


    ''' <summary>
    ''' To add a blank table to End of the Document
    ''' </summary>
    ''' <param name="rows"> No. of rows </param>
    ''' <param name="columns"> No. of columns </param>
    ''' <param name="enableBorder"> To enable all borders </param>
    ''' <param name="boldHeader"> To make the first row bold </param>
    ''' <param name="horizontallyCenter"> To centerize the cells horizontally </param>
    ''' <param name="verticallyCenter"> To centerize the cells vertically </param>
    ''' <param name="autofit"> To autofit to the content </param>
    ''' <returns> table </returns>
    Public Shared Function addTable(rows As Integer, columns As Integer, Optional enableBorder As Boolean = True, Optional boldHeader As Boolean = True, Optional horizontallyCenter As Boolean = True, Optional verticallyCenter As Boolean = True, Optional autofit As Boolean = True)

        Dim table As Word.Table
        Try
            Dim range As Word.Range
            Thread.Sleep(min)
            range = WDoc.Bookmarks.Item("\endofdoc").Range
            Thread.Sleep(min)
            Dim para1 As Word.Paragraph
            Thread.Sleep(min)
            WDoc.Paragraphs.Add()
            Thread.Sleep(min)
            para1 = WDoc.Paragraphs.Add(range)
            Thread.Sleep(min)
            para1.Range.Text = vbNewLine
            Thread.Sleep(min)
            table = WDoc.Tables.Add(para1.Range, rows, columns)
            If autofit Then
                Thread.Sleep(min)
                table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)
            End If
            With table
                If horizontallyCenter Then
                    Thread.Sleep(min)
                    .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                Else
                    Thread.Sleep(min)
                    .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                End If
                Thread.Sleep(min)
                .Range.ParagraphFormat.SpaceAfter = 0
                Thread.Sleep(min)
                .Borders.Enable = enableBorder
            End With

            Thread.Sleep(min)
            range = table.Range
            With range
                If verticallyCenter Then
                    Thread.Sleep(min)
                    .Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                Else
                    Thread.Sleep(min)
                    .Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
                End If
            End With

            Thread.Sleep(min)
            range = table.Rows(1).Range
            Thread.Sleep(min)
            range.Font.Bold = boldHeader
            Return table
        Catch ex As Exception
            MsgBox("Failed to add table!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            Return table
        End Try

    End Function

    ''' <summary>
    ''' To add values() to column
    ''' </summary>
    ''' <param name="table"> The preferred table </param>
    ''' <param name="columnIndex"> The specific column index </param>
    ''' <param name="values"> An array of values needs to be filled on the column </param>
    ''' <returns> table </returns>
    Public Shared Function addvaluestoColumn(table As Word.Table, columnIndex As Integer, values() As String)

        Try
            With table
                For x% = 1 To values.Length
                    Thread.Sleep(min)
                    .Cell(x, columnIndex).Range.Text = values(x - 1)
                Next
            End With
            Return table
        Catch ex As Exception
            Return table
            MsgBox("Failed to add values to Column!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try

    End Function

    ''' <summary>
    ''' To add values() to row
    ''' </summary>
    ''' <param name="table"> The preferred table </param>
    ''' <param name="rowIndex"> The specific row index </param>
    ''' <param name="values"> An array of values needs to be filled on the row </param>
    ''' <returns> table </returns>
    Public Shared Function addvaluestoRow(table As Word.Table, rowIndex As Integer, values() As String)

        Try
            With table
                For x% = 1 To values.Length
                    Thread.Sleep(min)
                    .Cell(rowIndex, x).Range.Text = values(x - 1)
                Next
            End With
            Return table
        Catch ex As Exception
            Return table
            MsgBox("Failed to add values to Row!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try

    End Function

    ''' <summary>
    ''' To a columns to end of the table
    ''' </summary>
    ''' <param name="table"> The table which needs to add columns </param>
    ''' <param name="noofColumns"> Count of columns </param>
    ''' <returns> table </returns>
    Public Shared Function addColumntoTable(table As Word.Table, noofColumns As Integer)

        For x% = 0 To noofColumns - 1
            Thread.Sleep(min)
            table.Columns.Add()
        Next
        Return table

    End Function


#End Region

End Class

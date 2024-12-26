Imports Word = Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.Threading
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.InteropServices
Imports System.IO
Imports IBR_Calculator_Live.GeneralFunctions
Imports IBR_Calculator_Live.Global_Declaration



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
    Enum Speed
        min = 30 '10
        med = 60 '20
        max = 90 '30
    End Enum
    'Public Shared WApp As Microsoft.Office.Interop.Word.Application
    'Public Shared WDoc As Microsoft.Office.Interop.Word.Document
    'Public Shared errorList As New List(Of Tuple(Of String, String, String))
    Public Const fontSize As Double = 9.5
#End Region

#Region "Functions"

    ''' <summary>
    ''' To make the program to pause.
    ''' </summary>
    ''' <param name="timing"> Time to pause in Milli second. </param>
    Public Shared Sub sleep(timing As Speed)
        Thread.Sleep(timing)
    End Sub

    ''' <summary>
    ''' This function is to invoke the Word Application (also kills every inst. before invoking).
    '''
    ''' </summary>
    ''' <param name="Basic_details"></param>
    ''' <param name="Component_Name"></param>
    ''' <returns> WDoc </returns>
    Public Shared Function PageSetup(Basic_details As List(Of List(Of String)), Component_Name As String, BasicDetailDGV As DataGridView) As Word.Document

        WriteLog("Page Setup", "")
        Try
            WriteLog("", $"Initiating Page Setup for {Basic_details(0)(7)}...")
            'System.Diagnostics.Process.WaitForInputIdle()
            'Thread.Sleep(Speed.med)
            Dim Folder() As String = Create_Folder("IBR CALCULATIONS", Basic_details(0)(0), Basic_details(0)(1) & " " & Component_Name)

            'Deserialize
            Dim serializedString As String = My.Resources.wordTemplate
            Dim _Bytes As Byte() = Convert.FromBase64String(serializedString)
            Dim formatter As New BinaryFormatter()
            Using stream As New MemoryStream(_Bytes)
                Dim deserializedBytes As Byte() = DirectCast(formatter.Deserialize(stream), Byte())
                File.WriteAllBytes(Folder(1) & "\1.docx", deserializedBytes)
            End Using

            killWordproc()

            WApp = New Word.Application()
            'WApp.WindowState = Word.WdWindowState.wdWindowStateMaximize
            WDoc = WApp.Documents.Add(Folder(1) & "\1.docx")
            'For Each addIn In WApp.COMAddIns
            '    addIn.connect = true
            'Next
            sleep(Speed.min)
            WDoc.ShowGrammaticalErrors = False
            sleep(Speed.min)
            WDoc.ShowRevisions = False
            sleep(Speed.min)
            WDoc.ShowSpellingErrors = False
            sleep(Speed.min)
            WApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
            'if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document
            sleep(Speed.min)
            WDoc.Paragraphs.WidowControl = True


            sleep(Speed.max)
            Dim _Range As Word.Range
            _Range = WDoc.Range(WDoc.Content.Start, WDoc.Content.End)
            _Range.WholeStory()
            _Range.Font.Name = "Segoe UI"
            _Range.Font.Size = fontSize
            Marshal.ReleaseComObject(_Range)
            'HOWEVER IT IS UNWANTED BECAUSE THE SERIALIZED DOC WILL HAVE THESE ATTRIBUTES BY DEFAULT, DELETE IF NEEDED
            With WDoc.PageSetup
                .PaperSize = Word.WdPaperSize.wdPaperA4
                .Orientation = Word.WdOrientation.wdOrientLandscape
            End With
            Dim headerTable As Word.Table
            headerTable = WDoc.Sections(1).Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Tables(1)
            'WApp.Visible = True
            'WApp.WindowState = Word.WdWindowState.wdWindowStateMaximize
            'WApp.Activate()
            WDoc.Activate()

            With headerTable
                sleep(Speed.min)
                .Range.Font.Name = "Calibri (Body)"
                sleep(Speed.min)
                .Range.Font.Size = 12
                sleep(Speed.min)
                'PROJECT NO
                If Basic_details(0)(1).ToUpper = "CUSTOM" Then
                    .Cell(2, 1).Range.Text = "Project No.: " & Basic_details(0)(9).ToUpper.ToString
                Else
                    .Cell(2, 1).Range.Text = "Project No.: " & Basic_details(0)(1).ToUpper.ToString
                End If
                sleep(Speed.min)
                'DOC TITLE
                .Cell(2, 2).Range.Text = "IBR CALCULATIONS FOR " & Basic_details(0)(8).ToUpper.ToString '"IBR Calculations for " & Component_Name.ToString
                sleep(Speed.min)
                'BOILER NO
                .Cell(3, 1).Range.Text = "Boiler No.: " & Basic_details(0)(0).ToUpper.ToString
                sleep(Speed.min)
                'DOC NO
                .Cell(3, 2).Range.Text = "Document No.: " & Basic_details(0)(7).ToUpper.ToString
                sleep(Speed.min)
                'REVISION
                .Cell(3, 3).Range.Text = "Revision: 00"
                sleep(Speed.min)
                'DATE
                .Cell(4, 3).Range.Text = "Date: " & System.DateTime.Now.ToString("dd-MM-yyyy")
            End With
            sleep(Speed.min)
            Dim footerTable As Word.Table = WDoc.Sections(1).Footers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Tables(1)

            With footerTable
                sleep(Speed.min)
                .Range.Font.Name = "Calibri (Body)"
                sleep(Speed.min)
                .Range.Font.Size = 12
                sleep(Speed.min)
                'PROJECT NO
                '.Cell(1, 2).Range.Text = Basic_details(0)(4).ToUpper.ToString
                'sleep(Speed.min)
                '.Cell(2, 2).Range.Text = Basic_details(0)(5).ToUpper.ToString
                'sleep(Speed.min)
                '    .Cell(3, 2).Range.Text = Basic_details(0)(6).ToUpper.ToString


                'Done to adjust multiple user requirement, once finalized it should be fixed
                sleep(Speed.min)
                .Borders.Enable = False
                .Cell(1, 1).Range.Text = ""
                sleep(Speed.min)
                .Cell(2, 1).Range.Text = ""
                sleep(Speed.min)
                .Cell(3, 1).Range.Text = ""
            End With
            Dim para1 As Word.Paragraph
            sleep(Speed.min)
            para1 = WDoc.Paragraphs.Add()
            '_Range = WDoc.Range(Start:=0, [End]:=0)
            '_Range.Select()

            Dim _table1 As Word.Table
            sleep(Speed.min)
            _table1 = WDoc.Tables.Add(para1.Range, 7, 2)
            Dim rng As Word.Range
            sleep(Speed.min)
            rng = _table1.Range
            With _table1
                sleep(Speed.min)
                .Borders.Enable = False
                sleep(Speed.min)
                .Columns(1).SetWidth(WApp.CentimetersToPoints(8.6), Word.WdRulerStyle.wdAdjustSameWidth)
                sleep(Speed.min)
                .Columns(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                sleep(Speed.min)
                rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                sleep(Speed.min)
                rng.ParagraphFormat.SpaceAfter = 0
                sleep(Speed.min)
                rng.ParagraphFormat.SpaceBefore = 0

                For Each cells As Word.Cell In .Columns(1).Cells
                    sleep(Speed.min)
                    cells.Range.Font.Bold = True
                Next
                sleep(Speed.min)
                .Cell(1, 1).Range.Text = "Project No."
                sleep(Speed.min)
                .Cell(2, 1).Range.Text = "Boiler No."
                sleep(Speed.min)
                .Cell(3, 1).Range.Text = "Client Name"
                sleep(Speed.min)
                .Cell(4, 1).Range.Text = "Code of Construction"
                sleep(Speed.min)
                .Cell(5, 1).Range.Text = "Type of boiler"
                sleep(Speed.min)
                .Cell(6, 1).Range.Text = "Component Name"
                sleep(Speed.min)
                .Cell(7, 1).Range.Text = "Document No."

                sleep(Speed.min)
                If Basic_details(0)(1).ToUpper = "CUSTOM" Then
                    .Cell(1, 2).Range.Text = ": " & Basic_details(0)(9).ToUpper.ToString
                Else
                    .Cell(1, 2).Range.Text = ": " & Basic_details(0)(1).ToUpper.ToString
                End If
                sleep(Speed.min)
                .Cell(2, 2).Range.Text = ": " & Basic_details(0)(0).ToUpper.ToString
                sleep(Speed.min)
                .Cell(3, 2).Range.Text = ": " & Basic_details(0)(2).ToUpper.ToString
                sleep(Speed.min)
                .Cell(4, 2).Range.Text = ": " & "I.B.R. 1950 with latest amendment"
                sleep(Speed.min)
                .Cell(5, 2).Range.Text = ": " & Basic_details(0)(3).ToUpper.ToString
                sleep(Speed.min)
                .Cell(6, 2).Range.Text = ": " & Component_Name.ToUpper.ToString
                'For x As Integer = 0 To Basic_details(1).Count - 1
                '    If Basic_details(1)(x) IsNot Nothing Then
                '        sleep(Speed.min)
                '        .Cell(7, 2).Range.Text = .Cell(7, 2).Range.Text & ": " & Basic_details(1)(x).ToUpper.ToString()
                '    End If
                'Next
                sleep(Speed.min)
                .Cell(7, 2).Range.Text = ": " & Basic_details(0)(7).ToUpper.ToString


            End With
            addLine("")

#Region "Revision Table"
            'sleep(Speed.min)
            'WDoc.Paragraphs.Add()
            'sleep(Speed.min)
            '    para1 = WDoc.Paragraphs.Add()
            'sleep(Speed.min)
            '    _table1 = WDoc.Tables.Add(para1.Range, 3, 7)
            '    sleep(Speed.min)
            'rng.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            'sleep(Speed.min)
            'rng = _table1.Rows(1).Range
            'sleep(Speed.min)
            'rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            'sleep(Speed.min)
            'rng.Font.Bold = True
            'With _table1
            '    sleep(Speed.min)
            '    .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            '    sleep(Speed.min)
            '    .Range.ParagraphFormat.SpaceAfter = 0
            '    sleep(Speed.min)
            '    .Borders.Enable = True
            '    sleep(Speed.min)
            '    .Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
            '    sleep(Speed.min)
            '        .Columns(1).SetWidth(WApp.CentimetersToPoints(2), Word.WdRulerStyle.wdAdjustSameWidth)
            '        sleep(Speed.min)
            '        .Columns(2).SetWidth(WApp.CentimetersToPoints(2.5), Word.WdRulerStyle.wdAdjustSameWidth)
            '        sleep(Speed.min)
            '        .Columns(3).SetWidth(WApp.CentimetersToPoints(2.5), Word.WdRulerStyle.wdAdjustSameWidth)
            '        sleep(Speed.min)
            '        .Columns(4).SetWidth(WApp.CentimetersToPoints(2.5), Word.WdRulerStyle.wdAdjustSameWidth)
            '        sleep(Speed.min)
            '        .Columns(5).SetWidth(WApp.CentimetersToPoints(2.5), Word.WdRulerStyle.wdAdjustSameWidth)
            '        sleep(Speed.min)
            '        .Columns(6).SetWidth(WApp.CentimetersToPoints(6.5), Word.WdRulerStyle.wdAdjustSameWidth)
            '        sleep(Speed.min)
            '        .Columns(7).SetWidth(WApp.CentimetersToPoints(2.5), Word.WdRulerStyle.wdAdjustSameWidth)

            '        sleep(Speed.min)
            '    .Cell(1, 1).Range.Text = "REVISION"
            '    sleep(Speed.min)
            '    .Cell(1, 2).Range.Text = "PREPARED BY"
            '    sleep(Speed.min)
            '    .Cell(1, 3).Range.Text = "CHECKED BY"
            '    sleep(Speed.min)
            '    .Cell(1, 4).Range.Text = "APPROVED BY"
            '    sleep(Speed.min)
            '        .Cell(1, 5).Range.Text = "DEPT."
            '        sleep(Speed.min)
            '        .Cell(1, 6).Range.Text = "REVISION NOTES"
            '        sleep(Speed.min)
            '        .Cell(1, 7).Range.Text = "DATE"

            '        sleep(Speed.min)
            '        .Cell(2, 1).Range.Text = "00"
            '    sleep(Speed.min)
            '    .Cell(2, 2).Range.Text = Basic_details(0)(4).ToUpper.ToString
            '    sleep(Speed.min)
            '    .Cell(2, 3).Range.Text = Basic_details(0)(5).ToUpper.ToString
            '    sleep(Speed.min)
            '        .Cell(2, 4).Range.Text = Basic_details(0)(6).ToUpper.ToString
            '        sleep(Speed.min)
            '        .Cell(2, 5).Range.Text = "Mechanical"
            '        sleep(Speed.min)
            '        .Cell(2, 7).Range.Text = System.DateTime.Now.ToString("dd-MM-yyyy")

            '    End With
#End Region

            If Basic_details(1).Count <> 0 Or Basic_details(1) IsNot Nothing Then
                sleep(Speed.min)
                'WDoc.Range(WDoc.Range.StoryLength - 1).InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)
                sleep(Speed.min)
                para1 = WDoc.Paragraphs.Add(WDoc.Bookmarks.Item("\endofdoc").Range)
                sleep(Speed.min)
                para1.Range.Text = "1. REFERENCE DRAWINGS: "
                sleep(Speed.min)
                para1.Style = Word.WdBuiltinStyle.wdStyleHeading1
                sleep(Speed.min)
                para1.Range.Font.Name = "Segoe UI"
                sleep(Speed.min)
                para1.Range.Font.Size = fontSize
                sleep(Speed.min)
                para1.Range.Font.Color = Word.WdColor.wdColorBlack
                sleep(Speed.min)
                para1.Range.Font.Bold = True
                sleep(Speed.min)
                WDoc.Paragraphs.Add()
                'sleep(Speed.min)
                'para1 = WDoc.Paragraphs.Add()
                sleep(Speed.min)
                _table1 = WDoc.Tables.Add(para1.Range, BasicDetailDGV.Rows.Count, BasicDetailDGV.Columns.GetColumnCount(DataGridViewElementStates.Displayed))
                sleep(Speed.min)
                _table1.Borders.Enable = False
                sleep(Speed.min)
                _table1.Range.Font.Bold = False
                sleep(Speed.min)
                _table1.Cell(1, 1).Range.Text = BasicDetailDGV.Columns(1).HeaderText
                sleep(Speed.min)
                _table1.Cell(1, 2).Range.Text = BasicDetailDGV.Columns(0).HeaderText
                sleep(Speed.min)
                _table1.Cell(1, 3).Range.Text = BasicDetailDGV.Columns(2).HeaderText
                sleep(Speed.min)
                _table1.Cell(1, 4).Range.Text = BasicDetailDGV.Columns(3).HeaderText
                For x% = 0 To BasicDetailDGV.Rows.Count - 2
                    sleep(Speed.min)
                    _table1.Cell(x + 2, 1).Range.Text = BasicDetailDGV.Rows(x).Cells(1).Value
                    sleep(Speed.min)
                    _table1.Cell(x + 2, 2).Range.Text = BasicDetailDGV.Rows(x).Cells(0).Value
                    sleep(Speed.min)
                    _table1.Cell(x + 2, 3).Range.Text = BasicDetailDGV.Rows(x).Cells(2).Value
                    sleep(Speed.min)
                    _table1.Cell(x + 2, 4).Range.Text = BasicDetailDGV.Rows(x).Cells(3).Value
                Next
                sleep(Speed.min)
                _table1.Rows(1).Range.Bold = True
                'sleep(Speed.min)
                '_table1.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                sleep(Speed.min)
                _table1.Borders.Enable = True
                sleep(Speed.min)
                _table1.Range.ParagraphFormat.SpaceAfter = 0
                'For x% = 0 To Basic_details(1).Count - 1
                '    If Basic_details(1).Item(x) Is Nothing Then
                '        sleep(Speed.min)
                '        _table1.Cell(x + 1, 1).Delete()
                '    Else
                '        sleep(Speed.min)
                '        _table1.Cell(x + 1, 1).Range.Text = vbTab & Basic_details(1).Item(x).ToString.ToUpper()
                '    End If
                'Next
            End If
            sleep(Speed.min)
            _table1.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)
            sleep(Speed.min)
            _table1.Columns(4).SetWidth(40, Word.WdRulerStyle.wdAdjustNone)
            sleep(Speed.min)
            _table1.Rows(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            sleep(Speed.min)
            _table1.Rows(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            sleep(Speed.min)
            _table1.Rows(1).Alignment = Word.WdRowAlignment.wdAlignRowCenter
            sleep(Speed.min)
            para1 = WDoc.Paragraphs.Add(WDoc.Bookmarks.Item("\endofdoc").Range)
            sleep(Speed.min)
            para1.Range.Text = "2. DESIGN DATA: "
            sleep(Speed.min)
            para1.Style = Word.WdBuiltinStyle.wdStyleHeading1
            sleep(Speed.min)
            para1.Range.Font.Name = "Segoe UI"
            sleep(Speed.min)
            para1.Range.Font.Size = fontSize
            sleep(Speed.min)
            para1.Range.Font.Color = Word.WdColor.wdColorBlack
            sleep(Speed.min)
            para1.Range.Font.Bold = True
            WriteLog("", "Page Setup completed!")
            Marshal.ReleaseComObject(para1)
            Marshal.ReleaseComObject(_table1)
            Marshal.ReleaseComObject(headerTable)
            Marshal.ReleaseComObject(footerTable)

        Catch ex As Exception
            MsgBox("Page Setup Error!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            WriteLog("", "Page Setup Error!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try
        WApp.Visible = True
        WApp.WindowState = Word.WdWindowState.wdWindowStateMaximize

        Return WDoc

#Region "Old_Code"

        'Dim Range As Word.Range = WDoc.Sections(1).Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range
        'Dim headerTable As Word.Table = WDoc.Tables.Add(Range, 2, 4)
        'headerTable.Borders.Enable = True
        'With headerTable
        '    .Cell(2, 1).Range.Text = "Date: " & System.DateTime.Now.ToString("dd-MM-yyyy")
        '    .Cell(2, 3).Range.Text = "Revision: X"
        '    .Cell(2, 3).Range.Font.Bold = True
        '    .Rows(2).Cells(1).Merge(headerTable.Rows(2).Cells(2))
        '    .Rows(2).Cells(2).Merge(headerTable.Rows(2).Cells(3))
        '    .Cell(1, 1).Range.Text = "Project No.: "
        '    .Cell(1, 2).Range.Text = "I.B.R Calculations for XXXXXX"
        '    .Cell(1, 2).Range.Font.Bold = True
        '    .Rows(1).Cells(2).Merge(headerTable.Rows(1).Cells(3))
        '    .Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        '    .Cell(1, 3).Range.Text = "Boiler No.: "
        'End With

        'Range = WDoc.Sections(1).Footers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range
        'headerTable = WDoc.Tables.Add(Range, 2, 3)
        'headerTable.Borders.Enable = True
        'WDoc.PageSetup.FooterDistance = WApp.CentimetersToPoints(0.5)
        'With headerTable
        '    .Cell(1, 1).Range.Text = "Prepared by: "
        '    .Cell(1, 2).Range.Text = "Checked by: "
        '    .Cell(1, 3).Range.Text = "Approved by: "
        '    .Rows(2).Cells(1).Merge(headerTable.Rows(2).Cells(3))
        '    'Dim pgRange As Word.Range = headerTable.Cell(2, 1).Range
        '    'pgRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        '    'pgRange.Fields.Add(pgRange, Word.WdFieldType.wdFieldPage)
        '    .Cell(2, 1).Select()
        '    WApp.Selection.Fields.Add(WApp.Selection.Range, Word.WdFieldType.wdFieldPage, "SDSD",)
        '    .Cell(2, 1).Range.Fields.Add(headerTable.Cell(1, 1).Range, Word.WdFieldType.wdFieldPage)



        'End With

#End Region
    End Function

    ''' <summary>
    ''' This particular block will forcefull end up all word instances without any time to shut down or cleanup
    ''' This is to avoid 'Call rejected by Callee'
    ''' </summary>
    Private Shared Sub killWordproc()

        Try
            WriteLog("", "Closing all active Word Instances")
            Dim wordInstances() As Process = Process.GetProcessesByName("WINWORD")
            For Each instance As Process In wordInstances
                instance.Kill()
            Next
            WriteLog("", "All active Word Instances are closed")
        Catch ex As Exception
            WriteLog("", "Failed to close all active Word Instances")
        End Try

    End Sub
    ''' <summary>
    ''' To serialize a document
    ''' </summary>
    Public Shared Sub Serialize()

        Try
            Dim filePath As String = "D:\OneDrive - Thermax Limited\_Main_Projects_\IBR Calculations\Support Docs\1.docx"

            Dim fileBytes As Byte() = File.ReadAllBytes(filePath)

            Dim formatter As New BinaryFormatter()
            Using stream As New MemoryStream()
                formatter.Serialize(stream, fileBytes)
                Dim base64String As String = Convert.ToBase64String(stream.ToArray())
                Console.WriteLine("Serialized Word Document as Base64 String: " & base64String)
            End Using

        Catch ex As Exception

        End Try

    End Sub



    ''' <summary>
    ''' To convert the .NET data grid view control to MS Word Table
    ''' </summary>
    ''' <param name="DGV"> Data grid view control to be converted </param>
    ''' <param name="columnstoOmit"> Indexes of Columns that needs to be neglected. 0 is the starting value </param>
    ''' <param name="remarkIndex"> Remark Notes index, includes highlighting too. </param>
    ''' <returns> table </returns>
    Public Shared Function DGVtoTable(DGV As DataGridView, Optional columnstoOmit() As Integer = Nothing, Optional remarkIndex As Integer = Nothing)

        WriteLog("Data Grid View to Word Table", "")
        Try
            WriteLog("", "Initiating Data Grid View to Word Table" & vbCrLf & vbCrLf & "Highlighting: " & remarkIndex.ToString())

            sleep(Speed.max)
            Dim range As Word.Range = WDoc.Range
            'range.Font.Bold = False
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            'range.Font.Bold = False
            Dim xg As Integer = 0
            'Dim xxy As Integer = 1 / xg
            Dim table As Word.Table
            If columnstoOmit Is Nothing Then
                table = WDoc.Tables.Add(range, DGV.Rows.Count + 1, DGV.Columns.GetColumnCount(DataGridViewElementStates.None))
            Else
                table = WDoc.Tables.Add(range, DGV.Rows.Count + 1, DGV.Columns.GetColumnCount(DataGridViewElementStates.None) - columnstoOmit.Length)
            End If

            With table
                sleep(Speed.min)
                .Columns.Borders.Enable = True
                sleep(Speed.min)
                .Range.Font.Bold = False
                sleep(Speed.min)
                .Range.Font.Size = fontSize
                sleep(Speed.min)
                .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                sleep(Speed.min)
                .Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
            End With

            sleep(Speed.min)
            WApp.ActiveWindow.ScrollIntoView(range)

            'This code block is to centerize the table in an easy (visual) way. Uncomment the next commented code for Conventional method
            sleep(Speed.max)
            table.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter


            sleep(Speed.max)
            table.Rows(1).Range.Font.Bold = True

            If columnstoOmit Is Nothing Then
                For f% = 0 To DGV.Columns.GetColumnCount(DataGridViewElementStates.None) - 1
                    sleep(Speed.min)
                    table.Cell(1, f + 1).Range.Text = DGV.Columns(f).HeaderText
                Next
            Else
                Dim colIndex = 0
                For f% = 0 To DGV.Columns.GetColumnCount(DataGridViewElementStates.None) - 1
                    sleep(Speed.min)
                    If Not columnstoOmit.Contains(f) Then
                        sleep(Speed.min)
                        table.Cell(1, colIndex + 1).Range.Text = DGV.Columns(f).HeaderText
                        colIndex += 1
                    End If
                Next
            End If

            'For row values
            sleep(Speed.max)
            If columnstoOmit Is Nothing Then
                Dim colIndex = 0
                For x% = 0 To DGV.Rows.Count - 1
                    For y% = 0 To DGV.Columns.GetColumnCount(DataGridViewElementStates.None) - 1
                        sleep(Speed.min)
                        table.Cell(x + 2, y + 1).Range.Text = DGV.Rows(x).Cells(y).Value.ToString()
                        If remarkIndex <> Nothing Then
                            If y = remarkIndex Then
                                sleep(Speed.min)
                                If Not DGV.Rows(x).Cells(y).Value.ToString().StartsWith("Safe", True, Globalization.CultureInfo.CurrentCulture) Then
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
                    For y% = 0 To DGV.Columns.GetColumnCount(DataGridViewElementStates.None) - 1
                        If Not columnstoOmit.Contains(y) Then
                            sleep(Speed.max)
                            table.Cell(x + 2, colIndex + 1).Range.Text = DGV.Rows(x).Cells(y).Value.ToString()
                            If remarkIndex <> Nothing Then
                                If y = remarkIndex Then
                                    sleep(Speed.min)
                                    If Not DGV.Rows(x).Cells(y).Value.ToString().StartsWith("Safe", True, Globalization.CultureInfo.CurrentCulture) Then
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


            sleep(Speed.min)
            table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)
            sleep(Speed.min)
            table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)
            sleep(Speed.min)
            table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
            sleep(Speed.min)
            table.Range.ParagraphFormat.SpaceAfter = 0
            sleep(Speed.min)
            table.Range.ParagraphFormat.SpaceBefore = 0
            Marshal.ReleaseComObject(range)
            WriteLog("", "Data Grid View to Word Table completed")

            'Marshal.ReleaseComObject(table)
            'If Err.Number <> 0 Then
            '    WriteLog("", "Data Grid View to Word Table Error!" & vbCrLf & Err.Description & vbCrLf & Err.Source)
            '    'MsgBox("Data Grid View to Word Table Error!" & vbCrLf & Err.Description & vbCrLf & Err.Source)
            'Else
            '    WriteLog("", "Data Grid View to Word Table completed")
            'End If
            Return table

        Catch ex As Exception

            'MsgBox("Data Grid View to Word Table Error!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            WriteLog("", "Data Grid View to Word Table Error!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
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
    Public Shared Function addLine(content As String, Optional addAnotherLine As Boolean = False, Optional bold As Boolean = False, Optional addNewPage As Boolean = False, Optional italic As Boolean = False, Optional underLine As Boolean = False, Optional addTab As Boolean = False, Optional heading1 As Boolean = False)
        WriteLog("Adding Line to Document: ", "")
        Try
            WriteLog("", "Adding Line initiated...")
            sleep(Speed.min)
            Dim para1 As Word.Paragraph
            If addNewPage Then
                WDoc.Range(WDoc.Range.StoryLength - 1).InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)
            End If
            WDoc.Paragraphs.Add()
            sleep(Speed.min)
            para1 = WDoc.Paragraphs.Add(WDoc.Bookmarks.Item("\endofdoc").Range)
            sleep(Speed.min)
            para1.Range.Font.Bold = False
            sleep(Speed.min)
            WApp.ActiveWindow.ScrollIntoView(para1.Range)
            If addAnotherLine Then
                para1 = WDoc.Paragraphs.Add(WDoc.Bookmarks.Item("\endofdoc").Range)
            End If
            WApp.ActiveWindow.ScrollIntoView(para1.Range)
            With para1
                sleep(Speed.min)
                If addTab Then
                    .Range.Text = vbTab & content
                Else
                    .Range.Text = content
                End If
                If heading1 Then
                    sleep(Speed.min)
                    .Style = Word.WdBuiltinStyle.wdStyleHeading1
                    sleep(Speed.min)
                    .Range.Font.Name = "Segoe UI"
                    sleep(Speed.min)
                    .Range.Font.Size = fontSize
                    sleep(Speed.min)
                    .Range.Font.Color = Word.WdColor.wdColorBlack
                End If
                If .Range.Font.Name <> "Segoe UI" Or .Range.Font.Size <> fontSize Then
                    sleep(Speed.min)
                    .Range.Font.Name = "Segoe UI"
                    sleep(Speed.min)
                    .Range.Font.Size = fontSize
                End If
                sleep(Speed.min)
                .Range.Font.Bold = bold
                sleep(Speed.min)
                .Range.Font.Italic = italic
                sleep(Speed.min)
                .Range.Font.Underline = underLine
            End With
            WriteLog("", "Line Added")
            WApp.ScreenRefresh()
            Return para1

        Catch ex As Exception
            'MsgBox("Failed to add line!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            WriteLog("", "Failed to add line!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            Return Nothing

        End Try
    End Function


    ''' <summary>
    ''' To append text to the last paragraph of the document.
    ''' </summary>
    ''' <param name="content"> Text to append </param>
    ''' <returns> WDoc </returns>
    Public Shared Function appendText(content As String)

        WriteLog("Appending Text to Document: ", "")
        Try
            WriteLog("", "Appending Text initiated...")
            sleep(Speed.min)
            Dim para1 As Word.Paragraph
            para1 = WDoc.Paragraphs(WDoc.Paragraphs.Count)
            Dim txt$ = para1.Range.Text
            WApp.ActiveWindow.ScrollIntoView(para1.Range)
            'To remove the para break at the end
            txt = Left(txt, Len(txt) - 1)
            txt = txt & content
            para1.Range.Text = txt
            WriteLog("", "Text appended")
        Catch ex As Exception
            'MsgBox("Failed to append text!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            WriteLog("", "Failed to append text!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
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
    Public Shared Function findnModify(stringtoFind As String, Optional bold As Boolean = False, Optional italic As Boolean = False, Optional underLine As Boolean = False, Optional highlight As Boolean = False, Optional fontColor As Word.WdColor = Word.WdColor.wdColorAutomatic, Optional fontSize As Single = fontSize)

        WriteLog("Find and Modify: ", "")
        Try
            WriteLog("", "Find and Modify Text initiated...")
            sleep(Speed.min)
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
                    Thread.Sleep(Speed.min / 2)
                    Range.Bold = bold
                    Thread.Sleep(Speed.min / 2)
                    Range.Italic = italic
                    Thread.Sleep(Speed.min / 2)
                    Range.Underline = underLine
                    Thread.Sleep(Speed.min / 2)
                    Range.Font.Color = fontColor
                    Thread.Sleep(Speed.min / 2)
                    Range.Font.Size = fontSize
                    sleep(Speed.min)
                    If highlight Then
                        Range.HighlightColorIndex = Word.WdColorIndex.wdYellow
                    End If
                    .Execute()
                Loop
            End With
            WriteLog("", "Found and Modified")
        Catch ex As Exception
            'MsgBox("Failed to Find and Modify!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            WriteLog("", "Failed to Find and Modify!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
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
        WriteLog("Adding formula to Document: ", "")
        Try
            WriteLog("", "Adding formula initiated...")
            sleep(Speed.min)
            Dim range As Word.Range
            range = WDoc.Bookmarks.Item("\endofdoc").Range
            WApp.ActiveWindow.ScrollIntoView(range)
            sleep(Speed.min)
            range.InsertParagraphAfter()
            sleep(Speed.min)
            WApp.ActiveWindow.ScrollIntoView(range)
            sleep(Speed.min)
            range.InsertAfter(formula)
            sleep(Speed.min)
            range.Bold = bold
            sleep(Speed.min)
            range = WDoc.OMaths.Add(range)
            sleep(Speed.min)
            range.OMaths.BuildUp()
            sleep(Speed.min)
            range.Text = range.Text & Text
            sleep(Speed.min)
            'Dim para1 As Word.Paragraph
            'sleep(Speed.min)
            'para1 = range.Paragraphs(1)
            'sleep(Speed.min)
            'para1.Range.Font.Bold = bold
            WriteLog("", "Formula added.")
        Catch ex As Exception
            'MsgBox("Failed to add formula!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            WriteLog("", "Failed to add formula!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try
        Return WDoc
    End Function


    ''' <summary>
    ''' To add a blank page
    ''' </summary>
    Public Shared Sub nextPage()
        WriteLog("Adding Blank page to Document: ", "")
        Try
            sleep(Speed.min)
            WDoc.Range(WDoc.Range.StoryLength - 1).InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)
            WriteLog("", "Blank page added.")
        Catch ex As Exception
            'MsgBox("Failed to add Blank Page!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            WriteLog("", "Failed to add Blank Page!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try
    End Sub


    ''' <summary>
    ''' To add a tab at End of the Document
    ''' </summary>
    Public Shared Sub addTabatEOD(Optional tabCount As Integer = 1)
        WriteLog("Adding tab to End of Document: ", "")
        Try
            Dim range As Word.Range
            sleep(Speed.min)
            range = WDoc.Bookmarks.Item("\endofdoc").Range

            sleep(Speed.min)
            range.InsertBefore(vbTab)
            sleep(Speed.min)
            WApp.ActiveWindow.ScrollIntoView(range)
            WriteLog("", "Tab to End of Document added.")
        Catch ex As Exception
            'MsgBox("Failed to add Tab to End of Document added!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            WriteLog("", "Failed to add Tab to End of Document added!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try
    End Sub


    ''' <summary>
    ''' To Insert a text watermark
    ''' </summary>
    ''' <param name="watermarkText"> Text </param>
    Public Shared Sub InsertWatermark(watermarkText As String)

        WriteLog("Adding watermark to Document: ", "")
Genesis:
        Try
            WriteLog("", "Adding watermark initiated...")

            Dim headerRange As Word.HeaderFooter
            Dim section As Word.Section
            sleep(Speed.min)
            section = WDoc.Sections(1)
            sleep(Speed.min)
            headerRange = section.Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary)
            'For Each shape As Word.Shape In headerRange.Shapes
            '    shape.Delete()
            'Next
            Dim shapeRange As Word.Shape
            sleep(Speed.min)
            shapeRange = headerRange.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect1, watermarkText, "Arial", 45, False, False, 0, 0)
            'shapeRange = headerRange.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 1000, 400)
            With shapeRange
                sleep(Speed.min)
                .Line.Visible = MsoTriState.msoTrue
                'Thread.Sleep(15)
                '.TextFrame.TextRange.Text = watermarkText
                'Thread.Sleep(15)
                '.TextFrame.TextRange.Font.Size = 60
                'Thread.Sleep(15)
                '.TextFrame.TextRange.Font.Bold = True
                'Thread.Sleep(15)
                '.TextFrame.TextRange.Font.Name = "Calibri"
                sleep(Speed.min)
                .Rotation = -45
                sleep(Speed.min)
                .Rotation = -45
                'Thread.Sleep(15)
                '.TextFrame.TextRange.Font.ColorIndex = Word.WdColorIndex.wdRed
                'Thread.Sleep(15)
                '.TextFrame.TextRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                sleep(Speed.min)
                '.Top = Word.WdShapePosition.wdShapeBottom
                .Top = 220
                .Left = 75
                '.Left = Word.WdShapePosition.wdShapeRight
            End With
            WriteLog("", "Watermark added.")

        Catch ex As Exception
            'MsgBox("Failed to add watermark!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            WriteLog("", "Failed to add watermark!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try


#Region "Old_Code"
        'shapeRange.Select()
        'With WApp
        '    '.Selection.ShapeRange.Group.Select(shapeRange)
        '    .Selection.ShapeRange.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
        '    '.Selection.ShapeRange.Height = .InchesToPoints(2.5)
        '    '.Selection.ShapeRange.Width = .InchesToPoints(10)
        '    .Selection.ShapeRange.WrapFormat.AllowOverlap = -1
        '    .Selection.ShapeRange.WrapFormat.Side = Microsoft.Office.Interop.Word.WdWrapSideType.wdWrapBoth
        '    .Selection.ShapeRange.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapNone
        '    .Selection.ShapeRange.RelativeHorizontalPosition = Microsoft.Office.Interop.Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        '    .Selection.ShapeRange.RelativeVerticalPosition = Microsoft.Office.Interop.Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
        '    .Selection.ShapeRange.Left = Microsoft.Office.Interop.Word.WdShapePosition.wdShapeCenter
        '    .Selection.ShapeRange.Top = Microsoft.Office.Interop.Word.WdShapePosition.wdShapeCenter
        'End With


        'shapeRange.TextFrame.TextRange.Text = watermarkText
        'shapeRange.TextFrame.TextRange.Font.Size = 70
        'shapeRange.TextFrame.TextRange.Font.Name = "Calibri"
        'shapeRange.Rotation = -45
        'shapeRange.TextFrame.TextRange.Font.ColorIndex = Word.WdColorIndex.wdRed
        'shapeRange.TextFrame.TextRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        'shapeRange.TextFrame.TextRange.ParagraphFormat.Borders.Enable = False
        'shapeRange.Line.Visible = MsoTriState.msoFalse
        'shapeRange.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
        'shapeRange.WrapFormat.AllowOverlap = -1
        'shapeRange.WrapFormat.Side = Microsoft.Office.Interop.Word.WdWrapSideType.wdWrapBoth
        'shapeRange.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapNone
        'shapeRange.RelativeHorizontalPosition = Microsoft.Office.Interop.Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        'shapeRange.RelativeVerticalPosition = Microsoft.Office.Interop.Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
        ''shapeRange.Left = Microsoft.Office.Interop.Word.WdShapePosition.wdShapeLeft
        ''shapeRange.Left = 30
        'shapeRange.Top = 150

#End Region
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

        WriteLog("Adding table to Document: ", "")

        Dim table As Word.Table
        Try
            WriteLog("", "Rows x Column = " & rows & " x " & "Adding table initiated...")
            Dim range As Word.Range
            sleep(Speed.min)
            range = WDoc.Bookmarks.Item("\endofdoc").Range
            WApp.ActiveWindow.ScrollIntoView(range)
            sleep(Speed.min)
            Dim para1 As Word.Paragraph
            sleep(Speed.min)
            WDoc.Paragraphs.Add()
            sleep(Speed.min)
            para1 = WDoc.Paragraphs.Add(range)
            sleep(Speed.min)
            para1.Range.Text = vbNewLine
            sleep(Speed.min)
            table = WDoc.Tables.Add(para1.Range, rows, columns)
            If autofit Then
                sleep(Speed.min)
                table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)
            End If
            With table
                If horizontallyCenter Then
                    sleep(Speed.min)
                    .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                Else
                    sleep(Speed.min)
                    .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                End If
                sleep(Speed.min)
                .Range.ParagraphFormat.SpaceAfter = 0
                sleep(Speed.min)
                .Borders.Enable = enableBorder
            End With

            sleep(Speed.min)
            range = table.Range
            With range
                If verticallyCenter Then
                    sleep(Speed.min)
                    .Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                Else
                    sleep(Speed.min)
                    .Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
                End If
                '.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            End With

            sleep(Speed.min)
            range = table.Rows(1).Range
            sleep(Speed.min)
            range.Font.Bold = boldHeader

            WriteLog("", "Table added.")

            Return table
        Catch ex As Exception
            'MsgBox("Failed to add table!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            WriteLog("", "Failed to add table!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            Return Nothing

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

        WriteLog("Adding values to Column: ", "")
        Try
            WriteLog("", "Column Index = " & columnIndex & " x " & "Adding values to column initiated...")
            With table
                For x% = 1 To values.Length
                    sleep(Speed.min)
                    .Cell(x, columnIndex).Range.Text = values(x - 1)
                Next
            End With
            WriteLog("", "Values added to Column.")
            Return table
        Catch ex As Exception
            Return table
            'MsgBox("Failed to add values to Column!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            WriteLog("", "Failed to add values to Column!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
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

        WriteLog("Adding values to Row: ", "")
        Try
            WriteLog("", "Row Index = " & rowIndex & " x " & "Adding values to row initiated...")
            With table
                For x% = 1 To values.Length
                    sleep(Speed.min)
                    .Cell(rowIndex, x).Range.Text = values(x - 1)
                Next
            End With
            WriteLog("", "Values added to Row.")
            Return table
        Catch ex As Exception
            Return table
            'MsgBox("Failed to add values to Row!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            WriteLog("", "Failed to add values to Row!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
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
            sleep(Speed.min)
            table.Columns.Add()
        Next
        Return table

    End Function

    'Incomplete
    Public Shared Sub splitTable(table As Word.Table)

        If table.Range.Information(Word.WdInformation.wdWithInTable) Then
            Dim range As Word.Range = table.Range
            Dim endPage As Integer = range.Information(Word.WdInformation.wdActiveEndPageNumber)
            range.Collapse(Word.WdCollapseDirection.wdCollapseStart)
            range.MoveEnd(Word.WdUnits.wdParagraph, 1)
            Dim startPage As Integer = range.Information(Word.WdInformation.wdActiveEndPageNumber)

            If startPage < endPage Then
                'Dim firstRow As Word.Row = table.Rows(1)
                'Dim firstRowRange As Word.Range = firstRow.Range
                'fi
                Dim pg As Word.Range = WDoc.Range(endPage)
                Dim tbl As Word.Table = pg.Tables(1)
                tbl.Split(5)
            End If
        End If

    End Sub


    ''' <summary>
    ''' To make a word table from given List of List(s)
    ''' </summary>
    ''' <param name="lists">List of List(s). See example</param>
    ''' <param name="enableBorder"> To enable all borders </param>
    ''' <param name="boldHeader"> To make the first row bold </param>
    ''' <param name="horizontallyCenter"> To centerize the cells horizontally </param>
    ''' <param name="verticallyCenter"> To centerize the cells vertically </param>
    ''' <param name="autofit"> To autofit to the content </param>
    ''' <param name="centerize"> To centerize the whole table </param>
    ''' <example>
    ''' <code>
    ''' Dim HeaderCalc As New List(Of List(Of String)) From {
    ''' New List(Of String) From {"S.No.", "Name", "Remarks"},
    ''' New List(Of String) From {1, "Kalidash Palaniappan","Q"}
    ''' }
    ''' </code>
    ''' </example>
    ''' <returns>Filled word Table</returns>
    Public Shared Function createTablefromLists(lists As List(Of List(Of String)), Optional enableBorder As Boolean = True, Optional boldHeader As Boolean = True, Optional horizontallyCenter As Boolean = True, Optional verticallyCenter As Boolean = True, Optional autofit As Boolean = True, Optional centerize As Boolean = False, Optional remarkIndex As Integer = Nothing)

        WriteLog("Adding table to Document: ", "")
        Try
            addLine("")
            Dim rowCount% = lists.Count
            Dim colCount% = If(rowCount > 0, lists(0).Count, 0)
            Dim range As Word.Range
            sleep(Speed.min)
            range = WDoc.Bookmarks.Item("\endofdoc").Range
            WApp.ActiveWindow.ScrollIntoView(range)
            sleep(Speed.min)
            Dim para1 As Word.Paragraph = WDoc.Paragraphs.Add(range)
            range = para1.Range
            sleep(Speed.min)
            range.Font.Bold = False
            sleep(Speed.min)
            Dim table As Word.Table = WDoc.Tables.Add(range, rowCount, colCount)

            For i As Integer = 1 To rowCount
                For j As Integer = 1 To colCount
                    sleep(Speed.min)
                    table.Cell(i, j).Range.Text = lists(i - 1)(j - 1)
                    If remarkIndex <> Nothing Then
                        If j = remarkIndex AndAlso i <> 1 Then
                            sleep(Speed.min)
                            If Not lists(i - 1)(j - 1).StartsWith("Safe", True, Globalization.CultureInfo.CurrentCulture) Then
                                errorList.Add(New Tuple(Of String, String, String)(table.Cell(i, 1).Range.Text, table.Cell(i, 2).Range.Text, lists(i - 1)(j - 1)))
                                sleep(Speed.med)
                                table.Cell(i, j).Shading.BackgroundPatternColor = Word.WdColor.wdColorYellow
                                sleep(Speed.med)
                                table.Cell(i, j).Range.Font.Bold = True
                                sleep(Speed.med)
                                table.Cell(i, j).Range.Font.Color = Word.WdColor.wdColorRed
                            End If

                            'If Not DGV.Rows(i).Cells(y).Value.ToString().StartsWith("Safe", True, Globalization.CultureInfo.CurrentCulture) Then
                            '        errorList.Add(New Tuple(Of String, String, String)((x + 1).ToString, DGV.Rows(x).Cells(1).Value.ToString(), DGV.Rows(x).Cells(y).Value.ToString()))
                            '        table.Cell(i + 2, colIndex + 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorYellow
                            '        table.Cell(i + 2, colIndex + 1).Range.Font.Bold = True
                            '        table.Cell(i + 2, colIndex + 1).Range.Font.Color = Word.WdColor.wdColorRed
                            '    End If
                        End If
                        'colIndex += 1
                    End If

                Next
            Next
            With table

                If autofit Then
                    sleep(Speed.min)
                    .AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)
                End If
                If centerize Then
                    sleep(Speed.min)
                    .Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
                End If
                If horizontallyCenter Then
                    sleep(Speed.min)
                    .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                Else
                    sleep(Speed.min)
                    .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                End If
                sleep(Speed.min)
                .Range.ParagraphFormat.SpaceAfter = 0
                sleep(Speed.min)
                .Borders.Enable = enableBorder

            End With

            sleep(Speed.min)
            range = table.Range
            With range
                If verticallyCenter Then
                    sleep(Speed.min)
                    .Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                Else
                    sleep(Speed.min)
                    .Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
                End If
                '.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            End With

            If boldHeader Then
                sleep(Speed.min)
                range = table.Rows(1).Range
                sleep(Speed.min)
                range.Font.Bold = True
                'range.ParagraphFormat.SpaceAfterAuto = 1
            End If

            WriteLog("", "Table added.")
            Return table
        Catch ex As Exception
            'MsgBox("Failed to add table!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            WriteLog("", "Failed to add table!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
            Return Nothing
        End Try

    End Function

    'incomplete
    ''' <summary>
    '''
    ''' </summary>
    ''' <param name="RHS"></param>
    ''' <param name="LHS"></param>
    Public Shared Sub createTablewithFormulafromLists(RHS As List(Of String), LHS As String)
        WriteLog("Adding formula table to Document: ", "")
        Try
            addLine("")
            Dim rowCount As Integer = RHS.Count
            Dim colCount As Integer = 2
            Dim range As Word.Range = WDoc.Bookmarks.Item("\endofdoc").Range
            WApp.ActiveWindow.ScrollIntoView(range)

            Dim para1 As Word.Paragraph = WDoc.Paragraphs.Add(range)
            range = para1.Range
            range.Font.Bold = False

            Dim table As Word.Table = WDoc.Tables.Add(range, rowCount, colCount)
            table.Columns(2).Cells.Merge()
            For i% = 1 To rowCount
                If RHS(i - 1).StartsWith("$$$") AndAlso RHS(i - 1).EndsWith("$$$") Then
                    table.Cell(i, 1).Range.Text = RHS(i - 1).Replace("$$$", "")
                    WDoc.OMaths.Add(table.Cell(i, 1).Range)
                    table.Cell(i, 1).Range.OMaths.BuildUp()
                    table.Cell(i, 1).Range.Bold = True
                Else
                    table.Cell(i, 1).Range.Text = RHS(i - 1)
                End If
            Next

            table.Cell(1, 2).Range.Text = LHS
            With table
                .AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)
                .Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
                '.Range.ParagraphFormat.SpaceAfter = 0
            End With
            table.Cell(1, 2).Range.ParagraphFormat.SpaceAfter = 0
            'table.Cell(1, 2).Borders.Enable = True
            table.Columns(2).Borders(Word.WdBorderType.wdBorderLeft).LineStyle = Word.WdLineStyle.wdLineStyleDashDot
        Catch ex As Exception

        End Try

    End Sub

    Public Shared Sub createTablewithFormulafromLists(lists As List(Of List(Of String)), Optional enableBorder As Boolean = True, Optional boldHeader As Boolean = True, Optional horizontallyCenter As Boolean = True, Optional verticallyCenter As Boolean = True, Optional autofit As Boolean = True, Optional centerize As Boolean = False)
        WriteLog("Adding table to Document: ", "")
        Try
            addLine("")
            Dim rowCount As Integer = lists.Count
            Dim colCount As Integer = If(rowCount > 0, lists(0).Count, 0)
            Dim range As Word.Range = WDoc.Bookmarks.Item("\endofdoc").Range
            WApp.ActiveWindow.ScrollIntoView(range)

            Dim para1 As Word.Paragraph = WDoc.Paragraphs.Add(range)
            range = para1.Range
            range.Font.Bold = False

            Dim table As Word.Table = WDoc.Tables.Add(range, rowCount, colCount)

            For i As Integer = 1 To rowCount
                For j As Integer = 1 To colCount
                    Dim cellText As String = lists(i - 1)(j - 1)
                    If cellText.Contains("$$$") Then
                        Dim formula As String = cellText.Substring(3, cellText.Length - 6)
                        ' table.Cell(i, j).Range.Formula = formula
                        table.Cell(i, j).Range.Text = (formula)
                        WDoc.OMaths.Add(table.Cell(i, j).Range)
                        table.Cell(i, j).Range.OMaths.BuildUp()

                    Else
                        table.Cell(i, j).Range.Text = cellText
                    End If
                Next
            Next

            With table
                If autofit Then .AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)
                If centerize Then .Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter
                .Range.ParagraphFormat.Alignment = If(horizontallyCenter, Word.WdParagraphAlignment.wdAlignParagraphCenter, Word.WdParagraphAlignment.wdAlignParagraphLeft)
                .Range.ParagraphFormat.SpaceAfter = 0
                .Borders.Enable = enableBorder
            End With

            range = table.Range
            range.Cells.VerticalAlignment = If(verticallyCenter, Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter, Word.WdCellVerticalAlignment.wdCellAlignVerticalTop)

            If boldHeader Then
                range = table.Rows(1).Range
                range.Font.Bold = True
            End If

            WriteLog("", "Table added.")
        Catch ex As Exception
            WriteLog("", "Failed to add table!" & vbCrLf & "Stack Trace- " & ex.StackTrace & vbCrLf & "Error string- " & ex.ToString)
        End Try
    End Sub


    ''' <summary>
    ''' To find whether the selective cells is not having specified remarkWord. If the cell value is not starting with remarkWord then it will be added to errorList.
    ''' </summary>
    ''' <param name="table"> Specified Word table to find errors </param>
    ''' <param name="errorList"> Error list to add errors </param>
    ''' <param name="remarkIndex"> Index where there is possibilities to error. Index starts with 1 </param>
    ''' <param name="remarkWord"> The specific word to be there in remarkIndex. If not starting with this word, the row will be added to the errorList </param>
    Public Shared Sub finderrorsinTable(table As Object, errorList As List(Of Tuple(Of String, String, String)), Optional remarkIndex As Integer = Nothing, Optional remarkWord As String = "Safe", Optional verticalTable As Boolean = False)
        Try
            If verticalTable Then
                Dim wordTable As Word.Table = table
                If remarkIndex = Nothing Then
                    remarkIndex = wordTable.Rows.Count
                End If

                ''1st Column will be header of the table. So ignored.
                For x% = 2 To wordTable.Columns.Count

                    If Not wordTable.Columns(x).Cells(remarkIndex).Range.Text.StartsWith(remarkWord) Then
                        sleep(Speed.min)
                        errorList.Add(New Tuple(Of String, String, String)("", wordTable.Columns(1).Cells(remarkIndex).Range.Text, wordTable.Columns(x).Cells(remarkIndex).Range.Text))
                        sleep(Speed.min)
                        wordTable.Columns(x).Cells(remarkIndex).Shading.BackgroundPatternColor = Word.WdColor.wdColorYellow
                        sleep(Speed.min)
                        wordTable.Columns(x).Cells(remarkIndex).Range.Font.Bold = True
                        sleep(Speed.min)
                        wordTable.Columns(x).Cells(remarkIndex).Range.Font.Color = Word.WdColor.wdColorRed
                    End If

                Next

            Else
                Dim wordTable As Word.Table = table
                If remarkIndex = Nothing Then
                    remarkIndex = wordTable.Columns.Count
                End If

                ''1st row will be header of the table. So ignored.
                For x% = 2 To wordTable.Rows.Count

                    If Not wordTable.Rows(x).Cells(remarkIndex).Range.Text.StartsWith(remarkWord) Then
                        sleep(Speed.min)
                        errorList.Add(New Tuple(Of String, String, String)(wordTable.Rows(x).Cells(1).Range.Text, wordTable.Rows(x).Cells(2).Range.Text, wordTable.Rows(x).Cells(remarkIndex).Range.Text))
                        sleep(Speed.min)
                        wordTable.Rows(x).Cells(remarkIndex).Shading.BackgroundPatternColor = Word.WdColor.wdColorYellow
                        sleep(Speed.min)
                        wordTable.Rows(x).Cells(remarkIndex).Range.Font.Bold = True
                        sleep(Speed.min)
                        wordTable.Rows(x).Cells(remarkIndex).Range.Font.Color = Word.WdColor.wdColorRed
                    End If

                Next

            End If
        Catch ex As Exception

        End Try

    End Sub


#End Region

End Class

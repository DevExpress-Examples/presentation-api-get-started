Imports System.Drawing
Imports System.IO
Imports DevExpress.Docs.Presentation

Namespace DxPresentationGetStarted

    Public Class Program
        Public Shared Sub Main(__ As String())

            Dim presentation As Presentation = New Presentation()
            presentation.Slides.Clear()

            Dim slideMaster = presentation.SlideMasters(0)
            slideMaster.Background = New CustomSlideBackground(New SolidFill(Color.FromArgb(194, 228, 249)))

            Dim slide1 As Slide = New Slide(slideMaster.Layouts.Get(SlideLayoutType.Title))
            For Each shape As Shape In slide1.Shapes
                If shape.PlaceholderSettings.Type = PlaceholderType.CenteredTitle Then
                    shape.TextArea = New TextArea("Daily Testing Status Report")
                End If
                If shape.PlaceholderSettings.Type = PlaceholderType.Subtitle Then
                    shape.TextArea = New TextArea($"{Date.Now: dddd, MMMM d, yyyy}")
                End If
            Next
            presentation.Slides.Add(slide1)

            Dim slide2 As Slide = New Slide(slideMaster.Layouts.GetOrCreate(SlideLayoutType.Object))
            For Each shape As Shape In slide2.Shapes
                If shape.PlaceholderSettings.Type = PlaceholderType.Title Then
                    shape.TextArea = New TextArea("Today’s Highlights")
                End If
                If shape.PlaceholderSettings.Type = PlaceholderType.Body Then
                    Dim textArea As TextArea = New TextArea()
                    textArea.Paragraphs.Clear()
                    textArea.Paragraphs.Add(New TextParagraph("5 successful builds"))
                    textArea.Paragraphs.Add(New TextParagraph("2 failed builds"))
                    textArea.Paragraphs.Add(New TextParagraph("12 new bugs reported"))
                    textArea.Paragraphs.Add(New TextParagraph("3 deployments"))
                    textArea.Paragraphs.Add(New TextParagraph("1 rollback"))
                    shape.TextArea = textArea
                End If
            Next
            presentation.Slides.Add(slide2)

            Dim slide3 As Slide = New Slide(slideMaster.Layouts.GetOrCreate(SlideLayoutType.Object))
            For Each shape As Shape In slide3.Shapes.ToList()
                If shape.PlaceholderSettings.Type = PlaceholderType.Title Then
                    shape.TextArea = New TextArea("Build Status")
                End If
                If shape.PlaceholderSettings.Type = PlaceholderType.Body Then

                    Dim rect = presentation.GetActualShapeBounds(slide3, shape)
                    slide3.Shapes.Remove(shape)

                    Dim table As Table = New Table(5, 5, rect.X, rect.Y, rect.Width, rect.Height)
                    slide3.Shapes.Add(table)

                    table(0, 0).TextArea.Text = "Build ID"
                    table(0, 1).TextArea.Text = "Branch"
                    table(0, 2).TextArea.Text = "Status"
                    table(0, 3).TextArea.Text = "Duration"
                    table(0, 4).TextArea.Text = "Triggered By"

                    table(1, 0).TextArea.Text = "#5421"
                    table(1, 1).TextArea.Text = "main"
                    table(1, 2).TextArea.Text = "✅  Passed"
                    table(1, 3).TextArea.Text = "4m 30s"
                    table(1, 4).TextArea.Text = "Auto-schedule"

                    table(2, 0).TextArea.Text = "#5420"
                    table(2, 1).TextArea.Text = "ui-fix"
                    table(2, 2).TextArea.Text = "❌  Failed"
                    table(2, 3).TextArea.Text = "2m 18s"
                    table(2, 4).TextArea.Text = "Push by dev1"

                    table(3, 0).TextArea.Text = "#5419"
                    table(3, 1).TextArea.Text = "main"
                    table(3, 2).TextArea.Text = "✅  Passed"
                    table(3, 3).TextArea.Text = "3m 52s"
                    table(3, 4).TextArea.Text = "Auto-schedule"

                    table(4, 0).TextArea.Text = "#5418"
                    table(4, 1).TextArea.Text = "hotfix"
                    table(4, 2).TextArea.Text = "❌  Failed"
                    table(4, 3).TextArea.Text = "5m 1s"
                    table(4, 4).TextArea.Text = "Manual"

                    table.Style = New ThemedTableStyle(TableStyleType.LightStyle1)
                    table.HasBandedRows = False
                End If
            Next
            presentation.Slides.Add(slide3)

            presentation.HeaderFooterManager.AddSlideNumberPlaceholder(presentation.Slides)
            presentation.HeaderFooterManager.AddFooterPlaceholder(presentation.Slides, "ProductXCompany")

            Dim outputStream As FileStream = New FileStream("..\..\..\data\my-presentation.pptx", FileMode.Create)
            presentation.SaveDocument(outputStream)
            outputStream.Dispose()

            presentation.ExportToPdf(New FileStream("..\..\..\data\exported-document.pdf", FileMode.Create))
        End Sub
    End Class
End Namespace

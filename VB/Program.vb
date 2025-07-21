Imports DevExpress.Docs.Presentation
Imports DevExpress.Drawing
Imports System.Drawing
Imports System.IO

Namespace DxPresentationGetStarted

    Public Class Program

        Public Shared Sub Main(ByVal underscore() As String)

            Dim presentation As New Presentation()
            presentation.Slides.Clear()

            Dim slideMaster As SlideMaster = presentation.SlideMasters(0)
            slideMaster.Background = New CustomSlideBackground(New SolidFill(Color.FromArgb(194, 228, 249)))

            Dim slide1 As Slide = New Slide(slideMaster.Layouts.[Get](SlideLayoutType.Title))

            For Each shape As Shape In slide1.Shapes

                If shape.PlaceholderSettings.Type = PlaceholderType.CenteredTitle Then
                    Dim textArea As TextArea = New TextArea()
                    textArea.Text = "Daily Testing Status Report"
                    shape.TextArea = textArea
                End If

                If shape.PlaceholderSettings.Type = PlaceholderType.Subtitle Then
                    Dim textArea As TextArea = New TextArea()
                    textArea.Text = $"{DateTime.Now}"
                    shape.TextArea = textArea
                End If
            Next
            presentation.Slides.Add(slide1)

            Dim slide2 As New Slide(slideMaster.Layouts.GetOrCreate(SlideLayoutType.Object))
            For Each shape As Shape In slide2.Shapes
                If shape.PlaceholderSettings.Type = PlaceholderType.Title Then
                    Dim textArea As New TextArea()
                    textArea.Text = "Today’s Highlights"
                    shape.TextArea = textArea
                End If
                If shape.PlaceholderSettings.Type = PlaceholderType.Object Then
                    Dim textArea As New TextArea()
                    Dim paragraph1 As New TextParagraph()
                    paragraph1.Runs.Add(New TextRun With {.Text = "5 successful builds"})
                    textArea.Paragraphs.Add(paragraph1)

                    Dim paragraph2 As New TextParagraph()
                    paragraph2.Runs.Add(New TextRun With {.Text = "2 failed builds"})
                    textArea.Paragraphs.Add(paragraph2)

                    Dim paragraph3 As New TextParagraph()
                    paragraph3.Runs.Add(New TextRun With {.Text = "12 new bugs reported"})
                    textArea.Paragraphs.Add(paragraph3)

                    Dim paragraph4 As New TextParagraph()
                    paragraph4.Runs.Add(New TextRun With {.Text = "3 deployments"})
                    textArea.Paragraphs.Add(paragraph4)

                    Dim paragraph5 As New TextParagraph()
                    paragraph5.Runs.Add(New TextRun With {.Text = "1 rollback"})
                    textArea.Paragraphs.Add(paragraph5)
                    shape.TextArea = textArea
                End If
            Next shape
            presentation.Slides.Add(slide2)

            Dim slide3 As New Slide(slideMaster.Layouts.GetOrCreate(SlideLayoutType.Object))
            For Each shape As Shape In slide3.Shapes
                If shape.PlaceholderSettings.Type = PlaceholderType.Title Then
                    Dim textArea As New TextArea()
                    textArea.Text = "Build Status"
                    shape.TextArea = textArea
                End If
                If shape.PlaceholderSettings.Type = PlaceholderType.Object Then
                    Dim textArea As New TextArea()
                    textArea.Text = " "
                    shape.TextArea = textArea
                    Dim imagePath As String = "..\..\..\data\table.png"
                    Dim stream As Stream = New FileStream(imagePath, FileMode.Open, FileAccess.Read)
                    Dim fill As New PictureFill(DXImage.FromStream(stream))
                    fill.Stretch = True
                    shape.Fill = fill
                End If
            Next shape
            presentation.Slides.Add(slide3)

            presentation.HeaderFooterManager.AddSlideNumberPlaceholder(presentation.Slides)
            presentation.HeaderFooterManager.AddFooterPlaceholder(presentation.Slides, "ProductXCompany")

            Dim outputStream As New FileStream("..\..\..\data\my-presentation.pptx", FileMode.Create)
            presentation.SaveDocument(outputStream)
            outputStream.Dispose()

            presentation.ExportToPdf(New FileStream("..\..\..\data\exported-document.pdf", FileMode.Create))
        End Sub
    End Class
End Namespace
Imports DevExpress.Docs.Presentation
Imports DevExpress.Drawing
Imports System.Drawing
Imports System.IO

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
                If shape.PlaceholderSettings.Type = PlaceholderType.Object Then
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
            For Each shape As Shape In slide3.Shapes
                If shape.PlaceholderSettings.Type = PlaceholderType.Title Then
                    shape.TextArea = New TextArea("Build Status")
                End If
                If shape.PlaceholderSettings.Type = PlaceholderType.Object Then
                    shape.TextArea = New TextArea(" ")
                    Dim imagePath = "..\..\..\data\table.png"

                    Using stream As Stream = New FileStream(imagePath, FileMode.Open, FileAccess.Read)
                        Dim fill As PictureFill = New PictureFill(DXImage.FromStream(stream))
                        fill.Stretch = True
                        shape.Fill = fill
                    End Using

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

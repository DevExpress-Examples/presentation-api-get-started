using DevExpress.Docs.Presentation;
using DevExpress.Drawing;
using System.Drawing;

namespace DxPresentationGetStarted;

public class Program {
    public static void Main(string[] _) {

        Presentation presentation = new Presentation();
        presentation.Slides.Clear();

        SlideMaster slideMaster = presentation.SlideMasters[0];
        slideMaster.Background = new CustomSlideBackground(new SolidFill(Color.FromArgb(194, 228, 249)));

        Slide slide1 = new Slide(slideMaster.Layouts.Get(SlideLayoutType.Title));
        foreach (Shape shape in slide1.Shapes) {
            if (shape.PlaceholderSettings.Type is PlaceholderType.CenteredTitle) {
                shape.TextArea = new TextArea("Daily Testing Status Report");
            }
            if (shape.PlaceholderSettings.Type is PlaceholderType.Subtitle) {
                shape.TextArea = new TextArea($"{DateTime.Now: dddd, MMMM d, yyyy}");
            }
        }
        presentation.Slides.Add(slide1);

        Slide slide2 = new Slide(slideMaster.Layouts.GetOrCreate(SlideLayoutType.Object));
        foreach (Shape shape in slide2.Shapes) {
            if (shape.PlaceholderSettings.Type is PlaceholderType.Title) {
                shape.TextArea = new TextArea("Today’s Highlights");
            }
            if (shape.PlaceholderSettings.Type is PlaceholderType.Object) {
                TextArea textArea = new TextArea();
                textArea.Paragraphs.Clear();
                textArea.Paragraphs.Add(new TextParagraph("5 successful builds"));
                textArea.Paragraphs.Add(new TextParagraph("2 failed builds"));
                textArea.Paragraphs.Add(new TextParagraph("12 new bugs reported"));
                textArea.Paragraphs.Add(new TextParagraph("3 deployments"));
                textArea.Paragraphs.Add(new TextParagraph("1 rollback"));
                shape.TextArea = textArea;
            }
        }
        presentation.Slides.Add(slide2);

        Slide slide3 = new Slide(slideMaster.Layouts.GetOrCreate(SlideLayoutType.Object));
        foreach (Shape shape in slide3.Shapes) {
            if (shape.PlaceholderSettings.Type is PlaceholderType.Title) {
                shape.TextArea = new TextArea("Build Status");
            }
            if (shape.PlaceholderSettings.Type is PlaceholderType.Object) {
                shape.TextArea = new TextArea(" ");
                string imagePath = @"..\..\..\data\table.png";
                Stream stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
                PictureFill fill = new PictureFill(DXImage.FromStream(stream));
                fill.Stretch = true;
                shape.Fill = fill;
            }
        }
        presentation.Slides.Add(slide3);

        presentation.HeaderFooterManager.AddSlideNumberPlaceholder(presentation.Slides);
        presentation.HeaderFooterManager.AddFooterPlaceholder(presentation.Slides, "ProductXCompany");

        FileStream outputStream = new FileStream(@"..\..\..\data\my-presentation.pptx", FileMode.Create);
        presentation.SaveDocument(outputStream);
        outputStream.Dispose();

        presentation.ExportToPdf(new FileStream(@"..\..\..\data\exported-document.pdf", FileMode.Create));
    }
}
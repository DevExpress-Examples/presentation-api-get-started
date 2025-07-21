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
                TextArea textArea = new TextArea();
                textArea.Text = "Daily Testing Status Report";
                shape.TextArea = textArea;
            }
            if (shape.PlaceholderSettings.Type is PlaceholderType.Subtitle) {
                TextArea textArea = new TextArea();
                textArea.Text = $"{DateTime.Now: dddd, MMMM d, yyyy}";
                shape.TextArea = textArea;
            }
        }
        presentation.Slides.Add(slide1);

        Slide slide2 = new Slide(slideMaster.Layouts.GetOrCreate(SlideLayoutType.Object));
        foreach (Shape shape in slide2.Shapes) {
            if (shape.PlaceholderSettings.Type is PlaceholderType.Title) {
                TextArea textArea = new TextArea();
                textArea.Text = "Today’s Highlights";
                shape.TextArea = textArea;
            }
            if (shape.PlaceholderSettings.Type is PlaceholderType.Object) {
                TextArea textArea = new TextArea();
                TextParagraph paragraph1 = new TextParagraph();
                paragraph1.Runs.Add(new TextRun { Text = "5 successful builds" });
                textArea.Paragraphs.Add(paragraph1);

                TextParagraph paragraph2 = new TextParagraph();
                paragraph2.Runs.Add(new TextRun { Text = "2 failed builds" });
                textArea.Paragraphs.Add(paragraph2);

                TextParagraph paragraph3 = new TextParagraph();
                paragraph3.Runs.Add(new TextRun { Text = "12 new bugs reported" });
                textArea.Paragraphs.Add(paragraph3);

                TextParagraph paragraph4 = new TextParagraph();
                paragraph4.Runs.Add(new TextRun { Text = "3 deployments" });
                textArea.Paragraphs.Add(paragraph4);

                TextParagraph paragraph5 = new TextParagraph();
                paragraph5.Runs.Add(new TextRun { Text = "1 rollback" });
                textArea.Paragraphs.Add(paragraph5);
                shape.TextArea = textArea;
            }
        }
        presentation.Slides.Add(slide2);

        Slide slide3 = new Slide(slideMaster.Layouts.GetOrCreate(SlideLayoutType.Object));
        foreach (Shape shape in slide3.Shapes) {
            if (shape.PlaceholderSettings.Type is PlaceholderType.Title) {
                TextArea textArea = new TextArea();
                textArea.Text = "Build Status";
                shape.TextArea = textArea;
            }
            if (shape.PlaceholderSettings.Type is PlaceholderType.Object) {
                TextArea textArea = new TextArea();
                textArea.Text = " ";
                shape.TextArea = textArea;
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
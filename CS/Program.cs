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
            if (shape.PlaceholderSettings.Type is PlaceholderType.Body) {
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
        foreach (Shape shape in slide3.Shapes.ToList()) {
            if (shape.PlaceholderSettings.Type is PlaceholderType.Title) {
                shape.TextArea = new TextArea("Build Status");
            }
            if (shape.PlaceholderSettings.Type is PlaceholderType.Body) {

                RectangleF rect = presentation.GetActualShapeBounds(slide3, shape);
                slide3.Shapes.Remove(shape);

                Table table = new Table(5, 5, rect.X, rect.Y, rect.Width, rect.Height);
                slide3.Shapes.Add(table);

                table[0, 0].TextArea.Text = "Build ID";
                table[0, 1].TextArea.Text = "Branch";
                table[0, 2].TextArea.Text = "Status";
                table[0, 3].TextArea.Text = "Duration";
                table[0, 4].TextArea.Text = "Triggered By";

                table[1, 0].TextArea.Text = "#5421";
                table[1, 1].TextArea.Text = "main";
                table[1, 2].TextArea.Text = "✅  Passed";
                table[1, 3].TextArea.Text = "4m 30s";
                table[1, 4].TextArea.Text = "Auto-schedule";

                table[2, 0].TextArea.Text = "#5420";
                table[2, 1].TextArea.Text = "ui-fix";
                table[2, 2].TextArea.Text = "❌  Failed";
                table[2, 3].TextArea.Text = "2m 18s";
                table[2, 4].TextArea.Text = "Push by dev1";

                table[3, 0].TextArea.Text = "#5419";
                table[3, 1].TextArea.Text = "main";
                table[3, 2].TextArea.Text = "✅  Passed";
                table[3, 3].TextArea.Text = "3m 52s";
                table[3, 4].TextArea.Text = "Auto-schedule";

                table[4, 0].TextArea.Text = "#5418";
                table[4, 1].TextArea.Text = "hotfix";
                table[4, 2].TextArea.Text = "❌  Failed";
                table[4, 3].TextArea.Text = "5m 1s";
                table[4, 4].TextArea.Text = "Manual";

                table.Style = new ThemedTableStyle(TableStyleType.LightStyle1);
                table.HasBandedRows = false;
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
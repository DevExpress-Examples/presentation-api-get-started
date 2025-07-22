<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/1023545041/25.1.4%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T1300596)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->

# DevExpress Presentation API Library – Get Started 

This example creates a new presentation, adds three slides, and populates slides with content.

### Slide #1

![slide 1](media/gs-resulting-slide1.png)

### Slide #2

![slide 2](media/gs-resulting-slide2.png)

### Slide #3
![slide 3](media/gs-resulting-slide3.png)

> [!Important]  
> You need a license for the [DevExpress Office File API Subscription](https://www.devexpress.com/products/net/office-file-api/) or [DevExpress Universal Subscription](https://www.devexpress.com/subscriptions/universal.xml) to use our Office File API library in production code. For pricing information, refer to the following webpage: [DevExpress Subscriptions](https://www.devexpress.com/Subscriptions/). 

## Implementation Details

* Use the `Presentation()` parameterless constructor to create a new presentation. 

    ```cs
    Presentation presentation = new Presentation();
    ```
* A newly created presentation contains a single default slide master. The Slide Master is a top-level template slide that you can use as a base for other slides. Slide masters store content, layouts and settings that you can share among derived slides. The default slide master goes first in the `Presentation.SlideMasters` collection.

    ```cs
    SlideMaster slideMaster = presentation.SlideMasters[0];
    ```

* To create a slide, initialize a `Slide` object and add it to the `Presentation.Slides` collection. For each slide, specify the layout type as a constructor parameter. In this example, slides use predefined layouts stored in the slide master. To obtain a layout from `SlideMaster.Layouts`, call the `Get` or `GetOrCreate` method.

    ```cs
    Slide slide1 = new Slide(slideMaster.Layouts.Get(SlideLayoutType.Title));
    //...
    presentation.Slides.Add(slide1);
    ```

* Layouts add placeholder shapes to slides. Use the `Slide.Shapes` collection to access placeholder shapes and populate them with content.

    ```cs
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
    ```
* You can save the resulting presentation to a PPTX file and export it to a PDF file:

    ```cs
    FileStream outputStream = new FileStream(@"..\..\..\data\my-presentation.pptx", FileMode.Create);
    presentation.SaveDocument(outputStream);
    outputStream.Dispose();

    presentation.ExportToPdf(new FileStream(@"..\..\..\data\exported-document.pdf", FileMode.Create));
    ```

## Files to Review

- [Program.cs](./CS/Program.cs) ([Program.vb](./VB/Program.vb))

## Documentation

- [Create Your First Presentation with DevExpress Presentation API Library](https://docs.devexpress.com/OfficeFileAPI/405404/presentation-api/create-first-presentation)

## More Examples

- [Extract Presentation Images, Notes, and Pictures](https://github.com/DevExpress-Examples/presentation-document-api-extract-content)
<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=presentation-api-get-started&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=presentation-api-get-started&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->

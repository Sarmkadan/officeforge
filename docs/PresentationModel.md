# PresentationModel

The `PresentationModel` class serves as the core container for presentation documents within the `officeforge` framework. It manages the structural hierarchy of a presentation, including its individual slides, embedded shapes, textual lines, and metadata. This class provides a centralized interface for constructing, querying, and serializing presentation data, enabling programmatic manipulation of presentation documents for various export and processing tasks.

## API

### Slides
`public List<SlideModel> Slides`
A list containing all `SlideModel` instances associated with the presentation. This member allows for direct access, reordering, or removal of slides within the document.

### AddSlide
`public SlideModel AddSlide`
A member that facilitates the addition of a new slide to the presentation. Invoking or accessing this member appends a new `SlideModel` to the `Slides` collection and returns the newly created instance.

### Title
`public string? Title`
The optional title of the presentation. If not explicitly set, this member returns `null`.

### Shapes
`public List<ShapeTextModel> Shapes`
A list of all `ShapeTextModel` instances found within the presentation, containing textual information associated with shapes.

### ToPlainText
`public string ToPlainText`
A member that generates and returns a string representation of the presentation's content formatted as plain text.

### Name
`public string? Name`
The optional name or identifier for the presentation file or entity. If no name is defined, this member returns `null`.

### Lines
`public List<string> Lines`
A list of strings representing the raw textual lines extracted from the presentation content.

## Usage

### Creating and populating a presentation
```csharp
var presentation = new PresentationModel();
presentation.Title = "Annual Report 2026";
presentation.Name = "Report_Q2";

var slide = presentation.AddSlide;
// Populate the slide (assuming SlideModel has appropriate properties)
// slide.Title = "Overview";
```

### Exporting presentation content to text
```csharp
var presentation = LoadPresentation("deck.pptx"); // Assuming a loading mechanism
string textContent = presentation.ToPlainText;

Console.WriteLine($"Presentation: {presentation.Title}");
Console.WriteLine(textContent);
```

## Notes

*   **Thread-Safety**: This class is not inherently thread-safe. Concurrent access to the `Slides`, `Shapes`, or `Lines` lists from multiple threads must be synchronized externally to avoid race conditions or collection modification exceptions.
*   **Nullability**: The `Title` and `Name` members are nullable (`string?`). Consumers should implement appropriate null-checking logic before accessing these members to prevent `NullReferenceException` in dependent operations.
*   **List Initialization**: While `Slides`, `Shapes`, and `Lines` are public lists, ensure they are properly initialized before attempting to add or remove elements to avoid `NullReferenceException`.

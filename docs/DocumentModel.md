# DocumentModel

The `DocumentModel` class acts as the central container for document content within the `officeforge` framework, providing the structural representation required to compose, read, and export documents. It manages a hierarchical collection of paragraphs and their associated text runs, allowing for the programmatic definition of document structure, formatting, and content extraction.

## API

### Properties and Methods

*   **`Paragraphs`** (`List<ParagraphModel>`)
    Gets the collection of `ParagraphModel` instances that constitute the main body of the document.

*   **`AddParagraph`** (`ParagraphModel`)
    Provides an interface or accessor for appending or obtaining a new paragraph within the document structure.

*   **`ToPlainText`** (`string`)
    Returns the entire content of the `DocumentModel` as a plain text string, stripped of formatting and structural metadata.

*   **`Runs`** (`List<RunModel>`)
    Gets the list of `RunModel` instances directly associated with the document structure.

*   **`Kind`** (`ParagraphKind`)
    Gets or sets the classification of the paragraph, defining its structural role (e.g., normal text, heading, list item).

*   **`HeadingLevel`** (`int`)
    Gets or sets the heading level for the document structure; typically relevant when `Kind` indicates a heading type.

### Types

*   **`RunModel`** (`sealed record`)
    A record representing a discrete segment of text within the document, sharing identical character-level styling.

*   **`RunStyle`** (`sealed record`)
    Defines the visual or character-level formatting attributes (e.g., font, size, weight) applicable to a `RunModel`.

*   **`Default`** (`static RunStyle`)
    The default `RunStyle` configuration utilized when no explicit styling is applied to a `RunModel`.

## Usage

### Example 1: Creating a Document and Adding Content

```csharp
using OfficeForge;

var doc = new DocumentModel();

// Adding a paragraph to the document
var paragraph = doc.AddParagraph;
paragraph.Kind = ParagraphKind.Heading1;
paragraph.HeadingLevel = 1;

// Creating and adding a run with default styling
var run = new RunModel { Text = "Title of the Document", Style = RunStyle.Default };
paragraph.Runs.Add(run);
```

### Example 2: Extracting Plain Text

```csharp
using OfficeForge;

// Assuming 'doc' is an already populated DocumentModel
string fullText = doc.ToPlainText;

Console.WriteLine("Document Content:");
Console.WriteLine(fullText);
```

## Notes

### Thread Safety
`DocumentModel` and its associated components (`ParagraphModel`, `RunModel`) are not thread-safe. Concurrent read or write access to the same `DocumentModel` instance from multiple threads must be synchronized externally.

### Edge Cases
- **Empty Documents:** A `DocumentModel` may be initialized with an empty `Paragraphs` list. `ToPlainText` will return an empty string in this state.
- **Null Values:** While the collections (`Paragraphs`, `Runs`) are pre-initialized, ensure that any `RunModel` or `ParagraphModel` added to these lists is properly instantiated to avoid potential `NullReferenceException` during operations like `ToPlainText` or serialization.
- **HeadingLevel:** The `HeadingLevel` property is intended for use when `Kind` is set appropriately. Setting `HeadingLevel` without a corresponding `ParagraphKind` that supports headings may not produce the intended results in exported formats.

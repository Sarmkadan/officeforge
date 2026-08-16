namespace OfficeForge.Tests;

internal static class TemplateFillerTestsConstants
{
    public const string PlaceholderKnown = "Dear {{name}}, welcome to {{company}}!";
    public const string PlaceholderWhitespace = "{{ name }}";
    public const string PlaceholderUnknown = "Hi {{missing}}";

    public const string RunPart1 = "Total: {{tot";
    public const string RunPart2 = "al}} EUR";
    public const string ExpectedRunText = "Total: 99 EUR";

    public const string StaticText = "static text";

    public const string CustomerPlaceholder = "{{customer}}";
    public const string CustomerFormula = "A1&\"{{customer}}\"";

    public const string InputKeyPair = "key=a=b";
    public const string InputOtherPair = "other=x";
    public const string ExpectedKeyValue = "a=b";
    public const string ExpectedOtherValue = "x";

    public const string MalformedNoValue = "novalue";
    public const string MalformedLeading = "=leading";

    public const string MarkdownHeader = "## Sales";
    public const string MarkdownHeaderRow = "| Region | Total |";
    public const string MarkdownSeparator = "| --- | --- |";
    public const string MarkdownDataRow = "| North | 150 |";

    public const string JsonSalesKey = "\"Sales\"";
    public const string JsonA2 = "\"A2\": \"North\"";
    public const string JsonB2 = "\"B2\": \"150\"";

    public const string PlainTextHeader = "Region\tTotal";
    public const string PlainTextData = "North\t150";

    public const string DocumentTitle = "Title";
    public const string DocumentListItem = "point";
    public const string DocumentStrong = "strong";

    public const string MarkdownTitle = "## Title";
    public const string MarkdownListItem = "- *point*";
    public const string MarkdownStrong = "**strong**";

    public const string HelloText = "hello";
    public const string JsonKindBody = "\"kind\": \"Body\"";
    public const string JsonTextHello = "\"text\": \"hello\"";

    public const int HeadingLevel = 2;
    public const int Number150 = 150;
    public const int Number99 = 99;
}

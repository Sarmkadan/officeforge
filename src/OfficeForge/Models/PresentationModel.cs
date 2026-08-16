namespace OfficeForge.Models;

public sealed class PresentationModel : IPresentationModel, IEquatable<PresentationModel>
{
    public bool Equals(PresentationModel? other)
    {
        return other != null &&
               EqualityComparer<List<SlideModel>>.Default.Equals(Slides, other.Slides);
    }

    public override bool Equals(object? obj) => ReferenceEquals(this, obj) || obj is PresentationModel other && Equals(other);

    public override int GetHashCode()
    {
        return HashCode.Combine(Slides);
    }

    public static bool operator ==(PresentationModel? left, PresentationModel? right) => EqualityComparer<PresentationModel>.Default.Equals(left, right);

    public static bool operator !=(PresentationModel? left, PresentationModel? right) => !(left == right);
    public List<SlideModel> Slides { get; } = [];

    public SlideModel AddSlide(string? title = null)
    {
        var slide = new SlideModel { Title = title };
        Slides.Add(slide);
        return slide;
    }
}

public sealed class SlideModel
{
    public string? Title { get; set; }
    public List<ShapeTextModel> Shapes { get; } = [];

    public string ToPlainText()
    {
        var lines = new List<string>();
        if (!string.IsNullOrEmpty(Title)) lines.Add(Title);
        lines.AddRange(Shapes.SelectMany(s => s.Lines));
        return string.Join(Environment.NewLine, lines);
    }
}

public sealed class ShapeTextModel
{
    public string? Name { get; set; }
    public List<string> Lines { get; } = [];
}

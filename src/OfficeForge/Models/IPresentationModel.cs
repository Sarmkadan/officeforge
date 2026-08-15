namespace OfficeForge.Models;

public interface IPresentationModel
{
    List<SlideModel> Slides { get; }
    SlideModel AddSlide(string? title = null);
}

namespace OfficeForge.Models;

public interface IPresentationModel
{
    List<SlideModel> Slides { get; set; }
    SlideModel AddSlide(string? title = null);
}

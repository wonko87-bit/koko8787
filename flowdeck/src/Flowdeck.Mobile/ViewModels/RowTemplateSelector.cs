namespace Flowdeck.Mobile.ViewModels;

/// <summary>
/// Picks a row layout by what the row is. The two lists share one screen, and a day group
/// holds whichever kind that screen is currently showing.
/// </summary>
public sealed class RowTemplateSelector : DataTemplateSelector
{
    public DataTemplate? TodoTemplate { get; set; }

    public DataTemplate? EventTemplate { get; set; }

    protected override DataTemplate? OnSelectTemplate(object item, BindableObject container) => item switch
    {
        TodoRow => TodoTemplate,
        EventRow => EventTemplate,
        _ => null,
    };
}

using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Models.Text
{
    [System.Flags]
    public enum CharStyle
    {
        None = 0,
        Bold = IVisio.Enums.VisCellVals.visBold,
        Italic = IVisio.Enums.VisCellVals.visItalic,
        UnderLine = IVisio.Enums.VisCellVals.visUnderLine,
        SmallCaps = IVisio.Enums.VisCellVals.visSmallCaps
    }
}
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Pages
{
    public enum PrintPageOrientation
    {
        SameAsPrinter = IVisio.Enums.VisCellVals.visPPOSameAsPrinter,
        Portrait = IVisio.Enums.VisCellVals.visPPOPortrait,
        Landscape = IVisio.Enums.VisCellVals.visPPOLandscape
    }
}
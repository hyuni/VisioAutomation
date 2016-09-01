using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Text
{
    public enum TabStopAlignment
    {
        Left = IVisio.Enums.VisCellVals.visTabStopLeft,
        Center = IVisio.Enums.VisCellVals.visTabStopCenter,
        Right = IVisio.Enums.VisCellVals.visTabStopRight,
        Decimal = IVisio.Enums.VisCellVals.visTabStopDecimal,
        Comma = IVisio.Enums.VisCellVals.visTabStopComma
    }
}
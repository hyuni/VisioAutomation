using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Shapes.ConnectionPoints
{
    public enum ConnectionPointType
    {
        Inward = IVisio.Enums.VisCellVals.visCnnctTypeInward,
        Outward = IVisio.Enums.VisCellVals.visCnnctTypeOutward,
        InwardOutward = IVisio.Enums.VisCellVals.visCnnctTypeInwardOutward
    }
}
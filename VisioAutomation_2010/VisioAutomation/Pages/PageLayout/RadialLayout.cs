using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Pages.PageLayout
{
    public class RadialLayout : Layout
    {
        public RadialLayout()
        {
            this.LayoutStyle = LayoutStyle.Radial;
            this.ConnectorStyle = ConnectorStyle.RightAngle;
        }

        protected override void SetPageCells(PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) IVisio.Enums.VisCellVals.visPLOPlaceDefault;
        }
    }
}
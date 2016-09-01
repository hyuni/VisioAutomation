using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Pages.PageLayout
{
    public class CircularLayout : Layout
    {
        public CircularLayout()
        {
            this.LayoutStyle = LayoutStyle.Circular;
            this.ConnectorStyle = ConnectorStyle.CenterToCenter;
        }

        protected override void SetPageCells(PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) IVisio.Enums.VisCellVals.visPLOPlaceCircular;
        }
    }
}
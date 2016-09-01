using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Pages.PageLayout
{
    public class FlowchartLayout : Layout
    {
        public Direction Direction { get; set; }

        public FlowchartLayout()
        {
            this.LayoutStyle = LayoutStyle.Flowchart;
            this.ConnectorStyle = ConnectorStyle.Flowchart;
            this.Direction = Direction.TopToBottom;
        }

        protected override void SetPageCells(PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) FlowchartLayout.GetPlaceStyle(this.Direction);
        }

        private static IVisio.Enums.VisCellVals GetPlaceStyle(Direction dir)
        {
            if (dir == Direction.TopToBottom)
            {
                return IVisio.Enums.VisCellVals.visPLOPlaceTopToBottom;
            }
            else if (dir == Direction.LeftToRight)
            {
                return IVisio.Enums.VisCellVals.visPLOPlaceLeftToRight;
            }
            else if (dir == Direction.BottomToTop)
            {
                return IVisio.Enums.VisCellVals.visPLOPlaceBottomToTop;
            }
            else if (dir == Direction.RightToLeft)
            {
                return IVisio.Enums.VisCellVals.visPLOPlaceRightToLeft;
            }
            else
            {
                throw new System.ArgumentException(nameof(dir));
            }
        }

        protected override IVisio.Enums.VisCellVals? ConnectorsStyleToRouteStyle()
        {
            var rs = base.ConnectorsStyleToRouteStyle();
            if (rs.HasValue)
            {
                return rs;
            }
            return this.ConnectorsStyleAndDirectionToRouteStyle(this.ConnectorStyle, this.Direction);
        }
    }
}
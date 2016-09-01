using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Pages.PageLayout
{
    public class CompactTreeLayout : Layout
    {
        public CompactTreeDirection Direction { get; set; }

        public CompactTreeLayout()
        {
            this.LayoutStyle = LayoutStyle.CompactTree;
            this.ConnectorStyle = ConnectorStyle.OrganizationChart;
            this.Direction = CompactTreeDirection.DownThenRight;
        }

        protected override void SetPageCells(PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) CompactTreeLayout.GetPlaceStyle(this.Direction);
        }

        private static IVisio.Enums.VisCellVals GetPlaceStyle(CompactTreeDirection dir)
        {
            if (dir == CompactTreeDirection.DownThenRight)
            {
                return IVisio.Enums.VisCellVals.visPLOPlaceCompactDownRight;
            }
            else if (dir == CompactTreeDirection.RightThenDown)
            {
                return IVisio.Enums.VisCellVals.visPLOPlaceCompactRightDown;
            }
            else if (dir == CompactTreeDirection.RightThenUp)
            {
                return IVisio.Enums.VisCellVals.visPLOPlaceCompactRightUp;
            }
            else if (dir == CompactTreeDirection.UpThenRigtht)
            {
                return IVisio.Enums.VisCellVals.visPLOPlaceCompactUpRight;
            }
            else if (dir == CompactTreeDirection.UpThenLeft)
            {
                return IVisio.Enums.VisCellVals.visPLOPlaceCompactUpLeft;
            }
            else if (dir == CompactTreeDirection.LeftThenUp)
            {
                return IVisio.Enums.VisCellVals.visPLOPlaceCompactLeftUp;
            }
            else if (dir == CompactTreeDirection.LeftThenDown)
            {
                return IVisio.Enums.VisCellVals.visPLOPlaceCompactLeftDown;
            }
            else if (dir == CompactTreeDirection.DownThenLeft)
            {
                return IVisio.Enums.VisCellVals.visPLOPlaceCompactDownLeft;
            }
            else
            {
                throw new System.ArgumentException(nameof(dir));
            }
        }
    }
}
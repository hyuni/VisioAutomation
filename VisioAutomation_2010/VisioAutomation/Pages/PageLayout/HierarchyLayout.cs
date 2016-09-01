using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Pages.PageLayout
{
    public class HierarchyLayout : Layout
    {
        public Direction Direction { get; set; }
        public HorizontalAlignment HorizontalAlignment { get; set; }
        public VerticalAlignment VerticalAlignment { get; set; }

        public HierarchyLayout()
        {
            this.LayoutStyle = LayoutStyle.Hierarchy;
            this.ConnectorStyle = ConnectorStyle.OrganizationChart;
            this.HorizontalAlignment = HorizontalAlignment.Center;
            this.VerticalAlignment = VerticalAlignment.Middle;
        }

        protected override void SetPageCells(PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) HierarchyLayout.GetPlaceStyle(this.Direction, this.HorizontalAlignment, this.VerticalAlignment);
        }

        private static IVisio.Enums.VisCellVals GetPlaceStyle(Direction dir, HorizontalAlignment halign, VerticalAlignment valign)
        {
            if (dir == Direction.BottomToTop)
            {
                if (halign == HorizontalAlignment.Left)
                {
                    return IVisio.Enums.VisCellVals.visPLOPlaceHierarchyBottomToTopLeft;
                }
                else if (halign == HorizontalAlignment.Center)
                {
                    return IVisio.Enums.VisCellVals.visPLOPlaceHierarchyBottomToTopCenter;
                }
                else if (halign == HorizontalAlignment.Right)
                {
                    return IVisio.Enums.VisCellVals.visPLOPlaceHierarchyBottomToTopRight;
                }
            }
            else if (dir == Direction.TopToBottom)
            {
                if (halign == HorizontalAlignment.Left)
                {
                    return IVisio.Enums.VisCellVals.visPLOPlaceHierarchyTopToBottomLeft;
                }
                else if (halign == HorizontalAlignment.Center)
                {
                    return IVisio.Enums.VisCellVals.visPLOPlaceHierarchyTopToBottomCenter;
                }
                else if (halign == HorizontalAlignment.Right)
                {
                    return IVisio.Enums.VisCellVals.visPLOPlaceHierarchyTopToBottomRight;
                }
            }
            else if (dir == Direction.LeftToRight)
            {
                if (valign == VerticalAlignment.Top)
                {
                    return IVisio.Enums.VisCellVals.visPLOPlaceHierarchyLeftToRightTop;
                }
                else if (valign == VerticalAlignment.Middle)
                {
                    return IVisio.Enums.VisCellVals.visPLOPlaceHierarchyLeftToRightMiddle;
                }
                else if (valign == VerticalAlignment.Bottom)
                {
                    return IVisio.Enums.VisCellVals.visPLOPlaceHierarchyLeftToRightBottom;
                }
            }
            else if (dir == Direction.RightToLeft)
            {
                if (valign == VerticalAlignment.Top)
                {
                    return IVisio.Enums.VisCellVals.visPLOPlaceHierarchyRightToLeftTop;
                }
                else if (valign == VerticalAlignment.Middle)
                {
                    return IVisio.Enums.VisCellVals.visPLOPlaceHierarchyRightToLeftMiddle;
                }
                else if (valign == VerticalAlignment.Bottom)
                {
                    return IVisio.Enums.VisCellVals.visPLOPlaceHierarchyRightToLeftBottom;
                }
                else
                {
                    throw new System.ArgumentOutOfRangeException(nameof(dir));
                }
            }
            throw new System.ArgumentOutOfRangeException(nameof(dir));
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
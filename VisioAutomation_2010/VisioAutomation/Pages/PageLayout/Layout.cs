using VisioAutomation.ShapeSheet.Writers;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Pages.PageLayout
{
    public abstract class Layout
    {
        public LayoutStyle LayoutStyle { get; set; }
        public ConnectorStyle ConnectorStyle { get; set; }
        public ConnectorAppearance ConnectorAppearance { get; set; }
        public Drawing.Size AvenueSize { get; set; }

        protected Layout()
        {
            this.AvenueSize = new Drawing.Size(0.375, 0.375);
        }

        protected virtual void SetPageCells(PageCells pagecells)
        {
            pagecells.AvenueSizeX = this.AvenueSize.Width;
            pagecells.AvenueSizeY = this.AvenueSize.Height;
            pagecells.LineRouteExt = (int) Layout.ConnectorAppearanceToLineRouteExt(this.ConnectorAppearance);

            var rs = this.ConnectorsStyleToRouteStyle();
            if (rs.HasValue)
            {
                pagecells.RouteStyle = (int) rs.Value;
            }
        }

        private static IVisio.Enums.VisCellVals ConnectorAppearanceToLineRouteExt(ConnectorAppearance ca)
        {
            if (ca == ConnectorAppearance.Default)
            {
                return IVisio.Enums.VisCellVals.visLORouteExtDefault;
            }
            else if (ca == ConnectorAppearance.Straight)
            {
                return IVisio.Enums.VisCellVals.visLORouteExtStraight;
            }
            else if (ca == ConnectorAppearance.Curved)
            {
                return IVisio.Enums.VisCellVals.visLORouteExtNURBS;
            }
            else
            {
                throw new System.ArgumentOutOfRangeException(nameof(ca));
            }
        }

        protected virtual IVisio.Enums.VisCellVals? ConnectorsStyleToRouteStyle()
        {
            var cs = this.ConnectorStyle;
            if (cs == ConnectorStyle.RightAngle)
            {
                return IVisio.Enums.VisCellVals.visLORouteRightAngle;
            }
            else if (cs == ConnectorStyle.Straight)
            {
                return IVisio.Enums.VisCellVals.visLORouteStraight;
            }
            else if (cs == ConnectorStyle.CenterToCenter)
            {
                return IVisio.Enums.VisCellVals.visLORouteCenterToCenter;
            }
            else if (cs == ConnectorStyle.Network)
            {
                return IVisio.Enums.VisCellVals.visLORouteNetwork;
            }
            else
            {
                return null;
            }
        }

        protected IVisio.Enums.VisCellVals ConnectorsStyleAndDirectionToRouteStyle(ConnectorStyle cs, Direction dir)
        {
            if (cs == ConnectorStyle.Flowchart)
            {
                if (dir == Direction.BottomToTop)
                {
                    return IVisio.Enums.VisCellVals.visLORouteFlowchartSN;
                }
                else if (dir == Direction.TopToBottom)
                {
                    return IVisio.Enums.VisCellVals.visLORouteFlowchartNS;
                }
                else if (dir == Direction.LeftToRight)
                {
                    return IVisio.Enums.VisCellVals.visLORouteFlowchartWE;
                }
                else if (dir == Direction.RightToLeft)
                {
                    return IVisio.Enums.VisCellVals.visLORouteFlowchartEW;
                }
            }
            else if (cs == ConnectorStyle.OrganizationChart)
            {
                if (dir == Direction.BottomToTop)
                {
                    return IVisio.Enums.VisCellVals.visLORouteOrgChartSN;
                }
                else if (dir == Direction.TopToBottom)
                {
                    return IVisio.Enums.VisCellVals.visLORouteOrgChartNS;
                }
                else if (dir == Direction.LeftToRight)
                {
                    return IVisio.Enums.VisCellVals.visLORouteOrgChartWE;
                }
                else if (dir == Direction.RightToLeft)
                {
                    return IVisio.Enums.VisCellVals.visLORouteOrgChartEW;
                }
            }
            else if (cs == ConnectorStyle.Simple)
            {
                if (dir == Direction.BottomToTop)
                {
                    return IVisio.Enums.VisCellVals.visLORouteSimpleSN;
                }
                else if (dir == Direction.TopToBottom)
                {
                    return IVisio.Enums.VisCellVals.visLORouteSimpleNS;
                }
                else if (dir == Direction.LeftToRight)
                {
                    return IVisio.Enums.VisCellVals.visLORouteSimpleWE;
                }
                else if (dir == Direction.RightToLeft)
                {
                    return IVisio.Enums.VisCellVals.visLORouteSimpleEW;
                }
            }
            throw new System.ArgumentOutOfRangeException(nameof(cs));
        }

        public void Apply(IVisio.IVPage page)
        {
            var pagecells = new PageCells();
            this.SetPageCells(pagecells);

            var writer = new FormulaWriterSRC();
            pagecells.SetFormulas(writer);
            var pagesheet = page.PageSheet;
            writer.Commit((IVisio.IVShape)pagesheet);
            page.Layout();
        }
    }
}
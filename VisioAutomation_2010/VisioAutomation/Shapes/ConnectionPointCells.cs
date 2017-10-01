using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class ConnectionPointCells : CellGroupMultiRow
    {
        public CellValueLiteral X { get; set; }
        public CellValueLiteral Y { get; set; }
        public CellValueLiteral DirX { get; set; }
        public CellValueLiteral DirY { get; set; }
        public CellValueLiteral Type { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.ConnectionPointX, this.X, nameof(SrcConstants.ConnectionPointX), nameof(this.X));
                yield return SrcValuePair.Create(SrcConstants.ConnectionPointY, this.Y, nameof(SrcConstants.ConnectionPointY), nameof(this.Y));
                yield return SrcValuePair.Create(SrcConstants.ConnectionPointDirX, this.DirX, nameof(SrcConstants.ConnectionPointDirX), nameof(this.DirX));
                yield return SrcValuePair.Create(SrcConstants.ConnectionPointDirY, this.DirY, nameof(SrcConstants.ConnectionPointDirY), nameof(this.DirY));
                yield return SrcValuePair.Create(SrcConstants.ConnectionPointType, this.Type, nameof(SrcConstants.ConnectionPointType), nameof(this.Type));
            }
        }

        public static List<List<ConnectionPointCells>> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCells(page, shapeids, type);
        }

        public static List<ConnectionPointCells> GetCells(IVisio.Shape shape, CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCells(shape, type);
        }

        private static readonly System.Lazy<ConnectionPointCellsReader> lazy_query = new System.Lazy<ConnectionPointCellsReader>();

        class ConnectionPointCellsReader : ReaderMultiRow<ConnectionPointCells>
        {
            public SectionQueryColumn DirX { get; set; }
            public SectionQueryColumn DirY { get; set; }
            public SectionQueryColumn Type { get; set; }
            public SectionQueryColumn X { get; set; }
            public SectionQueryColumn Y { get; set; }

            public ConnectionPointCellsReader()
            {
                var sec = this.query.SectionQueries.Add(IVisio.VisSectionIndices.visSectionConnectionPts);

                this.DirX = sec.Columns.Add(SrcConstants.ConnectionPointDirX, nameof(this.DirX));
                this.DirY = sec.Columns.Add(SrcConstants.ConnectionPointDirY, nameof(this.DirY));
                this.Type = sec.Columns.Add(SrcConstants.ConnectionPointType, nameof(this.Type));
                this.X = sec.Columns.Add(SrcConstants.ConnectionPointX, nameof(this.X));
                this.Y = sec.Columns.Add(SrcConstants.ConnectionPointY, nameof(this.Y));

            }

            public override ConnectionPointCells ToCellGroup(Utilities.ArraySegment<string> row)
            {
                var cells = new ConnectionPointCells();
                cells.X = row[this.X];
                cells.Y = row[this.Y];
                cells.DirX = row[this.DirX];
                cells.DirY = row[this.DirY];
                cells.Type = row[this.Type];

                return cells;
            }
        }

    }
}
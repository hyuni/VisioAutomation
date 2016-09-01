using VisioAutomation.ShapeSheet.Queries.Columns;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.ShapeSheet.Queries.CommonQueries
{
    class UserDefinedCellsQuery : Query
    {
        public ColumnSubQuery Value { get; set; }
        public ColumnSubQuery Prompt { get; set; }

        public UserDefinedCellsQuery()
        {
            var sec = this.AddSubQuery(IVisio.Enums.VisSectionIndices.visSectionUser);
            this.Value = sec.AddCell(SRCCON.User_Value, nameof(SRCCON.User_Value));
            this.Prompt = sec.AddCell(SRCCON.User_Prompt, nameof(SRCCON.User_Prompt));
        }

        public Shapes.UserDefinedCells.UserDefinedCell GetCells(ShapeSheet.CellData<string>[] row)
        {
            var cells = new Shapes.UserDefinedCells.UserDefinedCell();
            cells.Value = row[this.Value];
            cells.Prompt = row[this.Prompt];
            return cells;
        }
    }
}
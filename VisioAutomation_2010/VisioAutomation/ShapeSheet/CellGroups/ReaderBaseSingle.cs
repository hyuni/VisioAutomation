using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.ShapeSheet.CellGroups
{
    public abstract class ReaderBaseSingle<TCellGroup>
    {
        protected CellQuery query;

        protected ReaderBaseSingle()
        {
            this.query = new CellQuery();
        }

        protected abstract void validate_query();

        public abstract TCellGroup CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row);

    }

    public abstract class ReaderBaseMulti<TCellGroup>
    {
        protected SectionQuery query;

        protected ReaderBaseMulti()
        {
            this.query = new SectionQuery();
        }

        protected abstract void validate_query();

        public abstract TCellGroup CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row);

    }

}
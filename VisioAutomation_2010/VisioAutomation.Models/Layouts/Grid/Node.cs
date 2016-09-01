using IVisio = NetOffice.VisioApi;


namespace VisioAutomation.Models.Layouts.Grid
{
    public class Node
    {
        public IVisio.IVMaster Master { get; set; }
        public string Text { get; set; }
        public IVisio.IVShape Shape { get; set; }
        public Drawing.Rectangle Rectangle { get; set; }
        public short ShapeID { get; set; }
        public int Column { get; set; }
        public int Row { get; set; }
        public object Data { get; set; }
        public bool Draw { get; set; }

        public Dom.ShapeCells Cells { get; set; }

        public Node()
        {
            this.ShapeID = -1;
        }

    }
}
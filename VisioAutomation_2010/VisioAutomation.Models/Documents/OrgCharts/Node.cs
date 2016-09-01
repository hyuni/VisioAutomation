using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Models.Documents.OrgCharts
{
    public class Node
    {
        private readonly NodeList _children;
        internal Node _parent;

        public string Text { get; set; }
        public IVisio.Shape VisioShape { get; set; }
        public Dom.Node DOMNode { get; set; }
        public string URL { get; set; }
        public Drawing.Size? Size { get; set; }

        public Node()
        {
            this._children = new NodeList(this);
        }

        public Node(string name) :
            this ()
        {
            this.Text = name;
        }

        public NodeList Children
        {
            get { return this._children; }
        }

        public Node Parent
        {
            get { return this._parent; }
        }      
    }
}
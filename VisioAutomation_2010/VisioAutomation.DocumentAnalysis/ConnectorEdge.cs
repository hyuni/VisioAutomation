using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.DocumentAnalysis
{
    public struct ConnectorEdge
    {
        public IVisio.IVShape Connector { get; }
        public IVisio.IVShape From { get; }
        public IVisio.IVShape To { get; }

        public ConnectorEdge(IVisio.IVShape connectingshape, IVisio.IVShape fromshape, IVisio.IVShape toshape) : this()
        {
            if (fromshape == null)
            {
                throw new System.ArgumentNullException(nameof(fromshape));
            }

            if (toshape == null)
            {
                throw new System.ArgumentNullException(nameof(toshape));
            }

            this.Connector = connectingshape;
            this.From = fromshape;
            this.To = toshape;
        }

        public override string ToString()
        {
            string from_name = this.From !=null ? this.From.NameU : "null";
            string to_name = this.To != null ? this.To.NameU : "null";

            if (this.Connector != null)
            {
                var connector_name = this.Connector.NameU;
                return string.Format("({0}:{1}->{2})", connector_name, from_name, to_name);                
            }
            else
            {
                return string.Format("({0}->{1})", from_name, to_name);
            }
        }
    }
}
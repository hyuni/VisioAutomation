using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Models.Text
{
    public class FieldBase : Node
    {
        private const string placeholder_string = "[FIELD]";
        public IVisio.Enums.VisFieldFormats Format { get; set; }

        internal FieldBase(VisioAutomation.Models.Text.NodeType nt)
            : base(nt)
        {
        }
        
        public string PlaceholderText
        {
            get
            {
                return FieldBase.placeholder_string;
            }
        }
    }

}

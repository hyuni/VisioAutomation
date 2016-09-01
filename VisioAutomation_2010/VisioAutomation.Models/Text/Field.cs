using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Models.Text
{
    public class Field : FieldBase
    {
        public IVisio.Enums.VisFieldCategories Category { get; set; }
        public IVisio.Enums.VisFieldCodes Code { get; set; }

        public Field(IVisio.Enums.VisFieldCategories category, IVisio.Enums.VisFieldCodes code, IVisio.Enums.VisFieldFormats format) :
            base(NodeType.Field)
        {
            this.Category = category;
            this.Code = code;
            this.Format = format;
        }
    }
}

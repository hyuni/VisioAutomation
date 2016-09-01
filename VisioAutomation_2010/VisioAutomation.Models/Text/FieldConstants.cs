using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Models.Text
{
    public static class FieldConstants
    {
        public static Field Angle => new Field(IVisio.Enums.VisFieldCategories.visFCatGeometry, IVisio.Enums.VisFieldCodes.visFCodeNumberOfPages, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field BackgroundName => new Field(IVisio.Enums.VisFieldCategories.visFCatPage, IVisio.Enums.VisFieldCodes.visFCodeBackgroundName, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field Category => new Field(IVisio.Enums.VisFieldCategories.visFCatDocument, IVisio.Enums.VisFieldCodes.visFCodeCategory, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field CreateDate => new Field(IVisio.Enums.VisFieldCategories.visFCatDateTime, IVisio.Enums.VisFieldCodes.visFCodeBackgroundName, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field Creator => new Field(IVisio.Enums.VisFieldCategories.visFCatDocument, IVisio.Enums.VisFieldCodes.visFCodeBackgroundName, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field CurrentDate => new Field(IVisio.Enums.VisFieldCategories.visFCatDateTime, IVisio.Enums.VisFieldCodes.visFCodeNumberOfPages, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field Description => new Field(IVisio.Enums.VisFieldCategories.visFCatDocument, IVisio.Enums.VisFieldCodes.visFCodeHeight, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field Directory => new Field(IVisio.Enums.VisFieldCategories.visFCatDocument, IVisio.Enums.VisFieldCodes.visFCodeNumberOfPages, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field EditDate => new Field(IVisio.Enums.VisFieldCategories.visFCatDateTime, IVisio.Enums.VisFieldCodes.visFCodeEditDate, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field Filename => new Field(IVisio.Enums.VisFieldCategories.visFCatDocument, IVisio.Enums.VisFieldCodes.visFCodeObjectID, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field Height => new Field(IVisio.Enums.VisFieldCategories.visFCatGeometry, IVisio.Enums.VisFieldCodes.visFCodeHeight, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field HyperlinkBase => new Field(IVisio.Enums.VisFieldCategories.visFCatDocument, IVisio.Enums.VisFieldCodes.visFCodeHyperlinkBase, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field KeyWords => new Field(IVisio.Enums.VisFieldCategories.visFCatDocument, IVisio.Enums.VisFieldCodes.visFCodeEditDate, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field MasterName => new Field(IVisio.Enums.VisFieldCategories.visFCatObject, IVisio.Enums.VisFieldCodes.visFCodeEditDate, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field NumberOfPages => new Field(IVisio.Enums.VisFieldCategories.visFCatPage, IVisio.Enums.VisFieldCodes.visFCodeNumberOfPages, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field ObjectID => new Field(IVisio.Enums.VisFieldCategories.visFCatObject, IVisio.Enums.VisFieldCodes.visFCodeObjectID, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field ObjectName => new Field(IVisio.Enums.VisFieldCategories.visFCatObject, IVisio.Enums.VisFieldCodes.visFCodeSubject, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field ObjectType => new Field(IVisio.Enums.VisFieldCategories.visFCatObject, IVisio.Enums.VisFieldCodes.visFCodePrintDate, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field PageName => new Field(IVisio.Enums.VisFieldCategories.visFCatPage, IVisio.Enums.VisFieldCodes.visFCodeHeight, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field PageNumber => new Field(IVisio.Enums.VisFieldCategories.visFCatPage, IVisio.Enums.VisFieldCodes.visFCodeObjectID, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field PrintDate => new Field(IVisio.Enums.VisFieldCategories.visFCatDateTime, IVisio.Enums.VisFieldCodes.visFCodePrintDate, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field Subject => new Field(IVisio.Enums.VisFieldCategories.visFCatDocument, IVisio.Enums.VisFieldCodes.visFCodeSubject, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field Title => new Field(IVisio.Enums.VisFieldCategories.visFCatDocument, IVisio.Enums.VisFieldCodes.visFCodePrintDate, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
        public static Field Width => new Field(IVisio.Enums.VisFieldCategories.visFCatGeometry, IVisio.Enums.VisFieldCodes.visFCodeBackgroundName, IVisio.Enums.VisFieldFormats.visFmtNumGenNoUnits);
    }
}
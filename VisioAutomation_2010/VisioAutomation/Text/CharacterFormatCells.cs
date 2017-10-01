using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Text
{
    public class CharacterFormatCells : CellGroupMultiRow
    {
        public CellValueLiteral Color { get; set; }
        public CellValueLiteral Font { get; set; }
        public CellValueLiteral Size { get; set; }
        public CellValueLiteral Style { get; set; }
        public CellValueLiteral ColorTransparency { get; set; }
        public CellValueLiteral AsianFont { get; set; }
        public CellValueLiteral Case { get; set; }
        public CellValueLiteral ComplexScriptFont { get; set; }
        public CellValueLiteral ComplexScriptSize { get; set; }
        public CellValueLiteral DoubleStrikethrough { get; set; }
        public CellValueLiteral DoubleUnderline { get; set; }
        public CellValueLiteral LangID { get; set; }
        public CellValueLiteral Locale { get; set; }
        public CellValueLiteral LocalizeFont { get; set; }
        public CellValueLiteral Overline { get; set; }
        public CellValueLiteral Perpendicular { get; set; }
        public CellValueLiteral Pos { get; set; }
        public CellValueLiteral RTLText { get; set; }
        public CellValueLiteral FontScale { get; set; }
        public CellValueLiteral Letterspace { get; set; }
        public CellValueLiteral Strikethru { get; set; }
        public CellValueLiteral UseVertical { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.CharColor, this.Color, nameof(SrcConstants.CharColor), nameof(this.Color));
                yield return SrcValuePair.Create(SrcConstants.CharFont, this.Font, nameof(SrcConstants.CharFont), nameof(this.Font));
                yield return SrcValuePair.Create(SrcConstants.CharSize, this.Size, nameof(SrcConstants.CharSize), nameof(this.Size));
                yield return SrcValuePair.Create(SrcConstants.CharStyle, this.Style, nameof(SrcConstants.CharStyle), nameof(this.Style));
                yield return SrcValuePair.Create(SrcConstants.CharColorTransparency, this.ColorTransparency, nameof(SrcConstants.CharColorTransparency), nameof(this.ColorTransparency));
                yield return SrcValuePair.Create(SrcConstants.CharAsianFont, this.AsianFont, nameof(SrcConstants.CharAsianFont), nameof(this.AsianFont));
                yield return SrcValuePair.Create(SrcConstants.CharCase, this.Case, nameof(SrcConstants.CharCase), nameof(this.Case));
                yield return SrcValuePair.Create(SrcConstants.CharComplexScriptFont, this.ComplexScriptFont, nameof(SrcConstants.CharComplexScriptFont), nameof(this.ComplexScriptFont));
                yield return SrcValuePair.Create(SrcConstants.CharComplexScriptSize, this.ComplexScriptSize, nameof(SrcConstants.CharComplexScriptSize), nameof(this.ComplexScriptSize));
                yield return SrcValuePair.Create(SrcConstants.CharDoubleUnderline, this.DoubleUnderline, nameof(SrcConstants.CharDoubleUnderline), nameof(this.DoubleUnderline));
                yield return SrcValuePair.Create(SrcConstants.CharDoubleStrikethrough, this.DoubleStrikethrough, nameof(SrcConstants.CharDoubleStrikethrough), nameof(this.DoubleStrikethrough));
                yield return SrcValuePair.Create(SrcConstants.CharLangID, this.LangID, nameof(SrcConstants.CharLangID), nameof(this.LangID));
                yield return SrcValuePair.Create(SrcConstants.CharFontScale, this.FontScale, nameof(SrcConstants.CharFontScale), nameof(this.FontScale));
                yield return SrcValuePair.Create(SrcConstants.CharLangID, this.LangID, nameof(SrcConstants.CharLangID), nameof(this.LangID));
                yield return SrcValuePair.Create(SrcConstants.CharLetterspace, this.Letterspace, nameof(SrcConstants.CharLetterspace), nameof(this.Letterspace));
                yield return SrcValuePair.Create(SrcConstants.CharLocale, this.Locale, nameof(SrcConstants.CharLocale), nameof(this.Locale));
                yield return SrcValuePair.Create(SrcConstants.CharLocalizeFont, this.LocalizeFont, nameof(SrcConstants.CharLocalizeFont), nameof(this.LocalizeFont));
                yield return SrcValuePair.Create(SrcConstants.CharOverline, this.Overline, nameof(SrcConstants.CharOverline), nameof(this.Overline));
                yield return SrcValuePair.Create(SrcConstants.CharPerpendicular, this.Perpendicular, nameof(SrcConstants.CharPerpendicular), nameof(this.Perpendicular));
                yield return SrcValuePair.Create(SrcConstants.CharPos, this.Pos, nameof(SrcConstants.CharPos), nameof(this.Pos));
                yield return SrcValuePair.Create(SrcConstants.CharRTLText, this.RTLText, nameof(SrcConstants.CharRTLText), nameof(this.RTLText));
                yield return SrcValuePair.Create(SrcConstants.CharStrikethru, this.Strikethru, nameof(SrcConstants.CharStrikethru), nameof(this.Strikethru));
                yield return SrcValuePair.Create(SrcConstants.CharUseVertical, this.UseVertical, nameof(SrcConstants.CharUseVertical), nameof(this.UseVertical));
            }
        }

        public static List<List<CharacterFormatCells>> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCells(page, shapeids, type);
        }

        public static List<CharacterFormatCells> GetCells(IVisio.Shape shape, CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCells(shape, type);
        }

        private static readonly System.Lazy<CharacterFormatCellsReader> lazy_query = new System.Lazy<CharacterFormatCellsReader>();


        class CharacterFormatCellsReader : ReaderMultiRow<Text.CharacterFormatCells>
        {
            public SectionQueryColumn Font { get; set; }
            public SectionQueryColumn Style { get; set; }
            public SectionQueryColumn Color { get; set; }
            public SectionQueryColumn Size { get; set; }
            public SectionQueryColumn ColorTransparency { get; set; }
            public SectionQueryColumn AsianFont { get; set; }
            public SectionQueryColumn Case { get; set; }
            public SectionQueryColumn ComplexScriptFont { get; set; }
            public SectionQueryColumn ComplexScriptSize { get; set; }
            public SectionQueryColumn DoubleStrikethrough { get; set; }
            public SectionQueryColumn DoubleUnderline { get; set; }
            public SectionQueryColumn LangID { get; set; }
            public SectionQueryColumn Locale { get; set; }
            public SectionQueryColumn LocalizeFont { get; set; }
            public SectionQueryColumn Overline { get; set; }
            public SectionQueryColumn Perpendicular { get; set; }
            public SectionQueryColumn Pos { get; set; }
            public SectionQueryColumn RTLText { get; set; }
            public SectionQueryColumn FontScale { get; set; }
            public SectionQueryColumn Letterspace { get; set; }
            public SectionQueryColumn Strikethru { get; set; }
            public SectionQueryColumn UseVertical { get; set; }

            public CharacterFormatCellsReader()
            {
                var sec = this.query.SectionQueries.Add(IVisio.VisSectionIndices.visSectionCharacter);

                this.Color = sec.Columns.Add(SrcConstants.CharColor, nameof(this.Color));
                this.ColorTransparency = sec.Columns.Add(SrcConstants.CharColorTransparency, nameof(this.ColorTransparency));
                this.Font = sec.Columns.Add(SrcConstants.CharFont, nameof(this.Font));
                this.Size = sec.Columns.Add(SrcConstants.CharSize, nameof(this.Size));
                this.Style = sec.Columns.Add(SrcConstants.CharStyle, nameof(this.Style));
                this.AsianFont = sec.Columns.Add(SrcConstants.CharAsianFont, nameof(this.AsianFont));
                this.Case = sec.Columns.Add(SrcConstants.CharCase, nameof(this.Case));
                this.ComplexScriptFont = sec.Columns.Add(SrcConstants.CharComplexScriptFont, nameof(this.ComplexScriptFont));
                this.ComplexScriptSize = sec.Columns.Add(SrcConstants.CharComplexScriptSize, nameof(this.ComplexScriptSize));
                this.DoubleStrikethrough = sec.Columns.Add(SrcConstants.CharDoubleStrikethrough, nameof(this.DoubleStrikethrough));
                this.DoubleUnderline = sec.Columns.Add(SrcConstants.CharDoubleUnderline, nameof(this.DoubleUnderline));
                this.LangID = sec.Columns.Add(SrcConstants.CharLangID, nameof(this.LangID));
                this.Locale = sec.Columns.Add(SrcConstants.CharLocale, nameof(this.Locale));
                this.LocalizeFont = sec.Columns.Add(SrcConstants.CharLocalizeFont, nameof(this.LocalizeFont));
                this.Overline = sec.Columns.Add(SrcConstants.CharOverline, nameof(this.Overline));
                this.Perpendicular = sec.Columns.Add(SrcConstants.CharPerpendicular, nameof(this.Perpendicular));
                this.Pos = sec.Columns.Add(SrcConstants.CharPos, nameof(this.Pos));
                this.RTLText = sec.Columns.Add(SrcConstants.CharRTLText, nameof(this.RTLText));
                this.FontScale = sec.Columns.Add(SrcConstants.CharFontScale, nameof(this.FontScale));
                this.Letterspace = sec.Columns.Add(SrcConstants.CharLetterspace, nameof(this.Letterspace));
                this.Strikethru = sec.Columns.Add(SrcConstants.CharStrikethru, nameof(this.Strikethru));
                this.UseVertical = sec.Columns.Add(SrcConstants.CharUseVertical, nameof(this.UseVertical));

            }

            public override Text.CharacterFormatCells ToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
            {
                var cells = new Text.CharacterFormatCells();
                cells.Color = row[this.Color];
                cells.ColorTransparency = row[this.ColorTransparency];
                cells.Font = row[this.Font];
                cells.Size = row[this.Size];
                cells.Style = row[this.Style];
                cells.AsianFont = row[this.AsianFont];
                cells.AsianFont = row[this.AsianFont];
                cells.Case = row[this.Case];
                cells.ComplexScriptFont = row[this.ComplexScriptFont];
                cells.ComplexScriptSize = row[this.ComplexScriptSize];
                cells.DoubleStrikethrough = row[this.DoubleStrikethrough];
                cells.DoubleUnderline = row[this.DoubleUnderline];
                cells.FontScale = row[this.FontScale];
                cells.LangID = row[this.LangID];
                cells.Letterspace = row[this.Letterspace];
                cells.Locale = row[this.Locale];
                cells.LocalizeFont = row[this.LocalizeFont];
                cells.Overline = row[this.Overline];
                cells.Perpendicular = row[this.Perpendicular];
                cells.Pos = row[this.Pos];
                cells.RTLText = row[this.RTLText];
                cells.Strikethru = row[this.Strikethru];
                cells.UseVertical = row[this.UseVertical];

                return cells;
            }
        }
    }
}
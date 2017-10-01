using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Text
{
    public class ParagraphFormatCells : CellGroupMultiRow
    {
        public CellValueLiteral IndentFirst { get; set; }
        public CellValueLiteral IndentRight { get; set; }
        public CellValueLiteral IndentLeft { get; set; }
        public CellValueLiteral SpacingBefore { get; set; }
        public CellValueLiteral SpacingAfter { get; set; }
        public CellValueLiteral SpacingLine { get; set; }
        public CellValueLiteral HorizontalAlign { get; set; }
        public CellValueLiteral Bullet { get; set; }
        public CellValueLiteral BulletFont { get; set; }
        public CellValueLiteral BulletFontSize { get; set; }
        public CellValueLiteral LocalizeBulletFont { get; set; }
        public CellValueLiteral TextPosAfterBullet { get; set; }
        public CellValueLiteral Flags { get; set; }
        public CellValueLiteral BulletString { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.ParaIndentLeft, this.IndentLeft, nameof(SrcConstants.ParaIndentLeft), nameof(this.IndentLeft));
                yield return SrcValuePair.Create(SrcConstants.ParaIndentFirst, this.IndentFirst, nameof(SrcConstants.ParaIndentFirst), nameof(this.IndentFirst));
                yield return SrcValuePair.Create(SrcConstants.ParaIndentRight, this.IndentRight, nameof(SrcConstants.ParaIndentRight), nameof(this.IndentRight));
                yield return SrcValuePair.Create(SrcConstants.ParaSpacingAfter, this.SpacingAfter, nameof(SrcConstants.ParaSpacingAfter), nameof(this.SpacingAfter));
                yield return SrcValuePair.Create(SrcConstants.ParaSpacingBefore, this.SpacingBefore, nameof(SrcConstants.ParaSpacingBefore), nameof(this.SpacingBefore));
                yield return SrcValuePair.Create(SrcConstants.ParaSpacingLine, this.SpacingLine, nameof(SrcConstants.ParaSpacingLine), nameof(this.SpacingLine));
                yield return SrcValuePair.Create(SrcConstants.ParaHorizontalAlign, this.HorizontalAlign, nameof(SrcConstants.ParaHorizontalAlign), nameof(this.HorizontalAlign));
                yield return SrcValuePair.Create(SrcConstants.ParaBulletFont, this.BulletFont, nameof(SrcConstants.ParaBulletFont), nameof(this.BulletFont));
                yield return SrcValuePair.Create(SrcConstants.ParaBullet, this.Bullet, nameof(SrcConstants.ParaBullet), nameof(this.Bullet));
                yield return SrcValuePair.Create(SrcConstants.ParaBulletFontSize, this.BulletFontSize, nameof(SrcConstants.ParaBulletFontSize), nameof(this.BulletFontSize));
                yield return SrcValuePair.Create(SrcConstants.ParaLocalizeBulletFont, this.LocalizeBulletFont, nameof(SrcConstants.ParaLocalizeBulletFont), nameof(this.LocalizeBulletFont));
                yield return SrcValuePair.Create(SrcConstants.ParaTextPosAfterBullet, this.TextPosAfterBullet, nameof(SrcConstants.ParaTextPosAfterBullet), nameof(this.TextPosAfterBullet));
                yield return SrcValuePair.Create(SrcConstants.ParaFlags, this.Flags, nameof(SrcConstants.ParaFlags), nameof(this.Flags));
                yield return SrcValuePair.Create(SrcConstants.ParaBulletString, this.BulletString, nameof(SrcConstants.ParaBulletString), nameof(this.BulletString));
            }
        }

        public static List<List<ParagraphFormatCells>> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCells(page, shapeids, type);
        }

        public static List<ParagraphFormatCells> GetCells(IVisio.Shape shape, CellValueType type)
        {
            var query = lazy_query.Value;
            return query.GetCells(shape, type);
        }


        private static readonly System.Lazy<ParagraphFormatCellsReader> lazy_query = new System.Lazy<ParagraphFormatCellsReader>();


        class ParagraphFormatCellsReader : ReaderMultiRow<Text.ParagraphFormatCells>
        {
            public SectionQueryColumn Bullet { get; set; }
            public SectionQueryColumn BulletFont { get; set; }
            public SectionQueryColumn BulletFontSize { get; set; }
            public SectionQueryColumn BulletString { get; set; }
            public SectionQueryColumn Flags { get; set; }
            public SectionQueryColumn HorizontalAlign { get; set; }
            public SectionQueryColumn IndentFirst { get; set; }
            public SectionQueryColumn IndentLeft { get; set; }
            public SectionQueryColumn IndentRight { get; set; }
            public SectionQueryColumn LocalizeBulletFont { get; set; }
            public SectionQueryColumn SpaceAfter { get; set; }
            public SectionQueryColumn SpaceBefore { get; set; }
            public SectionQueryColumn SpaceLine { get; set; }
            public SectionQueryColumn TextPosAfterBullet { get; set; }

            public ParagraphFormatCellsReader()
            {
                var sec = this.query.SectionQueries.Add(IVisio.VisSectionIndices.visSectionParagraph);
                this.Bullet = sec.Columns.Add(SrcConstants.ParaBullet, nameof(this.Bullet));
                this.BulletFont = sec.Columns.Add(SrcConstants.ParaBulletFont, nameof(this.BulletFont));
                this.BulletFontSize = sec.Columns.Add(SrcConstants.ParaBulletFontSize, nameof(this.BulletFontSize));
                this.BulletString = sec.Columns.Add(SrcConstants.ParaBulletString, nameof(this.BulletString));
                this.Flags = sec.Columns.Add(SrcConstants.ParaFlags, nameof(this.Flags));
                this.HorizontalAlign = sec.Columns.Add(SrcConstants.ParaHorizontalAlign, nameof(this.HorizontalAlign));
                this.IndentFirst = sec.Columns.Add(SrcConstants.ParaIndentFirst, nameof(this.IndentFirst));
                this.IndentLeft = sec.Columns.Add(SrcConstants.ParaIndentLeft, nameof(this.IndentLeft));
                this.IndentRight = sec.Columns.Add(SrcConstants.ParaIndentRight, nameof(this.IndentRight));
                this.LocalizeBulletFont = sec.Columns.Add(SrcConstants.ParaLocalizeBulletFont, nameof(this.LocalizeBulletFont));
                this.SpaceAfter = sec.Columns.Add(SrcConstants.ParaSpacingAfter, nameof(this.SpaceAfter));
                this.SpaceBefore = sec.Columns.Add(SrcConstants.ParaSpacingBefore, nameof(this.SpaceBefore));
                this.SpaceLine = sec.Columns.Add(SrcConstants.ParaSpacingLine, nameof(this.SpaceLine));
                this.TextPosAfterBullet = sec.Columns.Add(SrcConstants.ParaTextPosAfterBullet, nameof(this.TextPosAfterBullet));
            }

            public override Text.ParagraphFormatCells ToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
            {
                var cells = new Text.ParagraphFormatCells();
                cells.IndentFirst = row[this.IndentFirst];
                cells.IndentLeft = row[this.IndentLeft];
                cells.IndentRight = row[this.IndentRight];
                cells.SpacingAfter = row[this.SpaceAfter];
                cells.SpacingBefore = row[this.SpaceBefore];
                cells.SpacingLine = row[this.SpaceLine];
                cells.HorizontalAlign = row[this.HorizontalAlign];
                cells.Bullet = row[this.Bullet];
                cells.BulletFont = row[this.BulletFont];
                cells.BulletFontSize = row[this.BulletFontSize];
                cells.LocalizeBulletFont = row[this.LocalizeBulletFont];
                cells.TextPosAfterBullet = row[this.TextPosAfterBullet];
                cells.Flags = row[this.Flags];
                cells.BulletString = row[this.BulletString];

                return cells;
            }
        }

    }
} 
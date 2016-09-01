using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Scripting.Commands
{
    public class AlignCommands : CommandSet
    {
        internal AlignCommands(Client client) :
            base(client)
        {

        }

        public void AlignHorizontal(TargetShapes targets, VisioAutomation.Drawing.Layout.AlignmentHorizontal align)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 2)
            {
                return;
            }

            IVisio.Enums.VisHorizontalAlignTypes halign;
            var valign = IVisio.Enums.VisVerticalAlignTypes.visVertAlignNone;

            switch (align)
            {
                case VisioAutomation.Drawing.Layout.AlignmentHorizontal.Left:
                    halign = IVisio.Enums.VisHorizontalAlignTypes.visHorzAlignLeft;
                    break;
                case VisioAutomation.Drawing.Layout.AlignmentHorizontal.Center:
                    halign = IVisio.Enums.VisHorizontalAlignTypes.visHorzAlignCenter;
                    break;
                case VisioAutomation.Drawing.Layout.AlignmentHorizontal.Right:
                    halign = IVisio.Enums.VisHorizontalAlignTypes.visHorzAlignRight;
                    break;
                default: throw new System.ArgumentOutOfRangeException();
            }

            const bool glue_to_guide = false;

            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Align Shapes"))
            {
                var selection = this._client.Selection.Get();
                selection.Align(halign, valign, glue_to_guide);
            }
        }

        public void AlignVertical(TargetShapes targets, VisioAutomation.Drawing.Layout.AlignmentVertical align)
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            int shape_count = targets.SetSelectionGetSelectedCount(this._client);
            if (shape_count < 2)
            {
                return;
            }

            // Set the align enums
            var halign = IVisio.Enums.VisHorizontalAlignTypes.visHorzAlignNone;
            IVisio.Enums.VisVerticalAlignTypes valign;
            switch (align)
            {
                case VisioAutomation.Drawing.Layout.AlignmentVertical.Top: valign = IVisio.Enums.VisVerticalAlignTypes.visVertAlignTop; break;
                case VisioAutomation.Drawing.Layout.AlignmentVertical.Center: valign = IVisio.Enums.VisVerticalAlignTypes.visVertAlignMiddle; break;
                case VisioAutomation.Drawing.Layout.AlignmentVertical.Bottom: valign = IVisio.Enums.VisVerticalAlignTypes.visVertAlignBottom; break;
                default: throw new System.ArgumentOutOfRangeException();
            }

            const bool glue_to_guide = false;

            // Perform the alignment
            var application = this._client.Application.Get();
            using (var undoscope = this._client.Application.NewUndoScope("Align Shapes"))
            {
                var selection = this._client.Selection.Get();
                selection.Align(halign, valign, glue_to_guide);
            }
        }

    }
}
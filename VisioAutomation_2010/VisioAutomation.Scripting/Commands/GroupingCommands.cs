using VisioAutomation.Scripting.Exceptions;
using IVisio=NetOffice.VisioApi;

namespace VisioAutomation.Scripting.Commands
{
    public class GroupingCommands: CommandSet
    {
        internal GroupingCommands(Client client) :
            base(client)
        {

        }


        public IVisio.IVShape Group()
        {
            this._client.Application.AssertApplicationAvailable();
            this._client.Document.AssertDocumentAvailable();

            // No shapes provided, use the active selection
            if (!this._client.Selection.HasShapes())
            {
                throw new VisioOperationException("No Selected Shapes to Group");
            }

            // the other way of doing this: this.Client.VisioApplication.DoCmd((short)IVisio.Enums.VisUICmds.visCmdObjectGroup);
            // but it doesn't return the group

            var selection = this._client.Selection.Get();
            var g = selection.Group();
            return g;
        }

        public void Ungroup(TargetShapes targets)
        {
            this._client.Application.AssertApplicationAvailable();
            if (targets.Shapes == null)
            {
                if (this._client.Selection.HasShapes())
                {
                    var application = this._client.Application.Get();
                    application.DoCmd((short)IVisio.Enums.VisUICmds.visCmdObjectUngroup);
                }
            }
            else
            {
                foreach (var shape in targets.Shapes)
                {
                    shape.Ungroup();
                }
            }
        }
    }
}
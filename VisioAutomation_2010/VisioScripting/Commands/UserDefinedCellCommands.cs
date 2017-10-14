using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Commands
{
    public class UserDefinedCellCommands : CommandSet
    {
        internal UserDefinedCellCommands(Client client) :
            base(client)
        {

        }

        public Dictionary<IVisio.Shape, Dictionary<string,UserDefinedCellCells>> GetHyperlinkCellsFromShapes(VisioScripting.Models.TargetShapes targets, CellValueType cvt)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument | CommandTargetFlags.ActivePage);


            var prop_dic = new Dictionary<IVisio.Shape, Dictionary<string, UserDefinedCellCells>>();

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return prop_dic;
            }

            var page = cmdtarget.ActivePage;
            var list_user_props = UserDefinedCellHelper.GetDictionary((IVisio.Page) page , targets.Shapes, cvt);

            for (int i = 0; i < targets.Shapes.Count; i++)
            {
                var shape = targets.Shapes[i];
                var props = list_user_props[i];
                prop_dic[shape] = props;
            }

            return prop_dic;
        }

        public List<bool> ShapesContainUserDefinedCellsWithName(VisioScripting.Models.TargetShapes targets, string name)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);

            if (name == null)
            {
                throw new System.ArgumentNullException(nameof(name));
            }

            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return new List<bool>();
            }

            var all_shapes = this._client.Selection.GetShapesInSelection();
            var results = all_shapes.Select(s => UserDefinedCellHelper.Contains(s, name)).ToList();

            return results;
        }
       
        public void DeleteUserDefinedCellsByName(VisioScripting.Models.TargetShapes targets, string name)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            } 

            if (name == null)
            {
                throw new System.ArgumentNullException("name cannot be null","name");
            }

            if (name.Length < 1)
            {
                throw new System.ArgumentException("name cannot be empty", nameof(name));
            }

            using (var undoscope = this._client.Application.NewUndoScope("Delete User-Defined Cell"))
            {
                foreach (var shape in targets.Shapes)
                {
                    UserDefinedCellHelper.Delete(shape, name);
                }
            }
        }

        public void SetUserDefinedCell(VisioScripting.Models.TargetShapes targets, VisioScripting.Models.UserDefinedCell cell)
        {
            var cmdtarget = this._client.GetCommandTarget( CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument);


            targets = targets.ResolveShapes(this._client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            using (var undoscope = this._client.Application.NewUndoScope("Set User-Defined Cell"))
            {
                foreach (var shape in targets.Shapes)
                {
                    UserDefinedCellHelper.Set(shape, cell.Name, cell.Cells);
                }
            }
        }
    }
}
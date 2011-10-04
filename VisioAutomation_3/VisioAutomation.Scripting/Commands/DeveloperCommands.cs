using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Scripting.Commands
{
    public class DeveloperCommands : CommandSet
    {
        public DeveloperCommands(Session session) :
            base(session)
        {

        }

        public void HelloWorld()
        {
            if (this.Session.VisioApplication == null)
            {
                this.Session.Application.New();
            }

            var doc = this.Session.Document.New(8.5, 11);
            var pages = doc.Pages;
            var page = pages.Add();

            var s0 = page.DrawRectangle(2, 2, 6, 6);
            s0.Text = "Hello World";
        }

        public System.Xml.Linq.XElement GetXMLDescription()
        {
            var el_shapes = new System.Xml.Linq.XElement("Shapes");
            if (!this.Session.HasSelectedShapes())
            {
                return el_shapes;
            }

            var page = this.Session.VisioApplication.ActivePage;
            var shapes = page.Shapes.AsEnumerable().ToList();
            var shapeids = shapes.Select(s => s.ID).ToList();

            var el_shape = VA.ShapeHelper.GetShapeDescriptionXML(page, shapeids);

            foreach (var x in el_shape)
            {
                el_shapes.Add(x);
            }

            return el_shapes;
        }

        public IVisio.Document DrawScriptingDocumentation()
        {
            var pagesize = new VA.Drawing.Size(8.5, 11);
            var docbuilder = new VA.Experimental.SimpleTextDoc.TextDocumentBuilder(this.Session.VisioApplication, pagesize);
            docbuilder.BodyParaSpacingAfter = 6.0;
            var lines = new List<string>();

            var cmdst_props = VA.Scripting.Session.GetCommandSetProperties().OrderBy(i=>i.Name).ToList();
            var sb = new System.Text.StringBuilder();
            var helpstr = new System.Text.StringBuilder();

            docbuilder.Start();
            foreach (var cmdset_prop in cmdst_props)
            {
                var cmdset_type = cmdset_prop.PropertyType;

                // Calculate the text
                var methods = CommandSet.GetCommandMethods(cmdset_type);
                lines.Clear();
                foreach (var method in methods)
                {
                    sb.Length = 0;
                    var method_params = method.GetParameters();
                    TextUtil.Join(sb, ", ", method_params.Select(param => string.Format("{0} {1}", ReflectionUtil.GetNiceTypeName(param.ParameterType), param.Name)));

                    if (method.ReturnType != typeof(void))
                    {
                        string line = string.Format("{0}({1}) -> {2}", method.Name, sb, ReflectionUtil.GetNiceTypeName(method.ReturnType));
                        lines.Add(line);
                    }
                    else
                    {
                        string line = string.Format("{0}({1})", method.Name, sb);
                        lines.Add(line);
                    }
                }

                lines.Sort();
                
                helpstr.Length = 0;
                TextUtil.Join(helpstr,"\r\n",lines);

                var docpage = new VisioAutomation.Experimental.SimpleTextDoc.TextPage();
                docpage.Title = cmdset_prop.Name + " commands";
                docpage.Body = helpstr.ToString();
                docpage.Name = cmdset_prop.Name + " commands";

                docbuilder.Draw(docpage);
            }

            docbuilder.Finish();
            docbuilder.VisioDocument.Subject = "VisioAutomation.Scripting Documenation";
            docbuilder.VisioDocument.Title = "VisioAutomation.Scripting Documenation";
            docbuilder.VisioDocument.Creator = "";
            docbuilder.VisioDocument.Company = "";

            return docbuilder.VisioDocument;
        }

        public IVisio.Document DrawInteropEnumDocumentation()
        {
            var pagesize = new VA.Drawing.Size(8.5, 11);
            var docbuilder = new VA.Experimental.SimpleTextDoc.TextDocumentBuilder(this.Session.VisioApplication, pagesize);
            //docbuilder.BodyParaSpacingAfter = 2.0;
            docbuilder.BodyTextSize = 8.0;
            var helpstr = new System.Text.StringBuilder();
            int chunksize = 70;

            var interop_enums = VA.Interop.InteropHelper.GetEnums();
            docbuilder.Start();
            int pagecount = 0;
            foreach (var enum_ in interop_enums)
            {


                int chunkcount = 0;

                var values = enum_.Values.OrderBy(i => i.Name).ToList();
                foreach (var chunk in Chunk(values, chunksize))
                {
                    helpstr.Length = 0;
                    foreach (var val in chunk)
                    {
                        helpstr.AppendFormat("0x{0}\t{1}\n", val.Value.ToString("x"),val.Name);

                    }

                    var docpage = new VA.Experimental.SimpleTextDoc.TextPage();
                    docpage.Title = enum_.Name;
                    docpage.Body = helpstr.ToString();
                    if (chunkcount == 0)
                    {
                        docpage.Name = string.Format("{0}", enum_.Name);
                        
                    }
                    else
                    {
                        docpage.Name = string.Format("{0} ({1})", enum_.Name, chunkcount + 1);
                    }

                    docbuilder.Draw(docpage);

                    var tabstops = new[]
                                 {
                                     new VA.Text.TabStop(1.5, VA.Text.TabStopAlignment.Left)
                                 };
                    VA.Text.TextHelper.SetTabStops(docpage.VisioBodyShape, tabstops);
                    
                    chunkcount++;
                    pagecount++;
                }

            }

            docbuilder.Finish();
            docbuilder.VisioDocument.Subject = "Visio Interop Enum Documenation";
            docbuilder.VisioDocument.Title = "Visio Interop Enum Documenation";
            docbuilder.VisioDocument.Creator = "";
            docbuilder.VisioDocument.Company = "";

            return docbuilder.VisioDocument;
        }

        private class PathTreeBuilder
        {
            public Dictionary<string, string> dic;
            public string root;

            public PathTreeBuilder()
            {
                this.dic = new Dictionary<string, string>();
            }

            public void Add(string path)
            {
                if (this.dic.ContainsKey(path))
                {
                    return;
                }

                var seps = new char[] { '.' };
                var tokens = path.Split(seps);
                if (tokens.Length == 0)
                {
                    throw new VA.AutomationException();
                }
                else if (tokens.Length == 1)
                {
                    this.root = tokens[0];
                    this.dic[this.root] = null;

                }
                else
                {
                    string ps = string.Join(".", tokens.Take(tokens.Length - 1));
                    this.Add(ps);
                    this.dic[path] = ps;
                }
                
            }

        }

        public IVisio.Document DrawVANamespaces()
        {
            var doc = this.Session.Document.New(8.5,11);

            var namespaces_1 = GetVANamespaces();
            var pb = new PathTreeBuilder();
            foreach (string ns in namespaces_1)
            {
                pb.Add(ns);
            }

            var namespaces_2 = pb.dic.Keys.ToList();

            var msagl_drawing = new VA.Layout.MSAGL.Drawing();
            var ns_shape_map = new Dictionary<string, VA.Layout.MSAGL.Shape>(namespaces_2.Count);
            foreach (string ns in namespaces_2)
            {
                string label = ns;
                int n = ns.LastIndexOf(".");
                if (n > 0)
                {
                    label = ns.Substring(n+1);
                }
                var s = msagl_drawing.AddShape(ns,label,"basic_u.vss","Rectangle");
                ns_shape_map[ns] = s;
            }

            foreach (string ns in namespaces_2)
            {
                var pns = pb.dic[ns];
                if (pns != null)
                {
                    msagl_drawing.Connect(pns+"_to_"+ns,ns_shape_map[pns], ns_shape_map[ns],null, VA.Connections.ConnectorType.Straight);
                    
                }
            }

            var layout_options = new VA.Layout.MSAGL.LayoutOptions();
            msagl_drawing.Render(doc.Application.ActivePage,layout_options);

            return doc;
        }

        private static HashSet<string> GetVANamespaces()
        {
            var types = VA.Experimental.Developer.DeveloperHelper.GetAllTypes();
            var namespaces = new HashSet<string>();
            foreach (var type in types)
            {
                namespaces.Add(type.Namespace);
            }
            return namespaces;
        }

        public IList<VA.Interop.EnumType> GetInteropEnums()
        {
            return VA.Interop.InteropHelper.GetEnums();
        }

        public VA.Interop.EnumType GetInteropEnum(string name)
        {
            return VA.Interop.InteropHelper.GetEnum(name);
        }

        private static IEnumerable<IEnumerable<T>> Chunk<T>(IEnumerable<T> source, int chunksize)
        {
            while (source.Any())
            {
                yield return source.Take(chunksize);
                source = source.Skip(chunksize);
            }
        }
    }
}


namespace VisioAutomation.Experimental.Developer
{
    public class DeveloperHelper
    {
        public static List<System.Type> GetTypes()
        {
            // find the VA assembly
            var vat = typeof (VisioAutomation.ApplicationHelper);
            var asm = vat.Assembly;

            // TODO: Consider filtering out types that should *not* be exposed despite being public
            var types = asm.GetExportedTypes().Where(t => t.IsPublic).ToList();
            return types;
        }

        public static List<System.Type> GetAllTypes()
        {
            // find the VA assembly
            var vat = typeof(VisioAutomation.ApplicationHelper);
            var asm = vat.Assembly;

            var types = asm.GetExportedTypes().ToList();
            return types;
        }

    }
}
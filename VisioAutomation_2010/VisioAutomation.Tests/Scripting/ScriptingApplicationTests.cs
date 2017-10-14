using System.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingApplicationTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Test_Application_Window()
        {
            this.Scripting_Test_Resize_Application_Window1();
            this.Scripting_Test_Resize_Application_Window2();
            this.Scripting_Test_App_to_Front();
        }

        public void Scripting_Test_Resize_Application_Window1()
        {
            var desired_size = new Size(600, 800);
            var client = this.GetScriptingClient();
            var old_rect = client.Window.GetApplicationWindowRectangle();
            var new_rect = new System.Drawing.Rectangle(old_rect.X, old_rect.Y, desired_size.Width, desired_size.Height);

            client.Window.SetApplicationWindowRectangle(new_rect);
            var actual_rect = client.Window.GetApplicationWindowRectangle();
            Assert.AreEqual(desired_size, actual_rect.Size);
        }

        public void Scripting_Test_Resize_Application_Window2()
        {
            var client = this.GetScriptingClient();
            var page_size = new VisioAutomation.Geometry.Size(10,5);
            var doc = client.Document.NewDocument(page_size);
            var pagesize = client.Page.GetActivePageSize();
            Assert.AreEqual(10.0, pagesize.Width);
            Assert.AreEqual(5.0, pagesize.Height);
            Assert.AreEqual(0, client.Selection.GetActiveSelection().Count);
            client.Draw.DrawRectangle(1, 1, 2, 2);
            Assert.AreEqual(1, client.Selection.GetActiveSelection().Count);

            client.Document.CloseActiveDocument(true);
        }

        public void Scripting_Test_App_to_Front()
        {
            var client = this.GetScriptingClient();
            client.Window.MoveApplicationWindowToFront();
        }

        [TestMethod]
        public void Scripting_Undo_Scenarios()
        {
            var client = this.GetScriptingClient();
            var page_size = new VisioAutomation.Geometry.Size(8.5,11);
            var drawing = client.Document.NewDocument(page_size);
            var page = client.Page.NewPage(page_size, false);
            Assert.AreEqual(0, page.Shapes.Count);
            page.DrawRectangle(1, 1, 3, 3);
            Assert.AreEqual(1, page.Shapes.Count);
            client.Application.Undo();
            Assert.AreEqual(0, page.Shapes.Count);
            client.Document.CloseActiveDocument(true);
        }

        [TestMethod]
        public void Scripting_CloseDocument_Scenarios()
        {
            var page_size = new VisioAutomation.Geometry.Size(8.5, 11);
            var client = this.GetScriptingClient();
            var doc1 = client.Document.NewDocument(page_size);
            var doc2 = client.Document.NewDocument(page_size);
            var doc3 = client.Document.NewDocument(page_size);

            client.Document.CloseAllDocumentsWithoutSaving();

            Assert.IsFalse(client.Document.HasActiveDocument);
            var application = client.Application.GetApplication();
            var documents = application.Documents;
            Assert.AreEqual(0, documents.Count);
        }

        [TestMethod]
        public void XXX()
        {
            var client = this.GetScriptingClient();
            client.Application.NewApplication();
            client.Document.NewDocument();
        }
    }
}
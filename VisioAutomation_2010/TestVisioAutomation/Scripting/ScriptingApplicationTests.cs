using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingApplicationTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Test_Resize_Application_Window()
        {
            var ss = GetScriptingSession();

            var old_size = ss.Application.Window.GetSize();
            var desired_size = new System.Drawing.Size(600, 800);

            ss.Application.Window.SetSize(desired_size.Width, desired_size.Height);

            var actual_size = ss.Application.Window.GetSize();
            Assert.AreEqual(desired_size, actual_size);
            ss.Application.Window.SetSize(old_size.Width, old_size.Height);
            actual_size = ss.Application.Window.GetSize();
            Assert.AreEqual(old_size, actual_size);
        }

        [TestMethod]
        public void Scripting_Test_Resize_Application_Window2()
        {
            var ss = GetScriptingSession();

            var doc = ss.Document.New(10, 5);

            Assert.IsTrue(ss.HasActiveDrawing);

            var pagesize = ss.Page.GetSize();
            Assert.AreEqual(10.0, pagesize.Width);
            Assert.AreEqual(5.0, pagesize.Height);
            Assert.AreEqual(0, ss.Layout.GetSelectedShapeCount());
            ss.Draw.Rectangle(1, 1, 2, 2);
            Assert.AreEqual(1, ss.Layout.GetSelectedShapeCount());

            ss.Document.CloseAllWithoutSaving();
        }

        [TestMethod]
        public void Scripting_Test_App_to_Front()
        {
            var ss = GetScriptingSession();
            ss.Application.Window.ToFront();
        }

        [TestMethod]
        public void Scripting_Test_Undo()
        {
            var ss = GetScriptingSession();
            var drawing = ss.Document.New(8.5, 11);
            var page = ss.Page.New(new VA.Drawing.Size(8.5, 11), false);
            Assert.AreEqual(0, page.Shapes.Count);
            page.DrawRectangle(1, 1, 3, 3);
            Assert.AreEqual(1, page.Shapes.Count);
            ss.Application.Undo();
            Assert.AreEqual(0, page.Shapes.Count);
            ss.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Test_close_all()
        {
            var ss = GetScriptingSession();
            var doc1 = ss.Document.New(10, 5);
            var doc2 = ss.Document.New(10, 5);
            var doc3 = ss.Document.New(10, 5);

            ss.Document.CloseAllWithoutSaving();

            Assert.IsFalse(ss.HasActiveDrawing);
            var application = ss.VisioApplication;
            var documents = application.Documents;
            Assert.AreEqual(0, documents.Count);
        }
    }
}
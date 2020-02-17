using System;
using System.Linq;
using OneNoteInterop = Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using Windows.Storage.Streams;
using Windows.UI.Input.Inking;
using System.Threading.Tasks;
using System.Collections;

namespace LibOneInk
{
    public class OneNote2016 : OneNote
    {
        private OneNoteInterop.Application _oneNoteApp = null;
        private readonly XNamespace ns;

        private readonly InkStrokeBuilder _strokeBuilder = new InkStrokeBuilder();

        public OneNote2016()
        {
            Logger.Info($"Starting OneNote2016");
            _oneNoteApp = new OneNoteInterop.Application();
            ns = GetNamespace();
        }

        private XNamespace GetNamespace()
        {
            if (_oneNoteApp == null)
                throw new InvalidOperationException("OneNote is not initialized");

            _oneNoteApp.GetHierarchy(null, OneNoteInterop.HierarchyScope.hsNotebooks, out string xml);

            var doc = XDocument.Parse(xml);
            return doc.Root.Name.Namespace;
        }

        protected override Task<string> GetObjectId(string parentId, OneNote.HierarchyScope scope, string objectName)
        {
            if (_oneNoteApp == null)
                throw new InvalidOperationException("OneNote is not initialized");

            OneNoteInterop.HierarchyScope intropScope = GetInteropScope(scope);
            _oneNoteApp.GetHierarchy(parentId, intropScope, out string xml);

            var doc = XDocument.Parse(xml);
            string nodeName = GetNodeName(scope);

            var node = doc.Descendants(ns + nodeName).FirstOrDefault(n => n.Attribute("name").Value == objectName);
            if (node == null)
                return Task.FromResult((string)null);
            return Task.FromResult(node.Attribute("ID").Value);
        }

        protected override Task<string> GetObjectName(string parentId, OneNote.HierarchyScope scope, string objectId)
        {
            if (_oneNoteApp == null)
                throw new InvalidOperationException("OneNote is not initialized");

            OneNoteInterop.HierarchyScope intropScope = GetInteropScope(scope);
            _oneNoteApp.GetHierarchy(parentId, intropScope, out string xml);

            var doc = XDocument.Parse(xml);
            string nodeName = GetNodeName(scope);

            var node = doc.Descendants(ns + nodeName).FirstOrDefault(n => n.Attribute("ID").Value == objectId);
            if (node == null)
                return Task.FromResult((string)null);
            return Task.FromResult(node.Attribute("name").Value);
        }

        private OneNoteInterop.HierarchyScope GetInteropScope(OneNote.HierarchyScope scope)
        {
            switch (scope)
            {
                case OneNote.HierarchyScope.Notebooks:
                    return OneNoteInterop.HierarchyScope.hsNotebooks;
                case OneNote.HierarchyScope.Pages:
                    return OneNoteInterop.HierarchyScope.hsPages;
                case OneNote.HierarchyScope.Sections:
                    return OneNoteInterop.HierarchyScope.hsSections;
            }
            return OneNoteInterop.HierarchyScope.hsChildren;
        }

        private string GetNodeName(OneNote.HierarchyScope scope)
        {
            switch (scope)
            {
                case OneNote.HierarchyScope.Notebooks:
                    return "Notebook";
                case OneNote.HierarchyScope.Pages:
                    return "Page";
                case OneNote.HierarchyScope.Sections:
                    return "Section";
            }
            return null;
        }

        public override Task<object> CreatePage(string name) => Task.FromResult((object)CreatePage(_sectionId, name));

        private XDocument CreatePage(string sectionId, string name)
        {
            _oneNoteApp.CreateNewPage(sectionId, out string pageId, OneNoteInterop.NewPageStyle.npsBlankPageWithTitle);

            _oneNoteApp.GetPageContent(pageId, out string xml, OneNoteInterop.PageInfo.piAll);

            XDocument page = XDocument.Parse(xml);
            var title = page.Descendants(ns + "T").First();
            title.Value = name;
            UpdatePageContent(page);
            return page;
        }

        public override Task UpdatePageContent(object page)
        {
            if (page == null)
            {
                Logger.Error("Invalid page");
                return Task.CompletedTask;
            }

            XDocument xdoc = page as XDocument;
            if (xdoc == null)
                throw new ArgumentException("doc is not a proper OneNote2016 document");

            _oneNoteApp.UpdatePageContent(xdoc.ToString());

            return Task.CompletedTask;
        }


        private async Task<string> SerializeStroke(InkManager ink)
        {
            using (InMemoryRandomAccessStream ms = new InMemoryRandomAccessStream())
            {
                await ink.SaveAsync(ms);
                byte[] res = new byte[ms.Position];
                DataReader dr = new DataReader(ms.GetInputStreamAt(0));
                await dr.LoadAsync((uint)ms.Position);
                dr.ReadBytes(res);
                return Convert.ToBase64String(res);
            }
        }


        public override async Task AddStrokeGroup(object page, OneInkStrokeGroup group)
        {
            if (page == null)
            {
                Logger.Error("Invalid document");
                return;
            }

            XDocument xpage = page as XDocument;
            if (xpage == null)
                throw new ArgumentException("doc is not a proper OneNote2016 document");

            var root = xpage.Descendants(ns + "Page").First();
            if (root == null)
            {
                Logger.Error("Unable to find root");
                return;
            }

            var outline = new XElement(ns + "Outline");
            root.Add(outline);
            var oechild = new XElement(ns + "OEChildren");
            outline.Add(oechild);
            var oe = new XElement(ns + "OE");
            oechild.Add(oe);

            InkManager ink = new InkManager();

            foreach (OneInkStroke stroke in group.Strokes)
            {
                ArrayList inkPoints = new ArrayList(stroke.Points.Count);
                foreach(var pt in stroke.Points)
                {
                    inkPoints.Add(new InkPoint(new Windows.Foundation.Point(pt.x, pt.y), pt.pressure));
                }
                
                InkStroke inkStroke = _strokeBuilder.CreateStrokeFromInkPoints(inkPoints.Cast<InkPoint>(), System.Numerics.Matrix3x2.Identity);
                ink.AddStroke(inkStroke);
            }

            string base64 = await SerializeStroke(ink);

            var drawing = new XElement(ns + "InkDrawing");
            oe.Add(drawing);
            var data = new XElement(ns + "Data");
            data.SetValue(base64);
            drawing.Add(data);
        }

        #region IDisposable Support
        private bool disposedValue = false;

        protected override void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    _oneNoteApp = null;
                    GC.Collect();
                }
                disposedValue = true;
            }
        }
        #endregion

    }

}


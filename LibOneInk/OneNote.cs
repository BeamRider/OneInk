using System;
using System.Threading.Tasks;

namespace LibOneInk
{
    public abstract class OneNote : IDisposable
    {
        public enum HierarchyScope
        {
            Notebooks,
            Sections,
            Pages
        }

        protected static readonly log4net.ILog Logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        protected string _notebookId = null;
        public string NotebookId { get { return _notebookId; } }

        protected string _notebook = null;
        public string Notebook { get { return _notebook; } }

        protected string _sectionId = null;
        public string SectionId { get { return _sectionId; } }

        protected string _section = null;
        public string Section { get { return _section; } }

        public async Task SetNotebookId(string notebookid)
        {
            _notebookId = notebookid;
            await UpdateNotebook();
        }

        public async Task SetNotebook(string notebook)
        {
            _notebook = notebook;
            await UpdateNotebookId();
        }

        public async Task SetSectionId(string sectionid)
        {
            _sectionId = sectionid;
            await UpdateSection();
        }

        public async Task SetSection(string section)
        {
            _section = section;
            await UpdateSectionId();
        }

        private async Task UpdateNotebook()
        {
            _notebook = await GetObjectName(null, HierarchyScope.Notebooks, _notebookId);
            if (_notebook == null)
                Logger.Error($"Unable to find Notebook \"{_notebookId}\"");
            else
                Logger.Debug($"Notebook {Notebook} {_notebookId}");
        }

        private async Task UpdateNotebookId()
        {
            _notebookId = await GetObjectId(null, HierarchyScope.Notebooks, _notebook);
            if (_notebookId == null)
                Logger.Error($"Unable to find Notebook \"{Notebook}\"");
            else
                Logger.Debug($"Notebook {Notebook} {_notebookId}");
        }

        private async Task UpdateSection()
        {
            _section = await GetObjectName(_notebookId, HierarchyScope.Sections, _sectionId);
            if (_section == null)
                Logger.Error($"Unable to find SectionId \"{_sectionId}\"");
            else
                Logger.Debug($"Section {Section} {_sectionId}");
        }

        private async Task UpdateSectionId()
        {
            if (_notebookId == null)
                throw new InvalidOperationException("Notebook has not been initialized or not valid");

            _sectionId = await GetObjectId(_notebookId, HierarchyScope.Sections, _section);
            if (_sectionId == null)
                Logger.Error($"Unable to find Section \"{Section}\"");
            else
                Logger.Debug($"Section {Section} {_sectionId}");
        }

        protected abstract Task<string> GetObjectId(string parentId, HierarchyScope scope, string objectName);
        protected abstract Task<string> GetObjectName(string parentId, HierarchyScope scope, string objectId);

        public abstract Task<object> CreatePage(string name);

        public abstract Task UpdatePageContent(object page);

        public abstract Task AddStrokeGroup(object page, OneInkStrokeGroup stroke);


        #region IDisposable Support

        protected abstract void Dispose(bool disposing);

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~OneNote()
        {
            Dispose(false);
        }
        #endregion
    }
}

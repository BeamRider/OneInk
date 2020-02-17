using System;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading;

namespace LibOneInk
{
    public class SingleInstanceApplicationLock : IDisposable
    {
        private static readonly log4net.ILog Logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        ~SingleInstanceApplicationLock()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public bool TryAcquireExclusiveLock()
        {
            try
            {
                if (!_mutex.WaitOne(1000, false))
                    return false;
            }
            catch (AbandonedMutexException)
            {
                Logger.Warn("Abandoned mutex");
            }

            return _hasAcquiredExclusiveLock = true;
        }

        private readonly Mutex _mutex;
        private bool _hasAcquiredExclusiveLock, _disposed;

        public SingleInstanceApplicationLock(string mtxid)
        {
            _mutex = CreateMutex(mtxid);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing && !_disposed && _mutex != null)
            {
                try
                {
                    if (_hasAcquiredExclusiveLock)
                        _mutex.ReleaseMutex();

                    _mutex.Dispose();
                }
                finally
                {
                    _disposed = true;
                }
            }
        }

        private static Mutex CreateMutex(string mutexId)
        {
            var sid = new SecurityIdentifier(WellKnownSidType.WorldSid, null);
            var allowEveryoneRule = new MutexAccessRule(sid,
                MutexRights.FullControl, AccessControlType.Allow);

            var securitySettings = new MutexSecurity();
            securitySettings.AddAccessRule(allowEveryoneRule);

            var mutex = new Mutex(false, mutexId);
            mutex.SetAccessControl(securitySettings);

            return mutex;
        }

    }
}

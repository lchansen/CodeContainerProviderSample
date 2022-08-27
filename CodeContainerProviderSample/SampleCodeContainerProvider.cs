using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.CodeContainerManagement;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using CodeContainer = Microsoft.VisualStudio.Shell.CodeContainerManagement.CodeContainer;
using ICodeContainerProvider = Microsoft.VisualStudio.Shell.CodeContainerManagement.ICodeContainerProvider;

namespace CodeContainerProviderSample
{
    [Guid(GuidString)]
    internal class SampleCodeContainerProvider : ICodeContainerProvider
    {
        public const string GuidString = "13562B13-0D64-4DEB-9464-6F6202511FA3";
        public static Guid Guid = new Guid(GuidString);

        public async Task<CodeContainer> AcquireCodeContainerAsync(IProgress<ServiceProgressData> downloadProgress, CancellationToken cancellationToken)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            // Instead of showing a folder browser dialog here, an implementer might:
            // - Show UI with list of available containers to download
            // - await the download of the container
            var acquiredFolder = ShowFolderBrowserDialog(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Simulate Downloading a Code Container");

            if (!string.IsNullOrEmpty(acquiredFolder))
            {
                string name = Path.GetFileName(acquiredFolder);
                CodeContainerSourceControlProperties sccProperties = new CodeContainerSourceControlProperties(name, acquiredFolder, Guid);
                CodeContainerLocalProperties localProperties = new CodeContainerLocalProperties(acquiredFolder, CodeContainerType.Folder, sccProperties);

                // You might costruct this if you have a remote URL that this code container exists at
                RemoteCodeContainer remoteContainer = null;

                // If certain we downloaded the asset, we return a CodeContainer and VS will save the record of it, and open it.
                return new CodeContainer(localProperties, remoteContainer, isFavorite: false, lastAccessed: DateTimeOffset.UtcNow);
            }
            else
            {
                // If user cancelled, or download failed, we can return null and VS will keep the "Get to Code" dialog open.
                return null;
            }
        }

        public Task<CodeContainer> AcquireCodeContainerAsync(Microsoft.VisualStudio.Shell.CodeContainerManagement.RemoteCodeContainer onlineCodeContainer, IProgress<ServiceProgressData> downloadProgress, CancellationToken cancellationToken)
        {
            // This would be called when VS can't find the local path. It might have been deleted, or the user logged into VS on a new machine.
            // This method would be used to re-download and get the new path for the code container.
            throw new NotImplementedException();
        }

        /// <summary>
        /// Show a VS platform folder browser dialog
        /// </summary>
        private string ShowFolderBrowserDialog(string initialFolder, string dialogTitle)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                IVsUIShell uiShell = CodeContainerProviderSamplePackage.GetGlobalService(typeof(SVsUIShell)) as IVsUIShell;
                if (uiShell == null)
                {
                    return null;
                }

                const int MaxBuffer = 2048;
                VSBROWSEINFOW[] browseInfo = new VSBROWSEINFOW[1];

                ErrorHandler.ThrowOnFailure(uiShell.GetDialogOwnerHwnd(out IntPtr ownerHwnd));

                IntPtr dirNamePtr = Marshal.AllocCoTaskMem(MaxBuffer);
                try
                {
                    browseInfo[0].pwzInitialDir = initialFolder;
                    browseInfo[0].lStructSize = (uint)Marshal.SizeOf(typeof(VSBROWSEINFOW));
                    browseInfo[0].pwzDlgTitle = dialogTitle;
                    browseInfo[0].dwFlags = 0;
                    browseInfo[0].nMaxDirName = 1024;
                    browseInfo[0].pwzDirName = dirNamePtr;
                    browseInfo[0].hwndOwner = ownerHwnd;

                    return ErrorHandler.Succeeded(uiShell.GetDirectoryViaBrowseDlg(browseInfo))
                        ? Marshal.PtrToStringAuto(browseInfo[0].pwzDirName)
                        : null;
                }
                finally
                {
                    Marshal.FreeCoTaskMem(browseInfo[0].pwzDirName);
                }
            }
            catch
            {
                return null;
            }
        }
    }
}

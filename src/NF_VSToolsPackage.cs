using Community.VisualStudio.Toolkit;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json.Linq;
using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

// memo: troubleshooting:
// run activity log: win+r:
//   devenv /RootSuffix Exp /Log
// then open %APPDATA%\Microsoft\VisualStudio\17.0Exp\ActivityLog.xml

namespace NF.VSTools
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(NF_VSToolsPackage.PackageGuidString)]

    // Define a UI context that is true when a solution exists (or, if you want, also when no solution is open)
    [ProvideUIContextRule(
        "C5F2A8C3-6A7E-4C1D-9A6B-5C9E9B5C1F01",              // any stable new GUID you generate
        name: "NF.VSTools AutoLoad",
        expression: "SolutionExists",
        termNames: new[] { "SolutionExists" },
        termValues: new[] { VSConstants.UICONTEXT.SolutionExists_string }
    )]

    // Tell VS to load this package when that UI context is true
    [ProvideAutoLoad("C5F2A8C3-6A7E-4C1D-9A6B-5C9E9B5C1F01", PackageAutoLoadFlags.BackgroundLoad)]

    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    public sealed class NF_VSToolsPackage : AsyncPackage
    {
        /// <summary>
        /// NF_VSToolsPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "ab87c5f2-98c6-4b96-a28b-594b14e5b654";

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="token">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken token, IProgress<ServiceProgressData> progress)
        {
            // The package itself can load on a background thread (AsyncPackage).
            // Switch to UI thread only when needed.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(token);
            await Log.InitAsync(this); // UI thread; creates IVs pane
            Log.Sync("Package InitializeAsync called.");

            // Ensure ALL [Command] classes in this assembly get registered
            await this.RegisterCommandsAsync();
            Log.Post("Registered all commands async.");

            // old: individual registration
            //await DuplicateFilePairCommand.InitializeAsync(this);
            //Log.Post("DuplicateFilePairCommand initialized.");

            // Display if command is discoverable:
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var mcs = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            var setGuid = new Guid(PackageGuids.guidNF_VSToolsPackageCmdSetString);
            var cmdId = new CommandID(setGuid, PackageIds.cmdDuplicateFilePairCommand);

            var found = mcs?.FindCommand(cmdId) != null;
            Log.Post($"FindCommand(cmdDuplicateFilePairCommand) => {found}");
        }

        #endregion
    }
}

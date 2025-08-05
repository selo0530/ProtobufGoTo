using System;
using System.Diagnostics.CodeAnalysis;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;

using EnvDTE;
using EnvDTE80;

namespace ProtobufGoTo
{
	[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
	[InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
	[ProvideMenuResource("Menus.ctmenu", 1)]
	[Guid(ProtobufGoToPackage.PackageGuidString)]
    [ProvideAutoLoad(Microsoft.VisualStudio.Shell.Interop.UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(Microsoft.VisualStudio.Shell.Interop.UIContextGuids80.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
	public sealed class ProtobufGoToPackage : AsyncPackage
	{
        public const string PackageGuidString = "3ca337c6-3f46-473d-8bc8-c22a921b0215";

		public DTE2 m_dte;

		public ProtobufGoToPackage()
		{
		}

		protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
		{
			await base.InitializeAsync(cancellationToken, progress);
			await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

			ProtobufGoTo.Initialize(this);
		}
	}
}

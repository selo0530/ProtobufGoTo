using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.VCProjectEngine;

using System.Reflection;
using System.IO;
using EnvDTE;
using EnvDTE80;
using System.Linq;
using System.Text.RegularExpressions;

namespace ProtobufGoTo
{
	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed class ProtobufGoTo
	{
		public const int CommandId = 0x0100;

		public static readonly Guid CommandSet = new Guid("7c132991-dea1-4719-8c67-c20b24b6775c");

		private readonly Package package;

		private ProtobufGoTo(Package package)
		{
			if (package == null)
			{
				throw new ArgumentNullException("package");
			}

			this.package = package;

			OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
			if (commandService != null)
			{
				var menuCommandID = new CommandID(CommandSet, CommandId);
				var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
				commandService.AddCommand(menuItem);
			}
		}

		public static ProtobufGoTo Instance
		{
			get;
			private set;
		}

		private IServiceProvider ServiceProvider
		{
			get
			{
				return this.package;
			}
		}

		private System.Diagnostics.Process FBProcess
		{
			get;
			set;
		}

		public static void Initialize(Package package)
		{
			Instance = new ProtobufGoTo(package);
		}

		private void MenuItemCallback(object sender, EventArgs e)
		{
            ProtobufGoToPackage ProtoPackage = (ProtobufGoToPackage)this.package;
            var dte = ProtoPackage.m_dte;
            if (dte == null)
            {
                // DTE가 아직 준비되지 않았으면 즉시 GetService로 시도
                dte = ServiceProvider.GetService(typeof(DTE)) as DTE2;
                if (dte == null)
                    return; // 또는 사용자에게 안내
                ProtoPackage.m_dte = dte;
            }
            var doc = dte.ActiveDocument;
            if (doc == null || !doc.Name.EndsWith(".proto", StringComparison.OrdinalIgnoreCase))
                return;

            TextSelection selection = doc.Selection as TextSelection;
            if (selection == null)
                return;

            // Always get the word under the cursor, regardless of selection
            int originalLine = selection.ActivePoint.Line;
            int originalColumn = selection.ActivePoint.DisplayColumn;
            selection.WordLeft(true);
            string leftWord = selection.Text;
            selection.WordRight(true);
            string word = leftWord + selection.Text;
            // Restore cursor
            selection.MoveToLineAndOffset(originalLine, originalColumn);
            string typeName = word.Trim();

            if (string.IsNullOrWhiteSpace(typeName))
                return;

            // Search for 'message XXX' or 'enum XXX' in the document
            var textDoc = doc.Object("TextDocument") as TextDocument;
            EditPoint startPoint = textDoc.StartPoint.CreateEditPoint();
            string allText = startPoint.GetText(textDoc.EndPoint);
            var regex = new Regex(@"^\s*(message|enum)\s+" + Regex.Escape(typeName) + @"\b", RegexOptions.Multiline);
            var match = regex.Match(allText);
            if (match.Success)
            {
                int charIndex = match.Index;
                int line = 1;
                for (int i = 0; i < charIndex; i++)
                {
                    if (allText[i] == '\n')
                    {
                        line++;
                    }
                }
                // Always move to the first column of the matched line
                EditPoint defPoint = textDoc.StartPoint.CreateEditPoint();
                defPoint.MoveToLineAndOffset(line + 1, 1);
                selection.MoveToPoint(defPoint, false);
                doc.Activate();
            }
		}
	}
}

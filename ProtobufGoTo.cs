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

		public static void Initialize(Package package)
		{
			Instance = new ProtobufGoTo(package);
		}

        private void GoToLocation(DTE2 dte, string path, int charIndex)
        {
            Window win = dte.ItemOperations.OpenFile(path);
            var doc = win.Document;
            if (doc == null) return;
            var textDoc = doc.Object("TextDocument") as TextDocument;
            if (textDoc == null) return;
            var selection = doc.Selection as TextSelection;
            if (selection == null) return;

            EditPoint defPoint = textDoc.StartPoint.CreateEditPoint();
            defPoint.MoveToAbsoluteOffset(charIndex + 1); // MoveToAbsoluteOffset is 1-based
            selection.MoveToPoint(defPoint, false);
            doc.Activate();
        }

        private void MenuItemCallback(object sender, EventArgs e)
		{
            ProtobufGoToPackage ProtoPackage = (ProtobufGoToPackage)this.package;
            var dte = ProtoPackage.m_dte;
            if (dte == null)
            {
                dte = ServiceProvider.GetService(typeof(DTE)) as DTE2;
                if (dte == null)
                    return;
                ProtoPackage.m_dte = dte;
            }
            var doc = dte.ActiveDocument;
            if (doc == null)
                return;

            if (doc.Name.EndsWith(".proto", StringComparison.OrdinalIgnoreCase))
            {
                TextSelection selection = doc.Selection as TextSelection;
                if (selection == null)
                    return;

                selection.WordLeft(true);
                string leftWord = selection.Text;
                selection.WordRight(true);
                string word = leftWord + selection.Text;
                selection.Cancel();
                string typeName = word.Trim();

                if (string.IsNullOrWhiteSpace(typeName))
                    return;

                var textDoc = doc.Object("TextDocument") as TextDocument;
                string allText = (doc.Object("TextDocument") as TextDocument).StartPoint.CreateEditPoint().GetText((doc.Object("TextDocument") as TextDocument).EndPoint);
                var regex = new Regex(@"^\s*(message|enum)\s+" + Regex.Escape(typeName) + @"\b", RegexOptions.Multiline);
                var match = regex.Match(allText);
                if (match.Success)
                {
                    GoToLocation(dte, doc.FullName, match.Index);
                    return;
                }

                var importRegex = new Regex(@"^\s*import\s+""([^""]+)"";", RegexOptions.Multiline);
                var importMatches = importRegex.Matches(allText);
                string currentDir = Path.GetDirectoryName(doc.FullName);
                foreach (Match importMatch in importMatches)
                {
                    string importPath = importMatch.Groups[1].Value;
                    string fullImportPath = Path.Combine(currentDir, importPath);
                    if (!File.Exists(fullImportPath))
                        continue;
                    string importText = File.ReadAllText(fullImportPath);
                    var importTypeMatch = regex.Match(importText);
                    if (importTypeMatch.Success)
                    {
                        GoToLocation(dte, fullImportPath, importTypeMatch.Index);
                        return;
                    }
                }

                var solution = dte.Solution;
                var protoFiles = new System.Collections.Generic.List<string>();
                void FindProtoFiles(ProjectItems items)
                {
                    if (items == null) return;
                    foreach (ProjectItem item in items)
                    {
                        try
                        {
                            if ((item.Kind == EnvDTE.Constants.vsProjectItemKindPhysicalFile || item.Kind == EnvDTE.Constants.vsProjectItemKindMisc) &&
                                item.Name.EndsWith(".proto", StringComparison.OrdinalIgnoreCase))
                            {
                                protoFiles.Add(item.FileNames[1]);
                            }
                            if (item.ProjectItems != null)
                                FindProtoFiles(item.ProjectItems);
                        }
                        catch { }
                    }
                }
                foreach (Project proj in solution.Projects)
                {
                    try
                    {
                        if (proj.ProjectItems != null)
                            FindProtoFiles(proj.ProjectItems);
                    }
                    catch { }
                }

                foreach (var protoPath in protoFiles)
                {
                    if (!File.Exists(protoPath))
                        continue;
                    string allText2 = File.ReadAllText(protoPath);
                    var match2 = regex.Match(allText2);
                    if (match2.Success)
                    {
                        GoToLocation(dte, protoPath, match2.Index);
                        return;
                    }
                }
            }
            else if (doc.Name.EndsWith(".h", StringComparison.OrdinalIgnoreCase) ||
                doc.Name.EndsWith(".cpp", StringComparison.OrdinalIgnoreCase))
            {
                TextSelection selection = doc.Selection as TextSelection;
                if (selection == null)
                    return;

                selection.WordLeft(true);
                string leftWord = selection.Text;
                selection.WordRight(true);
                string word = leftWord + selection.Text;
                selection.Cancel();
                string typeName = word.Trim();

                if (string.IsNullOrWhiteSpace(typeName))
                    return;

                if (typeName.StartsWith("PacketTypeReq_", StringComparison.OrdinalIgnoreCase) ||
                    typeName.StartsWith("PacketTypeRes_", StringComparison.OrdinalIgnoreCase))
                {
                    typeName = typeName.Replace("PacketTypeReq_", "").Replace("PacketTypeRes_", "");
                }

                var solution = dte.Solution;
                var protoFiles = new System.Collections.Generic.List<string>();
                void FindProtoFiles(ProjectItems items)
                {
                    if (items == null) return;
                    foreach (ProjectItem item in items)
                    {
                        try
                        {
                            if ((item.Kind == EnvDTE.Constants.vsProjectItemKindPhysicalFile || item.Kind == EnvDTE.Constants.vsProjectItemKindMisc) &&
                                item.Name.EndsWith(".proto", StringComparison.OrdinalIgnoreCase))
                            {
                                protoFiles.Add(item.FileNames[1]);
                            }
                            if (item.ProjectItems != null)
                                FindProtoFiles(item.ProjectItems);
                        }
                        catch { }
                    }
                }
                foreach (Project proj in solution.Projects)
                {
                    try
                    {
                        if (proj.ProjectItems != null)
                            FindProtoFiles(proj.ProjectItems);
                    }
                    catch { }
                }

                var regex = new Regex(@"^\s*(message|enum)\s+" + Regex.Escape(typeName) + @"\b", RegexOptions.Multiline);
                foreach (var protoPath in protoFiles)
                {
                    if (!File.Exists(protoPath))
                        continue;
                    string allText = File.ReadAllText(protoPath);
                    var match = regex.Match(allText);
                    if (match.Success)
                    {
                        GoToLocation(dte, protoPath, match.Index);
                        return;
                    }
                }

                var enumBlockRegex = new Regex(@"\benum\s+\w+\s*\{[\s\S]*?\}", RegexOptions.Multiline);
                var valueRegex = new Regex(@"\b" + Regex.Escape(typeName) + @"\b");

                foreach (var protoPath in protoFiles)
                {
                    if (!File.Exists(protoPath))
                        continue;

                    string allText = File.ReadAllText(protoPath);

                    foreach (Match enumBlockMatch in enumBlockRegex.Matches(allText))
                    {
                        var valueMatch = valueRegex.Match(enumBlockMatch.Value);
                        if (valueMatch.Success)
                        {
                            int charIndex = enumBlockMatch.Index + valueMatch.Index;
                            GoToLocation(dte, protoPath, charIndex);
                            return;
                        }
                    }
                }
            }
        }
	}
}

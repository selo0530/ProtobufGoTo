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
                    // Find the column offset of the typename in the matched line by analyzing the line text
                    EditPoint defPoint = textDoc.StartPoint.CreateEditPoint();
                    defPoint.MoveToLineAndOffset(line + 1, 1);
                    string lineText = defPoint.GetLines(line + 1, line + 2);
                    int columnOffset = lineText.IndexOf(typeName, StringComparison.Ordinal);
                    if (columnOffset >= 0)
                    {
                        defPoint.MoveToLineAndOffset(line + 1, columnOffset + 1);
                    }
                    selection.MoveToPoint(defPoint, false);
                    doc.Activate();
                    return;
                }

                // If not found, search imported proto files
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
                        // Open the imported file in the editor
                        Window importWin = dte.ItemOperations.OpenFile(fullImportPath);
                        var importDoc = importWin.Document;
                        var importTextDoc = importDoc.Object("TextDocument") as TextDocument;
                        int charIndex = importTypeMatch.Index;
                        int line = 1;
                        for (int i = 0; i < charIndex; i++)
                        {
                            if (importText[i] == '\n')
                            {
                                line++;
                            }
                        }
                        EditPoint defPoint = importTextDoc.StartPoint.CreateEditPoint();
                        defPoint.MoveToLineAndOffset(line + 1, 1);
                        string lineText = defPoint.GetLines(line + 1, line + 2);
                        int columnOffset = lineText.IndexOf(typeName, StringComparison.Ordinal);
                        if (columnOffset >= 0)
                        {
                            defPoint.MoveToLineAndOffset(line + 1, columnOffset + 1);
                        }
                        var importSelection = importDoc.Selection as TextSelection;
                        importSelection.MoveToPoint(defPoint, false);
                        importDoc.Activate();
                        return;
                    }
                }
            }
            else if (doc.Name.EndsWith(".h", StringComparison.OrdinalIgnoreCase) ||
                doc.Name.EndsWith(".cpp", StringComparison.OrdinalIgnoreCase))
            {
                // 커서 위치의 단어 추출
                TextSelection selection = doc.Selection as TextSelection;
                if (selection == null)
                    return;
                int originalLine = selection.ActivePoint.Line;
                int originalColumn = selection.ActivePoint.DisplayColumn;
                selection.WordLeft(true);
                string leftWord = selection.Text;
                selection.WordRight(true);
                string word = leftWord + selection.Text;
                string typeName = word.Trim();
                if (string.IsNullOrWhiteSpace(typeName))
                    return;

                // 솔루션 내 모든 .proto 파일 탐색
                var solution = dte.Solution;
                var protoFiles = new System.Collections.Generic.List<string>();
                void FindProtoFiles(ProjectItems items)
                {
                    foreach (ProjectItem item in items)
                    {
                        try
                        {
                            if ((item.Kind == EnvDTE.Constants.vsProjectItemKindPhysicalFile || item.Kind == EnvDTE.Constants.vsProjectItemKindMisc) &&
                                item.Name.EndsWith(".proto", StringComparison.OrdinalIgnoreCase))
                            {
                                string filePath = item.FileNames[1];
                                protoFiles.Add(filePath);
                            }
                            if (item.ProjectItems != null && item.ProjectItems.Count > 0)
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

                // 각 .proto 파일에서 message/enum 정의 찾기
                var regex = new Regex(@"^\s*(message|enum)\s+" + Regex.Escape(typeName) + @"\b", RegexOptions.Multiline);
                foreach (var protoPath in protoFiles)
                {
                    if (!File.Exists(protoPath))
                        continue;
                    string allText = File.ReadAllText(protoPath);
                    var match = regex.Match(allText);
                    if (match.Success)
                    {
                        int charIndex = match.Index;
                        int line = 1;
                        for (int i = 0; i < charIndex; i++)
                        {
                            if (allText[i] == '\n')
                                line++;
                        }
                        Window protoWin = dte.ItemOperations.OpenFile(protoPath);
                        var protoDoc = protoWin.Document;
                        var protoTextDoc = protoDoc.Object("TextDocument") as TextDocument;
                        EditPoint defPoint = protoTextDoc.StartPoint.CreateEditPoint();
                        defPoint.MoveToLineAndOffset(line + 1, 1);
                        string lineText = defPoint.GetLines(line + 1, line + 2);
                        int columnOffset = lineText.IndexOf(typeName, StringComparison.Ordinal);
                        if (columnOffset >= 0)
                        {
                            defPoint.MoveToLineAndOffset(line + 1, columnOffset + 1);
                        }
                        var protoSelection = protoDoc.Selection as TextSelection;
                        protoSelection.MoveToPoint(defPoint, false);
                        protoDoc.Activate();
                        return;
                    }
                }
            }
        }
	}
}

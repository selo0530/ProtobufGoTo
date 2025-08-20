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
using System.Diagnostics;
using System.Collections.Generic;

namespace ProtobufGoTo
{
	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed class ProtobufGoTo
	{
		public const int CommandId = 0x0100;
        public const int CommandFuncId = 0x0101;

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

                menuCommandID = new CommandID(CommandSet, CommandFuncId);
                menuItem = new MenuCommand(this.MenuItemFuncCallback, menuCommandID);
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

        private static readonly string[] CppExtensions = { ".cpp" };

        private bool IsCppFile(string fileName)
        {
            return CppExtensions.Any(ext => fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase));
        }

        private List<string> FindAllCppFiles(DTE2 dte)
        {
            var cppFiles = new List<string>();
            var solution = dte.Solution;

            if (solution == null || string.IsNullOrEmpty(solution.FullName))
                return cppFiles;

            try
            {
                // Method 1: Use DTE ProjectItems (existing method)
                FindCppFilesFromDTE(solution, cppFiles);

                // Method 2: File system search as fallback
                var solutionDir = Path.GetDirectoryName(solution.FullName + "//");
                if (!string.IsNullOrEmpty(solutionDir))
                {
                    FindCppFilesFromFileSystem(solutionDir, cppFiles);
                }

                Debug.WriteLine($"Total C++ files found: {cppFiles.Count}");
                foreach (var file in cppFiles)
                {
                    Debug.WriteLine($"C++ file: {file}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error finding C++ files: {ex.Message}");
            }

            return cppFiles.Distinct().ToList();
        }

        private void FindCppFilesFromDTE(Solution solution, List<string> cppFiles)
        {
            try
            {
                void FindCppFiles(ProjectItems items, int depth = 0)
                {
                    if (depth > 10) // Prevent infinite recursion
                        return;

                    if (items == null) return;

                    try
                    {
                        Debug.WriteLine($"FindCppFiles - Depth: {depth}, Items count: {items.Count}");
                        
                        for (int i = 1; i <= items.Count; i++) // DTE collections are 1-based
                        {
                            try
                            {
                                var item = items.Item(i);
                                if (item == null) continue;

                                Debug.WriteLine($"Checking item: {item.Name}, Kind: {item.Kind}");

                                if (IsCppFile(item.Name))
                                {
                                    try
                                    {
                                        if (item.Kind == EnvDTE.Constants.vsProjectItemKindPhysicalFile || 
                                            item.Kind == EnvDTE.Constants.vsProjectItemKindMisc)
                                        {
                                            if (item.FileCount > 0)
                                            {
                                                string filePath = item.FileNames[1]; // DTE is 1-based
                                                if (!cppFiles.Contains(filePath))
                                                {
                                                    cppFiles.Add(filePath);
                                                    Debug.WriteLine($"Added C++ file: {filePath}");
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine($"Error accessing file properties: {ex.Message}");
                                    }
                                }

                                // Recursively search subdirectories
                                if (item.ProjectItems != null && item.ProjectItems.Count > 0)
                                {
                                    FindCppFiles(item.ProjectItems, depth + 1);
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"Error processing project item {i}: {ex.Message}");
                                continue;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error iterating project items: {ex.Message}");
                    }
                }

                foreach (Project proj in solution.Projects)
                {
                    try
                    {
                        if (proj == null) continue;
                        
                        Debug.WriteLine($"Processing project: {proj.Name} (Kind: {proj.Kind})");
                        
                        // Handle different project types
                        if (proj.Kind == EnvDTE.Constants.vsProjectKindSolutionItems ||
                            proj.Kind == EnvDTE.Constants.vsProjectKindMisc)
                        {
                            // Solution folders or misc items
                            if (proj.ProjectItems != null)
                                FindCppFiles(proj.ProjectItems);
                        }
                        else if (proj.ProjectItems != null)
                        {
                            // Regular projects
                            FindCppFiles(proj.ProjectItems);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error processing project {proj?.Name}: {ex.Message}");
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in FindCppFilesFromDTE: {ex.Message}");
            }
        }

        private void FindCppFilesFromFileSystem(string solutionDir, List<string> cppFiles)
        {
            try
            {
                Debug.WriteLine($"Searching C++ files in file system: {solutionDir}");
                
                var allCppFiles = new List<string>();
                
                // Search for each C++ extension
                foreach (var extension in CppExtensions)
                {
                    var files = Directory.GetFiles(solutionDir, "*" + extension, SearchOption.AllDirectories);
                    allCppFiles.AddRange(files);
                }
                
                foreach (var file in allCppFiles)
                {
                    if (!cppFiles.Contains(file))
                    {
                        cppFiles.Add(file);
                        Debug.WriteLine($"Added C++ file from filesystem: {file}");
                    }
                }
                
                Debug.WriteLine($"File system search found {allCppFiles.Count} C++ files");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in C++ file system search: {ex.Message}");
            }
        }

        private List<string> FindAllProtoFiles(DTE2 dte)
        {
            var protoFiles = new List<string>();
            var solution = dte.Solution;

            if (solution == null || string.IsNullOrEmpty(solution.FullName))
                return protoFiles;

            try
            {
                // Method 1: Use DTE ProjectItems (existing method)
                FindProtoFilesFromDTE(solution, protoFiles);

                // Method 2: File system search as fallback
                var solutionDir = Path.GetDirectoryName(solution.FullName + "//");
                if (!string.IsNullOrEmpty(solutionDir))
                {
                    FindProtoFilesFromFileSystem(solutionDir, protoFiles);
                }

                Debug.WriteLine($"Total proto files found: {protoFiles.Count}");
                foreach (var file in protoFiles)
                {
                    Debug.WriteLine($"Proto file: {file}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error finding proto files: {ex.Message}");
            }

            return protoFiles.Distinct().ToList();
        }

        private void FindProtoFilesFromDTE(Solution solution, List<string> protoFiles)
        {
            try
            {
                void FindProtoFiles(ProjectItems items, int depth = 0)
                {
                    if (depth > 10) // Prevent infinite recursion
                        return;

                    if (items == null) return;

                    try
                    {
                        Debug.WriteLine($"FindProtoFiles - Depth: {depth}, Items count: {items.Count}");
                        
                        for (int i = 1; i <= items.Count; i++) // DTE collections are 1-based
                        {
                            try
                            {
                                var item = items.Item(i);
                                if (item == null) continue;

                                Debug.WriteLine($"Checking item: {item.Name}, Kind: {item.Kind}");

                                if (item.Name.EndsWith(".proto", StringComparison.OrdinalIgnoreCase))
                                {
                                    try
                                    {
                                        if (item.Kind == EnvDTE.Constants.vsProjectItemKindPhysicalFile || 
                                            item.Kind == EnvDTE.Constants.vsProjectItemKindMisc)
                                        {
                                            if (item.FileCount > 0)
                                            {
                                                string filePath = item.FileNames[1]; // DTE is 1-based
                                                if (!protoFiles.Contains(filePath))
                                                {
                                                    protoFiles.Add(filePath);
                                                    Debug.WriteLine($"Added proto file: {filePath}");
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine($"Error accessing file properties: {ex.Message}");
                                    }
                                }

                                // Recursively search subdirectories
                                if (item.ProjectItems != null && item.ProjectItems.Count > 0)
                                {
                                    FindProtoFiles(item.ProjectItems, depth + 1);
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"Error processing project item {i}: {ex.Message}");
                                continue;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error iterating project items: {ex.Message}");
                    }
                }

                foreach (Project proj in solution.Projects)
                {
                    try
                    {
                        if (proj == null) continue;
                        
                        Debug.WriteLine($"Processing project: {proj.Name} (Kind: {proj.Kind})");
                        
                        // Handle different project types
                        if (proj.Kind == EnvDTE.Constants.vsProjectKindSolutionItems ||
                            proj.Kind == EnvDTE.Constants.vsProjectKindMisc)
                        {
                            // Solution folders or misc items
                            if (proj.ProjectItems != null)
                                FindProtoFiles(proj.ProjectItems);
                        }
                        else if (proj.ProjectItems != null)
                        {
                            // Regular projects
                            FindProtoFiles(proj.ProjectItems);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error processing project {proj?.Name}: {ex.Message}");
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in FindProtoFilesFromDTE: {ex.Message}");
            }
        }

        private void FindProtoFilesFromFileSystem(string solutionDir, List<string> protoFiles)
        {
            try
            {
                Debug.WriteLine($"Searching file system in: {solutionDir}");
                
                var allProtoFiles = Directory.GetFiles(solutionDir, "*.proto", SearchOption.AllDirectories);
                
                foreach (var file in allProtoFiles)
                {
                    if (!protoFiles.Contains(file))
                    {
                        protoFiles.Add(file);
                        Debug.WriteLine($"Added proto file from filesystem: {file}");
                    }
                }
                
                Debug.WriteLine($"File system search found {allProtoFiles.Length} proto files");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in file system search: {ex.Message}");
            }
        }

        private void MenuItemFuncCallback(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

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

            if (typeName.StartsWith("PacketTypeReq_", StringComparison.OrdinalIgnoreCase) ||
                typeName.StartsWith("PacketTypeRes_", StringComparison.OrdinalIgnoreCase))
            {
                typeName = typeName.Replace("PacketTypeReq_", "").Replace("PacketTypeRes_", "");
            }

            var cppFiles = FindAllCppFiles(dte);
            var funcName = "ON_" + typeName;

            var regex = new Regex(@"\s+" + Regex.Escape(funcName) + @"\s*\(", RegexOptions.Multiline);
            foreach(var cppPath in cppFiles)
            {
                if (!File.Exists(cppPath))
                    continue;
                string allText = File.ReadAllText(cppPath);
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
                    Window protoWin = dte.ItemOperations.OpenFile(cppPath);
                    var protoDoc = protoWin.Document;
                    var protoTextDoc = protoDoc.Object("TextDocument") as TextDocument;
                    EditPoint defPoint = protoTextDoc.StartPoint.CreateEditPoint();
                    defPoint.MoveToLineAndOffset(line, 1);
                    string lineText = defPoint.GetLines(line, line + 2);
                    int columnOffset = lineText.IndexOf(funcName, StringComparison.Ordinal);
                    if (columnOffset >= 0)
                    {
                        defPoint.MoveToLineAndOffset(line, columnOffset + 1);
                    }
                    var protoSelection = protoDoc.Selection as TextSelection;
                    protoSelection.MoveToPoint(defPoint, false);
                    protoDoc.Activate();
                    return;
                }
            }
        }


        private void MenuItemCallback(object sender, EventArgs e)
		{
            ThreadHelper.ThrowIfNotOnUIThread();

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
                var regex = new Regex(@"(message|enum)\s+" + Regex.Escape(typeName) + @"\b", RegexOptions.Multiline);
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
                    defPoint.MoveToLineAndOffset(line, 1);
                    string lineText = defPoint.GetLines(line, line + 2);
                    int columnOffset = lineText.IndexOf(typeName, StringComparison.Ordinal);
                    if (columnOffset >= 0)
                    {
                        defPoint.MoveToLineAndOffset(line, columnOffset + 1);
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
                        defPoint.MoveToLineAndOffset(line, 1);
                        string lineText = defPoint.GetLines(line, line + 2);
                        int columnOffset = lineText.IndexOf(typeName, StringComparison.Ordinal);
                        if (columnOffset >= 0)
                        {
                            defPoint.MoveToLineAndOffset(line, columnOffset + 1);
                        }
                        var importSelection = importDoc.Selection as TextSelection;
                        importSelection.MoveToPoint(defPoint, false);
                        importDoc.Activate();
                        return;
                    }
                }

                // If not found, search proto files from the solution using improved method
                var protoFiles = FindAllProtoFiles(dte);

                // 각 .proto 파일에서 message/enum 정의 찾기
                var regex2 = new Regex(@"(message|enum)\s+" + Regex.Escape(typeName) + @"\b", RegexOptions.Multiline);
                foreach (var protoPath in protoFiles)
                {
                    if (!File.Exists(protoPath))
                        continue;
                    string allText2 = File.ReadAllText(protoPath);
                    var match2 = regex2.Match(allText2);
                    if (match2.Success)
                    {
                        int charIndex2 = match2.Index;
                        int line2 = 1;
                        for (int i = 0; i < charIndex2; i++)
                        {
                            if (allText2[i] == '\n')
                                line2++;
                        }
                        Window protoWin2 = dte.ItemOperations.OpenFile(protoPath);
                        var protoDoc2 = protoWin2.Document;
                        var protoTextDoc2 = protoDoc2.Object("TextDocument") as TextDocument;
                        EditPoint defPoint2 = protoTextDoc2.StartPoint.CreateEditPoint();
                        defPoint2.MoveToLineAndOffset(line2, 1);
                        string lineText2 = defPoint2.GetLines(line2, line2 + 2);
                        int columnOffset2 = lineText2.IndexOf(typeName, StringComparison.Ordinal);
                        if (columnOffset2 >= 0)
                        {
                            defPoint2.MoveToLineAndOffset(line2, columnOffset2 + 1);
                        }
                        var protoSelection2 = protoDoc2.Selection as TextSelection;
                        protoSelection2.MoveToPoint(defPoint2, false);
                        protoDoc2.Activate();
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
                // Restore cursor
                selection.MoveToLineAndOffset(originalLine, originalColumn);
                string typeName = word.Trim();
                if (string.IsNullOrWhiteSpace(typeName))
                    return;

                if (typeName.StartsWith("PacketTypeReq_", StringComparison.OrdinalIgnoreCase) ||
                    typeName.StartsWith("PacketTypeRes_", StringComparison.OrdinalIgnoreCase))
                {
                    typeName = typeName.Replace("PacketTypeReq_", "").Replace("PacketTypeRes_", "");
                }

                // 솔루션 내 모든 .proto 파일 탐색 - 개선된 방법 사용
                var protoFiles = FindAllProtoFiles(dte);

                // 각 .proto 파일에서 message/enum 정의 찾기
                //var regex = new Regex(@"^\s*(message|enum)\s+" + Regex.Escape(typeName) + @"\b", RegexOptions.Multiline);
                var regex = new Regex(@"(message|enum)\s+" + Regex.Escape(typeName) + @"\b", RegexOptions.Multiline);
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
                        defPoint.MoveToLineAndOffset(line, 1);
                        string lineText = defPoint.GetLines(line, line + 2);
                        int columnOffset = lineText.IndexOf(typeName, StringComparison.Ordinal);
                        if (columnOffset >= 0)
                        {
                            defPoint.MoveToLineAndOffset(line, columnOffset + 1);
                        }
                        var protoSelection = protoDoc.Selection as TextSelection;
                        protoSelection.MoveToPoint(defPoint, false);
                        protoDoc.Activate();
                        return;
                    }
                }

                // 각 .proto 파일에서 enum 값 찾기
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
                            defPoint.MoveToLineAndOffset(line, 1);
                            string lineText = defPoint.GetLines(line, line + 2);
                            int columnOffset = lineText.IndexOf(typeName, StringComparison.Ordinal);
                            if (columnOffset >= 0)
                            {
                                defPoint.MoveToLineAndOffset(line, columnOffset + 1);
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
}

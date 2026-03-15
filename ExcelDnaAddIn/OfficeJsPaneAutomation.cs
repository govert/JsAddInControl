using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using ExcelDna.Integration;

namespace Testing.ExcelDnaAddIn
{
    internal static class OfficeJsPaneAutomation
    {
        // These are the primary strings to change when adapting the automation to a real add-in.
        private const string HomeRibbonTabName = "Home";
        private const string DnaRibbonTabName = "Testing DNA";
        private const string JsShowRibbonButtonName = "Show Smiley Pane";
        private const string TaskPaneCloseButtonName = "Close pane";
        private const string TaskPaneOptionsButtonName = "Task Pane Options";
        private const string TaskPaneWindowClassName = "MsoWorkPane";
        private const string TaskPaneCloseButtonClassName = "NetUISimpleButton";
        private static readonly string[] TaskPaneAnchorTexts =
        {
            "Smiling test pane",
            "This custom task pane is coming from the local Office JavaScript add-in."
        };

        private static readonly string[] TaskPaneTitleCandidates =
        {
            "Testing JS Test Add-in",
            "Show Smiley Pane"
        };

        private const int SearchTimeoutMs = 10000;
        private const int PollIntervalMs = 200;
        private const int MaxDumpNodes = 1500;
        private const int MaxDumpDepth = 7;

        public static void ShowPane()
        {
            var root = GetExcelRootElement();

            SelectRibbonTab(root, HomeRibbonTabName);
            var button = WaitForElement(
                () => FindRibbonCommand(root, JsShowRibbonButtonName),
                $"Could not find ribbon button '{JsShowRibbonButtonName}' on the '{HomeRibbonTabName}' tab.");

            Invoke(button);
            QueueOnMainThread(() =>
            {
                try
                {
                    var refreshedRoot = GetExcelRootElement();
                    SelectRibbonTab(refreshedRoot, DnaRibbonTabName);
                }
                catch
                {
                    // Best effort only; the pane show command is the important action.
                }
            });
        }

        public static void HidePane()
        {
            var root = GetExcelRootElement();
            var closeButton = FindTaskPaneCloseButton(root);
            if (closeButton != null)
            {
                Invoke(closeButton);
                return;
            }

            throw new InvalidOperationException(
                "Could not find the Office JS task pane close button. Run 'Dump UIA Tree' to inspect the current Excel UI Automation tree.");
        }

        public static string DumpTreeToTextFile()
        {
            var root = GetExcelRootElement();
            var outputPath = Path.Combine(Path.GetTempPath(), "Testing.ExcelDna.UiAutomation.txt");
            var lines = new List<string>
            {
                $"Timestamp: {DateTime.Now:O}",
                $"ProcessId: {Process.GetCurrentProcess().Id}",
                $"MainWindowTitle: {Process.GetCurrentProcess().MainWindowTitle}",
                string.Empty
            };

            var nodes = new List<UiNodeRecord>();
            CollectNodes(root, 0, nodes);
            lines.AddRange(nodes.Where(IsRelevantNode).Select(FormatNode));
            File.WriteAllLines(outputPath, lines, Encoding.UTF8);
            return outputPath;
        }

        public static void DumpTreeToWorksheet(string outputPath)
        {
            var root = GetExcelRootElement();
            var nodes = new List<UiNodeRecord>();
            CollectNodes(root, 0, nodes);

            dynamic application = ExcelDnaUtil.Application;
            dynamic workbook = application.Workbooks.Add();
            dynamic worksheet = workbook.Worksheets.Add();
            worksheet.Name = $"UIA Dump {DateTime.Now:HHmmss}";

            worksheet.Cells[1, 1].Value2 = "Depth";
            worksheet.Cells[1, 2].Value2 = "ControlType";
            worksheet.Cells[1, 3].Value2 = "Name";
            worksheet.Cells[1, 4].Value2 = "AutomationId";
            worksheet.Cells[1, 5].Value2 = "ClassName";
            worksheet.Cells[1, 6].Value2 = "Bounds";
            worksheet.Cells[1, 7].Value2 = "Offscreen";
            worksheet.Cells[1, 8].Value2 = "Source";

            var row = 2;
            foreach (var node in nodes.Where(IsRelevantNode))
            {
                worksheet.Cells[row, 1].Value2 = node.Depth;
                worksheet.Cells[row, 2].Value2 = node.ControlType;
                worksheet.Cells[row, 3].Value2 = node.Name;
                worksheet.Cells[row, 4].Value2 = node.AutomationId;
                worksheet.Cells[row, 5].Value2 = node.ClassName;
                worksheet.Cells[row, 6].Value2 = node.Bounds;
                worksheet.Cells[row, 7].Value2 = node.IsOffscreen ? "true" : "false";
                worksheet.Cells[row, 8].Value2 = row == 2 ? outputPath : string.Empty;
                row++;
            }

            worksheet.Columns.AutoFit();
        }

        private static AutomationElement GetExcelRootElement()
        {
            dynamic application = ExcelDnaUtil.Application;
            var hwnd = (IntPtr)(int)application.Hwnd;
            var root = AutomationElement.FromHandle(hwnd);
            if (root == null)
            {
                throw new InvalidOperationException("Could not get the Excel main window for UI Automation.");
            }

            return root;
        }

        private static AutomationElement FindTaskPaneAnchor(AutomationElement root)
        {
            foreach (var anchorText in TaskPaneAnchorTexts)
            {
                var anchor = root.FindFirst(
                    TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.NameProperty, anchorText));

                if (anchor != null)
                {
                    return anchor;
                }
            }

            foreach (var title in TaskPaneTitleCandidates)
            {
                var titleElement = root.FindFirst(
                    TreeScope.Descendants,
                    new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, title),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)));

                if (titleElement != null)
                {
                    return titleElement;
                }
            }

            return null;
        }

        private static void SelectRibbonTab(AutomationElement root, string tabName)
        {
            var tab = WaitForElement(
                () => root.FindFirst(
                    TreeScope.Descendants,
                    new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
                        new PropertyCondition(AutomationElement.NameProperty, tabName))),
                $"Could not find ribbon tab '{tabName}'.");

            if (tab.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selectionPattern))
            {
                ((SelectionItemPattern)selectionPattern).Select();
                Thread.Sleep(250);
                return;
            }

            Invoke(tab);
            Thread.Sleep(250);
        }

        private static AutomationElement FindRibbonCommand(AutomationElement root, string exactName)
        {
            var exactButton = root.FindFirst(
                TreeScope.Descendants,
                new AndCondition(
                    new OrCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)),
                    new PropertyCondition(AutomationElement.NameProperty, exactName)));

            if (exactButton != null)
            {
                return exactButton;
            }

            return FindInRawView(
                root,
                element =>
                {
                    var type = SafeControlType(element);
                    if (!type.EndsWith("Button", StringComparison.OrdinalIgnoreCase) &&
                        !type.EndsWith("MenuItem", StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }

                    var name = SafeName(element);
                    return name.Equals(exactName, StringComparison.OrdinalIgnoreCase) ||
                           name.IndexOf("Smiley", StringComparison.OrdinalIgnoreCase) >= 0;
                },
                maxDepth: 12);
        }

        private static AutomationElement FindTaskPaneCloseButton(AutomationElement root)
        {
            var paneWindow = FindInRawView(
                root,
                element =>
                    SafeControlType(element).EndsWith("Window", StringComparison.OrdinalIgnoreCase) &&
                    SafeName(element).Equals("Testing JS Test Add-in", StringComparison.OrdinalIgnoreCase) &&
                    SafeCurrentProperty(element, AutomationElement.ClassNameProperty).Equals(TaskPaneWindowClassName, StringComparison.OrdinalIgnoreCase),
                maxDepth: 12);

            if (paneWindow != null)
            {
                var closeButton = FindInRawView(
                    paneWindow,
                    element =>
                        SafeControlType(element).EndsWith("Button", StringComparison.OrdinalIgnoreCase) &&
                        SafeName(element).Equals(TaskPaneCloseButtonName, StringComparison.OrdinalIgnoreCase) &&
                        SafeCurrentProperty(element, AutomationElement.ClassNameProperty).Equals(TaskPaneCloseButtonClassName, StringComparison.OrdinalIgnoreCase),
                    maxDepth: 8);

                if (closeButton != null)
                {
                    return closeButton;
                }
            }

            // Fallback: search globally for the specific task-pane close button exposed by the JS pane chrome.
            return FindInRawView(
                root,
                element =>
                    SafeControlType(element).EndsWith("Button", StringComparison.OrdinalIgnoreCase) &&
                    SafeName(element).Equals(TaskPaneCloseButtonName, StringComparison.OrdinalIgnoreCase) &&
                    SafeCurrentProperty(element, AutomationElement.ClassNameProperty).Equals(TaskPaneCloseButtonClassName, StringComparison.OrdinalIgnoreCase),
                maxDepth: 14);
        }

        private static AutomationElement FindButtonByName(AutomationElement root, string name)
        {
            return root.FindFirst(
                TreeScope.Descendants,
                new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                    new PropertyCondition(AutomationElement.NameProperty, name)));
        }

        private static AutomationElement FindDescendantButton(AutomationElement root, string name)
        {
            return root.FindFirst(
                TreeScope.Descendants,
                new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                    new PropertyCondition(AutomationElement.NameProperty, name)));
        }

        private static AutomationElement WaitForElement(Func<AutomationElement> resolver, string errorMessage)
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(SearchTimeoutMs);
            while (DateTime.UtcNow < deadline)
            {
                var element = resolver();
                if (element != null)
                {
                    return element;
                }

                Thread.Sleep(PollIntervalMs);
            }

            throw new InvalidOperationException(errorMessage);
        }

        private static void QueueOnMainThread(Action action)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    action();
                }
                catch
                {
                    // Best effort only for deferred UI cleanup steps.
                }
            });
        }

        private static AutomationElement FindInRawView(AutomationElement root, Func<AutomationElement, bool> predicate, int maxDepth)
        {
            var walker = TreeWalker.RawViewWalker;
            return FindInRawView(root, predicate, walker, 0, maxDepth);
        }

        private static AutomationElement FindInRawView(
            AutomationElement element,
            Func<AutomationElement, bool> predicate,
            TreeWalker walker,
            int depth,
            int maxDepth)
        {
            if (element == null || depth > maxDepth)
            {
                return null;
            }

            if (predicate(element))
            {
                return element;
            }

            var child = walker.GetFirstChild(element);
            while (child != null)
            {
                var match = FindInRawView(child, predicate, walker, depth + 1, maxDepth);
                if (match != null)
                {
                    return match;
                }

                child = walker.GetNextSibling(child);
            }

            return null;
        }

        private static void Invoke(AutomationElement element)
        {
            if (element.TryGetCurrentPattern(InvokePattern.Pattern, out var pattern))
            {
                ((InvokePattern)pattern).Invoke();
                return;
            }

            throw new InvalidOperationException(
                $"Element '{SafeName(element)}' was found but does not support InvokePattern.");
        }

        private static void CollectNodes(AutomationElement element, int depth, List<UiNodeRecord> nodes)
        {
            if (element == null || depth > MaxDumpDepth || nodes.Count >= MaxDumpNodes)
            {
                return;
            }

            nodes.Add(new UiNodeRecord
            {
                Depth = depth,
                ControlType = SafeControlType(element),
                Name = SafeName(element),
                AutomationId = SafeCurrentProperty(element, AutomationElement.AutomationIdProperty),
                ClassName = SafeCurrentProperty(element, AutomationElement.ClassNameProperty),
                Bounds = FormatBounds(SafeBoundingRectangle(element)),
                IsOffscreen = SafeIsOffscreen(element)
            });

            var walker = TreeWalker.ControlViewWalker;
            var child = walker.GetFirstChild(element);
            while (child != null && nodes.Count < MaxDumpNodes)
            {
                CollectNodes(child, depth + 1, nodes);
                child = walker.GetNextSibling(child);
            }
        }

        private static string FormatNode(UiNodeRecord node)
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0}- {1} | Name='{2}' | AutomationId='{3}' | Class='{4}' | Bounds={5} | Offscreen={6}",
                new string(' ', node.Depth * 2),
                node.ControlType,
                node.Name,
                node.AutomationId,
                node.ClassName,
                node.Bounds,
                node.IsOffscreen);
        }

        private static bool IsRelevantNode(UiNodeRecord node)
        {
            var name = node.Name ?? string.Empty;
            var className = node.ClassName ?? string.Empty;
            var controlType = node.ControlType ?? string.Empty;

            if (name.IndexOf("Testing", StringComparison.OrdinalIgnoreCase) >= 0 ||
                name.IndexOf("Smiley", StringComparison.OrdinalIgnoreCase) >= 0 ||
                name.IndexOf("Close pane", StringComparison.OrdinalIgnoreCase) >= 0 ||
                name.IndexOf("Task Pane Options", StringComparison.OrdinalIgnoreCase) >= 0 ||
                name.IndexOf("Home", StringComparison.OrdinalIgnoreCase) >= 0 ||
                name.IndexOf("pane", StringComparison.OrdinalIgnoreCase) >= 0 ||
                name.IndexOf("task", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if ((className.IndexOf("NetUI", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 className.IndexOf("MsoCommandBar", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 className.IndexOf("MsoWorkPane", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 className.IndexOf("NUIPane", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 className.IndexOf("WebView", StringComparison.OrdinalIgnoreCase) >= 0) &&
                name.IndexOf("Status Bar", StringComparison.OrdinalIgnoreCase) < 0 &&
                name.IndexOf("Accessibility", StringComparison.OrdinalIgnoreCase) < 0 &&
                name.IndexOf("Average", StringComparison.OrdinalIgnoreCase) < 0 &&
                name.IndexOf("Count", StringComparison.OrdinalIgnoreCase) < 0 &&
                name.IndexOf("Sum", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return true;
            }

            // Keep only ribbon/task-pane-shaped nodes, and skip worksheet/grid content.
            if ((controlType.EndsWith("Button", StringComparison.OrdinalIgnoreCase) ||
                 controlType.EndsWith("TabItem", StringComparison.OrdinalIgnoreCase) ||
                 controlType.EndsWith("Pane", StringComparison.OrdinalIgnoreCase) ||
                 controlType.EndsWith("Text", StringComparison.OrdinalIgnoreCase)) &&
                className.IndexOf("EXCEL7", StringComparison.OrdinalIgnoreCase) < 0 &&
                className.IndexOf("ExcelGrid", StringComparison.OrdinalIgnoreCase) < 0 &&
                className.IndexOf("SheetTab", StringComparison.OrdinalIgnoreCase) < 0 &&
                className.IndexOf("XLDESK", StringComparison.OrdinalIgnoreCase) < 0 &&
                className.IndexOf("XLCTL", StringComparison.OrdinalIgnoreCase) < 0 &&
                className.IndexOf("EDTBX", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return node.Depth <= 8;
            }

            return false;
        }

        private static string SafeName(AutomationElement element)
        {
            return SafeCurrentProperty(element, AutomationElement.NameProperty);
        }

        private static string SafeControlType(AutomationElement element)
        {
            try
            {
                return element.Current.ControlType?.ProgrammaticName ?? "<null>";
            }
            catch
            {
                return "<error>";
            }
        }

        private static string SafeCurrentProperty(AutomationElement element, AutomationProperty property)
        {
            try
            {
                var value = element.GetCurrentPropertyValue(property, true);
                return value == AutomationElement.NotSupported ? string.Empty : value?.ToString() ?? string.Empty;
            }
            catch
            {
                return "<error>";
            }
        }

        private static bool SafeIsOffscreen(AutomationElement element)
        {
            try
            {
                return element.Current.IsOffscreen;
            }
            catch
            {
                return false;
            }
        }

        private static System.Windows.Rect SafeBoundingRectangle(AutomationElement element)
        {
            try
            {
                return element.Current.BoundingRectangle;
            }
            catch
            {
                return System.Windows.Rect.Empty;
            }
        }

        private static string FormatBounds(System.Windows.Rect rect)
        {
            return rect == System.Windows.Rect.Empty
                ? string.Empty
                : $"{rect.Left:0},{rect.Top:0},{rect.Width:0},{rect.Height:0}";
        }

        private sealed class UiNodeRecord
        {
            public int Depth { get; set; }
            public string ControlType { get; set; }
            public string Name { get; set; }
            public string AutomationId { get; set; }
            public string ClassName { get; set; }
            public string Bounds { get; set; }
            public bool IsOffscreen { get; set; }
        }
    }
}

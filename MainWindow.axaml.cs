using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;

namespace SignatureGenerator
{
    public partial class MainWindow : Window
    {
        private readonly string _appDataDir;
        private readonly string _templatesDir;
        private readonly string _dataDir;

        private string? _htmlTemplatePath;
        private string? _csvPath;
        private List<string> _templatePlaceholders = new();

        // Parsed CSV kept in memory after validation
        private CsvTable? _csvTable;

        private enum ProcessState
        {
            MissingHtml,        // No HTML uploaded
            HtmlOnly,           // HTML uploaded, CSV not uploaded
            HtmlAndCsvInvalid,  // CSV uploaded but headers/rows invalid
            HtmlAndCsvValid     // CSV valid and matches placeholders
        }

        private ProcessState _state = ProcessState.MissingHtml;

        public MainWindow()
        {
            InitializeComponent();

            _appDataDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "SignatureGenerator");

            _templatesDir = Path.Combine(_appDataDir, "Templates");
            _dataDir      = Path.Combine(_appDataDir, "Data");

            Directory.CreateDirectory(_templatesDir);
            Directory.CreateDirectory(_dataDir);

            UpdateUiState();
        }

        // ========================= Upload HTML =========================
        private async void ChooseHtmlButton_OnClick(object? sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Title = "Select HTML Template",
                AllowMultiple = false,
                Filters =
                {
                    new FileDialogFilter { Name = "HTML files", Extensions = { "html", "htm" } },
                    new FileDialogFilter { Name = "All files", Extensions = { "*" } }
                }
            };

            var result = await ofd.ShowAsync(this);
            if (result is null || result.Length == 0) return;

            var sourcePath = result[0];

            try
            {
                if (!File.Exists(sourcePath))
                {
                    ShowMessage("The selected file does not exist.", MessageKind.Error);
                    return;
                }

                var destPath = Path.Combine(_templatesDir, Path.GetFileName(sourcePath));
                destPath = GetUniquePath(destPath);
                File.Copy(sourcePath, destPath, overwrite: false);

                _htmlTemplatePath = destPath;
                TemplateStatusText.Text = $"Template: {Path.GetFileName(destPath)}";
                DeleteHtmlButton.IsVisible = true;

                // Parse placeholders (double braces: {{field}})
                _templatePlaceholders = await ExtractPlaceholdersFromHtml(destPath);

                if (_templatePlaceholders.Count == 0)
                    ShowMessage("No placeholders found in the HTML (use {{Name}}, {{Email}}, etc.).", MessageKind.Warning);
                else
                    ShowMessage($"Template uploaded. Detected {_templatePlaceholders.Count} field(s): {string.Join(", ", _templatePlaceholders)}", MessageKind.Success);

                // Clear CSV state because headers might not match anymore
                _csvPath = null;
                _csvTable = null;
                DataStatusText.Text = "Data: Missing";
                DeleteCsvButton.IsVisible = false;

                UpdateUiState();
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to import template: {ex.Message}", MessageKind.Error);
            }
        }

        // ========================= Upload CSV =========================
        private async void ChooseCsvButton_OnClick(object? sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Title = "Select CSV Data",
                AllowMultiple = false,
                Filters =
                {
                    new FileDialogFilter { Name = "CSV", Extensions = { "csv" } },
                    new FileDialogFilter { Name = "All files", Extensions = { "*" } }
                }
            };

            var result = await ofd.ShowAsync(this);
            if (result is null || result.Length == 0) return;

            var sourcePath = result[0];

            try
            {
                if (!File.Exists(sourcePath))
                {
                    ShowMessage("The selected CSV does not exist.", MessageKind.Error);
                    return;
                }

                var destPath = Path.Combine(_dataDir, Path.GetFileName(sourcePath));
                destPath = GetUniquePath(destPath);
                File.Copy(sourcePath, destPath, overwrite: false);

                _csvPath = destPath;
                DataStatusText.Text = $"Data: {Path.GetFileName(destPath)}";
                DeleteCsvButton.IsVisible = true;

                ValidateCsvAgainstTemplate();
                UpdateUiState();
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to import CSV: {ex.Message}", MessageKind.Error);
            }
        }

        // ========================= Delete buttons =========================
        private void DeleteHtmlButton_OnClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(_htmlTemplatePath) && File.Exists(_htmlTemplatePath))
                    File.Delete(_htmlTemplatePath);
            }
            catch { /* ignore */ }

            _htmlTemplatePath = null;
            _templatePlaceholders.Clear();

            TemplateStatusText.Text = "Template: Missing";
            DeleteHtmlButton.IsVisible = false;

            _csvPath = null;
            _csvTable = null;
            DataStatusText.Text = "Data: Missing";
            DeleteCsvButton.IsVisible = false;

            ShowMessage("Template removed.", MessageKind.Info);
            UpdateUiState();
        }

        private void DeleteCsvButton_OnClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(_csvPath) && File.Exists(_csvPath))
                    File.Delete(_csvPath);
            }
            catch { /* ignore */ }

            _csvPath = null;
            _csvTable = null;

            DataStatusText.Text = "Data: Missing";
            DeleteCsvButton.IsVisible = false;

            ShowMessage("CSV removed.", MessageKind.Info);
            UpdateUiState();
        }

        // ========================= Primary CTA =========================
        private async void PrimaryActionButton_OnClick(object? sender, RoutedEventArgs e)
        {
            switch (_state)
            {
                case ProcessState.MissingHtml:
                    break;

                case ProcessState.HtmlOnly:
                {
                    if (_templatePlaceholders.Count == 0)
                    {
                        ShowMessage("No placeholders found to build a CSV template.", MessageKind.Warning);
                        return;
                    }

                    var sfd = new SaveFileDialog
                    {
                        Title = "Save CSV Template",
                        Filters = { new FileDialogFilter { Name = "CSV", Extensions = { "csv" } } },
                        InitialFileName = SuggestCsvFileName()
                    };
                    var savePath = await sfd.ShowAsync(this);
                    if (string.IsNullOrWhiteSpace(savePath)) return;

                    try
                    {
                        // Export using semicolon (Excel-friendly in PL/EU)
                        var csvHeader = BuildCsvHeader(_templatePlaceholders, ';');
                        await File.WriteAllTextAsync(savePath, csvHeader, Encoding.UTF8);
                        ShowMessage($"CSV template saved to: {savePath}", MessageKind.Success);
                    }
                    catch (Exception ex)
                    {
                        ShowMessage($"Failed to save CSV: {ex.Message}", MessageKind.Error);
                    }
                    break;
                }

                case ProcessState.HtmlAndCsvInvalid:
                    // Disabled; message already explains issues
                    break;

                case ProcessState.HtmlAndCsvValid:
                {
                    // Require signature name
                    var sigName = SignatureNameBox?.Text?.Trim();
                    if (string.IsNullOrWhiteSpace(sigName))
                    {
                        ShowMessage("Please enter a Signature name before generating.", MessageKind.Warning);
                        UpdateUiState();
                        return;
                    }

                    if (_htmlTemplatePath is null || _csvTable is null)
                    {
                        ShowMessage("Internal state error: missing template or CSV.", MessageKind.Error);
                        return;
                    }

                    var sfd = new SaveFileDialog
                    {
                        Title = "Save Signatures ZIP",
                        Filters = { new FileDialogFilter { Name = "ZIP", Extensions = { "zip" } } },
                        InitialFileName = SuggestZipFileName()
                    };
                    var zipPath = await sfd.ShowAsync(this);
                    if (string.IsNullOrWhiteSpace(zipPath)) return;

                    try
                    {
                        var htmlTemplate = await File.ReadAllTextAsync(_htmlTemplatePath, Encoding.UTF8);
                        GenerateZipOfSignatures(zipPath, htmlTemplate, _templatePlaceholders, _csvTable, sigName);
                        ShowMessage($"Signatures ZIP saved to: {zipPath}", MessageKind.Success);
                    }
                    catch (Exception ex)
                    {
                        ShowMessage($"Failed to generate ZIP: {ex.Message}", MessageKind.Error);
                    }
                    break;
                }
            }
        }

        private void SignatureNameBox_OnTextChanged(object? sender, TextChangedEventArgs e)
        {
            UpdateUiState();
        }

        // ========================= State/Validation =========================
        private void UpdateUiState()
        {
            // Determine current state from files
            if (string.IsNullOrWhiteSpace(_htmlTemplatePath))
            {
                _state = ProcessState.MissingHtml;
            }
            else if (string.IsNullOrWhiteSpace(_csvPath))
            {
                _state = ProcessState.HtmlOnly;
            }
            else
            {
                _state = (_csvTable != null) ? ProcessState.HtmlAndCsvValid : ProcessState.HtmlAndCsvInvalid;
            }

            // Reflect in UI + consider signature name requirement
            switch (_state)
            {
                case ProcessState.MissingHtml:
                    PrimaryActionButton.Content = "Generate Template";
                    PrimaryActionButton.IsEnabled = false;
                    break;

                case ProcessState.HtmlOnly:
                    PrimaryActionButton.Content = "Download CSV Template";
                    PrimaryActionButton.IsEnabled = true;
                    break;

                case ProcessState.HtmlAndCsvInvalid:
                    PrimaryActionButton.Content = "Fix CSV (see message)";
                    PrimaryActionButton.IsEnabled = false;
                    break;

                case ProcessState.HtmlAndCsvValid:
                {
                    var hasSigName = !string.IsNullOrWhiteSpace(SignatureNameBox?.Text);
                    PrimaryActionButton.Content = hasSigName ? "Download Signatures" : "Enter Signature Name";
                    PrimaryActionButton.IsEnabled = hasSigName;
                    break;
                }
            }
        }

        private void ValidateCsvAgainstTemplate()
        {
            _csvTable = null;

            if (_csvPath is null)
            {
                ShowMessage("No CSV selected.", MessageKind.Warning);
                return;
            }

            if (_templatePlaceholders.Count == 0)
            {
                ShowMessage("No placeholders found in the HTML to validate against.", MessageKind.Warning);
                return;
            }

            try
            {
                var table = ReadCsv(_csvPath);

                if (table.Headers.Count == 0)
                {
                    ShowMessage("CSV has no headers.", MessageKind.Error);
                    return;
                }

                // Exact set match (case-insensitive)
                var templateSet = new HashSet<string>(_templatePlaceholders, StringComparer.OrdinalIgnoreCase);
                var csvHeaderSet = new HashSet<string>(table.Headers, StringComparer.OrdinalIgnoreCase);

                var missing = templateSet.Except(csvHeaderSet, StringComparer.OrdinalIgnoreCase).ToList();
                var extra   = csvHeaderSet.Except(templateSet, StringComparer.OrdinalIgnoreCase).ToList();

                if (missing.Count > 0 || extra.Count > 0)
                {
                    var sb = new StringBuilder();
                    if (missing.Count > 0) sb.AppendLine("CSV is missing required columns: " + string.Join(", ", missing));
                    if (extra.Count > 0)   sb.AppendLine("CSV has unexpected columns: " + string.Join(", ", extra));
                    ShowMessage(sb.ToString().Trim(), MessageKind.Error);
                    return;
                }

                // Validate rows: all required fields present and non-empty
                var errors = new List<string>();
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    var row = table.Rows[i];
                    foreach (var req in templateSet)
                    {
                        if (!row.TryGetValue(req, out var val) || string.IsNullOrWhiteSpace(val))
                            errors.Add($"Row {i + 2}: missing value for “{req}”.");
                    }
                }

                if (errors.Count > 0)
                {
                    var preview = string.Join(Environment.NewLine, errors.Take(10));
                    var suffix = (errors.Count > 10) ? $" (+{errors.Count - 10} more…)" : "";
                    ShowMessage("CSV has empty/missing values:" + Environment.NewLine + preview + suffix, MessageKind.Error);
                    return;
                }

                _csvTable = table;
                ShowMessage($"CSV validated. {table.Rows.Count} row(s) ready.", MessageKind.Success);
            }
            catch (Exception ex)
            {
                ShowMessage($"Failed to read/validate CSV: {ex.Message}", MessageKind.Error);
            }
        }

        // ========================= CSV & ZIP helpers =========================
        // Auto-detect delimiter (comma, semicolon, or tab)
        private static char DetectDelimiter(string headerLine)
        {
            int commas = headerLine.Count(c => c == ',');
            int semis  = headerLine.Count(c => c == ';');
            int tabs   = headerLine.Count(c => c == '\t');

            if (semis >= commas && semis >= tabs) return ';';
            if (commas >= semis && commas >= tabs) return ',';
            if (tabs > 0) return '\t';
            return ','; // fallback
        }

        private static CsvTable ReadCsv(string path)
        {
            var lines = File.ReadAllLines(path);
            if (lines.Length == 0) return new CsvTable();

            // Header line (first non-empty)
            var headerLineIndex = Array.FindIndex(lines, l => !string.IsNullOrWhiteSpace(l));
            if (headerLineIndex < 0) return new CsvTable();

            var headerLine = lines[headerLineIndex];
            var delimiter  = DetectDelimiter(headerLine);

            var headers = SplitSeparatedLine(headerLine, delimiter)
                .Select(h => h.Trim())
                .Where(h => !string.IsNullOrWhiteSpace(h))
                .ToList();

            var rows = new List<Dictionary<string, string>>();

            for (int i = headerLineIndex + 1; i < lines.Length; i++)
            {
                var raw = lines[i];
                if (string.IsNullOrWhiteSpace(raw)) continue;

                var values = SplitSeparatedLine(raw, delimiter);
                var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                for (int c = 0; c < headers.Count; c++)
                {
                    var key = headers[c];
                    var val = (c < values.Count) ? values[c] : "";
                    row[key] = val?.Trim() ?? "";
                }

                rows.Add(row);
            }

            return new CsvTable { Headers = headers, Rows = rows };
        }

        // CSV splitter supporting quotes and commas/semicolons/tabs
        private static List<string> SplitSeparatedLine(string line, char delimiter)
        {
            var result = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                var ch = line[i];

                if (inQuotes)
                {
                    if (ch == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"')
                        {
                            sb.Append('"'); // escaped quote
                            i++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        sb.Append(ch);
                    }
                }
                else
                {
                    if (ch == delimiter)
                    {
                        result.Add(sb.ToString());
                        sb.Clear();
                    }
                    else if (ch == '"')
                    {
                        inQuotes = true;
                    }
                    else
                    {
                        sb.Append(ch);
                    }
                }
            }

            result.Add(sb.ToString());
            return result;
        }

        // Build header line with chosen separator (use ';' on export)
        private static string BuildCsvHeader(IEnumerable<string> headers, char separator)
        {
            string Escape(string s)
            {
                bool needsQuotes = s.Contains('"') || s.Contains(separator) || s.Contains('\n') || s.Contains('\r');
                if (!needsQuotes) return s;
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            }

            var line = string.Join(separator, headers.Select(Escape));
            return line + Environment.NewLine;
        }

        private static async System.Threading.Tasks.Task<List<string>> ExtractPlaceholdersFromHtml(string path)
        {
            var html = await File.ReadAllTextAsync(path, Encoding.UTF8);

            // Match {{placeholder}}
            var matches = Regex.Matches(html, @"\{\{([^{}\r\n]+)\}\}");

            return matches
                .Select(m => m.Groups[1].Value.Trim())
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string ReplacePlaceholders(string template, IReadOnlyCollection<string> placeholders, IDictionary<string, string> values)
        {
            var output = template;
            foreach (var field in placeholders)
            {
                var token = "{{" + field + "}}"; // double-brace
                var kv = values.FirstOrDefault(k => string.Equals(k.Key, field, StringComparison.OrdinalIgnoreCase));
                output = output.Replace(token, kv.Value ?? string.Empty, StringComparison.Ordinal);
            }
            return output;
        }

        private static string MakeSafeFileName(string? s, string fallback, int index)
        {
            var name = string.IsNullOrWhiteSpace(s) ? $"{fallback}_{index + 1}" : s.Trim();
            foreach (var c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            if (name.Length > 80) name = name[..80];
            return name;
        }

        private void GenerateZipOfSignatures(
            string zipPath,
            string htmlTemplate,
            IReadOnlyCollection<string> placeholders,
            CsvTable table,
            string? signatureName)
        {
            if (File.Exists(zipPath)) File.Delete(zipPath);

            var suffixRaw = string.IsNullOrWhiteSpace(signatureName) ? "signature" : signatureName!.Trim();

            using var fs = new FileStream(zipPath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Create, leaveOpen: false, entryNameEncoding: Encoding.UTF8);

            for (int i = 0; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];

                // Determine prefix: prefer 'name' value; else first header; else fallback
                string prefix;
                if (row.TryGetValue("name", out var nameVal) && !string.IsNullOrWhiteSpace(nameVal))
                {
                    prefix = nameVal;
                }
                else
                {
                    var firstHeader = table.Headers.FirstOrDefault();
                    if (!string.IsNullOrWhiteSpace(firstHeader) &&
                        row.TryGetValue(firstHeader, out var firstVal) &&
                        !string.IsNullOrWhiteSpace(firstVal))
                    {
                        prefix = firstVal;
                    }
                    else
                    {
                        prefix = $"row_{i + 1}";
                    }
                }

                var fileBaseRaw = $"{prefix}_{suffixRaw}";
                var content = ReplacePlaceholders(htmlTemplate, placeholders, row);

                var fileBase = MakeSafeFileName(fileBaseRaw, "signature", i);
                var fileName = fileBase.EndsWith(".html", StringComparison.OrdinalIgnoreCase)
                    ? fileBase
                    : fileBase + ".html";

                var entry = zip.CreateEntry(fileName, CompressionLevel.Optimal);
                using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                writer.Write(content);
            }
        }

        private string SuggestCsvFileName()
        {
            var baseName = "template_fields.csv";
            if (!string.IsNullOrEmpty(_htmlTemplatePath))
            {
                var n = Path.GetFileNameWithoutExtension(_htmlTemplatePath);
                baseName = $"{SanitizeFileName(n)}_fields.csv";
            }
            return baseName;
        }

        private string SuggestZipFileName()
        {
            var baseName = "signatures.zip";
            if (!string.IsNullOrEmpty(_htmlTemplatePath))
            {
                var n = Path.GetFileNameWithoutExtension(_htmlTemplatePath);
                baseName = $"{SanitizeFileName(n)}_signatures.zip";
            }
            return baseName;
        }

        private static string SanitizeFileName(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }

        private static string GetUniquePath(string path)
        {
            if (!File.Exists(path)) return path;
            var dir = Path.GetDirectoryName(path)!;
            var name = Path.GetFileNameWithoutExtension(path);
            var ext = Path.GetExtension(path);
            for (int i = 1; ; i++)
            {
                var candidate = Path.Combine(dir, $"{name} ({i}){ext}");
                if (!File.Exists(candidate))
                    return candidate;
            }
        }

        // ========================= Message area =========================
        private enum MessageKind { Info, Success, Warning, Error }

        private void ShowMessage(string text, MessageKind kind)
        {
            MessageText.Text = text;
            MessageArea.IsVisible = !string.IsNullOrWhiteSpace(text);

            var color = kind switch
            {
                MessageKind.Success => Color.FromRgb(0x22, 0x99, 0x55),
                MessageKind.Warning => Color.FromRgb(0xCC, 0x88, 0x00),
                MessageKind.Error   => Color.FromRgb(0xD6, 0x2F, 0x2F),
                _                   => Color.FromRgb(0x27, 0x32, 0x4A)
            };
            MessageText.Foreground = new SolidColorBrush(color);
        }

        // ========================= DTO =========================
        private sealed class CsvTable
        {
            public List<string> Headers { get; set; } = new();
            public List<Dictionary<string, string>> Rows { get; set; } = new();
        }
    }

    internal static class LinqSetExtensions
    {
        public static IEnumerable<T> Except<T>(this IEnumerable<T> first, IEnumerable<T> second, IEqualityComparer<T> comparer)
            => System.Linq.Enumerable.Except(first, second, comparer);
    }
}

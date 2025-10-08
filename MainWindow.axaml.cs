using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

namespace SignatureGenerator;

public partial class MainWindow : Window
{
    public class FooterResult
    {
        public string Email { get; set; } = "";
        public string FullName { get; set; } = "";
        public string FooterPath { get; set; } = "";
    }

    private readonly ObservableCollection<FooterResult> _generatedFooters = new();
    public ObservableCollection<string> LogLines { get; } = new();

    private string? _manualPath;

    // UI refs
    private TextBlock? _txtHtmlPath, _txtCsvPath, _txtStatus, _txtActionHint, _txtManualPreview;
    private RadioButton? _rbSendEmails, _rbDownloadZip, _rbAttachYes, _rbAttachNo;
    private StackPanel? _panelEmailFields, _panelManual;
    private Expander? _expSmtp;
    private TextBox? _tbSubject, _tbBody, _tbSmtpHost, _tbSmtpUser, _tbSmtpPassword;
    private NumericUpDown? _numSmtpPort;
    private ToggleSwitch? _tgUseSsl;
    private Border? _dropManualZone;
    private Button? _btnRun, _btnGenerate;
    private ProgressBar? _progress;
    private ItemsControl? _logList;

    public MainWindow()
    {
        InitializeComponent();
        CacheControls();

        // Fit window to current screen working area (in DIPs)
        this.Opened += (_, __) =>
        {
            var screen = Screens?.ScreenFromVisual(this);
            if (screen is null) return;

            var wa = screen.WorkingArea; // pixels
            var scale = RenderScaling;   // px per DIP
            var maxW = wa.Width / scale;
            var maxH = wa.Height / scale;

            MaxWidth  = Math.Max(600, maxW - 20);
            MaxHeight = Math.Max(500, maxH - 20);
            Width  = Math.Min(Width,  MaxWidth);
            Height = Math.Min(Height, MaxHeight);
        };

        // wire drag/drop for manual picker
        if (_dropManualZone is not null)
        {
            _dropManualZone.AddHandler(DragDrop.DragOverEvent, DropManualZone_OnDragOver, RoutingStrategies.Tunnel);
            _dropManualZone.AddHandler(DragDrop.DropEvent,      DropManualZone_OnDrop,     RoutingStrategies.Tunnel);
            _dropManualZone.PointerReleased += DropManualZone_OnPointerReleased;
        }

        if (_logList is not null) _logList.ItemsSource = LogLines;

        UpdateStatus();
        UpdateModeUI();
        UpdateAttachManualUI();
        UpdateActionButtons();
    }

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);
    private T? Get<T>(string name) where T : Control => this.FindControl<T>(name);

    private void CacheControls()
    {
        _txtHtmlPath = Get<TextBlock>("TxtHtmlPath");
        _txtCsvPath = Get<TextBlock>("TxtCsvPath");
        _txtStatus = Get<TextBlock>("TxtStatus");

        _rbSendEmails = Get<RadioButton>("RbSendEmails");
        _rbDownloadZip = Get<RadioButton>("RbDownloadZip");

        _panelEmailFields = Get<StackPanel>("PanelEmailFields");
        _expSmtp = Get<Expander>("ExpSmtp");

        _rbAttachYes = Get<RadioButton>("RbAttachYes");
        _rbAttachNo  = Get<RadioButton>("RbAttachNo");
        _panelManual = Get<StackPanel>("PanelManual");
        _dropManualZone = Get<Border>("DropManualZone");
        _txtManualPreview = Get<TextBlock>("TxtManualPreview");

        _tbSubject = Get<TextBox>("TbSubject");
        _tbBody    = Get<TextBox>("TbBody");

        _tbSmtpHost = Get<TextBox>("TbSmtpHost");
        _numSmtpPort= Get<NumericUpDown>("NumSmtpPort");
        _tgUseSsl   = Get<ToggleSwitch>("TgUseSsl");
        _tbSmtpUser = Get<TextBox>("TbSmtpUser");
        _tbSmtpPassword = Get<TextBox>("TbSmtpPassword");

        _txtActionHint = Get<TextBlock>("TxtActionHint");
        _btnRun = Get<Button>("BtnRun");
        _btnGenerate = Get<Button>("BtnGenerate");
        _progress = Get<ProgressBar>("Progress");
        _logList = Get<ItemsControl>("LogList");
    }

    // ---------- File pickers ----------
    private async void BtnChooseHtml_OnClick(object? s, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { AllowMultiple = false };
        dlg.Filters.Add(new FileDialogFilter { Name = "HTML", Extensions = { "html", "htm" } });
        var files = await dlg.ShowAsync(this);
        if (files is { Length: > 0 } && _txtHtmlPath is not null) _txtHtmlPath.Text = files[0];
        UpdateStatus();
    }

    private async void BtnChooseCsv_OnClick(object? s, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { AllowMultiple = false };
        dlg.Filters.Add(new FileDialogFilter { Name = "CSV", Extensions = { "csv" } });
        var files = await dlg.ShowAsync(this);
        if (files is { Length: > 0 } && _txtCsvPath is not null) _txtCsvPath.Text = files[0];
        UpdateStatus();
    }

    private async void BtnChooseManual_OnClick(object? s, RoutedEventArgs e)
        => await BtnChooseManual_WithFilterAsync();

    private async Task BtnChooseManual_WithFilterAsync()
    {
        var dlg = new OpenFileDialog { AllowMultiple = false };
        dlg.Filters.Add(new FileDialogFilter { Name = "PDF", Extensions = { "pdf" } });
        var files = await dlg.ShowAsync(this);
        if (files is { Length: > 0 })
        {
            _manualPath = files[0];
            if (_txtManualPreview is not null) _txtManualPreview.Text = Path.GetFileName(files[0]);
            Log($"Manual set: {_txtManualPreview?.Text}");
        }
    }

    private void UpdateStatus()
    {
        var templateOk = File.Exists(_txtHtmlPath?.Text ?? "");
        var csvOk      = File.Exists(_txtCsvPath?.Text ?? "");
        if (_txtStatus is not null)
            _txtStatus.Text = $"Template: {(templateOk ? "OK" : "Missing")} • Data: {(csvOk ? "OK" : "Missing")}";

        UpdateActionButtons();
    }

    private void UpdateActionButtons()
    {
        var templateOk = File.Exists(_txtHtmlPath?.Text ?? "");
        var csvOk      = File.Exists(_txtCsvPath?.Text ?? "");
        if (_btnGenerate is not null) _btnGenerate.IsEnabled = templateOk && csvOk;
        if (_btnRun is not null)      _btnRun.IsEnabled = _generatedFooters.Any();
    }

    // ---------- Drag/drop for manual ----------
    private void DropManualZone_OnDragOver(object? sender, DragEventArgs e)
    {
        if (e.Data.Contains(DataFormats.FileNames))
        {
            var paths = e.Data.GetFileNames()?.ToArray() ?? Array.Empty<string>();
            if (paths.Length == 1 && Path.GetExtension(paths[0]).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                e.DragEffects = DragDropEffects.Copy;
                e.Handled = true;
                return;
            }
        }
        e.DragEffects = DragDropEffects.None;
        e.Handled = true;
    }

    private void DropManualZone_OnDrop(object? sender, DragEventArgs e)
    {
        var paths = e.Data.GetFileNames()?.ToArray() ?? Array.Empty<string>();
        if (paths.Length == 1 && Path.GetExtension(paths[0]).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            _manualPath = paths[0];
            if (_txtManualPreview is not null) _txtManualPreview.Text = Path.GetFileName(paths[0]);
            Log($"Manual set: {_txtManualPreview?.Text}");
        }
        else
        {
            Log("Please drop a single .pdf file.");
        }
    }

    private async void DropManualZone_OnPointerReleased(object? s, PointerReleasedEventArgs e)
        => await BtnChooseManual_WithFilterAsync();

    // ---------- Mode & Attach Manual ----------
    private void DeliveryMode_Changed(object? s, RoutedEventArgs e) => UpdateModeUI();
    private void AttachManual_Changed(object? s, RoutedEventArgs e) => UpdateAttachManualUI();

    private void UpdateModeUI()
    {
        bool send = _rbSendEmails?.IsChecked == true;

        if (_panelEmailFields is not null) _panelEmailFields.IsVisible = send;
        if (_expSmtp is not null)          _expSmtp.IsVisible = send;

        if (_btnRun is not null)           _btnRun.Content = send ? "Send Emails" : "Download ZIP";
        if (_txtActionHint is not null)
            _txtActionHint.Text = send
                ? "Will send each recipient their HTML footer as an attachment using SMTP."
                : "Will package all generated footers into a ZIP (manual is not included).";
    }

    private void UpdateAttachManualUI()
    {
        bool attach = _rbAttachYes?.IsChecked == true;
        if (_panelManual is not null) _panelManual.IsVisible = attach;
    }

    // ---------- Generate ----------
    private async void BtnGenerate_OnClick(object? s, RoutedEventArgs e)
    {
        LogLines.Clear();

        var htmlPath = _txtHtmlPath?.Text;
        var csvPath  = _txtCsvPath?.Text;

        if (!File.Exists(htmlPath ?? "") || !File.Exists(csvPath ?? ""))
        {
            Log("Please select both an HTML template and a CSV file before generating.");
            UpdateActionButtons();
            return;
        }

        try
        {
            _generatedFooters.Clear();
            UpdateActionButtons();

            var template = await File.ReadAllTextAsync(htmlPath!);
            var lines = await File.ReadAllLinesAsync(csvPath!);

            if (lines.Length < 2)
            {
                Log("CSV seems empty (needs a header and at least one data row).");
                return;
            }

            var headers = ParseCsvLine(lines[0]).Select(h => h.Trim()).ToArray();
            var outDir = Path.Combine(Path.GetTempPath(), "signatures_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            Directory.CreateDirectory(outDir);

            if (_progress is not null) { _progress.IsVisible = true; _progress.Value = 0; _progress.Maximum = lines.Length - 1; }

            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                var cols = ParseCsvLine(lines[i]);
                if (cols.Length != headers.Length)
                {
                    Log($"Row {i}: column count mismatch, skipping.");
                    continue;
                }

                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int c = 0; c < headers.Length; c++)
                    map[headers[c]] = cols[c];

                var filled = ReplacePlaceholders(template, map);
                var safeName  = map.TryGetValue("name", out var nm) && !string.IsNullOrWhiteSpace(nm) ? nm : $"user{i}";
                var safeEmail = map.TryGetValue("email", out var em) && !string.IsNullOrWhiteSpace(em) ? em : $"noemail{i}";
                var fileName = $"{SanitizeFileName(safeName)}_{SanitizeFileName(safeEmail)}.html";
                var fullPath = Path.Combine(outDir, fileName);

                await File.WriteAllTextAsync(fullPath, filled, Encoding.UTF8);

                _generatedFooters.Add(new FooterResult
                {
                    Email = map.TryGetValue("email", out var eaddr) ? eaddr : "",
                    FullName = map.TryGetValue("name", out var fname) ? fname : safeName,
                    FooterPath = fullPath
                });

                if (_progress is not null) _progress.Value += 1;
            }

            Log($"Generated {_generatedFooters.Count} signatures into: {outDir}");
        }
        catch (Exception ex)
        {
            Log($"Generation error: {ex.Message}");
        }
        finally
        {
            if (_progress is not null) _progress.IsVisible = false;
            UpdateActionButtons();
        }
    }

    // ---------- Run ----------
    private async void BtnRun_OnClick(object? s, RoutedEventArgs e)
    {
        LogLines.Clear();
        if (!_generatedFooters.Any())
        {
            Log("Generate footers first!");
            UpdateActionButtons();
            return;
        }

        if (_rbSendEmails?.IsChecked == true)
            await SendEmailsAsync();
        else
            await DownloadZipAsync();
    }

    // ---------- Test SMTP ----------
    private async void BtnTestSmtp_OnClick(object? s, RoutedEventArgs e)
    {
        Log("Testing SMTP connection...");
        try
        {
            using var smtp = new SmtpClient();
            var secure = _tgUseSsl?.IsChecked == true ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTls;

            smtp.AuthenticationMechanisms.Remove("XOAUTH2");

            await smtp.ConnectAsync((_tbSmtpHost?.Text ?? "").Trim(), (int)(_numSmtpPort?.Value ?? 587), secure);
            await smtp.AuthenticateAsync((_tbSmtpUser?.Text ?? "").Trim(),
                                         (_tbSmtpPassword?.Text ?? "").Replace(" ", "").Trim());
            await smtp.DisconnectAsync(true);

            Log("✅ SMTP connection OK!");
        }
        catch (Exception ex)
        {
            Log($"❌ SMTP test failed: {ex.Message}");
        }
    }

    // ---------- Send Emails ----------
    private async Task SendEmailsAsync()
    {
        var host   = (_tbSmtpHost?.Text ?? "").Trim();
        var port   = (int)(_numSmtpPort?.Value ?? 587);
        var useSsl = _tgUseSsl?.IsChecked == true;
        var user   = (_tbSmtpUser?.Text ?? "").Trim();
        var pass   = (_tbSmtpPassword?.Text ?? "").Replace(" ", "").Trim();
        var subj   = _tbSubject?.Text ?? "";
        var body   = _tbBody?.Text ?? "";
        var attachManual = _rbAttachYes?.IsChecked == true;

        if (_progress is not null) { _progress.IsVisible = true; _progress.Value = 0; _progress.Maximum = _generatedFooters.Count; }

        try
        {
            using var smtp = new SmtpClient();
            var secure = useSsl ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTls;
            smtp.AuthenticationMechanisms.Remove("XOAUTH2");

            await smtp.ConnectAsync(host, port, secure);
            await smtp.AuthenticateAsync(user, pass);

            foreach (var item in _generatedFooters)
            {
                if (!IsDeliverableEmail(item.Email))
                {
                    Log($"Skipping reserved/invalid recipient: {item.Email}");
                    if (_progress is not null) _progress.Value += 1;
                    continue;
                }

                var msg = new MimeMessage();
                msg.From.Add(new MailboxAddress(user, user));
                msg.To.Add(MailboxAddress.Parse(item.Email));
                msg.Subject = subj;

                var builder = new BodyBuilder { TextBody = body };

                if (File.Exists(item.FooterPath))
                    builder.Attachments.Add(Path.GetFileName(item.FooterPath),
                        await File.ReadAllBytesAsync(item.FooterPath),
                        new ContentType("text", "html"));

                if (attachManual && !string.IsNullOrWhiteSpace(_manualPath) && File.Exists(_manualPath))
                    builder.Attachments.Add(Path.GetFileName(_manualPath),
                        await File.ReadAllBytesAsync(_manualPath),
                        new ContentType("application", "pdf"));

                msg.Body = builder.ToMessageBody();
                await smtp.SendAsync(msg);
                Log($"Sent to {item.FullName} <{item.Email}>");

                if (_progress is not null) _progress.Value += 1;
            }

            await smtp.DisconnectAsync(true);
            Log("All emails sent.");
        }
        catch (Exception ex)
        {
            Log($"Send error: {ex.Message}");
        }
        finally
        {
            if (_progress is not null) _progress.IsVisible = false;
        }
    }

    // ---------- ZIP (manual not included) ----------
    private async Task DownloadZipAsync()
    {
        var sfd = new SaveFileDialog
        {
            Title = "Save ZIP",
            InitialFileName = "signatures.zip",
            Filters = new List<FileDialogFilter> { new() { Name = "ZIP", Extensions = { "zip" } } }
        };

        var zipPath = await sfd.ShowAsync(this);
        if (string.IsNullOrWhiteSpace(zipPath)) return;

        if (_progress is not null) { _progress.IsVisible = true; _progress.Value = 0; _progress.Maximum = _generatedFooters.Count; }

        try
        {
            if (File.Exists(zipPath)) File.Delete(zipPath);
            using var zip = ZipFile.Open(zipPath, ZipArchiveMode.Create);

            foreach (var item in _generatedFooters)
            {
                if (File.Exists(item.FooterPath))
                    zip.CreateEntryFromFile(item.FooterPath, Path.GetFileName(item.FooterPath));
                if (_progress is not null) _progress.Value += 1;
            }

            Log($"ZIP saved: {zipPath} (manual excluded)");
        }
        catch (Exception ex)
        {
            Log($"ZIP error: {ex.Message}");
        }
        finally
        {
            if (_progress is not null) _progress.IsVisible = false;
        }
    }

    // ---------- Helpers ----------
    private static string SanitizeFileName(string s)
    {
        var invalid = Path.GetInvalidFileNameChars();
        return new string(s.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
    }

    private static string[] ParseCsvLine(string line)
    {
        var result = new List<string>();
        if (line is null) return Array.Empty<string>();

        var sb = new StringBuilder();
        bool inQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];

            if (c == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                { sb.Append('"'); i++; }
                else inQuotes = !inQuotes;
            }
            else if (c == ',' && !inQuotes)
            { result.Add(sb.ToString()); sb.Clear(); }
            else sb.Append(c);
        }

        result.Add(sb.ToString());
        return result.ToArray();
    }

    private static string ReplacePlaceholders(string template, Dictionary<string, string> map)
    {
        if (string.IsNullOrEmpty(template)) return template;
        string result = template;
        foreach (var kv in map)
        {
            var pattern = @"\{" + Regex.Escape(kv.Key) + @"\}";
            result = Regex.Replace(result, pattern, kv.Value ?? string.Empty, RegexOptions.IgnoreCase);
        }
        return result;
    }

    private static bool IsDeliverableEmail(string? email)
    {
        if (string.IsNullOrWhiteSpace(email)) return false;
        try
        {
            var mb = MimeKit.MailboxAddress.Parse(email.Trim());
            var domain = mb.Address.Split('@').Last().ToLowerInvariant();

            if (domain is "example.com" or "example.net" or "example.org" ||
                domain.EndsWith(".test") || domain.EndsWith(".example") ||
                domain.EndsWith(".invalid") || domain.EndsWith(".localhost"))
                return false;

            return true;
        }
        catch { return false; }
    }

    private void Log(string text) => LogLines.Add($"[{DateTime.Now:HH:mm:ss}] {text}");
}

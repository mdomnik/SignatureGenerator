using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Layout;
using Avalonia.Media;
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

    // UI refs (match the latest XAML you’re using)
    private TextBlock? _txtHtmlPath, _txtCsvPath, _txtStatus, _txtActionHint;
    private RadioButton? _rbSendEmails, _rbDownloadZip;
    private TextBox? _tbSubject, _tbBody;
    private TextBox? _tbSmtpHost, _tbSmtpUser, _tbSmtpPassword;
    private NumericUpDown? _numSmtpPort;
    private ToggleSwitch? _tgUseSsl;
    private TextBox? _tbFooterName;

    // Manual section (compact row)
    private CheckBox? _chkAttachManual;
    private TextBox? _tbManualPath;        // display-only
    private Button? _btnBrowseManual;

    private Button? _btnRun, _btnGenerate, _btnTestSmtp;
    private ProgressBar? _progress;
    private ItemsControl? _logList;

    public MainWindow()
    {
        InitializeComponent();
        CacheControls();

        // Fit window to working area
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

        // Ensure button text not clipped for the small outline “Test SMTP” button
        if (_btnTestSmtp is not null)
        {
            _btnTestSmtp.MinHeight = 34;
            _btnTestSmtp.Padding = new Thickness(16, 8);
        }

        // Wire up mode/checkbox changes (if XAML didn’t already)
        if (_rbSendEmails != null) _rbSendEmails.IsCheckedChanged += DeliveryMode_Changed;
        if (_rbDownloadZip != null) _rbDownloadZip.IsCheckedChanged += DeliveryMode_Changed;
        if (_chkAttachManual != null) _chkAttachManual.IsCheckedChanged += ChkAttachManual_Changed;

        if (_logList is not null) _logList.ItemsSource = LogLines;

        UpdateStatus();
        UpdateModeUI();
        UpdateActionButtons();
        ApplyManualUiState();
    }

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);
    private T? Get<T>(string name) where T : Control => this.FindControl<T>(name);

    private void CacheControls()
    {
        _txtHtmlPath = Get<TextBlock>("TxtHtmlPath");
        _txtCsvPath  = Get<TextBlock>("TxtCsvPath");
        _txtStatus   = Get<TextBlock>("TxtStatus");
        _txtActionHint = Get<TextBlock>("TxtActionHint");
        _tbFooterName = Get<TextBox>("TbFooterName");

        _rbSendEmails  = Get<RadioButton>("RbSendEmails");
        _rbDownloadZip = Get<RadioButton>("RbDownloadZip");

        _tbSubject = Get<TextBox>("TbSubject");
        _tbBody    = Get<TextBox>("TbBody");

        _tbSmtpHost = Get<TextBox>("TbSmtpHost");
        _numSmtpPort= Get<NumericUpDown>("NumSmtpPort");
        _tgUseSsl   = Get<ToggleSwitch>("TgUseSsl");
        _tbSmtpUser = Get<TextBox>("TbSmtpUser");
        _tbSmtpPassword = Get<TextBox>("TbSmtpPassword");

        // Manual row
        _chkAttachManual = Get<CheckBox>("ChkAttachManual");
        _tbManualPath    = Get<TextBox>("TbManualPath");
        _btnBrowseManual = Get<Button>("BtnBrowseManual");

        // Actions
        _btnRun      = Get<Button>("BtnRun");
        _btnGenerate = Get<Button>("BtnGenerate");
        _btnTestSmtp = Get<Button>("BtnTestSmtp");

        _progress = Get<ProgressBar>("Progress");
        _logList  = Get<ItemsControl>("LogList");
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
            if (_tbManualPath is not null) _tbManualPath.Text = Path.GetFileName(files[0]); // show filename only
            Log($"Manual set: {Path.GetFileName(files[0])}");
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

    // ---------- Mode & Manual UI ----------
    private void DeliveryMode_Changed(object? s, RoutedEventArgs e)
        => UpdateModeUI();

    private void UpdateModeUI()
    {
        bool send = _rbSendEmails?.IsChecked == true;

        // Enable/disable email-related fields when switching to ZIP
        SetEmailModeEnabled(send);

        if (_btnRun is not null) _btnRun.Content = send ? "Send Emails" : "Download ZIP";
        if (_txtActionHint is not null)
            _txtActionHint.Text = send
                ? "Each recipient will receive their own HTML file as an attachment."
                : "All generated footers will be packed into a ZIP file.";
    }

    private void SetEmailModeEnabled(bool enabled)
    {
        if (_tbSubject != null) _tbSubject.IsEnabled = enabled;
        if (_tbBody    != null) _tbBody.IsEnabled    = enabled;

        if (_chkAttachManual != null)
        {
            _chkAttachManual.IsEnabled = enabled;
            if (!enabled)
            {
                // In ZIP mode: clear and lock manual fields
                _chkAttachManual.IsChecked = false;
                _manualPath = null;
                if (_tbManualPath != null) _tbManualPath.Text = "";
            }
        }

        // Manual row: display is always read-only; browse only when both send & checked
        ApplyManualUiState();

        // SMTP controls
        if (_tbSmtpHost != null) _tbSmtpHost.IsEnabled = enabled;
        if (_numSmtpPort!= null) _numSmtpPort.IsEnabled= enabled;
        if (_tgUseSsl    != null) _tgUseSsl.IsEnabled   = enabled;
        if (_tbSmtpUser  != null) _tbSmtpUser.IsEnabled = enabled;
        if (_tbSmtpPassword != null) _tbSmtpPassword.IsEnabled = enabled;
        if (_btnTestSmtp != null) _btnTestSmtp.IsEnabled = enabled;
    }

    private void ChkAttachManual_Changed(object? s, RoutedEventArgs e) => ApplyManualUiState();

    private void ApplyManualUiState()
    {
        bool send = _rbSendEmails?.IsChecked == true;
        bool attach = _chkAttachManual?.IsChecked == true;

        // Display-only textbox stays non-interactive
        if (_tbManualPath != null)
        {
            _tbManualPath.IsReadOnly = true;
            _tbManualPath.IsHitTestVisible = false;
            _tbManualPath.Focusable = false;
            _tbManualPath.IsTabStop = false;
        }

        if (_btnBrowseManual != null)
            _btnBrowseManual.IsEnabled = send && attach;
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
                Log("The CSV file appears to be empty (requires header + at least one data row).");
                return;
            }

            var headers = ParseCsvLine(lines[0]).Select(h => h.Trim()).ToArray();
            var outDir = Path.Combine(Path.GetTempPath(), "footers_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            Directory.CreateDirectory(outDir);

            if (_progress is not null) { _progress.IsVisible = true; _progress.Value = 0; _progress.Maximum = lines.Length - 1; }

            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                var cols = ParseCsvLine(lines[i]);
                if (cols.Length != headers.Length)
                {
                    Log($"Row {i}: mismatched number of columns – skipping.");
                    continue;
                }

                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int c = 0; c < headers.Length; c++)
                    map[headers[c]] = cols[c];

                var filled = ReplacePlaceholders(template, map);
                var safeName  = map.TryGetValue("name", out var nm) && !string.IsNullOrWhiteSpace(nm) ? nm : $"user{i}";
                var safeEmail = map.TryGetValue("email", out var em) && !string.IsNullOrWhiteSpace(em) ? em : $"noemail{i}";
                var footerBase = _tbFooterName?.Text?.Trim();
                if (string.IsNullOrWhiteSpace(footerBase))
                    footerBase = "Signature"; // fallback default

                var fileName = $"{SanitizeFileName(footerBase)}({SanitizeFileName(safeEmail)}).html";
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

            Log($"Generated {_generatedFooters.Count} footers in: {outDir}");
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
            Log("Generate footers first.");
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
        Log("Testing SMTP connection…");
        try
        {
            using var smtp = new SmtpClient();
            var secure = _tgUseSsl?.IsChecked == true ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTls;

            smtp.AuthenticationMechanisms.Remove("XOAUTH2");

            await smtp.ConnectAsync((_tbSmtpHost?.Text ?? "").Trim(), (int)(_numSmtpPort?.Value ?? 587), secure);
            await smtp.AuthenticateAsync((_tbSmtpUser?.Text ?? "").Trim(),
                                         (_tbSmtpPassword?.Text ?? "").Replace(" ", "").Trim());
            await smtp.DisconnectAsync(true);

            Log("✅ SMTP connection successful.");
        }
        catch (Exception ex)
        {
            Log($"❌ SMTP test failed: {ex.Message}");
        }
    }

    // ---------- Email send ----------
    private async Task SendEmailsAsync()
    {
        var host   = (_tbSmtpHost?.Text ?? "").Trim();
        var port   = (int)(_numSmtpPort?.Value ?? 587);
        var useSsl = _tgUseSsl?.IsChecked == true;
        var user   = (_tbSmtpUser?.Text ?? "").Trim();
        var pass   = (_tbSmtpPassword?.Text ?? "").Replace(" ", "").Trim();
        var subj   = _tbSubject?.Text ?? "";
        var body   = _tbBody?.Text ?? "";
        var attachManual = _chkAttachManual?.IsChecked == true;

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
                    Log($"Skipping test/invalid address: {item.Email}");
                    if (_progress is not null) _progress.Value += 1;
                    continue;
                }

                // --- Create per-user ZIP file ---
                var footerBase = _tbFooterName?.Text?.Trim();
                if (string.IsNullOrWhiteSpace(footerBase))
                    footerBase = "Signature";

                var zipName = $"{SanitizeFileName(footerBase)}({SanitizeFileName(item.Email)}).zip";
                var tempZip = Path.Combine(Path.GetTempPath(), zipName);

                if (File.Exists(tempZip))
                    File.Delete(tempZip);

                using (var zip = ZipFile.Open(tempZip, ZipArchiveMode.Create))
                {
                    if (File.Exists(item.FooterPath))
                        zip.CreateEntryFromFile(item.FooterPath, Path.GetFileName(item.FooterPath));

                    if (attachManual && !string.IsNullOrWhiteSpace(_manualPath) && File.Exists(_manualPath))
                        zip.CreateEntryFromFile(_manualPath, Path.GetFileName(_manualPath));
                }

                // --- Compose message ---
                var msg = new MimeMessage();
                msg.From.Add(new MailboxAddress(user, user));
                msg.To.Add(MailboxAddress.Parse(item.Email));
                msg.Subject = subj;

                var builder = new BodyBuilder { TextBody = body };

                // Attach ZIP instead of raw HTML
                builder.Attachments.Add(zipName,
                    await File.ReadAllBytesAsync(tempZip),
                    new ContentType("application", "zip"));

                msg.Body = builder.ToMessageBody();

                await smtp.SendAsync(msg);
                Log($"Sent ZIP to: {item.FullName} <{item.Email}>");

                if (_progress is not null) _progress.Value += 1;
            }

            await smtp.DisconnectAsync(true);
            Log("All messages sent.");
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

    // ---------- ZIP (no manual in ZIP) ----------
    private async Task DownloadZipAsync()
    {
        var sfd = new SaveFileDialog
        {
            Title = "Save ZIP",
            InitialFileName = "footers.zip",
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

            Log($"Saved ZIP: {zipPath} (without manual).");
        }
        catch (Exception ex)
        {
            Log($"ZIP save error: {ex.Message}");
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

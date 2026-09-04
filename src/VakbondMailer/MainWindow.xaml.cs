using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using VakbondMailer.Models;
using VakbondMailer.Services;

namespace VakbondMailer;

public partial class MainWindow : Window
{
    private const string DefaultAccountLabel = "Standaardaccount van Outlook";
    private const string RequiredEmailDomain = "@fnv.nl";

    private readonly OutlookMailService _outlookService = new();
    private readonly ObservableCollection<LogEntry> _logEntries = new();
    private readonly ObservableCollection<string> _attachmentPaths = new();

    private string? _currentFilePath;
    private string? _templatesFolder;
    private ImportedRecipients? _imported;
    private TextBox? _lastFocusedTemplateBox;
    private bool _isSending;
    private bool _testRecipientIsCustom;
    private bool _suppressTestRecipientTracking;
    private bool _accountsLoaded;
    private CancellationTokenSource? _sendCts;

    private string? SelectedAccountName =>
        AccountComboBox.SelectedItem is string name && name != DefaultAccountLabel ? name : null;

    public MainWindow()
    {
        InitializeComponent();
        _lastFocusedTemplateBox = BodyTextBox;
        LogListBox.ItemsSource = _logEntries;
        AttachmentListControl.ItemsSource = _attachmentPaths;
        UpdateMergedPreview();

        AccountComboBox.ItemsSource = new[] { DefaultAccountLabel };
        AccountComboBox.SelectedIndex = 0;

        var settings = AppSettingsService.Load();
        if (!string.IsNullOrWhiteSpace(settings.TemplatesFolder) && Directory.Exists(settings.TemplatesFolder))
        {
            _templatesFolder = settings.TemplatesFolder;
            TemplateFolderText.Text = _templatesFolder;
            RefreshTemplateList();
        }
    }

    private void ChooseFileButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "CSV en Excel-bestanden (*.csv;*.xlsx)|*.csv;*.xlsx|Alle bestanden (*.*)|*.*",
        };

        if (dialog.ShowDialog() != true)
            return;

        // Eerst de vorige lijst loslaten: als het inlezen hierna misgaat, mag er geen oude
        // ledenlijst blijven staan waar je per ongeluk naartoe verstuurt.
        ClearLoadedList();

        try
        {
            _currentFilePath = dialog.FileName;
            FilePathText.Text = Path.GetFileName(_currentFilePath);
            FilePathText.Foreground = (System.Windows.Media.Brush)FindResource("InkBrush");

            var headers = RecipientImportService.ReadHeaders(_currentFilePath);
            EmailColumnComboBox.ItemsSource = headers;
            var guessed = RecipientImportService.GuessEmailColumn(headers);
            EmailColumnComboBox.SelectedItem = guessed ?? headers.FirstOrDefault();

            RenderPlaceholderChips(headers);

            // SelectionChanged op de ComboBox triggert hierna automatisch het inladen.
        }
        catch (Exception ex)
        {
            ClearLoadedList();
            MessageBox.Show(this, $"Kon het bestand niet lezen:\n{ex.Message}", "Fout bij inlezen",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// Zet alles wat van de ingeladen ledenlijst afhangt terug op leeg, zodat er nooit een
    /// half ingeladen of verouderde lijst achterblijft.
    /// </summary>
    private void ClearLoadedList()
    {
        _imported = null;
        _currentFilePath = null;
        PreviewDataGrid.ItemsSource = null;
        EmailColumnComboBox.ItemsSource = null;
        PlaceholderPanel.Children.Clear();
        RecipientCountText.Text = "Geen lijst geladen";
        FilePathText.Text = "Nog geen bestand geladen";
        FilePathText.Foreground = (System.Windows.Media.Brush)FindResource("InkFaintBrush");
        UpdateMergedPreview();
        UpdatePlaceholderWarning();
    }

    private void EmailColumnComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        ImportWithCurrentSettings();
    }

    private void ImportWithCurrentSettings()
    {
        if (_currentFilePath is null || EmailColumnComboBox.SelectedItem is not string emailColumn)
            return;

        try
        {
            _imported = RecipientImportService.Import(_currentFilePath, emailColumn);

            PreviewDataGrid.ItemsSource = BuildPreviewTable(_imported).DefaultView;
            RecipientCountText.Text = $"{_imported.Recipients.Count} ontvanger(s) geladen";

            foreach (var warning in _imported.Warnings)
                Log(warning);

            if (_imported.Warnings.Count > 0)
            {
                MessageBox.Show(this,
                    $"Let op: {_imported.Warnings.Count} rij(en) zijn overgeslagen bij het inladen (geen, ongeldig of dubbel e-mailadres). Zie het logboek onderaan voor details.",
                    "Rijen overgeslagen", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            UpdateMergedPreview();
            UpdatePlaceholderWarning();
        }
        catch (Exception ex)
        {
            // Ook hier: geen half/verouderd resultaat laten staan waar naartoe verstuurd kan worden.
            _imported = null;
            PreviewDataGrid.ItemsSource = null;
            RecipientCountText.Text = "Geen lijst geladen";
            UpdateMergedPreview();
            UpdatePlaceholderWarning();

            MessageBox.Show(this, $"Kon de lijst niet inladen:\n{ex.Message}", "Fout bij inlezen",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static DataTable BuildPreviewTable(ImportedRecipients imported)
    {
        var table = new DataTable();
        foreach (var header in imported.Headers)
            table.Columns.Add(header);

        foreach (var recipient in imported.Recipients)
        {
            var row = table.NewRow();
            foreach (var header in imported.Headers)
                row[header] = recipient.Fields.TryGetValue(header, out var value) ? value : string.Empty;
            table.Rows.Add(row);
        }

        return table;
    }

    private void RenderPlaceholderChips(IReadOnlyList<string> headers)
    {
        PlaceholderPanel.Children.Clear();
        foreach (var header in headers)
        {
            var button = new Button
            {
                Content = $"{{{{{header}}}}}",
                Style = (Style)FindResource("ChipButton"),
            };
            button.Click += (_, _) => InsertPlaceholder(header);
            PlaceholderPanel.Children.Add(button);
        }
    }

    private void SubjectTextBox_GotFocus(object sender, RoutedEventArgs e) => _lastFocusedTemplateBox = SubjectTextBox;

    private void BodyTextBox_GotFocus(object sender, RoutedEventArgs e) => _lastFocusedTemplateBox = BodyTextBox;

    private void InsertPlaceholder(string header)
    {
        var target = _lastFocusedTemplateBox ?? BodyTextBox;
        var placeholder = $"{{{{{header}}}}}";
        var caret = target.CaretIndex;
        target.Text = target.Text.Insert(caret, placeholder);
        target.CaretIndex = caret + placeholder.Length;
        target.Focus();
    }

    private void Template_TextChanged(object sender, TextChangedEventArgs e)
    {
        UpdateMergedPreview();
        UpdatePlaceholderWarning();
    }

    private void PreviewDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e) => UpdateMergedPreview();

    /// <summary>
    /// Het momenteel geselecteerde rijtje in de lijst-preview, of anders de eerste ontvanger.
    /// </summary>
    private Recipient? GetPreviewRecipient()
    {
        if (_imported is null || _imported.Recipients.Count == 0)
            return null;

        var index = PreviewDataGrid.SelectedIndex;
        return index >= 0 && index < _imported.Recipients.Count
            ? _imported.Recipients[index]
            : _imported.Recipients[0];
    }

    private void UpdateMergedPreview()
    {
        var recipient = GetPreviewRecipient();
        if (recipient is null)
        {
            PreviewRecipientLabel.Text = string.Empty;
            MergedPreviewSubject.Text = "(nog geen lijst geladen)";
            MergedPreviewBody.Text = "Laad een CSV of Excel-bestand om een echt voorbeeld te zien.";
            return;
        }

        PreviewRecipientLabel.Text = $"Voorbeeld van: {recipient.DisplayName} ({recipient.Email})";
        MergedPreviewSubject.Text = TemplateRenderer.Render(SubjectTextBox.Text, recipient);
        MergedPreviewBody.Text = TemplateRenderer.Render(BodyTextBox.Text, recipient);
    }

    /// <summary>
    /// Placeholders die in onderwerp/tekst gebruikt worden maar geen kolom in de ledenlijst zijn
    /// (bv. een tikfout), zodat die niet stilletjes als letterlijke tekst verstuurd worden.
    /// </summary>
    private IReadOnlyList<string> FindUnknownPlaceholders()
    {
        if (_imported is null)
            return Array.Empty<string>();

        return TemplateRenderer.ExtractPlaceholders(SubjectTextBox.Text)
            .Concat(TemplateRenderer.ExtractPlaceholders(BodyTextBox.Text))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(token => !_imported.Headers.Contains(token, StringComparer.OrdinalIgnoreCase))
            .ToList();
    }

    private void UpdatePlaceholderWarning()
    {
        var unknown = FindUnknownPlaceholders();
        if (unknown.Count == 0)
        {
            PlaceholderWarningText.Visibility = Visibility.Collapsed;
            return;
        }

        PlaceholderWarningText.Text =
            $"Onbekend veld: {string.Join(", ", unknown.Select(u => $"{{{{{u}}}}}"))} — komt niet voor als kolom in de ledenlijst.";
        PlaceholderWarningText.Visibility = Visibility.Visible;
    }

    private async void SendTestButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isSending)
            return;

        SetSendingState(true);
        try
        {
            // Outlook-COM-objecten zijn STA-gebonden: bewust op de UI-thread aanroepen
            // (niet via Task.Run) om apartment-threading-problemen te vermijden.
            var accountName = SelectedAccountName;
            var myEmail = _outlookService.GetAccountEmail(accountName);

            if (!IsFnvAddress(myEmail))
            {
                MessageBox.Show(this,
                    $"Dit account ({myEmail}) is geen FNV-adres. Kies via 'Accounts vernieuwen' een account dat eindigt op {RequiredEmailDomain}.",
                    "Alleen FNV-adressen toegestaan", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var testRecipient = string.IsNullOrWhiteSpace(TestRecipientTextBox.Text)
                ? myEmail
                : TestRecipientTextBox.Text.Trim();

            var sampleRecipient = GetPreviewRecipient()
                ?? new Recipient { Email = testRecipient, Fields = new Dictionary<string, string>() };

            var attachments = _attachmentPaths.ToList();
            var missingAttachment = attachments.FirstOrDefault(path => !File.Exists(path));
            if (missingAttachment is not null)
            {
                MessageBox.Show(this,
                    $"Deze bijlage bestaat niet meer:\n{missingAttachment}\n\nVerwijder hem uit de lijst of kies het bestand opnieuw.",
                    "Bijlage niet gevonden", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var isHtml = HtmlFormattingCheckBox.IsChecked == true;
            var subject = TemplateRenderer.Render(SubjectTextBox.Text, sampleRecipient);
            var renderedBody = TemplateRenderer.Render(BodyTextBox.Text, sampleRecipient);
            var body = isHtml ? SimpleHtmlFormatter.ToHtml(renderedBody) : renderedBody;

            _outlookService.SendMail(testRecipient, $"[TEST] {subject}", body, accountName, isHtml, attachments);
            await Task.Yield();

            LogSend("Testmail", testRecipient, true);
            MessageBox.Show(this, $"Testmail verstuurd naar {testRecipient}.", "Gelukt",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (OutlookNotAvailableException ex)
        {
            MessageBox.Show(this, ex.Message, "Outlook niet beschikbaar", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Onverwachte fout:\n{ex.Message}", "Fout", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetSendingState(false);
        }
    }

    private async void SendAllButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isSending)
            return;

        if (_imported is null || _imported.Recipients.Count == 0)
        {
            MessageBox.Show(this, "Laad eerst een ledenlijst met minstens één ontvanger.", "Geen ontvangers",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(SubjectTextBox.Text) || string.IsNullOrWhiteSpace(BodyTextBox.Text))
        {
            MessageBox.Show(this, "Vul eerst een onderwerp en tekst in.", "Sjabloon leeg",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var accountName = SelectedAccountName;
        string senderEmail;
        try
        {
            senderEmail = _outlookService.GetAccountEmail(accountName);
        }
        catch (OutlookNotAvailableException ex)
        {
            MessageBox.Show(this, ex.Message, "Outlook niet beschikbaar", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (!IsFnvAddress(senderEmail))
        {
            MessageBox.Show(this,
                $"Dit account ({senderEmail}) is geen FNV-adres. Kies via 'Accounts vernieuwen' een account dat eindigt op {RequiredEmailDomain}.",
                "Alleen FNV-adressen toegestaan", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var unknownPlaceholders = FindUnknownPlaceholders();
        if (unknownPlaceholders.Count > 0)
        {
            var confirmUnknown = MessageBox.Show(this,
                $"Deze velden komen niet voor in je ledenlijst en blijven letterlijk in de mail staan: " +
                $"{string.Join(", ", unknownPlaceholders.Select(u => $"{{{{{u}}}}}"))}.\n\nToch doorgaan?",
                "Onbekende velden", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (confirmUnknown != MessageBoxResult.Yes)
                return;
        }

        var count = _imported.Recipients.Count;
        var confirm = MessageBox.Show(this,
            $"Weet je zeker dat je deze mail wilt versturen naar {count} ontvanger(s), vanaf {senderEmail}?",
            "Bevestig verzenden", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
            return;

        var delay = TimeSpan.FromSeconds(ParseDelaySeconds(DelayTextBox.Text));
        var isHtml = HtmlFormattingCheckBox.IsChecked == true;
        var attachments = _attachmentPaths.ToList();

        var missingAttachment = attachments.FirstOrDefault(path => !File.Exists(path));
        if (missingAttachment is not null)
        {
            MessageBox.Show(this,
                $"Deze bijlage bestaat niet meer:\n{missingAttachment}\n\nVerwijder hem uit de lijst of kies het bestand opnieuw.",
                "Bijlage niet gevonden", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Onderwerp en tekst één keer vastleggen: als er tijdens het verzenden nog in de
        // velden getypt wordt, mogen de resterende mails daar niet door veranderen.
        var subjectTemplate = SubjectTextBox.Text;
        var bodyTemplate = BodyTextBox.Text;

        _sendCts = new CancellationTokenSource();
        var token = _sendCts.Token;

        SetSendingState(true, canCancel: true);
        SendProgressBar.Maximum = count;
        SendProgressBar.Value = 0;
        _logEntries.Clear();

        var results = new List<SendResult>();
        var cancelled = false;
        for (var i = 0; i < _imported.Recipients.Count; i++)
        {
            if (token.IsCancellationRequested)
            {
                cancelled = true;
                Log($"Verzending gestopt door gebruiker na {i} van {count}.");
                break;
            }

            var recipient = _imported.Recipients[i];
            var subject = TemplateRenderer.Render(subjectTemplate, recipient);
            var renderedBody = TemplateRenderer.Render(bodyTemplate, recipient);
            var body = isHtml ? SimpleHtmlFormatter.ToHtml(renderedBody) : renderedBody;

            try
            {
                // Bewust synchroon op de UI-thread: Outlook-COM-objecten zijn STA-gebonden,
                // aanroepen vanaf een threadpool-thread (Task.Run) kan RPC_E_WRONG_THREAD geven.
                _outlookService.SendMail(recipient.Email, subject, body, accountName, isHtml, attachments);
                results.Add(new SendResult { Email = recipient.Email, DisplayName = recipient.DisplayName, Success = true });
                LogSend(recipient.DisplayName, recipient.Email, true);
            }
            catch (Exception ex)
            {
                results.Add(new SendResult
                {
                    Email = recipient.Email,
                    DisplayName = recipient.DisplayName,
                    Success = false,
                    Error = ex.Message,
                });
                LogSend(recipient.DisplayName, recipient.Email, false, ex.Message);
            }

            SendProgressBar.Value = i + 1;
            StatusText.Text = $"{i + 1} / {count} verwerkt";

            if (i < _imported.Recipients.Count - 1)
            {
                try
                {
                    await Task.Delay(delay, token);
                }
                catch (TaskCanceledException)
                {
                    cancelled = true;
                    Log($"Verzending gestopt door gebruiker na {i + 1} van {count}.");
                    break;
                }
            }
        }

        var succeeded = results.Count(r => r.Success);
        var failed = results.Count(r => !r.Success);

        var directory = Path.GetDirectoryName(_currentFilePath) ?? Environment.CurrentDirectory;
        var reportPath = Path.Combine(directory, $"verzendrapport_{DateTime.Now:yyyyMMdd_HHmmss}.csv");

        try
        {
            SendReportService.Write(reportPath, results);
            Log($"Rapport opgeslagen: {reportPath}");
        }
        catch (Exception ex)
        {
            Log($"Kon rapport niet opslaan: {ex.Message}");
        }

        _sendCts?.Dispose();
        _sendCts = null;
        SetSendingState(false);

        var summary = cancelled
            ? $"Gestopt. {succeeded} verstuurd, {failed} mislukt, rest overgeslagen.\nRapport: {reportPath}"
            : $"Klaar. {succeeded} verstuurd, {failed} mislukt.\nRapport: {reportPath}";
        MessageBox.Show(this, summary, cancelled ? "Verzending gestopt" : "Verzenden voltooid",
            MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void CancelSendButton_Click(object sender, RoutedEventArgs e)
    {
        _sendCts?.Cancel();
        CancelSendButton.IsEnabled = false;
        CancelSendButton.Content = "Stoppen...";
    }

    private static double ParseDelaySeconds(string text)
    {
        var normalized = text.Trim().Replace(',', '.');
        if (double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out var seconds) && seconds >= 0)
            return Math.Max(seconds, 0.2);

        return 1;
    }

    private void AddAttachmentButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Multiselect = true, Filter = "Alle bestanden (*.*)|*.*" };
        if (dialog.ShowDialog() != true)
            return;

        foreach (var path in dialog.FileNames)
        {
            if (!_attachmentPaths.Contains(path))
                _attachmentPaths.Add(path);
        }
    }

    private void RemoveAttachmentButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: string path })
            _attachmentPaths.Remove(path);
    }

    private void ChooseTemplateFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog
        {
            InitialDirectory = _templatesFolder,
        };
        if (dialog.ShowDialog(this) != true || dialog.FolderName is not { } folder)
            return;

        _templatesFolder = folder;
        TemplateFolderText.Text = folder;
        TemplateFolderText.Foreground = (System.Windows.Media.Brush)FindResource("InkBrush");
        AppSettingsService.Save(new AppSettings { TemplatesFolder = folder });
        RefreshTemplateList();
    }

    private void RefreshTemplateList()
    {
        TemplateComboBox.ItemsSource = _templatesFolder is null
            ? null
            : TemplateLibraryService.ListTemplates(_templatesFolder);
        TemplateComboBox.SelectedIndex = -1;
    }

    private void TemplateComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (TemplateComboBox.SelectedItem is not TemplateSummary summary)
            return;

        try
        {
            var template = TemplateStorageService.Load(summary.FilePath);
            SubjectTextBox.Text = template.Subject;
            BodyTextBox.Text = template.Body;
            Log($"Sjabloon geladen: {summary.DisplayName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Kon sjabloon niet laden:\n{ex.Message}", "Fout", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SaveTemplateButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog
        {
            Filter = "Sjabloon (*.json)|*.json",
            FileName = SuggestTemplateFileName(SubjectTextBox.Text),
            InitialDirectory = _templatesFolder,
        };
        if (dialog.ShowDialog() != true)
            return;

        try
        {
            var name = string.IsNullOrWhiteSpace(SubjectTextBox.Text)
                ? Path.GetFileNameWithoutExtension(dialog.FileName)
                : SubjectTextBox.Text;
            TemplateStorageService.Save(dialog.FileName, new MailTemplate { Name = name, Subject = SubjectTextBox.Text, Body = BodyTextBox.Text });
            Log($"Sjabloon opgeslagen: {dialog.FileName}");

            if (_templatesFolder is not null && string.Equals(Path.GetDirectoryName(dialog.FileName), _templatesFolder, StringComparison.OrdinalIgnoreCase))
                RefreshTemplateList();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Kon sjabloon niet opslaan:\n{ex.Message}", "Fout", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static string SuggestTemplateFileName(string subject)
    {
        var name = string.IsNullOrWhiteSpace(subject) ? "standaardmail" : subject;
        foreach (var invalid in Path.GetInvalidFileNameChars())
            name = name.Replace(invalid, '-');
        return name;
    }

    private void RefreshAccountsButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var accounts = _outlookService.GetAccounts();
            var fnvAccounts = accounts.Where(a => IsFnvAddress(a.EmailAddress)).ToList();

            var items = new List<string>();
            var defaultEmail = _outlookService.GetAccountEmail(null);
            if (IsFnvAddress(defaultEmail))
                items.Add(DefaultAccountLabel);
            items.AddRange(fnvAccounts.Select(a => a.DisplayName));

            _accountsLoaded = true;
            AccountComboBox.ItemsSource = items;
            AccountComboBox.SelectedIndex = items.Count > 0 ? 0 : -1;

            if (items.Count == 0)
            {
                MessageBox.Show(this,
                    $"Geen account gevonden dat eindigt op {RequiredEmailDomain}. Voeg je FNV-account toe aan Outlook en klik daarna opnieuw op 'Accounts vernieuwen'.",
                    "Geen FNV-account gevonden", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                Log($"{items.Count} FNV-account(s) gevonden in Outlook.");
            }

            TryAutoFillTestRecipient();
        }
        catch (OutlookNotAvailableException ex)
        {
            MessageBox.Show(this, ex.Message, "Outlook niet beschikbaar", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Kon accounts niet ophalen:\n{ex.Message}", "Fout", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void AccountComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e) => TryAutoFillTestRecipient();

    /// <summary>
    /// Vult de testontvanger met het e-mailadres van het huidig gekozen verzendaccount, tenzij
    /// de gebruiker daar zelf al iets anders heeft ingetypt (zie <see cref="TestRecipientTextBox_TextChanged"/>).
    /// </summary>
    private void TryAutoFillTestRecipient()
    {
        // Pas ná "Accounts vernieuwen": anders zou het opvragen van het adres bij het opstarten
        // van de app Outlook op de achtergrond opstarten, wat niemand verwacht.
        if (_testRecipientIsCustom || !_accountsLoaded)
            return;

        try
        {
            var email = _outlookService.GetAccountEmail(SelectedAccountName);
            _suppressTestRecipientTracking = true;
            TestRecipientTextBox.Text = email;
            _suppressTestRecipientTracking = false;
        }
        catch
        {
            // Outlook nog niet beschikbaar; testontvanger blijft zoals hij is, gebruiker kan hem handmatig invullen.
        }
    }

    private void TestRecipientTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppressTestRecipientTracking)
            return;

        // Alleen als "eigen" beschouwen zolang er iets in staat; leegmaken zet 'm terug op automatisch.
        _testRecipientIsCustom = !string.IsNullOrWhiteSpace(TestRecipientTextBox.Text);
    }

    /// <param name="canCancel">
    /// Alleen bij een bulkverzending is er iets om af te breken; bij één testmail heeft een
    /// stopknop geen effect en moet hij dus ook niet in beeld komen.
    /// </param>
    private void SetSendingState(bool sending, bool canCancel = false)
    {
        _isSending = sending;
        SendTestButton.IsEnabled = !sending;
        SendAllButton.IsEnabled = !sending;
        ChooseFileButton.IsEnabled = !sending;
        EmailColumnComboBox.IsEnabled = !sending;
        AccountComboBox.IsEnabled = !sending;
        RefreshAccountsButton.IsEnabled = !sending;
        TestRecipientTextBox.IsEnabled = !sending;
        DelayTextBox.IsEnabled = !sending;
        HtmlFormattingCheckBox.IsEnabled = !sending;
        AddAttachmentButton.IsEnabled = !sending;

        // Sjabloon vastzetten tijdens het verzenden, zodat wat je ziet ook is wat er de deur uitgaat.
        SubjectTextBox.IsEnabled = !sending;
        BodyTextBox.IsEnabled = !sending;
        TemplateComboBox.IsEnabled = !sending;
        ChooseTemplateFolderButton.IsEnabled = !sending;
        SaveTemplateButton.IsEnabled = !sending;
        PlaceholderPanel.IsEnabled = !sending;

        CancelSendButton.Visibility = sending && canCancel ? Visibility.Visible : Visibility.Collapsed;
        CancelSendButton.IsEnabled = true;
        CancelSendButton.Content = "Stoppen";
    }

    private void Log(string message) => AddLogEntry(new LogEntry(NowStamp(), message, null, null));

    private void LogSend(string title, string email, bool success, string? error = null)
    {
        var detail = success ? email : $"{email} — {error}";
        AddLogEntry(new LogEntry(NowStamp(), title, detail, success));
    }

    private void AddLogEntry(LogEntry entry)
    {
        _logEntries.Add(entry);
        LogListBox.ScrollIntoView(entry);
    }

    private static string NowStamp() => DateTime.Now.ToString("HH:mm:ss");

    private static bool IsFnvAddress(string email) =>
        !string.IsNullOrWhiteSpace(email) && email.EndsWith(RequiredEmailDomain, StringComparison.OrdinalIgnoreCase);

    private sealed record LogEntry(string Time, string Title, string? Detail, bool? Success)
    {
        public bool HasDetail => !string.IsNullOrEmpty(Detail);

        public bool IsSuccess => Success == true;

        public bool IsFailure => Success == false;
    }
}

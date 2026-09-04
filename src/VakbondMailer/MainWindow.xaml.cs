using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
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

    private string? _currentFilePath;
    private string? _templatesFolder;
    private ImportedRecipients? _imported;
    private TextBox? _lastFocusedTemplateBox;
    private bool _isSending;

    private string? SelectedAccountName =>
        AccountComboBox.SelectedItem is string name && name != DefaultAccountLabel ? name : null;

    public MainWindow()
    {
        InitializeComponent();
        _lastFocusedTemplateBox = BodyTextBox;
        LogItemsControl.ItemsSource = _logEntries;
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
            MessageBox.Show(this, $"Kon het bestand niet lezen:\n{ex.Message}", "Fout bij inlezen",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
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

            UpdateMergedPreview();
        }
        catch (Exception ex)
        {
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

    private void Template_TextChanged(object sender, TextChangedEventArgs e) => UpdateMergedPreview();

    private void UpdateMergedPreview()
    {
        var sampleRecipient = _imported?.Recipients.FirstOrDefault();
        if (sampleRecipient is null)
        {
            MergedPreviewSubject.Text = "(nog geen lijst geladen)";
            MergedPreviewBody.Text = "Laad een CSV of Excel-bestand om een echt voorbeeld te zien.";
            return;
        }

        MergedPreviewSubject.Text = TemplateRenderer.Render(SubjectTextBox.Text, sampleRecipient);
        MergedPreviewBody.Text = TemplateRenderer.Render(BodyTextBox.Text, sampleRecipient);
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

            var sampleRecipient = _imported?.Recipients.FirstOrDefault()
                ?? new Recipient { Email = testRecipient, Fields = new Dictionary<string, string>() };

            var subject = TemplateRenderer.Render(SubjectTextBox.Text, sampleRecipient);
            var body = TemplateRenderer.Render(BodyTextBox.Text, sampleRecipient);

            _outlookService.SendMail(testRecipient, $"[TEST] {subject}", body, accountName);
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

        var count = _imported.Recipients.Count;
        var confirm = MessageBox.Show(this,
            $"Weet je zeker dat je deze mail wilt versturen naar {count} ontvanger(s), vanaf {senderEmail}?",
            "Bevestig verzenden", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
            return;

        SetSendingState(true);
        SendProgressBar.Maximum = count;
        SendProgressBar.Value = 0;
        _logEntries.Clear();

        var results = new List<SendResult>();
        for (var i = 0; i < _imported.Recipients.Count; i++)
        {
            var recipient = _imported.Recipients[i];
            var subject = TemplateRenderer.Render(SubjectTextBox.Text, recipient);
            var body = TemplateRenderer.Render(BodyTextBox.Text, recipient);

            try
            {
                // Bewust synchroon op de UI-thread: Outlook-COM-objecten zijn STA-gebonden,
                // aanroepen vanaf een threadpool-thread (Task.Run) kan RPC_E_WRONG_THREAD geven.
                _outlookService.SendMail(recipient.Email, subject, body, accountName);
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
                await Task.Delay(1000);
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

        SetSendingState(false);
        MessageBox.Show(this,
            $"Klaar. {succeeded} verstuurd, {failed} mislukt.\nRapport: {reportPath}",
            "Verzenden voltooid", MessageBoxButton.OK, MessageBoxImage.Information);
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

    private void SetSendingState(bool sending)
    {
        _isSending = sending;
        SendTestButton.IsEnabled = !sending;
        SendAllButton.IsEnabled = !sending;
        ChooseFileButton.IsEnabled = !sending;
        EmailColumnComboBox.IsEnabled = !sending;
        AccountComboBox.IsEnabled = !sending;
        RefreshAccountsButton.IsEnabled = !sending;
        TestRecipientTextBox.IsEnabled = !sending;
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
        LogScrollViewer.ScrollToEnd();
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

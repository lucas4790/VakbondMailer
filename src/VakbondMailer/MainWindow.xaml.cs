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
    private readonly OutlookMailService _outlookService = new();

    private string? _currentFilePath;
    private ImportedRecipients? _imported;
    private TextBox? _lastFocusedTemplateBox;
    private bool _isSending;

    public MainWindow()
    {
        InitializeComponent();
        _lastFocusedTemplateBox = BodyTextBox;
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

            var headers = RecipientImportService.ReadHeaders(_currentFilePath);
            EmailColumnComboBox.ItemsSource = headers;
            var guessed = RecipientImportService.GuessEmailColumn(headers);
            EmailColumnComboBox.SelectedItem = guessed ?? headers.FirstOrDefault();

            PlaceholderComboBox.ItemsSource = headers;

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

    private void SubjectTextBox_GotFocus(object sender, RoutedEventArgs e) => _lastFocusedTemplateBox = SubjectTextBox;

    private void BodyTextBox_GotFocus(object sender, RoutedEventArgs e) => _lastFocusedTemplateBox = BodyTextBox;

    private void InsertPlaceholderButton_Click(object sender, RoutedEventArgs e)
    {
        if (PlaceholderComboBox.SelectedItem is not string header)
            return;

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
            MergedPreviewTextBox.Text = "(nog geen lijst geladen)";
            return;
        }

        var subject = TemplateRenderer.Render(SubjectTextBox.Text, sampleRecipient);
        var body = TemplateRenderer.Render(BodyTextBox.Text, sampleRecipient);
        MergedPreviewTextBox.Text = $"Onderwerp: {subject}\n\n{body}";
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
            var myEmail = _outlookService.GetCurrentUserEmail();
            var sampleRecipient = _imported?.Recipients.FirstOrDefault()
                ?? new Recipient { Email = myEmail, Fields = new Dictionary<string, string>() };

            var subject = TemplateRenderer.Render(SubjectTextBox.Text, sampleRecipient);
            var body = TemplateRenderer.Render(BodyTextBox.Text, sampleRecipient);

            _outlookService.SendMail(myEmail, $"[TEST] {subject}", body);
            await Task.Yield();

            Log($"Testmail verstuurd naar {myEmail}");
            MessageBox.Show(this, $"Testmail verstuurd naar {myEmail}.", "Gelukt",
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

        var count = _imported.Recipients.Count;
        var confirm = MessageBox.Show(this,
            $"Weet je zeker dat je deze mail wilt versturen naar {count} ontvanger(s)?",
            "Bevestig verzenden", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
            return;

        SetSendingState(true);
        SendProgressBar.Maximum = count;
        SendProgressBar.Value = 0;
        LogListBox.Items.Clear();

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
                _outlookService.SendMail(recipient.Email, subject, body);
                results.Add(new SendResult { Email = recipient.Email, DisplayName = recipient.DisplayName, Success = true });
                Log($"Verstuurd: {recipient.DisplayName} <{recipient.Email}>");
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
                Log($"MISLUKT: {recipient.DisplayName} <{recipient.Email}> - {ex.Message}");
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

    private void SaveTemplateButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog { Filter = "Sjabloon (*.json)|*.json", FileName = "standaardmail.json" };
        if (dialog.ShowDialog() != true)
            return;

        try
        {
            TemplateStorageService.Save(dialog.FileName, new MailTemplate { Subject = SubjectTextBox.Text, Body = BodyTextBox.Text });
            Log($"Sjabloon opgeslagen: {dialog.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Kon sjabloon niet opslaan:\n{ex.Message}", "Fout", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void LoadTemplateButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "Sjabloon (*.json)|*.json" };
        if (dialog.ShowDialog() != true)
            return;

        try
        {
            var template = TemplateStorageService.Load(dialog.FileName);
            SubjectTextBox.Text = template.Subject;
            BodyTextBox.Text = template.Body;
            Log($"Sjabloon geladen: {dialog.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Kon sjabloon niet laden:\n{ex.Message}", "Fout", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SetSendingState(bool sending)
    {
        _isSending = sending;
        SendTestButton.IsEnabled = !sending;
        SendAllButton.IsEnabled = !sending;
        ChooseFileButton.IsEnabled = !sending;
        EmailColumnComboBox.IsEnabled = !sending;
    }

    private void Log(string message)
    {
        LogListBox.Items.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
        if (LogListBox.Items.Count > 0)
            LogListBox.ScrollIntoView(LogListBox.Items[^1]);
    }
}

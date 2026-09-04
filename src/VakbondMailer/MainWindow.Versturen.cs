using System.Linq;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using VakbondMailer.Models;
using VakbondMailer.Services;

namespace VakbondMailer;

/// <summary> Stap 3: accounts, testmail, bulkverzending, bijlagen en het logboek.</summary>
public partial class MainWindow
{
    private bool TryPrepareSend(out BulkSendOptions options, out string senderEmail)
    {
        options = null!;
        senderEmail = string.Empty;

        var accountName = SelectedAccountName;
        try
        {
            senderEmail = _outlookService.GetAccountEmail(accountName);
        }
        catch (OutlookNotAvailableException ex)
        {
            MessageBox.Show(this, ex.Message, "Outlook niet beschikbaar", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }

        if (!SendSettings.IsAllowedSender(senderEmail))
        {
            MessageBox.Show(this,
                $"Dit account ({senderEmail}) is geen FNV-adres. Kies via 'Accounts vernieuwen' een account dat eindigt op {SendSettings.RequiredEmailDomain}.",
                "Alleen FNV-adressen toegestaan", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        var attachments = _attachmentPaths.ToList();
        var missingAttachment = attachments.FirstOrDefault(path => !File.Exists(path));
        if (missingAttachment is not null)
        {
            MessageBox.Show(this,
                $"Deze bijlage bestaat niet meer:\n{missingAttachment}\n\nVerwijder hem uit de lijst of kies het bestand opnieuw.",
                "Bijlage niet gevonden", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        options = new BulkSendOptions
        {
            SubjectTemplate = SubjectTextBox.Text,
            BodyTemplate = BodyTextBox.Text,
            PlanningFields = CurrentPlanningFields,
            IsHtml = HtmlFormattingCheckBox.IsChecked == true,
            AttachmentPaths = attachments,
            AccountName = accountName,
            DelayBetweenMails = TimeSpan.FromSeconds(SendSettings.ParseDelaySeconds(DelayTextBox.Text)),
        };
        return true;
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
            if (!TryPrepareSend(out var options, out var myEmail))
                return;

            var testRecipient = string.IsNullOrWhiteSpace(TestRecipientTextBox.Text)
                ? myEmail
                : TestRecipientTextBox.Text.Trim();

            var sampleRecipient = GetPreviewRecipient()
                ?? new Recipient { Email = testRecipient, Fields = new Dictionary<string, string>() };

            var subject = TemplateRenderer.Render(options.SubjectTemplate, sampleRecipient, options.PlanningFields);
            var body = BulkMailSender.ComposeBody(options, sampleRecipient);

            _outlookService.SendMail(testRecipient, $"[TEST] {subject}", body, options.AccountName, options.IsHtml, options.AttachmentPaths);
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

        var recipients = GetSelectedRecipients();
        if (recipients.Count == 0)
        {
            MessageBox.Show(this, "Er staat geen enkele ontvanger aangevinkt in de lijst.", "Niemand geselecteerd",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        await SendToRecipientsAsync(recipients);
    }

    private async void RetryFailedButton_Click(object sender, RoutedEventArgs e)
    {
        if (_isSending || _lastFailedRecipients.Count == 0)
            return;

        var failed = _lastFailedRecipients.ToList();
        SelectOnly(failed);
        await SendToRecipientsAsync(failed);
    }

    private async Task SendToRecipientsAsync(IReadOnlyList<Recipient> recipients)
    {
        if (string.IsNullOrWhiteSpace(SubjectTextBox.Text) || string.IsNullOrWhiteSpace(BodyTextBox.Text))
        {
            MessageBox.Show(this, "Vul eerst een onderwerp en tekst in.", "Sjabloon leeg",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!TryPrepareSend(out var options, out var senderEmail))
            return;

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

        var count = recipients.Count;

        var alreadySent = SendHistoryService.CountRecentlySent(
            SendHistoryService.DefaultPath,
            options.SubjectTemplate,
            recipients.Select(r => r.Email),
            SendSettings.DuplicateSendWindow,
            DateTime.Now);

        if (alreadySent > 0)
        {
            var confirmDuplicate = MessageBox.Show(this,
                $"Deze mail is de afgelopen {SendSettings.DuplicateSendWindow.TotalDays:0} dagen al naar {alreadySent} van deze {count} ontvanger(s) gestuurd.\n\nToch (nogmaals) versturen?",
                "Mogelijk dubbel versturen", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (confirmDuplicate != MessageBoxResult.Yes)
                return;
        }

        var confirm = MessageBox.Show(this,
            $"Weet je zeker dat je deze mail wilt versturen naar {count} ontvanger(s), vanaf {senderEmail}?",
            "Bevestig verzenden", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
            return;

        _sendCts = new CancellationTokenSource();

        SetSendingState(true, canCancel: true);
        SendProgressBar.Maximum = count;
        SendProgressBar.Value = 0;
        _log.Clear();

        // Bewust op de UI-thread (geen Task.Run): Outlook-COM-objecten zijn STA-gebonden.
        var outcome = await BulkMailSender.SendAsync(
            _outlookService, recipients, options, OnSendProgress, _sendCts.Token);

        if (outcome.Cancelled)
            Log($"Verzending gestopt door gebruiker na {outcome.Results.Count} van {count}.");

        var results = outcome.Results;
        var succeeded = results.Count(r => r.Success);
        var failed = results.Count(r => !r.Success);
        var cancelled = outcome.Cancelled;

        _lastFailedRecipients = outcome.Failed.ToList();
        RetryFailedButton.Content = $"Mislukte opnieuw ({outcome.Failed.Count})";
        RetryFailedButton.Visibility = outcome.Failed.Count > 0 ? Visibility.Visible : Visibility.Collapsed;

        if (outcome.SentEmails.Count > 0)
        {
            try
            {
                SendHistoryService.Append(SendHistoryService.DefaultPath, options.SubjectTemplate, outcome.SentEmails, DateTime.Now);
            }
            catch (Exception ex)
            {
                Log($"Kon verzendgeschiedenis niet bijwerken: {ex.Message}");
            }
        }

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

    private void OnSendProgress(BulkSendProgress progress)
    {
        LogSend(progress.Result.DisplayName, progress.Result.Email, progress.Result.Success, progress.Result.Error);
        SendProgressBar.Value = progress.Processed;
        StatusText.Text = $"{progress.Processed} / {progress.Total} verwerkt";
    }

    private void CancelSendButton_Click(object sender, RoutedEventArgs e)
    {
        _sendCts?.Cancel();
        CancelSendButton.IsEnabled = false;
        CancelSendButton.Content = "Stoppen...";
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

    private void RefreshAccountsButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var accounts = _outlookService.GetAccounts();
            var fnvAccounts = accounts.Where(a => SendSettings.IsAllowedSender(a.EmailAddress)).ToList();

            var items = new List<string>();
            var defaultEmail = _outlookService.GetAccountEmail(null);
            if (SendSettings.IsAllowedSender(defaultEmail))
                items.Add(DefaultAccountLabel);
            items.AddRange(fnvAccounts.Select(a => a.DisplayName));

            _accountsLoaded = true;
            AccountComboBox.ItemsSource = items;
            AccountComboBox.SelectedIndex = items.Count > 0 ? 0 : -1;

            if (items.Count == 0)
            {
                MessageBox.Show(this,
                    $"Geen account gevonden dat eindigt op {SendSettings.RequiredEmailDomain}. Voeg je FNV-account toe aan Outlook en klik daarna opnieuw op 'Accounts vernieuwen'.",
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
        SelectAllButton.IsEnabled = !sending;
        SelectNoneButton.IsEnabled = !sending;
        RetryFailedButton.IsEnabled = !sending;
        PreviewDataGrid.IsEnabled = !sending;

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

    private void Log(string message) => ShowInLog(_log.Add(message));

    private void LogSend(string title, string email, bool success, string? error = null) =>
        ShowInLog(_log.AddSend(title, email, success, error));

    /// <summary>Laat de nieuwste regel meteen in beeld komen.</summary>
    private void ShowInLog(LogEntry entry) => LogListBox.ScrollIntoView(entry);
}

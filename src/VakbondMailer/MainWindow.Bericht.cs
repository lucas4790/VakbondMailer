using System.Linq;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using VakbondMailer.Models;
using VakbondMailer.Services;

namespace VakbondMailer;

/// <summary> Stap 2: het bericht — sjablonen, velden, gastles-planning en het live voorbeeld.</summary>
public partial class MainWindow
{
    private void RenderPlaceholderChips(IReadOnlyList<string> headers)
    {
        PlaceholderPanel.Children.Clear();

        // Kolommen uit de ledenlijst, plus de planningsvelden die de app zelf invult.
        foreach (var field in headers.Concat(PlanningFields.Keys))
        {
            var button = new Button
            {
                Content = $"{{{{{field}}}}}",
                Style = (Style)FindResource("ChipButton"),
            };
            button.Click += (_, _) => InsertPlaceholder(field);
            PlaceholderPanel.Children.Add(button);
        }
    }

    /// <summary>
    /// Vult de maandkeuze met de komende twaalf maanden; standaard de volgende maand, want een
    /// gastles plan je zelden nog voor deze maand in.
    /// </summary>
    private void InitializeMonthOptions()
    {
        var options = PlanningFields.NextMonths(12, DateTime.Today)
            .Select(month => new MonthOption(month, PlanningFields.FormatMonthYear(month)))
            .ToList();

        MonthComboBox.ItemsSource = options;
        MonthComboBox.SelectedIndex = options.Count > 1 ? 1 : 0;
    }

    private DateTime SelectedMonth =>
        MonthComboBox.SelectedItem is MonthOption option ? option.Value : DateTime.Today;

    private IReadOnlyDictionary<string, string> CurrentPlanningFields =>
        PlanningFields.Build(SelectedMonth, ProposalCalendar.SelectedDates.Cast<DateTime>());

    private void MonthComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        // Datums uit een andere maand horen niet bij dit voorstel, dus die vervallen.
        ProposalCalendar.SelectedDates.Clear();

        var month = SelectedMonth;
        var lastDay = month.AddMonths(1).AddDays(-1);
        ProposalCalendar.DisplayDateStart = month;
        ProposalCalendar.DisplayDateEnd = lastDay;
        ProposalCalendar.DisplayDate = month;

        BlockOutDatesYouWouldNeverPropose(month, lastDay);
        UpdatePlanningSummary();
    }

    /// <summary>
    /// Weekenden en dagen die al voorbij zijn kun je een docent niet voorstellen, dus die
    /// zijn niet aan te klikken.
    /// </summary>
    private void BlockOutDatesYouWouldNeverPropose(DateTime firstDay, DateTime lastDay)
    {
        ProposalCalendar.BlackoutDates.Clear();

        for (var day = firstDay; day <= lastDay; day = day.AddDays(1))
        {
            if (day.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday || day.Date < DateTime.Today)
                ProposalCalendar.BlackoutDates.Add(new CalendarDateRange(day));
        }
    }

    private void ProposalCalendar_SelectedDatesChanged(object sender, SelectionChangedEventArgs e) => UpdatePlanningSummary();

    private void ClearDatesButton_Click(object sender, RoutedEventArgs e)
    {
        ProposalCalendar.SelectedDates.Clear();
        UpdatePlanningSummary();
    }

    private void UpdatePlanningSummary()
    {
        var dateOptions = PlanningFields.FormatDateOptions(ProposalCalendar.SelectedDates.Cast<DateTime>());
        DateOptionsPreviewText.Text = string.IsNullOrEmpty(dateOptions)
            ? "Geen datums gekozen — {{Datumopties}} blijft leeg. Klik datums in de kalender om ze voor te stellen."
            : $"{{{{Datumopties}}}} wordt: {dateOptions}";

        UpdateMergedPreview();
        UpdatePlaceholderWarning();
    }

    private sealed record MonthOption(DateTime Value, string Label);

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
    private Recipient? GetPreviewRecipient() => _selection?.At(PreviewDataGrid.SelectedIndex);

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

        var planning = CurrentPlanningFields;
        PreviewRecipientLabel.Text = $"Voorbeeld van: {recipient.DisplayName} ({recipient.Email})";
        MergedPreviewSubject.Text = TemplateRenderer.Render(SubjectTextBox.Text, recipient, planning);
        MergedPreviewBody.Text = TemplateRenderer.Render(BodyTextBox.Text, recipient, planning);
    }

    private IReadOnlyList<string> FindUnknownPlaceholders() =>
        _imported is null
            ? Array.Empty<string>()
            : TemplateRenderer.FindUnknownPlaceholders(
                SubjectTextBox.Text, BodyTextBox.Text, _imported.Headers.Concat(PlanningFields.Keys));

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

    /// <summary>
    /// De controles die voor elke verzending gelden (welk account, mag dat account, bestaan de
    /// bijlagen nog) plus wat er per mail ingevuld moet worden. Eén plek, zodat de testmail en
    /// de echte verzending niet uit elkaar kunnen gaan lopen.

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
            FileName = TemplateStorageService.SuggestFileName(SubjectTextBox.Text),
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
}

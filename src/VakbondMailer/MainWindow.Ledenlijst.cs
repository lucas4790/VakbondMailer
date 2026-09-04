using System.Linq;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using VakbondMailer.Models;
using VakbondMailer.Services;

namespace VakbondMailer;

/// <summary> Stap 1: de ledenlijst inlezen, tonen en aanvinken wie mail krijgt.</summary>
public partial class MainWindow
{
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
        _selection = null;
        _lastFailedRecipients = new List<Recipient>();
        RetryFailedButton.Visibility = Visibility.Collapsed;
        PreviewDataGrid.ItemsSource = null;
        EmailColumnComboBox.ItemsSource = null;
        RenderPlaceholderChips(Array.Empty<string>()); // planningsvelden blijven wel bruikbaar
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
            _lastFailedRecipients = new List<Recipient>();
            RetryFailedButton.Visibility = Visibility.Collapsed;

            _selection = RecipientSelection.From(_imported);
            _selection.Table.ColumnChanged += (_, args) =>
            {
                if (args.Column?.ColumnName == RecipientSelection.SelectionColumnName)
                    UpdateRecipientCount();
            };
            PreviewDataGrid.ItemsSource = _selection.Table.DefaultView;
            UpdateRecipientCount();

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
            _selection = null;
            PreviewDataGrid.ItemsSource = null;
            RecipientCountText.Text = "Geen lijst geladen";
            UpdateMergedPreview();
            UpdatePlaceholderWarning();

            MessageBox.Show(this, $"Kon de lijst niet inladen:\n{ex.Message}", "Fout bij inlezen",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void PreviewDataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
    {
        // Alleen het vinkje is bewerkbaar; de gegevens uit de ledenlijst blijven kijkwerk.
        var isSelectionColumn = e.PropertyName == RecipientSelection.SelectionColumnName;
        e.Column.IsReadOnly = !isSelectionColumn;

        if (isSelectionColumn)
        {
            e.Column.Header = "✓";
            e.Column.CanUserResize = false;
        }
    }

    private IReadOnlyList<Recipient> GetSelectedRecipients() =>
        _selection?.Selected ?? Array.Empty<Recipient>();

    private void SelectAllButton_Click(object sender, RoutedEventArgs e) => SetAllSelected(true);

    private void SelectNoneButton_Click(object sender, RoutedEventArgs e) => SetAllSelected(false);

    private void SetAllSelected(bool selected)
    {
        if (_selection is null)
            return;

        // Een openstaand vinkje moet eerst vastgelegd zijn, anders overschrijft de grid het weer.
        PreviewDataGrid.CommitEdit(DataGridEditingUnit.Row, exitEditingMode: true);
        _selection.SetAll(selected);
        UpdateRecipientCount();
    }

    private void SelectOnly(IReadOnlyCollection<Recipient> recipients)
    {
        if (_selection is null)
            return;

        PreviewDataGrid.CommitEdit(DataGridEditingUnit.Row, exitEditingMode: true);
        _selection.SelectOnly(recipients);
        UpdateRecipientCount();
    }

    private void UpdateRecipientCount() =>
        RecipientCountText.Text = _selection?.CountLabel ?? "Geen lijst geladen";
}

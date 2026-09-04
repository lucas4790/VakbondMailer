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

    private readonly OutlookMailService _outlookService = new();
    private readonly SendLog _log = new();
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
    private RecipientSelection? _selection;
    private List<Recipient> _lastFailedRecipients = new();

    private string? SelectedAccountName =>
        AccountComboBox.SelectedItem is string name && name != DefaultAccountLabel ? name : null;

    /// <summary>
    /// Zet een vinkje op de stappenbalk zodra een stap af is, zodat die balk iets zegt in plaats
    /// van alleen versiering te zijn.
    /// </summary>
    private void UpdateStepTracker()
    {
        MarkStep(StepOneBadgeBorder, StepOneBadge, "1", _selection is { All.Count: > 0 });

        MarkStep(StepTwoBadgeBorder, StepTwoBadge, "2",
            !string.IsNullOrWhiteSpace(SubjectTextBox.Text) && !string.IsNullOrWhiteSpace(BodyTextBox.Text));

        MarkStep(StepThreeBadgeBorder, StepThreeBadge, "3", _accountsLoaded && AccountComboBox.SelectedItem is not null);
    }

    private void MarkStep(Border badge, TextBlock label, string nummer, bool klaar)
    {
        badge.Background = (System.Windows.Media.Brush)FindResource(klaar ? "GoodBrush" : "AccentBrush");
        label.Text = klaar ? "\u2713" : nummer;
    }

    /// <summary>
    /// Laat het venster meegroeien met het scherm: op een grote monitor hoeft er dan veel minder
    /// gescrold te worden tussen de drie stappen. Wel begrensd, want de inhoud staat gecentreerd
    /// en wordt van eindeloze breedte niet beter leesbaar.
    /// </summary>
    private void SizeToScreen()
    {
        var beschikbaar = SystemParameters.WorkArea;

        Width = Math.Clamp(beschikbaar.Width * 0.6, MinWidth, 1500);
        Height = Math.Clamp(beschikbaar.Height * 0.9, MinHeight, 1300);
    }

    public MainWindow()
    {
        InitializeComponent();
        SizeToScreen();
        _lastFocusedTemplateBox = BodyTextBox;
        LogListBox.ItemsSource = _log.Entries;
        AttachmentListControl.ItemsSource = _attachmentPaths;
        InitializeMonthOptions();
        RenderPlaceholderChips(Array.Empty<string>());
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
}

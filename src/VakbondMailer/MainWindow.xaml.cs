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

    public MainWindow()
    {
        InitializeComponent();
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

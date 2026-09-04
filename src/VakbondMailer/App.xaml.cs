using System.Globalization;
using System.Threading;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;

namespace VakbondMailer;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        UseDutchEverywhere();
        UseFnvAccentColor();

        // Op een machine zonder actieve schermsessie levert een schermopname van een
        // GPU-gerenderd WPF-venster een leeg beeld op. Met VAKBONDMAILER_SOFTWARE_RENDERING=1
        // rendert WPF via software en is het venster wél vast te leggen (zie
        // scripts/Maak-Screenshot.ps1). Voor normaal gebruik verandert er niets.
        if (Environment.GetEnvironmentVariable("VAKBONDMAILER_SOFTWARE_RENDERING") == "1")
            RenderOptions.ProcessRenderMode = RenderMode.SoftwareOnly;
    }

    /// <summary>
    /// De Fluent-theme kleurt selecties, focusranden en de kalender standaard met de
    /// Windows-accentkleur van de gebruiker. Hier moet dat het FNV-blauw zijn, zodat het scherm
    /// één geheel blijft op welke computer de app ook draait.
    /// </summary>
    private void UseFnvAccentColor()
    {
        var fnvBlauw = (Color)ColorConverter.ConvertFromString("#009CDE");

        Resources["SystemAccentColor"] = fnvBlauw;
        Resources["SystemAccentColorPrimary"] = fnvBlauw;
        Resources["SystemAccentColorSecondary"] = fnvBlauw;
        Resources["SystemAccentColorTertiary"] = fnvBlauw;
        Resources["AccentTextFillColorPrimaryBrush"] = new SolidColorBrush(fnvBlauw);
        Resources["AccentFillColorDefaultBrush"] = new SolidColorBrush(fnvBlauw);
    }

    /// <summary>
    /// De app is volledig Nederlandstalig, maar besturingselementen zoals de kalender volgen
    /// standaard de Windows-taalinstelling. Dat zou "October" naast "oktober" opleveren op een
    /// Engelstalige Windows, dus zetten we de taal expliciet vast.
    /// </summary>
    private static void UseDutchEverywhere()
    {
        var dutch = new CultureInfo("nl-NL");

        CultureInfo.DefaultThreadCurrentCulture = dutch;
        CultureInfo.DefaultThreadCurrentUICulture = dutch;
        Thread.CurrentThread.CurrentCulture = dutch;
        Thread.CurrentThread.CurrentUICulture = dutch;

        FrameworkElement.LanguageProperty.OverrideMetadata(
            typeof(FrameworkElement),
            new FrameworkPropertyMetadata(XmlLanguage.GetLanguage(dutch.IetfLanguageTag)));
    }
}

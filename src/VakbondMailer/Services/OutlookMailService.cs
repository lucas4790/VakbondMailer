using System.Runtime.InteropServices;

namespace VakbondMailer.Services;

public sealed class OutlookNotAvailableException : Exception
{
    public OutlookNotAvailableException(string message, Exception? inner = null) : base(message, inner)
    {
    }
}

/// <summary>
/// Verstuurt mail via de al-lopende, al-ingelogde klassieke Outlook desktop-app (late-bound COM),
/// zodat er geen Azure-app-registration of opgeslagen wachtwoord nodig is.
/// </summary>
public sealed class OutlookMailService
{
    private const int OlMailItem = 0;

    // Marshal.GetActiveObject bestaat alleen op .NET Framework; op .NET (Core) moet de
    // Running Object Table zelf via P/Invoke aangesproken worden.
    [DllImport("ole32.dll")]
    private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid clsid);

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    public void SendMail(string toEmail, string subject, string body)
    {
        dynamic outlookApp = GetOrCreateOutlookApplication();
        dynamic mailItem = outlookApp.CreateItem(OlMailItem);
        try
        {
            mailItem.To = toEmail;
            mailItem.Subject = subject;
            mailItem.Body = body;
            mailItem.Send();
        }
        finally
        {
            Marshal.ReleaseComObject(mailItem);
        }
    }

    public string GetCurrentUserEmail()
    {
        dynamic outlookApp = GetOrCreateOutlookApplication();
        dynamic session = outlookApp.Session;
        try
        {
            dynamic currentUser = session.CurrentUser;
            try
            {
                dynamic exchangeUser = currentUser.AddressEntry.GetExchangeUser();
                if (exchangeUser is not null)
                    return (string)exchangeUser.PrimarySmtpAddress;
            }
            catch (COMException)
            {
                // Geen Exchange-account gevonden; val terug op het adres van de huidige gebruiker.
            }

            return (string)currentUser.Address;
        }
        finally
        {
            Marshal.ReleaseComObject(session);
        }
    }

    private static object GetOrCreateOutlookApplication()
    {
        try
        {
            return GetRunningOutlookApplication();
        }
        catch (COMException)
        {
            var type = Type.GetTypeFromProgID("Outlook.Application")
                ?? throw new OutlookNotAvailableException(
                    "Outlook is niet gevonden op deze computer. Zorg dat de klassieke Outlook desktop-app is geïnstalleerd (niet 'nieuwe Outlook').");

            try
            {
                return Activator.CreateInstance(type)
                    ?? throw new OutlookNotAvailableException("Kon Outlook niet starten.");
            }
            catch (COMException ex)
            {
                throw new OutlookNotAvailableException(
                    "Kon geen verbinding maken met Outlook. Zorg dat de klassieke Outlook desktop-app (niet 'nieuwe Outlook') open staat en dat je bent ingelogd.", ex);
            }
        }
    }

    private static object GetRunningOutlookApplication()
    {
        var hr = CLSIDFromProgID("Outlook.Application", out var clsid);
        if (hr != 0)
            throw new COMException("Outlook.Application is niet geregistreerd op dit systeem.", hr);

        GetActiveObject(ref clsid, IntPtr.Zero, out var instance);
        return instance;
    }
}

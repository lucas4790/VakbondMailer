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
            return Marshal.GetActiveObject("Outlook.Application");
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
}

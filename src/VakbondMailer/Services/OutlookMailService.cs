using System.Collections.Generic;
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

    /// <param name="accountName">
    /// Weergavenaam van het Outlook-account waarmee verstuurd moet worden (uit <see cref="GetAccountNames"/>).
    /// Bij null/leeg wordt Outlook's eigen standaardaccount gebruikt.
    /// </param>
    public void SendMail(string toEmail, string subject, string body, string? accountName = null)
    {
        dynamic outlookApp = GetOrCreateOutlookApplication();
        dynamic mailItem = outlookApp.CreateItem(OlMailItem);
        try
        {
            mailItem.To = toEmail;
            mailItem.Subject = subject;
            mailItem.Body = body;

            if (!string.IsNullOrWhiteSpace(accountName))
            {
                var account = FindAccount(outlookApp, accountName);
                if (account is not null)
                    mailItem.SendUsingAccount = account;
            }

            mailItem.Send();
        }
        finally
        {
            Marshal.ReleaseComObject(mailItem);
        }
    }

    /// <summary>
    /// Weergavenamen van alle mailaccounts die in deze Outlook-installatie zijn geconfigureerd,
    /// zodat de gebruiker kan kiezen vanaf welk account verstuurd wordt.
    /// </summary>
    public IReadOnlyList<string> GetAccountNames()
    {
        dynamic outlookApp = GetOrCreateOutlookApplication();
        dynamic session = outlookApp.Session;
        try
        {
            var names = new List<string>();
            dynamic accounts = session.Accounts;
            int count = accounts.Count;
            for (var i = 1; i <= count; i++)
            {
                dynamic account = accounts[i];
                names.Add((string)account.DisplayName);
            }

            return names;
        }
        finally
        {
            Marshal.ReleaseComObject(session);
        }
    }

    /// <summary>
    /// E-mailadres van het opgegeven account (of, bij null/leeg of wanneer het adres niet
    /// bepaald kan worden, van Outlook's huidige/standaardgebruiker).
    /// </summary>
    public string GetAccountEmail(string? accountName = null)
    {
        dynamic outlookApp = GetOrCreateOutlookApplication();
        dynamic session = outlookApp.Session;
        try
        {
            if (!string.IsNullOrWhiteSpace(accountName))
            {
                var account = FindAccount(outlookApp, accountName);
                if (account is not null)
                {
                    try
                    {
                        var smtp = (string)((dynamic)account).SmtpAddress;
                        if (!string.IsNullOrWhiteSpace(smtp))
                            return smtp;
                    }
                    catch (COMException)
                    {
                        // SmtpAddress niet beschikbaar voor dit accounttype; val terug op CurrentUser hieronder.
                    }
                }
            }

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

    private static object? FindAccount(dynamic outlookApp, string accountName)
    {
        dynamic session = outlookApp.Session;
        try
        {
            dynamic accounts = session.Accounts;
            int count = accounts.Count;
            for (var i = 1; i <= count; i++)
            {
                dynamic account = accounts[i];
                if (string.Equals((string)account.DisplayName, accountName, StringComparison.OrdinalIgnoreCase))
                    return account;
            }

            return null;
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

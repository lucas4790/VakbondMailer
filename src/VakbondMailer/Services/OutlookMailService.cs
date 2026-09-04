using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace VakbondMailer.Services;

public sealed class OutlookNotAvailableException : Exception
{
    public OutlookNotAvailableException(string message, Exception? inner = null) : base(message, inner)
    {
    }
}

public sealed record OutlookAccount(string DisplayName, string EmailAddress);

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
    /// Weergavenaam van het Outlook-account waarmee verstuurd moet worden (uit <see cref="GetAccounts"/>).
    /// Bij null/leeg wordt Outlook's eigen standaardaccount gebruikt.
    /// </param>
    /// <param name="isHtml">Wanneer true wordt <paramref name="body"/> als HTML geïnterpreteerd (zie <see cref="SimpleHtmlFormatter"/>).</param>
    /// <param name="attachmentPaths">Volledige bestandspaden van bijlagen om mee te sturen.</param>
    public void SendMail(
        string toEmail,
        string subject,
        string body,
        string? accountName = null,
        bool isHtml = false,
        IReadOnlyList<string>? attachmentPaths = null)
    {
        dynamic outlookApp = GetOrCreateOutlookApplication();
        dynamic mailItem = outlookApp.CreateItem(OlMailItem);
        try
        {
            mailItem.To = toEmail;
            mailItem.Subject = subject;

            if (isHtml)
                mailItem.HTMLBody = body;
            else
                mailItem.Body = body;

            if (attachmentPaths is not null)
            {
                foreach (var path in attachmentPaths)
                    mailItem.Attachments.Add(path);
            }

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
    /// Alle mailaccounts die in deze Outlook-installatie zijn geconfigureerd, met naam én
    /// e-mailadres, zodat de gebruiker (of de app) kan filteren op welk account verstuurd mag worden.
    /// </summary>
    public IReadOnlyList<OutlookAccount> GetAccounts()
    {
        dynamic outlookApp = GetOrCreateOutlookApplication();
        dynamic session = outlookApp.Session;
        try
        {
            var result = new List<OutlookAccount>();
            dynamic accounts = session.Accounts;
            int count = accounts.Count;
            for (var i = 1; i <= count; i++)
            {
                dynamic account = accounts[i];
                string displayName = account.DisplayName;
                result.Add(new OutlookAccount(displayName, TryGetSmtpAddress(account) ?? string.Empty));
            }

            return result;
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
                var smtp = account is null ? null : TryGetSmtpAddress(account);
                if (smtp is not null)
                    return smtp;
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

    private static string? TryGetSmtpAddress(dynamic account)
    {
        try
        {
            string smtp = account.SmtpAddress;
            return string.IsNullOrWhiteSpace(smtp) ? null : smtp;
        }
        catch (COMException)
        {
            // SmtpAddress niet beschikbaar voor dit accounttype (bv. sommige POP/IMAP-accounts).
            return null;
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

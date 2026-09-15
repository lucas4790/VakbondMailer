namespace VakbondMailer.Models;

public sealed class Recipient
{
    public required string Email { get; init; }

    public required IReadOnlyDictionary<string, string> Fields { get; init; }

    public string DisplayName
    {
        get
        {
            var nameField = Fields.FirstOrDefault(f =>
                f.Key.Contains("naam", StringComparison.OrdinalIgnoreCase) ||
                f.Key.Contains("name", StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(nameField.Value))
                return nameField.Value;

            var anyField = Fields.FirstOrDefault(f => !string.IsNullOrWhiteSpace(f.Value) && f.Value != Email);
            return !string.IsNullOrWhiteSpace(anyField.Value) ? anyField.Value : Email;
        }
    }
}

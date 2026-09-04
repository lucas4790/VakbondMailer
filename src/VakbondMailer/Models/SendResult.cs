namespace VakbondMailer.Models;

public sealed class SendResult
{
    public required string Email { get; init; }

    public required string DisplayName { get; init; }

    public required bool Success { get; init; }

    public string? Error { get; init; }
}

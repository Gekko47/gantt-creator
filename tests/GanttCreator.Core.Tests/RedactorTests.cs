using GanttCreator.Core.Logging;

namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for <see cref="Redactor"/> redaction patterns.
/// </summary>
public sealed class RedactorTests
{
    private readonly IRedactor _redactor = new Redactor();

    [Fact]
    public void Redact_email_is_masked()
    {
        var input = "Contact user@example.com for details.";
        var output = _redactor.Redact(input);
        Assert.Equal("Contact [email] for details.", output);
    }

    [Fact]
    public void Redact_multiple_emails_all_masked()
    {
        var input = "a@b.com and x@y.org";
        var output = _redactor.Redact(input);
        Assert.Equal("[email] and [email]", output);
    }

    [Fact]
    public void Redact_windows_absolute_path_is_masked()
    {
        var input = "File at C:\\Users\\name\\Documents\\file.txt";
        var output = _redactor.Redact(input);
        Assert.Equal("File at [path]", output);
    }

    [Fact]
    public void Redact_unc_path_is_masked()
    {
        var input = "Share \\\\server\\share\\folder\\file.dll";
        var output = _redactor.Redact(input);
        Assert.Equal("Share [path]", output);
    }

    [Fact]
    public void Redact_unix_path_is_masked()
    {
        var input = "Config /etc/app/config.yaml";
        var output = _redactor.Redact(input);
        Assert.Equal("Config [path]", output);
    }

    [Fact]
    public void Redact_iso_datetime_is_masked()
    {
        var input = "Started 2026-09-05T14:30:00.123Z";
        var output = _redactor.Redact(input);
        Assert.Equal("Started [datetime]", output);
    }

    [Fact]
    public void Redact_iso_datetime_with_offset_is_masked()
    {
        var input = "Event 2026-09-05 14:30:00+02:00";
        var output = _redactor.Redact(input);
        Assert.Equal("Event [datetime]", output);
    }

    [Fact]
    public void Redact_guid_is_masked()
    {
        var input = "Trace 123e4567-e89b-12d3-a456-426614174000 end";
        var output = _redactor.Redact(input);
        Assert.Equal("Trace [guid] end", output);
    }

    [Fact]
    public void Redact_long_hex_token_is_masked()
    {
        var input = "Hash abcdef1234567890abcdef1234567890";
        var output = _redactor.Redact(input);
        Assert.Equal("Hash [token]", output);
    }

    [Fact]
    public void Redact_short_hex_not_masked()
    {
        var input = "Value abc123";
        var output = _redactor.Redact(input);
        Assert.Equal("Value abc123", output);
    }

    [Fact]
    public void Redact_null_or_empty_returns_same()
    {
        Assert.Equal("", _redactor.Redact(""));
        Assert.Null(_redactor.Redact(null!));
    }

    [Fact]
    public void Redact_no_sensitive_data_unchanged()
    {
        var input = "Simple log message with numbers 123 and words.";
        var output = _redactor.Redact(input);
        Assert.Equal(input, output);
    }

    [Fact]
    public void Redact_mixed_content_all_patterns_masked()
    {
        var input = "User alice@example.com saved C:\\temp\\file.log at 2026-09-05T12:00:00Z with id 123e4567-e89b-12d3-a456-426614174000 and hash deadbeefcafebabe1234567890abcdef";
        var output = _redactor.Redact(input);
        Assert.Contains("[email]", output);
        Assert.Contains("[path]", output);
        Assert.Contains("[datetime]", output);
        Assert.Contains("[guid]", output);
        Assert.Contains("[token]", output);
        Assert.DoesNotContain("alice@example.com", output);
        Assert.DoesNotContain("C:\\temp\\file.log", output);
        Assert.DoesNotContain("2026-09-05T12:00:00Z", output);
        Assert.DoesNotContain("123e4567-e89b-12d3-a456-426614174000", output);
        Assert.DoesNotContain("deadbeefcafebabe1234567890abcdef", output);
    }
}
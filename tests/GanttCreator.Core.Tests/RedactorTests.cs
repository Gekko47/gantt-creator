using GanttCreator.Core.Logging;

namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for <see cref="Redactor"/> redaction patterns.
/// </summary>
public sealed class RedactorTests
{
    private readonly Redactor _redactor = new();

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

    // The LongHexTokenPattern requires at least one hexadecimal letter, so
    // purely decimal identifiers of 16+ digits must survive redaction.
    [Fact]
    public void Redact_purely_decimal_16_digit_identifier_not_masked()
    {
        var input = "Card 1234567890123456 processed";
        var output = _redactor.Redact(input);
        Assert.Equal(input, output);
    }

    [Fact]
    public void Redact_purely_decimal_20_digit_identifier_not_masked()
    {
        var input = "Sequence 00000000000000000000 is not a token";
        var output = _redactor.Redact(input);
        Assert.Equal(input, output);
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
        Assert.Contains("[email]", output, StringComparison.Ordinal);
        Assert.Contains("[path]", output, StringComparison.Ordinal);
        Assert.Contains("[datetime]", output, StringComparison.Ordinal);
        Assert.Contains("[guid]", output, StringComparison.Ordinal);
        Assert.Contains("[token]", output, StringComparison.Ordinal);
        Assert.DoesNotContain("alice@example.com", output, StringComparison.Ordinal);
        Assert.DoesNotContain("C:\\temp\\file.log", output, StringComparison.Ordinal);
        Assert.DoesNotContain("2026-09-05T12:00:00Z", output, StringComparison.Ordinal);
        Assert.DoesNotContain("123e4567-e89b-12d3-a456-426614174000", output, StringComparison.Ordinal);
        Assert.DoesNotContain("deadbeefcafebabe1234567890abcdef", output, StringComparison.Ordinal);
    }

    [Fact]
    public void Redact_iso_date_only_is_masked()
    {
        var input = "Event on 1990-05-10 for review.";
        var output = _redactor.Redact(input);
        Assert.Equal("Event on [datetime] for review.", output);
    }

    [Fact]
    public void Redact_prose_with_slashes_is_not_masked()
    {
        // The Unix path pattern must only fire on tokens that begin with the
        // slash. Prose slashes, ratios, and slash-separated dates are content,
        // not paths, and redacting them destroys log usefulness.
        var input = "Compare and/or combine; due 12/31/2026; ratio 3/4 and x/y.";
        Assert.Equal(input, _redactor.Redact(input));
    }

    [Fact]
    public void Redact_scheme_urls_are_not_masked()
    {
        // "p://" inside a scheme must not be masked as a Windows drive, and
        // UnixPathPattern cannot start at a '/' preceded by a word character,
        // so the whole URL survives redaction unchanged.
        var input = "See http://example.com/docs/index.html for details.";
        var output = _redactor.Redact(input);
        Assert.Equal(input, output);
    }

    [Theory]
    [InlineData("Config /etc/app/config.yaml", "Config [path]")]
    [InlineData("Read /usr/share/doc for details", "Read [path] for details")]
    [InlineData("(see /var/log/messages)", "(see [path])")]
    [InlineData("/etc/passwd was read", "[path] was read")]
    public void Redact_unix_absolute_paths_still_masked(string input, string expected)
    {
        Assert.Equal(expected, _redactor.Redact(input));
    }
}

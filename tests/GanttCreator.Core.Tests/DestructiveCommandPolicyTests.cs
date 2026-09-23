using System.Globalization;
using GanttCreator.Core;
using CommandClass = GanttCreator.Core.DestructiveCommandPolicy.CommandClass;
using ConfirmationKind = GanttCreator.Core.DestructiveCommandPolicy.ConfirmationKind;
using static GanttCreator.Core.DestructiveCommandPolicy;

namespace GanttCreator.Core.Tests;

/// <summary>
/// Contract tests for <see cref="DestructiveCommandPolicy"/>. Pure Core —
/// no Office, no interop, no UI. Every new guard and every new exact-string
/// pin has a positive test (AGENTS.md validator rule; checklist section I).
/// </summary>
public class DestructiveCommandPolicyTests
{
    // ---- Classification truth table ----

    [Theory]
    [InlineData(CommandClass.ClearTableBody)]
    [InlineData(CommandClass.RecolumniseTable)]
    [InlineData(CommandClass.ResetCatalogues)]
    public void Every_supported_destructive_command_is_destructive_but_confirmed(
        CommandClass commandClass)
    {
        ConfirmationKind kind = DestructiveCommandPolicy.Classify(commandClass);

        Assert.Equal(ConfirmationKind.DestructiveButConfirmed, kind);
    }

    [Fact]
    public void Classify_rejects_an_unrecognised_command_class()
    {
        // An unrecognised value (outside the three supported classes) is a
        // programmer error: the policy layer must throw rather than silently
        // returning no confirmation.
        CommandClass unrecognised = (CommandClass)99;

        Assert.Throws<ArgumentNullException>(
            () => DestructiveCommandPolicy.Classify(unrecognised));
    }

    [Fact]
    public void Classify_rejects_a_command_class_outside_the_supported_set()
    {
        // Force a value that is not one of the three supported classes and is
        // also not the specific value used in the "unrecognised command class"
        // test, so the two tests are independent.
        const int badValue = 255;
        CommandClass bad = (CommandClass)badValue;

        Assert.Throws<ArgumentNullException>(
            () => DestructiveCommandPolicy.Classify(bad));
    }

    // ---- Confirmation-text exact-string pins (ADR-0008 D3/D2) ----

    [Fact]
    public void BuildConfirmationText_pins_the_caption()
    {
        string text = DestructiveCommandPolicy.BuildConfirmationText(
            CommandClass.ClearTableBody);

        Assert.StartsWith("caption=Gantt Creator", text, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConfirmationText_pins_the_main_instruction()
    {
        string text = DestructiveCommandPolicy.BuildConfirmationText(
            CommandClass.ClearTableBody);

        Assert.Contains(
            "mainInstruction=Gantt Creator is about to make a destructive change",
            text,
            StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConfirmationText_pins_the_exact_no_undo_line()
    {
        // ADR-0008 D3/D2: the body must explicitly state the change cannot be
        // undone, using exactly that phrasing.
        string text = DestructiveCommandPolicy.BuildConfirmationText(
            CommandClass.ClearTableBody);

        // The body is the third field; assert the exact line is present as a
        // complete sentence.
        Assert.Contains(
            "This change cannot be undone.",
            text,
            StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConfirmationText_is_deterministic_and_count_only()
    {
        // Count-only: the builder never enumerates rows from any workbook, so
        // the same command always produces the same text regardless of any
        // workbook state. Proved by building the text twice and asserting
        // exact equality (deterministic), then asserting it contains no numeric
        // row/row-count/token that could only come from a workbook.
        string first = DestructiveCommandPolicy.BuildConfirmationText(
            CommandClass.ClearTableBody);
        string second = DestructiveCommandPolicy.BuildConfirmationText(
            CommandClass.ClearTableBody);

        Assert.Equal(first, second);

        // The body must not contain any row count, table row count, sheet
        // count, or GUID-like token that would require a live workbook. The
        // only numbers present are the enum discriminant and the literal "1"
        // in "cannot be undone" (none) — i.e. the body is purely nominal.
       //
        // Concretely: the ClearTableBody body is exactly
        // "Clear table body will delete all rows from the Gantt data table.
        // This change cannot be undone." — no count.
        string bodyField = ExtractBodyField(first);

        Assert.DoesNotContain('0', bodyField);
        Assert.DoesNotContain('1', bodyField);
        Assert.DoesNotContain('2', bodyField);
        Assert.DoesNotContain('3', bodyField);
        Assert.DoesNotContain('4', bodyField);
        Assert.DoesNotContain('5', bodyField);
        Assert.DoesNotContain('6', bodyField);
        Assert.DoesNotContain('7', bodyField);
        Assert.DoesNotContain('8', bodyField);
        Assert.DoesNotContain('9', bodyField);
    }

    [Theory]
    [InlineData(CommandClass.ClearTableBody)]
    [InlineData(CommandClass.RecolumniseTable)]
    [InlineData(CommandClass.ResetCatalogues)]
    public void BuildConfirmationText_builds_a_complete_confirmation_for_every_command(
        CommandClass commandClass)
    {
        string text = DestructiveCommandPolicy.BuildConfirmationText(commandClass);

        // A complete confirmation has three fields: caption, mainInstruction,
        // body.
        string[] fields = text.Split('\n', StringSplitOptions.None);
        Assert.Equal(3, fields.Length);

        Assert.StartsWith("caption=", fields[0], StringComparison.Ordinal);
        Assert.StartsWith("mainInstruction=", fields[1], StringComparison.Ordinal);
        Assert.StartsWith("body=", fields[2], StringComparison.Ordinal);

        // The body names the command.
        string commandName = DestructiveCommandPolicy.CommandDisplayName(commandClass);
        Assert.Contains(commandName, ExtractBodyField(text), StringComparison.Ordinal);

        // The body states the no-undo line.
        Assert.Contains(
            "This change cannot be undone.",
            ExtractBodyField(text),
            StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConfirmationText_rejects_an_unrecognised_command_class()
    {
        // An unrecognised value (outside the three supported classes) is a
        // programmer error: the policy layer must throw rather than silently
        // returning no confirmation.
        CommandClass unrecognised = (CommandClass)99;

        Assert.Throws<ArgumentNullException>(
            () => DestructiveCommandPolicy.BuildConfirmationText(unrecognised));
    }

    // ---- Display name / subject pins ----

    [Theory]
    [InlineData(CommandClass.ClearTableBody, "Clear table body")]
    [InlineData(CommandClass.RecolumniseTable, "Recolumnise table")]
    [InlineData(CommandClass.ResetCatalogues, "Reset configuration catalogues")]
    public void CommandDisplayName_returns_the_pinned_name(
        CommandClass commandClass, string expected)
    {
        Assert.Equal(expected, DestructiveCommandPolicy.CommandDisplayName(commandClass));
    }

    [Fact]
    public void CommandDisplayName_rejects_an_unrecognised_command_class()
    {
        // An unrecognised value (outside the three supported classes) is a
        // programmer error: the policy layer must throw rather than silently
        // returning no confirmation.
        CommandClass unrecognised = (CommandClass)99;

        Assert.Throws<ArgumentNullException>(
            () => DestructiveCommandPolicy.CommandDisplayName(unrecognised));
    }

    [Theory]
    [InlineData(CommandClass.ClearTableBody, "will delete all rows from the Gantt data table.")]
    [InlineData(CommandClass.RecolumniseTable, "will recreate the Gantt data table with the current columns.")]
    [InlineData(CommandClass.ResetCatalogues, "will restore the configuration catalogues to their defaults.")]
    public void CommandSubject_returns_the_pinned_subject(
        CommandClass commandClass, string expected)
    {
        Assert.Equal(expected, DestructiveCommandPolicy.CommandSubject(commandClass));
    }

    [Fact]
    public void CommandSubject_rejects_an_unrecognised_command_class()
    {
        // An unrecognised value (outside the three supported classes) is a
        // programmer error: the policy layer must throw rather than silently
        // returning no confirmation.
        CommandClass unrecognised = (CommandClass)99;

        Assert.Throws<ArgumentNullException>(
            () => DestructiveCommandPolicy.CommandSubject(unrecognised));
    }

    // ---- Culture invariance ----

    [Fact]
    public void BuildConfirmationText_is_culture_invariant_under_tr_tr()
    {
        CultureInfo original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("tr-TR");

            string text = DestructiveCommandPolicy.BuildConfirmationText(
                CommandClass.ClearTableBody);

            // The pinned English phrasing must still appear exactly; culture-
            // invariant formatting must not localise the body.
            Assert.Contains(
                "This change cannot be undone.",
                text,
                StringComparison.Ordinal);
            Assert.StartsWith("caption=Gantt Creator", text, StringComparison.Ordinal);
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    // ---- Helpers ----

    /// <summary>
    /// Extracts the body field from the three-field confirmation text.
    /// </summary>
    private static string ExtractBodyField(string text)
    {
        string[] fields = text.Split('\n', StringSplitOptions.None);
        Assert.Equal(3, fields.Length);
        // body=...  -> strip the prefix.
        return fields[2].Substring(StartingLength("body="));
    }

    private static int StartingLength(string prefix) => prefix.Length;
}

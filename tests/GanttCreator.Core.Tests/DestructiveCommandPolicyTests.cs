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
        // returning no confirmation. The default 0 value is NOT unrecognised —
        // it is the valid, named class ClearTableBody (pinned by the theory
        // above) — so only an out-of-set value may reach the throw.
        CommandClass unrecognised = (CommandClass)99;

        ArgumentOutOfRangeException error = Assert.Throws<ArgumentOutOfRangeException>(
            () => DestructiveCommandPolicy.Classify(unrecognised));

        // The parameter and the offending value must travel with the throw, so
        // a caller can tell which class was rejected.
        Assert.Equal("commandClass", error.ParamName);
        Assert.Equal(unrecognised, error.ActualValue);
    }

    [Fact]
    public void Classify_rejects_a_command_class_outside_the_supported_set()
    {
        // Force a value that is not one of the three supported classes and is
        // also not the specific value used in the "unrecognised command class"
        // test, so the two tests are independent.
        const int badValue = 255;
        CommandClass bad = (CommandClass)badValue;

        Assert.Throws<ArgumentOutOfRangeException>(
            () => DestructiveCommandPolicy.Classify(bad));
    }

    // ---- Confirmation-text exact-string pins (ADR-0008 D3/D2) ----

    [Fact]
    public void BuildConfirmationText_pins_the_caption()
    {
        ConfirmationText text = DestructiveCommandPolicy.BuildConfirmationText(
            CommandClass.ClearTableBody);

        Assert.Equal(DestructiveCommandPolicy.ConfirmationCaption, text.Caption);
    }

    [Fact]
    public void BuildConfirmationText_pins_the_main_instruction()
    {
        ConfirmationText text = DestructiveCommandPolicy.BuildConfirmationText(
            CommandClass.ClearTableBody);

        Assert.Equal(
            DestructiveCommandPolicy.ConfirmationMainInstruction,
            text.MainInstruction);
    }

    [Fact]
    public void BuildConfirmationText_pins_the_exact_no_undo_line()
    {
        // ADR-0008 D3/D2: the body must explicitly state the change cannot be
        // undone, using exactly that phrasing.
        ConfirmationText text = DestructiveCommandPolicy.BuildConfirmationText(
            CommandClass.ClearTableBody);

        Assert.Contains(
            "This change cannot be undone.",
            text.Body,
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
        ConfirmationText first = DestructiveCommandPolicy.BuildConfirmationText(
            CommandClass.ClearTableBody);
        ConfirmationText second = DestructiveCommandPolicy.BuildConfirmationText(
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
        string body = first.Body;

        Assert.DoesNotContain('0', body);
        Assert.DoesNotContain('1', body);
        Assert.DoesNotContain('2', body);
        Assert.DoesNotContain('3', body);
        Assert.DoesNotContain('4', body);
        Assert.DoesNotContain('5', body);
        Assert.DoesNotContain('6', body);
        Assert.DoesNotContain('7', body);
        Assert.DoesNotContain('8', body);
        Assert.DoesNotContain('9', body);
    }

    [Theory]
    [InlineData(CommandClass.ClearTableBody)]
    [InlineData(CommandClass.RecolumniseTable)]
    [InlineData(CommandClass.ResetCatalogues)]
    public void BuildConfirmationText_builds_a_complete_confirmation_for_every_command(
        CommandClass commandClass)
    {
        ConfirmationText text = DestructiveCommandPolicy.BuildConfirmationText(commandClass);

        // A complete confirmation carries all three dialog fields.
        Assert.Equal(DestructiveCommandPolicy.ConfirmationCaption, text.Caption);
        Assert.Equal(DestructiveCommandPolicy.ConfirmationMainInstruction, text.MainInstruction);
        Assert.NotEmpty(text.Body);

        // The body names the command.
        string commandName = DestructiveCommandPolicy.CommandDisplayName(commandClass);
        Assert.Contains(commandName, text.Body, StringComparison.Ordinal);

        // The body states the no-undo line.
        Assert.Contains(
            "This change cannot be undone.",
            text.Body,
            StringComparison.Ordinal);
    }

    [Fact]
    public void BuildConfirmationText_rejects_an_unrecognised_command_class()
    {
        // An unrecognised value (outside the three supported classes) is a
        // programmer error: the policy layer must throw rather than silently
        // returning no confirmation.
        CommandClass unrecognised = (CommandClass)99;

        Assert.Throws<ArgumentOutOfRangeException>(
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

        Assert.Throws<ArgumentOutOfRangeException>(
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

        Assert.Throws<ArgumentOutOfRangeException>(
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

            ConfirmationText text = DestructiveCommandPolicy.BuildConfirmationText(
                CommandClass.ClearTableBody);

            // The pinned English phrasing must still appear exactly; culture-
            // invariant formatting must not localise the body.
            Assert.Contains(
                "This change cannot be undone.",
                text.Body,
                StringComparison.Ordinal);
            Assert.Equal(DestructiveCommandPolicy.ConfirmationCaption, text.Caption);
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }
}

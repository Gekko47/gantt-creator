using GanttCreator.Office;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Copy Editable and Send to PowerPoint present the same pre-flight refusals to the user, so they
/// must say the same thing about the same reason, differing only in the command's own voice
/// ("copy again" versus "send again", "Copy Editable could not find" versus "Send to PowerPoint
/// could not find").
/// </summary>
/// <remarks>
/// This is the guard for two failures the compiler cannot see. A refusal reason added to one of
/// the shared enums lands on the trailing arm in both commands and would reach the user as the
/// undefined-value sentence; and the two commands can drift apart in wording, which is what
/// happened to the four catalogue reasons before this test existed.
/// </remarks>
public class ExportRefusalWordingTests
{
    [Fact]
    public void Read_refusals_are_described_per_reason_and_both_commands_agree() =>
        AssertAllDescribed<GanttTableReadRefusalReason>(
            CopyEditableCommand.TranslateRead,
            SendToPowerPointCommand.TranslateRead);

    [Fact]
    public void Config_refusals_are_described_per_reason_and_both_commands_agree() =>
        AssertAllDescribed<ConfigReadRefusalReason>(
            CopyEditableCommand.TranslateConfig,
            SendToPowerPointCommand.TranslateConfig);

    [Fact]
    public void Measurement_refusals_are_described_per_reason_and_both_commands_agree() =>
        AssertAllDescribed<PanelGridRefusalReason>(
            CopyEditableCommand.TranslateMeasurement,
            SendToPowerPointCommand.TranslateMeasurement);

    /// <summary>Checks every defined reason against both invariants.</summary>
    /// <typeparam name="TReason">The refusal enum the commands translate.</typeparam>
    /// <param name="copy">Copy Editable's translator.</param>
    /// <param name="send">Send to PowerPoint's translator.</param>
    private static void AssertAllDescribed<TReason>(
        Func<TReason, string> copy,
        Func<TReason, string> send)
        where TReason : struct, Enum
    {
        // A cast can produce any integer, so the trailing arm is reachable and its sentence is
        // what a newly added member would inherit silently.
        TReason undefined = (TReason)Enum.ToObject(typeof(TReason), 9999);
        string copyFallback = copy(undefined);
        string sendFallback = send(undefined);
        List<string> problems = [];

        foreach (TReason reason in Enum.GetValues<TReason>())
        {
            string copySentence = copy(reason);
            string sendSentence = send(reason);

            if (string.Equals(copySentence, copyFallback, StringComparison.Ordinal)
                || string.Equals(sendSentence, sendFallback, StringComparison.Ordinal))
            {
                problems.Add($"{reason}: falls through to the undefined-value sentence");
            }

            if (!string.Equals(InSendVoice(copySentence), sendSentence, StringComparison.Ordinal))
            {
                problems.Add($"{reason}: copy says \"{copySentence}\" but send says \"{sendSentence}\"");
            }
        }

        Assert.Empty(problems);
    }

    /// <summary>Rewrites a Copy Editable sentence in Send to PowerPoint's voice.</summary>
    /// <param name="copySentence">The Copy Editable sentence.</param>
    /// <returns>The same sentence as Send to PowerPoint would word it.</returns>
    private static string InSendVoice(string copySentence) =>
        copySentence
            .Replace("Copy Editable", "Send to PowerPoint", StringComparison.Ordinal)
            .Replace("copy", "send", StringComparison.Ordinal);
}

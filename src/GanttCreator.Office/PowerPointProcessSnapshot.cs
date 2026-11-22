using System.ComponentModel;
using System.Diagnostics;

namespace GanttCreator.Office;

/// <summary>
/// Reads the <c>POWERPNT</c> process facts that decide whether an attach-or-start call
/// created the host (R7.1 D2, work item Step 0).
/// </summary>
/// <remarks>
/// The snapshot is the measured basis for <c>StartedByUs</c>: a typed
/// <c>new Application()</c> both attaches to a running instance and starts a stopped one, so
/// the only evidence of ownership is a process that did not exist before the call. These
/// helpers take and return primitives so the derivation and host resolution are testable
/// without a host.
/// </remarks>
internal static class PowerPointProcessSnapshot
{
    /// <summary>The PowerPoint executable name, without the extension.</summary>
    internal const string ProcessName = "POWERPNT";

    /// <summary>Captures the current PowerPoint process identifiers, sorted ascending.</summary>
    /// <returns>The process identifiers; empty when no PowerPoint is running.</returns>
    internal static int[] Capture()
    {
        Process[] processes = Process.GetProcessesByName(ProcessName);
        try
        {
            int[] processIds = new int[processes.Length];
            for (int index = 0; index < processes.Length; index++)
            {
                processIds[index] = processes[index].Id;
            }

            Array.Sort(processIds);
            return processIds;
        }
        finally
        {
            foreach (Process process in processes)
            {
                process.Dispose();
            }
        }
    }

    /// <summary>
    /// Derives whether this call created the host: no process before and at least one after.
    /// </summary>
    /// <param name="before">The process identifiers captured before the typed call.</param>
    /// <param name="after">The process identifiers captured after the typed call.</param>
    /// <returns>
    /// <see langword="true"/> only when PowerPoint was not running before the call and was
    /// running after it; a pre-existing host is never owned.
    /// </returns>
    internal static bool DeriveStartedByUs(int[] before, int[] after) =>
        before.Length == 0 && after.Length > 0;

    /// <summary>Resolves the host process the adapter is attached to.</summary>
    /// <param name="before">The process identifiers captured before the typed call.</param>
    /// <param name="after">The process identifiers captured after the typed call.</param>
    /// <returns>
    /// The first process created by the call when there is one, otherwise the first running
    /// PowerPoint process, otherwise zero.
    /// </returns>
    internal static int ResolveHostProcessId(int[] before, int[] after)
    {
        foreach (int processId in after)
        {
            if (Array.IndexOf(before, processId) < 0)
            {
                return processId;
            }
        }

        return after.Length > 0 ? after[0] : 0;
    }

    /// <summary>Reads the host's file version for the technical record, or <see langword="null"/>.</summary>
    /// <param name="processId">The host process identifier, or zero when it is unknown.</param>
    /// <returns>The file version string, or <see langword="null"/> when it cannot be read.</returns>
    internal static string? DescribeBuild(int processId)
    {
        if (processId <= 0)
        {
            return null;
        }

        try
        {
            using Process process = Process.GetProcessById(processId);
            return process.MainModule?.FileVersionInfo.FileVersion;
        }
        catch (Exception exception) when (
            exception is ArgumentException
                or InvalidOperationException
                or Win32Exception
                or NotSupportedException)
        {
            // The build string is diagnostics only; losing it must not lose the operation.
            return null;
        }
    }
}

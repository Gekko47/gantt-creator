using System.IO;
using Xunit;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// Positive tests for safety guards in <c>scripts/check-status.ps1</c>.
///
/// The script is a PowerShell file; we cannot import and call it from a
/// C# xUnit harness, so we assert on the source text. These are
/// tripwires: if someone deletes or weakens a guard, the next commit
/// fails here and the developer reads the diff, not the runtime error.
/// </summary>
public sealed class CheckStatusScriptTests
{
    private static string ReadCheckStatus()
    {
        var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(CheckStatusScriptTests).Assembly.Location)!);
        while (dir is not null)
        {
            var candidate = Path.Combine(dir.FullName, "scripts", "check-status.ps1");
            if (File.Exists(candidate)) return File.ReadAllText(candidate);
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException("Could not locate scripts/check-status.ps1.");
    }

    [Fact]
    public void CheckStatus_scopes_git_to_repo_root()
    {
        // The script computes $repoRoot for a reason: git invocations must
        // run against that root, not the caller's CWD. A regression that
        // drops the -C flag would break when run from any directory other
        // than the repo root.
        var text = ReadCheckStatus();
        Assert.Contains("git -C $repoRoot rev-parse", text, System.StringComparison.Ordinal);
    }

    [Fact]
    public void CheckStatus_normalises_paths_and_rejects_traversal()
    {
        // The path-existence check must resolve each candidate and reject
        // any path that escapes the repo root via '..' or absolute paths.
        // A backslash-only resolution would have allowed a STATUS entry
        // like '`../../etc/passwd.md`' to pass; this assertion would
        // catch that regression on commit.
        var text = ReadCheckStatus();
        Assert.Contains("GetFullPath", text, System.StringComparison.Ordinal);
        Assert.Contains("resolves outside the repository", text, System.StringComparison.Ordinal);
    }
}

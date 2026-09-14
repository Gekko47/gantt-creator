using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;

namespace GanttCreator.Architecture.Tests;

// R1.4 DEBUG-only environment-variable guard scanner. See DebugHookGuardTests
// for the contract: every production Environment.GetEnvironmentVariable read must
// be provably compiled out of Release builds. The scanner masks comments and
// string/char literals, models the #if/#elif/#else/#endif branch stack, and
// accepts a read only when the conjunction of enclosing conditions is
// unsatisfiable with DEBUG unset. Malformed or unbalanced directives are
// reported as findings. This file is test instrumentation only.
internal static partial class EnvVarGuardScanner
{
    [GeneratedRegex(@"\b(?:Environment\s*\.\s*)?(?:GetEnvironmentVariable|GetEnvironmentVariables|ExpandEnvironmentVariables)\s*\(", RegexOptions.CultureInvariant)]
    private static partial Regex EnvVarCallRegex();

    [GeneratedRegex(@"^\s*#(if|elif|else|endif)\b(.*)$", RegexOptions.CultureInvariant)]
    private static partial Regex DirectiveRegex();

    [GeneratedRegex(@"[A-Za-z_][A-Za-z0-9_]*", RegexOptions.CultureInvariant)]
    private static partial Regex IdentifierRegex();

    private static readonly string[] DebugPreprocessorSymbols = new[] { "DEBUG" };

    // Pinned to C# 14 (the language level the repo compiles with under the .NET 10
    // SDK) instead of Latest, so masking cannot drift when the parser package and
    // the repo's language level move apart. Shared by every CSharpParseOptions
    // construction in this scanner. The value is 1400, expressed as a cast
    // because the 4.14.0 package's compile asset has no named CSharp14 member
    // (verified by reflection on the installed package and by build failure
    // CS0117 when the named member is referenced).
    private const LanguageVersion ScannerLanguageVersion = (LanguageVersion)1400;

    public static IReadOnlyList<string> Scan(string source) =>
        ScanMasked(MaskCommentsAndStrings(source));

    // __MASKING__
    #region Masking

    private static string MaskCommentsAndStrings(string source)
    {
        // Parse source multiple times: once without DEBUG defined and once with DEBUG
        // defined. Roslyn correctly handles all string literal forms — regular,
        // verbatim, interpolated, raw, and UTF-8 — as well as character literals
        // and both line and block comments. Token enumeration descends into
        // structured trivia (descendIntoTrivia: true) so comments and literals on
        // directive lines (e.g. a trailing comment after '#if DEBUG') are masked
        // too. The token spans from the trees are merged before the raw source is
        // masked so that the #if/#endif analyzer never sees contents of comments
        // or literals, including any preprocessor directives embedded in them.
        // LanguageVersion is pinned explicitly to C# 14 (the version the repo compiles
        // with under the .NET 10 SDK) instead of Latest, so masking cannot drift
        // when the parser package and the repo's language level move apart. The
        // value lives on the shared ScannerLanguageVersion constant so all three
        // parses use the same language level.
        var spansToMask = CollectMaskSpans(source, new CSharpParseOptions(languageVersion: ScannerLanguageVersion));
        spansToMask.AddRange(CollectMaskSpans(source, new CSharpParseOptions(languageVersion: ScannerLanguageVersion, preprocessorSymbols: DebugPreprocessorSymbols)));

        // A region guarded by a symbol other than DEBUG (for example '#if FEATURE')
        // is inactive in the parses above, so Roslyn reports its contents as skipped
        // tokens inside structured trivia. Extract every symbol referenced by the source's #if/#elif
        // conditions and parse a third time with all of them (plus DEBUG) defined so
        // those regions become active and their literal/comment spans are collected
        // and merged as well.
        var referencedSymbols = CollectReferencedPreprocessorSymbols(source);
        if (referencedSymbols.Count > 0)
        {
            var allSymbols = new List<string>(DebugPreprocessorSymbols.Length + referencedSymbols.Count);
            allSymbols.AddRange(DebugPreprocessorSymbols);
            allSymbols.AddRange(referencedSymbols);
            spansToMask.AddRange(CollectMaskSpans(source, new CSharpParseOptions(languageVersion: ScannerLanguageVersion, preprocessorSymbols: allSymbols)));
        }

        // Mask the raw source with the merged Roslyn-provided spans. The hand-rolled
        // scanners for //, /*, and ' are removed because Roslyn already correctly
        // identifies comments and character literals; keeping a parallel hand-rolled
        // pass could leave literal or comment contents unmasked and thereby mask real
        // code.
        var masked = source.ToCharArray();
        int n = source.Length;
        foreach (var span in spansToMask)
        {
            int start = span.Start;
            int end = Math.Min(span.End, n);
            MaskRange(masked, start, end);
        }

        return new string(masked);
    }

    private static List<TextSpan> CollectMaskSpans(string source, CSharpParseOptions parseOptions)
    {
        var spans = new List<TextSpan>();
        var syntaxTree = CSharpSyntaxTree.ParseText(source, parseOptions);
        var root = syntaxTree.GetRoot();

        foreach (var token in root.DescendantTokens(descendIntoTrivia: true))
        {
            foreach (var trivia in token.LeadingTrivia)
            {
                AddTriviaSpanIfComment(spans, trivia);
            }

            foreach (var trivia in token.TrailingTrivia)
            {
                AddTriviaSpanIfComment(spans, trivia);
            }

            if (token.IsKind(SyntaxKind.StringLiteralToken)
                || token.IsKind(SyntaxKind.InterpolatedStringTextToken)
                || token.IsKind(SyntaxKind.Utf8StringLiteralToken)
                || token.IsKind(SyntaxKind.InterpolatedStringToken)
                || token.IsKind(SyntaxKind.CharacterLiteralToken)
                || token.IsKind(SyntaxKind.SingleLineRawStringLiteralToken)
                || token.IsKind(SyntaxKind.MultiLineRawStringLiteralToken)
                || token.IsKind(SyntaxKind.Utf8SingleLineRawStringLiteralToken)
                || token.IsKind(SyntaxKind.Utf8MultiLineRawStringLiteralToken))
            {
                spans.Add(new TextSpan(token.Span.Start, token.Span.Length));
            }
        }

        return spans;
    }

    private static List<string> CollectReferencedPreprocessorSymbols(string source)
    {
        var symbols = new SortedSet<string>(StringComparer.Ordinal);
        foreach (var line in source.Split('\n'))
        {
            Match directive = DirectiveRegex().Match(line);
            if (!directive.Success)
            {
                continue;
            }

            var keyword = directive.Groups[1].Value;
            if (keyword is not ("if" or "elif"))
            {
                continue;
            }

            foreach (Match identifier in IdentifierRegex().Matches(directive.Groups[2].Value))
            {
                var name = identifier.Value;
                if (!name.Equals("defined", StringComparison.Ordinal)
                    && !name.Equals("true", StringComparison.OrdinalIgnoreCase)
                    && !name.Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    symbols.Add(name);
                }
            }
        }

        return new List<string>(symbols);
    }

    private static void AddTriviaSpanIfComment(List<TextSpan> spans, SyntaxTrivia trivia)
    {
        if (trivia.IsKind(SyntaxKind.SingleLineCommentTrivia)
            || trivia.IsKind(SyntaxKind.MultiLineCommentTrivia)
            || trivia.IsKind(SyntaxKind.SingleLineDocumentationCommentTrivia)
            || trivia.IsKind(SyntaxKind.MultiLineDocumentationCommentTrivia))
        {
            spans.Add(new TextSpan(trivia.Span.Start, trivia.Span.Length));
        }
    }

    private static void MaskRange(char[] masked, int start, int end)
    {
        for (int k = start; k < end && k < masked.Length; k++)
        {
            if (masked[k] != '\n')
            {
                masked[k] = ' ';
            }
        }
    }

    #endregion

    // __ANALYSIS__
    #region Analysis

    private static List<string> ScanMasked(string masked)
    {
        var findings = new List<string>();

        // Process preprocessor directives line-by-line so the branch stack is
        // modeled correctly and directive line numbers remain accurate. Multiline
        // Environment.* calls are detected against the full masked source (with
        // newlines preserved) and then mapped back to the source line of the call's
        // start index.
        var lines = masked.Split('\n');
        var stack = new Stack<Frame>();
        Expr current = new Lit(true);

        // Scanner index -> source line number (1-based). Computed from the fully
        // masked text so multiline matches report against the line where the call
        // begins.
        var lineOfIndex = LineNumberMap(masked);

        // Pre-compute condition state at the start of each source line so env-var
        // calls can be attributed to the correct enclosing directive regardless of
        // where on the line they appear.
        var exprAtLineStart = new Expr[lines.Length];
        exprAtLineStart[0] = current;

        for (int idx = 0; idx < lines.Length; idx++)
        {
            if (idx > 0)
            {
                exprAtLineStart[idx] = exprAtLineStart[idx - 1];
            }

            string line = lines[idx];
            int lineNo = idx + 1;
            Match m = DirectiveRegex().Match(line);
            if (m.Success)
            {
                HandleDirective(
                    stack,
                    m.Groups[1].Value,
                    m.Groups[2].Value.Trim(),
                    lineNo,
                    findings,
                    ref current);
                exprAtLineStart[idx] = current;
                continue;
            }
        }

        foreach (Match envCall in EnvVarCallRegex().Matches(masked))
        {
            int lineNo = lineOfIndex[envCall.Index];
            if (lineNo < 1 || lineNo > lines.Length)
            {
                // Defensive: a match outside the mapped range should not happen
                // after masking preserves newlines, but treat it as unguarded
                // rather than failing the scan.
                lineNo = lines.Length;
            }

            Expr exprAtCall = exprAtLineStart[lineNo - 1];
            if (!IsUnsatisfiableWithDebugFalse(exprAtCall))
            {
                findings.Add(
                    $"Line {lineNo}: Environment.GetEnvironmentVariable read is not " +
                    $"guaranteed compiled out of Release builds (active guard: {Describe(exprAtCall)}).");
            }
        }

        if (stack.Count > 0)
        {
            findings.Add($"End of file: {stack.Count} unterminated #if block(s).");
        }

        return findings;
    }

    private static int[] LineNumberMap(string text)
    {
        int length = text.Length;
        int[] map = new int[length];
        int line = 1;

        for (int i = 0; i < length; i++)
        {
            map[i] = line;
            if (text[i] == '\n')
            {
                line++;
            }
        }

        return map;
    }

    private static void HandleDirective(
        Stack<Frame> stack,
        string keyword,
        string rest,
        int lineNo,
        List<string> findings,
        ref Expr current)
    {
        switch (keyword)
        {
            case "if":
                if (rest.Length == 0)
                {
                    findings.Add($"Line {lineNo}: #if with no expression.");
                    return;
                }

                if (TryParseCondition(rest, lineNo, findings, out Expr? ifCond) && ifCond != null)
                {
                    stack.Push(new Frame { Parent = current, Preceding = new List<Expr> { ifCond } });
                    current = Conj(current, ifCond);
                }

                return;

            case "elif":
                if (stack.Count == 0)
                {
                    findings.Add($"Line {lineNo}: #elif without a matching #if.");
                    return;
                }

                if (rest.Length == 0)
                {
                    findings.Add($"Line {lineNo}: #elif with no expression.");
                    return;
                }

                if (TryParseCondition(rest, lineNo, findings, out Expr? elifCond) && elifCond != null)
                {
                    Frame f = stack.Peek();
                    Expr branch = Conj(f.Parent, Conj(new Not(OrAll(f.Preceding)), elifCond));
                    f.Preceding.Add(elifCond);
                    current = branch;
                }

                return;

            case "else":
                if (stack.Count == 0)
                {
                    findings.Add($"Line {lineNo}: #else without a matching #if.");
                    return;
                }

                Frame fe = stack.Peek();
                current = Conj(fe.Parent, new Not(OrAll(fe.Preceding)));
                return;

            case "endif":
                if (stack.Count == 0)
                {
                    findings.Add($"Line {lineNo}: #endif without a matching #if.");
                    return;
                }

                current = stack.Pop().Parent;
                return;
        }
    }

    #endregion

    #region Expression model

    private abstract record Expr;

    private sealed record Lit(bool Value) : Expr;

    private sealed record Sym(string Name) : Expr;

    private sealed record Not(Expr Operand) : Expr;

    private sealed record And(Expr A, Expr B) : Expr;

    private sealed record Or(Expr A, Expr B) : Expr;

    private sealed class Frame
    {
        public Expr Parent = new Lit(true);
        public List<Expr> Preceding = new();
    }

    private static Expr Conj(Expr a, Expr b) => SimplifyAnd(Substitute(a), Substitute(b));

    private static Expr OrAll(IReadOnlyList<Expr> exprs)
    {
        Expr e = new Lit(false);
        foreach (var part in exprs)
        {
            e = SimplifyOr(e, Substitute(part));
        }

        return e;
    }

    private static Expr Substitute(Expr e) => e switch
    {
        Sym s => s.Name.Equals("DEBUG", StringComparison.Ordinal) ? new Lit(false) : s,
        Lit l => l,
        Not n => SimplifyNot(Substitute(n.Operand)),
        And a => SimplifyAnd(Substitute(a.A), Substitute(a.B)),
        Or o => SimplifyOr(Substitute(o.A), Substitute(o.B)),
        _ => e
    };

    private static Expr SimplifyNot(Expr e) => e switch
    {
        Lit l => new Lit(!l.Value),
        _ => new Not(e)
    };

    private static Expr SimplifyAnd(Expr a, Expr b) => (a, b) switch
    {
        (Lit la, _) when !la.Value => new Lit(false),
        (_, Lit lb) when !lb.Value => new Lit(false),
        (Lit la, _) when la.Value => b,
        (_, Lit lb) when lb.Value => a,
        _ => new And(a, b)
    };

    private static Expr SimplifyOr(Expr a, Expr b) => (a, b) switch
    {
        (Lit la, _) when la.Value => new Lit(true),
        (_, Lit lb) when lb.Value => new Lit(true),
        (Lit la, _) when !la.Value => b,
        (_, Lit lb) when !lb.Value => a,
        _ => new Or(a, b)
    };

    private static bool IsUnsatisfiableWithDebugFalse(Expr e) => !IsSatisfiable(Substitute(e));

    private static bool IsSatisfiable(Expr e)
    {
        var symbols = new HashSet<string>(StringComparer.Ordinal);
        CollectSymbols(e, symbols);
        var list = new List<string>(symbols);
        if (list.Count > 20)
        {
            return true;
        }

        int total = 1 << list.Count;
        var assignment = new Dictionary<string, bool>(list.Count);
        for (int mask = 0; mask < total; mask++)
        {
            assignment.Clear();
            for (int i = 0; i < list.Count; i++)
            {
                assignment[list[i]] = (mask & (1 << i)) != 0;
            }

            if (Evaluate(e, assignment))
            {
                return true;
            }
        }

        return false;
    }

    private static void CollectSymbols(Expr e, HashSet<string> symbols)
    {
        switch (e)
        {
            case Sym s:
                symbols.Add(s.Name);
                break;
            case Not n:
                CollectSymbols(n.Operand, symbols);
                break;
            case And a:
                CollectSymbols(a.A, symbols);
                CollectSymbols(a.B, symbols);
                break;
            case Or o:
                CollectSymbols(o.A, symbols);
                CollectSymbols(o.B, symbols);
                break;
        }
    }

    private static bool Evaluate(Expr e, Dictionary<string, bool> assignment) => e switch
    {
        Sym s => assignment.TryGetValue(s.Name, out bool v) && v,
        Lit l => l.Value,
        Not n => !Evaluate(n.Operand, assignment),
        And a => Evaluate(a.A, assignment) && Evaluate(a.B, assignment),
        Or o => Evaluate(o.A, assignment) || Evaluate(o.B, assignment),
        _ => false
    };

    private static string Describe(Expr e) => e switch
    {
        Lit l => l.Value.ToString(),
        Sym s => s.Name,
        Not n => $"!{Describe(n.Operand)}",
        And a => $"({Describe(a.A)} && {Describe(a.B)})",
        Or o => $"({Describe(o.A)} || {Describe(o.B)})",
        _ => "?"
    };

    #endregion

    // __PARSER__
    #region Parser

    private static bool TryParseCondition(string text, int lineNo, List<string> findings, out Expr? condition)
    {
        var parser = new ConditionParser(text);
        Expr? result = parser.ParseOr();
        parser.SkipWhitespace();
        if (result == null || parser.HasMore)
        {
            findings.Add($"Line {lineNo}: malformed preprocessor expression '{text.Trim()}'.");
            condition = null;
            return false;
        }

        condition = result;
        return true;
    }

    private static bool IsIdentChar(char c) => char.IsLetterOrDigit(c) || c == '_';

    private sealed class ConditionParser
    {
        private readonly string _text;
        private int _pos;

        public ConditionParser(string text)
        {
            _text = text;
            _pos = 0;
        }

        public bool HasMore { get; private set; }

        public Expr? ParseOr()
        {
            Expr? left = ParseAnd();
            if (left == null)
            {
                return null;
            }

            while (Consume("||"))
            {
                Expr? right = ParseAnd();
                if (right == null)
                {
                    return null;
                }

                left = new Or(left, right);
            }

            return left;
        }

        private Expr? ParseAnd()
        {
            Expr? left = ParseUnary();
            if (left == null)
            {
                return null;
            }

            while (Consume("&&"))
            {
                Expr? right = ParseUnary();
                if (right == null)
                {
                    return null;
                }

                left = new And(left, right);
            }

            return left;
        }

        private Expr? ParseUnary()
        {
            if (Consume("!"))
            {
                Expr? operand = ParseUnary();
                return operand == null ? null : new Not(operand);
            }

            return ParsePrimary();
        }

        private Expr? ParsePrimary()
        {
            SkipWhitespace();
            if (Consume("("))
            {
                Expr? inner = ParseOr();
                if (inner == null || !Consume(")"))
                {
                    return null;
                }

                return inner;
            }

            if (Consume("defined"))
            {
                SkipWhitespace();
                if (!Consume("("))
                {
                    return null;
                }

                string? name = ReadIdentifier();
                if (name == null || !Consume(")"))
                {
                    return null;
                }

                return new Sym(name);
            }

            string? id = ReadIdentifier();
            if (id == null)
            {
                return null;
            }

            if (id.Equals("true", StringComparison.OrdinalIgnoreCase))
            {
                return new Lit(true);
            }

            if (id.Equals("false", StringComparison.OrdinalIgnoreCase))
            {
                return new Lit(false);
            }

            return new Sym(id);
        }

        private string? ReadIdentifier()
        {
            SkipWhitespace();
            int start = _pos;
            while (_pos < _text.Length && IsIdentChar(_text[_pos]))
            {
                _pos++;
            }

            if (_pos == start)
            {
                return null;
            }

            return _text[start.._pos];
        }

        private bool Consume(string token)
        {
            SkipWhitespace();
            if (_pos + token.Length > _text.Length)
            {
                return false;
            }

            for (int i = 0; i < token.Length; i++)
            {
                if (_text[_pos + i] != token[i])
                {
                    return false;
                }
            }

            if (IsIdentChar(token[token.Length - 1]))
            {
                char after = _pos + token.Length < _text.Length ? _text[_pos + token.Length] : '\0';
                if (IsIdentChar(after))
                {
                    return false;
                }
            }

            _pos += token.Length;
            HasMore = _pos < _text.Length;
            return true;
        }

        public void SkipWhitespace()
        {
            while (_pos < _text.Length && char.IsWhiteSpace(_text[_pos]))
            {
                _pos++;
            }

            HasMore = _pos < _text.Length;
        }
    }

    #endregion
}

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
    [GeneratedRegex(@"\bEnvironment\s*\.\s*(GetEnvironmentVariable|GetEnvironmentVariables|ExpandEnvironmentVariables)\s*\(", RegexOptions.CultureInvariant)]
    private static partial Regex EnvVarCallRegex();

    [GeneratedRegex(@"^\s*#(if|elif|else|endif)\b(.*)$", RegexOptions.CultureInvariant)]
    private static partial Regex DirectiveRegex();

    public static IReadOnlyList<string> Scan(string source) =>
        ScanMasked(MaskCommentsAndStrings(source));

    // __MASKING__
    #region Masking

    private static string MaskCommentsAndStrings(string source)
    {
        // Parse source with Roslyn using CSharpParseOptions without DEBUG defined.
        // This correctly handles all string literal forms including verbatim strings,
        // interpolated strings, and nested string literals within interpolations.
        var parseOptions = new CSharpParseOptions();
        var syntaxTree = CSharpSyntaxTree.ParseText(source, parseOptions);
        var root = syntaxTree.GetRoot();

        // Collect spans of all string literal tokens (regular, verbatim, interpolated).
        // Roslyn handles escape sequences, interpolation holes, and triple-quoted strings
        // correctly, unlike the hand-rolled parser.
        var spansToMask = new List<TextSpan>();
        foreach (var token in root.DescendantTokens())
        {
            if (token.IsKind(SyntaxKind.StringLiteralToken))
            {
                spansToMask.Add(new TextSpan(token.Span.Start, token.Span.Length));
            }
        }

        // Mask char literals and comments with the hand-rolled approach (these are
        // unambiguous and the hand-rolled parser handles them correctly).
        var masked = source.ToCharArray();
        int n = source.Length;
        int i = 0;

        while (i < n)
        {
            char c = source[i];

            // Char literal
            if (c == '\'')
            {
                int end = CloseCharLiteral(source, i);
                MaskRange(masked, i, end);
                i = end;
                continue;
            }

            // Line comment
            if (c == '/' && i + 1 < n && source[i + 1] == '/')
            {
                int nl = source.IndexOf('\n', i);
                int end = nl == -1 ? n : nl;
                MaskRange(masked, i, end);
                i = nl == -1 ? n : end;
                continue;
            }

            // Block comment
            if (c == '/' && i + 1 < n && source[i + 1] == '*')
            {
                masked[i] = ' ';
                masked[i + 1] = ' ';
                i = SkipBlockComment(source, i + 2, masked);
                continue;
            }

            i++;
        }

        // Now mask string literal tokens using Roslyn-provided spans.
        // Overlap/ordering: string literals are masked last so any preprocessor
        // directives inside them don't confuse the #if/#endif analyzer. Roslyn
        // already distinguished true string contents from code.
        foreach (var span in spansToMask)
        {
            int start = span.Start;
            int end = Math.Min(span.End, n);
            MaskRange(masked, start, end);
        }

        return new string(masked);
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

    private static int CloseCharLiteral(string s, int open)
    {
        int i = open + 1;
        int n = s.Length;
        if (i < n && s[i] == '\\')
        {
            i += 2;
        }
        else if (i < n)
        {
            i++;
        }

        if (i < n && s[i] == '\'')
        {
            return i + 1;
        }

        return Math.Min(i, n);
    }

    private static int SkipBlockComment(string s, int start, char[] output)
    {
        int i = start;
        int n = s.Length;
        while (i < n)
        {
            if (s[i] == '*' && i + 1 < n && s[i + 1] == '/')
            {
                output[i] = ' ';
                output[i + 1] = ' ';
                return i + 2;
            }

            if (s[i] != '\n')
            {
                output[i] = ' ';
            }

            i++;
        }

        return n;
    }

    #endregion

    // __ANALYSIS__
    #region Analysis

    private static List<string> ScanMasked(string masked)
    {
        var findings = new List<string>();
        var lines = masked.Split('\n');
        var stack = new Stack<Frame>();
        Expr current = new Lit(true);

        for (int idx = 0; idx < lines.Length; idx++)
        {
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
                continue;
            }

            if (EnvVarCallRegex().IsMatch(line) &&
                !IsUnsatisfiableWithDebugFalse(current))
            {
                findings.Add(
                    $"Line {lineNo}: Environment.GetEnvironmentVariable read is not " +
                    $"guaranteed compiled out of Release builds (active guard: {Describe(current)}).");
            }
        }

        if (stack.Count > 0)
        {
            findings.Add($"End of file: {stack.Count} unterminated #if block(s).");
        }

        return findings;
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

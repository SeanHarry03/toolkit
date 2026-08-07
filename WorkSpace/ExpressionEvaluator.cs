using System.Globalization;

namespace ExcelToJson;

/// <summary>
/// 简单算术表达式：字段名、数字、+ - * /、括号。
/// 字段按 double 取值（缺失为 0）。
/// </summary>
public static class ExpressionEvaluator
{
    public static Func<IReadOnlyDictionary<string, object?>, object?> Compile(string expression)
    {
        if (string.IsNullOrWhiteSpace(expression))
            throw new ArgumentException("表达式为空");

        var tokens = Tokenize(expression);
        var pos = 0;
        var root = ParseExpression(tokens, ref pos);
        if (pos != tokens.Count)
            throw new FormatException($"表达式未解析完: {expression}");

        return row => Evaluate(root, row);
    }

    public static bool TryCompile(string expression, out Func<IReadOnlyDictionary<string, object?>, object?>? func, out string? error)
    {
        try
        {
            func = Compile(expression);
            error = null;
            return true;
        }
        catch (Exception ex)
        {
            func = null;
            error = ex.Message;
            return false;
        }
    }

    private abstract class Node;

    private sealed class NumberNode(double value) : Node
    {
        public double Value { get; } = value;
    }

    private sealed class IdentNode(string name) : Node
    {
        public string Name { get; } = name;
    }

    private sealed class BinaryNode(char op, Node left, Node right) : Node
    {
        public char Op { get; } = op;
        public Node Left { get; } = left;
        public Node Right { get; } = right;
    }

    private sealed class UnaryNode(char op, Node inner) : Node
    {
        public char Op { get; } = op;
        public Node Inner { get; } = inner;
    }

    private static double Evaluate(Node node, IReadOnlyDictionary<string, object?> row) => node switch
    {
        NumberNode n => n.Value,
        IdentNode i => GetDouble(row, i.Name),
        UnaryNode u => u.Op == '-' ? -Evaluate(u.Inner, row) : Evaluate(u.Inner, row),
        BinaryNode b => b.Op switch
        {
            '+' => Evaluate(b.Left, row) + Evaluate(b.Right, row),
            '-' => Evaluate(b.Left, row) - Evaluate(b.Right, row),
            '*' => Evaluate(b.Left, row) * Evaluate(b.Right, row),
            '/' => Evaluate(b.Right, row) == 0 ? 0 : Evaluate(b.Left, row) / Evaluate(b.Right, row),
            _ => throw new InvalidOperationException($"未知运算符: {b.Op}"),
        },
        _ => 0,
    };

    public static double GetDouble(IReadOnlyDictionary<string, object?> row, string key)
    {
        if (!row.TryGetValue(key, out var val) || val == null)
            return 0;
        if (val is double d)
            return d;
        if (val is long l)
            return l;
        if (val is int i)
            return i;
        if (val is float f)
            return f;
        if (double.TryParse(val.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
            return parsed;
        return 0;
    }

    private enum TokKind { Number, Ident, Op, LParen, RParen }

    private readonly record struct Tok(TokKind Kind, string Text);

    private static List<Tok> Tokenize(string expr)
    {
        var list = new List<Tok>();
        var i = 0;
        while (i < expr.Length)
        {
            var c = expr[i];
            if (char.IsWhiteSpace(c))
            {
                i++;
                continue;
            }

            if (char.IsDigit(c) || (c == '.' && i + 1 < expr.Length && char.IsDigit(expr[i + 1])))
            {
                var start = i;
                i++;
                while (i < expr.Length && (char.IsDigit(expr[i]) || expr[i] == '.'))
                    i++;
                list.Add(new Tok(TokKind.Number, expr[start..i]));
                continue;
            }

            if (char.IsLetter(c) || c == '_' )
            {
                var start = i;
                i++;
                while (i < expr.Length && (char.IsLetterOrDigit(expr[i]) || expr[i] == '_'))
                    i++;
                list.Add(new Tok(TokKind.Ident, expr[start..i]));
                continue;
            }

            if (c is '+' or '-' or '*' or '/')
            {
                list.Add(new Tok(TokKind.Op, c.ToString()));
                i++;
                continue;
            }

            if (c == '(')
            {
                list.Add(new Tok(TokKind.LParen, "("));
                i++;
                continue;
            }

            if (c == ')')
            {
                list.Add(new Tok(TokKind.RParen, ")"));
                i++;
                continue;
            }

            throw new FormatException($"非法字符 '{c}' 在表达式中");
        }

        return list;
    }

    // expression = term ((+|-) term)*
    private static Node ParseExpression(List<Tok> tokens, ref int pos)
    {
        var left = ParseTerm(tokens, ref pos);
        while (pos < tokens.Count && tokens[pos].Kind == TokKind.Op &&
               (tokens[pos].Text[0] is '+' or '-'))
        {
            var op = tokens[pos++].Text[0];
            var right = ParseTerm(tokens, ref pos);
            left = new BinaryNode(op, left, right);
        }

        return left;
    }

    // term = unary ((*|/) unary)*
    private static Node ParseTerm(List<Tok> tokens, ref int pos)
    {
        var left = ParseUnary(tokens, ref pos);
        while (pos < tokens.Count && tokens[pos].Kind == TokKind.Op &&
               (tokens[pos].Text[0] is '*' or '/'))
        {
            var op = tokens[pos++].Text[0];
            var right = ParseUnary(tokens, ref pos);
            left = new BinaryNode(op, left, right);
        }

        return left;
    }

    private static Node ParseUnary(List<Tok> tokens, ref int pos)
    {
        if (pos < tokens.Count && tokens[pos].Kind == TokKind.Op && tokens[pos].Text[0] is '+' or '-')
        {
            var op = tokens[pos++].Text[0];
            var inner = ParseUnary(tokens, ref pos);
            return op == '+' ? inner : new UnaryNode('-', inner);
        }

        return ParsePrimary(tokens, ref pos);
    }

    private static Node ParsePrimary(List<Tok> tokens, ref int pos)
    {
        if (pos >= tokens.Count)
            throw new FormatException("表达式不完整");

        var t = tokens[pos++];
        if (t.Kind == TokKind.Number)
            return new NumberNode(double.Parse(t.Text, CultureInfo.InvariantCulture));

        if (t.Kind == TokKind.Ident)
            return new IdentNode(t.Text);

        if (t.Kind == TokKind.LParen)
        {
            var inner = ParseExpression(tokens, ref pos);
            if (pos >= tokens.Count || tokens[pos].Kind != TokKind.RParen)
                throw new FormatException("缺少右括号");
            pos++;
            return inner;
        }

        throw new FormatException($"意外的 token: {t.Text}");
    }
}

// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using FluentAssertions;
using OfficeCli.Core;
using Xunit;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Tests.Core;

public class FormulaParserTests
{
    // Parse() uses WrapInOfficeMath which returns the element directly when
    // only a single node is produced. Descendants<T>() does not include self,
    // so we use FindFirst<T> that checks the root element as well.
    private static T FindFirst<T>(OpenXmlElement root) where T : OpenXmlElement
        => root as T ?? root.Descendants<T>().FirstOrDefault();

    private static IEnumerable<T> FindAll<T>(OpenXmlElement root) where T : OpenXmlElement
        => root is T self
            ? new[] { self }.Concat(root.Descendants<T>())
            : root.Descendants<T>();

    // ==================== RewriteOver (via round-trip) ====================

    [Fact]
    public void RewriteOver_BasicFraction_EquivalentToFrac()
    {
        var viaOver = FormulaParser.ToLatex(FormulaParser.Parse(@"{a \over b}"));
        var viaFrac = FormulaParser.ToLatex(FormulaParser.Parse(@"\frac{a}{b}"));
        viaOver.Should().Be(viaFrac);
    }

    [Fact]
    public void RewriteOver_NoOver_Unchanged()
    {
        var result = FormulaParser.ToLatex(FormulaParser.Parse(@"\frac{x}{y}"));
        result.Should().Be(@"\frac{x}{y}");
    }

    [Fact]
    public void RewriteOver_Nested_InnerFirst()
    {
        // {a \over {b \over c}} → \frac{a}{\frac{b}{c}}
        var result = FormulaParser.ToLatex(FormulaParser.Parse(@"{a \over {b \over c}}"));
        result.Should().Contain(@"\frac");
        result.Should().Contain("b").And.Contain("c");
    }

    [Fact]
    public void RewriteOver_MalformedNoOuterBrace_DoesNotThrow()
    {
        var act = () => FormulaParser.Parse(@"a \over b");
        act.Should().NotThrow();
    }

    // ==================== Parse: output element structure ====================

    [Fact]
    public void Parse_ReturnsNonNull()
    {
        var result = FormulaParser.Parse(@"\frac{1}{2}");
        result.Should().NotBeNull();
    }

    [Fact]
    public void Parse_EmptyString_ReturnsElement()
    {
        var result = FormulaParser.Parse("");
        result.Should().NotBeNull();
    }

    [Fact]
    public void Parse_SingleChar_ProducesRun()
    {
        var result = FormulaParser.Parse("x");
        var run = FindFirst<M.Run>(result);
        run.Should().NotBeNull();
        run!.InnerText.Should().Be("x");
    }

    [Fact]
    public void Parse_Fraction_ProducesFElement()
    {
        var result = FormulaParser.Parse(@"\frac{a}{b}");
        var fraction = FindFirst<M.Fraction>(result);
        fraction.Should().NotBeNull();
    }

    [Fact]
    public void Parse_Fraction_NumeratorDenominator()
    {
        var result = FormulaParser.Parse(@"\frac{1}{2}");
        var f = FindFirst<M.Fraction>(result)!;
        f.Numerator?.InnerText.Should().Be("1");
        f.Denominator?.InnerText.Should().Be("2");
    }

    [Fact]
    public void Parse_Subscript_SingleChar_Shorthand()
    {
        var result = FormulaParser.Parse("H_2");
        var sub = FindFirst<M.Subscript>(result);
        sub.Should().NotBeNull();
        sub!.SubArgument?.InnerText.Should().Be("2");
    }

    [Fact]
    public void Parse_Subscript_MultiChar()
    {
        var result = FormulaParser.Parse("x_{abc}");
        var sub = FindFirst<M.Subscript>(result);
        sub.Should().NotBeNull();
        sub!.SubArgument?.InnerText.Should().Be("abc");
    }

    [Fact]
    public void Parse_Superscript_SingleChar()
    {
        var result = FormulaParser.Parse("x^2");
        var sup = FindFirst<M.Superscript>(result);
        sup.Should().NotBeNull();
        sup!.SuperArgument?.InnerText.Should().Be("2");
    }

    [Fact]
    public void Parse_SubAndSup_ProducesSubscriptOrSubSuperscript()
    {
        var result = FormulaParser.Parse("x_i^2");
        var hasSubSup = FindFirst<M.SubSuperscript>(result) != null
            || (FindFirst<M.Subscript>(result) != null && FindFirst<M.Superscript>(result) != null);
        hasSubSup.Should().BeTrue();
    }

    [Fact]
    public void Parse_Sqrt_ProducesRadical()
    {
        var result = FormulaParser.Parse(@"\sqrt{x}");
        var rad = FindFirst<M.Radical>(result);
        rad.Should().NotBeNull();
    }

    [Fact]
    public void Parse_SqrtN_HasDegree()
    {
        var result = FormulaParser.Parse(@"\sqrt[3]{x}");
        var rad = FindFirst<M.Radical>(result)!;
        rad.Degree?.InnerText.Should().Be("3");
    }

    [Fact]
    public void Parse_Sqrt_DegreeIsHidden()
    {
        var result = FormulaParser.Parse(@"\sqrt{x}");
        var rad = FindFirst<M.Radical>(result)!;
        var radPr = rad.RadicalProperties;
        var degHide = radPr?.HideDegree;
        if (degHide != null)
        {
            var val = degHide.GetAttribute("val", "http://schemas.openxmlformats.org/officeDocument/2006/math").Value;
            (val == "1" || val == "true").Should().BeTrue();
        }
        else
        {
            rad.Degree?.InnerText.Should().BeNullOrEmpty();
        }
    }

    [Fact]
    public void Parse_Sum_ProducesNary()
    {
        var result = FormulaParser.Parse(@"\sum_{i=1}^{n}");
        var nary = FindFirst<M.Nary>(result);
        nary.Should().NotBeNull();
    }

    [Fact]
    public void Parse_Integral_ProducesNary()
    {
        var result = FormulaParser.Parse(@"\int_0^1 f(x)");
        var nary = FindFirst<M.Nary>(result);
        nary.Should().NotBeNull();
    }

    [Fact]
    public void Parse_GreekAlpha_ProducesAlphaChar()
    {
        var result = FormulaParser.Parse(@"\alpha");
        result.InnerText.Should().Contain("α");
    }

    [Fact]
    public void Parse_GreekPi_ProducesPiChar()
    {
        var result = FormulaParser.Parse(@"\pi");
        result.InnerText.Should().Contain("π");
    }

    [Fact]
    public void Parse_LeftRightDelimiters_ProducesDelimiter()
    {
        var result = FormulaParser.Parse(@"\left(\frac{a}{b}\right)");
        var delim = FindFirst<M.Delimiter>(result);
        delim.Should().NotBeNull();
    }

    [Fact]
    public void Parse_Text_ContainsText()
    {
        var result = FormulaParser.Parse(@"\text{hello}");
        result.InnerText.Should().Contain("hello");
    }

    [Fact]
    public void Parse_Overset_ProducesLimUpp()
    {
        var result = FormulaParser.Parse(@"\overset{\triangle}{\rightarrow}");
        var limUpp = FindFirst<M.LimitUpper>(result);
        limUpp.Should().NotBeNull();
    }

    [Fact]
    public void Parse_Underset_ProducesLimLow()
    {
        var result = FormulaParser.Parse(@"\underset{k}{\rightarrow}");
        var limLow = FindFirst<M.LimitLower>(result);
        limLow.Should().NotBeNull();
    }

    [Fact]
    public void Parse_Matrix_Pmatrix()
    {
        var result = FormulaParser.Parse(@"\begin{pmatrix} a & b \\ c & d \end{pmatrix}");
        var matrix = FindFirst<M.Matrix>(result);
        matrix.Should().NotBeNull();
        var rows = matrix!.Elements<M.MatrixRow>().ToList();
        rows.Should().HaveCount(2);
    }

    [Fact]
    public void Parse_NestedFraction()
    {
        var result = FormulaParser.Parse(@"\frac{\frac{a}{b}}{c}");
        // Outer fraction is the result itself, inner fraction is a descendant
        var fractions = FindAll<M.Fraction>(result).ToList();
        fractions.Should().HaveCountGreaterThanOrEqualTo(2);
    }

    [Fact]
    public void Parse_Overline_ProducesBar()
    {
        var result = FormulaParser.Parse(@"\overline{AB}");
        var bar = FindFirst<M.Bar>(result);
        bar.Should().NotBeNull();
    }

    [Fact]
    public void Parse_Hat_ProducesAccent()
    {
        var result = FormulaParser.Parse(@"\hat{x}");
        var accent = FindFirst<M.Accent>(result);
        accent.Should().NotBeNull();
    }

    // ==================== ToLatex round-trip ====================

    [Theory]
    [InlineData(@"\frac{a}{b}")]
    [InlineData(@"\frac{1}{2}")]
    [InlineData(@"\sqrt{x}")]
    [InlineData(@"\sqrt[3]{x}")]
    public void ToLatex_RoundTrip(string latex)
    {
        var omml = FormulaParser.Parse(latex);
        var result = FormulaParser.ToLatex(omml);
        result.Should().Be(latex);
    }

    [Fact]
    public void ToLatex_Subscript_RoundTrip()
    {
        var omml = FormulaParser.Parse("H_2");
        var result = FormulaParser.ToLatex(omml);
        result.Should().Be("H_2");
    }

    [Fact]
    public void ToLatex_Superscript_RoundTrip()
    {
        var omml = FormulaParser.Parse("x^2");
        var result = FormulaParser.ToLatex(omml);
        result.Should().Be("x^2");
    }

    [Fact]
    public void ToLatex_MultiCharSubscript_UsesBraces()
    {
        var omml = FormulaParser.Parse("x_{abc}");
        var result = FormulaParser.ToLatex(omml);
        result.Should().Be("x_{abc}");
    }

    [Fact]
    public void ToLatex_Sum_ContainsExpectedParts()
    {
        var latex = @"\sum_{i=1}^{n}";
        var omml = FormulaParser.Parse(latex);
        var result = FormulaParser.ToLatex(omml);
        result.Should().Contain(@"\sum");
        result.Should().Contain("i=1");
        result.Should().Contain("n");
    }

    [Fact]
    public void ToLatex_GreekLetter_DoesNotThrow()
    {
        var omml = FormulaParser.Parse(@"\alpha + \beta");
        var act = () => FormulaParser.ToLatex(omml);
        act.Should().NotThrow();
    }

    // ==================== ToReadableText ====================

    [Fact]
    public void ToReadableText_Fraction_IsNonEmpty()
    {
        var omml = FormulaParser.Parse(@"\frac{1}{2}");
        var text = FormulaParser.ToReadableText(omml);
        text.Should().NotBeNullOrEmpty();
    }

    [Theory]
    [InlineData(@"\frac{a}{b}")]
    [InlineData(@"\sqrt{x}")]
    [InlineData(@"\sum_{i=1}^n i^2")]
    [InlineData(@"\alpha\beta\gamma")]
    public void ToReadableText_DoesNotThrow(string input)
    {
        var omml = FormulaParser.Parse(input);
        var act = () => FormulaParser.ToReadableText(omml);
        act.Should().NotThrow(because: $"input: {input}");
    }
}

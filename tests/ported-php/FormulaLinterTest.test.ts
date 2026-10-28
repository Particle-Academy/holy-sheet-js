import { describe, it, expect } from "vitest";
import { Agent } from "../../src";

// Ported from PHP tests/Unit/FormulaLinterTest.php.

describe("formula linter (ported PHP FormulaLinterTest)", () => {
  it("returns no issues for a workbook with valid formulas", () => {
    const schema = {
      sheets: [
        {
          name: "Q4",
          rows: [
            ["Region", "Revenue", "Doubled"],
            ["NA", 100, { formula: "B2*2" }],
            ["EU", 200, { formula: "B3*2" }],
            ["Total", { formula: "SUM(B2:B3)" }, { formula: "SUM(C2:C3)" }],
          ],
        },
      ],
    };
    expect(Agent.lint(schema)).toEqual([]);
  });

  it("catches the header-row off-by-one bug and suggests the correct row", () => {
    const schema = {
      sheets: [
        {
          name: "Q4",
          rows: [
            ["Region", "Annual", "Monthly"],
            ["NA", 12000, { formula: "B1*12" }],
          ],
        },
      ],
    };
    const issues = Agent.lint(schema);
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#VALUE!");
    expect(issues[0].address).toBe("C2");
    expect(issues[0].hint).toContain('B1 = "Annual" (string)');
    expect(issues[0].hint).toContain("Did you mean B2");
  });

  it("catches division by zero", () => {
    const schema = {
      sheets: [{ name: "D", rows: [["x"], [0], [{ formula: "100/A2" }]] }],
    };
    const issues = Agent.lint(schema);
    expect(issues[0].error).toBe("#DIV/0!");
  });

  it("catches circular references", () => {
    const schema = {
      sheets: [
        {
          name: "C",
          rows: [["x"], [{ formula: "A1+1" }], [{ formula: "A2" }]],
        },
      ],
    };
    const issues = Agent.lint(schema);
    expect(issues.length).toBeGreaterThan(0);
  });

  it("detects true circular dependency (A1 = B1, B1 = A1)", () => {
    const schema = {
      sheets: [
        {
          name: "C",
          cells: {
            A1: { formula: "B1" },
            B1: { formula: "A1" },
          },
        },
      ],
    };
    const issues = Agent.lint(schema);
    expect(issues).not.toHaveLength(0);
    expect(issues[0].error).toBe("#CIRC!");
  });

  it("flags unknown function names as #NAME?", () => {
    const schema = {
      sheets: [{ name: "F", rows: [["x"], [{ formula: "BOGUSFN(1,2)" }]] }],
    };
    expect(Agent.lint(schema)[0].error).toBe("#NAME?");
  });

  it("evaluates cross-sheet references", () => {
    const schema = {
      sheets: [
        { name: "Detail", rows: [["x"], [100], [200]] },
        {
          name: "Summary",
          cells: {
            A1: { value: "Total" },
            B1: { formula: "SUM(Detail!A2:A3)" },
          },
        },
      ],
    };
    expect(Agent.lint(schema)).toEqual([]);
  });

  it("handles SUM across a numeric column with no errors", () => {
    const schema = {
      sheets: [
        {
          name: "S",
          rows: [["x"], [10], [20], [30], [{ formula: "SUM(A2:A4)" }], [{ formula: "AVERAGE(A2:A4)" }]],
        },
      ],
    };
    expect(Agent.lint(schema)).toEqual([]);
  });

  it("catches arithmetic on a string in the middle of an expression", () => {
    const schema = {
      sheets: [
        {
          name: "M",
          rows: [
            ["x", "y"],
            [10, "oops"],
            [20, 30],
            [{ formula: "A2+B2" }],
          ],
        },
      ],
    };
    const issues = Agent.lint(schema);
    expect(issues[0].error).toBe("#VALUE!");
    expect(issues[0].hint).toContain('B2 = "oops"');
  });

  // Sheet names that need quoting — holy-sheet issue #6. Ported case for case.
  const twoSheets = (first: string, formula: string) => ({
    sheets: [
      { name: first, columns: [{ header: "Deal Size", type: "number" }], rows: [[9407], [9750]] },
      { name: "Summary", columns: [{ header: "Total", type: "number" }], rows: [[{ formula }]] },
    ],
  });

  it("lints a reference to a quoted sheet name like an unquoted one", () => {
    expect(Agent.lint(twoSheets("My Earnings Projection", "SUM('My Earnings Projection'!A2:A3)"))).toEqual([]);
  });

  it("resolves a quoted sheet reference to the real cell", () => {
    const issues = Agent.lint(twoSheets("My Earnings Projection", "'My Earnings Projection'!A1*2"));
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#VALUE!");
    expect(issues[0].hint).toContain('A1 = "Deal Size" (string)');
    expect(issues[0].hint).toContain("Did you mean A2");
  });

  it("reads Excel's doubled quote as one literal quote in a sheet name", () => {
    expect(Agent.lint(twoSheets("Q3 'Final'", "SUM('Q3 ''Final'''!A2:A3)"))).toEqual([]);
  });

  it("reports #REF! for a quoted sheet that does not exist, naming it", () => {
    const issues = Agent.lint(twoSheets("Deals", "SUM('No Such Sheet'!A2:A3)"));
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#REF!");
    // Byte-for-byte the PHP and Python hint.
    expect(issues[0].hint).toBe(
      "The formula refers to a sheet named 'No Such Sheet', and this workbook has no such sheet. Its sheets are: Deals, Summary. Quote a name that contains spaces or punctuation: 'My Sheet'!A1.",
    );
  });

  it("reports #REF! for an unquoted sheet that does not exist", () => {
    const issues = Agent.lint(twoSheets("Deals", "SUM(Nope!A2:A3)"));
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#REF!");
  });

  it("matches sheet names case-insensitively, as Excel does", () => {
    expect(Agent.lint(twoSheets("Deals", "SUM(deals!A2:A3)"))).toEqual([]);
    expect(Agent.lint(twoSheets("My Deals", "SUM('MY DEALS'!A2:A3)"))).toEqual([]);
    const issues = Agent.lint(twoSheets("Deals", "DEALS!A1*2"));
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#VALUE!");
  });

  it("reports #NAME? for a quote that never closes", () => {
    const issues = Agent.lint(twoSheets("My Deals", "SUM('My Deals!A2:A3)"));
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#NAME?");
  });

  it("reports #NAME? for a quoted name that is not a sheet reference", () => {
    const issues = Agent.lint(twoSheets("My Deals", "'My Deals'+1"));
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#NAME?");
  });
});

/*
 * Excel semantics (fancy-formula plan, decision D17). Ported from the PHP
 * 2.11.1 tests, case for case, plus lazy IF.
 *
 * The linter only reports errors, so a value is observed through a division:
 * `1/(<claim>)` lints clean exactly when the claim is TRUE (1/1), and as
 * #DIV/0! when it is FALSE (1/0). Division rather than IF because it is the
 * smaller claim: it does not depend on IF being lazy, which is tested on its
 * own below. holy-sheet writes for Excel, so Excel decides.
 */

type Cells = Record<string, unknown>;

function lintFormula(formula: string, cells: Cells = {}) {
  return Agent.lint({ sheets: [{ name: "S", cells: { ...cells, Z1: { formula } } }] });
}

function lintClaim(claim: string, cells: Cells = {}) {
  return lintFormula(`1/(${claim})`, cells);
}

function z1(issues: ReturnType<typeof lintFormula>) {
  return issues.filter((i) => i.address === "Z1");
}

/** The tail of every syntax hint, byte for byte the PHP and Python text. */
const SYNTAX_TAIL =
  ". Excel cannot parse this formula. Check for an unbalanced parenthesis or quote, a missing operator or argument, or a character a formula cannot contain. Text goes in double quotes, and a quote inside text is written twice: " +
  '"a""b".';
const SUPPORTED =
  "Holy Sheet supports: SUM, AVERAGE, COUNT, COUNTA, MIN, MAX, IF, ROUND, ABS, LEN, UPPER, LOWER, CONCAT.";

describe("formula linter: Excel semantics (ported PHP 2.11.1)", () => {
  it("treats % as postfix percent, not a binary operator", () => {
    expect(lintClaim("50%=0.5")).toEqual([]);
    expect(lintClaim("A1%=0.5", { A1: { value: 50 } })).toEqual([]);
    expect(lintClaim("-50%=-0.5")).toEqual([]);
    expect(lintClaim("200%*3=6")).toEqual([]);
    expect(lintClaim("(1+1)%=0.02")).toEqual([]);
  });

  it("reports a binary use of % as an error, since Excel rejects it", () => {
    // The old parser read 10%2 as 10/100*2 and passed it.
    const issues = lintFormula("10%2");
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#NAME?");
  });

  it("binds ^ tighter than * and /", () => {
    expect(lintClaim("2*3^2=18")).toEqual([]);
    expect(lintClaim("10/2^2=2.5")).toEqual([]);
    expect(lintClaim("2^3*2=16")).toEqual([]);
    // Excel's ^ is left-associative, and negation binds tighter than ^.
    expect(lintClaim("2^3^2=64")).toEqual([]);
    expect(lintClaim("-2^2=4")).toEqual([]);
    expect(lintClaim("2^-1=0.5")).toEqual([]);
    // Percent binds tighter still: (50%)^2.
    expect(lintClaim("50%^2=0.25")).toEqual([]);
  });

  it("reads a doubled quote inside a text literal as one quote", () => {
    // ="a""b" is the three characters a"b.
    expect(lintClaim('LEN("a""b")=3')).toEqual([]);
    expect(lintClaim('"a""b"="a"&"""b"')).toEqual([]);
    expect(lintClaim('LEN("""")=1')).toEqual([]);
    expect(lintClaim('LEN("")=0')).toEqual([]);
  });

  it("reads TRUE and FALSE text as logicals in IF, in any case, as Excel does", () => {
    // The text "FALSE" WAS already falsy; that part matched Excel. What did not
    // was every other spelling: "false" counted as TRUE.
    expect(lintClaim('IF("FALSE",1,2)=2')).toEqual([]);
    expect(lintClaim('IF("false",1,2)=2')).toEqual([]);
    expect(lintClaim('IF("True",1,2)=1')).toEqual([]);
    expect(lintClaim("IF(A1,1,2)=2", { A1: { value: "FALSE" } })).toEqual([]);
    // An empty cell is FALSE, a non-zero number TRUE.
    expect(lintClaim("IF(A1,1,2)=2")).toEqual([]);
    expect(lintClaim("IF(-3,1,2)=1")).toEqual([]);
  });

  it("reports #VALUE! when IF is given text that is not TRUE or FALSE", () => {
    // The old linter called any other text TRUE and "" FALSE. Excel errors.
    for (const formula of ['IF("abc",1,2)', 'IF("",1,2)', "IF(A1,1,2)"]) {
      const issues = lintFormula(formula, { A1: { value: "yes" } });
      expect(issues, formula).toHaveLength(1);
      expect(issues[0].error).toBe("#VALUE!");
    }
  });

  it("propagates an error in the IF condition instead of choosing a branch", () => {
    const issues = z1(lintFormula("IF(A1,1,2)", { A1: { formula: "1/0" } }));
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#DIV/0!");
  });

  it("counts characters in LEN, not bytes", () => {
    expect(lintClaim('LEN("é")=1')).toEqual([]);
    expect(lintClaim('LEN("naïve café")=10')).toEqual([]);
    expect(lintClaim("LEN(A1)=2", { A1: { value: "日本" } })).toEqual([]);
    // Excel counts UTF-16 code units: a character outside the BMP is 2.
    expect(lintClaim('LEN("😀")=2')).toEqual([]);
  });

  it("compares numeric text as text, which Excel sorts above every number", () => {
    // ="10">9 is TRUE in Excel, but so is 10>9, so it cannot tell the two
    // readings apart. "5">9 can.
    expect(lintClaim('"5">9')).toEqual([]);
    expect(lintClaim('"10">9')).toEqual([]);
    expect(lintClaim('"10"<>10')).toEqual([]);
    expect(lintClaim('"2">"10"')).toEqual([]);
    expect(lintClaim("A1>99", { A1: { value: "10" } })).toEqual([]);
  });

  it("compares the rest as Excel does: text ignores case, logicals sort last, blanks match", () => {
    expect(lintClaim('"a"="A"')).toEqual([]);
    expect(lintClaim('"abc"<"ABD"')).toEqual([]);
    expect(lintClaim("TRUE>1")).toEqual([]);
    expect(lintClaim('TRUE>"zzz"')).toEqual([]);
    expect(lintClaim("FALSE<TRUE")).toEqual([]);
    expect(lintClaim("A1=0")).toEqual([]);
    expect(lintClaim('A1=""')).toEqual([]);
    expect(lintClaim("A1=FALSE")).toEqual([]);
  });

  it("orders text by code point, as PHP and Python do", () => {
    // U+1F600 sorts above U+FF5E by code point; UTF-16's `<` would put its
    // leading surrogate (0xD83D) below 0xFF5E and answer the other way.
    expect(lintClaim('"😀">"～"')).toEqual([]);
  });

  it("tokenizes with PHP's ASCII character classes", () => {
    // A no-break space is not whitespace to PHP's ctype_space, and "é" is not
    // a letter to ctype_alpha: both are syntax errors there, so here too.
    for (const formula of ["1" + String.fromCharCode(160) + "+1", "é1+1"]) {
      const issues = lintFormula(formula);
      expect(issues, formula).toHaveLength(1);
      expect(issues[0].hint, formula).toBe("Syntax error: unexpected non-ASCII character outside a text literal" + SYNTAX_TAIL);
    }
  });

  it("propagates an error through a comparison", () => {
    const issues = z1(lintFormula("A1>0", { A1: { formula: "1/0" } }));
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#DIV/0!");
  });

  it("says a malformed formula is a syntax error, never an unknown function", () => {
    // #NAME? stays the label. The hint used to be one sentence for every
    // cause, "references an unknown function or has a syntax error", which
    // sent an agent hunting for a function that was fine.
    const malformed = ["SUM(1,2", "(1+2", "1 2", "1+", "1 # 2", "10%2", '"abc', "'S'+1", "SUM(1,,"];
    for (const formula of malformed) {
      const issues = lintFormula(formula);
      expect(issues, formula).toHaveLength(1);
      expect(issues[0].error).toBe("#NAME?");
      expect(issues[0].hint.startsWith("Syntax error: "), formula).toBe(true);
      expect(issues[0].hint.toLowerCase()).not.toContain("unknown function");
    }
  });

  it("names the unknown function in its hint", () => {
    const issues = lintFormula("VLOOKUP(1, A1:B2, 2)");
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#NAME?");
    expect(issues[0].hint.startsWith("Unknown function VLOOKUP.")).toBe(true);
    expect(issues[0].hint).toContain("SUM, AVERAGE");
  });

  it("reports an unknown function even in an IF branch that is not taken", () => {
    // Excel resolves names before it evaluates. The old linter returned
    // #NAME? as a VALUE, and IF discarded it.
    const issues = lintFormula("IF(FALSE, BOGUSFN(1), 1)");
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#NAME?");
    expect(issues[0].hint.startsWith("Unknown function BOGUSFN.")).toBe(true);
  });

  it("points an inherited #NAME? at the cell it came from", () => {
    const issues = z1(lintFormula("A1+1", { A1: { formula: "BOGUSFN(1)" } }));
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#NAME?");
    expect(issues[0].hint).toContain("A1");
    expect(issues[0].hint.toLowerCase()).not.toContain("unknown function");
  });

  it("rounds half away from zero on the decimal value, as Excel ROUND does", () => {
    // A GUARD in PHP, where round() already gives Excel's answer. Here it was a
    // defect: Math.round(x * 100) rounds the binary double, so 1.005 (which
    // is 1.00499999... as a double) rounded to 1.
    const cases = [
      "ROUND(2.675,2)=2.68", "ROUND(-2.675,2)=-2.68", "ROUND(1.005,2)=1.01",
      "ROUND(0.285,2)=0.29", "ROUND(8.325,2)=8.33", "ROUND(2.5,0)=3",
      "ROUND(-2.5,0)=-3", "ROUND(-0.5,0)=-1", "ROUND(1250,-2)=1300",
      "ROUND(-1250,-2)=-1300", "ROUND(1.45,1)=1.5",
    ];
    for (const claim of cases) {
      expect(lintClaim(claim), claim).toEqual([]);
    }
  });

  it("reports a text literal that never closes", () => {
    // The old tokenizer ran the literal to the end of the formula and passed it.
    const issues = lintFormula('"abc');
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#NAME?");
  });

  it("writes the #NAME? hints byte for byte as PHP and Python do", () => {
    const cases: Array<[string, string]> = [
      ["SUM(1,2", "Syntax error: expected ')', but the formula ends" + SYNTAX_TAIL],
      ["1 2", "Syntax error: unexpected number 2 after a complete expression" + SYNTAX_TAIL],
      ["1+", "Syntax error: the formula ends where a value was expected" + SYNTAX_TAIL],
      ["1 # 2", "Syntax error: unexpected character '#'" + SYNTAX_TAIL],
      ["1 é 2", "Syntax error: unexpected non-ASCII character outside a text literal" + SYNTAX_TAIL],
      ['"abc', 'Syntax error: a text literal opens with " and never closes' + SYNTAX_TAIL],
      ["'S'+1", "Syntax error: the quoted name 'S' is not followed by ! and a cell reference" + SYNTAX_TAIL],
      ["SUM(1,,", "Syntax error: unexpected ',' where a value was expected" + SYNTAX_TAIL],
      ['"a""b" 3', "Syntax error: unexpected number 3 after a complete expression" + SYNTAX_TAIL],
      ["VLOOKUP(1, A2:B3, 2)", "Unknown function VLOOKUP. " + SUPPORTED],
      ["A1+1", "This formula depends on A1, whose value is #NAME?. Fix that cell first."],
    ];
    for (const [formula, hint] of cases) {
      const issues = z1(lintFormula(formula, { A1: { formula: "BOGUSFN(1)" } }));
      expect(issues, formula).toHaveLength(1);
      expect(issues[0].hint, formula).toBe(hint);
    }
  });
});

/*
 * IF is lazy, as in Excel: only the branch it takes is evaluated. The linter
 * evaluates while it parses, so the untaken branch is still PARSED (a syntax
 * error or an unknown function there is still reported) but never EVALUATED.
 */
describe("formula linter: lazy IF", () => {
  it("does not evaluate the branch IF does not take", () => {
    // The guard every spreadsheet writes. Eager IF reported #DIV/0! on it.
    expect(lintFormula("IF(A2=0,0,1/A2)", { A2: { value: 0 } })).toEqual([]);
    expect(lintFormula("IF(TRUE,1,1/0)")).toEqual([]);
    expect(lintFormula("IF(FALSE,1/0,1)")).toEqual([]);
    expect(lintFormula("IF(1,2,1/0)+1")).toEqual([]);
    // An untaken branch reads no cells, so their text and errors stay out.
    expect(lintFormula("IF(TRUE,1,A1*2)", { A1: { value: "Annual" } })).toEqual([]);
    expect(lintFormula("IF(FALSE,Nope!A1,1)")).toEqual([]);
    expect(lintFormula("IF(TRUE,1,foo)")).toEqual([]);
    expect(lintFormula("IF(TRUE,1,A1:foo)")).toEqual([]);
  });

  it("still evaluates the branch IF takes", () => {
    for (const formula of ["IF(TRUE,1/0,1)", "IF(FALSE,1,1/0)", "IF(A2=0,1/A2,0)"]) {
      const issues = lintFormula(formula, { A2: { value: 0 } });
      expect(issues, formula).toHaveLength(1);
      expect(issues[0].error).toBe("#DIV/0!");
    }
  });

  it("reads an untaken IF branch without evaluating anything inside it, nested IFs included", () => {
    expect(lintFormula("IF(TRUE,1,IF(1/0,1/0,1/0))")).toEqual([]);
    expect(lintFormula("IF(FALSE,IF(1/0,1,2),3)")).toEqual([]);
    expect(lintFormula('IF(TRUE,1,SUM(1/0,"x"*2)&-("y")%)')).toEqual([]);
    expect(lintFormula("IF(TRUE,IF(FALSE,1/0,2),3)")).toEqual([]);
  });

  it("takes the else branch as FALSE when IF has none", () => {
    // =IF(FALSE,1) is FALSE in Excel, and FALSE is not an error.
    expect(lintFormula("IF(FALSE,1/0)")).toEqual([]);
    // Through &, where FALSE is "FALSE" and a blank is "": IF(...)=FALSE
    // cannot tell them apart, since a blank equals FALSE.
    expect(lintClaim('IF(FALSE,1)&""="FALSE"')).toEqual([]);
  });

  it("still reports a syntax error inside an untaken IF branch", () => {
    for (const formula of ["IF(TRUE,1,2+)", "IF(FALSE,(1,2)", 'IF(TRUE,1,"abc)', "IF(TRUE,1,2 3)"]) {
      const issues = lintFormula(formula);
      expect(issues, formula).toHaveLength(1);
      expect(issues[0].error).toBe("#NAME?");
      expect(issues[0].hint.startsWith("Syntax error: "), formula).toBe(true);
    }
  });

  it("parses the whole IF before an error in its condition propagates", () => {
    // Excel refuses a formula it cannot parse, so the syntax error is the one
    // to report. (A computed error such as 1/0 still stops the parse where it
    // arises, in IF as anywhere else in a formula; this is the error the
    // condition evaluates to.)
    const issues = z1(lintFormula("IF(A1,1,2+)", { A1: { formula: "1/0" } }));
    expect(issues).toHaveLength(1);
    expect(issues[0].error).toBe("#NAME?");
    expect(issues[0].hint.startsWith("Syntax error: ")).toBe(true);
  });

  it("still reports #VALUE! for an IF with fewer than two arguments", () => {
    for (const formula of ["IF(TRUE)", "IF()"]) {
      const issues = lintFormula(formula);
      expect(issues, formula).toHaveLength(1);
      expect(issues[0].error).toBe("#VALUE!");
    }
  });
});

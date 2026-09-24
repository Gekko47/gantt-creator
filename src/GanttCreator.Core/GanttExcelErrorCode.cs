namespace GanttCreator.Core;

/// <summary>Neutral Excel error codes without a dependency on Office PIA types.</summary>
public enum GanttExcelErrorCode
{
    /// <summary>An error code not known by this application version.</summary>
    Unknown = 0,

    /// <summary><c>#NULL!</c></summary>
    Null = 2000,

    /// <summary><c>#DIV/0!</c></summary>
    DivisionByZero = 2007,

    /// <summary><c>#VALUE!</c></summary>
    Value = 2015,

    /// <summary><c>#REF!</c></summary>
    Reference = 2023,

    /// <summary><c>#NAME?</c></summary>
    Name = 2029,

    /// <summary><c>#NUM!</c></summary>
    Number = 2036,

    /// <summary><c>#N/A</c></summary>
    NotAvailable = 2042,

    /// <summary><c>#GETTING_DATA</c></summary>
    GettingData = 2043,

    /// <summary><c>#SPILL!</c></summary>
    Spill = 2045,

    /// <summary><c>#CONNECT!</c></summary>
    Connect = 2046,

    /// <summary><c>#BLOCKED!</c></summary>
    Blocked = 2047,

    /// <summary><c>#UNKNOWN!</c></summary>
    ExcelUnknown = 2048,

    /// <summary><c>#FIELD!</c></summary>
    Field = 2049,

    /// <summary><c>#CALC!</c></summary>
    Calculation = 2050,
}

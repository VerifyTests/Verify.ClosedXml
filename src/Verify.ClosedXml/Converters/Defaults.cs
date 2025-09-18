static class Defaults
{
    public static bool IsDefault(this IXLNumberFormat format) =>
        format is {NumberFormatId: 0, Format: ""};

    public static bool IsDefault(this IXLProtection protection) =>
        protection is {Locked: true, Hidden: false};

    public static bool IsDefault(this IXLAlignment alignment) =>
        alignment.Horizontal == XLAlignmentHorizontalValues.General &&
        alignment is
        {
            JustifyLastLine: false,
            ShrinkToFit: false,
            WrapText: false,
            TopToBottom: false
        };

    public static bool IsDefault(this IXLFill fill) =>
        fill.BackgroundColor.Color == System.Drawing.Color.Transparent &&
        fill.PatternColor.Color == System.Drawing.Color.Transparent &&
        fill.PatternType == XLFillPatternValues.None;

    public static bool IsDefault(this IXLBorder border) =>
        border.LeftBorder == XLBorderStyleValues.None &&
        border.LeftBorderColor == XLColor.Black &&
        border.RightBorder == XLBorderStyleValues.None &&
        border.RightBorderColor == XLColor.Black &&
        border.TopBorder == XLBorderStyleValues.None &&
        border.TopBorderColor == XLColor.Black &&
        border.BottomBorder == XLBorderStyleValues.None &&
        border.BottomBorderColor == XLColor.Black &&
        border.DiagonalBorder == XLBorderStyleValues.None &&
        border.DiagonalBorderColor == XLColor.Black &&
        border is {DiagonalUp: false, DiagonalDown: false};
}
package org.doctribute.tools;

public class FontInfo {

    private final String fontFamily;
    private final double fontSize;
    private final String fontStyle;
    private final String fontWeight;
    private final String fillColor;

    public FontInfo() {
        this.fontFamily = "Calibri";
        this.fontSize = 10.0;
        this.fontStyle = "normal";
        this.fontWeight = "bold";
        this.fillColor = "#000000";
    }

    public FontInfo(
            String fontFamily, double fontSize, String fontStyle, String fontWeight, String fillColor) {

        this.fontFamily = fontFamily;
        this.fontSize = fontSize;
        this.fontStyle = fontStyle;
        this.fontWeight = fontWeight;
        this.fillColor = fillColor;
    }

    public String getFontFamily() {
        return fontFamily;
    }

    public double getFontSize() {
        return fontSize;
    }

    public String getFontStyle() {
        return fontStyle;
    }

    public String getFontWeight() {
        return fontWeight;
    }

    public String getFillColor() {
        return fillColor;
    }

}

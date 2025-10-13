/*
 * Copyright (C) 2025 Glavo
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */
package org.glavo.meow;

import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHighlightColor;

import java.util.Locale;
import java.util.Objects;
import java.util.Set;

public final class Meow {
    public static final boolean DEBUG = "true".equals(System.getProperty("meow.debug"));

    private static final Set<String> MEOW = Set.of(
            "喵", "喵呜", // zh_CN
            "にゃん", "にゃ", "nya", "nyan", // ja_JP
            "야옹", // ko_KR
            "мяу", "мјау", // ru_RU
            "miyav", // tr_TR
            "miau", // de_DE
            "miaou", // fr_FR
            "miauw", // nl_NL
            "μιάου", // el_GR
            "mjau", "mjäu", // sv_SE
            "nyávog", // hu_HU
            "म्याऊं", // hi_IN
            "meo meo", // vi_VN
            "mjá", // is_IS
            "mjav", // no_NO
            "mjal", // sq_AL
            "მიაუ", // ka_GE
            "myau", // lv_LV
            "ngiyavuma", // zu_ZA
            "ማው", // am_ET
            "meow", "maw", // en_US
            "meong", // fi_FI
            "miao" // it_IT
    );

    public static Meow of(XWPFRun term) {
        if (term == null) {
            return null;
        }

        if (!term.getEmbeddedPictures().isEmpty()) {
            return null;
        }

        if (MEOW.contains(term.text().trim().toLowerCase(Locale.ROOT))) {
            return new Meow(term);
        } else {
            return null;
        }
    }

    public static Meow builtin(int color, STHighlightColor.Enum highlightColor) {
        return new Meow(
                false,
                UnderlinePatterns.NONE,
                color,
                highlightColor,
                false,
                null,
                null
        );
    }

    private final String rawName;
    private final boolean isBracket;
    private final UnderlinePatterns underline;
    private final int color;
    private final STHighlightColor.Enum highlightColor;
    private final boolean isBold;
    private final String font;
    private final Double fontSize;

    public Meow(XWPFRun term) {
        this.rawName = term.text();
        this.isBracket = term.isItalic();
        this.color = term.getColor() == null ? 0 : Integer.parseInt(term.getColor(), 16);
        this.highlightColor = term.getTextHighlightColor();
        this.isBold = term.isBold();
        this.underline = term.getUnderline();
        this.font = term.getFontFamily();
        this.fontSize = term.getFontSizeAsDouble();
    }

    public Meow(boolean isBracket, UnderlinePatterns underline, int color, STHighlightColor.Enum highlightColor, boolean isBold, String font, Double fontSize) {
        this.rawName = "";
        this.isBracket = isBracket;
        this.underline = underline;
        this.color = color;
        this.highlightColor = highlightColor;
        this.isBold = isBold;
        this.font = font;
        this.fontSize = fontSize;
    }


    public boolean isBracket() {
        return isBracket;
    }

    public UnderlinePatterns getUnderline() {
        return underline;
    }

    public int getColor() {
        return color;
    }

    public STHighlightColor.Enum getHighlightColor() {
        return highlightColor;
    }

    public String getFont() {
        return font;
    }

    public Double getFontSize() {
        return fontSize;
    }

    @Override
    public int hashCode() {
        return Objects.hash(isBracket, color, isBold, highlightColor, underline, font, fontSize);
    }

    @Override
    public boolean equals(Object obj) {
        return obj instanceof Meow that
               && this.isBracket == that.isBracket
               && this.color == that.color
               && this.highlightColor == that.highlightColor
               && this.isBold == that.isBold
               && this.underline == that.underline
               && Objects.equals(this.font, that.font)
               && Objects.equals(this.fontSize, that.fontSize);
    }

    @Override
    public String toString() {
        return "Meow[%s, isBracket=%s, color=%06X, highlightColor=%s, isBold=%s, underline=%s, font=%s, fontSize=%s]"
                .formatted(rawName, isBracket, color, highlightColor, isBold, underline, font, fontSize);
    }
}

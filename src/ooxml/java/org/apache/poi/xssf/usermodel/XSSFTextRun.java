/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package org.apache.poi.xssf.usermodel;

import java.awt.Color;

import org.apache.poi.ooxml.util.POIXMLUnits;
import org.apache.poi.sl.usermodel.TextRun;
import org.apache.poi.util.Removal;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.text.StrikeType;
import org.apache.poi.xddf.usermodel.text.UnderlineType;
import org.apache.poi.xddf.usermodel.text.XDDFTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSRgbColor;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSolidColorFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextField;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextLineBreak;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextNormalAutofit;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextStrikeType;
import org.openxmlformats.schemas.drawingml.x2006.main.STTextUnderlineType;

/**
 * Represents a run of text within the containing text body. The run element is the
 * lowest level text separation mechanism within a text body.
 */
public class XSSFTextRun extends XDDFTextRun {

    XSSFTextRun(CTTextLineBreak r, XSSFTextParagraph p){
        super(r, p);
    }

    XSSFTextRun(CTTextField r, XSSFTextParagraph p){
        super(r, p);
    }

    XSSFTextRun(CTRegularTextRun r, XSSFTextParagraph p){
        super(r, p);
    }

    public void setFontColor(Color color){
        CTTextCharacterProperties rPr = getRPr();
        CTSolidColorFillProperties fill = rPr.isSetSolidFill() ? rPr.getSolidFill() : rPr.addNewSolidFill();
        CTSRgbColor clr = fill.isSetSrgbClr() ? fill.getSrgbClr() : fill.addNewSrgbClr();
        clr.setVal(new byte[]{(byte)color.getRed(), (byte)color.getGreen(), (byte)color.getBlue()});

        if(fill.isSetHslClr()) fill.unsetHslClr();
        if(fill.isSetPrstClr()) fill.unsetPrstClr();
        if(fill.isSetSchemeClr()) fill.unsetSchemeClr();
        if(fill.isSetScrgbClr()) fill.unsetScrgbClr();
        if(fill.isSetSysClr()) fill.unsetSysClr();

    }

    public Color getFontColor(){

        CTTextCharacterProperties rPr = getRPr();
        if(rPr.isSetSolidFill()){
            CTSolidColorFillProperties fill = rPr.getSolidFill();

            if(fill.isSetSrgbClr()){
                CTSRgbColor clr = fill.getSrgbClr();
                byte[] rgb = clr.getVal();
                return new Color(0xFF & rgb[0], 0xFF & rgb[1], 0xFF & rgb[2]);
            }
        }

        return new Color(0, 0, 0);
    }

    /**
     * Specifies the typeface, or name of the font that is to be used for this text run.
     *
     * @param typeface  the font to apply to this text run.
     * The value of <code>null</code> unsets the Typeface attribute from the underlying xml.
     */
    public void setFont(String typeface){
        setFontFamily(typeface, (byte)-1, (byte)-1, false);
    }

    public void setFontFamily(String typeface, byte charset, byte pictAndFamily, boolean isSymbol){
        CTTextCharacterProperties rPr = getRPr();

        if(typeface == null){
            if(rPr.isSetLatin()) rPr.unsetLatin();
            if(rPr.isSetCs()) rPr.unsetCs();
            if(rPr.isSetSym()) rPr.unsetSym();
        } else {
            if(isSymbol){
                CTTextFont font = rPr.isSetSym() ? rPr.getSym() : rPr.addNewSym();
                font.setTypeface(typeface);
            } else {
                CTTextFont latin = rPr.isSetLatin() ? rPr.getLatin() : rPr.addNewLatin();
                latin.setTypeface(typeface);
                if(charset != -1) latin.setCharset(charset);
                if(pictAndFamily != -1) latin.setPitchFamily(pictAndFamily);
            }
        }
    }

    /**
     * @return  font family or null if not set
     */
    public String getFontFamily(){
        CTTextCharacterProperties rPr = getRPr();
        CTTextFont font = rPr.getLatin();
        if(font != null){
            return font.getTypeface();
        }
        return XSSFFont.DEFAULT_FONT_NAME;
    }

    public byte getPitchAndFamily(){
        CTTextCharacterProperties rPr = getRPr();
        CTTextFont font = rPr.getLatin();
        if(font != null){
            return font.getPitchFamily();
        }
        return 0;
    }

    /**
     * Specifies whether a run of text will be formatted as single strikethrough text.
     *
     * @param strike whether a run of text will be formatted as single strikethrough text.
     *
     * @deprecated prefer {@link #setStrikeThrough(StrikeType)}
     */
    @Deprecated
    @Removal(version = "6.0.0")
    public void setStrikethrough(boolean strike) {
        getRPr().setStrike(strike ? STTextStrikeType.SNG_STRIKE : STTextStrikeType.NO_STRIKE);
    }

    /**
     *  Set the baseline for both the superscript and subscript fonts.
     *  <p>
     *     The size is specified using a percentage.
     *     Positive values indicate superscript, negative values indicate subscript.
     *  </p>
     *
     * @param baselineOffset
     *
     * @deprecated prefer {@link #setBaseline(Double)}
     */
    @Deprecated
    @Removal(version = "6.0.0")
    public void setBaselineOffset(double baselineOffset){
        setBaseline(baselineOffset);
    }

    /**
     * Set whether the text in this run is formatted as superscript.
     * Default base line offset is 30%
     *
     * @deprecated prefer {@link #setSuperscript(Double)}
     */
    @Deprecated
    @Removal(version = "6.0.0")
    public void setSuperscript(boolean flag){
        setSuperscript(flag ? 30.0 : null);
    }

    /**
     * Set whether the text in this run is formatted as subscript.
     * Default base line offset is -25%.
     *
     * @deprecated prefer {@link #setSubscript(Double)}
     */
    @Deprecated
    @Removal(version = "6.0.0")
    public void setSubscript(boolean flag){
        setSubscript(flag ? -25.0 : null);
    }

    /**
     * @return whether a run of text will be formatted as a small caps, all caps or normal text.
     *
     * @deprecated prefer {@link #getCapitals()}
     */
    @Deprecated
    @Removal(version = "6.0.0")
    public TextCap getTextCap() {
        switch (getCapitals()) {
            case ALL: return TextCap.ALL;
            case SMALL: return TextCap.SMALL;
            default: return TextCap.NONE;
        }
    }

    /**
     * @param underline whether this run of text is formatted as underlined text
     *
     * @deprecated use {@link #setUnderline(UnderlineType)} instead
     */
    @Deprecated
    @Removal(version = "6.0.0")
    public void setUnderline(boolean underline) {
        super.setUnderline(underline ? UnderlineType.SINGLE : UnderlineType.NONE);
    }

    protected CTTextCharacterProperties getRPr(){
        return getOrCreateProps();
    }

    @Override
    public String toString(){
        return "[" + getClass() + "]" + getText();
    }
}

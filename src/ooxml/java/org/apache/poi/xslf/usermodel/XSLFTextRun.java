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
package org.apache.poi.xslf.usermodel;

import java.awt.Color;
import java.util.function.Consumer;

import org.apache.poi.common.usermodel.fonts.FontCharset;
import org.apache.poi.common.usermodel.fonts.FontFamily;
import org.apache.poi.common.usermodel.fonts.FontGroup;
import org.apache.poi.common.usermodel.fonts.FontInfo;
import org.apache.poi.common.usermodel.fonts.FontPitch;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.sl.draw.DrawPaint;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.sl.usermodel.PaintStyle.SolidPaint;
import org.apache.poi.sl.usermodel.TextRun;
import org.apache.poi.util.Beta;
import org.apache.poi.util.Internal;
import org.apache.poi.util.POILogFactory;
import org.apache.poi.util.POILogger;
import org.apache.poi.util.Removal;
import org.apache.poi.xddf.usermodel.text.CapitalsType;
import org.apache.poi.xddf.usermodel.text.StrikeType;
import org.apache.poi.xddf.usermodel.text.UnderlineType;
import org.apache.poi.xddf.usermodel.text.XDDFTextRun;
import org.apache.poi.xslf.model.CharacterPropertyFetcher;
import org.apache.poi.xslf.model.CharacterPropertyFetcher.CharPropFetcher;
import org.apache.poi.xslf.usermodel.XSLFPropertiesDelegate.XSLFFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTFontCollection;
import org.openxmlformats.schemas.drawingml.x2006.main.CTFontScheme;
import org.openxmlformats.schemas.drawingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSchemeColor;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeStyle;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSolidColorFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextField;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextLineBreak;

/**
 * Represents a run of text within the containing text body. The run element is the
 * lowest level text separation mechanism within a text body.
 */
@Beta
public class XSLFTextRun extends XDDFTextRun implements TextRun {
    private static final POILogger LOG = POILogFactory.getLogger(XSLFTextRun.class);

    protected XSLFTextRun(CTTextLineBreak run, XSLFTextParagraph parent) {
        super(run, parent);
    }

    protected XSLFTextRun(CTTextField run, XSLFTextParagraph parent) {
        super(run, parent);
    }

    protected XSLFTextRun(CTRegularTextRun run, XSLFTextParagraph parent) {
        super(run, parent);
    }

    @Override
    public String getRawText(){
        return super.getText();
    }

    @Override
    public void setFontColor(Color color) {
        setFontColor(DrawPaint.createSolidPaint(color));
    }

    @Override
    public void setFontColor(PaintStyle color) {
        if (!(color instanceof SolidPaint)) {
            LOG.log(POILogger.WARN, "Currently only SolidPaint is supported!");
            return;
        }
        SolidPaint sp = (SolidPaint)color;
        Color c = DrawPaint.applyColorTransform(sp.getSolidColor());

        CTTextCharacterProperties rPr = getRPr(true);
        CTSolidColorFillProperties fill = rPr.isSetSolidFill() ? rPr.getSolidFill() : rPr.addNewSolidFill();

        XSLFSheet sheet = getParagraph().getParentShape().getSheet();
        XSLFColor col = new XSLFColor(fill, sheet.getTheme(), fill.getSchemeClr(), sheet);
        col.setColor(c);
    }

    @Override
    public PaintStyle getFontColor(){
        XSLFShape shape = getParagraph().getParentShape();
        final boolean hasPlaceholder = shape.getPlaceholder() != null;
        return fetchCharacterProperty((props, val) -> fetchFontColor(props, val, shape, hasPlaceholder));
    }

    private static void fetchFontColor(CTTextCharacterProperties props, Consumer<PaintStyle> val, XSLFShape shape, boolean hasPlaceholder) {
        if (props == null) {
            return;
        }

        CTShapeStyle style = shape.getSpStyle();
        CTSchemeColor phClr = null;
        if (style != null && style.getFontRef() != null) {
            phClr = style.getFontRef().getSchemeClr();
        }

        XSLFFillProperties fp = XSLFPropertiesDelegate.getFillDelegate(props);
        XSLFSheet sheet = shape.getSheet();
        PackagePart pp = sheet.getPackagePart();
        XSLFTheme theme = sheet.getTheme();
        PaintStyle ps = shape.selectPaint(fp, phClr, pp, theme, hasPlaceholder);

        if (ps != null)  {
            val.accept(ps);
        }
    }

    @Override
    public void setFontFamily(String typeface) {
        FontGroup fg = FontGroup.getFontGroupFirst(getRawText());
        new XSLFFontInfo(fg).setTypeface(typeface);
    }

    @Override
    public void setFontFamily(String typeface, FontGroup fontGroup) {
        new XSLFFontInfo(fontGroup).setTypeface(typeface);
    }

    @Override
    public void setFontInfo(FontInfo fontInfo, FontGroup fontGroup) {
        new XSLFFontInfo(fontGroup).copyFrom(fontInfo);
    }

    @Override
    public String getFontFamily() {
        FontGroup fg = FontGroup.getFontGroupFirst(getRawText());
        return new XSLFFontInfo(fg).getTypeface();
    }

    @Override
    public String getFontFamily(FontGroup fontGroup) {
        return new XSLFFontInfo(fontGroup).getTypeface();
    }

    @Override
    public FontInfo getFontInfo(FontGroup fontGroup) {
        XSLFFontInfo fontInfo = new XSLFFontInfo(fontGroup);
        return (fontInfo.getTypeface() != null) ? fontInfo : null;
    }

    @Override
    public byte getPitchAndFamily(){
        FontGroup fg = FontGroup.getFontGroupFirst(getRawText());
        XSLFFontInfo fontInfo = new XSLFFontInfo(fg);
        FontPitch pitch = fontInfo.getPitch();
        if (pitch == null) {
            pitch = FontPitch.VARIABLE;
        }
        FontFamily family = fontInfo.getFamily();
        if (family == null) {
            family = FontFamily.FF_SWISS;
        }
        return FontPitch.getNativeId(pitch, family);
    }

    /**
     * @deprecated use {@link #setStrikeThrough(StrikeType)} instead in order to use {@link StrikeType#DOUBLE_STRIKE}
     * @param strike whether the text run has a single strike or no strike.
     */
    @Override
    @Deprecated
    @Removal(version = "6.0.0")
    public void setStrikeThrough(boolean strike) {
        super.setStrikeThrough(strike ? StrikeType.SINGLE_STRIKE : StrikeType.NO_STRIKE);
    }

    /**
     * Set whether the text in this run is formatted as superscript.
     * Default base line offset is 30%
     *
     * @see #setBaseline(Double)
     */
    @SuppressWarnings("WeakerAccess")
    public void setSuperscript(boolean flag){
        setBaseline(flag ? 30. : 0.);
    }

    /**
     * Set whether the text in this run is formatted as subscript.
     * Default base line offset is -25%.
     *
     * @see #setBaseline(Double)
     */
    @SuppressWarnings("WeakerAccess")
    public void setSubscript(boolean flag){
        setBaseline(flag ? -25.0 : 0.);
    }

    /**
     * @return whether a run of text will be formatted as capitals text.
     */
    @Override
    public TextCapitals getTextCapitals() {
        switch (getCapitals()) {
            case ALL: return TextCapitals.ALL;
            case SMALL: return TextCapitals.SMALL;
            default: return TextCapitals.NONE;
        }
    }

    @Override
    public void setBold(boolean bold){
        super.setBold(bold);
    }

    @Override
    public void setItalic(boolean italic){
        super.setItalic(italic);
    }

    /**
     * @deprecated use {@link #setUnderline(UnderlineType)} instead
     */
    @Deprecated
    @Removal(version = "6.0.0")
    @Override
    public void setUnderlined(boolean underline) {
        super.setUnderline(underline ? UnderlineType.SINGLE : UnderlineType.NONE);
    }

    /**
     * Return the character properties
     *
     * @param create if true, create an empty character properties object if it doesn't exist
     * @return the character properties or null if create was false and the properties haven't exist
     */
    @Internal
    CTTextCharacterProperties getRPr(boolean create) {
        if (create) {
            return getOrCreateProps();
        } else {
            return getProperties();
        }
    }

    @Override
    public String toString(){
        return "[" + getClass() + "]" + getRawText();
    }

    @Override
    public XSLFHyperlink createHyperlink(){
        XSLFHyperlink hl = getHyperlink();
        if (hl != null) {
            return hl;
        }

        CTTextCharacterProperties rPr = getRPr(true);
        return new XSLFHyperlink(rPr.addNewHlinkClick(), getParagraph().getParentShape().getSheet());
    }

    @Override
    public XSLFHyperlink getHyperlink(){
        CTTextCharacterProperties rPr = getRPr(false);
        if (rPr == null) {
            return null;
        }
        CTHyperlink hl = rPr.getHlinkClick();
        if (hl == null) {
            return null;
        }
        return new XSLFHyperlink(hl, getParagraph().getParentShape().getSheet());
    }

    private <T> T fetchCharacterProperty(CharPropFetcher<T> fetcher){
        final XSLFTextShape shape = getParagraph().getParentShape();
        return new CharacterPropertyFetcher<>(this, fetcher).fetchProperty(shape);
    }

    void copy(XSLFTextRun r){
        String srcFontFamily = r.getFontFamily();
        if(srcFontFamily != null && !srcFontFamily.equals(getFontFamily())){
            setFontFamily(srcFontFamily);
        }

        PaintStyle srcFontColor = r.getFontColor();
        if(srcFontColor != null && !srcFontColor.equals(getFontColor())){
            setFontColor(srcFontColor);
        }

        Double srcFontSize = r.getFontSize();
        if (srcFontSize == null) {
            if (getFontSize() != null) {
                setFontSize(null);
            }
        } else if(!srcFontSize.equals(getFontSize())) {
            setFontSize(srcFontSize);
        }

        boolean bold = r.isBold();
        if(bold != isBold()) {
            setBold(bold);
        }

        boolean italic = r.isItalic();
        if(italic != isItalic()) {
            setItalic(italic);
        }

        UnderlineType underline = r.getUnderline();
        if (!underline.equals(getUnderline())) {
            setUnderline(underline);
        }

        StrikeType strike = r.getStrikeThrough();
        if (!strike.equals(getStrikeThrough())) {
            setStrikeThrough(strike);
        }

        CapitalsType capitals = r.getCapitals();
        if (!capitals.equals(getCapitals())) {
            setCapitals(capitals);
        }

        XSLFHyperlink hyperSrc = r.getHyperlink();
        if (hyperSrc != null) {
            XSLFHyperlink hyperDst = getHyperlink();
            hyperDst.copy(hyperSrc);
        }
    }


    @Override
    public FieldType getFieldType() {
        if (isField()) {
            CTTextField tf = (CTTextField) getXmlObject();
            if ("slidenum".equals(tf.getType())) {
                return FieldType.SLIDE_NUMBER;
            }
        }
        return null;
    }


    private final class XSLFFontInfo implements FontInfo {
        private final FontGroup fontGroup;

        private XSLFFontInfo(FontGroup fontGroup) {
            this.fontGroup = (fontGroup != null) ? fontGroup : FontGroup.getFontGroupFirst(getRawText());
        }

        void copyFrom(FontInfo fontInfo) {
            CTTextFont tf = getXmlObject(true);
            if (tf == null) {
                return;
            }
            setTypeface(fontInfo.getTypeface());
            setCharset(fontInfo.getCharset());
            FontPitch pitch = fontInfo.getPitch();
            FontFamily family = fontInfo.getFamily();
            if (pitch == null && family == null) {
                if (tf.isSetPitchFamily()) {
                    tf.unsetPitchFamily();
                }
            } else {
                setPitch(pitch);
                setFamily(family);
            }
        }

        @Override
        public String getTypeface() {
            CTTextFont tf = getXmlObject(false);
            return (tf != null && tf.isSetTypeface()) ? tf.getTypeface() : null;
        }

        @Override
        public void setTypeface(String typeface) {
            if (typeface != null) {
                final CTTextFont tf = getXmlObject(true);
                if (tf != null) {
                    tf.setTypeface(typeface);
                }
                return;
            }

            CTTextCharacterProperties props = getRPr(false);
            if (props == null) {
                return;
            }
            FontGroup fg = FontGroup.getFontGroupFirst(getRawText());
            switch (fg) {
            default:
            case LATIN:
                if (props.isSetLatin()) {
                    props.unsetLatin();
                }
                break;
            case EAST_ASIAN:
                if (props.isSetEa()) {
                    props.unsetEa();
                }
                break;
            case COMPLEX_SCRIPT:
                if (props.isSetCs()) {
                    props.unsetCs();
                }
                break;
            case SYMBOL:
                if (props.isSetSym()) {
                    props.unsetSym();
                }
                break;
            }
        }

        @Override
        public FontCharset getCharset() {
            CTTextFont tf = getXmlObject(false);
            return (tf != null && tf.isSetCharset()) ? FontCharset.valueOf(tf.getCharset()&0xFF) : null;
        }

        @Override
        public void setCharset(FontCharset charset) {
            CTTextFont tf = getXmlObject(true);
            if (tf == null) {
                return;
            }
            if (charset != null) {
                tf.setCharset((byte)charset.getNativeId());
            } else {
                if (tf.isSetCharset()) {
                    tf.unsetCharset();
                }
            }
        }

        @Override
        public FontFamily getFamily() {
            CTTextFont tf = getXmlObject(false);
            return (tf != null && tf.isSetPitchFamily()) ? FontFamily.valueOfPitchFamily(tf.getPitchFamily()) : null;
        }

        @Override
        public void setFamily(FontFamily family) {
            CTTextFont tf = getXmlObject(true);
            if (tf == null || (family == null && !tf.isSetPitchFamily())) {
                return;
            }
            FontPitch pitch = (tf.isSetPitchFamily())
                ? FontPitch.valueOfPitchFamily(tf.getPitchFamily())
                : FontPitch.VARIABLE;
            byte pitchFamily = FontPitch.getNativeId(pitch, family != null ? family : FontFamily.FF_SWISS);
            tf.setPitchFamily(pitchFamily);
        }

        @Override
        public FontPitch getPitch() {
            CTTextFont tf = getXmlObject(false);
            return (tf != null && tf.isSetPitchFamily()) ? FontPitch.valueOfPitchFamily(tf.getPitchFamily()) : null;
        }

        @Override
        public void setPitch(FontPitch pitch) {
            CTTextFont tf = getXmlObject(true);
            if (tf == null || (pitch == null && !tf.isSetPitchFamily())) {
                return;
            }
            FontFamily family = (tf.isSetPitchFamily())
                ? FontFamily.valueOfPitchFamily(tf.getPitchFamily())
                : FontFamily.FF_SWISS;
            byte pitchFamily = FontPitch.getNativeId(pitch != null ? pitch : FontPitch.VARIABLE, family);
            tf.setPitchFamily(pitchFamily);
        }

        private CTTextFont getXmlObject(boolean create) {
            if (create) {
                return getCTTextFont(getRPr(true), true);
            }

            return fetchCharacterProperty((props, val) -> {
                CTTextFont font = getCTTextFont(props, false);
                if (font != null) {
                    val.accept(font);
                }
            });
        }

        private CTTextFont getCTTextFont(CTTextCharacterProperties props, boolean create) {
            if (props == null) {
                return null;
            }

            CTTextFont font;
            switch (fontGroup) {
            default:
            case LATIN:
                font = props.getLatin();
                if (font == null && create) {
                    font = props.addNewLatin();
                }
                break;
            case EAST_ASIAN:
                font = props.getEa();
                if (font == null && create) {
                    font = props.addNewEa();
                }
                break;
            case COMPLEX_SCRIPT:
                font = props.getCs();
                if (font == null && create) {
                    font = props.addNewCs();
                }
                break;
            case SYMBOL:
                font = props.getSym();
                if (font == null && create) {
                    font = props.addNewSym();
                }
                break;
            }

            if (font == null) {
                return null;
            }

            String typeface = font.isSetTypeface() ? font.getTypeface() : "";
            if (typeface.startsWith("+mj-") || typeface.startsWith("+mn-")) {
                //  "+mj-lt".equals(typeface) || "+mn-lt".equals(typeface)
                final XSLFTheme theme = getParagraph().getParentShape().getSheet().getTheme();
                CTFontScheme fontTheme = theme.getXmlObject().getThemeElements().getFontScheme();
                CTFontCollection coll = typeface.startsWith("+mj-")
                    ? fontTheme.getMajorFont() : fontTheme.getMinorFont();
                // TODO: handle LCID codes
                // see https://blogs.msdn.microsoft.com/officeinteroperability/2013/04/22/office-open-xml-themes-schemes-and-fonts/
                String fgStr = typeface.substring(4);
                if ("ea".equals(fgStr)) {
                    font = coll.getEa();
                } else if ("cs".equals(fgStr)) {
                    font = coll.getCs();
                } else {
                    font = coll.getLatin();
                }
                // SYMBOL is missing

                if (font == null || !font.isSetTypeface() || "".equals(font.getTypeface())) {
                    // don't fallback to latin but bubble up in the style hierarchy (slide -> layout -> master -> theme)
                    return null;
//                    font = coll.getLatin();
                }
            }

            return font;
        }
    }

    @Override
    public XSLFTextParagraph getParagraph() {
        return (XSLFTextParagraph)_parent;
    }
}

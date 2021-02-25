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
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Optional;
import java.util.function.Function;
import java.util.function.Predicate;

import org.apache.poi.ooxml.util.POIXMLUnits;
import org.apache.poi.util.Internal;
import org.apache.poi.util.Removal;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.text.TextAlignment;
import org.apache.poi.xddf.usermodel.text.TextContainer;
import org.apache.poi.xddf.usermodel.text.XDDFSpacing;
import org.apache.poi.xddf.usermodel.text.XDDFSpacingPercent;
import org.apache.poi.xddf.usermodel.text.XDDFSpacingPoints;
import org.apache.poi.xddf.usermodel.text.XDDFTextBody;
import org.apache.poi.xddf.usermodel.text.XDDFTextParagraph;
import org.apache.poi.xddf.usermodel.text.XDDFTextRun;
import org.apache.poi.xssf.model.ParagraphPropertyFetcher;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.*;

/**
 * Represents a paragraph of text within the containing text body.
 * The paragraph is the highest level text separation mechanism.
 */
public class XSSFTextParagraph extends XDDFTextParagraph implements TextContainer, Iterable<XSSFTextRun>{

    XSSFTextParagraph(CTTextParagraph p, XDDFTextBody parent){
        super(p, parent);

        _runs.clear();
        for(XmlObject r : _p.selectPath("*")){
            if (r instanceof CTTextLineBreak) {
                _runs.add(new XSSFTextRun((CTTextLineBreak)r, this));
            } else if (r instanceof CTTextField) {
                _runs.add(new XSSFTextRun((CTTextField)r, this));
            } else if (r instanceof CTRegularTextRun) {
                _runs.add(new XSSFTextRun((CTRegularTextRun)r, this));
            }
        }
    }

    @Internal
    public CTTextParagraph getXmlObject(){
        return _p;
    }

    public List<XSSFTextRun> getTextRuns(){
        final ArrayList<XSSFTextRun> runs = new ArrayList<>(_runs.size());
        for (XDDFTextRun r : _runs) {
            runs.add((XSSFTextRun) r);
        }
        return runs;
    }

    public Iterator<XSSFTextRun> iterator(){
        return getTextRuns().iterator();
    }

    /**
     * Add a new run of text
     *
     * @return a new run of text
     */
    public XSSFTextRun addRegularRun(String text) {
        XDDFTextRun xddfTextRun = super.appendRegularRun(text);
        _runs.remove(xddfTextRun);
        XSSFTextRun run = new XSSFTextRun((CTRegularTextRun)xddfTextRun.getXmlObject(), this);
        _runs.add(run);
        return run;
    }

    /**
     * Insert a line break
     *
     * @return text run representing this line break ('\n')
     */
    public XSSFTextRun addLineBreak(){
        XDDFTextRun xddfTextRun = super.appendLineBreak();
        _runs.remove(xddfTextRun);
        XSSFTextRun run = new XSSFTextRun((CTTextLineBreak)xddfTextRun.getXmlObject(), this);
        _runs.add(run);
        return run;
    }

    /**
     * Add a single tab stop to be used on a line of text when there are one or more tab characters
     * present within the text.
     *
     * @param value the position of the tab stop relative to the left margin
     */
    public void addTabStop(double value){
        CTTextParagraphProperties pr = _p.isSetPPr() ? _p.getPPr() : _p.addNewPPr();
        CTTextTabStopList tabStops = pr.isSetTabLst() ? pr.getTabLst() : pr.addNewTabLst();
        tabStops.addNewTab().setPos(Units.toEMU(value));
    }

    /**
     * Specifies the particular level text properties that this paragraph will follow.
     * The value for this attribute formats the text according to the corresponding level
     * paragraph properties defined in the list of styles associated with the body of text
     * that this paragraph belongs to (therefore in the parent shape).
     * <p>
     * Note that the closest properties object to the text is used, therefore if there is
     * a conflict between the text paragraph properties and the list style properties for
     * this level then the text paragraph properties will take precedence.
     * </p>
     *
     * @param level the level (0 ... 4)
     */
    public void setLevel(int level){
        CTTextParagraphProperties pr = _p.isSetPPr() ? _p.getPPr() : _p.addNewPPr();

        pr.setLvl(level);
    }

    /**
     * Returns the level of text properties that this paragraph will follow.
     *
     * @return the text level of this paragraph (0-based). Default is 0.
     */
    public int getLevel(){
        CTTextParagraphProperties pr = _p.getPPr();
        if(pr == null) return 0;

        return pr.getLvl();
    }


    /**
     * Set this paragraph as an automatic numbered bullet point
     *
     * @param scheme type of auto-numbering
     * @param startAt the number that will start number for a given sequence of automatically
     *        numbered bullets (1-based).
     */
    public void setBullet(ListAutoNumber scheme, int startAt) {
        if(startAt < 1) throw new IllegalArgumentException("Start Number must be greater or equal that 1") ;
        CTTextParagraphProperties pr = _p.isSetPPr() ? _p.getPPr() : _p.addNewPPr();
        CTTextAutonumberBullet lst = pr.isSetBuAutoNum() ? pr.getBuAutoNum() : pr.addNewBuAutoNum();
        lst.setType(STTextAutonumberScheme.Enum.forInt(scheme.ordinal() + 1));
        lst.setStartAt(startAt);

        if(!pr.isSetBuFont()) pr.addNewBuFont().setTypeface("Arial");
        if(pr.isSetBuNone()) pr.unsetBuNone();
        // remove these elements if present as it results in invalid content when opening in Excel.
        if(pr.isSetBuBlip()) pr.unsetBuBlip();
        if(pr.isSetBuChar()) pr.unsetBuChar();
    }

    /**
     * Set this paragraph as an automatic numbered bullet point
     *
     * @param scheme type of auto-numbering
     */
    public void setBullet(ListAutoNumber scheme) {
        CTTextParagraphProperties pr = _p.isSetPPr() ? _p.getPPr() : _p.addNewPPr();
        CTTextAutonumberBullet lst = pr.isSetBuAutoNum() ? pr.getBuAutoNum() : pr.addNewBuAutoNum();
        lst.setType(STTextAutonumberScheme.Enum.forInt(scheme.ordinal() + 1));

        if(!pr.isSetBuFont()) pr.addNewBuFont().setTypeface("Arial");
        if(pr.isSetBuNone()) pr.unsetBuNone();
        // remove these elements if present as it results in invalid content when opening in Excel.
        if(pr.isSetBuBlip()) pr.unsetBuBlip();
        if(pr.isSetBuChar()) pr.unsetBuChar();
    }


    @Override
    public String toString(){
        return "[" + getClass() + "]" + getText();
    }

    @Override
    public <R> Optional<R> findDefinedParagraphProperty(Predicate<CTTextParagraphProperties> isSet, Function<CTTextParagraphProperties, R> getter) {
        // only required for the compiler not to fail
        return super.findDefinedParagraphProperty(isSet,getter);
    }

    @Override
    public <R> Optional<R> findDefinedRunProperty(Predicate<CTTextCharacterProperties> isSet, Function<CTTextCharacterProperties, R> getter) {
        // only required for the compiler not to fail
        return super.findDefinedRunProperty(isSet, getter);
    }
}

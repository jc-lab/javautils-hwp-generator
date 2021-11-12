package kr.jclab.javautils.hwptemplate;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.control.Control;
import kr.dogfoot.hwplib.object.bodytext.control.ControlTable;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.CtrlHeaderGso;
import kr.dogfoot.hwplib.object.bodytext.control.ctrlheader.gso.*;
import kr.dogfoot.hwplib.object.bodytext.control.gso.ControlPicture;
import kr.dogfoot.hwplib.object.bodytext.control.gso.ControlRectangle;
import kr.dogfoot.hwplib.object.bodytext.control.gso.GsoControlType;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.ShapeComponentNormal;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.lineinfo.*;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.shadowinfo.ShadowInfo;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponent.shadowinfo.ShadowType;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponenteach.ShapeComponentPicture;
import kr.dogfoot.hwplib.object.bodytext.control.gso.shapecomponenteach.polygon.PositionXY;
import kr.dogfoot.hwplib.object.bodytext.control.table.Cell;
import kr.dogfoot.hwplib.object.bodytext.control.table.Row;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.HWPChar;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.HWPCharType;
import kr.dogfoot.hwplib.object.docinfo.BinData;
import kr.dogfoot.hwplib.object.docinfo.bindata.BinDataCompress;
import kr.dogfoot.hwplib.object.docinfo.bindata.BinDataState;
import kr.dogfoot.hwplib.object.docinfo.bindata.BinDataType;
import kr.dogfoot.hwplib.object.docinfo.borderfill.fillinfo.PictureEffect;
import kr.dogfoot.hwplib.object.docinfo.borderfill.fillinfo.PictureInfo;
import kr.dogfoot.hwplib.object.etc.Color4Byte;
import kr.jclab.javautils.hwptemplate.generator.GetItemValueHandler;
import kr.jclab.javautils.hwptemplate.generator.ImageItemValue;
import kr.jclab.javautils.hwptemplate.generator.ItemValue;
import kr.jclab.javautils.hwptemplate.generator.TextItemValue;
import kr.jclab.javautils.hwptemplate.generator.intl.TraversalState;

import java.awt.*;
import java.io.UnsupportedEncodingException;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class TemplatedHwpGenerator {
    private static final String BEGIN_ELEMENT_CODE = "##BEGIN##";
    private static final String END_ELEMENT_CODE = "##END##";
    private static final Pattern BEGIN_ELEMENT_PATTERN = Pattern.compile(BEGIN_ELEMENT_CODE);
    private static final Pattern END_ELEMENT_PATTERN = Pattern.compile(END_ELEMENT_CODE);
    private static final Pattern DYNAMIC_VALUE_PATTERN = Pattern.compile("\\$\\{([^}]+)}");

    @lombok.Getter
    private final HWPFile hwpFile;
    @lombok.Getter
    private int templateSectionBeginIndex = 0;
    @lombok.Getter
    private int templateElementCount = 0;

    private Section templateSection = null;

    public TemplatedHwpGenerator(HWPFile hwpFile) {
        this.hwpFile = hwpFile;
    }

    public boolean prepare() {
        TraversalState<Void> traversalState = TraversalState.<Void>builder()
                .build();

        int index = 0;
        for (Iterator<Section> sectionIterator = this.hwpFile.getBodyText().getSectionList().iterator(); sectionIterator.hasNext(); index++) {
            Section section = sectionIterator.next();
            traversalSection(traversalState, 0, section);
            if (traversalState.getTemplateCount() > 0) {
                this.templateSection = section;
                sectionIterator.remove();
                this.templateSectionBeginIndex = index;
            }
        }

        this.templateElementCount = traversalState.getTemplateCount();

        return this.templateElementCount > 0;
    }

    public <ItemType> HWPFile generate(
            List<ItemType> items,
            GetItemValueHandler<ItemType> handler
    ) {
        HWPFile newFile = this.hwpFile.clone(true);
        TraversalState<ItemType> state = TraversalState.<ItemType>builder()
                .generateMode(true)
                .targetHwpFile(newFile)
                .items(new ArrayDeque<>(items))
                .getItemValueHandler(handler)
                .build();
        int index = this.templateSectionBeginIndex;
        while (state.getRemaining() > 0) {
            Section newSection = this.templateSection.clone();
            traversalSection(state, 0, newSection);
            newFile.getBodyText().getSectionList().add(index++, newSection);
        }
        return newFile;
    }

    private void replaceToText(TraversalState<?> state, Paragraph paragraph, String text) throws UnsupportedEncodingException {
        List<HWPChar> charList = paragraph.getText().getCharList();
        List<HWPChar> lastChars = new ArrayList<>();
        if (!charList.isEmpty()) {
            for (HWPChar c : charList) {
                if (c.getType() == HWPCharType.Normal) {
                    lastChars.clear();
                } else {
                    lastChars.add(c);
                }
            }
        }
        paragraph.getText().getCharList().clear();
        paragraph.getText().addString(text);
        paragraph.getText().getCharList().addAll(lastChars);
//        paragraph.createCharShape();
//        paragraph.getCharShape().addParaCharShape(0, 5);
    }

    private void replaceToImage(TraversalState<?> state, Paragraph paragraph, Rectangle shapePosition, ImageItemValue imageItemValue) {
        int streamId = state.getTargetHwpFile().getBinData().getEmbeddedBinaryDataList().size() + 1;
        String streamName = getStreamName(streamId, imageItemValue.getExt());
        state.getTargetHwpFile().getBinData().addNewEmbeddedBinaryData(
                streamName,
                imageItemValue.getData(),
                BinDataCompress.Compress
        );
        int binDataID = addBinDataInDocInfo(state.getTargetHwpFile(), streamId);
        paragraph.createText();
        paragraph.getText().addExtendCharForGSO();
        ControlPicture controlPicture = (ControlPicture) paragraph.addNewGsoControl(GsoControlType.Picture);
        CtrlHeaderGso header = controlPicture.getHeader();
        GsoHeaderProperty prop = header.getProperty();
        prop.setLikeWord(false);
        prop.setApplyLineSpace(false);
        prop.setVertRelTo(VertRelTo.Para);
        prop.setVertRelativeArrange(RelativeArrange.TopOrLeft);
        prop.setHorzRelTo(HorzRelTo.Para);
        prop.setHorzRelativeArrange(RelativeArrange.TopOrLeft);
        prop.setVertRelToParaLimit(true);
        prop.setAllowOverlap(true);
        prop.setWidthCriterion(WidthCriterion.Absolute);
        prop.setHeightCriterion(HeightCriterion.Absolute);
        prop.setProtectSize(false);
        prop.setTextFlowMethod(TextFlowMethod.FitWithText);
        prop.setTextHorzArrange(TextHorzArrange.BothSides);
        prop.setObjectNumberSort(ObjectNumberSort.Figure);

        header.setyOffset(shapePosition.y);
        header.setxOffset(shapePosition.x);
        header.setWidth(shapePosition.width);
        header.setHeight(shapePosition.height);
        header.setzOrder(0);
        header.setOutterMarginLeft(0);
        header.setOutterMarginRight(0);
        header.setOutterMarginTop(0);
        header.setOutterMarginBottom(0);
        header.setInstanceId(0x5bb840e1);
        header.setPreventPageDivide(false);
        // header.setExplanation(null);
        ShapeComponentPicture scp = controlPicture.getShapeComponentPicture();
        Color4Byte borderColor = scp.getBorderColor();
        borderColor.setValue(0);

        LineInfoProperty borderProperty = scp.getBorderProperty();
        borderProperty.setLineEndShape(LineEndShape.Flat);
        borderProperty.setStartArrowShape(LineArrowShape.None);
        borderProperty.setStartArrowSize(LineArrowSize.MiddleMiddle);
        borderProperty.setEndArrowShape(LineArrowShape.None);
        borderProperty.setEndArrowSize(LineArrowSize.MiddleMiddle);
        borderProperty.setLineType(LineType.None);
        borderProperty.setFillStartArrow(true);
        borderProperty.setFillEndArrow(true);

        PositionXY leftTop = scp.getLeftTop();
        PositionXY leftBottom = scp.getLeftBottom();
        PositionXY rightBottom = scp.getRightBottom();
        PositionXY rightTop = scp.getRightTop();

        leftTop.setX(0);
        leftTop.setY(0);
        rightTop.setX(shapePosition.width);
        rightTop.setY(0);
        leftBottom.setX(0);
        leftBottom.setY(shapePosition.height);
        rightBottom.setX(shapePosition.width);
        rightBottom.setY(shapePosition.height);

        PictureInfo pictureInfo = scp.getPictureInfo();
        pictureInfo.setBinItemID(binDataID);
        pictureInfo.setEffect(PictureEffect.RealPicture);

        ShapeComponentNormal sc = (ShapeComponentNormal) controlPicture.getShapeComponent();
        sc.setOffsetX(0);
        sc.setOffsetY(0);
        sc.setGroupingCount(0);
        sc.setLocalFileVersion(1);
        sc.setWidthAtCreate(shapePosition.width);
        sc.setHeightAtCreate(shapePosition.height);
        sc.setWidthAtCurrent(shapePosition.width);
        sc.setHeightAtCurrent(shapePosition.height);
        sc.setRotateAngle(0);
        sc.setRotateXCenter(shapePosition.width / 2);
        sc.setRotateYCenter(shapePosition.height / 2);

        sc.createLineInfo();
        LineInfo li = sc.getLineInfo();
        li.getProperty().setLineEndShape(LineEndShape.Flat);
        li.getProperty().setStartArrowShape(LineArrowShape.None);
        li.getProperty().setStartArrowSize(LineArrowSize.MiddleMiddle);
        li.getProperty().setEndArrowShape(LineArrowShape.None);
        li.getProperty().setEndArrowSize(LineArrowSize.MiddleMiddle);

        li.getProperty().setFillStartArrow(true);
        li.getProperty().setFillEndArrow(true);
        li.getProperty().setLineType(LineType.None);
        li.setOutlineStyle(OutlineStyle.Normal);
        li.setThickness(0);
        li.getColor().setValue(0);

        sc.createShadowInfo();
        ShadowInfo si = sc.getShadowInfo();
        si.setType(ShadowType.None);
        si.getColor().setValue(0xc4c4c4);
        si.setOffsetX(0);
        si.setOffsetY(0);
        si.setTransparent((short) 0);

        sc.setMatrixsNormal();

        scp.setImageHeight(shapePosition.height);
        scp.setImageWidth(shapePosition.width);
        scp.setInstanceId(0x5bb810e1);

        scp.setTopAfterCutting(0);
        scp.setLeftAfterCutting(0);
        scp.setBottomAfterCutting(0);
        scp.setRightAfterCutting(0);
    }

    private <ItemType> boolean traversalPara(TraversalState<ItemType> state, int depth, Paragraph paragraph, Rectangle parentSize) {
        if (paragraph.getControlList() != null) {
            for (Control control : paragraph.getControlList()) {
                if (control instanceof ControlTable) {
                    ControlTable controlImpl = (ControlTable) control;
                    for (Row row : controlImpl.getRowList()) {
                        for (Cell cell : row.getCellList()) {
                            for (int i = 0; i < cell.getParagraphList().getParagraphCount(); i++) {
                                Paragraph currentParagraph = cell.getParagraphList().getParagraph(i);
                                if (!traversalPara(state, depth + 1, currentParagraph, new Rectangle(
                                        0,
                                        0,
                                        (int) cell.getListHeader().getWidth(),
                                        (int) cell.getListHeader().getHeight()
                                ))) {
                                    return false;
                                }
                            }
                        }
                    }
                } else if (control instanceof ControlRectangle) {
                    ControlRectangle controlImpl = (ControlRectangle) control;
                    if (controlImpl.getTextBox() != null) {
                        for (Paragraph subPara : controlImpl.getTextBox().getParagraphList()) {
                            if (!traversalPara(state, depth + 1, subPara, new Rectangle(
                                    0,
                                    0,
                                    (int) controlImpl.getHeader().getWidth(),
                                    (int) controlImpl.getHeader().getHeight()
                            ))) {
                                return false;
                            }
                        }
                    }
                }
            }
        } else {
            try {
                String normalString = paragraph.getNormalString();

                if (BEGIN_ELEMENT_PATTERN.matcher(normalString).find()) {
                    if (state.isGenerateMode()) {
                        replaceToText(state, paragraph, normalString.replaceAll(BEGIN_ELEMENT_CODE, ""));
                    }
                    if (!state.begin()) return false;
                } else if (END_ELEMENT_PATTERN.matcher(normalString).find()) {
                    if (state.isGenerateMode()) {
                        replaceToText(state, paragraph, normalString.replaceAll(END_ELEMENT_CODE, ""));
                    }
                    if (!state.end()) return false;
                } else {
                    Matcher m = DYNAMIC_VALUE_PATTERN.matcher(normalString);
                    if (state.isGenerateMode() && m.find()) {
                        String key = m.group(1);
                        ItemValue itemValue = state.getItemValue(key);
                        if (itemValue != null) {
                            if (itemValue instanceof TextItemValue) {
                                TextItemValue textItemValue = (TextItemValue) itemValue;
                                replaceToText(state, paragraph, m.replaceAll(textItemValue.getText()));
                            } else if (itemValue instanceof ImageItemValue) {
                                ImageItemValue imageItemValue = (ImageItemValue) itemValue;
                                replaceToImage(state, paragraph, parentSize, imageItemValue);
                            }
                        }
                    }
                }
            } catch (UnsupportedEncodingException e) {
                e.printStackTrace();
            }
        }
        return true;
    }

    private <ItemType> void traversalSection(TraversalState<ItemType> state, int depth, Section section) {
        for (Paragraph paragraph : section.getParagraphs()) {
            if (!traversalPara(state, depth + 1, paragraph, null)) {
                break;
            }
        }
    }

    private static String getStreamName(int streamIndex, String imageFileExt) {
        return "Bin" + String.format("%04X", streamIndex) + "." + imageFileExt;
    }

    private static int addBinDataInDocInfo(HWPFile hwpFile, int streamIndex) {
        BinData bd = new BinData();
        bd.getProperty().setType(BinDataType.Embedding);
        bd.getProperty().setCompress(BinDataCompress.Compress);
        bd.getProperty().setState(BinDataState.NotAccess);
        bd.setBinDataID(streamIndex);
        bd.setExtensionForEmbedding("png");
        hwpFile.getDocInfo().getBinDataList().add(bd);
        return hwpFile.getDocInfo().getBinDataList().size();
    }
}

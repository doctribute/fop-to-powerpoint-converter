package org.doctribute.tools;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Point;
import java.awt.Rectangle;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URI;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.xml.parsers.DocumentBuilderFactory;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlipFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRelativeRect;
import org.openxmlformats.schemas.presentationml.x2006.main.CTPicture;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

public class FOPToPowerPointConverter {

    private static final Pattern TRANSLATE_TRANSFORM_PATTERN = Pattern.compile("translate\\((\\d+),(\\d+)\\)");

    public static void convert(InputStream inputStream, OutputStream outputStream) throws Exception {

        XMLSlideShow pptx = new XMLSlideShow();
        Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(inputStream);
        NodeList nodelist = document.getDocumentElement().getElementsByTagName("page");
        for (int i = 0; i < nodelist.getLength(); i++) {
            Element page = (Element) nodelist.item(i);
            if (i == 0) {
                pptx.setPageSize(getDimension(page));
            }
            processPage(pptx, page);
        }

        pptx.write(outputStream);
    }

    private static void processPage(XMLSlideShow pptx, Element page) {

        XSLFSlide slide = pptx.createSlide();

        processChildren(
                pptx, slide, new PageContext(), new ShapeContext(),
                (Element) page.getElementsByTagName("content").item(0));
    }

    private static void processChildren(
            XMLSlideShow pptx, XSLFSlide slide, PageContext pageContext, ShapeContext shapeContext, Element element) {

        NodeList nodelist = element.getChildNodes();
        if (nodelist.getLength() > 0) {
            for (int i = 0; i < nodelist.getLength(); i++) {
                if (nodelist.item(i) instanceof Element child) {
                    switch (child.getNodeName()) {
                        case "viewport":
                            processViewport(pptx, slide, pageContext, shapeContext, child);
                            break;
                        case "g":
                            processGroup(pptx, slide, pageContext, shapeContext, child);
                            break;
                        case "rect":
                            processRect(pptx, slide, pageContext, shapeContext, child);
                            break;
                        case "clip-rect":
                            shapeContext = shapeContext.clone(getRectangle(child));
                            break;
                        case "image":
                            processImage(pptx, slide, pageContext, shapeContext, child);
                            break;
                        case "text":
                            processText(pptx, slide, pageContext, shapeContext, child);
                            break;
                        case "font":
                            updateFontInfo(pageContext, child);
                            break;
                    }
                }
            }
        }
    }

    private static void processViewport(
            XMLSlideShow pptx, XSLFSlide slide, PageContext pageContext, ShapeContext shapeContext, Element element) {
        //TODO
        processChildren(pptx, slide, pageContext, shapeContext, element);
    }

    private static void processGroup(
            XMLSlideShow pptx, XSLFSlide slide, PageContext pageContext, ShapeContext shapeContext, Element element) {

        if (element.hasAttribute("transform")) {
            String transform = element.getAttribute("transform");
            Matcher matcher = TRANSLATE_TRANSFORM_PATTERN.matcher(transform);
            if (matcher.find()) {
                ShapeContext newShapeContext = shapeContext.clone(
                        Double.parseDouble(matcher.group(1)), Double.parseDouble(matcher.group(2)));
                processChildren(pptx, slide, pageContext, newShapeContext, element);

            } else {
                System.out.println("Unsupported transform: " + transform);
            }

        } else {
            processChildren(pptx, slide, pageContext, shapeContext, element);
        }
    }

    private static void processRect(
            XMLSlideShow pptx, XSLFSlide slide, PageContext pageContext, ShapeContext shapeContext, Element element) {

        if (element.hasAttribute("fill")) {
            Rectangle rectangle = getRectangle(element);
            rectangle.translate(toPoints(shapeContext.getTranslateX()), toPoints(shapeContext.getTranslateY()));

            XSLFAutoShape shape = slide.createAutoShape();
            shape.setFillColor(Color.decode(element.getAttribute("fill")));
            shape.setAnchor(rectangle);
        }
    }

    private static void processImage(
            XMLSlideShow pptx, XSLFSlide slide, PageContext pageContext, ShapeContext shapeContext, Element element) {

        Path imagePath = Paths.get(URI.create(element.getAttribute("xlink:href")));

        try {
            PictureData.PictureType pictureType = getPictureType(imagePath);

            Rectangle rectangle = getRectangle(element);
            rectangle.translate(toPoints(shapeContext.getTranslateX()), toPoints(shapeContext.getTranslateY()));

            try (InputStream inputStream = Files.newInputStream(imagePath)) {
                XSLFPictureData pictureData = pptx.addPicture(inputStream, pictureType);
                XSLFPictureShape shape = slide.createPicture(pictureData);
                if (!shapeContext.getClipRectangle().isEmpty()) {
                    CTPicture ct = (CTPicture) shape.getXmlObject();
                    CTBlipFillProperties bfp = ct.getBlipFill();
                    bfp.unsetStretch();
                    bfp.addNewStretch();

                    setSrcRect(bfp, rectangle, shapeContext.getClipRectangle());

                    rectangle = rectangle.intersection(shapeContext.getClipRectangle());
                }
                shape.setAnchor(rectangle);
            }

        } catch (IOException e) {
            System.out.println("Unsupported image type: " + imagePath);
        }
    }

    private static void processText(
            XMLSlideShow pptx, XSLFSlide slide, PageContext pageContext, ShapeContext shapeContext, Element element) {

        Rectangle textAnchor = new Rectangle(
                getOrigin(element),
                getDimension((Element) element.getParentNode()));
        textAnchor.translate(toPoints(shapeContext.getTranslateX()), toPoints(shapeContext.getTranslateY()));

        XSLFTextBox shape = slide.createTextBox();
        shape.setAnchor(textAnchor);
        XSLFTextParagraph p = shape.addNewTextParagraph();
        XSLFTextRun textRun = p.addNewTextRun();

        FontInfo fontInfo = pageContext.getFontInfo();
        textRun.setFontFamily(fontInfo.getFontFamily());
        textRun.setFontSize(fontInfo.getFontSize());
        if (fontInfo.getFontStyle().equals("italic")) {
            textRun.setItalic(true);
        }
        if (fontInfo.getFontWeight().equals("400")) {
            textRun.setBold(true);
        }
        textRun.setFontColor(Color.decode(fontInfo.getFillColor()));

        textRun.setText(element.getTextContent().replaceAll("\\s{2,}", " ").strip());
    }

    private static void updateFontInfo(PageContext pageContext, Element element) {

        FontInfo fontInfo = pageContext.getFontInfo();
        if (fontInfo == null) {
            fontInfo = new FontInfo();
        }

        String fontFamily = fontInfo.getFontFamily();
        double fontSize = fontInfo.getFontSize();
        String fontStyle = fontInfo.getFontStyle();
        String fontWeight = fontInfo.getFontWeight();
        String fillColor = fontInfo.getFillColor();

        if (element.hasAttribute("family")) {
            fontFamily = element.getAttribute("family");
        }
        if (element.hasAttribute("size")) {
            fontSize = toPoints(Integer.parseInt(element.getAttribute("size")));
        }
        if (element.hasAttribute("style")) {
            fontStyle = element.getAttribute("style");
        }
        if (element.hasAttribute("weight")) {
            fontWeight = element.getAttribute("weight");
        }
        if (element.hasAttribute("color")) {
            fillColor = element.getAttribute("color");
        }

        pageContext.setFontInfo(new FontInfo(fontFamily, fontSize, fontStyle, fontWeight, fillColor));
    }

    private static Rectangle getRectangle(Element element) {
        return new Rectangle(getOrigin(element), getDimension(element));
    }

    private static Point getOrigin(Element element) {
        int x = Integer.parseInt(element.getAttribute("x"));
        int y = Integer.parseInt(element.getAttribute("y"));
        return new Point(toPoints(x), toPoints(y));
    }

    private static Dimension getDimension(Element element) {
        int width = Integer.parseInt(element.getAttribute("width"));
        int height = Integer.parseInt(element.getAttribute("height"));
        return new Dimension(toPoints(width), toPoints(height));
    }

    private static void setSrcRect(
            CTBlipFillProperties bfp, Rectangle shapeRectangle, Rectangle clipRectangle) {

        CTRelativeRect srcRect = bfp.addNewSrcRect();

        int deltaWidth = shapeRectangle.width - clipRectangle.width;
        int deltaHeight = shapeRectangle.height - clipRectangle.height;

        if (deltaWidth > 0) {
            int left = clipRectangle.x - shapeRectangle.x;
            int right = deltaWidth - left;

            srcRect.setL(100000 * left / shapeRectangle.width);
            srcRect.setR(100000 * right / shapeRectangle.width);
        }

        if (deltaHeight > 0) {
            int top = clipRectangle.y - shapeRectangle.y;
            int bottom = deltaHeight - top;

            srcRect.setT(100000 * top / shapeRectangle.height);
            srcRect.setB(100000 * bottom / shapeRectangle.height);
        }
    }

    private static PictureData.PictureType getPictureType(Path imagePath) throws IOException {
        String contentType = Files.probeContentType(imagePath);
        return switch (contentType) {
            case "image/jpeg" -> PictureData.PictureType.JPEG;
            case "image/png" -> PictureData.PictureType.PNG;
            case "image/gif" -> PictureData.PictureType.GIF;
            case "image/svg+xml" -> PictureData.PictureType.SVG;
            default -> throw new IOException();
        };
    }

    private static int toPoints(double value) {
        return (int) value / 1000;
    }

}

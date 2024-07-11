package org.doctribute.tools;

import java.awt.Rectangle;

public class ShapeContext {

    private final double translateX;
    private final double translateY;
    private final Rectangle clipRectangle;

    public ShapeContext() {
        this.translateX = 0;
        this.translateY = 0;
        this.clipRectangle = new Rectangle();
    }

    public ShapeContext(double translateX, double translateY, Rectangle clipRectangle) {
        this.translateX = translateX;
        this.translateY = translateY;
        this.clipRectangle = clipRectangle;
    }

    public ShapeContext clone(double translateX, double translateY) {
        return new ShapeContext(translateX, translateY, this.clipRectangle);
    }

    public ShapeContext clone(Rectangle clipRectangle) {
        return new ShapeContext(this.translateX, this.translateY, clipRectangle);
    }

    public double getTranslateX() {
        return translateX;
    }

    public double getTranslateY() {
        return translateY;
    }

    public Rectangle getClipRectangle() {
        return clipRectangle;
    }

}

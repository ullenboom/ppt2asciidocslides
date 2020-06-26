package com.tutego.adoc.dom.text;

public class Title{
    int layer; // 0 = DocumentHeader
    Paragraph text;
    public Title(Paragraph text, int layer){
        this.layer = layer;
        this.text = text;
    }
    public Title(String text, int layer){
        this.text = new Paragraph(text);
    }

    public String toAdoc(){
        return null;
    }
}

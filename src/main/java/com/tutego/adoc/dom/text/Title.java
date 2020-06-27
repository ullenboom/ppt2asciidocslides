package com.tutego.adoc.dom.text;

import com.tutego.adoc.dom.Content;

public class Title implements Content {
    int layer; // 0 = DocumentHeader
    Paragraph text;
    public Title(Paragraph text, int layer){
        this.layer = layer;
        this.text = text;
    }
    public Title(String text, int layer){
        this.text = new Paragraph(text);
    }

    @Override
    public String toAdoc() {
        return null;
    }
}

package com.tutego.adoc.dom.text;

import java.util.List;

public class Title extends Paragraph {
    int layer; // 0 = DocumentHeader
    public Title(List<Text> texts, int layer){
        super(texts);
        this.layer = layer;
    }

    public String toAdoc(){
        return null;
    }
}

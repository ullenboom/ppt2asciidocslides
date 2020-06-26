package com.tutego.adoc.dom.text;

import java.util.ArrayList;
import java.util.List;

public class Paragraph {

    List<Text> texts;

    public Paragraph(String text, Style... style){
        texts = new ArrayList<>();
        texts.add(new Text(text, style));
    }

    public Paragraph(List<Text> texts){
        this.texts = texts;
    }

    public Paragraph(String text){
        texts = new ArrayList<>();
        texts.add(new Text(text));
    }

    public Paragraph(){
        texts = new ArrayList<>();
    }

    public String toAdoc(){
        return null;
    }

    public void addText(Text text){
        texts.add(text);
    }

    public List<Text> getTexts() {
        return texts;
    }
}

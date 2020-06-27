package com.tutego.adoc.dom.text;

import com.tutego.adoc.dom.Content;

public class Title implements Content {
    int level;
    Paragraph text;

    public Title( int level, Paragraph text){
        this.level = level;
        this.text = text;
        validate();
    }

    public Title( int level, String text){
        this.level = level;
        this.text = new Paragraph(text);
        validate();
    }


    public Title(int level, Text... text){
        this.level = level;
        this.text = new Paragraph(text);
        validate();
    }

    @Override
    public String toAdoc() {
        String result = " " + text.toAdoc();
        for (int i = 0; i < level; i++){
            result = "=" + result;
        }
        return result;
    }

    @Override
    public void validate() {
        if (level > 5 || level < 1){
            throw new IllegalArgumentException("Layer has to be between 0 and 5.");
        }
    }
}

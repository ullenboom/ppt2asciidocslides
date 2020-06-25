package com.tutego.adoc.dom.text;

public class Link{
    String url;
    Paragraph text;
    public Link(Paragraph text, String url){
        this.url = url;
        this.text = text;
    }

    public String toAdoc(){
        return null;
    }
}

package com.tutego.adoc.dom.text;

public class Text{
    Style[] styles;
    String text;
    public Text(String text, Style... styles){
        this.text = text;
        this.styles = styles;
    }
    public String toAdoc(){
        return null;
    }
}

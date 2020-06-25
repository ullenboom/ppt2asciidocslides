package com.tutego.adoc.dom.text;

public class Text {
    enum style {NORMAL, BOLD, ITALIC, MARKED, UNDERLINED, CROSSED}
    style[] styles;
    String text;
    public Text(String text, style... styles){
        this.text = text;
        this.styles = styles;
    }
    public String toAdoc(){
        return null;
    }
}

package com.tutego.adoc.dom.container;

import com.tutego.adoc.dom.text.Paragraph;
import com.tutego.adoc.dom.text.Text;

import java.util.List;

public class ListItem extends Paragraph {
    int indent;
    public ListItem(List<Text> texts, int indent){
        super(texts);
        this.indent = indent;
    }
    public String toAdoc(){
        return null;
    }

}

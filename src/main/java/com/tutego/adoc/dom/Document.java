package com.tutego.adoc.dom;

import java.util.ArrayList;
import java.util.List;

public class Document implements Content{

    String Author;

    List<Object> contents = new ArrayList<>();

    public void addObject(Content content){
        contents.add(content);
    }

    @Override
    public String toAdoc(){
        return null;
    }

    @Override
    public void validate() {
    }
}

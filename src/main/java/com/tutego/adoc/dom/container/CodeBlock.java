package com.tutego.adoc.dom.container;

import java.util.Arrays;
import java.util.List;

public class CodeBlock extends Block{
    List<String> languages;
    public CodeBlock(String text, String... languages) {
        super(text);
        this.languages = Arrays.asList(languages);
    }
    public String toAdoc(){
        return null;
    }

}
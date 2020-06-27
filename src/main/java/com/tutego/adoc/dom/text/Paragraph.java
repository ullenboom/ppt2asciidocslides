package com.tutego.adoc.dom.text;

import com.tutego.adoc.dom.Content;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Paragraph implements Content {

  private final List<Text> texts;

  public Paragraph( String text, Text.Style... style ) {
    texts = new ArrayList<>();
    texts.add( new Text( text, style ) );
  }

  public Paragraph( Text... texts ) {
    this.texts = Arrays.asList( texts );
  }

  public void addText( Text text ) {
    texts.add( text );
  }

  public List<Text> getTexts() {
    return texts;
  }

  @Override
  public String toAdoc() {
    return null;
  }
}

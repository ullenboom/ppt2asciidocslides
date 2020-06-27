package com.tutego.adoc.dom.text;

import com.tutego.adoc.dom.Content;
import org.apache.commons.collections4.ListUtils;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.IntStream;

public class Paragraph implements Content {

  private List<Text> texts;

  public Paragraph( String text, Text.Style... style ) {
    texts = new ArrayList<>();
    texts.add( new Text( text, style ) );
  }

  public Paragraph( Text... texts ) {
    this.texts = Arrays.asList( texts );
  }

  public void appendText( Text text ) {
    texts.add( text );
  }

  public List<Text> getTexts() {
    return texts;
  }

  public void merge(Paragraph paragraph){
    texts = ListUtils.union(texts, paragraph.getTexts());
  }

  @Override
  public String toAdoc() {
    System.out.println(texts.size());
    StringBuilder result = new StringBuilder();
    IntStream.range(0, texts.size()).forEach(index ->
            result.append(index != texts.size()-1 ? texts.get(index).toAdoc() + " " : texts.get(index).toAdoc()));

    return result.toString();
  }

  @Override
  public void validate() {

  }
}

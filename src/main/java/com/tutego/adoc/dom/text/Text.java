package com.tutego.adoc.dom.text;

import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.Objects;

public class Text {

  public enum Style {
    ITALIC( "_" ),TOP("^"), DOWN("~"), MONOSPACE( "`" ), BOLD( "*" );

    String key;

    Style( String key ) {
      this.key = key;
    }

    String getKey() {
      return key;
    }
  }

  private final Style[] styles;
  private final String text;

  public Text( String text, Style... styles ) {
    this.text = Objects.requireNonNull( text );
    this.styles = Objects.requireNonNull( styles );
    validate();
  }

  public String toAdoc() {
    Collections.sort(Arrays.asList(styles), Comparator.naturalOrder());
    String rs = text;
    for ( Style style : styles ) {
      rs = applyStyle( rs, style );
    }
    return rs;
  }

  private String applyStyle( String text, Style style ) {
    String key = style.getKey();
    return String.format( "%s%s%s", key, text, key );
  }

  private void validate(){
    if((Arrays.asList(styles).contains(Style.TOP) || Arrays.asList(styles).contains(Style.DOWN)) && text.contains(" ")){
      throw new IllegalArgumentException("TOP and DOWN can only be applied to one continuous String.");
    }
    if (text.isEmpty()){
      throw new IllegalArgumentException("The text can not be an empty String");
    }
  }
}

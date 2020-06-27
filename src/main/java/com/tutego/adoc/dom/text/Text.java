package com.tutego.adoc.dom.text;

import java.util.*;

public class Text {

  public enum Style {
    ITALIC( "_" ), MONOSPACE( "`" ), BOLD( "*" );

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
  }

  public String toAdoc() {
    //sortStyles();
    /*Collections.sort(Arrays.asList(styles), new Comparator<Style>() {
      @Override
      public int compare(Style style1, Style style2) {
        return style1.compareTo(style2);
      }
    });*/

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
}

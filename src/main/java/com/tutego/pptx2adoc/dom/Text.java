package com.tutego.pptx2adoc.dom;

import com.tutego.pptx2adoc.writer.DocumentWriter;

import java.util.EnumSet;
import java.util.Set;

/// A styled text run — the atomic inline unit of the DOM.
/// Carries the raw text and a set of semantic [Style] markers.
public final class Text implements TextSpan {

  /// Semantic text styles supported by the DOM. The actual markup characters
  /// are defined by the respective [DocumentWriter] implementation.
  public enum Style {
    ITALIC, TOP, DOWN, MONOSPACE, BOLD
  }

  private final EnumSet<Style> styles;
  private final String text;

  public Text( String text ) {
    if ( text == null || text.isEmpty() )
      throw new IllegalArgumentException( "Text cannot be null or empty" );

    this.text = text;
    this.styles = EnumSet.noneOf( Style.class );
  }

  public Text style( Style style ) {
    if ( (style == Style.TOP || style == Style.DOWN) && text.contains( " " ) ) {
      throw new IllegalArgumentException(
          "TOP and DOWN can only be applied to one continuous String." );
    }
    styles.add( style );
    return this;
  }

  public String text() {
    return text;
  }

  public Set<Style> styles() {
    return styles;
  }
}

package com.tutego.pptx2adoc.dom;

import java.util.Objects;

/// A plain paragraph block containing inline [TextSpan] elements.
///
public final class Paragraph extends Texts implements Block {

  public Paragraph add( TextSpan text ) {
    super.add( Objects.requireNonNull( text ) );
    return this;
  }
}

package com.tutego.pptx2adoc.dom;

import java.util.ArrayList;
import java.util.List;

/// An ordered collection of [TextSpan] elements forming a paragraph-like
/// text body. Serves as base class for [Header], [Paragraph], and [ListItem].
///
public class Texts {

  private final List<TextSpan> textSpans = new ArrayList<>();

  public Texts add( TextSpan text ) {
    textSpans.add( text );
    return this;
  }

  public List<TextSpan> texts() {
    return textSpans;
  }
}

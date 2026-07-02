package com.tutego.pptx2adoc.dom;

import java.util.Objects;

/// A block quote with text, optional author, and optional reference.
public final class Quote implements Block {
  private String text = "";
  private String author = "";
  private String reference = "";

  public Quote text( String text ) {
    this.text = Objects.requireNonNull( text );
    return this;
  }

  public String text() {
    return text;
  }

  public Quote author( String author ) {
    this.author = Objects.requireNonNull( author );
    return this;
  }

  public String author() {
    return author;
  }

  public Quote reference( String reference ) {
    this.reference = Objects.requireNonNull( reference );
    return this;
  }

  public String reference() {
    return reference;
  }
}

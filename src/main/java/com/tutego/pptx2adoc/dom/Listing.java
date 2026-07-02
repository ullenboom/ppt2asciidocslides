package com.tutego.pptx2adoc.dom;

import java.util.Objects;

/// A source code listing block with a language identifier and code text.
///
public final class Listing implements Block {

  private String language = "";
  private String code = "";

  public Listing language( String language ) {
    this.language = Objects.requireNonNull( language );
    return this;
  }

  public String language() {
    return language;
  }

  public Listing code( String code ) {
    this.code = Objects.requireNonNull( code );
    return this;
  }

  public String code() {
    return code;
  }
}

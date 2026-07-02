package com.tutego.pptx2adoc.dom;

import java.util.Objects;

/// A hyperlink span with a URL and optional link text.
///
public final class Link implements TextSpan {

  private Text text;
  private String url = "";

  public Link url( String url ) {
    this.url = Objects.requireNonNull( url );
    return this;
  }

  public String url() {
    return url;
  }

  public Link text( Text text ) {
    this.text = text;
    return this;
  }

  public Text text() {
    return text;
  }
}

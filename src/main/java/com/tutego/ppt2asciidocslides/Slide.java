package com.tutego.ppt2asciidocslides;

/**
 * Represents a simple Asciidoc slide with a title and a body.
 * 
 * @author Christian Ullenboom
 */
class Slide {

  String title = "";
  String body  = "";

  @Override
  public String toString() {
    return title + (body.isEmpty() ? "" : "\n" + body);
  }
}

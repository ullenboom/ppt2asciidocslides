package com.tutego.pptx2adoc.dom;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/// Root node of the DOM tree, holding an ordered list of [Block] elements
/// and optional document metadata such as the author.
public class Document {

  private String author = "";
  private final List<Block> blocks = new ArrayList<>();

  public Document add( Block block ) {
    blocks.add( block );
    return this;
  }

  public List<Block> blocks() {
    return Collections.unmodifiableList( blocks );
  }

  public Document author( String author ) {
    this.author = author;
    return this;
  }

  public String author() {
    return author;
  }
}

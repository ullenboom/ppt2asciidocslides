package com.tutego.pptx2adoc.dom;

import java.util.Objects;

/// An image block with a path, optional title, and optional width/height.
public final class Image implements Block {

  private String title = "";
  private String path = "";
  private int width = -1;
  private int height = -1;

  public Image path( String path ) {
    Objects.requireNonNull( path );

    if ( path.isEmpty() )
      throw new IllegalArgumentException( "Path can not be an empty String." );

    this.path = path;
    return this;
  }

  public String path() {
    return path;
  }

  public Image size( int width, int height ) {
    this.width = width;
    this.height = height;
    return this;
  }

  public Image size( int width ) {
    this.width = width;
    return this;
  }

  public int width() {
    return width;
  }

  public int height() {
    return height;
  }

  public Image title( String title ) {
    this.title = title;
    return this;
  }

  public String title() {
    return title;
  }
}

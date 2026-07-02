package com.tutego.pptx2adoc.dom;

/// A heading block with a level (1–5) and inline text content.
/// In Marp output, each [Header] starts a new slide separated by `---`.
///
public final class Header extends Texts implements Block {
  private int level;

  public Header level( int level ) {
    if ( level < 1 || level > 5 )
      throw new IllegalArgumentException( "Level has to be between 1 and 5" );

    this.level = level;
    return this;
  }

  public int level() {
    return level;
  }

  public Header text( Text text ) {
    super.add( text );
    return this;
  }
}

package com.tutego.pptx2adoc.dom;

/// A list item block with an indentation level (1–5), a [Type] (ordered or unordered),
/// and inline text content.
///
public final class ListItem extends Texts implements Block {

  private int level;
  private Type type;

  /// Whether a list item is ordered (numbered) or unordered (bulleted).
  ///
  public enum Type {
    UNORDERED, ORDERED
  }

  public ListItem level( int level ) {
    if ( level < 1 || level > 5 )
      throw new IllegalArgumentException( "Level has to be between 1 and 5." );

    this.level = level;
    return this;
  }

  public ListItem type( Type type ) {
    this.type = type;
    return this;
  }

  public int level() {
    return level;
  }

  public Type type() {
    return type;
  }

  @Override
  public ListItem add( TextSpan text ) {
    super.add( text );
    return this;
  }
}

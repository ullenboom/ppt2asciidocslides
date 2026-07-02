package com.tutego.pptx2adoc.writer;

import com.tutego.pptx2adoc.dom.*;

import java.io.PrintWriter;
import java.util.List;
import java.util.Map;

import static com.tutego.pptx2adoc.dom.ListItem.Type.ORDERED;
import static com.tutego.pptx2adoc.dom.ListItem.Type.UNORDERED;
import static com.tutego.pptx2adoc.dom.Text.Style.*;

/// Converts a [Document] DOM tree to AsciiDoc output via pattern matching.
/// All AsciiDoc-specific syntax lives here; the DOM is renderer-neutral.
public class AsciidocWriter implements DocumentWriter {

  private static final Map<Text.Style, String> STYLE_MARKERS = Map.of(
      ITALIC, "_",
      TOP, "^",
      DOWN, "~",
      MONOSPACE, "`",
      BOLD, "*"
  );

  private static final Map<ListItem.Type, String> LIST_INDICATORS = Map.of(
      UNORDERED, "*",
      ORDERED, "."
  );

  public void write( Document document, PrintWriter out ) {
    for ( Block block : document.blocks() ) {
      renderBlock( block, out );
      out.println();
    }
  }

  private void renderBlock( Block block, PrintWriter out ) {
    switch ( block ) {
      case Header h    -> renderHeader( h, out );
      case Paragraph p -> renderTexts( p, out );
      case ListItem li -> renderListItem( li, out );
      case Image img   -> renderImage( img, out );
      case Listing l   -> renderListing( l, out );
      case Quote q     -> renderQuote( q, out );
    }
  }

  private void renderHeader( Header h, PrintWriter out ) {
    out.println();
    out.println();
    out.print( "=".repeat( h.level() ) );
    out.print( ' ' );
    renderTexts( h, out );
  }

  private void renderListItem( ListItem li, PrintWriter out ) {
    out.print( LIST_INDICATORS.get( li.type() ).repeat( li.level() ) );
    out.print( ' ' );
    renderTexts( li, out );
  }

  private void renderImage( Image img, PrintWriter out ) {
    out.printf( "image::%s[%s%s%s]",
                img.path(),
                img.title().isEmpty() ? "" : img.title(),
                img.width() == -1 ? "" : "," + img.width(),
                img.height() == -1 ? "" : "," + img.height() );
  }

  private void renderListing( Listing l, PrintWriter out ) {
    out.print( "[source," );
    out.print( l.language() );
    out.println( "]" );
    out.println( "----" );
    out.print( l.code() );
    out.println( "----" );
    out.println();
  }

  private void renderQuote( Quote q, PrintWriter out ) {
    out.printf( "[quote,%s,%s]%n", q.author(), q.reference() );
    out.print( q.text() );
  }

  private void renderTexts( Texts texts, PrintWriter out ) {
    List<TextSpan> spans = texts.texts();
    for ( TextSpan span : spans )
      renderTextSpan( span, out );
    out.println();
  }

  private void renderTextSpan( TextSpan span, PrintWriter out ) {
    switch ( span ) {
      case Text t  -> renderText( t, out );
      case Link l  -> renderLink( l, out );
    }
  }

  private void renderText( Text t, PrintWriter out ) {
    String text = t.text();
    String leading = text.substring( 0, text.length() - text.stripLeading().length() );
    String trailing = text.substring( text.stripTrailing().length() );
    String core = text.strip();

    StringBuilder inner = new StringBuilder( core );
    for ( Text.Style style : t.styles() )
      inner.insert( 0, STYLE_MARKERS.get( style ) ).append( STYLE_MARKERS.get( style ) );

    out.print( leading );
    out.print( inner );
    out.print( trailing );
  }

  private void renderLink( Link l, PrintWriter out ) {
    if ( l.url().isEmpty() )
      throw new IllegalStateException( "Link url must be set before rendering." );
    out.print( l.url() );
    out.print( '[' );
    if ( l.text() != null )
      renderText( l.text(), out );
    out.print( ']' );
  }
}

package com.tutego.pptx2adoc.writer;

import com.tutego.pptx2adoc.dom.*;

import java.io.PrintWriter;
import java.util.List;
import java.util.Map;

import static com.tutego.pptx2adoc.dom.Text.Style.*;

/// Converts a [Document] DOM tree to Marp Markdown output (CommonMark + Marpit extensions).
/// Slides are separated by `---`, detected by [Header] blocks.
public class MarpWriter implements DocumentWriter {

  private static final Map<Text.Style, String> INLINE_MARKERS = Map.of(
      BOLD, "**",
      ITALIC, "*",
      MONOSPACE, "`"
  );

  @Override
  public void write( Document document, PrintWriter out ) {
    writeFrontMatter( out );

    boolean firstSlide = true;
    for ( Block block : document.blocks() ) {
      if ( block instanceof Header ) {
        if ( !firstSlide ) {
          out.println();
          out.println( "---" );
          out.println();
        }
        firstSlide = false;
      }
      renderBlock( block, out );
      out.println();
    }
  }

  private void writeFrontMatter( PrintWriter out ) {
    out.println( "---" );
    out.println( "theme: default" );
    out.println( "paginate: true" );
    out.println( "---" );
    out.println();
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
    out.print( "#".repeat( h.level() ) );
    out.print( ' ' );
    renderTexts( h, out );
  }

  private void renderListItem( ListItem li, PrintWriter out ) {
    String indent = "  ".repeat( li.level() - 1 );
    String marker = li.type() == ListItem.Type.ORDERED ? "1. " : "- ";
    out.print( indent );
    out.print( marker );
    renderTexts( li, out );
  }

  private void renderImage( Image img, PrintWriter out ) {
    StringBuilder alt = new StringBuilder();
    if ( !img.title().isEmpty() )
      alt.append( img.title() );
    if ( img.width() != -1 )
      alt.append( alt.isEmpty() ? "" : " " ).append( "w:" ).append( img.width() );
    if ( img.height() != -1 )
      alt.append( alt.isEmpty() ? "" : " " ).append( "h:" ).append( img.height() );
    out.printf( "![%s](%s)", alt, img.path() );
    out.println();
  }

  private void renderListing( Listing l, PrintWriter out ) {
    out.print( "```" );
    out.print( l.language() );
    out.println();
    out.print( l.code() );
    out.println();
    out.println( "```" );
    out.println();
  }

  private void renderQuote( Quote q, PrintWriter out ) {
    out.printf( "> %s%n", q.text() );
    if ( !q.author().isEmpty() ) {
      out.print( "> — " );
      out.print( q.author() );
      if ( !q.reference().isEmpty() )
        out.printf( " (%s)", q.reference() );
      out.println();
    }
  }

  private void renderTexts( Texts texts, PrintWriter out ) {
    List<TextSpan> spans = texts.texts();
    for ( TextSpan span : spans )
      renderTextSpan( span, out );
    out.println();
  }

  private void renderTextSpan( TextSpan span, PrintWriter out ) {
    switch ( span ) {
      case Text t -> renderText( t, out );
      case Link l -> renderLink( l, out );
    }
  }

  private void renderText( Text t, PrintWriter out ) {
    String text = t.text();
    String leading = text.substring( 0, text.length() - text.stripLeading().length() );
    String trailing = text.substring( text.stripTrailing().length() );
    String core = text.strip();

    StringBuilder inner = new StringBuilder( core );
    for ( Text.Style style : t.styles() ) {
      String marker = INLINE_MARKERS.get( style );
      if ( marker != null )
        inner.insert( 0, marker ).append( marker );
      else if ( style == TOP )
        inner.insert( 0, "<sup>" ).append( "</sup>" );
      else if ( style == DOWN )
        inner.insert( 0, "<sub>" ).append( "</sub>" );
    }

    out.print( leading );
    out.print( inner );
    out.print( trailing );
  }

  private void renderLink( Link l, PrintWriter out ) {
    out.print( '[' );
    if ( l.text() != null )
      renderText( l.text(), out );
    out.print( "](" );
    out.print( l.url() );
    out.print( ')' );
  }
}

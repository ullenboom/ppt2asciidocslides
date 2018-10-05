package com.tutego.pptx2adoc;

import java.awt.Color;
import java.io.IOException;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Optional;
import java.util.Scanner;
import java.util.Stack;
import java.util.regex.Pattern;
import org.apache.poi.sl.usermodel.PaintStyle.SolidPaint;
import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 * Converter from Microsoft PowerPoint to Asciidoc.
 * 
 * @author Christian Ullenboom
 */
class Converter {

  private final Configuration configuration;

  Converter( Configuration configuration ) {
    this.configuration = configuration;
  }

  /**
   * Reads a PPT slide from the configuration and writes an Asciidoc file.
   * 
   * @throws IOException
   */
  void writePptSlides() throws IOException {
    // Read all PPT slides
    var slides = extractAllPptSlides( configuration.input );
    // Write all slides in Asciidoc format
    try ( var writer = Files.newBufferedWriter( configuration.output ) ) {
      for ( var slide : slides )
        writer.write( slide.toString() );
    }
  }

  /**
   * Opens a given PowerPoint, extract all slides and converts to a domain model.
   * 
   * @param filename
   * @return
   * @throws IOException
   */
  private Slides extractAllPptSlides( Path filename ) throws IOException {
    var slides = new Slides();
    try ( var input = Files.newInputStream( filename );
          var slideShow = new XMLSlideShow( input ) ) {
      slideShow.getSlides().forEach( slide -> slides.addSlide( exctractPptSlide( slide ) ) );
    }
    return slides;
  }

  /**
   * Iterates through all elements of a slide and converts to Asciidoc. Recognizes basically text an images.
   * 
   * @param pptSlide
   * @return
   */
  private Slide exctractPptSlide( XSLFSlide pptSlide ) {
    var slide = new Slide();
    Optional<XSLFTextShape> headerTextShape = getTextShapeByType( pptSlide, Placeholder.TITLE );
    slide.title = headerTextShape.isPresent() ? ("\n== " + headerTextShape.get().getText() + "\n")
                                              : "\n== KEINE ÜBERSCHRIFT\n";

    var body = "";
    for ( var shape : pptSlide.getShapes() ) {

      // Skip title text shape, we got that already
      if ( shape.equals( headerTextShape.orElse( null ) ) )
        continue;

      if ( shape instanceof XSLFTextShape ) {
        var text = extractPptTextShape( (XSLFTextShape) shape );
        if ( !text.trim().isEmpty() )
          body += text;
      }
      else if ( shape instanceof XSLFPictureShape ) {
        if ( configuration.writeImages ) {
          var picShape = (XSLFPictureShape) shape;
          var picData  = picShape.getPictureData();

          var imageDir = configuration.output.getParent().resolve( configuration.imageFolder );

          try {
            if ( Files.notExists( imageDir ) )
              Files.createDirectories( imageDir );

            var checksum = Math.abs( new BigInteger( picData.getChecksum() ).hashCode() );
            var imageFile = Paths.get( checksum + "-" + picData.getFileName() );
            var imagePath = imageDir.resolve( imageFile );

            Files.copy( picData.getInputStream(), imagePath, StandardCopyOption.REPLACE_EXISTING );
            body += "image::" + configuration.imageFolder + "/" + imageFile + "[]\n";
          }
          catch ( Exception e ) {
            e.printStackTrace();
          }
        }
      }
    }

    slide.body = body;

    return slide;
  }

  /**
   * Converts a shape to Asciidoc.
   * @param tsh
   * @return
   */
  private static String extractPptTextShape( XSLFTextShape tsh ) {
    var para = "";
    for ( XSLFTextParagraph pptPara : tsh ) {
      // ignore pptPara.isBullet(), just use indent level
      switch ( pptPara.getIndentLevel() ) {
        case 0: break;
        case 1: para += "* "; break;
        case 2: para += " * "; break;
      }

      for ( var textRun : pptPara ) {
        // Code?
        // TODO: some blanks get swallowed, so for now place a space in front.
        SolidPaint solidPaint = (SolidPaint) textRun.getFontColor();
        var color = solidPaint == null ? Color.BLACK : solidPaint.getSolidColor().getColor();
        int gray = fromRGBtoGray( color.getRGB() );
        if ( textRun.getFontFamily().equalsIgnoreCase( "Consolas" ) || gray > 20 )
          para += "`" + textRun.getRawText() + "`";
        else
          para = convertFontAttributes( para, textRun );
      }

      // join two consecutive code segments
      para = para.replaceAll( "``", "" ).trim();
      para += "\n\n";
    } // end for

    // Remove blank link in between 2 code lines
    para = para.replaceAll( "`(\\r?\\n)*`", "`\n`" );

    // Remove leading space before ` (code)
    para = para.replaceAll( "^ +`", "`" );

    // Convert consecutive code sequence lines from `xxx` to ...\nxxx\n...\n
    StringBuilder buffer = new StringBuilder( 1024 );

    try ( var scanner = new Scanner( para ) ) {
      while ( scanner.hasNextLine() ) {
        var line = scanner.nextLine();
        if ( line.startsWith( "`" ) && line.endsWith( "`" ) && line.length() > 1 )
          buffer.append( "....\n" ).append( line.substring( 1, line.length() - 1 ) )
                  .append( "\n....\n" );
        else
          buffer.append( line.trim() ).append( "\n" );
      }
    }

    // join two consecutive code blocks
    String codeTag = Pattern.quote( "...." );
    String result = buffer.toString().replaceAll( codeTag + "\n" + codeTag + "\n", "" );

    // remove all tabs
    result = result.replaceAll( "\t", "" );

    // remove empty lines before end of code blocks, like \n\n....
    result = result.replaceAll( "(\\r?\\n)*" + codeTag, "\n...." );

    result = result.replaceAll( "," + "`" + "\\s", "," + " " + "`" );

    // // Remove consequitive new lines
    // result = result.replaceAll( "(\\r?\\n)+", "\n" );

    return result;
  }

  /**
   * Converts a  {@code texxt} to Asciidoc and appends the result to a {@code para}.
   * 
   * @param para
   * @param text
   * @return
   */
  private static String convertFontAttributes( String para, XSLFTextRun text ) {
    var stack = new Stack<String>();
    if ( text.isBold() ) {
      para += "**";
      stack.add( "**" );
    }
    if ( text.isItalic() ) {
      para += "__";
      stack.add( "__" );
    }
    if ( text.isSubscript() ) {
      para += "~";
      stack.add( "~" );
    }
    if ( text.isSuperscript() ) {
      para += "^";
      stack.add( "^" );
    }

    para += text.getRawText();

    while ( !stack.isEmpty() )
      para += stack.pop();

    return para;
  }

  /**
   * Converts a RGB into gray. 0 is black and 255 is white.
   * 
   * @param rgb
   * @return
   */
  private static int fromRGBtoGray( int rgb ) {
    return (((rgb >> 16) & 0xFF) + ((rgb >> 8) & 0xFF) + (rgb & 0xFF)) / 3;
  }

  /**
   * Searches in a slide for a given type and returns the first {@code XSLFTextShape}. 
   * 
   * @param slide
   * @param type
   * @return
   */
  private static Optional<XSLFTextShape> getTextShapeByType( XSLFSlide slide, Placeholder type ) {
    for ( var shape : slide.getShapes() ) {
      if ( shape instanceof XSLFTextShape ) {
        var textShape = (XSLFTextShape) shape;
        if ( textShape.getTextType() == type )
          return Optional.ofNullable( textShape );
      }
    }
    return Optional.empty();
  }
}

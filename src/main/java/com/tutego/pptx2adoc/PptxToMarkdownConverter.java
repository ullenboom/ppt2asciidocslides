package com.tutego.pptx2adoc;

import com.tutego.pptx2adoc.dom.*;
import com.tutego.pptx2adoc.writer.DocumentWriter;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.xslf.usermodel.*;

import java.io.IOException;
import java.io.PrintWriter;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Predicate;
import java.util.regex.Pattern;

import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;

/// Converts PowerPoint slides into a renderer-neutral DOM tree and delegates
/// output generation to a [DocumentWriter].
///
public class PptxToMarkdownConverter {

  private final Configuration configuration;
  private final DocumentWriter writer;
  private final Document document = new Document();

  public PptxToMarkdownConverter( Configuration configuration, DocumentWriter writer ) {
    this.configuration = configuration;
    this.writer = writer;
  }

  public void convertSlide( XSLFSlide slide ) {
    getHeaderTextShape( slide )
        .map( XSLFTextShape::getText )
        .or( () -> Optional.of( "No Header" ) )
        .map( Text::new )
        .map( title -> new Header().level( 3 ).text( title ) )
        .ifPresent( document::add );

    slide.getShapes().stream()
         .filter( shape -> !isHeaderShape( shape, slide ) )
         .forEach( this::processShape );
  }

  private Optional<XSLFTextShape> getHeaderTextShape( XSLFSlide slide ) {
    return slide.getShapes().stream()
                .filter( XSLFTextShape.class::isInstance )
                .map( XSLFTextShape.class::cast )
                .filter( shape -> shape.getTextType() == Placeholder.TITLE )
                .findFirst();
  }

  private boolean isHeaderShape( XSLFShape shape, XSLFSlide slide ) {
    return getHeaderTextShape( slide )
        .map( headerShape -> Objects.equals( headerShape, shape ) )
        .orElse( false );
  }

  private void processShape( XSLFShape shape ) {
    if ( shape instanceof XSLFTextShape textShape )
      processTextShape( textShape );
    else if ( shape instanceof XSLFPictureShape pictureShape )
      processPictureShape( pictureShape );
  }

  private void processPictureShape( XSLFPictureShape shape ) {
    if ( configuration.writeImages() )
      convertPictureToImage( shape ).ifPresent( document::add );
  }

  private void processTextShape( XSLFTextShape textShape ) {
    for ( XSLFTextParagraph paragraph : textShape ) {
      if ( paragraph.getIndentLevel() == 0 )
        document.add( createParagraph( paragraph ) );
      else
        document.add( createListItem( paragraph ) );
    }
  }

  private Paragraph createParagraph( XSLFTextParagraph textParagraph ) {
    var para = new Paragraph();
    extractTextsFromParagraph( textParagraph ).texts().forEach( para::add );
    return para;
  }

  private ListItem createListItem( XSLFTextParagraph textParagraph ) {
    int level = Math.clamp( textParagraph.getIndentLevel(), 1, 5 );
    var listItem = new ListItem()
        .level( level )
        .type( ListItem.Type.UNORDERED );
    extractTextsFromParagraph( textParagraph ).texts().forEach( listItem::add );
    return listItem;
  }

  private Optional<Image> convertPictureToImage( XSLFPictureShape pictureShape ) {
    Path targetDir = Objects.requireNonNull( configuration.target().getParent() );
    Path imagesDir = targetDir.resolve( configuration.imageDirectory() );
    XSLFPictureData pictureData = pictureShape.getPictureData();

    int checksum = new BigInteger( pictureData.getChecksum() ).hashCode() & 0x7FFF_FFFF;
    Path pictureFile = Paths.get( checksum + "-" + pictureData.getFileName() );
    Path picturePath = imagesDir.resolve( pictureFile );
    String relativePath = targetDir.relativize( picturePath ).toString().replace( '\\', '/' );

    try {
      if ( Files.notExists( imagesDir ) )
        Files.createDirectories( imagesDir );

      try ( var inputStream = pictureData.getInputStream() ) {
        Files.copy( inputStream, picturePath, REPLACE_EXISTING );
      }

      return Optional.of( new Image().path( relativePath ) );
    }
    catch ( IOException e ) {
      System.err.println( "Error saving image file. " + e );
      return Optional.empty();
    }
  }

  private Texts extractTextsFromParagraph( XSLFTextParagraph textParagraph ) {
    Texts texts = new Texts();

    for ( XSLFTextRun textRun : textParagraph )
      texts.add( createText( textRun ) );

    return texts;
  }

  private Text createText( XSLFTextRun textRun ) {
    var text = new Text( textRun.getRawText().replace( "\u00A0", " " ) );

    if ( isProbablyCodeFont( textRun ) )
      text.style( Text.Style.MONOSPACE );
    else
      applyTextStyles( text, textRun );

    return text;
  }

  private void applyTextStyles( Text text, XSLFTextRun textRun ) {
    if ( textRun.isBold() )
      text.style( Text.Style.BOLD );
    if ( textRun.isItalic() )
      text.style( Text.Style.ITALIC );
    if ( textRun.isSubscript() )
      text.style( Text.Style.DOWN );
    if ( textRun.isSuperscript() )
      text.style( Text.Style.TOP );
  }

  private static final Predicate<String> IS_MONOSPACE_FONT =
      Pattern.compile( "consolas|courier", Pattern.CASE_INSENSITIVE ).asPredicate();

  private static boolean isProbablyCodeFont( XSLFTextRun textRun ) {
    return Optional.ofNullable( textRun.getFontFamily() )
                   .filter( IS_MONOSPACE_FONT )
                   .isPresent()
        || isLighterColor( textRun.getFontColor() );
  }

  private static boolean isLighterColor( PaintStyle paintStyle ) {
    return paintStyle instanceof PaintStyle.SolidPaint solidPaint &&
        fromRGBtoGray( solidPaint.getSolidColor().getColor().getRGB() ) > 20;
  }

  private static int fromRGBtoGray( int rgb ) {
    return (((rgb >> 16) & 0xFF) + ((rgb >> 8) & 0xFF) + (rgb & 0xFF)) / 3;
  }

  public void writeAdoc( PrintWriter printWriter ) {
    writer.write( document, printWriter );
  }
}

package com.tutego.pptx2adoc;

import com.tutego.pptx2adoc.writer.AsciidocWriter;
import com.tutego.pptx2adoc.writer.DocumentWriter;
import com.tutego.pptx2adoc.writer.MarpWriter;
import joptsimple.OptionException;
import joptsimple.OptionParser;
import org.apache.poi.xslf.usermodel.XMLSlideShow;

import java.io.IOException;
import java.io.PrintWriter;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/// Main application class to start from the command line.
///
/// @author Christian Ullenboom
public class Main {

  public static void main( String... args ) {

    // Options: -i noimages -h help -v Version Operands: Input Output Imagefolder

    OptionParser parser = new OptionParser();
    parser.accepts( "noimages", "do not write images for slides" );
    parser.accepts( "target", "name target file to create" ).withRequiredArg();
    parser.accepts( "imagefolder", "name of folder containing images" ).withRequiredArg()
          .defaultsTo( "images" );
    parser.accepts( "format", "output format: adoc or marp" ).withRequiredArg()
          .defaultsTo( "adoc" );
    parser.accepts( "help", "show help" ).forHelp();

    var config = new Configuration();
    boolean printHelp = false;

    try {
      var options = parser.parse( args );
      var arguments = options.nonOptionArguments();

      if ( options.has( "help" ) ) {
        printHelp = true;
      }
      else if ( arguments.isEmpty() ) {
        System.err.println( "Input file is required" );
        printHelp = true;
      }
      else {
        config.format( options.valueOf( "format" ).toString().toLowerCase() );
        String extension = config.format().equals( "marp" ) ? ".md" : ".adoc";

        config.source( Paths.get( arguments.getFirst().toString() ).toAbsolutePath() );

        if ( config.source().toString().toLowerCase().endsWith( ".ppt" ) ) {
          printHelp = true;
          System.err.println(
              "This tool can only process PPTX files. Please convert your PPT to PPTX first." );
        }
        else {
          if ( options.has( "target" ) )
            config.target( Paths.get( options.valueOf( "target" ).toString() ).toAbsolutePath() );
          else {
            Path source = config.source();
            String fileName = source.getFileName().toString().replace( ".pptx", extension );
            config.target( source.resolveSibling( fileName ).toAbsolutePath() );
          }

          if ( options.has( "noimages" ) )
            config.writeImages( false );

          if ( options.has( "imagefolder" ) )
            config.imageDirectory( options.valueOf( "imagefolder" ).toString() );
        }
      }
    }
    catch ( OptionException e ) {
      System.err.printf( "Unrecognized option: %s%n%n", e.getLocalizedMessage() );
      printHelp = true;
    }
    finally {
      if ( printHelp )
        try {
          parser.printHelpOn( System.err );
        }
        catch ( IOException e ) {
          // Intentionally left empty.
        }
    }

    if ( printHelp )
      return;

    DocumentWriter writer = config.format().equals( "marp" ) ?
                            new MarpWriter() :
                            new AsciidocWriter();

    try ( var fileWriter = Files.newBufferedWriter( config.target() );
          var printWriter = new PrintWriter( fileWriter );
          var is = Files.newInputStream( config.source() );
          var xmlSlideShow = new XMLSlideShow( is ) ) {
      var converter = new PptxToMarkdownConverter( config, writer );
      xmlSlideShow.getSlides().forEach( converter::convertSlide );
      converter.writeAdoc( printWriter );
    }
    catch ( IOException e ) {
      System.err.println( "Error processing file: " + e.getLocalizedMessage() );
    }
    catch ( Exception e ) {
      System.err.println( "Error: " + e.getLocalizedMessage() );
    }
  }
}

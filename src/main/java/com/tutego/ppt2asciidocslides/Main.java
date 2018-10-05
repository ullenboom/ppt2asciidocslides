package com.tutego.ppt2asciidocslides;

import java.io.IOException;
import java.nio.file.Paths;
import joptsimple.OptionException;
import joptsimple.OptionParser;

/**
 * Main application class to start from the command line.
 * 
 * @author Christian Ullenboom
 */
public class Application {

  public static void main( String[] args ) {

    /*
     * Options: -i noimages -h help -v Version Operands: Input Output Imagefolder
     */

    OptionParser parser = new OptionParser();
    parser.accepts( "noimages", "do not write images for slides" );
    parser.accepts( "output", "name output file to create" ).withRequiredArg();
    parser.accepts( "imagefolder", "name of folder containing images" ).withRequiredArg().defaultsTo( "Images" );
    parser.accepts( "help", "show help" ).forHelp();

    var configuration = new Configuration();

    boolean printHelp = false;

    try {
      var options = parser.parse( args );
      var arguments = options.nonOptionArguments();

      if ( arguments.isEmpty() ) {
        System.err.println( "Input file is required" );
        printHelp = true;
        return;
      }

      configuration.input = Paths.get( arguments.get( 0 ).toString() ).toAbsolutePath();

      if ( options.has( "output" ) )
        configuration.output = Paths.get( options.valueOf( "output" ).toString() ).toAbsolutePath();
      else
        configuration.output = Paths.get( arguments.get( 0 ).toString().replace( ".pptx", ".adoc" ) ).toAbsolutePath();

      if ( options.has( "noimages" ) )
        configuration.writeImages = false;

      if ( options.has( "help" ) ) {
        printHelp = true;
        return;
      }
    }
    catch ( OptionException e ) {
      System.err.printf( "Unrecognized option: %s%n%n", e.getLocalizedMessage() );
      printHelp = true;
    }
    finally {
      if ( printHelp )
        try { parser.printHelpOn( System.err ); } catch ( IOException e ) { }
    }

    try {
      var converter = new Converter( configuration );
      converter.writePptSlides();
    }
    catch ( IOException e ) {
      System.err.println( "Error processing file: " + e.getMessage() );
    }
  }
}

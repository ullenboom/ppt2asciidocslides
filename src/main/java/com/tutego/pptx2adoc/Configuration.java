package com.tutego.pptx2adoc;

import java.nio.file.Path;
import java.nio.file.Paths;

/// Configuration for the PPTX-to-Markdown converter, holding source path,
/// target path, image settings, and image directory.
public class Configuration {

  private Path source;
  private Path target;
  private boolean writeImages = true;
  private String imageDirectory = "Images";
  private String format = "adoc";

  public Configuration format( String format ) {
    this.format = format;
    return this;
  }

  public String format() {
    return format;
  }

  public Configuration target( Path output ) {
    this.target = output;
    return this;
  }

  public Path target() {
    if ( target == null ) {
      target = Paths.get( source.toAbsolutePath().toString().replace( ".pptx", ".adoc" ) );
    }
    return target;
  }

  public Configuration source( Path source ) {
    this.source = source;
    return this;
  }

  public Path source() {
    return source;
  }

  public Configuration writeImages( boolean writeImages ) {
    this.writeImages = writeImages;
    return this;
  }

  public boolean writeImages() {
    return writeImages;
  }

  public Configuration imageDirectory( String imageDirectory ) {
    this.imageDirectory = imageDirectory;
    return this;
  }

  public String imageDirectory() {
    return imageDirectory;
  }
}

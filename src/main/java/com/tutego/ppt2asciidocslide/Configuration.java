package com.tutego.ppt2asciidocslide;

import java.nio.file.Path;

/**
 * Holds configuration for the options on the command line.
 * 
 * @author Christian Ullenboom
 */
class Configuration {

  Path  input;
  Path  output;
  boolean writeImages = true;
  String imageFolder = "Images";

  @Override
  public String toString() {
    return "Configuration [input=" + input + ", output=" + output + ", writeImages=" + writeImages
           + ", imageFolder=" + imageFolder + "]";
  }
}

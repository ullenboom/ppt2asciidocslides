package com.tutego.pptx2adoc;

import java.nio.file.Path;

/**
 * Holds configuration for the options of the command line.
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

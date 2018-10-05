package com.tutego.pptx2adoc;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * A collection of slides.
 * 
 * @author Christian Ullenboom
 */
class Slides implements Iterable<Slide> {

  private final List<Slide> slides = new ArrayList<>( 48 );

  void addSlide( Slide slide ) {
    slides.add( slide );
  }

  @Override
  public Iterator<Slide> iterator() {
    return slides.iterator();
  }
}

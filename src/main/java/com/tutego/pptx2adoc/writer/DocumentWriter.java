package com.tutego.pptx2adoc.writer;

import com.tutego.pptx2adoc.dom.Document;

import java.io.PrintWriter;

/// Strategy interface for rendering a [Document] DOM tree to a specific markup language.
public interface DocumentWriter {

  /// Renders the given document to the provided writer.
  ///
  /// @param document the document to render
  /// @param out      the target writer
  void write( Document document, PrintWriter out );
}

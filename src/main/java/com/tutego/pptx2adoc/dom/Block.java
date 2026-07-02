package com.tutego.pptx2adoc.dom;

/// Sealed interface for all top-level block elements that can appear in a [Document].
///
/// Permitted types: [Header], [Paragraph], [ListItem], [Image], [Listing], [Quote].
///
public sealed interface Block permits Header, Paragraph, ListItem, Image, Listing, Quote {
}

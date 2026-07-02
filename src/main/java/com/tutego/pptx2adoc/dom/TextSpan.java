package com.tutego.pptx2adoc.dom;

/// Sealed interface for inline text spans within a [Texts] container.
///
/// Permitted types: [Text], [Link].
///
public sealed interface TextSpan permits Text, Link {
}

package com.tutego.adoc.dom.container;

import com.tutego.adoc.dom.text.Paragraph;
import com.tutego.adoc.dom.text.Text;

import java.util.List;

public class ListItem extends Paragraph {
  int indent;

  public ListItem( int indent, Text... texts ) {
    super( texts );
    this.indent = indent;
  }

  public String toAdoc() {
    return null;
  }

}

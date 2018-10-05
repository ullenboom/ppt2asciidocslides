# PPT to Asciidoc converter

Asciidoc is a great tool for publishing technical information. Some time ago I changed my presentation slides from PowerPoint to Asciidoc. Because I had so many PPT slides I wrote this converter to get a first draft.

The converter is a command line tool written in Java. It uses Maven as a build tool and `org.apache.poi:poi-ooxml´ to access the raw PPTX files.

From the command line use:

    $ java com.tutego.ppt2asciidocslide.Main Slide.pptx

Use can use the option

| Option      | Description  | Default  |
| ----------- | ------------ | -------- |
| output      | Specifies the target filename | Input filename and suffix _.adoc_ |
| writeImages | If imagages get extracted and written or not | true |
| imageFolder | Folder for the generated images |  _Images_ |

PPTX2HTML
==========
[![MIT License][license-image]][license-url]

PPTX2HTML can convert MS-PPTX file to Reveal.js HTML by using pure javascript.  
Support Chrome, Firefox, IE>=10 and Edge.  
Here is the [Online DEMO](https://rhildred.github.io/PPTX2HTML) page.

I (Rich Hildred) needed this for remote teaching during the COVID 19 pandemic. I needed to convert the PowerPoint from when my 3rd year database course was taught in a classroom and a lab. I could then insert the lab demo right in to the now Reveal.js presentation.

Many thanks to g21589 for doing the heavy lifting.

Version
----

0.2.8 (Beta Test)

Support Objects
----
* Text
  * Font size
  * Font family
  * Font style: blod, italic, underline
  * Color
  * Location
  * hyperlink
* Picture
  * Type: jpg/jpeg, png, gif
  * Location
* Graph
  * Bar chart
  * Line chart
  * Pie chart
  * Scatter chart
* Table
  * Location
  * Size
* Text block (convert to Div)
  * Align (Horizontal and Vertical)
  * Background color (single color)
  * Border (borderColor, borderWidth, borderType, strokeDasharray)
* Drawing (convert to SVG)
  * Simple block (rect, ellipse, roundRect)
  * Background color (single color)
  * Align (Horizontal and Vertical)
  * Border (borderColor, borderWidth, borderType, strokeDasharray)
* Group/Multi-level Group
  * Level (z-index)
* Theme/Layout


License
----

MIT

[license-image]: http://img.shields.io/badge/license-MIT-blue.svg?style=flat
[license-url]: LICENSE
[Online DEMO]: http://g21589.github.io/PPTX2HTML

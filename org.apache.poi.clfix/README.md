# Apache Poi CLFIX

This fragment is required, so that we can use classes from
`org.apache.poi.poifs.crypt.agile` in our poi code.
This package is part of poi-ooxml, but not imported by the base poi bundle
itself. This leads to class loading issues during runtime.

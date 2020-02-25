
# PyECMA376-2

A Python implementation of the Open Packaging Conventions (OPC).

ECMA 376 Part 2 defines the “Open Packaging Conventions”, which is the packaging format to be used by the Office Open XML file formats.
It specifies, how to represent multiple logical files (“Parts”) within a physical Package (as a ZIP container), how to express semantic relationships between those Parts (using accompanying XML Parts), and how to add meta data and cryptographic signatures to the Package.
The format is defined in two steps: an abstract logic package model with Parts, Content Types and Relationships, and a physical mapping of this package model to PKZIP files.

This Python package aims to implement both, the logical model and physical mapping of OPC package files, to allow reading and writing such files.
However, it does not provide functionality to deal with the packages' payload, i.e. there is not functionality included to parse MS Word Documents from .docx files etc.


## Features of PyECMA376-2

* reading OPC package files
  * listing contained Parts (incl. Content Type)
  * reading Parts as file-like objects (incl. interleaved Parts)
  * parsing and following Relationships

* writing OPC package files
  * creating and writing Parts (via writable file-like objects, incl. interleaved Parts)
  * adding Relationships (as simple Python objects)
  * adding Content Type information

Modifying packages in-place is **not** supported.


### Currently Missing Features

* parsing/Writing package meta data (“Core Properties”) **(WIP)**
* retrieving/adding package thumbnail image
* reading/verifying/creating cryptographic signatures


## License

This package is developed by Michael Thies at the Chair of Process Control Engineering (PLT) at RWTH Aachen University.

It is published under the terms of Apache License v2.
See LICENSE file for details.

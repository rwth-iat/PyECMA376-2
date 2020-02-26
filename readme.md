
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


### Dependencies

This package requires `lxml` for XML reading and writing (with proper XML namespaces support).
Apart from that only the Python standard library is required.

The Python interpreter must support Python 3.6 or higher.


## Usage

Short example of reading an OPC package file:

```python
import pyecma376_2

with pyecma376_2.ZipPackageReader("document.docx") as reader:
    # List parts in package
    for part_name, content_type in reader.list_parts():
        print(part_name)
    
    # Get Relationship of type "…/officeDocument" from package-level Relationships
    document_part_name = reader.get_related_parts_by_type("/")[
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'][0]

    # Open part as (binary) file-like object
    with reader.open_part(document_part_name) as part:
        # XML parsing and document interpretation goes here
        print(part.read().decode())
```

Short example of creating and writing into an OPC package file:

```python
import pyecma376_2

with pyecma376_2.ZipPackageWriter("new_document.myx") as writer:
    # Add a part
    with writer.open_part("example/document.txt", "text/plain") as part:
        part.write("Lorem ipsum dolor sit amet.".encode())
    
    # Write the packages root relationships
    writer.write_relationships([
        pyecma376_2.OPCRelationship("r1", "http://example.com/my-package-relationship-id", "http://example.com",
                                    pyecma376_2.OPCTargetMode.EXTERNAL),
        pyecma376_2.OPCRelationship("r2", "http://example.com/my-document-rel", "example/document.txt",
                                    pyecma376_2.OPCTargetMode.INTERNAL),
    ])
    
    # The Content Types Stream with all parts' ContentTypes is automatically added when closing the package
    # Modify `writer.content_types` to change Content Types representation and use `writer.write_content_types_stream()`
    # for premature serialization/output.
```


## Package Architecture

The architecture of this package follows the logical concept of the ECMA standard:
The `package_model` module defines abstract `OPCPackageReader` and `OPCPackageWriter` classes that implement all the logical package model functionality, but omit the physical mapping to ZIP files.
This mapping is reflected in the abstract methods `list_items()`, `open_item()` and `create_item()` which are then implemented by the `ZipPackageReader` and `ZipPackageWriter` classes from the `zip_package` module.

Auxiliary classes and functions like `OPCRelationship`, `part_realpath` and `normalize_part_name` are also contained in the `package_model` module.


## License

This package is developed by Michael Thies at the Chair of Process Control Engineering (PLT) at RWTH Aachen University.

It is published under the terms of Apache License v2.
See LICENSE file for details.

 docx2pdf_rs

A command line utility to quickly convert large DOCX files to unformatted, minimal pdf files for LLM API consumption.

*The output is **not** very human.*

Performance overhead, this program compresses all non-jpeg media files into jepg via turbojpeg in PARALLEL using Rayon, 
this is fast for large documents with many images, but will incur a large mem tax as it needs to load all image files into memory.

Besides this, everything else is done via streaming and should be fine.

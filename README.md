# odt_to_txt #

**odt_to_txt.py** is a utility that uses [LibreOffice](https://libreoffice.org/) to convert document files to text files. This utility does not directly do file format conversion. It requires LibreOffice is installed.

The LibreOffice `--convert-to` command line parameter is used to do the conversion.

- [Starting LibreOffice Software With Parameters](https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html)

- [File Conversion Filter Names](https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html)

In addition to the **.odt** format used by LibreOffice, it can read Microsoft Word **.docx** and **.doc** files.


## Command Line Usage ##

```
usage: odt_to_txt.py [-h] [-r] [-o] [-d] [paths ...]

Run LibreOffice to convert document files to text files. Handles .odt, .doc,
and .docx file formats.

positional arguments:
  paths               One or more files and/or directories to process. Files
                      must be type (extension) '.odt', '.doc', or '.docx'. For
                      directories, all files with one of those extensions will
                      be processed.

options:
  -h, --help          show this help message and exit
  -r, --recurse       Do recursive search for document files in sub-
                      directories.
  -o, --overwrite     Overwrite existing output files. By default, existing
                      files are not replaced.
  -d, --datetime-tag  Add a [date_time] tag, based on the source document last
                      modified timestamp, to the output file names.
```

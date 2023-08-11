# Convert2PDF

![Visual Studio Code](https://img.shields.io/badge/Visual%20Studio%20Code-0078d7.svg?style=for-the-badge&logo=visual-studio-code&logoColor=white)
![GitHub](https://img.shields.io/badge/github-%23121011.svg?style=for-the-badge&logo=github&logoColor=white)
![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)

## Table of Contents

- [Convert2PDF](#convert2pdf)
  - [Table of Contents](#table-of-contents)
  - [About](#about)
  - [Getting Started](#getting-started)
    - [Requirements](#requirements)
  - [Usage](#usage)
    - [To convert all files in a directory](#to-convert-all-files-in-a-directory)
    - [To convert specific formats](#to-convert-specific-formats)
      - [For all Word Document files](#for-all-word-document-files)
      - [For all Powerpoint files](#for-all-powerpoint-files)
      - [For all Excel Spreadsheets](#for-all-excel-spreadsheets)
      - [For all Image files](#for-all-image-files)
    - [Missing file formats](#missing-file-formats)
  - [Roadmap](#roadmap)
  - [License](#license)

## About

A Python application that converts multiple Office documents into PDF format which removes the need of looking for online converter tools or converting the documents manually.

`Convert2PDF` takes in a file type as input and exports all matching file extensions for that Office format (eg. `.doc` or `.docx` for Word documents) and saves them in a separate directory, thus saving the hassle of converting all documents manually or looking for online converters.

_Inspired by [Convert2PDF](https://github.com/aditeyabaral/convert2pdf/tree/master)._

> _Note:_
>
> Since `comtypes` primarily supports Windows, `Convert2PDF` will not work on other platforms.
>
> This program also requires the user to have Microsoft Office applications installed on the computer.

## Getting Started

### Requirements

Only Python 3.11+ is tested and guaranteed to work.

To run the program, users can install the required libraries using `pip` and the `requirements.txt` file using the following command:

```python
pip install -r requirements.txt
```

## Usage

### To convert all files in a directory

Users can convert all files in a directory using `python Convert2PDF.py` or `python Convert2PDF.py -f *`.

### To convert specific formats

Users can also explicitly mention which files they would like to convert. To specify a particular type, pass in the respective format parameter as a command line argument.

#### For all Word Document files

```python
python cli.py -f word
```

#### For all Powerpoint files

```python
python cli.py -f ppt
```

#### For all Excel Spreadsheets

```python
python cli.py -f excel
```

#### For all Image files

```python
python cli.py -f img
```

### Missing file formats

A list of various file formats has been declared at the top section of the code. Don't see a file extension you need? You can add it in!

## Roadmap

- [ ] Create a Tkinter GUI for the application.

## License

Distributed under the MIT License. See `LICENSE` for more information.

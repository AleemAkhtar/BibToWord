# BibTeX to Word Bibliography Converter

This VBA macro facilitates the conversion of BibTeX-formatted references into Microsoft Word's bibliography format. By parsing BibTeX entries, it generates the corresponding XML structure and integrates the sources directly into your Word document.

## Features

- **BibTeX Parsing:** Processes various BibTeX entry types, including `@book`, `@article`, `@inproceedings`, and more.
- **Field Extraction:** Extracts essential fields such as title, year, volume, pages, and authors.
- **Author Handling:** Supports multiple authors, parsing their names into first, middle, and last components.
- **Word Integration:** Adds the converted sources to the document's bibliography and inserts citations at the current cursor position.

## Usage Instructions

1. **Prepare BibTeX Data:** Ensure your BibTeX entries are correctly formatted and accessible.
2. **Execute the Macro:** Run the `CommandButton1_Click` subroutine in your VBA environment.
3. **Review the Output:** The macro will add the parsed sources to your Word document's bibliography and insert citations where applicable.

## Limitations

- **Special Characters:** The current implementation does not handle special characters, such as the ampersand (`&`). Ensure your BibTeX data uses the word "and" instead.
- **Language Support:** The macro is designed for English-language entries and may not correctly process characters from other languages.

## Contribution

Contributions to enhance the macro's functionality, such as adding support for special characters and additional languages, are welcome. Please feel free to update and improve the code.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.

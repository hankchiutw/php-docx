# php-docx
Simple PHP docx parser class.

### How it works
A docx file is in fact some compressed files. After uncompressing, the contents are stored in `word/document.xml`.

Simply use PHP class `ZipArchive` to uncompress and `DOMDocument::loadXML` to parse the xml file.

### Features
- Convert contents to text
- Convert contents to html format with some style retained

### Demo

Execute
```sh
php -S localhost:3000
```
Then browse http://localhost:3000

# LOConverter
---

## Purpose
**LOConverter** is a library for converting between different document formats using LibreOffice.
It serves as a lightweight alternative to **jodconverter** with a simpler API.

---
## Usage

### Document Conversion
If the source file is accessible via an absolute path:

```java
LOConverter converter = new LOConverter(8080, "absolute/path/to/LibbreOffice/home");

try {
	converter.convert("absolute/path/to/source/file.extension1", "absolute/path/to/target/file.extension2");
} finally {
	converter.terminateOfficeProcess();
}
```

If the source file is available as an `InputStream`:

```java
LOConverter converter = new LOConverter(8080, "absolute/path/to/LibbreOffice/home");

try {
	converter.convert(inputStreamOfSourceFile, "absolute/path/to/target/file.extension");
} finally {
	converter.terminateOfficeProcess();
}
```

### Document Comparison
Currently, comparison is supported **only for DOCX files**:
```java
LOConverter converter = new LOConverter(8080, "absolute/path/to/LibbreOffice/home");

try {
	converter.compare(
		"absolute/path/to/first/file.docx",
		"absolute/path/to/second/file.docx",
		"absolute/path/to/result/file.docx"
	);
} finally {
	converter.terminateOfficeProcess();
}
```
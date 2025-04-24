package com.bakuard.lo;

public enum PdfFormat {
	PDFX("PDF/X", 0),
	PDFX1A2001("PDF/X-1a:2001", 1),
	PDFX32002("PDF/X-3:2002", 2),
	PDFA1A("PDF/A-1a", 3),
	PDFA1B("PDF/A-1b", 4);

	public static PdfFormat findByFormatName(String formatName) {
		for(PdfFormat format : values())
			if (format.getFormatName().equalsIgnoreCase(formatName))
				return format;
		return PDFA1B;
	}

	private final String pdfFormatName;
	private final int version;

	private PdfFormat(String pdfFormatName, int version) {
		this.pdfFormatName = pdfFormatName;
		this.version = version;
	}

	public String getFormatName() {
		return pdfFormatName;
	}

	public int getVersion() {
		return version;
	}
}

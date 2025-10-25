package com.bakuard.lo;

import com.sun.star.beans.PropertyValue;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

public class Properties {

	// Parameters for converting to PDF
	private static final String Pdf = "pdf";
	private static final String Format = "format";
	private static final String PdfOptionVersion = "SelectPdfVersion";
	private static final String PdfOptionPermissionPassword = "PermissionPassword";
	private static final String PdfOptionRestrictPermissions = "RestrictPermissions";
	private static final String PdfOptionChanges = "Changes";
	private static final int ChangesDeniedValue = 0;

	public static PropertyValue property(String name, Object value) {
		PropertyValue property = new PropertyValue();
		property.Name = name;
		property.Value = value;
		return property;
	}

	public static PropertyValue[] properties(String name, Object value) {
		return new PropertyValue[]{ property(name, value) };
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public static PropertyValue[] properties(Map<String, ?> parameters) {
		ArrayList<PropertyValue> properties = new ArrayList<>();

		for(Map.Entry<String, ?> entry : parameters.entrySet()) {
			String key = entry.getKey();
			Object value = entry.getValue();
			if(value instanceof Map)
				value = properties((Map)value);
			properties.add(property(key, value));
		}

		return properties.toArray(new PropertyValue[0]);
	}

	public static PropertyValue[] conversionProperties(String targetFileExtension, Map<String, ?> parameters) {
		return targetFileExtension.equalsIgnoreCase(Pdf) ? conversionPdfProperties(parameters) : properties(parameters);
	}


	private static PropertyValue[] conversionPdfProperties(Map<String, ?> parameters) {
		PdfFormat pdfVersion = PdfFormat.findByFormatName((String) parameters.get(Format));

		Map<String, Object> pdfOptions = new HashMap<>();
		pdfOptions.put(PdfOptionVersion, pdfVersion.getVersion());

		String permissionPassword = (String) parameters.get(PdfOptionPermissionPassword);

		if (permissionPassword != null && !permissionPassword.isEmpty()) {
			pdfOptions.put(PdfOptionRestrictPermissions, true);
			pdfOptions.put(PdfOptionPermissionPassword, permissionPassword);
			pdfOptions.put(PdfOptionChanges, ChangesDeniedValue);
		}

		return new PropertyValue[] {
				Properties.property("FilterName", "writer_pdf_Export"),
				Properties.property("FilterData", Properties.properties(pdfOptions))
		};
	}
}

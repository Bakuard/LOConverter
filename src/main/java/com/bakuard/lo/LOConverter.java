package com.bakuard.lo;

import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.DisposedException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XCloseable;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

public class LOConverter {

	private static final int ConnectionAttempts = 10;

	private final LibreOfficeProcess process;
	private final LibreOfficeContext currentContext;

	private final PropertiesSettings propertiesSettings;

	//Параметры для конвертации в PDF
	public static final String Format = "format";
	private static final String PdfOptionVersion = "SelectPdfVersion";
	private static final String PdfOptionPermissionPassword = "PermissionPassword";
	private static final String PdfOptionRestrictPermissions = "RestrictPermissions";
	private static final String PdfOptionChanges = "Changes";
	private static final int ChangesDeniedValue = 0;

	public LOConverter(int portNumber) {
		process = new LibreOfficeProcess(portNumber);
		currentContext = new LibreOfficeContext(process);
		propertiesSettings = new PropertiesSettings();
	}

	private void startOfficeProcess() {
		if(!isConnectionAlive()) {
			System.out.println("Start libreOffice process and connect...");
			process.start();
			currentContext.connectOfficeProcess(ConnectionAttempts);
		}
	}

	private boolean isConnectionAlive() {
		if(currentContext.getDispatchHelperInterface() == null)
			return false;

		try {
			currentContext.refreshCurrentFrame();
			currentContext.getDispatchHelperInterface().executeDispatch(
					currentContext.getDispatchProvider(),
					".uno:About",
					"", 0,
					new PropertyValue[]{}
			);
			return true;
		} catch(Exception e) {
			return false;
		}
	}

	public void terminateOfficeProcess() {
		try {
			currentContext.closeConnection();
		} catch (Exception e) {
			throw new RuntimeException("Fail to close XConnection", e);
		} finally {
			process.terminate();
		}
	}


	public void compareDocuments(String firstDocumentAbsolutPath, String secondDocumentAbsolutPath, String resultDocumentAbsolutPath) {
		startOfficeProcess();

		openDocument(firstDocumentAbsolutPath);
		compareDocument(secondDocumentAbsolutPath);
		saveDocumentAs(resultDocumentAbsolutPath, Properties.properties("FilterName", "MS Word 2007 XML"));
		closeDocument();
	}

	public void convertToPdf(InputStream source, String targetFileAbsolutPath, Map<String, String> parameters) {
		startOfficeProcess();

		openDocument(source);
		PropertyValue[] propertyValues = preparePropertiesForPdfConversion(parameters);
		saveDocumentAs(targetFileAbsolutPath, propertyValues);
		closeDocument();
	}

	public void convertTo(InputStream source, String targetFileAbsolutPath) {
		startOfficeProcess();

		openDocument(source);
		String documentFamily = getCurrentDocumentFamily();
		PropertyValue[] properties = propertiesSettings.getStoreProperties(documentFamily, FilenameUtils.getExtension(targetFileAbsolutPath));
		saveDocumentAs(targetFileAbsolutPath, properties);
		closeDocument();
	}


	private void openDocument(String sourceFileAbsolutPath) {
		try {
			XComponentLoader componentLoader = currentContext.getCompLoader();
			XComponent component = componentLoader.loadComponentFromURL(
					filePathToUri(sourceFileAbsolutPath), "_blank", 0,
					Properties.properties(Map.of(
							"ReadOnly", true,
							"UpdateDocMode", 0,
							"Hidden", true
					))
			);
			currentContext.setCurrentDocument(component);
			currentContext.refreshCurrentFrame();
		} catch(DisposedException e) {
			System.out.println("Connection with LibreOffice process was abrupted. Fail to open document: " + e);
			throw e;
		} catch (Exception e) {
			throw new RuntimeException("Fail to open document with LibreOffice.", e);
		}
	}

	private void openDocument(InputStream documentSource) {
		File tmpFile = inputStreamToTempFile(documentSource);
		openDocument(tmpFile.getAbsolutePath());
		tmpFile.delete();
	}

	private void closeDocument() {
		XCloseable closeable = UnoRuntime.queryInterface(XCloseable.class, currentContext.getCurrentDocument());
		try {
			if (closeable != null)
				closeable.close(true);
			else
				currentContext.getCurrentDocument().dispose();
		} catch(DisposedException e) {
			System.out.println("Connection with LibreOffice process was abrupted. Fail to close document: " + e);
			throw e;
		} catch (Exception e) {
			throw new RuntimeException("Fail to close document", e);
		}
	}

	private void compareDocument(String comparedFileAbsolutPath) {
		try {
			currentContext.getDispatchHelperInterface().executeDispatch(
					currentContext.getDispatchProvider(),
					".uno:CompareDocuments",
					"",
					0,
					new PropertyValue[]{
							Properties.property("URL", filePathToUri(comparedFileAbsolutPath))
					}
			);
		} catch(DisposedException e) {
			System.out.println("Connection with LibreOffice process was abrupted. Fail to compare document: " + e);
			throw e;
		} catch (Exception e) {
			throw new RuntimeException("Fail to compare with document: " + comparedFileAbsolutPath, e);
		}
	}

	private void saveDocumentAs(String newFileAbsolutPath, PropertyValue[] properties) {
		XStorable storable = UnoRuntime.queryInterface(XStorable.class, currentContext.getCurrentDocument());
		try {
			storable.storeToURL(filePathToUri(newFileAbsolutPath), properties);
		} catch(DisposedException e) {
			System.out.println("Connection with LibreOffice process was abrupted. Fail to save document as " + newFileAbsolutPath + ", reason: " + e);
			throw e;
		} catch(Exception e) {
			throw new RuntimeException("Fail to save document as " + newFileAbsolutPath, e);
		}
	}

	private String getCurrentDocumentFamily() {
		XServiceInfo serviceInfo = UnoRuntime.queryInterface(XServiceInfo.class, currentContext.getCurrentDocument());
		if (serviceInfo.supportsService("com.sun.star.text.WebDocument")) {
			return "WEB";
		} else if (serviceInfo.supportsService("com.sun.star.text.GenericTextDocument")) {
			return "TEXT";
		} else if (serviceInfo.supportsService("com.sun.star.sheet.SpreadsheetDocument")) {
			return "SPREADSHEET";
		} else if (serviceInfo.supportsService("com.sun.star.presentation.PresentationDocument")) {
			return "PRESENTATION";
		} else {
			return serviceInfo.supportsService("com.sun.star.drawing.DrawingDocument") ? "DRAWING" : null;
		}
	}


	private String filePathToUri(String filePath) {
		return new File(filePath).toURI().toString();
	}

	private File inputStreamToTempFile(InputStream in) {
		FileOutputStream out = null;
		try {
			final File tempFile = File.createTempFile("doczilla", "tmp-lo");
			out = new FileOutputStream(tempFile);
			IOUtils.copy(in, out);
			return tempFile;
		} catch(IOException e) {
			throw new RuntimeException("Fail copy InputStream to temp file while convert document with LibreOffice.", e);
		} finally {
			IOUtils.closeQuietly(out);
		}
	}

	private PropertyValue[] preparePropertiesForPdfConversion(Map<String, String> parameters) {
		PdfFormat pdfVersion = PdfFormat.findByFormatName(parameters.get(Format));

		Map<String, Object> pdfOptions = new HashMap<>();
		pdfOptions.put(PdfOptionVersion, pdfVersion.getVersion());

		String permissionPassword = parameters.get(PdfOptionPermissionPassword);

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

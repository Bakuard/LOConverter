package com.bakuard.lo;

import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.DisposedException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XCloseable;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.TimeUnit;

public class LOConverter {

	private static final Logger logger = LoggerFactory.getLogger(LOConverter.class.getName());

	private static final int ConnectionAttempts = 10;

	private final LOProcess process;
	private final LOContext currentContext;

	private final PropertiesSettings propertiesSettings;

	public LOConverter(int portNumber, String officeHome) {
		process = new LOProcess(portNumber, officeHome);
		currentContext = new LOContext(process);
		propertiesSettings = new PropertiesSettings();
	}

	private void startOfficeProcessAndConnect() {
		if(!isConnectionAlive()) {
			logger.info("Start libreOffice process and connect...");
			process.start();
			currentContext.connectOfficeProcess(ConnectionAttempts);
		} else {
			logger.debug("LibreOffice is already running and connected.");
		}
	}

	public void terminateOfficeProcess() {
		try {
			logger.info("Close connection and terminate libreOffice process...");
			currentContext.closeConnection();
		} catch (Exception e) {
			throw new RuntimeException("Fail to close XConnection correctly.", e);
		} finally {
			process.terminate();
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


	public void compare(String firstDocumentAbsolutPath, String secondDocumentAbsolutPath, String resultDocumentAbsolutPath) {
		startOfficeProcessAndConnect();

		openDocument(firstDocumentAbsolutPath);
		compareDocument(secondDocumentAbsolutPath);
		saveDocumentAs(resultDocumentAbsolutPath, Properties.properties("FilterName", "MS Word 2007 XML"));
		closeDocument();

		logger.info("Task 'compareDocuments' was completed.");
	}

	public void convert(InputStream source, String targetFileAbsolutPath) {
		convert(source, targetFileAbsolutPath, null);
	}

	public void convert(InputStream source, String targetFileAbsolutPath, Map<String, String> optionalParameters) {
		startOfficeProcessAndConnect();

		TimeoutTimer timer = null;
		try {
			timer = new TimeoutTimer(TimeUnit.MINUTES, 2, this);
			timer.start();

			openDocument(source);

			PropertyValue[] properties = null;
			String documentFamily = getCurrentDocumentFamily();
			String targetExtension = FilenameUtils.getExtension(targetFileAbsolutPath);

			if(optionalParameters != null && !optionalParameters.isEmpty())
				properties = Properties.conversionProperties(targetExtension, optionalParameters);
			else
				properties = propertiesSettings.getStorePropertiesByFileFamily(documentFamily, targetExtension);

			saveDocumentAs(targetFileAbsolutPath, properties);
			closeDocument();

			logger.info("Conversion from document family '{}' to file with extension '{}' was completed.", documentFamily, targetExtension);
		} finally {
			timer.cancel();
		}
	}

	public void convert(String sourceFileAbsolutPath, String targetFileAbsolutPath) {
		convert(sourceFileAbsolutPath, targetFileAbsolutPath, null);
	}

	public void convert(String sourceFileAbsolutPath, String targetFileAbsolutPath, Map<String, String> optionalParameters) {
		startOfficeProcessAndConnect();

		TimeoutTimer timer = null;
		try {
			timer = new TimeoutTimer(TimeUnit.MINUTES, 2, this);
			timer.start();

			openDocument(sourceFileAbsolutPath);

			PropertyValue[] properties = null;
			String targetExtension = FilenameUtils.getExtension(targetFileAbsolutPath);
			String sourceExtension = FilenameUtils.getExtension(sourceFileAbsolutPath);

			if(optionalParameters != null && !optionalParameters.isEmpty())
				properties = Properties.conversionProperties(targetExtension, optionalParameters);
			else
				properties = propertiesSettings.getStorePropertiesByExtensions(sourceExtension, targetExtension);

			saveDocumentAs(targetFileAbsolutPath, properties);
			closeDocument();

			logger.info("Conversion from '{}' to '{}' was completed.", sourceExtension, targetExtension);
		} finally {
			timer.cancel();
		}
	}


	private void openDocument(String sourceFileAbsolutPath) {
		try {
			HashMap<String, Object> defaultProperties = new HashMap<>();
			defaultProperties.put("UpdateDocMode", 0);
			defaultProperties.put("Hidden", true);

			XComponentLoader componentLoader = currentContext.getCompLoader();
			XComponent component = componentLoader.loadComponentFromURL(filePathToUri(sourceFileAbsolutPath), "_blank", 0, Properties.properties(defaultProperties));

			currentContext.setCurrentDocument(component);
			currentContext.refreshCurrentFrame();
		} catch(DisposedException e) {
			logger.error("Connection with LibreOffice process was abrupted. Fail to open document.", e);
			throw e;
		} catch (Exception e) {
			throw new RuntimeException("Fail to open document with LibreOffice.", e);
		}
	}

	private void openDocument(InputStream documentSource) {
		Path tmpFile = inputStreamToTempFile(documentSource);
		openDocument(tmpFile.toAbsolutePath().toString());
		try {
			Files.deleteIfExists(tmpFile);
		} catch(IOException e) {
			logger.warn("Fail to delete temporary file with document source.", e);
		}
	}

	private void closeDocument() {
		XCloseable closeable = UnoRuntime.queryInterface(XCloseable.class, currentContext.getCurrentDocument());
		try {
			if (closeable != null)
				closeable.close(true);
			else
				currentContext.getCurrentDocument().dispose();
		} catch(DisposedException e) {
			logger.error("Connection with LibreOffice process was abrupted. Fail to close document.", e);
			throw e;
		} catch (Exception e) {
			throw new RuntimeException("Fail to close document", e);
		}
	}

	private void compareDocument(String comparedFileAbsolutPath) {
		try {
			XFrame frame = UnoRuntime.queryInterface(XTextDocument.class, currentContext.getCurrentDocument()).getCurrentController().getFrame();
			XDispatchProvider dispatchProvider = UnoRuntime.queryInterface(XDispatchProvider.class, frame);

			currentContext.getDispatchHelperInterface().executeDispatch(
					dispatchProvider,
					".uno:CompareDocuments",
					frame.getName(),
					0,
					Properties.properties("URL", filePathToUri(comparedFileAbsolutPath))
			);
		} catch(DisposedException e) {
			logger.error("Connection with LibreOffice process was abrupted. Fail to compare document.", e);
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
			logger.error("Connection with LibreOffice process was abrupted. Fail to save document as {}", newFileAbsolutPath, e);
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
		return Paths.get(filePath).toUri().toString();
	}

	private Path inputStreamToTempFile(InputStream in) {
		try {
			Path tempFile = Files.createTempFile("loConverter-", "-document-from-InputStream");
			Files.copy(in, tempFile, StandardCopyOption.REPLACE_EXISTING);
			return tempFile;
		} catch(IOException e) {
			throw new RuntimeException("Fail copy InputStream to temp file while convert document with LibreOffice.", e);
		} finally {
			IOUtils.closeQuietly(in);
		}
	}


	private static class TimeoutTimer implements Runnable {

		private final LOConverter LOConverter;
		private final TimeUnit timeUnit;
		private final long duration;
		private Thread thread;

		public TimeoutTimer(TimeUnit timeUnit, long duration, LOConverter LOConverter) {
			this.timeUnit = timeUnit;
			this.duration = duration;
			this.LOConverter = LOConverter;
		}

		@Override
		public void run() {
			try {
				logger.debug("LibreOffice task interrupt timer has been started.");
				timeUnit.sleep(duration);
				logger.debug("Interrupt current LibreOffice task.");
				LOConverter.terminateOfficeProcess();
			} catch(InterruptedException e) {
				logger.debug("LibreOffice task interrupt timer has been canceled.");
			}
		}

		public void start() {
			thread = new Thread(this);
			thread.start();
		}

		public void cancel() {
			thread.interrupt();
		}
	}
}

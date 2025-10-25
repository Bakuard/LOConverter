package com.bakuard.lo;

import com.sun.star.beans.XPropertySet;
import com.sun.star.bridge.XBridge;
import com.sun.star.bridge.XBridgeFactory;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.connection.XConnection;
import com.sun.star.connection.XConnector;
import com.sun.star.frame.*;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lib.uno.helper.UnoUrl;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.concurrent.TimeUnit;

public class LOContext {

	private static final Logger logger = LoggerFactory.getLogger(LOContext.class.getName());

	private final LOProcess officeProcess;

	private final String unoUrlAsString;
	private final UnoUrl unoUrl;
	private XBridge bridge;

	private XComponentContext context;
	private XMultiComponentFactory multiCompFactory;

	private XDesktop desktopInterface;
	private XComponentLoader compLoader;
	private XDispatchHelper dispatchHelperInterface;
	private XDispatchProvider dispatchProvider;
	private XComponent currentDocument;

	public LOContext(LOProcess officeProcess) {
		this.officeProcess = officeProcess;
		this.unoUrlAsString = String.format("socket,host=127.0.0.1,port=%d,tcpNoDelay=1", officeProcess.getPort());
		this.unoUrl = getUnoUrl(unoUrlAsString);
	}

	/*
	* The LibreOffice process is not immediately ready to accept connections after it starts.
	* If the connection attempt fails at first, you should wait for a short while and then try again.
	* (This approach was taken from the JODConverter library.)
	*/
	public void connectOfficeProcess(int attemptNumber) {
		for(int i = 0; i < attemptNumber; i++) {
			try {
				officeProcess.waitOffice(1, TimeUnit.SECONDS);

				connect();
				initializeCompFactoryAndComponentContext();
				initializeDesktop();

				logger.info("LibreOffice process connection established");
				return;
			} catch(ProcessUnavailableException e) {
				throw e;
			} catch(Exception e) {
				logger.debug("Fail to connect to LibreOffice. Attempt {}/{}. Reason: {}", i + 1, attemptNumber, e.getMessage());
			}
		}

		throw new ProcessUnavailableException("Failed to connect to LibreOffice process after " + attemptNumber + " attempts.");
	}

	public void closeConnection() {
		if(bridge != null) {
			XComponent bridgeComp = UnoRuntime.queryInterface(XComponent.class, bridge);
			bridgeComp.dispose();
			logger.info("Connection with LibreOffice process was closed.");
		}
	}

	public XComponentLoader getCompLoader() {
		return compLoader;
	}

	public XDispatchHelper getDispatchHelperInterface() {
		return dispatchHelperInterface;
	}

	public XDispatchProvider getDispatchProvider() {
		return dispatchProvider;
	}

	public void refreshCurrentFrame() {
		XFrame frame = desktopInterface.getCurrentFrame();
		dispatchProvider = UnoRuntime.queryInterface(XDispatchProvider.class, frame);
	}

	public void setCurrentDocument(XComponent currentDocument) {
		this.currentDocument = currentDocument;
	}

	public XComponent getCurrentDocument() {
		return currentDocument;
	}


	private void connect() throws Exception {
		XComponentContext context = Bootstrap.createInitialComponentContext(null);
		XMultiComponentFactory multiCompFactory = context.getServiceManager();
		XConnection connection = createXConnection(context, multiCompFactory);
		bridge = createBridge(connection, context, multiCompFactory);
	}

	private void initializeCompFactoryAndComponentContext() throws Exception {
		Object remoteBridge = bridge.getInstance(unoUrl.getRootOid());
		multiCompFactory = UnoRuntime.queryInterface(XMultiComponentFactory.class, remoteBridge);
		XPropertySet propertySetInterface = UnoRuntime.queryInterface(XPropertySet.class, multiCompFactory);
		Object defaultContext = propertySetInterface.getPropertyValue("DefaultContext");
		context = UnoRuntime.queryInterface(XComponentContext.class, defaultContext);
	}

	private void initializeDesktop() throws Exception {
		Object desktopService = multiCompFactory.createInstanceWithContext("com.sun.star.frame.Desktop", context);
		desktopInterface = UnoRuntime.queryInterface(XDesktop.class, desktopService);
		compLoader = UnoRuntime.queryInterface(XComponentLoader.class, desktopService);
		Object dispatchHelperService = multiCompFactory.createInstanceWithContext("com.sun.star.frame.DispatchHelper", context);
		dispatchHelperInterface = UnoRuntime.queryInterface(XDispatchHelper.class, dispatchHelperService);
	}


	private XConnection createXConnection(XComponentContext context, XMultiComponentFactory multiCompFactory) throws Exception {
		Object connector = multiCompFactory.createInstanceWithContext("com.sun.star.connection.Connector", context);
		XConnector connectorInterface = UnoRuntime.queryInterface(XConnector.class, connector);
		return connectorInterface.connect(unoUrlAsString);
	}

	private XBridge createBridge(XConnection connection, XComponentContext context, XMultiComponentFactory multiCompFactory) throws Exception {
		Object bridgeFactory = multiCompFactory.createInstanceWithContext("com.sun.star.bridge.BridgeFactory", context);
		XBridgeFactory bridgeFactoryInterface = UnoRuntime.queryInterface(XBridgeFactory.class, bridgeFactory);
		String bridgeName = "DocumentConverterBridge(port=" + officeProcess.getPort() + ")";
		return bridgeFactoryInterface.createBridge(bridgeName, unoUrl.getProtocolAndParametersAsString(), connection, null);
	}

	private static UnoUrl getUnoUrl(String url) {
		try {
			return UnoUrl.parseUnoUrl(url + ";urp;StarOffice.ServiceManager");
		} catch (IllegalArgumentException e) {
			throw new java.lang.IllegalArgumentException(e.getMessage(), e);
		}
	}
}

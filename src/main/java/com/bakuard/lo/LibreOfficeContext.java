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

import java.util.concurrent.TimeUnit;

public class LibreOfficeContext {

	private final LibreOfficeProcess officeProcess;

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

	public LibreOfficeContext(LibreOfficeProcess officeProcess) {
		this.officeProcess = officeProcess;
		this.unoUrlAsString = String.format("socket,host=127.0.0.1,port=%d,tcpNoDelay=1", officeProcess.getPort());
		this.unoUrl = getUnoUrl(unoUrlAsString);
	}

	/*
	 * Процесс LibreOffice не сразу готов принимать соединения после своего запуска. Если
	 * не получилось сразу соединиться, нужно подождать некоторое время, а затем попробовать снова.
	 * (Этот подход был взят из библиотеки JODConverter).
	 */
	public void connectOfficeProcess(int attemptNumber) {
		Exception officeConnectException = null;

		for(int i = 0; i < attemptNumber; i++) {
			try {
				connectOfficeProcess();
				return;
			} catch(Exception e) {
				System.out.println("Fail to connect to LibreOffice. Attempt " + (i + 1) + "/" + attemptNumber);
				officeConnectException = e;
			}
		}

		throw new RuntimeException("Fail to connect to LibreOffice process.", officeConnectException);
	}

	private void connectOfficeProcess() throws Exception {
		officeProcess.waitOffice(1, TimeUnit.SECONDS);

		connect();
		initializeCompFactoryAndComponentContext();
		initializeDesktop();

		System.out.println("LibreOffice process connection established on port " + officeProcess.getPort());
	}

	public void closeConnection() {
		if(bridge != null) {
			XComponent bridgeComp = UnoRuntime.queryInterface(XComponent.class, bridge);
			bridgeComp.dispose();
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

	public XComponent getCurrentDocument() {
		return currentDocument;
	}

	public void refreshCurrentFrame() {
		XFrame frame = desktopInterface.getCurrentFrame();
		dispatchProvider = UnoRuntime.queryInterface(XDispatchProvider.class, frame);
	}

	public void setCurrentDocument(XComponent currentDocument) {
		this.currentDocument = currentDocument;
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

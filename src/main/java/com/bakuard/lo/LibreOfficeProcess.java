package com.bakuard.lo;

import org.apache.commons.io.IOUtils;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class LibreOfficeProcess {

	private final String acceptString;
	private final ProcessBuilder processBuilder;
	private final int port;
	private Process process;

	public LibreOfficeProcess(int port) {
		this.port = port;
		this.acceptString = String.format("socket,host=127.0.0.1,port=%d,tcpNoDelay=1;urp;StarOffice.ServiceManager", port);
		this.processBuilder = new ProcessBuilder(prepareCommandForStartOffice());
	}

	public void start() {
		if(isLibreOfficeProcessExists())
			return;

		try {
			process = processBuilder.start();
			System.out.println("LibreOffice process was started.");
		} catch (IOException e) {
			throw new RuntimeException("Fail to start LibreOffice process.", e);
		}
	}

	public void terminate() {
		if(process == null)
			return;

		process.destroy();
		if(process.isAlive())
			process.destroyForcibly();
		process = null;

		System.out.println("LibreOffice process was terminated with exit code.");
	}

	public void waitOffice(long timeUnitNumber, TimeUnit timeUnit) throws Exception {
		if(process != null)
			process.waitFor(timeUnitNumber, timeUnit);
	}

	public int getPort() {
		return port;
	}


	private List<String> prepareCommandForStartOffice() {
		List<String> args = new ArrayList<>();
		args.add(getOfficeExecutable());
		args.add("--accept=" + acceptString);
		args.add("--headless");
		args.add("--invisible");
		args.add("--nocrashreport");
		args.add("--nodefault");
		args.add("--nofirststartwizard");
		args.add("--nolockcheck");
		args.add("--nologo");
		args.add("--norestore");
		args.add("-env:UserInstallation=" + createTempFileUri());
		return args;
	}

	private String createTempFileUri() {
		File tempDir = new File(System.getProperty("java.io.tmpdir"));
		String fileName = ".z8converter_socket_host-127.0.0.1_port-" + port + "_tcpNoDelay-1";
		return new File(tempDir, fileName).toURI().toString();
	}

	private static String getOfficeExecutable() {
		String officeHome = "C:/Program Files/LibreOffice";
		File file = new File(officeHome);
		if(!file.isDirectory())
			return officeHome;

		return new File(officeHome, "program/soffice" + (isWindows() ? ".exe" : "")).getAbsolutePath();
	}

	private static boolean isWindows() {
		String osName = System.getProperty("os.name");
		return osName != null && osName.toLowerCase().startsWith("windows");
	}


	private boolean isLibreOfficeProcessExists() {
		ProcessBuilder processBuilder = null;
		Pattern processGetLine = null;

		if(isWindows()) {
			processBuilder = new ProcessBuilder("cmd", "/c", "wmic process where(name like 'soffice%') get commandline,processid");
			processGetLine = Pattern.compile("^\\s*(?<CommandLine>.*?)\\s+(?<Pid>\\d+)\\s*$");
		} else {
			processBuilder = new ProcessBuilder( "/bin/sh", "-c", "/bin/ps -e -o pid,args | /bin/grep soffice | /bin/grep -v grep");
			processGetLine = Pattern.compile("^\\s*(?<Pid>\\d+)\\s+(?<CommandLine>.*)$");
		}

		List<String> commands = findAnyOfficeProcesses(processBuilder);
		boolean isProcessExists = matchOfficeProcesses(commands, processGetLine);

		System.out.println("Check if LibreOffice process already exists: " + isProcessExists);
		return isProcessExists;
	}

	private List<String> findAnyOfficeProcesses(ProcessBuilder officeProcessesFinder) {
		List<String> commands = new ArrayList<>();

		BufferedReader reader = null;
		try {
			Process process = officeProcessesFinder.start();
			reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
			String line = reader.readLine();
			while(line != null) {
				if(!line.isBlank())
					commands.add(line);
				line = reader.readLine();
			}
		} catch(IOException e) {
			System.out.println("Fail to read existed process: " + e);
		} finally {
			IOUtils.closeQuietly(reader);
		}
		return commands;
	}

	private boolean matchOfficeProcesses(List<String> processCommands, Pattern processLineExtracter) {
		final Pattern commandPattern = Pattern.compile(Pattern.quote("soffice") + ".*" + Pattern.quote("--accept=" + acceptString));
		for(String line : processCommands) {
			Matcher matcher = processLineExtracter.matcher(line);
			if(matcher.matches()) {
				String commandLine = matcher.group("CommandLine");
				Matcher commandMatcher = commandPattern.matcher(commandLine);
				if(commandMatcher.find()) {
					return true;
				}
			}
		}
		return false;
	}
}

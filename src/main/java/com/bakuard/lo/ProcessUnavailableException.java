package com.bakuard.lo;

public class ProcessUnavailableException extends RuntimeException {

	public ProcessUnavailableException(String message) {
		super(message);
	}

	public ProcessUnavailableException(String message, Throwable cause) {
		super(message, cause);
	}

}

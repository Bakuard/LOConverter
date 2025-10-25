package com.bakuard.lo;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import com.sun.star.beans.PropertyValue;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.*;
import java.util.stream.Collectors;

public class PropertiesSettings {

	private static class FamilyAndExtension {
		private final String sourceFileFamily;
		private final String targetFileExtension;

		public FamilyAndExtension(String sourceFileFamily, String targetFileExtension) {
			this.sourceFileFamily = sourceFileFamily;
			this.targetFileExtension = targetFileExtension;
		}

		@Override
		public boolean equals(Object o) {
			if(o == null || getClass() != o.getClass()) return false;
			FamilyAndExtension that = (FamilyAndExtension) o;
			return Objects.equals(sourceFileFamily, that.sourceFileFamily)
						   && Objects.equals(targetFileExtension, that.targetFileExtension);
		}

		@Override
		public int hashCode() {
			return Objects.hash(sourceFileFamily, targetFileExtension);
		}
	}

	private final Map<FamilyAndExtension, PropertyValue[]> storeProperties = new HashMap<>();
	private final Set<String> supportedFileFormats = new HashSet<>();
	private final Map<String, String> fileFamilyByExtension = new HashMap<>();

	public PropertiesSettings() {
		parseStoreProperties(loadSettings());
	}

	public PropertyValue[] getStorePropertiesByExtensions(String sourceFileExtension, String targetFileExtension) {
		String sourceFileFamily = fileFamilyByExtension.get(sourceFileExtension);
		return getStorePropertiesByFileFamily(sourceFileFamily, targetFileExtension);
	}

	public PropertyValue[] getStorePropertiesByFileFamily(String sourceFileFamily, String targetFileExtension) {
		assertSourceFileFamilyNotNull(sourceFileFamily);
		assertTargetFileFormatAreSupported(targetFileExtension);

		PropertyValue[] properties = storeProperties.get(new FamilyAndExtension(sourceFileFamily, targetFileExtension));
		if(properties == null)
			throw new RuntimeException("Unsupported conversion: " + sourceFileFamily + " --> " + targetFileExtension);
		return properties;
	}


	private JsonArray loadSettings() {
		try(InputStream inputStream = getClass().getClassLoader().getResourceAsStream("documents-formats.json")) {
			BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream));
			String data = reader.lines().collect(Collectors.joining("\n"));
			return JsonParser.parseString(data).getAsJsonArray();
		} catch(IOException e) {
			throw new RuntimeException("Fail to load settings for file conversion with LibreOffice.", e);
		}
	}

	private void parseStoreProperties(JsonArray settings) {
		for(int i = 0; i < settings.size(); i++) {
			JsonObject fileFormatSettings = settings.get(i).getAsJsonObject();
			String inputFamily = fileFormatSettings.get("inputFamily").getAsString();
			JsonArray extensions = fileFormatSettings.get("extensions").getAsJsonArray();

			for(int j = 0; j < extensions.size(); j++) {
				String extension = extensions.get(j).getAsString();
				supportedFileFormats.add(extension);
				fileFamilyByExtension.put(extension, inputFamily);
			}

			Map<String, Object> fileFormatProperties = new HashMap<>();
			JsonObject rowStoreProperties = fileFormatSettings.get("storeProperties").getAsJsonObject();
			for(String fileFamilyName : rowStoreProperties.keySet()) {
				JsonObject fileFamily = rowStoreProperties.get(fileFamilyName).getAsJsonObject();
				for(String propertyName : fileFamily.keySet())
					fileFormatProperties.put(propertyName, fileFamily.get(propertyName).getAsString());

				extensions = fileFormatSettings.get("extensions").getAsJsonArray();
				for(int j = 0; j < extensions.size(); j++) {
					String extension = extensions.get(j).getAsString();
					FamilyAndExtension familyAndExtension = new FamilyAndExtension(fileFamilyName, extension);
					storeProperties.put(familyAndExtension, Properties.properties(fileFormatProperties));
				}
			}
		}
	}

	private void assertTargetFileFormatAreSupported(String fileFormat) {
		if(!supportedFileFormats.contains(fileFormat))
			throw new RuntimeException("Unsupported file format: " + fileFormat);
	}

	private void assertSourceFileFamilyNotNull(String sourceFileFamily) {
		if(sourceFileFamily == null)
			throw new RuntimeException("LibreOffice can't determine file family for source file.");
	}
}

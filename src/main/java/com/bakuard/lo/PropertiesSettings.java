package com.bakuard.lo;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import com.sun.star.beans.PropertyValue;
import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

public class PropertiesSettings {

	private record FamilyAndExtension(String sourceFileFamily, String targetFileExtension) {}

	private final Map<FamilyAndExtension, PropertyValue[]> storeProperties;
	private final Set<String> supportedFileFormats;

	public PropertiesSettings() {
		JsonArray settings = loadSettings();
		storeProperties = extractStoreProperties(settings);
		supportedFileFormats = extractSupportedFileFormats(settings);
	}

	public PropertyValue[] getStoreProperties(String sourceFileFamily, String targetFileExtension) {
		assertSourceFileFamilyNotNull(sourceFileFamily);
		assertFileFormatSettingsExists(targetFileExtension);
		FamilyAndExtension familyAndExtension = new FamilyAndExtension(sourceFileFamily, targetFileExtension);
		PropertyValue[] properties = storeProperties.get(familyAndExtension);
		if(properties == null)
			throw new RuntimeException("Unsupported conversion: " + sourceFileFamily + " --> " + targetFileExtension);
		return properties;
	}


	private JsonArray loadSettings() {
		try {
			File file = new File(getClass().getClassLoader().getResource("./documents-formats.json").getFile());
			String data = FileUtils.readFileToString(file, "UTF-8");
			return JsonParser.parseString(data).getAsJsonArray();
		} catch(IOException e) {
			throw new RuntimeException("Fail to load settings for file conversion with LibreOffice.", e);
		}
	}

	private Map<FamilyAndExtension, PropertyValue[]> extractStoreProperties(JsonArray settings) {
		Map<FamilyAndExtension, PropertyValue[]> storePropertiesByFamilyAndExtension = new HashMap<>();

		for(int i = 0; i < settings.size(); i++) {
			JsonObject fileFormatSettings = settings.get(i).getAsJsonObject();

			Map<String, Object> storeProperties = new HashMap<>();
			JsonObject rowStoreProperties = fileFormatSettings.get("storeProperties").getAsJsonObject();
			for(String fileFamilyName : rowStoreProperties.keySet()) {
				JsonObject fileFamily = rowStoreProperties.get(fileFamilyName).getAsJsonObject();
				for(String propertyName : fileFamily.keySet())
					storeProperties.put(propertyName, fileFamily.get(propertyName).getAsString());

				JsonArray extensions = fileFormatSettings.get("extensions").getAsJsonArray();
				for(int j = 0; j < extensions.size(); j++) {
					String extension = extensions.get(j).getAsString();
					FamilyAndExtension familyAndExtension = new FamilyAndExtension(fileFamilyName, extension);
					storePropertiesByFamilyAndExtension.put(familyAndExtension, Properties.properties(storeProperties));
				}
			}
		}

		return storePropertiesByFamilyAndExtension;
	}

	private Set<String> extractSupportedFileFormats(JsonArray settings) {
		Set<String> supportedFileFormats = new HashSet<>();
		for(int i = 0; i < settings.size(); i++) {
			JsonObject fileFormatSettings = settings.get(i).getAsJsonObject();
			JsonArray extensions = fileFormatSettings.get("extensions").getAsJsonArray();
			for(int j = 0; j < extensions.size(); j++)
				supportedFileFormats.add(extensions.get(j).getAsString());
		}
		return supportedFileFormats;
	}

	private void assertFileFormatSettingsExists(String fileFormat) {
		if(!supportedFileFormats.contains(fileFormat))
			throw new RuntimeException("Unsupported file format: " + fileFormat);
	}

	private void assertSourceFileFamilyNotNull(String sourceFileFamily) {
		if(sourceFileFamily == null)
			throw new RuntimeException("LibreOffice can't determine file family for source file.");
	}
}

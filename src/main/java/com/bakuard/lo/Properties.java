package com.bakuard.lo;

import com.sun.star.beans.PropertyValue;

import java.util.ArrayList;
import java.util.Map;

public class Properties {

	public static PropertyValue property(String name, Object value) {
		PropertyValue property = new PropertyValue();
		property.Name = name;
		property.Value = value;
		return property;
	}

	public static PropertyValue[] properties(String name, Object value) {
		return new PropertyValue[]{property(name, value)};
	}

	public static PropertyValue[] properties(Map<String, Object> parameters) {
		ArrayList<PropertyValue> properties = new ArrayList<>();

		for(Map.Entry<String, Object> entry : parameters.entrySet())
			properties.add(property(entry.getKey(), entry.getValue()));

		return properties.toArray(new PropertyValue[0]);
	}

}

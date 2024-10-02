package core

// Helper function to get a boolean value from a map
func getBoolValue(targetMap map[string]interface{}, key string, defaultValue bool) bool {
	if val, ok := targetMap[key]; ok && val != nil {
		return val.(bool)
	}
	return defaultValue
}

// Helper function to get a float64 value from a map
func getFloat64Value(targetMap map[string]interface{}, key string, defaultValue float64) float64 {
	if val, ok := targetMap[key]; ok && val != nil {
		return val.(float64)
	}
	return defaultValue
}

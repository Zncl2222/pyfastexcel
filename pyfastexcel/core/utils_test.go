package core

import (
	"testing"
)

// Test for getBoolValue function
func TestGetBoolValue(t *testing.T) {
	tests := []struct {
		name         string
		targetMap    map[string]interface{}
		key          string
		defaultValue bool
		expected     bool
	}{
		{
			name:         "Key exists and value is true",
			targetMap:    map[string]interface{}{"key1": true},
			key:          "key1",
			defaultValue: false,
			expected:     true,
		},
		{
			name:         "Key exists and value is false",
			targetMap:    map[string]interface{}{"key1": false},
			key:          "key1",
			defaultValue: true,
			expected:     false,
		},
		{
			name:         "Key does not exist, use default value",
			targetMap:    map[string]interface{}{"key2": false},
			key:          "key1",
			defaultValue: true,
			expected:     true,
		},
		{
			name:         "Key exists but value is nil, use default value",
			targetMap:    map[string]interface{}{"key1": nil},
			key:          "key1",
			defaultValue: true,
			expected:     true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := getBoolValue(tt.targetMap, tt.key, tt.defaultValue)
			if result != tt.expected {
				t.Errorf("expected %v, got %v", tt.expected, result)
			}
		})
	}
}

// Test for getFloat64Value function
func TestGetFloat64Value(t *testing.T) {
	tests := []struct {
		name         string
		targetMap    map[string]interface{}
		key          string
		defaultValue float64
		expected     float64
	}{
		{
			name:         "Key exists and value is valid float64",
			targetMap:    map[string]interface{}{"key1": 3.14},
			key:          "key1",
			defaultValue: 1.23,
			expected:     3.14,
		},
		{
			name:         "Key does not exist, use default value",
			targetMap:    map[string]interface{}{"key2": 2.71},
			key:          "key1",
			defaultValue: 1.23,
			expected:     1.23,
		},
		{
			name:         "Key exists but value is nil, use default value",
			targetMap:    map[string]interface{}{"key1": nil},
			key:          "key1",
			defaultValue: 1.23,
			expected:     1.23,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := getFloat64Value(tt.targetMap, tt.key, tt.defaultValue)
			if result != tt.expected {
				t.Errorf("expected %v, got %v", tt.expected, result)
			}
		})
	}
}

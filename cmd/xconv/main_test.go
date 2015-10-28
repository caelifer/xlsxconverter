package main

import (
	"bytes"
	"io"
	"os"
	"testing"
)

func TestMain(t *testing.T) {
	// Mock os.Stdout
	r, w, _ := os.Pipe()
	out := os.Stdout
	os.Stdout = w

	// Instrument output capturing
	oc := make(chan string)
	go func() {
		var buf bytes.Buffer
		// t.Log("About to io.Copy")
		if _, err := io.Copy(&buf, r); err != nil {
			t.Errorf("Failed to caputer output: %v", err)
		}
		// t.Log("Done io.Copy")
		oc <- buf.String()
	}()

	// Call tested function in a closure
	c := make(chan struct{})
	go func() {
		defer func() {
			close(c)
			if r := recover(); r != nil {
				t.Fatalf("panic in main(): %v", r)
			}
		}()
		main()
	}()
	<-c

	// Closing mocked os.Stdout to allow io.Copy
	w.Close()
	// Capture actual output
	res := <-oc
	// Restore state
	os.Stdout = out

	// Test captured output
	want := "Hello\n"
	if res != want {
		t.Errorf("Expected: %q, got: %q", want, res)
	}
}

// vim: :ts=4:sw=4:noexpandtab:ai

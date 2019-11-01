# doc-to-text

WORK IN PROGRESS I mean it can't read whole document for now.

Atm your best chance to extract text from .doc file is to use 
catdoc command tool with some simple Go code as follows:

```go
func TestCatDoc(t *testing.T) {
	path := filepath.Join("testdata", "file-sample_100kB.doc")
	fileData, err := ioutil.ReadFile(path)
	if err != nil {
		t.Error(err)
	}
	cmd := exec.Command("catdoc")
	stdin, err := cmd.StdinPipe()
	if err != nil {
		t.Error(err)
	}
	go func() {
		defer stdin.Close()
		_, err := io.WriteString(stdin, string(fileData))
		if err != nil {
			t.Error(err)
		}
	}()
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Error(err)
	}
	fmt.Printf("%s\n", out)
}
```

pure go library to read text from .doc files without using any external tools, commands etc

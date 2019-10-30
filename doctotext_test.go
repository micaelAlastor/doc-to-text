package doctotext

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"testing"

	"github.com/richardlehane/mscfb"
)

func TestDocToText(t *testing.T) {
	path := filepath.Join("testdata", "file-sample_100kB.doc")
	file, err := os.Open(path)
	defer file.Close()
	text, err := DocToText(file)
	if err != nil {
		t.Error(err)
	}
	fmt.Println(text)
}

func TestDocToText2(t *testing.T) {
	path := filepath.Join("testdata", "lefthand.doc")
	file, err := os.Open(path)
	defer file.Close()
	text, err := DocToText(file)
	if err != nil {
		t.Error(err)
	}
	fmt.Println(text)
}

func TestMscfbDoc(t *testing.T) {
	path := filepath.Join("testdata", "file-sample_100kB.doc")
	file, err := os.Open(path)
	defer file.Close()
	doc, err := mscfb.New(file)
	if err != nil {
		log.Fatal(err)
	}
	for entry, err := doc.Next(); err == nil; entry, err = doc.Next() {
		fmt.Println(entry.Name)
		fmt.Println(entry.Size)
		buf := make([]byte, 512)
		i, _ := doc.Read(buf)
		if i > 0 {
			fmt.Println(buf[:i])
		}
		fmt.Println("---------------------")
	}
}

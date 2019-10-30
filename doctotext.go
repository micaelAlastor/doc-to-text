package doctotext

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"os"
	"unicode/utf16"
	"unicode/utf8"

	"github.com/richardlehane/mscfb"
	//"golang.org/x/text/encoding/charmap"
)

func ReadBytesAt(source []byte, offset int64, size int64) ([]byte, error) {
	result := make([]byte, size)
	reader := bytes.NewReader(source)
	_, err := reader.ReadAt(result, offset)
	return result, err
}

func ToInt32(source []byte, offset int64) (uint32, error) {
	data, err := ReadBytesAt(source, offset, 4)
	return binary.LittleEndian.Uint32(data), err
}

// UTF16BytesToString converts UTF-16 encoded bytes, in big or little endian byte order,
// to a UTF-8 encoded string.
func UTF16BytesToString(b []byte, o binary.ByteOrder) string {
	utf := make([]uint16, (len(b)+(2-1))/2)
	for i := 0; i+(2-1) < len(b); i += 2 {
		utf[i/2] = o.Uint16(b[i:])
	}
	if len(b)/2 < len(utf) {
		utf[len(utf)-1] = utf8.RuneError
	}
	return string(utf16.Decode(utf))
}

func DocToText(file *os.File) (string, error) {
	result := ""
	doc, err := mscfb.New(file)
	if err != nil {
		return "", err
	}

	var wordDocumentEntry *mscfb.File
	var tableEntry *mscfb.File
	fib := make([]byte, 1472)

	for entry, err := doc.Next(); err == nil; entry, err = doc.Next() {
		if entry.Name == "WordDocument" {
			wordDocumentEntry = entry
		}
		if entry.Name == "0Table" || entry.Name == "1Table" {
			tableEntry = entry
		}
	}

	if wordDocumentEntry == nil {
		return "", err
	}
	_, err = wordDocumentEntry.Read(fib)
	if err != nil {
		return "", err
	}

	fcOffset := 0x01A2
	lcbOffset := 0x01A6

	fcClx, err := ToInt32(fib, int64(fcOffset))
	lcbClx, err := ToInt32(fib, int64(lcbOffset))
	if err != nil {
		return "", err
	}

	clx := make([]byte, lcbClx)

	if tableEntry == nil {
		return "", err
	}
	_, err = tableEntry.ReadAt(clx, int64(fcClx))
	if err != nil {
		return "", err
	}

	/*The clx byte array can contain multiple substructures and only one of these substructures is the piece
	table. Each substructure starts with a byte which denotes the type of the substructure, followed by a
	value indicating the length of the substructure.
		If the substructure describes a piece table the value of this byte is 2, otherwise 1. The length of the
	entry is a 32 bit value for a piece table and an 8 bit value for all other entries.*/

	var pieceTable []byte
	var lcbPieceTable uint32
	pos := 0
	goOn := true

	for goOn {
		typeEntry := clx[pos]

		if typeEntry == 2 {
			goOn = false
			lcbPieceTable, err = ToInt32(clx, int64(pos+1))
			if err != nil {
				return "", err
			}
			pieceTable = make([]byte, lcbPieceTable)
			copy(pieceTable, clx[pos+5:pos+int(lcbPieceTable)])
		} else if typeEntry == 1 {
			//skip this entry
			pos = pos + 1 + 1 + int(clx[pos+1])
		} else {
			goOn = false
		}
	}
	fmt.Println(pieceTable)

	/*The piece table itself contains two arrays:
	The first array contains n+1 logical character positions (n is the number of pieces). The
	entries are the logical start and end positions of the pieces in the text sequence, i.e. the first
	piece starts at logical position 1 and extends to position 2, the second starts at position 2,
		etc. Logical position x means that this is the x-th character in the document, i.e. this is not the
	file character position in the WordDocument stream. The positions are 32 bit values.
		The second array contains n piece descriptor structures. Each structure has a length of 8
	bytes. The physical location of the piece inside of the WordDocument stream and the
	encoding of the text can be found in these 8 bytes from byte 3 to byte 6. This file character
	(FC) position is a 32 bit integer value. The second most significant bit is a flag that specifies
	the encoding of the piece: if the bit is set, the piece is CP1252-encoded and the FC is a word
	pointer; otherwise, the piece is Unicode-encoded and the FC is a byte pointer.*/

	pieceCount := (lcbPieceTable - 4) / 12

	var cpStart, cpEnd uint32
	var pieceDescriptor []byte

	for i := 0; i < int(pieceCount); i++ {
		//get the position
		cpStart, err = ToInt32(pieceTable, int64(i*4))
		cpEnd, err = ToInt32(pieceTable, int64((i+1)*4))
		if err != nil {
			return "", err
		}

		//get the descriptor
		pieceDescriptor = make([]byte, 8)
		offsetPieceDescriptor := ((pieceCount + 1) * 4) + uint32(i*8)
		copy(pieceDescriptor, pieceTable[offsetPieceDescriptor:offsetPieceDescriptor+8])

		//The interpretation of the encoding flag and the calculation of the FC pointer are as follows:
		fcValue, err := ToInt32(pieceDescriptor, 2)
		if err != nil {
			return "", err
		}
		isANSI := (fcValue & 0x40000000) == 0x40000000
		fc := fcValue & 0xBFFFFFFF
		cb := cpEnd - cpStart
		encoding := "1251"
		if !isANSI {
			encoding = "UTF8"
			cb *= 2
		}
		fmt.Println("encoding is " + encoding)
		bytesOfText := make([]byte, cb)
		_, err = wordDocumentEntry.ReadAt(bytesOfText, int64(fc))
		if err != nil {
			return "", err
		}

		text := UTF16BytesToString(bytesOfText, binary.LittleEndian)

		fmt.Println(text)

		result += text
	}

	return result, nil
}

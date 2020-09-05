package main

import (
	"archive/zip"
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"

	"github.com/jessevdk/go-flags"
	log "github.com/sirupsen/logrus"
)

var opts struct {
	ListOfDirs      []string `long:"dir" description:"A list of dirs" required:"true"`
	Search          string   `long:"search" description:"Searched string" required:"true"`
	CaseInsensitive bool     `long:"case-insensitive" description:"Case insensitive"`
}

func main() {
	_, err := flags.ParseArgs(&opts, os.Args)
	if err != nil {
		log.Errorln(err.Error())
	}

	for _, dir := range opts.ListOfDirs {
		filepath.Walk(dir, func(path string, _ os.FileInfo, err error) error {
			if err != nil {
				log.Errorf("%s : %s\n", path, err.Error())
			}

			findSearchInFile(path, opts.Search, opts.CaseInsensitive)
			return nil
		})
	}
}

func findSearchInFile(path string, search string, caseInsensitive bool) {
	if strings.HasSuffix(path, ".xlsx") {
		findSearchInXlsx(path, search, caseInsensitive)
	} else if strings.HasSuffix(path, ".ods") {
		findSearchInOds(path, search, caseInsensitive)
	} else if strings.HasSuffix(path, ".docx") {
		findSearchInDocx(path, search, caseInsensitive)
	} else if strings.HasSuffix(path, ".odt") {
		findSearchInOdt(path, search, caseInsensitive)
	}
}

func findSearchInFileInsideZip(filename, search string, caseInsensitive bool, predicate func(*zip.File) bool) {
	r, err := zip.OpenReader(filename)

	if err != nil {
		return
	}
	defer r.Close()

	for _, f := range r.File {
		if predicate(f) {
			rc, err := f.Open()
			if err != nil {
				log.Errorln(err)
				continue
			}
			defer rc.Close()

			dat, err := ioutil.ReadAll(rc)
			if err != nil {
				log.Errorln(err)
				continue
			}

			datString := string(dat)
			if caseInsensitive && strings.Contains(
				strings.ToLower(datString),
				strings.ToLower(search),
			) {
				fmt.Println(strings.Replace(filename, " ", "\\ ", -1))
				continue
			}
			
			if strings.Contains(datString, search) {
				fmt.Println(strings.Replace(filename, " ", "\\ ", -1))
			}
		}
	}
}

func findSearchInXlsx(filename string, search string, caseInsensitive bool) {
	findSearchInFileInsideZip(filename, search, caseInsensitive, func(f *zip.File) bool {
		isSheet := strings.HasPrefix(f.Name, "xl/worksheets/") && strings.HasSuffix(f.Name, ".xml")
		isSharedStrings := f.Name == "xl/sharedStrings.xml"
		return isSheet || isSharedStrings
	})
}

func findSearchInOds(filename string, search string, caseInsensitive bool) {
	findSearchInFileInsideZip(filename, search, caseInsensitive, func(f *zip.File) bool {
		return f.Name == "content.xml"
	})
}

func findSearchInDocx(filename string, search string, caseInsensitive bool) {
	findSearchInFileInsideZip(filename, search, caseInsensitive, func(f *zip.File) bool {
		return f.Name == "word/document.xml"
	})
}

func findSearchInOdt(filename string, search string, caseInsensitive bool) {
	findSearchInFileInsideZip(filename, search, caseInsensitive, func(f *zip.File) bool {
		return f.Name == "content.xml"
	})
}

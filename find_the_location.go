package main

import (
	"image"
	"os"

	_ "image/png"

	"github.com/deluan/lookup"
	"github.com/kbinani/screenshot"
	"github.com/kpango/glg"
)

// Helper function to load an image from the filesystem
func loadImage(imgPath string) image.Image {
	imageFile, _ := os.Open(imgPath)
	defer imageFile.Close()
	img, _, _ := image.Decode(imageFile)
	return img
}

func get_position_from_img(img string) []lookup.GPoint {
	// Load full image
	bounds := screenshot.GetDisplayBounds(0)

	screenshot_img, err := screenshot.CaptureRect(bounds)
	if err != nil {
		glg.Error(err)
	}

	// Create a lookup for that image
	l := lookup.NewLookup(screenshot_img)

	// Load a template to search inside the image
	_, err = os.Stat(img)
	if err != nil {
		glg.Error("没有找到图片", err)
		os.Exit(11)
	}
	template := loadImage(img)

	// Find all occurrences of the template in the image
	pp, err := l.FindAll(template, 0.9)
	if err != nil {
		glg.Error("没有找到匹配", err)
	}
	return pp

}

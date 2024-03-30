# VBA-RapidOCR

[![Latest Version](https://img.shields.io/badge/Latest-v1.0.0-green.svg)]()
[![MIT License](https://img.shields.io/github/license/mashape/apistatus.svg)]()
[![Made with Love](https://img.shields.io/badge/Made%20with-%E2%9D%A4-red.svg?colorB=e31b23)]()

VBA Wrapper for RapidOCR: A library that empowers VBA users to extract text from images using the robust RapidOCR engine.

## Features

* Get Text From Image File.
* Easy to use.

## Install RapidOCR and ONNX Models

* Execute [Run-DownloadRapidOCRLibrary.bat](Run-DownloadRapidOCRLibrary.bat)

## Usage

##### Basic use:

```VB

Sub TestRapidOCRSimple()

    Dim ocr As New RapidOCR
    ocr.MsgBox ocr.ImageToText(ThisWorkbook.Path & "\images\Image1.png")
 
End Sub

```

<!-- ##### More examples [here.](/Examples) -->

## Release History

See [CHANGELOG.md](CHANGELOG.md)

<!-- ## Acknowledgments & Credits -->

## License

Usage is provided under the [MIT](https://choosealicense.com/licenses/mit/) License.

Copyright © 2024, [Danysys.](https://www.danysys.com)
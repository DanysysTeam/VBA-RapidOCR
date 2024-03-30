

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#download 7zip to temp path
$url7zip = "https://www.7-zip.org/a/7zr.exe"
$outputPath7zip = "$env:temp\7zr.exe"
#download if does not exist
if (!(Test-Path -Path $outputPath7zip)) {
    Invoke-WebRequest -Uri $url7zip -OutFile $outputPath7zip 
}

#donwload RapidOCR RapidOcrOnnx
$urlRapidOCR = "https://github.com/RapidAI/RapidOcrOnnx/releases/download/1.2.3/windows-clib-vs2022-mt.7z"
$outputPathRapidOCR = "$env:temp\windows-clib-vs2022-mt.7z"
#download if does not exist
if (!(Test-Path -Path $outputPathRapidOCR)) {
    Invoke-WebRequest -Uri $urlRapidOCR -OutFile $outputPathRapidOCR
}


& $outputPath7zip x $outputPathRapidOCR -o"$env:temp" -y

#create RapidOCR into script's path and create x86 and x64 subfolders
If (!(test-path -PathType container $PSScriptRoot\RapidOCR)) {
    New-Item -ItemType Directory -Path $PSScriptRoot\RapidOCR
}
If (!(test-path -PathType container $PSScriptRoot\RapidOCR\x86)) {
    New-Item -ItemType Directory -Path $PSScriptRoot\RapidOCR\x86
}
If (!(test-path -PathType container $PSScriptRoot\RapidOCR\x64)) {
    New-Item -ItemType Directory -Path $PSScriptRoot\RapidOCR\x64
}

#copy files ovwrite
Copy-Item -Path "$env:temp\windows-clib-vs2022-mt\win-CLIB-CPU-x64\bin\*" -Destination "$PSScriptRoot\RapidOCR\x64" -Recurse -Force
Copy-Item -Path "$env:temp\windows-clib-vs2022-mt\win-CLIB-CPU-Win32\bin\*" -Destination "$PSScriptRoot\RapidOCR\x86" -Recurse -Force


#create onnx-models folder
If (!(test-path -PathType container $PSScriptRoot\RapidOCR\onnx-models)) {
    New-Item -ItemType Directory -Path $PSScriptRoot\RapidOCR\onnx-models
}


# #download  onnx-models in temp and decompress to RapidOCR\onnx-models
$url_det = "https://github.com/Kazuhito00/PaddleOCR-ONNX-Sample/raw/507132aeab35c62336bf0153c21c40c6e9e4c68e/ppocr_onnx/model/det_model/ch_PP-OCRv3_det_infer.onnx"
$url_cls = "https://github.com/Kazuhito00/PaddleOCR-ONNX-Sample/raw/507132aeab35c62336bf0153c21c40c6e9e4c68e/ppocr_onnx/model/cls_model/ch_ppocr_mobile_v2.0_cls_infer.onnx"
$url_rec = "https://github.com/Kazuhito00/PaddleOCR-ONNX-Sample/raw/507132aeab35c62336bf0153c21c40c6e9e4c68e/ppocr_onnx/model/rec_model/ch_PP-OCRv3_rec_infer.onnx"
$url_keys = "https://raw.githubusercontent.com/PaddlePaddle/PaddleOCR/release/2.7/ppocr/utils/ppocr_keys_v1.txt"


#check if no exist into onnx-models
if (!(Test-Path -Path "$env:temp\ch_PP-OCRv3_det_infer.onnx")) {
    Invoke-WebRequest -Uri $url_det -OutFile "$env:temp\ch_PP-OCRv3_det_infer.onnx"
}

if (!(Test-Path -Path "$env:temp\ch_ppocr_mobile_v2.0_cls_infer.onnx")) {   
    Invoke-WebRequest -Uri $url_cls -OutFile "$env:temp\ch_ppocr_mobile_v2.0_cls_infer.onnx"
}

if (!(Test-Path -Path "$env:temp\ch_PP-OCRv3_rec_infer.onnx")) {
    Invoke-WebRequest -Uri $url_rec -OutFile "$env:temp\ch_PP-OCRv3_rec_infer.onnx"
}

if (!(Test-Path -Path "$env:temp\ppocr_keys_v1.txt")) {
    Invoke-WebRequest -Uri $url_keys -OutFile "$env:temp\ppocr_keys_v1.txt"
}   
   

if (!(Test-Path -Path "$PSScriptRoot\RapidOCR\onnx-models\ch_PP-OCRv3_det_infer.onnx")) {   
    Copy-Item -Path "$env:temp\ch_PP-OCRv3_det_infer.onnx" -Destination "$PSScriptRoot\RapidOCR\onnx-models\" -Force
}

if (!(Test-Path -Path "$PSScriptRoot\RapidOCR\onnx-models\ch_ppocr_mobile_v2.0_cls_infer.onnx")) {   
    Copy-Item -Path "$env:temp\ch_ppocr_mobile_v2.0_cls_infer.onnx" -Destination "$PSScriptRoot\RapidOCR\onnx-models\" -Force
}

if (!(Test-Path -Path "$PSScriptRoot\RapidOCR\onnx-models\ch_PP-OCRv3_rec_infer.onnx")) {   
    Copy-Item -Path "$env:temp\ch_PP-OCRv3_rec_infer.onnx" -Destination "$PSScriptRoot\RapidOCR\onnx-models\" -Force
}

if (!(Test-Path -Path "$PSScriptRoot\RapidOCR\onnx-models\ppocr_keys_v1.txt")) {   
    Copy-Item -Path "$env:temp\ppocr_keys_v1.txt" -Destination "$PSScriptRoot\RapidOCR\onnx-models\" -Force
}


Write-Host -NoNewLine 'Press any key to exit...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');





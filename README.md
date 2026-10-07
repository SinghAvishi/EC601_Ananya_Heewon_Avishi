# EC601_Ananya_Heewon_Avishi
This repository is for our EC601 project to place our documents and track our progress. 

Project Background:
This project builds an open-source compatibility testing tool for hospital AI medical imaging software procurement. Given DICOM images from a hospital's existing CT scanner and a vendor's publicly available FDA 510(k) clearance document, the tool automatically extracts the scanner parameters the vendor validated on, generates a quantitative fingerprint of the hospital's scanner, and compares the two, producing a structured compatibility report that flags specific mismatches or information gaps before a contract is signed. The tool addresses a documented gap in hospital AI governance: no standardized, independent methodology currently exists for assessing whether a procured AI tool will perform as validated on a specific institution's hardware, and vendor disclosure documents are inconsistently detailed, leaving procurement committees without the technical basis to make informed compatibility decisions.


Basic steps:
Creating a “compatibility software” to match new AI medical imaging software to existing scanners/hospital hardware
Inputs (from the hospital): sample DICOM images from scanners and vendor’s 510(k) document

Method: the tool takes information from the DICOM images to create a “fingerprint” which includes information about noise magnitude, texture, resolution, etc. 

From the 510(k) document, you might be able to pull out information about slice thickness, in addition to manufacturer they tested on and how they reconstructed the image

Then using the information from your scanner’s DICOM images and the vendor’s reconstruction algorithm information, you match them to see if they would be a good fit (if the vendor’s reconstructed image that they tested on matches the “quality” of the images from the hospital’s scanners). Then you can get a good idea on if whatever the medical imaging software is doing (like segmentation, etc.) can work on the images being produced by your hospital’s scanners already
Outputs (from our software): compatibility that’s honest about what it doesn’t know. There might be a lot that can’t be extracted from the 510(k) documents. This is really meant to be used as a guide for hospitals, so if there’s information they can’t extract using this software, they can just do it manually


Initial next steps:

510(k) document review:

Review some of these for some medical imaging software and see what kinds of information commonly show up (slice thickness, info about reconstruction, etc.).

Start building a tool to extract this information from the document


DICOM fingerprinting:
Get some sample images from a publicly available dataset and see if we can extract DICOM metadata
In addition to metadata, we’d need information on noise magnitude/other elements. Figure out how to extract these from images

Also finalize a list of components we’d need to focus on extracting from the image

Then put them together and create a user interface to have the users add their inputs, then they’d get a compatibility report as an output

# QuickScanEFR : OCR for Pulmonary Function Testing

This project offers an end-to-end pipeline for extracting textual content from PTF (Pulmonary Function Testing) scan reports, saved in PDF format, and then saving this extracted data as structured content. The pipeline is specifically optimized for PTF reports and ensures accuracy in data extraction.

## Features

-PDF to Image Conversion: Converts EFR PDF reports, especially those with embedded images or complex layouts, into high-resolution PNG images.
-Image Preprocessing: Enhances the quality of images from EFR reports to make them suitable for OCR. This includes resizing, thresholding, erosion, and cropping to focus on regions of interest.
-OCR Processing using Tesseract and Camelot Plubmer: Extracts text from preprocessed EFR images using Tesseract andthe Camelot library, which is integrated into the pipeline via a plumber. The extracted data is then structured for saving in Excel format.

## Optimization

The pipeline is optimized for performance using parallel processing wherever possible. This ensures faster processing of multiple EFR reports, especially on multi-core systems.

## Future Enhancements

-Adding support for multi-page EFR PDF reports.
-Enhancing OCR accuracy with fine-tuning
-GUI for easier user interaction with the pipeline.

## Contributing

If you'd like to contribute to this project, please fork the repository and submit a pull request.

## License

This project is licensed under the MIT License.


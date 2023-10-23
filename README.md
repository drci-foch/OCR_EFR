# QuickScanEFR: OCR for Pulmonary Function Testing

QuickScanEFR provides a dedicated pipeline for extracting textual content from PFT (Pulmonary Function Testing) scan reports saved in PDF format. Optimized specifically for PFT reports, it ensures precision in data extraction.

---

## 🚀 Features

- **PDF to Image Conversion**: 
  - Transform EFR PDF reports with embedded images or complex layouts into high-resolution PNG images.
  
- **Image Preprocessing**: 
  - Improve image quality from EFR reports for better OCR results.
  - Techniques include resizing, thresholding, erosion, and cropping to focus on regions of interest.
  
- **OCR Processing**: 
  - Utilize Tesseract and Camelot integrated through a plumber.
  - Extract and structure the data for saving in Excel format.

---

## ⚙️ Optimization

- Employs parallel processing for enhanced performance.
- Swift processing of multiple EFR reports, especially beneficial for multi-core systems.

---

## 🔮 Future Enhancements

- Support for multi-page EFR PDF reports.
- Boost OCR accuracy with further fine-tuning.
- Intuitive GUI for a user-friendly experience.

---

## 🤝 Contributing

- Interested in making a difference? Fork the repository and submit a pull request!

---

## 📜 License

- QuickScanEFR is proudly licensed under the MIT License.

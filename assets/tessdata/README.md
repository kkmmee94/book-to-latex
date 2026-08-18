# Bundled Arabic OCR data

`ara.traineddata` is the Arabic LSTM model from the official [`tesseract-ocr/tessdata_fast`](https://github.com/tesseract-ocr/tessdata_fast) repository.

- Upstream file: `ara.traineddata`
- Intended engine: Tesseract 4 or 5 LSTM
- SHA-256: `E3206D3DC87FD50C24A0FB9F01838615911D25168F4E64415244B67D2BB3E729`.
- Upstream licensing: Apache-2.0; see the upstream repository for notices and source/model details.

The traineddata file is not executable. The Book to LaTeX app supplies it through Tesseract's `--tessdata-dir` option when Arabic OCR is selected.

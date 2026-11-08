# Vendored libraries (pinned, served from this origin with Subresource Integrity)

| Library | Version | Licence | Source |
|---|---|---|---|
| Firebase JS SDK (compat bundles: app, auth, firestore) | 10.14.1 | Apache-2.0 | npm package `firebase` |
| pdf.js (`pdf.min.js`, `pdf.worker.min.js`) | 3.11.174 | Apache-2.0 | npm package `pdfjs-dist` |
| Tesseract.js (`tesseract/tesseract.min.js`, `tesseract/worker.min.js`) | 6.0.1 | Apache-2.0 | npm package `tesseract.js` |
| Tesseract core, LSTM builds (`tesseract/core/tesseract-core-lstm.wasm.js`, `tesseract-core-simd-lstm.wasm.js`) | 6.0.0 | Apache-2.0 | npm package `tesseract.js-core` |
| Arabic and English models (`tesseract/lang/ara.traineddata.gz`, `eng.traineddata.gz`, tessdata `4.0.0_best_int`) | 1.0.0 | Apache-2.0 | npm packages `@tesseract.js-data/ara`, `@tesseract.js-data/eng` |

To upgrade: copy the new files here, recompute `openssl dgst -sha384 -binary FILE | openssl base64 -A`, and update the `integrity` attributes in `index.html` and `attach.js`.

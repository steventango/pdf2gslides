# pdf2gslides
Converts PDFs to Google Slides!

## Why?
PDFs count against your Google Drive's storage quota, Google Slides do not.
University has given me endless number of space consuming PDFs.

## Installation
 1. Install [LibreOffice](https://www.libreoffice.org/download/download/).
 2. Install Python dependencies: `$ pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib numpy`.
 3. [Enable the Drive and Slides API](https://developers.google.com/drive/api/v3/quickstart/python).
 4. Download `credentials.json` and put into pdf2gslides folder.
 5. Run `$ python pdf2gslides.py`.
 6. Place PDF files to convert in the folder: `in\`.
 7. Run `$ python pdf2gslides.py`.
 8. Authenticate the program by signing into Google.
 9. Check your Google Drive for the GSlides converted from your PDFs.

## Dependencies
 - [google-api-python-client](https://github.com/googleapis/google-api-python-client) 1.12.3
 - [google-auth-httplib2](https://github.com/googleapis/google-auth-library-python-httplib2) 0.0.4
 - [google-auth-oauthlib](https://github.com/googleapis/google-auth-library-python-oauthlib) 0.4.1
 - [LibreOffice](https://www.libreoffice.org/) 7.0.1.2
 - [numpy](https://numpy.org/) 1.18.5

## Environment
 - Windows 10.0 Build 19041
 
 This program has only been tested on the above environment,
 there are no guarantees that it will work in other environments.

## TODO
 - Implement parallelization.

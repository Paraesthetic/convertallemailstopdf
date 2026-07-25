# Batch MSG and EML to PDF Converter

Convert a folder of Outlook MSG and standard EML email files into readable PDFs without using an online conversion service.

The script provides a folder picker, extracts core message details and writes one PDF for each supported email. It is useful for preparing correspondence bundles, preserving readable copies of exported email, reviewing records outside an email client and creating material for later indexing or document management.

## Supported input

* Outlook MSG files
* Standard EML files

For each message, the PDF includes:

* From
* To
* Sent date
* Subject
* Message body

The output is placed in a new folder named email pdf inside the selected input folder. Each PDF keeps the base filename of its source email.

## Requirements

* Python 3 with Tkinter
* extract-msg
* xhtml2pdf
* beautifulsoup4

The script attempts to install missing packages automatically. A dedicated virtual environment is recommended. To install them in advance:

    python -m pip install extract-msg xhtml2pdf beautifulsoup4

## Run

    python "Convert all emails in folder to PDF.py"

Select the folder containing the exported messages. The script processes supported files in that folder and prints the result of each conversion to the console.

No online document conversion service is used. Email content is read locally and PDFs are written locally. Internet access may still be used if the script needs to download missing Python packages.

## What is not included

* Attachments are not extracted or embedded.
* Messages in subfolders are not processed.
* Outlook specific formatting is not reproduced exactly.
* EML processing prefers the first plain text body part and ignores attached parts.
* Rich HTML only messages may have incomplete or blank body content.
* CC, BCC, importance, categories, delivery information and other metadata are not included.
* The script does not combine individual messages into a single PDF.

## Privacy and evidentiary limitations

The source message is rendered into a convenient reading copy. The resulting PDF is not a forensic preservation format and does not retain all original headers, metadata, attachments or cryptographic properties. Keep the original MSG or EML files when authenticity, completeness, legal hold, records compliance or later forensic examination may matter.

Email content is inserted into an HTML template before PDF rendering. Process messages from trusted sources and review the PDFs before opening or distributing them. Output filenames can also collide when different source formats share the same base filename.

## Licence

GNU General Public License version 3. See LICENSE for the complete terms.

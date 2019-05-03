# PGE186 Excel Tools

PGE186 Tools is an MS Excel add-in to provide a code-driven alternative to human cut-and-paste tasks needed to develop an MS Word-based report for a major California utility.  No proprietary utility data is included in this repo.

## Install

To use:
See Excel - PGE186 Tools.xlam Installation.docx or copy the .xlam file to your C:\Users\\*username*\AppData\Roaming\Microsoft\AddIns directory

*but really, none of these tools will be that helpful if you don't have proprietary utility data in your Excel sheet*

## Tools

The following tools are provided inside a custom ribbon tab, the (guidance|https://www.thespreadsheetguru.com/blog/step-by-step-instructions-create-first-excel-ribbon-vba-addin) for which was provided by Chris Newman, the Spreadsheet Guru.

### for the Excel sheet

#### Color Row by Change Type
Apply row background color based on cell value in first column.
#### Update Tower Numbers
Change text in #/# format  to 00#/00# format.
#### Add Change IDs
Add three columns at the end of the sheet to add a letter code based on cell value in first column, an incremented number based on the universe of similar letter codes, and a concatenation of those.

### for the Word summary
Each of the following tools creates a new tab containing a subset of the source data, suitable for cutting and pasting in the Word document.
#### Ratings Requested
#### Relay Request
#### Equipment Added
#### Equipment Retired


#### Source Docs Used
Summarizes a comments field to pull document references into a summary of documents used in the analysis
#### Source Docs Reference
Another new tab creator to contain a subset of the source data, suitable for cutting and pasting in the Word document.


**Exchange Defend: PDF** (xdpdf) is designed to quickly and transparently render inert potentially malicious parts of a PDF document traversing a Microsoft Exchange server. Whenever it changes a PDF document, it will advise the recipient of the email and keep a copy for administrative review if necessary.

xdpdf works in the following way:

1. It detects a PDF
2. It scans the PDF for potentially undesirable keywords
3. If a keyword is matched, xdpdf will:
  1. copy the PDF to an administrator-specified location
  2. overwrite the first two bytes of the keyword, preventing execution by the PDF reader
  3. notify the user, instructing them to contact their mail administrator if they require the original document

xdpdf is highly configurable.  It has multiple log levels, as well as allowing administrative configuration of log and quarantine directories.  In addition, the administrator can specify if xdpdf should believe the attachment extension or MIME-type supplied by Exchange, or perform speedy PDF detection on all incoming attachments.

System Requirements
===================
Microsoft Exchange Server 2010 SP1
.NET Framework 3.5

Note: xdpdf will work on Exchange Server 2010 without SP1 or Exchange Server 2007, however the DLL will likely need to be recompiled against the local versions of Microsoft.Exchange.Data.Common and Microsoft.Exchange.Data.Transport

Installation
============

1. Copy xdpdf.dll and xdpdf.config to a directory on the mail server
2. Set configuration options, creating the LogPath and QuarantinePath directories
3. In the Exchange Management Shell on the Exchange server, run 'Install-TransportAgent -Name xdpdf -AssemblyPath drive:\path\to\dll\xdpdf.dll> -TransportAgentFactory xdpdf.xdpdfFactory'
4. In the Exchange Management Shell on the Exchange server, run 'Enable-TransportAgent -Identity xdpdf'
5. In the Exchange Management Shell on the Exchange server, run 'Restart-Service MSExchangeTransport'
6. Monitor LogPath for log files, and QuarantinePath for quarantined files

Configuration Change
====================

1. In the Exchange Management Shell on the Exchange server, run 'Disable-TransportAgent -Identity xdpdf'
2. In the Exchange Management Shell on the Exchange server, run 'Restart-Service MSExchangeTransport'
3. Make changes to the configuration file
4. In the Exchange Management Shell on the Exchange server, run 'Enable-TransportAgent -Identity xdpdf'
5. In the Exchange Management Shell on the Exchange server, run 'Restart-Service MSExchangeTransport'

Uninstallation
==============

1. In the Exchange Management Shell on the Exchange server, run 'Disable-TransportAgent -Identity xdpdf'
2. In the Exchange Management Shell on the Exchange server, run 'Uninstall-TransportAgent -Identity xdpdf'
3. In the Exchange Management Shell on the Exchange server, run 'Restart-Service MSExchangeTransport'
4. Delete xdpdf files as desired

Configuration
=============
The configuration file contains the following settings:

Parameter | Options | Description
--- | --- | ---
Logging | True/False		| Enable or disable logging globally
Loglevel | 0 - 2		| Specify the log level:
	|			| 0) Log only malicious PDFs
	|			| 1) Log start, end and PDF processing details for each email
	|			| 2) Verbose; detailed logging (recommended for testing purposes only)
LogPath | drive:\path		| Path to the log directory.  It should already exist.
	|			| xdpdf will create one log file per day within this directory
QuarantinePath | drive:\path	| Path to the quarantine directory.  It should already exist.
	|			| xdpdf will create one subdirectory per email containing a potentially malicious PDF, and place the unmodified PDF within.
Keywords | string list		| Keywords should be of the form '/Keyword', noting the leading slash and case correctness according to the PDF spec
ScanAllAttachments | True/False	| If True, xdpdf will use its internal PDF file magic. If False, it will believe the file extension or MIME type supplied by Exchange.

Implementation Notes
====================
Given the potential volume of files being scanned in 'ScanAllAttachments' mode, it has been designed to be as fast and lightweight as possible.  It will read the first 1024 bytes of a file, and attempt to locate the bytes "%PDF" within - this is the extent of PDF file magic according to the spec.  During testing on a 2 x 2.6GHz core VM with 4gb of memory, Exchange could not measure the time it took to scan non-PDF files - it considered the entire process took place within the same millisecond.  If you run in ScanAllAttachments mode at log level 2, you will be able to see the metrics for your own server; of course you should do this in a testing environment.  On the other hand, you may decide to take the risk of misidentified PDF files getting through in exchange for not scanning all files.

xdpdf was deliberately designed to be as quick and thorough as possible.  This means that it makes no attempt to understand the structure or content of the file; rather it simply looks for the specified keywords (whether obfuscated or not), and renders inert those it finds.  It is entirely possible it will encounter the keyword strings within document content (either plain text or encoded images etc), especially for the shorter keywords such as '/AA' and '/JS'.  This is the main reason for keeping a copy of the original PDF - during testing, 21 PDFs from 7950 I tested (around 5.4gb worth) rendered incorrectly once they had been modified by xdpdf.  The burden would be on the end user to notify the mail administrator, and the mail administrator to use the log and original file to make a decision on the safety of the PDF, before passing it on to the end user.  Note that around a third of the testing PDFs were modified by xdpdf, and the vast majority displayed correctly following this.

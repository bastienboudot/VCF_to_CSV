**VCF to CSV Converter (Excel-Safe, Multi-Phone)**

**Overview**

This macOS script converts vCard (.vcf) contact files into a single Excel-compatible CSV file.

It supports:
	•	A single VCF file containing multiple contacts
	•	A folder containing multiple VCF files
	•	Contacts with multiple phone numbers
	•	Clean CSV output that opens correctly in Microsoft Excel

The script uses:
	•	AppleScript for user interaction
	•	Embedded Python for reliable vCard parsing and CSV generation

No third-party Python libraries are required.

⸻

**Features**
	•	Works with macOS Contacts exports
	•	Handles folded vCard lines
	•	Supports both EMAIL and E-MAIL fields
	•	Extracts multiple phone numbers into separate columns
	•	Prevents Excel column shifting
	•	Excludes the NOTE field by design
	•	Can be exported as a standalone macOS application

⸻

**Extracted Fields**

The script extracts the following fields from each contact:

CSV Column	Description
First Name	Contact first name
Last Name	Contact last name
E-mail	Primary email address
Organization	Company name
Department	Department / service
Job Title	Job title
Phone 1	First phone number
Phone 2	Second phone number (if any)
Phone 3	Additional phone numbers (as needed)

The number of Phone X columns is automatically adjusted based on the maximum number of phone numbers found across all contacts.

⸻

**How It Works**
	1.	The script asks whether you want to convert:
  	•	A single VCF file, or
  	•	A folder containing multiple VCF files
	2.	You select the source file or folder.
	3.	You choose the output CSV file name and location.
	4.	The script:
  	•	Parses all vCard entries
  	•	Collects all phone numbers
  	•	Builds a consistent CSV structure
  	•	Writes an Excel-safe CSV file
	5.	A confirmation dialog appears when the conversion is complete.

⸻

**Requirements**
	•	macOS (tested on Tahoe 26.2)
	•	Python 3 (included by default on recent macOS versions)
	•	Microsoft Excel (or any CSV-compatible spreadsheet software, not used to generate CSV file)

No additional installations are required.

⸻

**Usage**

Option 1 – Run from Script Editor
	1.	Open Script Editor
	2.	Paste the script
	3.	Click Run
	4.	Follow the on-screen prompts

Option 2 – Export as an Application
	1.	Open Script Editor
	2.	Choose File → Export
	3.	Select File Format: Application
	4.	Save the app
	5.	Double-click the app to run it

This allows the script to be used like a regular macOS application.

⸻

**Output Format Notes**
	•	The CSV uses comma separators
	•	All fields are properly quoted
	•	The file opens correctly in Excel without column misalignment
	•	Empty fields remain empty (no placeholder text)

This format is suitable for:
	•	Outlook mail merge
	•	CRM imports
	•	Email campaign tools

⸻

**Limitations**
	•	The NOTE field is intentionally ignored
	•	Only the first email address is kept if multiple are present
	•	Phone numbers are not labeled (mobile, home, work), only enumerated

⸻

**Customization**

The script can be easily adapted to:
	•	Add more vCard fields
	•	Rename CSV headers
	•	Export to multiple CSV files
	•	Apply Outlook-specific formatting rules

If needed, keep modifications inside the embedded Python section for best reliability.

⸻

**License**

Free to use and modify for personal or professional purposes.

⸻

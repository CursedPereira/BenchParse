# 🛡️ BenchParse - CIS Benchmark to Excel Converter

## Project Description

**BenchParse** is a Python script designed to automate the extraction of CIS benchmark controls from PDFs and convert them into an Excel sheet. It simplifies the process of analyzing security compliance data, making it easier for security professionals to review, audit, and implement CIS benchmarks efficiently.

> **BenchParse** streamlines the tedious task of manually parsing CIS benchmark documents, allowing security professionals to focus on implementation rather than data extraction. The tool enhances productivity by providing a structured and easily accessible format of CIS controls in Excel.

> Whether you're a security analyst, auditor, or compliance officer, **BenchParse** offers an efficient way to handle CIS benchmarks, ensuring a more seamless compliance assessment workflow.

---

## Features

- **Automated CIS Benchmark Extraction**: Parses PDF documents to extract relevant CIS benchmark controls.
- **Excel Output Generation**: Converts the extracted data into a structured Excel sheet.
- **Structured Data Extraction**:
  - Control Number
  - Control Title
  - Description
  - Rationale
  - Impact
  - Profile Applicability
  - Recommendations
- **Efficient Parsing with `pdfplumber`**: Uses `pdfplumber` for precise text extraction from PDFs.
- **Enhanced Readability with `rich`**: Implements colorful console outputs for better user experience.
- **Spinner Animation with `halo`**: Displays progress indicators while parsing.

---

## Installation

### Prerequisites

1. **Python**: Ensure Python 3.6 or higher is installed. Download from the [official Python website](https://www.python.org/downloads/).
2. **Required Python Libraries**: Install dependencies using:
   ```bash
   pip install -r requirements.txt
   ```
3. **CIS Benchmark PDFs**: Obtain CIS benchmark documents in PDF format for parsing.

### Installation Steps

1. Clone the repository:
   ```bash
   git clone https://github.com/Cursed271/BenchParse.git
   ```
2. Navigate to the project directory:
   ```bash
   cd BenchParse
   ```
3. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

---

## Usage

**Ensure that the CIS Benchmark PDF resides in the same folder**
1. Run the script:
   ```bash
   python BenchParse.py
   ```
2. Provide the path to the CIS benchmark PDF when prompted.
3. The script will extract CIS controls and save them in an Excel file.

---

## Screenshots
**Script Screenshot**
![Script](https://github.com/CursedPereira/BenchParse/blob/main/BenchParse.png)
**Output**
![Output](https://github.com/CursedPereira/BenchParse/blob/main/Output.png)


---

## Contribution

1. Fork the repository on GitHub.
2. Create a new branch for your modifications.
3. Implement your improvements or bug fixes.
4. Commit your changes and push them to your fork.
5. Submit a pull request.

---

## License  

This software (**BenchParse**) is licensed under a **Proprietary License**.  

- **Attribution Required**: You **must** credit **Cursed** in any public or private use of this software.  
- **Modification Restricted**: Modification is **only allowed under specific conditions** mentioned in the LICENSE file.  
- **No Reselling**: You are **not allowed** to sell or distribute this software without written permission.  
- **Modification Notification**: If you modify the software, you **must** notify me via a GitHub issue in this repository.  

For full terms, see the **[LICENSE](./LICENSE)** file.  

---

## About Me

Hey there, I'm **Cursed**, a Senior Cybersecurity Consultant specializing in **Cloud Security, Red Teaming, and Penetration Testing**. I develop security tools that simplify complex security workflows and empower security professionals.

Connect with me:
- **GitHub:** [@Cursed271](https://github.com/Cursed271)
- **LinkedIn:** [@Cursed271](https://www.linkedin.com/in/cursed271/)
- **Website:** [CursedSec](https://github.com/Cursed271)

---

## FAQ

**Q**: What input does BenchParse require?

**A**: It requires a **CIS benchmark PDF** as input to extract security controls and convert them into an Excel file.

**Q**: Can I use BenchParse for any PDF?

**A**: BenchParse is specifically designed for CIS benchmark PDFs. Parsing other PDFs may result in inaccurate data extraction.

**Q**: Does BenchParse require a license key?

**A**: No, but usage is restricted under the Proprietary Software License.

**Q**: How can I report issues or suggest improvements?

**A**: Open an issue on the GitHub repository or contact me via LinkedIn/GitHub.

---


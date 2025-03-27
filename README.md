# Better TMS API

Better TMS API is a collection of Python scripts that perform mass-actions on MercuryGate TMS with a focus on speed and the ability to handle large datasets. By using modern frameworks and multi-threading, these tools perform faster than the built-in MercuryGate APIs, making bulk updates and operations more efficient.

![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54) ![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)

## Table of Contents

- [Overview](#overview)
- [Script Overview](#scripts-overview)
- [Excel & CSV File Structure](#excel--csv-file-structure)
- [Setup & Installation](#setup--installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [Troubleshooting](#troubleshooting)
- [License](#license)

## Overview

This repository contains scripts designed to streamline mass-updates and actions on MercuryGate TMS. They are optimized for speed and high-volume operations, particularly when working with extensive Excel datasets. The provided scripts outperform the standard MercuryGate APIs by implementing multi-threading where applicable.

## Script Overview

| Script Name          | Purpose                                         | Dependencies                                                                        | Performance |
| -------------------- | ----------------------------------------------- | ----------------------------------------------------------------------------------- | ----------- |
| addPricesheet.py     | Creates and adds new pricesheet to load (EL/SO) | Source Excel file                                                               | ![Static Badge](https://img.shields.io/badge/multi--threaded-darkgreen)        |
| editPricesheet.py    | Updates existing pricesheet (EL/SO)             | Source Excel file                                                               | ![Static Badge](https://img.shields.io/badge/multi--threaded-darkgreen)        |
| editStatusMessage.py | Updates or adds status messages (EL/SO)         | Source Excel file, secondary Excel file with transport_order_id lookup | ![Static Badge](https://img.shields.io/badge/single--threaded-orange)      |


## Excel & CSV File Structure

### Expected Excel File Structure

Your Excel file should contain at least two worksheets:

1. **config**  
   This sheet contains key-value pairs in the following format:  
   - **Column A:** Configuration key  
   - **Column B:** Configuration value  
   
   **Required Variables:**  
   - `PRIMARY_SERVER`  
   - `AUTH_COOKIE`  
   - `ENTERPRISE_OID`  
   - `EVENT_SUFFIX`  
   - `STATUS_MESSAGE`  
   - `TRANSPORT_ORDER_SUFFIX`  
   - `SCAC`  

2. **lookup**  
   This sheet holds the data to be processed. The first row should include headers such as:  
   - `pri_ref` (the SO number)  
   - `OTM_COST` (the new cost value)  
   - `pricesheet_is` (the PriceSheet ID)  
   - `transport_id`  
   - `transport_order_id` (optional if a CSV mapping is provided)  

   The script processes rows starting from row 2 and stops at the first blank row in the `pri_ref` column.

### Optional CSV Mapping File

If available, the CSV mapping file should include a header row with at least the following columns:  
- `transport_id`  
- `transport_order_id`  

This file is used to provide additional mapping for transport IDs, enhancing the flexibility of the scripts.

## Setup & Installation

### Prerequisites

- Python 3.6 or higher
- The following Python packages:
  - `openpyxl`
  - `requests`
  - (Standard libraries: `csv`, `datetime`, `urllib`, `concurrent.futures`)

### Installation

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/yourusername/better-tms-api.git
   cd better-tms-api```

2. **Create a Virtual Environment:**

   ```python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate` instead ```

3. **Install Dependencies:**

   ```pip install -r requirements.txt```

## Usage

Before running any of the scripts, ensure that your Excel file (and optional CSV mapping file) is structured correctly as outlined above.

### Running a Script
For example, to run the Edit Pricesheets script:

   ```python EditPricesheets.py```

If your CSV mapping file is available, ensure that the script is pointed to the correct file path by updating the corresponding variable in the script or via command-line arguments (if implemented).

## Contributing

Contributions are welcome! Please follow these steps:
- Fork the repository.
- Create a new branch (`git checkout -b feature/your-feature`).
- Commit your changes.
- Push to the branch and create a Pull Request.

Ensure that your code follows the project style guidelines and that you have updated the README if your changes affect the usage or configuration.

## Troubleshooting

- **Excel Configuration Issues:**  
  Ensure that the `config` and `lookup` sheets exist and are properly formatted. Missing required keys in the `config` sheet will raise errors.

- **Mapping CSV Issues:**  
  Verify that your CSV file uses the correct headers (`transport_id`, `transport_order_id`). Incorrect formatting may lead to mapping errors.

- **HTTP Request Errors:**  
  If you encounter errors related to HTTP requests (e.g., priming the session or POST requests), double-check the `PRIMARY_SERVER` and `AUTH_COOKIE` values in your configuration.

For additional support or to report bugs, please open an issue on GitHub.

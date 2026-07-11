# 🛵 Swiggy Instamart PID Scraper

A Python-based Swiggy Instamart scraper that automates the extraction of product information using Swiggy Instamart Product IDs (PIDs).

The script reads Product IDs from an Excel or CSV file, visits the corresponding Swiggy Instamart product pages, extracts key product details, and exports the results into a formatted Excel file. It is designed for e-commerce research, pricing analysis, catalog management, and marketplace data collection.

---

## ✨ Features

* 📥 Read Product IDs (PIDs) from Excel or CSV
* 🔍 Scrape product information directly from Swiggy Instamart
* 🏷️ Extract Product Name
* 💰 Extract Selling Price
* 💵 Extract MRP
* 📉 Calculate Discount Percentage
* 📦 Extract Quantity / Size
* ✅ Check Product Availability
* 🏢 Extract Brand & Category
* ⭐ Extract Product Rating
* 📝 Extract Rating Count
* 📄 Extract Product Description
* 🖼️ Extract Product Image URL
* 📤 Export results to a formatted Excel file
* 📍 Automatically sets delivery location
* 🔄 Multi-layer data extraction with intelligent fallbacks

---

## 📂 Repository Structure

```text
swiggy-instamart-pid-scraper/
│
├── swiggy_instamart_scraper.py
├── Swiggy Input.xlsx
└── Swiggy Output.xlsx
```

---

## 📥 Input

Populate **Swiggy Input.xlsx** (or CSV) with the Product IDs (PIDs) or Swiggy Instamart product URLs.

Example:

| PID       |
| --------- |
| 123456789 |
| 987654321 |

---

## 📤 Output

After execution, the scraper generates **Swiggy Output.xlsx** containing:

* PID
* Product Name
* Brand
* Category
* Selling Price
* MRP
* Discount %
* Quantity / Size
* Availability
* Delivery Time
* Rating
* Rating Count
* Description
* Product Image URL
* Scraping Status
* Source URL

---

## 🛠 Tech Stack

* Python
* Selenium
* BeautifulSoup
* Pandas
* OpenPyXL
* WebDriver Manager

---

## 🚀 Getting Started

1. Clone this repository.

```bash
git clone https://github.com/AdityaSinha13/swiggy-instamart-pid-scraper.git
```

2. Open **swiggy_instamart_scraper.py** in your preferred Python IDE.

3. Install the required Python libraries.

4. Before running the script, update the configuration section:

   * Input File Path
   * Output File Path
   * Delivery Location / Pincode
   * Headless Mode (Optional)
   * Delay Between Requests (Optional)

5. Add your Product IDs to **Swiggy Input.xlsx**.

6. Run the script.

7. The extracted data will be saved in **Swiggy Output.xlsx**.

---

## ⚙ Configuration

The following settings can be customized directly inside the script:

* Input File Path
* Output File Path
* Delivery Location / Pincode
* Browser Mode (Headless / Visible)
* Delay Between Requests
* Fields to Extract

---

## 🚀 Smart Extraction Engine

The scraper uses multiple extraction methods to maximize accuracy:

1. **Next.js (`__NEXT_DATA__`) JSON**
2. **JSON-LD Structured Data**
3. **Open Graph & Meta Tags**
4. **Live DOM Extraction**

If one method fails, the scraper automatically falls back to the next available source.

---

## 💡 Use Cases

* Product Research
* Price Monitoring
* Competitor Analysis
* Marketplace Data Collection
* Catalog Management
* E-commerce Automation

---

## ⚠️ Disclaimer

This project is intended for educational and research purposes only.

Please comply with Swiggy's Terms of Service and avoid excessive or abusive scraping.

---

## 🚧 Future Enhancements

* Automatic dependency installation
* Resume interrupted scraping
* Proxy support
* CAPTCHA handling
* Parallel scraping
* Streamlit GUI
* Scheduled scraping
* Product image downloading

---

## 👨‍💻 Author

**Aditya Sinha**

If you found this project useful, consider giving it a ⭐ on GitHub!

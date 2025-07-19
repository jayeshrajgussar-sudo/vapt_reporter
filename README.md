# 🛡️ VAPT Reporting Framework

A Python-Flask based framework that automates the creation of industry-grade VAPT (Vulnerability Assessment and Penetration Testing) reports — reducing hours or days of manual work into just a few minutes.

---

## 📌 Background

In traditional VAPT workflows, security analysts manually extract data from tools like **Nessus**, **OpenVAS**, or **Burp Suite**, and paste findings into structured formats (PowerPoint/Excel). This process is:

- Time-consuming  
- Error-prone  
- Hard to scale in large assessments  

---

## 🛠️ Project Description

To solve these inefficiencies, this framework provides:

- 🚀 **Flask-based GUI** for ease of use  
- 📊 **Excel/CSV ingestion** of raw vulnerability data  
- 📎 **Auto-generated PPTX reports** (Executive + Technical format)  
- 🎨 **Support for logos, custom headings, and client branding**  
- 📁 **Project-wise output organization**  
- 📦 **Reusable templates across assessments**  

### 🔧 Tech Stack

- **Python**  
- **Flask** (GUI/Web Framework)  
- **Pandas** (Data Processing)  
- **python-pptx** (PowerPoint Automation)

---

## ⚙️ Features & Workflow

### ✅ Upload Raw Data
- Upload Excel/CSV files directly from GUI.

### ✅ Data Processing
- Automatically cleans, formats, and organizes input.

### ✅ Report Generation
- Outputs structured `.pptx` reports with client branding.

### ✅ Output Management
- Reports stored under organized folders per project & phase.

---

## 🚀 Getting Started

### 1. Install Dependencies

```bash
pip install -r requirements.txt
````

### 2. Run the App

```bash
python app.py
```

Access the GUI via browser at: `http://127.0.0.1:5000`

---

## 📄 How to Use

### Step-by-step:

1. Choose **report phase** (e.g., Phase 1 or 2)
2. Enter **client folder name**
3. Upload the **.xlsx file**
4. Click **Generate Report**

> ⚠️ Ensure the Excel file has exact column names as expected in the template.

---

## 🖼️ Custom Branding

To add your logo and customize the report:

### Steps:

1. Go to the `add on/` folder

2. Place:

   * Your PPTX template
   * `logo.png` (your logo)

3. Run the customization script:

```bash
python add_ons.py
```

4. Provide the following:

   * PPTX filename
   * Logo filename
   * Client Name
   * Service Provider Name

---

## 📊 Sample Comparison

| Traditional         | This Framework                |
| ------------------- | ----------------------------- |
| Manual copy-paste   | Fully automated               |
| Hours of formatting | Done in minutes               |
| Inconsistent output | Structured, professional PPTX |
| Static templates    | Dynamic branding & logos      |
| No version control  | Organized per project/phase   |

---

## 📁 Deliverables

* ✅ Full source code
* ✅ Flask GUI
* ✅ Branded PPTX Templates
* ✅ Sample reports

---

## 🚀 Future Scope

* 📡 Real-time API integration (OpenVAS, Nessus)
* 📥 PDF/Excel export options
* 📊 Interactive dashboards
* 👥 Multi-user report access/versioning

---

## 📄 License

MIT License (or your preferred license)

---

## 🤝 Contributions

Pull requests and feedback are welcome! Let's make VAPT reporting seamless.

---

## 👨‍💻 Maintainer

**Jayesh Raj Gussar**
Cybersecurity & Python Automation Enthusiast

```

---

Let me know if you'd like a `LICENSE`, `requirements.txt`, or `.gitignore` file to go along with this. I can generate those too.
```

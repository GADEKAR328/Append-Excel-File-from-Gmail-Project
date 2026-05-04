# 📊 Gmail to Power Query – Append Excel Files Automatically

## 🚀 Project Overview
This project demonstrates how to **automatically extract Excel file attachments from Gmail** and **append them into a single dataset using Power Query**.

Since Power Query does not directly support Gmail integration, we use **Google Apps Script as a bridge API** to fetch email attachments and load them into Excel.

---

## 🎯 Objective
- Fetch Excel attachments from Gmail
- Convert and process file data in Power Query
- Append multiple files into a single dataset
- Automate data consolidation process

---

## 🛠️ Tools & Technologies Used
- Microsoft Excel
- Power Query
- Google Apps Script
- Gmail (IMAP enabled)
- Web API Integration

---

## 🔄 Workflow Architecture

Gmail → Google Apps Script (API) → Power Query → Data Transformation → Final Dataset

---

## ⚙️ Step-by-Step Implementation

### 🔹 Step 1: Enable IMAP in Gmail
- Open Gmail
- Go to **Settings → See all settings**
- Navigate to **Forwarding and POP/IMAP**
- Enable **IMAP access**
- Save changes

---

### 🔹 Step 2: Create Google Apps Script API
- Go to: https://script.google.com/home
- Click **New Project**
- Paste the Gmail attachment extraction script
- Click **Deploy → New Deployment**
- Select type: **Web App**
- Set access to: **Anyone**
- Authorize permissions
- Copy the **Web App URL**

---

### 🔹 Step 3: Connect Power Query to API
- Open Excel
- Go to **Data → Get Data → From Web**
- Paste the **Web App URL**
- Load the data into Power Query

---

### 🔹 Step 4: Convert Base64 to Binary
In Power Query:

- Go to **Add Column → Custom Column**
- Use formula:

```powerquery
Binary.FromText([Binary], BinaryEncoding.Base64)
```

- This converts encoded data into readable file format

---

### 🔹 Step 5: Extract Attachments

Depending on file type:

#### 📄 Excel Files (.xlsx, .csv)
```powerquery
Excel.Workbook([Binary], null, true)
```

#### 📄 Text Files (.txt, .csv)
```powerquery
Text.FromBinary([Binary])
```

#### 📄 JSON Files
```powerquery
Json.Document([Binary])
```

---

### 🔹 Step 6: Append Data
- Expand extracted tables
- Clean & transform data
- Append multiple files into one final dataset

---

## 📈 Key Features
- ✅ Automated email attachment extraction
- ✅ Dynamic file handling
- ✅ Supports multiple file formats
- ✅ Real-time data refresh capability
- ✅ No manual file downloads required

---

## 📂 Use Cases
- Sales report consolidation
- Daily automated reporting
- Finance & accounting data merging
- HR data collection
- Business intelligence workflows

---

## 📸 Project Preview
<p align="center">
  <img src="https://raw.githubusercontent.com/GADEKAR328/Append-Excel-File-from-Gmail-Project/f9d55f972b7118463ed6bdc78615ee03da228197/Append%20Gmail%20To%20Excel%20.jpg" width="900"/>
</p>

---

## 🔗 Important Links
- Google Apps Script: https://script.google.com
- Project Guide Script: https://script.google.com/d/1hF3JEE5uHwV4eDAQogA4NFI7y98Nv3lXhqTeM0GxfzKFNYniT6k8EGnB/edit?usp=sharing
- GitHub Repository: https://github.com/GADEKAR328

---

## 🙏 Acknowledgment
Special thanks to **Mr. Satish Dhawale** (CEO & Founder of SkillCourse) for guidance and support.

---

## 📢 Connect With Me
- LinkedIn: https://www.linkedin.com/in/yogesh-gadekar-a1231b189
- GitHub: https://github.com/GADEKAR328

---

## ⭐ If you found this project useful, don't forget to star the repository!

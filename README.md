# 🎓 Certificate Automation System

This project automates the process of sending digital certificates to students.  
You simply upload an **Excel sheet** containing student details (like name, email, and event name), and the system generates and emails personalized certificates to each student.

---

## 🚀 Features

- 📂 Upload Excel file containing student details  
- 🧾 Auto-generates certificates for each student  
- ✉️ Sends certificates as email attachments  
- ⚙️ Configurable email templates and SMTP settings  
- 🧠 Error handling and logs for failed deliveries  

---

## 🧰 Tech Stack

- **Python 3.x**  
- **pandas** — for reading and processing Excel files  
- **reportlab / PIL** — for generating certificates  
- **smtplib / email.mime** — for sending emails  
- **jinja2 (optional)** — for templating certificate text  

---

## 📋 Prerequisites

Make sure you have:
- Python 3.8+ installed  
- SMTP credentials (e.g., Gmail app password)  
- Excel sheet in `.xlsx` format with the following columns:
  | Column | Description |
  |:--------|:------------|
  | Name | Student's full name |
  | Email | Student's email address |
  | Course | Course or event name |

---

## ⚙️ Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/muskan9191/certifier.git
   cd certifier
   ```

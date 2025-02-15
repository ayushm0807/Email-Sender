# Email-Sender

### **📧 AI-Powered Automated Email Sender**

This project is a **fully AI-generated** Python program designed to automate **email sending with attachments** using data from an **Excel sheet**.  
The goal of this project was to **enhance my prompt engineering skills** by learning how to instruct AI to generate **complex, structured code**.

---

## **🚀 Features**
✅ **Reads recipient details** from an Excel file  
✅ **Supports PDF attachments** per recipient  
✅ **Uses environment variables** for security (no hardcoded credentials)  
✅ **Prevents duplicate emails** by tracking sent status  
✅ **Skips empty rows** to optimize processing  
✅ **Preserves Excel formatting** while updating values  


---

## **🛠️ Setup & Usage**

### **1️⃣ Install Dependencies**
First, install the required Python libraries:
```bash
pip install smtplib openpyxl python-dotenv
```

---

### **2️⃣ Set Up Environment Variables**
Create a **`.env`** file in the project directory and add:
```
EMAIL_SENDER=your-email@gmail.com
EMAIL_PASSWORD=your-app-password
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
```
🔹 **Never use your real password!** Use a **Google App Password** for security.

---

### **3️⃣ Prepare Your Excel File**
The script reads emails from an Excel file (`email_list.xlsx`) with this format:

| S.No. | Name  | Mail               | Company  | Subject | Description | PDF Path            | Send | Sent |
|-------|-------|--------------------|----------|---------|-------------|----------------------|------|------|
| 1     | John  | john@example.com    | ABC Ltd  | Offer   | Welcome!    | `/path/to/file.pdf` | Yes  | Yes  |
| 2     | Sarah | sarah@example.com   | XYZ Inc  | Update  | News        | `/path/to/file.pdf` | Yes  | No   |
| 3     | Mike  | mike@example.com    | PQR Ltd  | Event   | Invitation  |                      | No   | No   |

🔹 **Send Column (H):** Set `"Yes"` to send an email  
🔹 **Sent Column (I):** Automatically updates to `"Yes"` after successful sending  

---

### **4️⃣ Run the Script**
Execute the script to send emails:
```bash
python email_sender.py
```

---


## **🌟 Why This Project?**
This project is a **learning experiment** where I explored how to:  
✔ Use AI to generate **functional, real-world code**  
✔ Apply **prompt engineering** to refine AI-generated responses  
✔ Understand **code quality, security, and workflow** in AI-assisted programming  

---

**Made with AI-powered creativity 🚀**  

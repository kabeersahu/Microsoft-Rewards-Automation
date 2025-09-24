# **Automated Microsoft Rewards Script**

### **🚀 Project Goal: The Smart Way to Earn Rewards**

This Python script is designed to automate the process of earning Microsoft Rewards points by leveraging the fact that a single PC can manage multiple accounts through separate Chrome user profiles. This project focuses on performing realistic web searches to mimic human behavior, thereby avoiding detection and potential account bans. It's a robust solution that showcases skills in web automation, data management, and creating reliable tools that handle real-world challenges.

### **🧠 How It Works: A Realistic Automation Flow**

Here’s a step-by-step visual guide to the script's process.

**1\. Set Up & Clean Up**

* **Launch**: You start the zhel.py script.  
* **Process Check**: It first automatically finds and shuts down any leftover Chrome windows that might be running. This keeps things clean and error-free.

**2\. Data & Accounts**

* **Data Source**: The script pulls a list of search queries from your Excel file (Book1.xlsx).  
* **Multi-Profile Loop**: It cycles through each of your designated Chrome user profiles, one at a time.

**3\. Human-Like Execution**

* **Navigate**: For each profile, the script opens a new browser window.  
* **Realistic Search**: It performs searches using these techniques:  
  * **Time Lag**: It adds a random delay between each search to simulate a person's natural pause.  
  * **Scroll Simulation**: It scrolls down the page to mimic a user looking at the results.

**4\. The End**

* **Cycle Complete**: After finishing all the searches for one account, it closes the chrome profile 1 and movees to chrome file 2 (i.e.) moves on to the next one.

### **🌟 Core Strengths**

| Feature | Why It's Impressive |
| :---- | :---- |
| ✅ **Multi-Account Support** | Manages workflows across multiple user sessions. |
| ✅ **Realistic User Simulation** | Use of anti-bot measures, robust and almost undetectable  |
| ✅ **Data-Driven & Scalable** | Manage search queries separately from the code for easy scaling and updates. |
| ✅ **Automated Process Management** | Self-sufficient tool that doesn't rely on manual intervention. |
| ✅ **Strong Error Handling** | The script is designed to be resilient and won't crash when faced with unexpected web page layouts. |

### **⚙️ Getting Started**

1. **Clone the Repository:** Download the project files.  
2. **Install Dependencies:** Run this command in your terminal to get everything you need:  
   pip install selenium openpyxl psutil

3. **Update Paths:** In MS_MONEY.py, you'll need to set the file paths for your ChromeDriver and Chrome user profiles.  
4. **Add Your Data:** Put your search terms in the Book1.xlsx file.  
5. **Run the Script:** Open your terminal in the project folder and run:  
   python zhel.py

### **📂 Project Files**

* MS_MONEY.py: The core automation script.  
* Book1.xlsx: The Excel file with your search queries.  
* New Microsoft Word Document.docx: The document containing the email and password lists for the different accounts.

# **Automated Microsoft Rewards Script**

### **Project Goal**

This Python script is designed to automate the process of earning Microsoft Rewards points by leveraging the fact that a single PC can manage multiple accounts through separate Chrome user profiles. This project focuses on performing realistic web searches to mimic human behavior, thereby avoiding detection and potential account bans. It's a robust solution that showcases skills in web automation, data management, and creating reliable tools that handle real-world challenges.

### **How It Works**

The script follows a simple, step-by-step process.

1. **Start the Script:** You run the zhel.py file on your computer.  
2. **Clean Up:** The script first checks for and closes any extra Chrome browser windows that might be running. This is important for a smooth process\!  
3. **Get Your Data:** It reads a list of search words from your Excel file (Book1.xlsx).  
4. **Loop Through Accounts:** It goes through each of your Chrome user profiles one at a time. Each profile acts like a different person's browser session.  
5. **Perform Realistic Searches:** For each profile, it opens a Chrome browser and performs searches with a human-like touch:  
   * It picks a random search word from your Excel file.  
   * It adds a random pause (time lag) between each search, so it doesn't look like a super-fast robot.  
   * It scrolls down the page to simulate a person actually looking at the search results.  
6. **Done\!** After all the searches are done for one account, it closes the browser and moves to the next one.

### **Project Strengths**

✅ **Multi-Account Automation:** This script can use many different Chrome accounts one after the other. It's perfect for managing multiple Microsoft Rewards accounts in a single run.

✅ **Realistic User Simulation:** This project is built to act like a real person. By using random delays and automatic scrolling, it avoids raising red flags and looks more authentic.

✅ **Data-Driven Searches:** All the search words are in an Excel file. This means you can easily change what the script searches for without ever touching the code.

✅ **Automated Process Management:** The script is smart. It finds and closes any extra Chrome browser windows that might get stuck. This keeps your computer running smoothly and prevents errors.

✅ **Strong Error Handling:** The script is built to handle problems and keep going. It won't crash if it can't find a button or a link on a website.

### **How to Use It**

1. **Get the Files:** Download or clone this project from GitHub.  
2. **Install the Tools:** Open your command line and run this command:  
   pip install selenium openpyxl psutil

3. **Set Your Paths:** In the zhel.py file, you need to update the file paths for your ChromeDriver, Excel file, and Chrome profiles.  
4. **Add Your Search Words:** Make sure your Excel file (Book1.xlsx) has a column with the search words you want the script to use.  
5. **Run the Script:** Open your command line in the project folder and type:  
   python zhel.py

### **Project Files**

* zhel.py: The main script that does all the work.  
* Book1.xlsx: The Excel file where your search words are stored.  
* New Microsoft Word Document.docx: This document contains the email and password list for the different accounts.

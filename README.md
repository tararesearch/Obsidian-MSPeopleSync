# 📘 How to Use This Plugin (Microsoft People Sync)

This plugin helps you sync your Microsoft Graph `/me/contacts` into clean, minimal Obsidian notes — one note per person — using a customizable template.

## 🧩 1. Create a People Folder

Create a folder:
People/

## ⚙️ 2. Plugin Settings
- Access Token
- People Folder
- File Prefix
- Template customization

![[msconfig.png]]

## 🚀 3. Sync Contacts
Use command palette:
Microsoft People Sync: Sync contacts from Microsoft Graph


## 📂 4. Check Generated Files
People/@Name.md

## 🔗 5. Use in Notes
Using in note [[@  <- Will show People list from people folder
Using in note ![[@ <- will show embedded people information

Have a nice day.


Example Template:
#### {{displayName}} • 🧑‍💼 {{jobTitle}}
📧 {{primaryEmail}}  
📱 {{mobilePhone}}  
🏢 {{department}} • {{companyName}} • {{officeLocation}}  
👔 {{title}}  
☎️ {{businessPhones}}


This plug-in can working with other plug in
obsidian://show-plugin?id=at-people
obsidian://show-plugin?id=obsidian-completr


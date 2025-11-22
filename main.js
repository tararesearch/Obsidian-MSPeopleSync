
const { Plugin, PluginSettingTab, Setting, TFile, Notice } = require('obsidian');

const DEFAULT_TEMPLATE = [
  "#### {{displayName}} • 🧑‍💼 {{jobTitle}}",
  "",
  "📧 {{primaryEmail}}  ",
  "📱 {{mobilePhone}}  ",
  "🏢 {{department}} • {{companyName}} • {{officeLocation}}  ",
  "👔 {{title}}  ",
  "☎️ {{businessPhones}}  ",
].join("\n").trim();

const DEFAULT_SETTINGS = {
  accessToken: "",
  peopleFolder: "People",
  template: DEFAULT_TEMPLATE,
  filePrefix: "@"
};

class PeopleSyncPlugin extends Plugin {
  async onload() {
    console.log("Loading Microsoft People Sync plugin");
    await this.loadSettings();

    this.addCommand({
      id: "ms-people-sync",
      name: "Sync contacts from Microsoft Graph",
      callback: () => this.syncContacts()
    });

    this.addSettingTab(new PeopleSyncSettingTab(this.app, this));
  }

  onunload() {
    console.log("Unloading Microsoft People Sync plugin");
  }

  async loadSettings() {
    const loaded = await this.loadData();
    this.settings = Object.assign({}, DEFAULT_SETTINGS, loaded || {});
  }

  async saveSettings() {
    await this.saveData(this.settings);
  }

  async syncContacts() {
    const token = (this.settings.accessToken || "").trim();
    if (!token) {
      new Notice("People Sync: กรุณาใส่ Access Token ใน Settings ก่อน", 5000);
      return;
    }

    new Notice("People Sync: เริ่มดึงข้อมูลจาก Microsoft Graph...", 3000);

    try {
      const contacts = await this.fetchAllContacts(token);
      const result = await this.writeContactFiles(contacts);
      new Notice(`People Sync: สร้าง/อัปเดต ${result.written} ไฟล์ (ข้าม ${result.skipped})`, 5000);
    } catch (err) {
      console.error("People Sync error:", err);
      new Notice("People Sync error: " + (err && err.message ? err.message : err), 8000);
    }
  }

  async fetchAllContacts(accessToken) {
    let url = "https://graph.microsoft.com/v1.0/me/contacts" +
              "?$top=50" +
              "&$select=displayName,title,jobTitle,companyName,department,officeLocation,mobilePhone,businessPhones,emailAddresses";

    const all = [];
    let page = 1;

    while (url) {
      console.log(`PeopleSync: Fetching page ${page}: ${url}`);
      const res = await fetch(url, {
        headers: {
          "Authorization": "Bearer " + accessToken,
          "Content-Type": "application/json"
        }
      });

      if (!res.ok) {
        const text = await res.text();
        throw new Error(`Graph ${res.status}: ${text}`);
      }

      const data = await res.json();
      if (Array.isArray(data.value)) {
        all.push(...data.value);
        console.log(`   ➕ got ${data.value.length} contacts (total: ${all.length})`);
      }

      url = data["@odata.nextLink"] || null;
      page++;
    }

    console.log(`PeopleSync: Loaded ${all.length} contacts total`);
    return all;
  }

  renderTemplate(tpl, data) {
    let out = tpl.replace(/{{\s*([\w]+)\s*}}/g, (match, key) => {
      return data[key] || "";
    });

    const lines = out.split("\n");
    const cleaned = lines.filter(line => {
      const stripped = line.replace(/[*`_]/g, "").trim();
      return stripped.length > 0;
    });

    return cleaned.join("\n").trim() + "\n";
  }

  contactToData(c) {
    const safe = v => (v || "").toString().trim();

    const displayName = safe(c.displayName);
    const title = safe(c.title);
    const jobTitle = safe(c.jobTitle);
    const companyName = safe(c.companyName);
    const department = safe(c.department);
    const officeLocation = safe(c.officeLocation);
    const mobilePhone = safe(c.mobilePhone);

    const businessPhonesArr = Array.isArray(c.businessPhones) ? c.businessPhones : [];
    const businessPhones = businessPhonesArr.map(safe).filter(Boolean);

    const emailsArr = Array.isArray(c.emailAddresses) ? c.emailAddresses : [];
    const primaryEmail = emailsArr[0] && emailsArr[0].address
      ? safe(emailsArr[0].address)
      : "";

    const hasAny = [
      displayName,
      jobTitle,
      companyName,
      department,
      officeLocation,
      mobilePhone,
      primaryEmail,
      ...businessPhones
    ].some(v => v && v !== "");

    if (!hasAny) return null;

    return {
      displayName,
      title,
      jobTitle,
      companyName,
      department,
      officeLocation,
      mobilePhone,
      businessPhones: businessPhones.join(", "),
      primaryEmail
    };
  }

  async writeContactFiles(contacts) {
    const vault = this.app.vault;
    const folder = (this.settings.peopleFolder || "People").trim();
    const prefix = this.settings.filePrefix || "@";

    const folderPath = folder.endsWith("/") ? folder.slice(0, -1) : folder;
    try {
      await vault.createFolder(folderPath);
    } catch (e) {
      // ignore exists
    }

    let written = 0;
    let skipped = 0;

    for (const c of contacts) {
      const data = this.contactToData(c);
      if (!data) {
        skipped++;
        continue;
      }

      const baseName = data.displayName || data.primaryEmail || "Unknown";
      const fileName = prefix + this.sanitizeFileName(baseName) + ".md";
      const filePath = folderPath + "/" + fileName;

      const content = this.renderTemplate(this.settings.template, data);

      const existing = vault.getAbstractFileByPath(filePath);
      if (existing instanceof TFile) {
        await vault.modify(existing, content);
      } else {
        await vault.create(filePath, content);
      }

      written++;
      console.log("PeopleSync: wrote", filePath);
    }

    return { written, skipped };
  }

  sanitizeFileName(name) {
    return (name || "Unknown").replace(/[\\\/:*?"<>|]/g, "").trim();
  }
}

class PeopleSyncSettingTab extends PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display() {
    const { containerEl } = this;
    containerEl.empty();

    containerEl.createEl("h2", { text: "Microsoft People Sync Settings" });

    new Setting(containerEl)
      .setName("Access Token")
      .setDesc("Access Token สำหรับ Microsoft Graph (Contacts.Read). ระวัง: ข้อมูลสำคัญ ห้ามแชร์.")
      .addText(text =>
        text
          .setPlaceholder("eyJ0eXAiOiJKV1QiLCJhbGciOi...")
          .setValue(this.plugin.settings.accessToken)
          .onChange(async (value) => {
            this.plugin.settings.accessToken = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("People folder")
      .setDesc("โฟลเดอร์ปลายทางใน Vault (เช่น People)")
      .addText(text =>
        text
          .setPlaceholder("People")
          .setValue(this.plugin.settings.peopleFolder)
          .onChange(async (value) => {
            this.plugin.settings.peopleFolder = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("File name prefix")
      .setDesc("Prefix (เช่น @ Using for autocomplete)")
      .addText(text =>
        text
          .setPlaceholder("@")
          .setValue(this.plugin.settings.filePrefix)
          .onChange(async (value) => {
            this.plugin.settings.filePrefix = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Template")
      .setDesc("ใช้ {{field}} เพื่อแทนค่าจาก contact. ค่า default .")
      .addTextArea(area => {
        area
          .setValue(this.plugin.settings.template)
          .onChange(async (value) => {
            this.plugin.settings.template = value;
            await this.plugin.saveSettings();
          });
        area.inputEl.rows = 10;
        area.inputEl.cols = 50;
      });

    new Setting(containerEl)
      .setName("Reset template")
      .setDesc("Reset template to standard")
      .addButton(btn =>
        btn
          .setButtonText("Reset")
          .onClick(async () => {
            this.plugin.settings.template = DEFAULT_TEMPLATE;
            await this.plugin.saveSettings();
            this.display();
            new Notice("People Sync: Reset template done");
          })
      );

    containerEl.createEl("h3", { text: "Available fields (for {{field}})" });
    containerEl.createEl("p", {
      text: "displayName, title, jobTitle, companyName, department, officeLocation, mobilePhone, businessPhones, primaryEmail"
    });
  }
}

module.exports = PeopleSyncPlugin;

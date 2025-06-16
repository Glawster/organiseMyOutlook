import win32com.client

import tkinter as tk
from tkinter import ttk
from datetime import datetime

import re
import os
import threading
import pythoncom

from setupLogging import setupLogging

# Setup logging
logger = setupLogging("organiseMyOutlook")

class OrganiseMyOutlook:

    def ensureRequiredSubfolders(self, folder):
        required = ["Inbox", "Sent Items"]
        existing = {f.Name for f in folder.Folders}
        for name in required:
            if name not in existing:
                try:
                    folder.Folders.Add(name)
                    logger.info(f"...created missing folder '{name}' in '{folder.Name}'")
                except Exception as e:
                    logger.error(f"could not create folder '{name}' in '{folder.Name}': {e}")

  
    def __init__(self, root):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.root = root
        self.root.title("Organise My Outlook")

        self.sourceFolder = None
        self.destinationFolder = None

        self.buildForm()

    def extractAccountFromPstName(self, folderName):
      
        # Format 1: andyw@glawster.com (2025)
        email_match = re.match(r"(.+@[\w\.]+) \(\d{4}\)", folderName)
        if email_match:
            return email_match.group(1).lower()

        # Format 2: Andy @ Glawster (2025)
        name_match = re.match(r"(.+?) @ .+? \(\d{4}\)", folderName)
        if name_match:
            return name_match.group(1).strip().lower()

        # Format 3: Just a name, no account or year
        return folderName.strip().lower()

    def onSourceSelected(self, event):
        self.chkPSTCreate.config(state="normal")
        self.updateDestinationList()

    def onCreateMissingToggle(self):
        if self.createMissingVar.get() and self.sourceCombo.get():
            self.statusLabel.config(text=f"creating missing PST files...")
            accountName = self.extractAccountFromPstName(self.sourceCombo.get())
            sourceFolder = self.outlook.Folders[self.sourceCombo.get()]
            self.checkAndCreateMissingPSTs(accountName, sourceFolder)

    def buildForm(self):
        logger.info("Starting...")
        self.pstFolderPath = self.getDefaultPstFolder()
        logger.info (f"...default PST folder path: {self.pstFolderPath}")
        allFolders = [f.Name for f in self.outlook.Folders]
        logger.info("...found these folders")
        for folder in sorted(allFolders):
            logger.info(f"...{folder}")

        self.filterVar = tk.BooleanVar(value=True)
        self.dryRunVar = tk.BooleanVar(value=True)
        self.createMissingVar = tk.BooleanVar(value=False)
        
        sortedFolders = sorted([f.Name for f in self.outlook.Folders])

        ttk.Label(self.root, text="Select Source PST/Account:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.sourceCombo = ttk.Combobox(self.root, width=50, state="readonly", values=sortedFolders)
        self.sourceCombo.grid(row=0, column=1, padx=10, pady=5)
        self.sourceCombo.bind("<<ComboboxSelected>>", self.onSourceSelected)

        ttk.Label(self.root, text="Select Destination PST (must include year):").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.destinationCombo = ttk.Combobox(self.root, width=50, state="readonly", values=sortedFolders)
        self.destinationCombo.grid(row=1, column=1, padx=10, pady=5)
        self.destinationCombo.bind("<<ComboboxSelected>>", self.onDestinationSelected)

        ttk.Label(self.root, text="Override Year (optional):").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.overrideYearEntry = ttk.Entry(self.root, width=10)
        self.overrideYearEntry.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        self.chkFilter = ttk.Checkbutton(self.root, text="Filter destinations by source account", variable=self.filterVar, command=self.updateDestinationList)
        self.chkFilter.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        self.chkDryRun = ttk.Checkbutton(self.root, text="Dry Run (don't actually move emails)", variable=self.dryRunVar)
        self.chkDryRun.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        
        self.statusLabel = ttk.Label(self.root, text="", foreground="blue")
        self.statusLabel.grid(row=6, column=0, columnspan=2, padx=10, pady=5, sticky="we")

        self.progressBar = ttk.Progressbar(self.root, mode='indeterminate')
        self.progressBar.grid(row=6, column=0, columnspan=2, pady=10, padx=10, sticky="we")
        self.progressBar.grid_remove()  # hide by default
        
        # right-side action buttons
        frmActions = ttk.Frame(self.root)
        frmActions.grid(row=3, column=1, padx=10, pady=5, sticky="e")

        btnCreatePST = ttk.Button(frmActions, text="Create Missing PSTs", command=self.onCreateMissingToggle)
        btnCreatePST.pack(side="right", padx=5)

        btnScan = ttk.Button(frmActions, text="Scan for Move Candidates", command=self.startScanInThread)
        btnScan.pack(side="right", padx=5)

        # button frame
        frmButton = ttk.Frame(self.root)
        frmButton.grid(row=7, column=0, columnspan=2, pady=10)

        btnScan = ttk.Button(frmButton, text="Scan for Move Candidates", command=self.startScanInThread)
        btnScan.pack(side="left", padx=10)
 
        btnRun = ttk.Button(frmButton, text="Move Emails", command=self.startMoveInThread)
        btnRun.pack(side="left", padx=10)

        btnExit = ttk.Button(frmButton, text="Exit", command=self.root.quit)
        btnExit.pack(side="left", padx=10)

    def updateDestinationList(self):
      
        def normalize(name):
            return name.replace(" ", "").lower()

        source = self.sourceCombo.get()
        normalizedSource = normalize(source)

        def matchByAccount(f):
            return normalize(f.Name).startswith(normalizedSource.split("(")[0])

        if source and self.filterVar.get():
            filtered = sorted([f.Name for f in self.outlook.Folders if matchByAccount(f)])
        else:
            filtered = sorted([f.Name for f in self.outlook.Folders])

        self.destinationCombo['values'] = filtered
        self.destinationCombo.set("")
        account = self.extractAccountFromPstName(source)
        self.statusLabel.config(text=f"Destination options {'filtered' if self.filterVar.get() else 'reset'} for account: {account if account else 'all'}")

    def extractYearFromPstName(self, folderName):
        match = re.search(r"\((\d{4})\)", folderName)
        return int(match.group(1)) if match else None

    def onDestinationSelected(self, event):
        source = self.sourceCombo.get()
        dest = self.destinationCombo.get()
        if source and dest:
            year = self.extractYearFromPstName(dest)
            if year:
                self.overrideYearEntry.delete(0, tk.END)
                self.overrideYearEntry.insert(0, str(year))
            self.statusLabel.config(text=f"Moving {year} emails from '{source}' to '{dest}'")
            year = self.extractYearFromPstName(dest)
            self.statusLabel.config(text=f"Moving {year} emails from '{source}' to '{dest}'")

    def startMoveInThread(self):
        
        threading.Thread(target=self.onMoveEmails).start()

    def onMoveEmails(self):
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            sourceName = self.sourceCombo.get()
            sourceFolder = outlook.Folders[sourceName]
            destName = self.destinationCombo.get()
            destinationFolder = outlook.Folders[destName]

            self.ensureRequiredSubfolders(sourceFolder)
            self.ensureRequiredSubfolders(destinationFolder)
        except Exception as e:
            self.showError("Selection Error", f"Error resolving folders: {e}")
            return

        if self.createMissingVar.get():
            accountName = self.extractAccountFromPstName(sourceFolder.Name)
            self.checkAndCreateMissingPSTs(accountName, sourceFolder)

        yearOverride = self.overrideYearEntry.get()
        if yearOverride.isdigit():
            year = int(yearOverride)
            logger.info(f"...override year used: {year}")
        else:
            year = self.extractYearFromPstName(destinationFolder.Name)
            logger.info(f"...year: {year}")
            if not year:
                self.showError("Year Not Found", "Could not extract year from destination PST name.")
                return

        try:
            inboxSource = sourceFolder.Folders("Inbox")
            logger.info(f"...inboxSource: {sourceFolder.Name} > Inbox")
            sentSource = sourceFolder.Folders("Sent Items")
            logger.info(f"...sentSource: {sourceFolder.Name} > Sent Items")
            inboxDest = destinationFolder.Folders("Inbox")
            logger.info(f"...inboxDest: {destinationFolder.Name} > Inbox")
            sentDest = destinationFolder.Folders("Sent Items")
            logger.info(f"...sentDest: {destinationFolder.Name} > Sent Items")
        except Exception as e:
            self.showError("Folder Error", f"Expected subfolders 'Inbox' and 'Sent Items' in both PSTs: {e}")
            return

        logger.info(f"starting email move... from '{sourceFolder.Name}' to '{destinationFolder.Name}' for year {year}")
        movedInbox = self.moveEmailsByYear(inboxSource, inboxDest, year, "Inbox")
        movedSent = self.moveEmailsByYear(sentSource, sentDest, year, "Sent Items")
        logger.info(f"...completed email move. inbox: {movedInbox}, sent items: {movedSent}")

        print("Done", f"Moved {movedInbox} from Inbox, {movedSent} from Sent Items for year {year}.")
        self.statusLabel.config(text="Done...")
        
    def showError(self, title, message):
        logger.error(message)
        print(title, message)

    def moveEmailsByYear(self, sourceFolder, destFolder, year, folderLabel):
        dryRun = self.dryRunVar.get()
        if dryRun:
            logger.info("...dry run selected")
        
        movedCount = 0
        items = list(sourceFolder.Items)
        logger.info(f"...total items in {folderLabel}: {len(items)}")

        for item in items:
            try:
                mailDate = getattr(item, "SentOn", getattr(item, "ReceivedTime", None))
                if mailDate and mailDate.year == year:
                    subject = getattr(item, "Subject", "(No Subject)")
                    logger.info(f"moving: {folderLabel} | {mailDate.date()} | {subject}")
                    if not dryRun:
                        item.Move(destFolder)
                    movedCount += 1
                    logger.info(f"...movedCount ({folderLabel}): {movedCount}")
            except Exception as e:
                logger.error(f"Error moving item in {folderLabel}: {e}")
        return movedCount

    def checkAndCreateMissingPSTs(self, baseAccount, sourceFolder):
        existing = [f.Name for f in self.outlook.Folders]
        normalized = lambda s: s.replace(" ", "").lower()

        foundYears = set()
        for name in existing:
            if normalized(baseAccount) in normalized(name):
                year = self.extractYearFromPstName(name)
                if year:
                    foundYears.add(year)

        # gather all years used in source inbox + sent items
        requiredYears = set()
        for folderName in ["Inbox", "Sent Items"]:
            try:
                folder = sourceFolder.Folders(folderName)
                for item in folder.Items:
                    mailDate = getattr(item, "SentOn", getattr(item, "ReceivedTime", None))
                    if mailDate:
                        requiredYears.add(mailDate.year)
            except Exception as e:
                logger.error(f"Could not scan {folderName} for year detection: {e}")

        sourceYear = self.extractYearFromPstName(sourceFolder.Name)
        requiredYears = {y for y in requiredYears if 1980 <= y <= datetime.now().year and y != sourceYear}
        logger.info(f"...required years from email dates: {sorted(requiredYears)}")
        
        for year in requiredYears:
            if year not in foundYears:
                pstName = f"{baseAccount} ({year})"
                if self.dryRunVar.get():
                    logger.info(f"...missing PST for {year}, would create: {pstName}")
                else:
                    logger.info(f"...missing PST for {year}, creating: {pstName}")
                    try:
                        # Add PST and rename display name
                        pstPath = os.path.join(self.pstFolderPath, f"{pstName}.pst")
                        self.outlook.AddStoreEx(pstPath, 1)  # Unicode PST
                        storeFolder = self.outlook.Folders[self.outlook.Folders.Count - 1]
                        storeFolder.Name = pstName  # Rename to match expected format

                        # Ensure "Inbox" and "Sent Items" exist
                        for folderName in ["Inbox", "Sent Items"]:
                            try:
                                storeFolder.Folders.Add(folderName)
                                logger.info(f"...created folder: {folderName}")
                            except Exception as e:
                                logger.error(f"Could not create folder '{folderName}' in {pstName}: {e}")

                    except Exception as e:
                        logger.error(f"Could not create PST for {year}: {e}")

        self.statusLabel.config(text=f"PST files created...")
        self.createMissingVar.set(False)
        self.chkPSTCreate.config(state="disabled")
    
    def startScanInThread(self):
        threading.Thread(target=self.scanForMoveCandidates).start()

    def scanForMoveCandidates(self):
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        logger.info("scanning PSTs for all years...")
        results = {}

        for folder in outlook.Folders:
            try:
                name = folder.Name
                for subfolderName in ["Inbox", "Sent Items"]:
                    self.statusLabel.config(text=f"Scanning: {name} > {subfolderName}")
                    self.root.update_idletasks()
                    try:
                        subfolder = folder.Folders(subfolderName)
                        for item in subfolder.Items:
                            mailDate = getattr(item, "SentOn", getattr(item, "ReceivedTime", None))
                            if mailDate:
                                year = mailDate.year
                                key = (name, subfolderName, year)
                                results[key] = results.get(key, 0) + 1
                    except:
                        continue
            except:
                continue
        
        groupedResults = [(pst, folder, year, count) for (pst, folder, year), count in results.items()]
        self.showScanResults(groupedResults)
        msg = "...done" if groupedResults else "no emails found"
        self.statusLabel.config(text=msg)

    def showScanResults(self, results):
        win = tk.Toplevel(self.root)
        win.title("Scan Results by Year")
        tree = ttk.Treeview(win, columns=("pst", "folder", "year", "count"), show="headings")
        tree.heading("pst", text="PST Name")
        tree.heading("folder", text="Folder")
        tree.heading("year", text="Year")
        tree.heading("count", text="Emails")
        tree.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        for name, subfolder, year, count in sorted(results, key=lambda x: (x[0], x[2], x[1])):
            # Skip entries where the PST name includes (year)
            if f"({year})" in name:
                continue
            tree.insert("", "end", values=(name, subfolder, year, count))

        win.grid_rowconfigure(0, weight=1)
        win.grid_columnconfigure(0, weight=1)


    def getDefaultPstFolder(self):
            try:
                firstStore = self.outlook.Folders[0].Store
                path = firstStore.FilePath
                return os.path.dirname(path)
            except Exception as e:
                logger.error(f"Could not determine default PST folder path: {e}")
                return os.path.expanduser("~")

if __name__ == "__main__":
    root = tk.Tk()
    app = OrganiseMyOutlook(root)
    root.mainloop()

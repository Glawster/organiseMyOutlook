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
  
    def __init__(self, root):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.root = root
        self.root.title("Organise My Outlook")

        self.sourceFolder = None
        self.destinationFolder = None

        self.buildForm()

    def extractAccountFromPstName(self, folderName):
        match = re.match(r"(.+@[\w\.]+) \(\d{4}\)", folderName)
        return match.group(1) if match else None

    def onSourceSelected(self, event):
        self.updateDestinationList()

    def buildForm(self):
        logger.info("Starting...")
        allFolders = [f.Name for f in self.outlook.Folders]
        logger.info("...found these folders")
        for folder in sorted(allFolders):
            logger.info(f"... {folder}")
        self.filterVar = tk.BooleanVar(value=True)
        self.dryRunVar = tk.BooleanVar(value=True)
        sortedFolders = sorted([f.Name for f in self.outlook.Folders if self.isValidPstName(f.Name)])

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

        filterCheck = ttk.Checkbutton(self.root, text="Filter destinations by source account", variable=self.filterVar, command=self.updateDestinationList)
        filterCheck.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        ttk.Checkbutton(self.root, text="Dry Run (don't actually move emails)", variable=self.dryRunVar).grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="w")

        self.statusLabel = ttk.Label(self.root, text="", foreground="blue")
        self.statusLabel.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky="we")

        self.progressBar = ttk.Progressbar(self.root, mode='indeterminate')
        self.progressBar.grid(row=5, column=0, columnspan=2, pady=10, padx=10, sticky="we")
        self.progressBar.grid_remove()  # hide by default

        # button frame
        frmButton = ttk.Frame(self.root)
        frmButton.grid(row=6, column=0, columnspan=2, pady=10)
        
        btnRun = ttk.Button(frmButton, text="Move Emails", command=self.startMoveInThread)
        btnRun.pack(side="left", padx=10)

        btnExit = ttk.Button(frmButton, text="Exit", command=self.root.quit)
        btnExit.pack(side="left", padx=10)


    def isValidPstName(self, folderName):
        return re.match(r".+@\w+\.\w+ \(\d{4}\)", folderName)

    def updateDestinationList(self):
        source = self.sourceCombo.get()
        account = self.extractAccountFromPstName(source)
        if account and self.filterVar.get():
            filtered = sorted([f.Name for f in self.outlook.Folders if account in f.Name and self.isValidPstName(f.Name)])
        else:
            filtered = sorted([f.Name for f in self.outlook.Folders if self.isValidPstName(f.Name)])
        self.destinationCombo['values'] = filtered
        self.destinationCombo.set("")
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
        self.progressBar.grid()
        self.progressBar.start()
        threading.Thread(target=self.onMoveEmails).start()

    def onMoveEmails(self):
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            sourceName = self.sourceCombo.get()
            sourceFolder = outlook.Folders[sourceName]
            destName = self.destinationCombo.get()
            destinationFolder = outlook.Folders[destName]
        except Exception as e:
            self.showError("Selection Error", f"Error resolving folders: {e}")
            return

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

        self.progressBar.stop()
        self.progressBar.grid_remove()
        print("Done", f"Moved {movedInbox} from Inbox, {movedSent} from Sent Items for year {year}.")
        self.statusLabel.config(text="Done...")
        

    def showError(self, title, message):
        self.progressBar.stop()
        self.progressBar.grid_remove()
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

if __name__ == "__main__":
    root = tk.Tk()
    app = OrganiseMyOutlook(root)
    root.mainloop()

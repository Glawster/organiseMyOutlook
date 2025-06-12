# 🧭 Project Guidelines

This document defines the coding, logging, and naming conventions used across all Python projects for consistency and maintainability.

---

## 📁 Project Structure

```
project-name/
├── src/                    # Core application code
├── tests/                  # Unit tests
├── logs/                  # Runtime logs
├── requirements.txt        # Runtime dependencies
├── dev-requirements.txt    # Dev tools like linters, formatters
├── README.md
├── .gitignore
├── project_guidelines.md   # This file
```

---

## ✍️ Code Style Conventions

| Element      | Convention         | Example                    |
|--------------|--------------------|----------------------------|
| Variables    | camelCase          | emailList, logFilePath     |
| Functions    | camelCase          | moveEmailsByYear()         |
| Classes      | TitleCase          | OutlookEmailMoverApp       |
| Constants    | UPPER_SNAKE_CASE   | LOG_LEVEL, DEFAULT_YEAR    |
| File Names   | camelCase.py       | organiseMyOutlook.py       |

Use `black` for formatting (4 spaces per indent, ≤100 character lines).

---

## 🪪 Logging Guidelines

### Format:
```python
'%(asctime)s [%(module)s] %(levelname)s %(message)s'
```

### Message Style:
| Context                  | Example                         |
|--------------------------|---------------------------------|
| General action           | ...starting email move          |
| Completion               | emails moved successfully...    |
| Data output              | ...emails processed: 45         |
| Error (Sentence case)    | ERROR - Folder not found        |

### Logging Example:
```python
logger.info("...scanning folder")
logger.info("emails found: 230")
logger.error("Failed to open PST file")
```

---

## 🔁 Reusable Logging Setup

```python
def setupLogging(title: str) -> logging.Logger:
    import os
    import logging
    import datetime

    title = title.replace(" ", "")
    logger = logging.getLogger(title)
    if not logger.handlers:
        logDir = os.getcwd()
        os.makedirs(logDir, exist_ok=True)
        logDate = datetime.datetime.now().strftime("%Y%m%d")
        logFilePath = os.path.join(logDir, f"{title}.{logDate}.log")

        handler = logging.FileHandler(logFilePath)
        formatter = logging.Formatter('%(asctime)s [%(module)s] %(levelname)s %(message)s')
        handler.setFormatter(formatter)

        logger.setLevel(logging.INFO)
        logger.addHandler(handler)

    return logger
```

---

## 🖼 GUI Naming Conventions

| Widget Type    | Prefix Example       |
|----------------|----------------------|
| Button         | btnSubmit            |
| Entry          | entryUsername        |
| Label          | lblTitle             |
| Frame          | frmMain              |
| Text           | txtLogOutput         |
| Listbox        | lstEmails            |
| Checkbutton    | chkRememberMe        |
| Radiobutton    | rdoOptionOne         |
| Combobox       | cmbFolderList        |
| Handlers       | onLoadData()         |

Enforced by `guiNamingLinter.py`.

---

## ✅ Linting & Enforcement Tools

| Tool         | Role                            |
|--------------|----------------------------------|
| black        | Auto-formatting Python code     |
| pytest       | Unit testing framework          |
| guiNamingLinter.py | Custom GUI naming/static checker |
| pre-commit   | Auto-run checks before commits  |

### Example dev-requirements.txt
```text
black
pytest
pre-commit
```

---

## 🧪 Testing Conventions

- Use `pytest`
- Place all test files in `tests/`
- Name tests like `test_functionName.py`
- Avoid testing GUI directly; test logic only

---

## 📦 Packaging & Distribution

If needed:
- Use `pyinstaller` to bundle into .exe
- Add setup.py or pyproject.toml for pip distribution

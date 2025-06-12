import logging

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

        handler = logging.FileHandler(logFilePath, encoding='utf-8')
        formatter = logging.Formatter('%(asctime)s [%(module)s] %(levelname)s %(message)s')
        handler.setFormatter(formatter)

        logger.setLevel(logging.INFO)
        logger.addHandler(handler)

    return logger

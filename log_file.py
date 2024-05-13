import logging


def setup_logger(log_file):
    logger = logging.getLogger(log_file)
    # Check if handlers are already added to the logger
    if len(logger.handlers) == 0:
        logger.setLevel(logging.DEBUG)
        handler = logging.FileHandler(log_file)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', '%H:%M')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    return logger

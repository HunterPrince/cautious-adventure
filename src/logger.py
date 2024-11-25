import logging

def setup_logging(log_file="./data/log/app.log"):
    """
    Sets up logging configuration.

    Parameters
    ----------
    log_file : str, optional
        The name of the log file. Default is "app.log".
    """
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s - %(filename)s:%(lineno)d',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    logger = logging.getLogger(__name__)
    return logger

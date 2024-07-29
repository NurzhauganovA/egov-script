import os
import logging


logging.basicConfig(level=logging.DEBUG)


def get_file_logger(file_name):
    if not os.path.exists('logs'):
        os.makedirs('logs')

    if file_name not in os.listdir('logs'):
        open(f'logs/{file_name}', 'w').close()

    logger = logging.getLogger(file_name)
    logger.setLevel(logging.INFO)
    file_handler = logging.FileHandler(f'logs/{file_name}')
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    return logger

import os
import configparser

# 获取项目根目录路径
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(ROOT_DIR, "config.ini")


def getConfigByKey(key='DEFAULT'):
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE, encoding="utf-8")
    return config[key]

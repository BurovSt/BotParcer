import configparser

config = None


def get_config():

    global config

    if config is None:
        config = configparser.ConfigParser()
        config.read('config.ini')
        if not config.sections():
            print('read error')
    return config


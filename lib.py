# -*- coding: utf-8 -*-
# Общая библиотека функций
# ver 1.0

import string
from configparser import ConfigParser

def lenl(a):
    try:
        if a != None:
            a = str(a).strip()
            if  a != '':
                a = ''.join([char for char in a if char in string.digits])
                return len(a)
        return 0
    except TypeError:
        return 0


def l(a):
    try:
        if a != None:
            a = str(a).strip()
            if  a != '':
                a = ''.join([char for char in a if char in string.digits])
                if len(a) > 0:
                    return int(a)
                else:
                    return 0
        return 0
    except TypeError:
        return 0

def s(a):
    try:
        if a != None:
            return str(a).strip().replace(u"\xa0", u" ")
        return ''
    except TypeError:
        return ''

def read_config(filename='config.ini', section='mysql'):
    """ Read database configuration file and return a dictionary object
    :param filename: name of the configuration file
    :param section: section of database configuration
    :return: a dictionary of database parameters
    """
    # create parser and read ini configuration file
    parser = ConfigParser()
    parser.read(filename)

    # get section, default to mysql
    db = {}
    if parser.has_section(section):
        items = parser.items(section)
        for item in items:
            db[item[0]] = item[1]
    else:
        raise Exception('{0} not found in the {1} file'.format(section, filename))

    return db

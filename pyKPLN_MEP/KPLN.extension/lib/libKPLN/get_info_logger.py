# coding: utf-8

import logging


class InfoLogger:
    """ This class export user data to special .txt file.
    In input - he need string user name and human name of the script.
    To get user name use this:
    >>> from System.Security.Principal import WindowsIdentity
    >>>
    and:
    >>> user_name = WindowsIdentity.GetCurrent().Name
    """

    def __init__(self, user_name, script_name):
        self.user_name = user_name
        self.script_name = script_name

        logger_createor = logging.getLogger("{}:{}".format(script_name,
                                                           user_name))
        log_handler = logging.FileHandler("Z:\\Отдел BIM\\Куцко Тимофей\\DATA_pyRevit\\all_pyKPLN_data.log")
        log_format = logging.Formatter("%(levelname)s:%(name)s:%(asctime)s",
                                       datefmt='%d-%b-%y:%H_%M_%S')
        log_handler.setFormatter(log_format)
        logger_createor.addHandler(log_handler)
        logger_createor.info(self.user_name)

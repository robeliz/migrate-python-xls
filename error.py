import logging

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# create a file handler
handler = logging.FileHandler('log/ERROR.log')

handler.setLevel(logging.DEBUG)
# create a logging format
formatter = logging.Formatter('%(levelname)s - %(message)s')
handler.setFormatter(formatter)
# add the handlers to the logger
logger.addHandler(handler)


class Error(object):
    excel_name = ""
    row_number = ""
    col_number = ""
    function = ""
    description = ""
    type = ""
    sql = ""

    def save_error(self):
        pass

    def log_error(self):
        logger.error(self.excel_name + " | " + str(self.row_number) + " | " + str(self.col_number) + " | " + self.function + " | " + self.description + " | " + self.sql + " | ")
        pass

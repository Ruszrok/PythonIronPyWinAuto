
try:
    from logger import logger
    foundLogger = True
except:
    foundLogger = False



class ActionLogger:

    def __init__(self, logFilePath = None):
        if foundLogger:
            self.logger = logger.Logger()

    def log(self, msg):
        if foundLogger:
            self.logger.message(msg)
        else:
            print(msg)



class QueueManager:

    __instance = None

    @staticmethod
    def getInstance():
        """ Static access method. """
        if not QueueManager.__instance:
            QueueManager()
        return QueueManager.__instance

    def __init__(self):
        if not QueueManager.__instance:

            import logging
            import queue

            logging.basicConfig(level=logging.INFO, format='{asctime} {levelname} ({threadName:11s}) {message}',
                                style='{')

            logging.info(f"Creating QueueManager Instance")
            QueueManager.__instance = self
            self.marketdataQueue = queue.Queue()

    def sendMarketData(self,msg):
        self.marketdataQueue.put(msg)

    def readMarketData(self):
        return self.marketdataQueue.get()
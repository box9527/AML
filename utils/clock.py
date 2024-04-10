#!/usr/bin/env python3
# coding: utf-8
# @carl9527


import datetime
import logging
import time
import threading
from loguru import logger


class Clock(threading.Thread):
    """Class to display the time every seconds

    Every 5 seconds, the time is displayed using the logging.ERROR level
    to show that different colors are associated to the log levels
    """
    def __init__(self):
        super().__init__()
        self._stop_event = threading.Event()

    def run(self):
        logger.debug('Clock started')
        previous = -1
        while not self._stop_event.is_set():
            now = datetime.datetime.now()
            if previous != now.second:
                previous = now.second
                if now.second % 5 == 0:
                    level = logging.ERROR
                else:
                    level = logging.INFO
                logger.log(level, now)
            time.sleep(0.2)

    def stop(self):
        self._stop_event.set()

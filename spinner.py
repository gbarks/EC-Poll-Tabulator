#!/usr/bin/env python3

# ==========================================================
#  Multithreaded console loading spinner, stolen from:
#   https://gist.github.com/anonymous/c4ad2bbe2d5552af05c2
# ==========================================================

import itertools
import sys
import time
import threading

class Spinner(object):
    spinner_cycle = itertools.cycle(['|', '/', '-', '\\'])

    def __init__(self):
        self.stop_running = threading.Event()
        self.spin_thread = threading.Thread(target=self.init_spin)

    def start(self):
        self.spin_thread.start()

    def stop(self):
        self.stop_running.set()
        self.spin_thread.join()

    def init_spin(self):
        while not self.stop_running.is_set():
            sys.stdout.write(next(self.spinner_cycle))
            sys.stdout.flush()
            time.sleep(0.25)
            sys.stdout.write('\b')

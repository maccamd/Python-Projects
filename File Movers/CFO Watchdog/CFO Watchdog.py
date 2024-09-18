import time
import logging
import multiprocessing
from watchdog.observers import Observer
from watchdog.events import LoggingEventHandler

event_handler = LoggingEventHandler()
observer = Observer()

folder = r'W:\EDC Schedules\RDEC Weekly Checks Results\VDC 1\CFO CHECKS'
folder_1 = r'W:\EDC Schedules\RDEC Weekly Checks Results\VDC 2\CFO CHECKS'

def monitor(folder):
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    observer.schedule(event_handler, folder, recursive=True)
    observer.start()
    

    try:
        while True:
            time.sleep(60)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

if __name__ == "__main__":
    m = multiprocessing.Process(target=monitor, args=(folder,))
    m1 = multiprocessing.Process(target=monitor, args=(folder,))
    m.start()
    m1.start()
    
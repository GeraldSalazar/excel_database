import promptlib
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import os
import shutil
import pandas as pd


class Watcher:
    DIRECTORY_TO_WATCH = "path"

    def __init__(self):
        processedDir = os.path.join(os.getcwd(), 'Processed')
        notApplicableDir = os.path.join(os.getcwd(), 'Not applicable')
        try: 
            os.mkdir(processedDir) 
            os.mkdir(notApplicableDir) 
        except OSError as error: 
            print('Folders already exists') 
        self.observer = Observer()

    def run(self):
        prompter = promptlib.Files()
        dir = prompter.dir()
        self.DIRECTORY_TO_WATCH = dir
        print (self.DIRECTORY_TO_WATCH)
        event_handler = Handler()
        self.observer.schedule(event_handler, self.DIRECTORY_TO_WATCH, recursive=True)
        self.observer.start()
        try:
            while True:
                time.sleep(5)
        except:
            self.observer.stop()
            print ("Error")

        self.observer.join()


class Handler(FileSystemEventHandler):

    @staticmethod
    def on_any_event(event):
        if event.is_directory:
            return None

        elif event.event_type == 'created':
            # Check if file created is a excel file
            print (event.src_path)
            print (os.getcwd())

            if(isExcelFile(event.src_path)):

                try:
                    shutil.move(event.src_path, os.getcwd()+'\\Processed')
                except:
                    print('Done')
            else:
                try:
                    shutil.move(event.src_path, os.getcwd()+'\\Not applicable')
                except:
                    print('Done')

def isExcelFile(path):
    ext = os.path.splitext(path)[-1].lower()
    return ext == '.xls'

if __name__ == '__main__':
    w = Watcher()
    w.run()

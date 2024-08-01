import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import subprocess

# Define the event handler
class NewFileHandler(FileSystemEventHandler):
    def __init__(self, script_path):
        self.script_path = script_path
    def on_created(self, event):
        # Execute your script when a new file is created
        subprocess.run(['python', self.script_path, event.src_path])

# Define the folder to monitor
folder_to_watch = 'E:/excel-converter-script/'
script_path = 'E:/excel-converter-script/etl_script.py'
# Create the event handler and observer
event_handler = NewFileHandler(script_path)
observer = Observer()
observer.schedule(event_handler, folder_to_watch, recursive=False)

# Start monitoring the folder
observer.start()

try:
    # Keep the script running
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    # Stop the observer if the script is interrupted
    observer.stop()

# Wait for the observer to join
observer.join()
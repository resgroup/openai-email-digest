import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


class MyHandler(FileSystemEventHandler):
    def on_modified(self, event):
        # This function is called when a file is modified.
        if event.is_directory:
            return None

        if event.event_type == "modified":
            # Take whatever action you want here when a file is modified.
            print(f"File {event.src_path} has been modified")


def main():
    path = "C:\\path\\to\\your\\folder"  # CHANGE THIS TO THE FOLDER YOU WANT TO MONITOR
    event_handler = MyHandler()
    observer = Observer()
    observer.schedule(event_handler, path, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


if __name__ == "__main__":
    main()

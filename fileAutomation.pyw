import os
from shutil import move
from time import sleep
import json
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

with open("config.json", "r") as file:
    global config
    config = json.load(file)

source_dir = config["source_dir"]
dest_dir_image = config["dest_dir_image"]
dest_dir_video = config["dest_dir_video"]
dest_dir_music = config["dest_dir_music"]
dest_dir_documents = config["dest_dir_documents"]

image_extensions = [".jpg", ".jpeg", ".jpe", ".jif", ".jfif", ".jfi", ".png", ".gif", ".webp", ".tiff", ".tif", ".psd", ".raw", ".arw", ".cr2", ".nrw", ".k25", ".bmp", ".dib", ".heif", ".heic", ".ind", ".indd", ".indt", ".jp2", ".j2k", ".jpf", ".jpf", ".jpx", ".jpm", ".mj2", ".svg", ".svgz", ".ai", ".eps", ".ico"]
video_extensions = [".webm", ".mpg", ".mp2", ".mpeg", ".mpe", ".mpv", ".ogg", ".mp4", ".mp4v", ".m4v", ".avi", ".wmv", ".mov", ".qt", ".flv", ".swf", ".avchd"]
audio_extensions = [".m4a", ".flac", "mp3", ".wav", ".wma", ".aac"]
document_extensions = [".doc", ".docx", ".odt", ".pdf", ".xls", ".xlsx", ".ppt", ".pptx"]

def make_unique(dest, path):
    filename, extension = os.path.splitext(path)
    counter = 1

    while os.path.exists(f"{dest}/{path}"):
        path = f"{filename} ({counter}){extension}"
        counter += 1

    return path

def move_file(dest, entry, name):
    unique_name = name
    if os.path.exists(f"{dest}/{name}"):
        unique_name = make_unique(dest, name)
        os.chdir(source_dir)
        os.rename(name, unique_name)

    move(f"{source_dir}/{unique_name}", dest)

class MoverHandler(FileSystemEventHandler):
    def on_modified(self, event):
        global config
        with os.scandir(source_dir) as entries:
            for entry in entries:
                name = entry.name
                if config["dir_image_check"]: self.check_image_files(entry, name)
                if config["dir_video_check"]: self.check_video_files(entry, name)
                if config["dir_music_check"]: self.check_music_files(entry, name)
                if config["dir_documents_check"]: self.check_document_files(entry, name)

    def check_music_files(self, entry, name):
        global config
        if config["dir_music_check"]:
            for audio_extension in audio_extensions:
                if name.endswith(audio_extension) or name.endswith(audio_extension.upper()):
                    dest = dest_dir_music
                    move_file(dest, entry, name)

    def check_video_files(self, entry, name):
        global config
        if config["dir_video_check"]:
            for video_extension in video_extensions:
                if name.endswith(video_extension) or name.endswith(video_extension.upper()):
                    move_file(dest_dir_video, entry, name)

    def check_image_files(self, entry, name):
        global config
        if config["dir_image_check"]:
            for image_extension in image_extensions:
                if name.endswith(image_extension) or name.endswith(image_extension.upper()):
                    move_file(dest_dir_image, entry, name)

    def check_document_files(self, entry, name):
        global config
        if config["dir_documents_check"]:
            for documents_extension in document_extensions:
                if name.endswith(documents_extension) or name.endswith(documents_extension.upper()):
                    move_file(dest_dir_documents, entry, name)

if __name__ == "__main__":
    path = source_dir
    event_handler = MoverHandler()
    observer = Observer()
    observer.schedule(event_handler, path, recursive=True)
    observer.start()
    try:
        while True:
            sleep(10)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

import win32api
import win32print
from watchdog.observers import Observer
from watchdog.events import *
import json
import sys
import traceback



def printer_loading(filename, printer):
    open(filename, "r")
    win32api.ShellExecute(
        0,
        "print",
        filename,
        '/d:"%s"' % printer,
        ".",
        0
    )

class FileEventHandler(FileSystemEventHandler):
    def __init__(self,printer):
        FileSystemEventHandler.__init__(self)
        self.printer = printer

    def on_created(self, event):
        if event.is_directory:
            print("directory created:{0}".format(event.src_path))
        else:
            print("file created:{0}".format(event.src_path))
            if event.src_path.endswith("docx") or event.src_path.endswith("doc"):
                if "~" not in event.src_path:
                    print(f"打印{event.src_path}")
                    printer_loading(event.src_path, self.printer)

import time
if __name__ == "__main__":

    try:
        observers = []
        print("当前检测到的打印机:")
        maxL = 2
        Num = 2
        for i in range(1, 40):
            # print(i)
            r = win32print.EnumPrinters(i)
            l = len(r)
            if l > maxL:
                maxL = l
                Num = i

        for i, p in enumerate(list(win32print.EnumPrinters(Num))):
            print(f"\t{i}:{p[1]}\n")

        # f = open('./settings.txt', 'r', encoding="utf-8")
        settings = {}
        # print("当前配置:")
        # for line in f.readlines():
        #     k, v = line.split(",")
        #     settings[k.strip()] = v.strip()
        #     print(f"\t{k}:{v}")



        print("当前配置:")
        with open("./settings.json", 'rb') as load_f:
            ds = json.load(load_f)
            for d in ds:
                settings[d['url'].strip()] = d["printer"].strip()
                print(f"\t{d}")
        pass
        for path, printer in settings.items():
            observer = Observer()
            event_handler = FileEventHandler(printer)
            observer.schedule(event_handler, path, True)
            observer.start()
            observers.append(observer)

        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            for observer in observers:
                observer.stop()

        for observer in observers:
            observer.join()
    except Exception as e:
        print(e)
        traceback.print_exc()
        input()
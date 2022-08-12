import os

import win32com.client
# app = win32com.client.DispatchEx("kwps.Application")
## 设置不可见
# app.visible = 0
# doc = app.Documents.Open(r"C:\Users\ljing\Desktop\translate.docm")
# app.Application.Run("addPage")
# doc.Close()
# app.Application.Quit()

# import win32com.client
# o = win32com.client.Dispatch("kwps.application")
# o.Visible=True
# doc = o.Documents.Add()
# doc.Content.text="Hello world!"
import time
import sys

sys.meta_path.append(os.path.dirname(__file__))


class OperateWps:
    def __init__(self, filename):
        self.app = win32com.client.DispatchEx("kwps.Application")
        self.filename = filename
        self.app.visible = 0
        self.doc = self.app.Documents.Open(self.filename)

    def run(self, *args):
        self.app.Application.Run("addPage")
        # self.app.Application.Run("find_content", ("简介", 10, "黑体"))
        pass

    def __del__(self):
        self.doc.Close()


if __name__ == '__main__':
    a = OperateWps(r".\translate.docm")
    a.run()

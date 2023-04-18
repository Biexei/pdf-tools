import wx
import os
from PyPDF2 import PdfFileReader, PdfFileWriter
import time


def pdf_merge(out_put_path: str, *input_files) -> None:
    """
    合并pdf
    :param out_put_path: 合并结果输出路径
    :param input_files: 待合并pdf文件
    :return: None
    """
    pdf_writer = PdfFileWriter()
    out_path = out_put_path + r"\合并结果%s.pdf" % str(int(time.time()))
    for file in input_files:
        reader = PdfFileReader(open(file=file, mode='rb'))
        page_size = reader.getNumPages()
        for i in range(page_size):
            pdf_writer.addPage(reader.getPage(i))
    with open(file=out_path, mode='wb'):
        pdf_writer.write(out_path)


def pdf_split(in_put_file: str, size_range: str, out_put_path: str) -> None:
    """
    拆分pdf
    :param in_put_file: 被拆分pdf文件
    :param size_range: 起始页（包含）-截止页（包含）,如1-3
    :param out_put_path: 拆分结果输出目录
    :return: None
    """
    rg = size_range.split(",")
    for _ in rg:
        range_split = _.split("-")
        start = int(range_split[0])
        end = int(range_split[1])
        writer = PdfFileWriter()
        reader = PdfFileReader(open(file=in_put_file, mode='rb'))
        page_size = reader.getNumPages()
        if start < 1:
            raise Exception("起始页参数错误")
        if start > end:
            raise Exception("参数错误")
        if end > page_size:
            raise Exception("截止页参数错误，超出最大页码数：%s" % str(page_size))
        for i in range(start - 1, end):
            writer.addPage(reader.getPage(i))
        path = out_put_path + r"\\%s-%s.pdf" % (start, end)
        with open(file=path, mode='wb'):
            writer.write(path)


class SiteLog(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, title='PDF工具', size=(640, 480))
        # PDF合并
        self.SelBtn = wx.Button(self, label='PDF合并', pos=(5, 5), size=(70, 70))
        self.file_name_text_1 = wx.StaticText(self, label='文件1', pos=(80, 5), size=(50, 25))
        self.file_path_1 = wx.TextCtrl(self, pos=(140, 5), size=(230, 25))
        self.choose_btn_1 = wx.Button(self, label='选择文件1', pos=(400, 5), size=(80, 25))
        self.choose_btn_1.Bind(wx.EVT_BUTTON, self.on_open_file1)
        self.file_name_text_2 = wx.StaticText(self, label='文件2', pos=(80, 50), size=(50, 25))
        self.file_path_2 = wx.TextCtrl(self, pos=(140, 50), size=(230, 25))
        self.choose_btn_2 = wx.Button(self, label='选择文件2', pos=(400, 50), size=(80, 25))
        self.choose_btn_2.Bind(wx.EVT_BUTTON, self.on_open_file2)
        self.merge_btn = wx.Button(self, label='合并', pos=(520, 5), size=(70, 70))
        self.merge_btn.Bind(wx.EVT_BUTTON, self.on_merge)

        # PDF拆分
        self.SelBtn = wx.Button(self, label='PDF拆分', pos=(5, 200), size=(70, 70))
        self.file_name_text = wx.StaticText(self, label='文件', pos=(80, 200), size=(50, 25))
        self.file_path = wx.TextCtrl(self, pos=(140, 200), size=(230, 25))
        self.choose_btn = wx.Button(self, label='选择文件', pos=(400, 200), size=(80, 25))
        self.choose_btn.Bind(wx.EVT_BUTTON, self.on_open_file)
        self.size_range = wx.StaticText(self, label='区间', pos=(80, 245), size=(50, 25))
        self.size_range_value = wx.TextCtrl(self, pos=(140, 245), size=(230, 25))
        self.size_range_value.SetHint("如1-4，多个区间如1-3,3-4")
        self.split_btn = wx.Button(self, label='拆分', pos=(520, 200), size=(70, 70))
        self.split_btn.Bind(wx.EVT_BUTTON, self.on_split)

    def on_open_file1(self, event):
        """
        PDF合并文件1选择事件
        :param event:
        :return:
        """
        wildcard = 'Allfiles(*.*)|*.*'
        dialog = wx.FileDialog(None, 'select', os.getcwd(), '', wildcard, wx.FC_OPEN)
        if dialog.ShowModal() == wx.ID_OK:
            self.file_path_1.SetValue(dialog.GetPath())
            dialog.Destroy()

    def on_open_file2(self, event):
        """
        PDF合并文件2选择事件
        :param event:
        :return:
        """
        wildcard = 'Allfiles(*.*)|*.*'
        dialog = wx.FileDialog(None, 'select', os.getcwd(), '', wildcard, wx.FC_OPEN)
        if dialog.ShowModal() == wx.ID_OK:
            self.file_path_2.SetValue(dialog.GetPath())
            dialog.Destroy()

    def on_open_file(self, event):
        """
        PDF拆分文件选择事件
        :param event:
        :return:
        """
        wildcard = 'Allfiles(*.*)|*.*'
        dialog = wx.FileDialog(None, 'select', os.getcwd(), '', wildcard, wx.FC_OPEN)
        if dialog.ShowModal() == wx.ID_OK:
            self.file_path.SetValue(dialog.GetPath())
            dialog.Destroy()

    def on_merge(self, event):
        """
        PDF合并事件
        :param event:
        :return:
        """
        crt_path = os.getcwd()
        pdf_merge(crt_path, self.file_path_1.GetValue(), self.file_path_2.GetValue())
        toast = wx.MessageDialog(None, "合并成功")
        if toast.ShowModal() == wx.ID_YES:
            toast.Destroy()

    def on_split(self, event):
        """
        PDF拆分事件
        :param event:
        :return:
        """
        toast = "拆分成功"
        crt_path = os.getcwd()
        try:
            pdf_split(self.file_path.GetValue(), self.size_range_value.GetValue(), crt_path)
        except Exception as e:
            toast = str(e)
        toast = wx.MessageDialog(None, toast)
        if toast.ShowModal() == wx.ID_YES:
            toast.Destroy()


if __name__ == '__main__':
    app = wx.App()
    SiteFrame = SiteLog()
    SiteFrame.Show()
    app.MainLoop()

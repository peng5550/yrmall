# coding:utf-8
import threading
from tkinter.ttk import Scrollbar
from mttkinter import mtTkinter as tk
from tkinter.messagebox import showinfo, showwarning, showerror
from tkinter import ttk
from tkinter import filedialog
import aiohttp
import asyncio
from openpyxl import Workbook


class Application(object):

    def __init__(self):
        self.__creat_UI()

    def __creat_UI(self):
        self.window = tk.Tk()
        self.window.title("yrmall")
        self.window.geometry("500x500+500+50")
        self.label_id = tk.Label(self.window, text="查询商品ID")
        self.label_id.place(x=20, y=20, width=80, height=30)

        self.label_id_from = tk.Label(self.window, text="From")
        self.label_id_from.place(x=95, y=20, width=50, height=30)

        self.label_id_end = tk.Label(self.window, text="End")
        self.label_id_end.place(x=95, y=70, width=50, height=30)

        self.entry_id_from = tk.Entry(self.window)
        self.entry_id_from.place(x=150, y=20, width=120, height=30)

        self.entry_id_end = tk.Entry(self.window)
        self.entry_id_end.place(x=150, y=70, width=120, height=30)

        self.btn_start = tk.Button(self.window, text="开始采集", command=lambda: self.thread_it(self.start_task))
        self.btn_start.place(x=310, y=20, width=100, height=30)

        self.btn_excel = tk.Button(self.window, text="导出Excel", command=lambda: self.thread_it(self.save2excel))
        self.btn_excel.place(x=310, y=70, width=100, height=30)

        self.label_show_data = tk.Label(self.window, text="数据展示    ")
        self.label_show_data.place(x=20, y=120, width=80, height=30)

        title = ['1', '2', '3', '4', '5', '6', '7', '8']
        self.box = ttk.Treeview(self.window, columns=title, show='headings')
        self.box.place(x=50, y=180, width=400, height=300)
        self.box.column('1', width=50, anchor='center')
        self.box.column('2', width=50, anchor='center')
        self.box.column('3', width=100, anchor='center')
        self.box.column('4', width=50, anchor='center')
        self.box.column('5', width=100, anchor='center')
        self.box.column('6', width=100, anchor='center')
        self.box.column('7', width=100, anchor='center')
        self.box.column('8', width=500, anchor='center')
        self.box.heading('1', text='序号')
        self.box.heading('2', text='ID')
        self.box.heading('3', text='标题')
        self.box.heading('4', text='价格')
        self.box.heading('5', text='尺码')
        self.box.heading('6', text='颜色')
        self.box.heading('7', text='运费')
        self.box.heading('8', text='图片链接')

        self.VScroll1 = Scrollbar(self.box, orient='vertical', command=self.box.yview)
        self.VScroll1.pack(side="right", fill="y")
        self.VScroll2 = Scrollbar(self.box, orient='horizontal', command=self.box.xview)
        self.VScroll2.pack(side="bottom", fill="x")
        self.box.configure(yscrollcommand=self.VScroll1.set)
        self.box.configure(xscrollcommand=self.VScroll2.set)

    def __make_url(self):
        id_from = self.entry_id_from.get().strip()
        if not id_from:
            showerror("警告", "请输入关键词!")
            return
        id_end = self.entry_id_end.get().strip()
        if not id_end:
            id_end = id_from
        base_url = "https://yrmall.net/api/goods/get/{}"
        return [base_url.format(id) for id in range(int(id_from), int(id_end) + 1)]

    async def __get_content(self, semaphore, link):
        goodsId = link.split("/")[-1]
        conn = aiohttp.TCPConnector(verify_ssl=False)
        headers = {
            'Accept': 'application/json',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Cache-Control': 'no-cache',
            'Connection': 'keep-alive',
            'Pragma': 'no-cache',
            'Referer': f'https://yrmall.net/products/details?id={goodsId}',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-origin',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.135 Safari/537.36'
        }
        async with semaphore:
            async with aiohttp.ClientSession(connector=conn, headers=headers) as sess:
                async with sess.get(link) as resp:
                    content = await resp.json()
                    if content["code"] == "200":
                        return content, goodsId
                    return

    async def __crawler(self, semaphore, link):
        content, goodsId = await self.__get_content(semaphore, link)
        if content:
            self.data_index += 1
            size = []
            color = []
            for opt in content["data"]["options"]:
                if opt["option_name"] == "Size":
                    for s in opt["items"]:
                        size.append(s["item_name"])
                elif opt["option_name"] == "color":
                    for c in opt["items"]:
                        color.append(c["item_name"])

            product_info = [
                self.data_index,
                goodsId,
                content["data"]["goods_title"],
                content["data"]["price"],
                ", ".join(size),
                ", ".join(color),
                content["data"]["pre_delivery_fee"],
                ", ".join(content["data"]["other_images"]),
            ]
            self.datas.append(product_info)
            self.box.insert("", "end", values=product_info)
            self.box.yview_moveto(1.0)

    def save2excel(self):
        if not self.datas:
            showwarning("警告", "当前不存在任何数据!")
            return
        filePath = filedialog.asksaveasfilename(title="保存文件", filetypes=[("xlsx", ".xlsx")])
        file_name = f"{filePath}.xlsx"
        wb = Workbook()
        ws = wb.active
        for line in self.datas:
            ws.append(line)
        wb.save(file_name)
        showinfo("提示信息", "保存成功！")

    async def task_manager(self, url_list, func):
        tasks = []
        sem = asyncio.Semaphore(5)
        if url_list:
            for url in url_list:
                task = asyncio.ensure_future(func(sem, url))
                tasks.append(task)
            await asyncio.gather(*tasks)

    def start(self, url_list):
        new_loop = asyncio.new_event_loop()
        asyncio.set_event_loop(new_loop)
        loop = asyncio.get_event_loop()
        loop.run_until_complete(self.task_manager(url_list, self.__crawler))

    def start_task(self):
        self.box.delete(*self.box.get_children())
        self.datas = [["序号", "ID", "标题", "价格", "尺码", "颜色", "运费", "图片链接"]]
        self.data_index = 0
        products_urls = self.__make_url()
        self.start(products_urls)
        showinfo("提示信息", "采集完成!")

    @staticmethod
    def thread_it(func, *args):
        t = threading.Thread(target=func, args=args)
        t.setDaemon(True)
        t.start()

    def run(self):
        self.window.mainloop()


if __name__ == '__main__':
    demo = Application()
    demo.run()

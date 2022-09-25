
# coding: utf-8

# In[18]:


#文件路径选取和运行窗口
from tkinter import *
from tkinter.filedialog import askopenfilename,askdirectory
from tkinter.messagebox import showinfo
import os
import time
import logging


pb_x = 0
pb_y = 0
pb_width = 300
pb_height = 20

pb_bg = "white"
pb_fg = "green"
pb_frame_size = 0
canvas = Canvas





class progress:
    def __init__(self):
        pass

    # 初始化，创建Canvas实例，设定坐标和宽高
    def init(self, master, x=pb_x, y=pb_y,
             width=pb_width, height=pb_height,
             bg=pb_bg, fg=pb_fg,
             frame_size=pb_frame_size):
        global pb_x
        pb_x = x

        global pb_y
        pb_y = y

        global pb_width
        pb_width = width

        global pb_height
        pb_height = height

        global pb_bg
        pb_bg = bg

        global pb_fg
        pb_fg = fg

        global pb_frame_size
        pb_frame_size = frame_size

        global canvas
        canvas = Canvas(master, width=width, height=height, bg=pb_bg)
        canvas.place(x=x, y=y)

    # 运行进度条
    def run(self, master, percentage, text=None):
        global canvas
        fill_line = canvas.create_rectangle(1.5, 1.5, 0, 23, width=0, fill=pb_fg)
        canvas.coords(fill_line, (0, 0, percentage, 60))
        if text:
            label = Label(master, text=text)
            label.place(x=pb_x + pb_width, y=pb_y)
        master.update()


def fileopen():
    v.set('')#clear text
    
    file_name=askopenfilename()
    if file_name:
        v.set(file_name)
    print(v.get())#获得路径
    global select_path #寻找全局变量去修改
    select_path = str(v.get())#转str
    print(select_path)
    global folder_path
    folder_path=os.path.dirname(select_path)
    global pdf_name
    pdf_name=os.path.basename(select_path).split('.')[0]
        
def run():
    import threading
    root = Tk()
    root.title("progressBar")
    root.geometry("400x50+750+450")
    print('开始转换')
   
    progress.init(self=progress, master=root, x=10, y=10)
    #更新进度条
    progress.run(self=progress, master=root, percentage=10, text="转换进度："+str(10)+"%")
    
    thread = threading.Thread(target=main_thread(root=root))
    #守护线程
    thread.setDaemon(True)
    thread.start()
    root.mainloop()
    

def main_thread(root):
    
    #更新进度条
    progress.run(self=progress, master=root, percentage=35, text="转换进度："+str(35)+"%")
    
    
    #提取文字
    from io import StringIO
    from pdfminer.high_level import extract_text_to_fp
    from pdfminer.layout import LAParams
    output_string = StringIO()
    with open(select_path, 'rb') as fin:
         extract_text_to_fp(fin, output_string, laparams=LAParams(),
                        output_type='html', codec=None,output_dir=folder_path+'/tempImage',debug=False)
    
    #更新进度条
    progress.run(self=progress, master=root, percentage=50, text="转换进度："+str(50)+"%")
    

    img_path=folder_path+'/tempImage'
    
    #写入到html
    file=open(folder_path+'/test.html','w',encoding='utf-8')
    html_path=folder_path+'/test.html'
    file.write(output_string.getvalue().strip())
    file.close()
    
   
    
    #获取图片列表的三种属性
    import os
    def listdir(path, list_src):  # 传入存储的list
        for file in os.listdir(path): 
            file_name=os.path.join(file)
            file_path = os.path.join(path, file)
            file_builttime=os.path.getmtime(file_path)
            #windows下产生右斜杠，与python不兼容
            file_path = file_path.replace('\\', '/')
            if os.path.isdir(file_path):
                listdir(file_path, list_src)
            else:
                list_src.append(file_path)
                list_name.append(file_name)
                list_builttime.append(file_builttime)
     #更新进度条
    progress.run(self=progress, master=root, percentage=75, text="转换进度："+str(75)+"%")

    list_name=[]
    list_builttime=[]
    list_src=[]
    series=[]
    path=img_path   #文件夹路径
    listdir(path,list_src)

    for i in range(len(list_name)):
        item=[]
        item.append(list_name[i])
        item.append(list_builttime[i])
        item.append(list_src[i])
        series.append(item)


    #加入dataframe按pdf中的顺序排列
    import pandas as pd
    column=['name','builtDate','src']
    df=pd.DataFrame(series,columns=column)
    #默认升序排列，从小到大
    df=df.sort_values(by='builtDate')
    df=df.reset_index(drop=True)
    
    #更新进度条
    progress.run(self=progress, master=root, percentage=80, text="转换进度："+str(80)+"%")
    
    #写入img src到html
    #a是append,w是write
    file=open(html_path,'a',encoding='utf-8')
    for i in range(len(df)):
        src=df.loc[i,'src']
        file.write("<img src='%s'>" % src ) 
    file.close()
    
    #更新进度条
    progress.run(self=progress, master=root, percentage=90, text="转换进度："+str(90)+"%")

    import pypandoc
    #html转doc
    # convert_file('原文件','目标格式','目标文件')
    output = pypandoc.convert_file(html_path, 'docx', outputfile=folder_path+"/"+pdf_name+".docx")

    
    #删除暂时的html和img
    import os

    import shutil
    if os.path.exists(html_path):
        os.remove(html_path) 
    if os.path.exists(img_path):

        shutil.rmtree(path) 
    
    #更新进度条
    progress.run(self=progress, master=root, percentage=100, text="转换进度："+str(100)+"%")
    root.destroy()
   
   
frameT=Tk()
frameT.geometry('500x100+700+400')
frameT.title('选择需要转换的pdf文件')
frame=Frame(frameT)
frame.pack(padx=10,pady=10)
frame1=Frame(frameT)
frame1.pack(padx=10,pady=10)
v=StringVar()
select_path=''
folder_path=''
pdf_name=''

ent=Entry(frame,width=50,textvariable=v).pack(fill=X,side=LEFT)
btn= Button(frame,width=20,text='选择文件',font=("宋体",14),command=fileopen).pack(fill=X,padx=10)
ext= Button(frame1,width=10,text='运行',font=("宋体",14),command=run).pack(fill=X,side=LEFT)
etb= Button(frame1,width=10,text='退出',font=("宋体",14),command=frameT.destroy).pack(fill=Y,padx=10)

frameT.mainloop()
print(pdf_name)


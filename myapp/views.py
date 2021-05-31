import enum
from django.shortcuts import render,HttpResponse
from tap import settings
from win32com.client import Dispatch
import numpy
import pandas
import re
import pythoncom
import os
import os.path
from PIL import Image
from aip import AipOcr


#download
from django.utils.encoding import escape_uri_path
from django.http import StreamingHttpResponse
from django.http import JsonResponse
from rest_framework import status

# Create your views here.

###保存到服务端excel，下载到本地excel
def download(n,sixth):
    pythoncom.CoInitialize()
    word = Dispatch('Word.Application')     # 打开word应用程序
    word.Visible = 0        # 后台运行,不显示
    word.DisplayAlerts = 0  # 不警告

    rtf_path=os.getcwd().replace('\\','/')+'/static/rtf'
    excel_path=os.getcwd().replace('\\','/')+'/static/excel'

    
    path = r'{}/{}'.format(rtf_path,n) # 写绝对路径，相对路径会拨错
    
    doc = word.Documents.Open(FileName=path, Encoding='gbk')

    f=open('{}/1.txt'.format(rtf_path),'w',encoding='gbk')
    for para in doc.paragraphs:
        
        f.write(para.Range.Text)
        # print(para.Range.Text)

    doc.Close()
    word.Quit()
    f.close()

    f=open('{}/1.txt'.format(rtf_path),'r',encoding='gbk')
    text=f.readlines()
    f.close()
    for index,value in enumerate(text):
        if '各脏器生物活性状态' in value:
            start=index
        if '健康干预方案' in value:
            end=index

    # print(start,end)
    # print(text[start:end])
    res=[re.sub('\s','',x) for x in text[start+3:end-3]]
    num,name=[],[]
    # print(res)
    for x in res:
        num.append(x[x.find('[')+1:x.find(']')])
        name.append(x[x.find(']')+1:])

    # print(len(num),len(name))
    total_list=[]
    for index,value in enumerate(num):
        total_list.append([name[index],value])

    file_name=n[:-4]+'.xlsx'

    writer=pandas.ExcelWriter("{}/{}".format(excel_path,file_name))
    data=numpy.array(total_list)
    df=pandas.DataFrame(data)
    df2=pandas.DataFrame(tap(text)[0])
    df3=pandas.DataFrame(tap(text)[1])
    df4=pandas.DataFrame(tap(text)[2])
    df5=pandas.DataFrame(tap(text)[3])
    df6=pandas.DataFrame(tap(text)[4])
    df7=pandas.DataFrame(sixth)
    df.to_excel(writer,sheet_name='各脏器生物活性状态',header=False,index=False)
    df2.to_excel(writer,sheet_name='间质的离子分析',header=False,index=False)
    df3.to_excel(writer,sheet_name='酸碱平衡',header=False,index=False)
    df4.to_excel(writer,sheet_name='神经递质',header=False,index=False)
    df5.to_excel(writer,sheet_name='激素水平',header=False,index=False)
    df6.to_excel(writer,sheet_name='生化相对指标',header=False,index=False)
    df7.to_excel(writer,sheet_name='各系统风险值',header=False,index=False)

    writer.save()

    download_file_path=excel_path+'/'+file_name
    response = big_file_download(download_file_path,file_name)
    
    if response:
        return response

### 改变图片尺寸
def ResizeImage(path2):
    filein = path2
    fileout = path2
    width = 500
    height = 900
    img = Image.open(filein)
    out = img.resize((width, height),Image.ANTIALIAS)
    out.save(fileout)
    img.close()

### 读取图片指标
def get_file_content(filepath):
    APP_ID = '24269656'
    API_KEY = '5oXOngw1HxCdZqsDKkWDhB3M'
    SECRET_KEY = 'WGIKYNOkAB3x1WdR4qYVENpNEo9TLRQa'
    
    client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
    with open(filepath, 'rb') as fp:
        image = fp.read()
    fp.close()
     # 定义参数变量
    options = {
        # 定义图像方向
        'detect-direction': 'true',
        'language-type': 'CHN_ENG'
    }
    result = client.general(image, options)
    _list,num,total_list=[],[],[]
    for word in result['words_result']:
        _list.append(word['words'])

    for x in _list:
        try:
            num.append(re.findall('\d+%',x)[0])
        except:
            pass
    name=['呼吸系统的风险','消化系统的风险','免疫系统的风险','变性疾病危险','泌尿生殖器和肾脏的风险',
            '骨骼以及神经筋的风险','心血管系统的风险','内分泌系统','神经功能','氧化压力','过敏的风险',
            '潜在的情况','感染的风险','皮肤疾病的风险','耳鼻喉的风险'
    ]
    
    for index,value in enumerate(name):
        total_list.append([value,num[index]])
    sixth=numpy.array(total_list)
    return sixth

###将前端传输过来的文件保存在服务端
def save_file(f,img,path,path2):
    with open(path,'wb') as fp:
        if f.multiple_chunks: #判断到上传文件为大于2.5MB的大文件
            for buf in f.chunks(): #迭代写入文件
                fp.write(buf)
        else:
            fp.write(f.read())
    fp.close()

    with open(path2,'wb') as fp:
        if img.multiple_chunks: #判断到上传文件为大于2.5MB的大文件
            for buf in img.chunks(): #迭代写入文件
                fp.write(buf)
        else:
            fp.write(img.read())
    fp.close()


###视图函数
def index(request):
    if request.method == "POST":
        f = request.FILES.get("upload_file")
        img = request.FILES.get("img")
        try:
            if f.name[-3:]=='rtf' and img.name[-3:] in ['jpg','png']:
                path = os.path.join(settings.STATICFILES_DIRS[0],'rtf/'+f.name)
                path2 = os.path.join(settings.STATICFILES_DIRS[0],'img/'+img.name)
                
                save_file(f,img,path,path2)
                ResizeImage(path2)
                sixth=get_file_content(path2)
                res=download(f.name,sixth)
                return res
            else:
                return HttpResponse('请选择正确的文件')
        except:
            return HttpResponse('请选择正确的文件')
    return render(request, 'index.html',locals())


###下载本地文件生成器
def file_iterator(file_path, chunk_size=512):
    with open(file_path, mode='rb') as f:
        while True:
            c = f.read(chunk_size)
            if c:
                yield c
            else:
                break

###返回下载本地文件响应
def big_file_download(download_file_path, filename):
    try:
        response = StreamingHttpResponse(file_iterator(download_file_path))
        # 增加headers
        response['Content-Type'] = 'application/octet-stream'
        response['Access-Control-Expose-Headers'] = "Content-Disposition, Content-Type"
        response['Content-Disposition'] = "attachment; filename={}".format(escape_uri_path(filename))
        return response
    except Exception:
        return JsonResponse({'status': status.HTTP_400_BAD_REQUEST, 'msg': 'Excel下载失败'},
                            status=status.HTTP_400_BAD_REQUEST)

def tap(text):
    content=text
    res=[]
    ##间质的离子分析
    for index,value in enumerate(content):
        if '间质的离子分析' in value:
            start=index+2
        if '间质的铁' in value:
            end=index+1

    first=[x.strip() for x in content[start:end]]
    total_list=[]
    for x in first:
        index=x.find(': ')
        total_list.append([x[:index],x[index+1:]])
    total_list[0][0]=total_list[0][0][1:]
    first=numpy.array(total_list)
    res.append(first)


    ##酸碱平衡
    for index,value in enumerate(content):
        if '（标准值：N对应值）' in value:
            start=index+1
        if 'iSO2' in value:
            end=index+1
    second=[x.strip() for x in content[start:end]]
    total_list=[]
    for x in second:
        index=x.find('=')
        total_list.append([x[:index],re.findall('-*\d+.\d+',x)[0]])
    total_list[0][0]=total_list[0][0][1:]
    second=numpy.array(total_list)
    res.append(second)


    ##神经递质
    for index,value in enumerate(content):
        if '间质的5-羟色胺' in value:
            start=index
        if '间质的乙酰胆碱' in value:
            end=index+1
    third=[x.strip() for x in content[start:end]]

    total_list=[]
    for x in third:
        index=x.find('=')
        total_list.append([x[:index],x[index+1:]])
    total_list[0][0]=total_list[0][0][1:]
    third=numpy.array(total_list)
    res.append(third)


    ##激素水平
    for index,value in enumerate(content):
        if '间质的促甲状腺激素' in value:
            start=index
        if '间质的促肾上腺皮质激素' in value:
            end=index+1
    fourth=[x.strip().split(',') for x in content[start:end]]
    a=[]
    for x in fourth:
        for y in x:
            a.append(y)
    total_list=[]
    for x in a:
        index=x.find('=')
        total_list.append([x[:index],x[index+1:]])
    total_list[0][0]=total_list[0][0][1:]
    fourth=numpy.array(total_list)
    res.append(fourth)


    ##生化相对指标
    for index,value in enumerate(content):
        if '间质的甘油三酯' in value:
            start=index
        if '间质的低密度脂蛋白' in value:
            end=index+1
    fifth=[x.strip() for x in content[start:end]]

    total_list=[]
    for x in fifth:
        index=x.find('=')
        total_list.append([x[:index],x[index+1:]])
    total_list[0][0]=total_list[0][0][1:]
    fifth=numpy.array(total_list)
    res.append(fifth)

    return res
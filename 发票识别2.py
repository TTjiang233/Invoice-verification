import urllib
import base64
import json
import urllib.parse
import urllib.request
import xlwt
import pandas as pd
#from pdf2image import convert_from_path
import os.path
#from aip import AipOcr
#from PIL import Image as PI

import datetime
import os

import fitz  # fitz就是pip install PyMuPDF

client_id ='OBh8iy35TWzwZ8V6wT6szQbe'
client_secret ='8dVCTI8vtT4OgEfpcvfnWKywtW7I5jBm'

def get_token():
    host = 'https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=' + client_id + '&client_secret=' + client_secret
    request = urllib.request.Request(host)
    request.add_header('Content-Type', 'application/json; charset=UTF-8')
    response = urllib.request.urlopen(request)
    token_content = response.read()
    if token_content:
        token_info = json.loads(token_content)
        token_key = token_info['access_token']
    return token_key

def pyMuPDF_fitz(pdfPath, imagePath,n):
    
    
    startTime_pdf2img = datetime.datetime.now()  # 开始时间

    print("imagePath=" + imagePath)
    pdfDoc = fitz.open(pdfPath)
    for pg in range(pdfDoc.pageCount):
        page = pdfDoc[pg]
        rotate = int(0)
        zoom_x = 1.33333333  # (1.33333333-->1056x816)   (2-->1584x1224)
        zoom_y = 1.33333333
        mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
        pix = page.getPixmap(matrix=mat, alpha=False)

        if not os.path.exists(imagePath):  # 判断存放图片的文件夹是否存在
            os.makedirs(imagePath)  # 若图片文件夹不存在就创建

        pix.writePNG(imagePath + f'\images_{n}.png')  # 将图片写入指定的文件夹内
        return [imagePath + f'\images_{n}.png',n]


    
def vat_invoice(filename):
    #filename =r"C:\Users\lenovo\Desktop\点点智联面试题\发票查验\images_0.png"
    
    request_url ="https://aip.baidubce.com/rest/2.0/ocr/v1/vat_invoice"
    #二进制打开图片
    f = open(filename[0], 'rb')
    img = base64.b64encode(f.read())

    params = dict()
    params['image'] = img
    params['show'] = 'true'
    params = urllib.parse.urlencode(params).encode("utf-8")
    #params = json.dumps(params).encode('utf-8')

    access_token = get_token()
    request_url = request_url + "?access_token=" + access_token
    request = urllib.request.Request(url=request_url, data=params)
    request.add_header('Content-Type', 'application/x-www-form-urlencoded')
    response = urllib.request.urlopen(request)
    content = response.read()
    if content:
        #print(content)
        content=content.decode('utf-8')
        #print(content)
        data = json.loads(content)
        #print(data)
        words_result=data['words_result']
        
        #f.save(r"C:\Users\lenovo\Desktop\点点智联面试题\发票查验\发票结果.xls") #保存文件
        #list2=[words_result['InvoiceNum'],words_result['AmountInWords'],words_result['NoteDrawer'],words_result['llerAddress'],words_result['SellerRegisterNum'],words_result['SellerBank'],	words_result['CheckCode'],words_result['InvoiceNum'],words_result['InvoiceDate'],	words_result['PurchaserRegisterNum'],words_result['PurchaserBank'],words_result['InvoiceTypeOrg'],words_result['PurchaserName'],	words_result['PurchaserBank'],words_result['TotalTax'],words_result['AmountInFiguers']]
        d={
               '发票号':words_result['InvoiceNum'],
               '大写总金额':words_result['AmountInWords'],
               '开票人':words_result['NoteDrawer'],
               '销售方地址':words_result['SellerAddress'],
               '销售方纳税人识别号':words_result['SellerRegisterNum'],
               '销售方开户行账号':words_result['SellerBank'],
               '校验码':words_result['CheckCode'],
               '发票代码':words_result['InvoiceNum'],
               '发票日期':words_result['InvoiceDate'],
               '购买纳税人识别号':words_result['PurchaserRegisterNum'],
               '购买方开户行':words_result['PurchaserBank'],
               '发票类型':words_result['InvoiceTypeOrg'],
               '购买方名称':words_result['PurchaserName'],
               '购买方开户账号':words_result['PurchaserBank'],
               '税额':words_result['TotalTax'],
               '价税总计':words_result['AmountInFiguers'],
               '图片地址':filename[0] }
        df=pd.DataFrame(d,index=[0])
        print(df)

                


if __name__ == "__main__":
    # 1、PDF地址
    paths=r'C:\Users\lenovo\Desktop\点点智联面试题\发票查验'
    """f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok=True) #创建sheet
    list1=['发票号',	'大写总金额	','开票人','销售方地址',	'销售方纳税人识别号',	'销售方开户行账号','校验码',	'发票代码','发票日期','购买纳税人识别号'	,'购买方开户行','发票类型','购买方名称' ,'购买方开户账号','税额','价税总计','图片地址']
    for i in range(len(list1)):
        sheet1.write(0,i,list1[i])"""
    n=0
    for each in os.listdir(paths):
        if 'pdf' in each:
            pdfPath = paths + f'\{each}'
            # 2、需要储存图片的目录
            imagePath = r'C:\Users\lenovo\Desktop\点点智联面试题\发票查验'
            pyMuPDF_fitz(pdfPath, imagePath,n)
            vat_invoice(pyMuPDF_fitz(pdfPath, imagePath,n))
            n=n+1
    #f.save(r"C:\Users\lenovo\Desktop\点点智联面试题\发票查验\发票结果.xls") #保存文件
 #   pdfPath = r'C:\Users\lenovo\Desktop\点点智联面试题\发票查验\滴滴电子发票A.pdf'
        
        
    


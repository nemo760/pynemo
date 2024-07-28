#coding=utf-8
import pandas as pd
import openpyxl
import shutil
import random
import string

def randomStr(lenth):
    # 创建一个包含所有数字和大小写字母的字符串
    characters = string.ascii_letters + string.digits
    random_string = ''.join(random.choice(characters) for _ in range(lenth))
    return random_string

def newFilePath(Batch,Name):
    return "C:/test/%s_%s.zip"%(Batch,Name)

def CopyZipFile(new_file_path):
    # 源zip文件路径
    src_file_path = 'C:/source/BATCH01_extmid.zip'
    # 使用shutil模块的copy2函数复制文件，并保持源文件数据不变
    shutil.copy2(src_file_path, new_file_path)
def dataPort(extmid,CardNum):
    dataList=[]
    ExternalMerchantID =extmid
    dataList.append(ExternalMerchantID)
    MerchantType = 'Enterprise'
    dataList.append(MerchantType)
    FullName = 'Test_FullName'
    dataList.append(FullName)
    ShortName = 'Test_ShortName'
    dataList.append(ShortName)
    Province = 12
    dataList.append(Province)
    City = 385
    dataList.append(City)
    District = 5415
    dataList.append(District)
    Postcode = 70612
    dataList.append(Postcode)
    Address = 'GSHDGDHJS GSHDGDHJSGSHDGDHJSGSHDGDHJSGSHDGDHJSGSHDGDHJSGSHDGDHJS  GSHDGDHJSGSHDGDHJSGSHDGDHJSGSHDGDHJS dada'
    dataList.append(Address)
    KTP = '8738746362888733'
    dataList.append(KTP)
    ContactName = 'Test_ContactName'
    dataList.append(ContactName)
    ContactNumber = '834324747'
    dataList.append(ContactName)
    Email = 'nemomy@163.com'
    dataList.append(Email)
    MCC = '0763'
    dataList.append(MCC)
    NMID = ''
    dataList.append(NMID)
    TerminalNumber = 'TA001'
    dataList.append(TerminalNumber)
    Criteria = '02'
    dataList.append(Criteria)
    AccountName = 'WENWEN'
    dataList.append(AccountName)
    Issuer = 'BNC'
    dataList.append(Issuer)
    CardNumber = CardNum
    dataList.append(CardNumber)
    BusinessLicensingNumber = 'BLN123456789'
    dataList.append(BusinessLicensingNumber)
    NPWPNumber = 'NPWP123456789'
    dataList.append(NPWPNumber)
    NIBNumber = 'NIB12345678'
    dataList.append(NIBNumber)
    return  dataList

def openyxlExcel(filePath,data):
    # 编写进件文件的函数，传入需要编辑的excel，准备好进件数据
    # 打开Excel文件
    workbook = openpyxl.load_workbook(filePath)
    # 选择要写入的工作表
    worksheet = workbook['Sheet1']
    # 将数据写入工作表的指定行
    worksheet.append(data)
    # 保存Excel文件
    workbook.save(filePath)

if __name__ == "__main__":
    for i in range(20):
        batch = 'BATCH01'
        extmid = randomStr(10)
        #生成新文件名
        newPath = newFilePath(batch, extmid)
        #根据新文件名复制原进件附件文件
        CopyZipFile(newPath)
        #生成进件参数，写入进件excel文件
        data = dataPort(extmid, '5859457100780470')
        filePath = 'C:/test/QRIS_000590000545_20240112_SUBMERCHANT_BATCH01.xlsx'
        openyxlExcel(filePath, data)
    print('-----------------the end ---------------------')
    print('----------20240728-01---------------')
    print('----------20240728-02---------------')
    print('----------20240728-03---------------')

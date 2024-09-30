from datetime import datetime
import os,sys
import configparser
import win32com.client
def logExp(ex):
    strLogPath = getRelativePath("log")
    strTime =  datetime.now().strftime("%Y%m%d")
    strFilePath = strLogPath + r"\log_" + strTime + ".txt"
    strContents = "[{}] {}".format( datetime.now().strftime("%Y%m%d %H:%M:%S"),ex)
    if not os.path.exists(strFilePath):
        with open(strFilePath,"x",encoding='utf-8') as logFile:
            logFile.writelines("\n")
            logFile.writelines(strContents)
    else:
        with open(strFilePath,"a",encoding='utf-8') as logFile:
            logFile.writelines("\n")
            logFile.writelines(strContents)   
def getRelativePath(folderPath):
    strAbsPath = os.path.abspath(sys.argv[0])
    strCrrPath = os.path.dirname(strAbsPath)
    strFilePath = os.path.join(strCrrPath,folderPath)
    if not os.path.exists(strFilePath):
        os.makedirs(strFilePath)
        return strFilePath
    else:      
        return strFilePath
def countFileInFolder(folder_name):
    sys_Path = os.path.abspath(sys.argv[0])
    base_Path = os.path.dirname(sys_Path)
    fol_path = os.path.join(base_Path,folder_name)
    all_files = os.listdir(fol_path)
    print(len(all_files))
    return len(all_files)
def chayVBA(file_path):
    

    # Đường dẫn tới file Excel của bạn
    macro_name = 'Callback2'  # Thay đổi tên macro bạn muốn chạy

    # Bước 1: Tạo đối tượng Excel
    excel = win32com.client.Dispatch('Excel.Application')

    # Bước 2: Mở workbook
    workbook = excel.Workbooks.Open(file_path)

    # Bước 3: Chạy macro
    excel.Application.Run(macro_name)

    # Bước 4: Lưu và đóng workbook
    workbook.Save()
    workbook.Close()

    # Bước 5: Đóng Excel
    excel.Quit()

    print(f"Đã chạy macro '{macro_name}' trong file '{file_path}'")
    
def readIni(strSec,strKey):
    try:
        #xu li file ini de doc dc bang configParse
        strAbsPath = os.path.abspath(sys.argv[0])
        strCurrPath = os.path.dirname(strAbsPath)    
        with open(strCurrPath + r"\config.ini", 'rb') as f:
            content = f.read()
        # Kiểm tra và loại bỏ BOM nếu tồn tại
        if content.startswith(b'\xef\xbb\xbf'):
            content = content[3:]
        # Chuyển đổi nội dung thành chuỗi
        content_str = content.decode('utf-8')  


        #Tao Doi Tuong ConfigParse
        config = configparser.ConfigParser()
        #Doc File INI
        config.read_string(content_str)
        sec = strSec
        op = strKey
        if config.has_section(sec):
            if config.has_option(sec,op):
                value = config.get(sec,op)
                return value
            else:
                logExp(f'Key {op} of Section {sec} is not Exists!')
        else:
            logExp(f'Section {sec} is not Exists!') 
    except Exception as ex:
        logExp(ex)
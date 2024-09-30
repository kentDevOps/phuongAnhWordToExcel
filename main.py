from log import *
from wordProcess import *
import time

def mainPro():
    try:
        start_time = time.time()
        #Kiểm Tra và tạo folder doc
        docPath = getRelativePath("doc")
        #Kiểm Tra tạo folder report
        reportPath = getRelativePath("report")
        #count Doc File
        count_doc = countFileInFolder("doc")
        if count_doc == 0:
            print(f'File Doc Không Tồn Tại , Hãy Copy Vào Folder doc')
            return    
        #Duyệt qua các file trong doc 
        arrDoc = os.listdir(docPath)
        dem = 1
        for item in arrDoc:
            sample_path = os.path.join(docPath,item)
            # Lọc dữ liệu cần fill vào excel
            cauho = filter_congchuc_lines(sample_path)
            daA = filter_DapAnA_lines(sample_path,'A.')
            daB = filter_DapAnB_lines(sample_path,'B.')
            daC = filter_DapAnC_lines(sample_path,'C.')
            daD = filter_DapAnD_lines(sample_path,'D.')
            ketQua = extract_correct_answers(sample_path)
            giaiThich = find_explanation_index(sample_path)
            dem+=1
            strAbsPath = os.path.abspath(sys.argv[0])
            strCrrPath = os.path.dirname(strAbsPath)
            vbaFile = os.path.join(strCrrPath,"rp.xlsx")            
            append_to_excel(cauho,daA,daB,daC,daD,ketQua,giaiThich,vbaFile,os.path.join(reportPath,item.split('.')[0] + '.xlsx'),dem)
            #text_without_accents = unidecode(item.split('.')[0])
            '''strAbsPath = os.path.abspath(sys.argv[0])
            strCrrPath = os.path.dirname(strAbsPath)
            vbaFile = os.path.join(strCrrPath,"WordReport.xlsm")
            chayVBA(vbaFile)'''
            #boido(os.path.join(reportPath,item.split('.')[0] + '.xlsx'))
        end_time = time.time()
        print('------------------------------------------Kết Quả Xử Lí-----------------------------------------')
        print(f'Thời gian thực thi: {end_time - start_time} giây')     
        print('------------------------------------------------End---------------------------------------------')
    except Exception as ex:
        logExp(str(ex))

#check xem có phải Hàm main không và show form
if __name__ == "__main__":
    mainPro()
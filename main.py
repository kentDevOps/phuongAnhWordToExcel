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
        end_time = time.time()
        print(f'Thời gian thực thi: {end_time - start_time} giây')   
    except Exception as ex:
        logExp(str(ex))

#check xem có phải Hàm main không và show form
if __name__ == "__main__":
    mainPro()
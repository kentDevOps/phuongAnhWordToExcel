from docx import Document
import re
'''def read_word_file(file_path):
    # Mở và đọc file Word
    doc = Document(file_path)
    full_text = ""

    # Lặp qua tất cả các đoạn văn trong tài liệu và chuyển thành văn bản
    for para in doc.paragraphs:
        full_text += para.text + "\n"
    
    return full_text

def find_lines_containing_keywords(text, keywords):
    # Tìm các dòng chứa từ khóa
    lines = text.splitlines()
    matched_lines = [line for line in lines if any(keyword in line for keyword in keywords)]
    
    return matched_lines

# Đường dẫn đến file Word của bạn
file_path = "DONE.docx"  # Thay thế bằng đường dẫn tới file của bạn

# Đọc toàn bộ nội dung của file Word
text = read_word_file(file_path)

# Tìm các dòng chứa từ khóa "Định Nghĩa" và "[CÔNG CHỨC"
keywords = ["Định Nghĩa", "[CÔNG CHỨC"]
matched_lines = find_lines_containing_keywords(text, keywords)

# In các dòng tìm thấy
for line in matched_lines:
    print(line)'''
def contains_roman_numerals(text):
    # Regular expression to match Roman numerals
    roman_pattern = r'\b[IVXLCDM]+\.'
    return re.search(roman_pattern, text) is not None
txt = ' Mark the letter A, B, C, or D on your answer sheet to indicate the correct answer to each of the following questions.'
a = contains_roman_numerals(txt)
print(a)

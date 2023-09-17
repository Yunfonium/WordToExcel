from docx import Document
import pandas as pd

def docxToXls(file_path):
    f = open(file_path,'rb') 
    document = Document(f)
    f.close()
    df = pd.DataFrame(columns=['姓','名','性別','身分證','戶籍地'])

    for p in document.paragraphs:
        if p.text == '':
            continue
        text_list = p.text.split(',')
        name_list = text_list[0].replace(' ',',').split(',')
        last_name = name_list[0]
        first_name = name_list[1]
        id = text_list[1]
        df = df.append({'姓': first_name, '名':last_name, '身分證':id}, ignore_index= True)
    print('DataFrame is created successfully, transform it to xlsx.')
    df.to_excel('./excel_data/tmp.xlsx',index=False)

if __name__ == '__main__':
    print('Start to transform docx to xlsx')
    docxToXls('./word_data/身分資料文件.docx')
    print('Finish!')
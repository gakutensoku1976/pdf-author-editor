import sys
import pypdf
import tkinter as tk
from tkinter import filedialog

def edit_pdf_author(input_file_path, author):
    output_file_path = input_file_path.replace('.pdf', '_edited.pdf')
    src_pdf = pypdf.PdfReader(input_file_path)
    dst_pdf = pypdf.PdfWriter()
    dst_pdf.clone_reader_document_root(src_pdf)

    d = {key: src_pdf.metadata[key] for key in src_pdf.metadata.keys()}

    if author is not None:
        d['/Author'] = author
    elif '/Author' in d:
        del d['/Author']

    dst_pdf.add_metadata(d)
    dst_pdf.write(output_file_path)

    return output_file_path

def get_pdf_author(file_path):
    metadata = pypdf.PdfReader(file_path).metadata
    if '/Author' in metadata:
        return metadata['/Author']
    else:
        return None

def edit_pdf(input_pdf_path, new_author):
    old_author = get_pdf_author(input_pdf_path)

    if old_author is None:
        print(f'入力ファイル {input_pdf_path} には作成者情報が含まれていません。')
    else:
        print(f'入力ファイル {input_pdf_path} の作成者情報は {old_author} です。')

    print(f'作成者情報を {new_author} に変更しますか？(y/n)')

    if input() == 'y':
        output_pdf_path = edit_pdf_author(input_pdf_path, new_author)
        check = get_pdf_author(output_pdf_path)
        print(f'【処理結果】')
        print(f'出力ファイル {output_pdf_path} の作成者情報は {check} です。')

    print('-' * 80)

def delete_pdf(input_pdf_path):
    old_author = get_pdf_author(input_pdf_path)

    if old_author is None:
        print(f'入力ファイル {input_pdf_path} には作成者情報が含まれていないため、削除する必要はありません。')
    else:
        print(f'入力ファイル {input_pdf_path} の作成者情報は {old_author} です。')
        print(f'作成者情報を削除しますか？(y/n)')

        if input() == 'y':
            output_pdf_path = edit_pdf_author(input_pdf_path, None)

            check = get_pdf_author(output_pdf_path)
            
            print(f'【処理結果】')
            if check is None:
                print(f'出力ファイル {output_pdf_path} には作成者情報が含まれていません。')
                print(f'作成者情報の削除に成功しました。')

            else:
                print(f'出力ファイル {output_pdf_path} に作成者情報 {check} が含まれています。')
                print(f'作成者情報の削除に失敗しました。')

    print('-' * 80)


def main():

    print('-' * 80)

    if len(sys.argv) == 1:

        root = tk.Tk()
        root.withdraw()

        fTyp = [('PDFファイル', '*.pdf')]
        excelFileNames = filedialog.askopenfilename(filetypes=fTyp, multiple=True)

        for excelFileName in excelFileNames:
            delete_pdf(excelFileName)

    elif len(sys.argv) == 2:

        delete_pdf(sys.argv[1])

    elif len(sys.argv) == 3:

        edit_pdf(sys.argv[1], sys.argv[2])  

    else:
        print('【使用方法】')
        print('\t入力PDFファイル名と新しい作成者名を指定してください')
        print('\t新しい作成者名を指定しない場合は、作成者情報を削除します')
    
    sys.exit(0)

if __name__ == '__main__':
    main()

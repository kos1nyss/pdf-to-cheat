from converter import Converter

conv = Converter()

# 1) Вариант с одним файлом, pdf превращается в docx
conv.execute('one.docx', pdf_path='pdf/example.pdf')

# 2) Вариант с папкой, берутся все файлы из папки и собираются в один docx
conv.execute('many.docx', folder_with_pdfs_path='/pdf')

from docx import Document

def replace_asterisks_in_docx(path, x):
    # 打开输入的 docx 文件
    doc = Document(path)
    # 将段落中的所有 'x' 替换为空字符
    for paragraph in doc.paragraphs:        
        paragraph.text = paragraph.text.replace(x, '')

    # 保存修改后的文档
    doc.save(path)
    
def escape_backslashes(path):
    # 转义反斜杠
    path = path.replace('\\', '\\\\')
    path = path.replace('"', '')
    return path     

def main():
    path = input("请输入要处理的docx文件路径： ")
    path = escape_backslashes(path)
    print("输入Q或q以结束程序...")
    while True:
        x = input("请输入要删除的字符： ")
        if x == 'Q' or x == 'q':
            print("程序结束")
            break
        else:
            replace_asterisks_in_docx(path, x)
            print('Done!')

if __name__ == "__main__":
    main()
from docx import Document
from docx.oxml import OxmlElement, ns

def convert_path(input_path):
    return input_path.replace('/', '\\')

def config_table_properties(table):
    tbl_pr = table._element.xpath('w:tblPr')[0]

    # Thiết lập chiều rộng 100% cho bảng
    tbl_width = OxmlElement('w:tblW')
    tbl_width.set(ns.qn('w:w'), '5000')
    tbl_width.set(ns.qn('w:type'), 'pct')
    tbl_pr.append(tbl_width)

    # Căn chỉnh bảng ở giữa
    jc = OxmlElement('w:jc')
    jc.set(ns.qn('w:val'), 'center')
    tbl_pr.append(jc)

    # Thiết lập bọc văn bản là "None"
    tbl_wrap = OxmlElement('w:tblpPr')
    tbl_wrap.set(ns.qn('w:leftFromText'), '0')
    tbl_wrap.set(ns.qn('w:rightFromText'), '0')
    tbl_wrap.set(ns.qn('w:topFromText'), '0')
    tbl_wrap.set(ns.qn('w:bottomFromText'), '0')
    tbl_pr.append(tbl_wrap)

    # Thiết lập khoảng cách từ bên trái (Indent from left)
    indent = OxmlElement('w:tblInd')
    indent.set(ns.qn('w:w'), '0')
    indent.set(ns.qn('w:type'), 'dxa')
    tbl_pr.append(indent)

    # Thiết lập left và right cell margins là 1.5 mm
    cell_mar_left = OxmlElement('w:left')
    cell_mar_left.set(ns.qn('w:w'), '85.2')
    cell_mar_left.set(ns.qn('w:type'), 'dxa')

    cell_mar_right = OxmlElement('w:right')
    cell_mar_right.set(ns.qn('w:w'), '85.2')
    cell_mar_right.set(ns.qn('w:type'), 'dxa')

    cell_margins = OxmlElement('w:tblCellMar')
    cell_margins.append(cell_mar_left)
    cell_margins.append(cell_mar_right)

    tbl_pr.append(cell_margins)


    # Thiết lập đường viền cho bảng với độ rộng 1 pt
    tbl_border = OxmlElement('w:tblBorders')
    tbl_border.set(ns.qn('w:val'), 'single')  # Kiểu đường viền đơn
    tbl_border.set(ns.qn('w:sz'), '4')  # Độ rộng 1 pt (4 là đơn vị dxa tương ứng với 1 pt)
    tbl_border.set(ns.qn('w:space'), '0')
    tbl_border.set(ns.qn('w:color'), 'auto')

    tbl_pr.append(tbl_border)

def main():
    # Đường dẫn cố định tới tài liệu Word
    input_path = r"C:\Users\HaoNguyen\Desktop\MySkill\English\CodingRules\Code Rule.docx"
    path = convert_path(input_path)

    # Mở tài liệu Word đã có sẵn
    doc = Document(input_path)

    # Lặp qua các bảng trong tài liệu
    for table in doc.tables:
        config_table_properties(table)  # Thiết lập các thuộc tính cho mỗi bảng

    # Lưu tài liệu với các thay đổi
    doc.save(path)

if __name__ == "__main__":
    main()

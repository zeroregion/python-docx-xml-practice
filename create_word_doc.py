from xml.etree.ElementTree import Element, SubElement, tostring, register_namespace
from xml.dom import minidom
import os
import zipfile

# 注册命名空间
register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
register_namespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
register_namespace('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture')


def create_content_types():
    """创建[Content_Types].xml文件"""
    types = Element('Types')
    types.set('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types')

    # 添加默认类型
    default1 = SubElement(types, 'Default')
    default1.set('Extension', 'rels')
    default1.set('ContentType', 'application/vnd.openxmlformats-package.relationships+xml')

    default2 = SubElement(types, 'Default')
    default2.set('Extension', 'xml')
    default2.set('ContentType', 'application/xml')

    # 添加覆盖类型
    override1 = SubElement(types, 'Override')
    override1.set('PartName', '/word/document.xml')
    override1.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')

    # 添加页眉页脚类型
    override2 = SubElement(types, 'Override')
    override2.set('PartName', '/word/header1.xml')
    override2.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')

    override3 = SubElement(types, 'Override')
    override3.set('PartName', '/word/footer1.xml')
    override3.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')

    override4 = SubElement(types, 'Override')
    override4.set('PartName', '/word/header2.xml')
    override4.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml')

    override5 = SubElement(types, 'Override')
    override5.set('PartName', '/word/footer2.xml')
    override5.set('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')

    return types


def create_settings():
    """创建settings.xml文件"""
    settings = Element('w:settings')
    settings.set('xmlns:w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

    # 添加基本设置
    zoom = SubElement(settings, 'w:zoom')
    zoom.set('w:percent', '100')

    return settings


def create_header():
    """创建页眉"""
    header = Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hdr')

    p = SubElement(header, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    pPr = SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    pStyle = SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
    pStyle.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'Header')

    # 添加居中对齐
    jc = SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
    jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'center')

    r = SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    t = SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    t.text = '文档标题'

    return header


def create_footer(start_page_number=1):
    """创建页脚（带页码）
    Args:
        start_page_number: 起始页码
    """
    footer = Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ftr')

    p = SubElement(footer, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    pPr = SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    pStyle = SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
    pStyle.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'Footer')

    # 添加居中对齐
    jc = SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
    jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'center')

    r = SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    t = SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    t.text = '第 '

    # 添加页码字段
    fldChar1 = SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldChar')
    fldChar1.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType', 'begin')

    instrText = SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instrText')
    instrText.text = f'{start_page_number}'

    fldChar2 = SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldChar')
    fldChar2.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType', 'separate')

    r = SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    t = SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    t.text = str(start_page_number)

    fldChar3 = SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldChar')
    fldChar3.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType', 'end')

    r = SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    t = SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    t.text = ' 页'

    return footer


def create_styles():
    """创建styles.xml文件"""
    styles = Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styles')

    # 添加默认段落样式
    docDefaults = SubElement(styles, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}docDefaults')
    runDefaults = SubElement(docDefaults, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPrDefault')
    runProps = SubElement(runDefaults, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')

    # 添加默认字体设置
    font = SubElement(runProps, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
    font.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', 'Times New Roman')
    font.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '宋体')
    font.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi', 'Times New Roman')

    # 添加标题样式
    style = SubElement(styles, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style')
    style.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'paragraph')
    style.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId', 'Title')
    name = SubElement(style, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name')
    name.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'Title')
    runProps = SubElement(style, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    font = SubElement(runProps, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
    font.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', 'Times New Roman')
    font.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '宋体')
    font.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi', 'Times New Roman')
    size = SubElement(runProps, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
    size.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '36')
    size = SubElement(runProps, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}szCs')
    size.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '36')

    # 添加页眉样式
    headerStyle = SubElement(styles, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style')
    headerStyle.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'paragraph')
    headerStyle.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId', 'Header')
    headerName = SubElement(headerStyle, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name')
    headerName.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'Header')

    # 添加页脚样式
    footerStyle = SubElement(styles, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}style')
    footerStyle.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'paragraph')
    footerStyle.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId', 'Footer')
    footerName = SubElement(footerStyle, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name')
    footerName.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'Footer')

    return styles


def add_section_break(body, has_header_footer=False, start_page_number=None, last_page=False):
    """添加分节符"""
    p = SubElement(body, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    pPr = SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    sectPr = SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')

    if has_header_footer:
        # 添加页眉引用
        headerReference = SubElement(sectPr,
                                     '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}headerReference')
        headerReference.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'default')
        headerReference.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', 'rId3')

        # 添加页脚引用
        footerReference = SubElement(sectPr,
                                     '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footerReference')
        footerReference.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'default')
        footerReference.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', 'rId4')

        # 设置页码格式
        pgNumType = SubElement(sectPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgNumType')
        if start_page_number is not None:
            pgNumType.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}start',
                          str(start_page_number))
        pgNumType.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fmt', 'decimal')

        # 设置页眉页脚属性
        headerFooter = SubElement(sectPr,
                                  '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}headerFooter')
        headerFooter.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}differentFirst', '0')
        headerFooter.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}differentOddEven', '0')

    else:
        # 添加空的页眉页脚引用，表示不使用页眉页脚
        headerReference = SubElement(sectPr,
                                     '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}headerReference')
        headerReference.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'default')
        headerReference.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', 'rId5')

        footerReference = SubElement(sectPr,
                                     '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footerReference')
        footerReference.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'default')
        footerReference.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', 'rId6')

    # 设置页面边距
    pgMar = SubElement(sectPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgMar')
    pgMar.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}top', '1440')
    pgMar.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}right', '1440')
    pgMar.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bottom', '1440')
    pgMar.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left', '1440')
    pgMar.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}header', '720')
    pgMar.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footer', '720')
    pgMar.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}gutter', '0')

    # 设置分节符类型
    type = SubElement(sectPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type')
    type.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'nextPage')

    if last_page:
        pass

    return p


def create_cover_page(body):
    """创建封面"""
    # 创建标题段落
    title_p = SubElement(body, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    title_p.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId', 'Title')
    title_run = SubElement(title_p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    title_text = SubElement(title_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    title_text.text = '文档标题'

    # 添加空行
    for _ in range(10):
        SubElement(body, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')

    # 添加日期
    date_p = SubElement(body, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    date_run = SubElement(date_p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    date_text = SubElement(date_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    date_text.text = '2025年5月'

    print(f"和封面分割")
    add_section_break(body, has_header_footer=False, last_page=False)


def create_back_cover(body):
    """创建封底"""
    print(f"和封底分割")

    # 添加空行
    for _ in range(20):
        SubElement(body, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')

    p = SubElement(body, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    pPr = SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    sectPr = SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')

    # 设置分节符类型
    # type = SubElement(sectPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type')
    # type.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'nextPage')

    # 添加空的页眉页脚引用，表示不使用页眉页脚
    # headerReference = SubElement(sectPr,
    #                              '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}headerReference')
    # headerReference.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'default')
    # headerReference.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', 'rId5')
    #
    # footerReference = SubElement(sectPr,
    #                              '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footerReference')
    # footerReference.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'default')
    # footerReference.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', 'rId6')

    # 设置不链接到前一节
    headerFooter = SubElement(sectPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}headerFooter')
    headerFooter.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}differentFirst', '0')
    headerFooter.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}differentOddEven', '0')


def create_content_pages(body, num_pages):
    """创建指定页数的正文内容"""

    for page in range(num_pages):
        print(f"创建第{page + 1}页")
        # 添加页面标题
        title_p = SubElement(body, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
        title_run = SubElement(title_p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
        title_text = SubElement(title_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
        title_text.text = f'第{page + 1}页内容'

        # 添加一些示例内容
        for _ in range(10):
            p = SubElement(body, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
            r = SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
            t = SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
            t.text = f'这是第{page + 1}页的示例内容。'

        # 如果不是最后一页，添加分页符
        if page != num_pages - 1:
            print(f"和正文内容下一页分割")
            add_section_break(body, has_header_footer=True, start_page_number=page + 1, last_page=False)
        else:
            print(f"正文内容最后一页")
            add_section_break(body, has_header_footer=True, start_page_number=page + 1, last_page=True)


def create_word_document(num_content_pages=3):
    # 创建文档关系XML
    relationships = Element('Relationships')
    relationships.set('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships')

    # 添加文档关系
    rel1 = SubElement(relationships, 'Relationship')
    rel1.set('Id', 'rId1')
    rel1.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument')
    rel1.set('Target', 'word/document.xml')

    # 创建主文档XML
    document = Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}document')

    # 创建文档主体
    body = SubElement(document, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body')

    # 创建封面（不带页眉页脚）
    create_cover_page(body)

    # 创建指定页数的正文内容
    create_content_pages(body, num_content_pages)

    # 创建封底（不带页眉页脚的分节符）
    create_back_cover(body)

    # 创建完整的目录结构
    os.makedirs('output/_rels', exist_ok=True)
    os.makedirs('output/word', exist_ok=True)
    os.makedirs('output/word/_rels', exist_ok=True)

    # 保存所有XML文件
    with open('output/[Content_Types].xml', 'w', encoding='utf-8') as f:
        f.write(prettify(create_content_types()))

    with open('output/_rels/.rels', 'w', encoding='utf-8') as f:
        f.write(prettify(relationships))

    with open('output/word/document.xml', 'w', encoding='utf-8') as f:
        f.write(prettify(document))

    with open('output/word/settings.xml', 'w', encoding='utf-8') as f:
        f.write(prettify(create_settings()))

    with open('output/word/styles.xml', 'w', encoding='utf-8') as f:
        f.write(prettify(create_styles()))

    # 保存页眉页脚
    with open('output/word/header1.xml', 'w', encoding='utf-8') as f:
        f.write(prettify(create_header()))

    with open('output/word/footer1.xml', 'w', encoding='utf-8') as f:
        f.write(prettify(create_footer(1)))  # 从第1页开始

    # 创建空的页眉页脚文件（用于封底）
    empty_header = Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hdr')
    empty_footer = Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ftr')

    with open('output/word/header2.xml', 'w', encoding='utf-8') as f:
        f.write(prettify(empty_header))

    with open('output/word/footer2.xml', 'w', encoding='utf-8') as f:
        f.write(prettify(empty_footer))

    # 创建word/_rels目录和关系文件
    word_rels = Element('Relationships')
    word_rels.set('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships')

    rel1 = SubElement(word_rels, 'Relationship')
    rel1.set('Id', 'rId1')
    rel1.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings')
    rel1.set('Target', 'settings.xml')

    rel2 = SubElement(word_rels, 'Relationship')
    rel2.set('Id', 'rId2')
    rel2.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles')
    rel2.set('Target', 'styles.xml')

    rel3 = SubElement(word_rels, 'Relationship')
    rel3.set('Id', 'rId3')
    rel3.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header')
    rel3.set('Target', 'header1.xml')

    rel4 = SubElement(word_rels, 'Relationship')
    rel4.set('Id', 'rId4')
    rel4.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer')
    rel4.set('Target', 'footer1.xml')

    rel5 = SubElement(word_rels, 'Relationship')
    rel5.set('Id', 'rId5')
    rel5.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header')
    rel5.set('Target', 'header2.xml')

    rel6 = SubElement(word_rels, 'Relationship')
    rel6.set('Id', 'rId6')
    rel6.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer')
    rel6.set('Target', 'footer2.xml')

    with open('output/word/_rels/document.xml.rels', 'w', encoding='utf-8') as f:
        f.write(prettify(word_rels))

    # 打包成docx文件
    create_docx()


def create_docx():
    """将output目录打包成docx文件"""
    # 创建临时zip文件
    with zipfile.ZipFile('output.docx', 'w', zipfile.ZIP_DEFLATED) as zipf:
        # 遍历output目录
        for root, dirs, files in os.walk('output'):
            for file in files:
                file_path = os.path.join(root, file)
                # 计算相对路径
                arcname = os.path.relpath(file_path, 'output')
                zipf.write(file_path, arcname)

    print("已生成output.docx文件")


def prettify(elem):
    """将XML元素转换为格式化的字符串"""
    rough_string = tostring(elem, 'utf-8')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")


if __name__ == '__main__':
    # 可以在这里指定正文页数
    create_word_document(num_content_pages=3)

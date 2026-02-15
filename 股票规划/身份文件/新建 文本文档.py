from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

def create_residence_certificate(file_name="住所使用证明_吉林蓝疆.docx"):
    # 创建文档对象
    doc = Document()

    # 设置中文字体支持
    doc.styles['Normal'].font.name = u'宋体'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')

    # --- 标题 ---
    title = doc.add_paragraph('住所（经营场所）使用证明')
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title.runs[0]
    run_title.font.size = Pt(18)  # 二号/小二
    run_title.font.bold = True
    
    # --- 抬头 ---
    doc.add_paragraph('') # 空行
    header = doc.add_paragraph('大安市市场监督管理局：')
    header.runs[0].font.size = Pt(14)
    header.runs[0].font.bold = True

    # --- 正文内容 ---
    # 定义正文样式函数
    def add_content_line(text, bold=False):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5 # 1.5倍行距
        p.paragraph_format.first_line_indent = Cm(0.74) # 首行缩进
        run = p.add_run(text)
        run.font.size = Pt(12) # 小四
        run.font.name = '宋体'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        if bold:
            run.font.bold = True
        return p

    # 正文第1段
    content1 = "兹有 曹______ （身份证号：____________________________________），系 ____________________________________ 房屋的所有权人。"
    add_content_line(content1)

    # 正文第2段 (填入公司名)
    content2 = "该房屋面积约 ______ 平方米，现同意将该房屋（及附属院落）无偿提供给 "
    p2 = doc.add_paragraph()
    p2.paragraph_format.line_spacing = 1.5
    p2.paragraph_format.first_line_indent = Cm(0.74)
    run2a = p2.add_run(content2)
    run2a.font.size = Pt(12)
    
    # 加粗公司名字
    run2b = p2.add_run("吉林蓝疆生态牧业科技有限责任公司")
    run2b.font.size = Pt(12)
    run2b.font.bold = True
    run2b.font.underline = True # 加下划线突出

    run2c = p2.add_run(" 作为住所（经营场所）使用，使用期限为 ______ 年（建议填写20年）。")
    run2c.font.size = Pt(12)

    # 正文第3段
    content3 = "该房屋不存在权属纠纷，不属于非法建筑、危险建筑、被征收房屋。若因经营活动产生噪音、污染等扰民问题，或涉及拆迁征收，企业承诺自行解决并承担相应法律责任。"
    add_content_line(content3)

    content4 = "特此证明。"
    add_content_line(content4)
    
    doc.add_paragraph('') # 空行
    doc.add_paragraph('') # 空行

    # --- 签字区域 ---
    # 房东签字
    sign_p1 = doc.add_paragraph()
    sign_p1.paragraph_format.left_indent = Cm(8) # 右侧缩进
    run_s1 = sign_p1.add_run("所有权人签字：__________________")
    run_s1.font.size = Pt(12)
    
    phone_p1 = doc.add_paragraph()
    phone_p1.paragraph_format.left_indent = Cm(8)
    run_p1 = phone_p1.add_run("联系电话：__________________")
    run_p1.font.size = Pt(12)

    doc.add_paragraph('') # 空行

    # 企业签字
    sign_p2 = doc.add_paragraph()
    sign_p2.paragraph_format.left_indent = Cm(8)
    run_s2 = sign_p2.add_run("企业（申请人）签字：__________________")
    run_s2.font.size = Pt(12)
    
    phone_p2 = doc.add_paragraph()
    phone_p2.paragraph_format.left_indent = Cm(8)
    run_p2 = phone_p2.add_run("联系电话：__________________")
    run_p2.font.size = Pt(12)

    doc.add_paragraph('') # 空行
    doc.add_paragraph('') # 空行
    
    # --- 盖章区域分隔线 ---
    line = doc.add_paragraph()
    line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_line = line.add_run("-----------------------------------------------------------")

    doc.add_paragraph('') # 空行

    # --- 村委会意见 ---
    opinion_title = doc.add_paragraph('【村委会/社区 确认意见】')
    opinion_title.runs[0].font.bold = True
    opinion_title.runs[0].font.size = Pt(12)

    opinion_text = doc.add_paragraph('   情况属实。同意该房屋作为企业住所（经营场所）登记使用。')
    opinion_text.runs[0].font.size = Pt(12)

    doc.add_paragraph('') # 空行
    doc.add_paragraph('') # 空行

    # 盖章位
    stamp_p = doc.add_paragraph()
    stamp_p.paragraph_format.left_indent = Cm(9)
    run_stamp = stamp_p.add_run("单位公章：(请在此处盖章)")
    run_stamp.font.size = Pt(12)

    date_p = doc.add_paragraph()
    date_p.paragraph_format.left_indent = Cm(10)
    run_date = date_p.add_run("2026 年    月    日")
    run_date.font.size = Pt(12)

    # 保存文件
    doc.save(file_name)
    print(f"成功生成文件：{file_name}")
    print("请打开文件，在下划线处填入父亲姓名、身份证号、具体地址和面积，并打印出来。")

if __name__ == "__main__":
    create_residence_certificate()
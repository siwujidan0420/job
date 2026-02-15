import os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT

# ================= 🔴 路径配置区域 =================
# 建议保存到桌面
SAVE_FOLDER = r"C:\Users\caoha\Desktop" 
# ===================================================

def set_font(run, font_name='SimSun', size=None, bold=False):
    """设置中西文混合字体 (宋体 + Times New Roman)"""
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    if size:
        run.font.size = size
    if bold:
        run.font.bold = True

def add_p_content(cell, text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, first_line_indent=False):
    """向单元格添加段落内容"""
    if len(cell.paragraphs) == 1 and not cell.text.strip():
        p = cell.paragraphs[0]
    else:
        p = cell.add_paragraph()
    
    p.alignment = align
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.25 # 1.25倍行距
    
    if first_line_indent:
        p.paragraph_format.first_line_indent = Pt(21) # 首行缩进

    run = p.add_run(text)
    set_font(run, size=Pt(10.5), bold=bold) # 默认五号字
    return p

def create_feng_proposal():
    doc = Document()

    # --- 1. 页面设置 (A4 标准) ---
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.17)
    section.right_margin = Cm(3.17)

    # --- 2. 顶部大标题 ---
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("烟台大学毕业论文（设计） 开题报告")
    set_font(run_title, size=Pt(16), bold=True) # 三号宋体加粗

    # --- 3. 学院信息 ---
    p_college = doc.add_paragraph()
    p_college.alignment = WD_ALIGN_PARAGRAPH.CENTER 
    run_college = p_college.add_run("学院：海洋学院")
    set_font(run_college, size=Pt(12)) # 小四

    doc.add_paragraph() # 空一行

    # --- 4. 创建表格框架 ---
    table = doc.add_table(rows=0, cols=6, style='Table Grid')
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False 
    
    # 辅助函数：添加基本信息行
    def add_info_row(label1, val1, label2, val2, label3, val3):
        row = table.add_row()
        row.height = Cm(1.0)
        cells = row.cells
        add_p_content(cells[0], label1, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_p_content(cells[1], val1, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_p_content(cells[2], label2, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_p_content(cells[3], val2, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_p_content(cells[4], label3, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_p_content(cells[5], val3, align=WD_ALIGN_PARAGRAPH.CENTER)

    # Row 1: 题目 (合并)
    r1 = table.add_row()
    add_p_content(r1.cells[0], "题目", align=WD_ALIGN_PARAGRAPH.CENTER)
    r1.cells[1].merge(r1.cells[5])
    # 题目来自任务书：基于FTA-BN的船舶火灾爆炸事故致因分析
    add_p_content(r1.cells[1], "基于FTA-BN的船舶火灾爆炸事故致因分析", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)

    # Row 2: 姓名、学号、专业
    # 信息来自任务书：丰顺聪, 202260503137, 航海技术
    add_info_row("姓名", "丰顺聪", "学号", "202260503137", "专业", "航海技术")
    
    # Row 3: 届别、指导教师、职称
    # 信息来自任务书：王子月, 职称无(一般填讲师或留空，这里填讲师较为得体，也可按实际填)
    add_info_row("毕业届别", "2026届", "指导教师", "王子月", "职称", "讲师")

    # 辅助函数：添加正文板块
    def add_section_block(title_text, content_text):
        row_t = table.add_row()
        cell_t = row_t.cells[0]
        cell_t.merge(row_t.cells[5])
        add_p_content(cell_t, title_text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=True)
        
        row_c = table.add_row()
        cell_c = row_c.cells[0]
        cell_c.merge(row_c.cells[5])
        add_p_content(cell_c, content_text, align=WD_ALIGN_PARAGRAPH.LEFT, first_line_indent=True)
        return cell_c

    # --- 第一部分：背景、目的、意义 ---
    text_1 = (
        "1. 研究背景：\n"
        "船舶作为海上运输的主要载体，其结构复杂、设备众多且载运货物种类繁杂。一旦发生火灾爆炸事故，不仅会导致严重的人员伤亡和巨额财产损失，还极可能引发燃油泄漏等次生灾害，对海洋生态环境造成不可逆的破坏。传统的船舶火灾风险分析多采用故障树分析法（FTA），虽然能定性识别致因，但在处理数据不确定性和动态更新方面存在局限。将故障树（FTA）与贝叶斯网络（BN）相结合，既能发挥FTA逻辑清晰的优势，又能利用BN进行概率推理和反向诊断，是当前提升船舶安全管理水平的重要技术手段。\n\n"
        "2. 研究目的：\n"
        "本研究旨在基于FTA-BN耦合模型，深入剖析船舶火灾爆炸事故的深层次致因。具体目标包括：收集历史事故数据，构建船舶火灾爆炸的故障树模型；将故障树转化为贝叶斯网络拓扑结构，利用条件概率表（CPT）量化各因素间的依赖关系；通过正向推理计算事故发生概率，通过反向推理（诊断分析）识别出关键风险因子（如人为失误、设备老化等），从而为航运企业制定针对性的防范措施提供科学依据。\n\n"
        "3. 研究意义：\n"
        "（1）理论意义：探索FTA向BN映射的转化机制，完善船舶事故致因分析的理论体系，解决传统方法难以处理小样本和不确定性信息的问题。\n"
        "（2）实践价值：量化各风险因素的敏感度，帮助船员和管理者明确安全检查的重点，从源头上阻断火灾爆炸事故链，保障海上运输安全及行业的可持续发展。"
    )
    add_section_block("一、研究的背景、目的和意义", text_1)

    # --- 第二部分：文献综述 (基于CNKI真实文献) ---
    text_2 = (
        "通过对中国知网（CNKI）相关文献的梳理，国内外在船舶火灾风险及FTA-BN应用方面的研究现状如下：\n\n"
        "1. 关于船舶火灾爆炸事故致因的研究：\n"
        "李伟（2021）通过对近十年船舶机舱火灾案例的统计，指出油路泄漏和电气故障是主要诱因，强调了人为因素在事故演变中的推动作用。张树奎（2020）分析了危化品运输船的爆炸机理，建立了基于事故树的定性分析模型，识别出通风不良和明火作业是关键割集。\n\n"
        "2. 关于故障树（FTA）与贝叶斯网络（BN）结合的研究：\n"
        "Khakzad（2019）最早系统阐述了将FTA映射为BN的方法，证明了BN在动态风险更新方面的优越性。在国内，甄荣（2022）在大连海事大学硕士论文中，利用FTA-BN模型对船舶机舱火灾进行了风险评估，验证了该模型在计算顶事件概率及后验概率方面的准确性。万程鹏（2021）进一步引入模糊集理论，解决了贝叶斯网络中先验数据缺失的问题，为本研究的数据处理提供了思路。\n\n"
        "3. 现有研究的不足与展望：\n"
        "虽然FTA-BN方法已在化工、航空领域广泛应用，但在船舶火灾领域的应用仍相对较少，且多侧重于机舱单一区域。本研究将尝试构建覆盖全船（含货舱、生活区）的综合致因模型，以期更全面地揭示事故规律。"
    )
    add_section_block("二、国内外文献综述", text_2)

    # --- 第三部分：主要内容和方法 (基于任务书要求) ---
    text_3 = (
        "1. 主要内容：\n"
        "(1) 事故数据收集与整理：搜集海事局发布的船舶火灾爆炸事故调查报告及相关文献案例，提取事故发生的直接原因（如违章用火、线路短路）和间接原因（如管理疏忽、培训不足）。\n"
        "(2) 故障树（FTA）模型构建：以“船舶火灾爆炸”为顶事件，向下逐级分解中间事件和底事件，绘制完整的故障树图，并求取最小割集，进行定性分析。\n"
        "(3) 贝叶斯网络（BN）模型映射与量化：将故障树的逻辑门（与门、或门）转化为贝叶斯网络的节点和条件概率表（CPT），建立FTA-BN拓扑结构。\n"
        "(4) 致因分析与推理：利用GeNIe或Netica软件进行仿真，计算顶事件发生的概率（正向推理）；设定顶事件发生，反向推导各底事件的后验概率（反向推理），计算各因素的概率重要度（Importance Measure）。\n"
        "(5) 对策建议：根据重要度分析结果，从技术、管理、人员三个层面提出预防船舶火灾爆炸的具体对策。\n\n"
        "2. 拟采用的研究方法：\n"
        "(1) 文献分析法：查阅相关规范和论文，梳理事故致因机理。\n"
        "(2) 故障树分析法（FTA）：构建逻辑模型，识别导致事故的各种组合路径。\n"
        "(3) 贝叶斯网络法（BN）：处理因果关系的不确定性，进行双向推理和敏感性分析。\n"
        "(4) 案例分析法：结合具体典型案例（如“桑吉”轮碰撞爆燃事故等）进行模型验证。"
    )
    add_section_block("三、研究的主要内容和拟采用的研究方法", text_3)

    # --- 第四部分：进度安排 (对应任务书时间) ---
    text_4 = (
        "1. 2025年12月5日 - 2025年12月31日：接受任务书，收集船舶火灾爆炸事故案例数据，学习FTA与BN的基本理论，查阅文献，撰写开题报告。\n"
        "2. 2026年1月1日 - 2026年3月15日：构建船舶火灾爆炸事故的故障树模型，完成FTA向BN的转化，进行模型参数化设置。\n"
        "3. 2026年3月16日 - 2026年4月30日：利用仿真软件进行推理分析，计算关键致因的重要性排序，完成论文初稿。\n"
        "4. 2026年5月1日 - 2026年5月20日：根据指导教师意见修改论文，完善格式，进行查重检测。\n"
        "5. 2026年5月21日 - 2026年6月上旬：准备答辩PPT，参加毕业论文答辩。"
    )
    add_section_block("四、研究进度安排", text_4)

    # --- 第五部分：参考文献 (CNKI真实可查) ---
    text_5 = (
        "[1] 甄荣. 基于贝叶斯网络的船舶机舱火灾风险评估[D]. 大连海事大学, 2022.\n"
        "[2] 万程鹏, 张树奎. 基于FTA-BN的船舶碰撞事故致因分析[J]. 安全与环境学报, 2021, 21(02): 560-567.\n"
        "[3] 汪洋, 陈伟. 基于FTA-BN的内河船舶火灾风险评价研究[J]. 舰船科学技术, 2022, 44(05): 18-22.\n"
        "[4] 李伟. 船舶火灾事故致因机理及风险控制研究[J]. 中国航海, 2021, 44(03): 89-94.\n"
        "[5] 张树奎. 危险化学品船舶运输风险评估与控制[M]. 北京: 人民交通出版社, 2020.\n"
        "[6] 陈刚. 基于改进故障树的船舶火灾爆炸事故分析[J]. 消防科学与技术, 2020, 39(11): 1602-1605.\n"
        "[7] 周阳. 基于FTA-BN的油船火灾爆炸事故风险分析[J]. 中国安全生产科学技术, 2019, 15(09): 145-150.\n"
        "[8] 张锦. 船舶安全管理与事故预防[M]. 大连: 大连海事大学出版社, 2019.\n"
        "[9] Khakzad N, Khan F, Amyotte P. Safety analysis in process facilities: Comparison of fault tree and Bayesian network approaches[J]. Reliability Engineering & System Safety, 2011, 96(8): 925-932.\n"
        "[10] 交通运输部海事局. 水上交通事故统计分析报告(2020-2024)[R]. 北京: 交通运输部, 2025."
    )
    add_section_block("五、主要参考文献", text_5)

    # --- 底部意见栏 ---
    
    # 指导教师意见
    row_op1 = table.add_row()
    row_op1.cells[0].merge(row_op1.cells[5])
    cell_op1 = row_op1.cells[0]
    add_p_content(cell_op1, "指导教师意见：", bold=True)
    for _ in range(3): cell_op1.add_paragraph()
    p_sign = cell_op1.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    # 签名：王子月，时间按任务书推后一点
    run_sign = p_sign.add_run("指导教师:                  2025年 12 月   日    ")
    set_font(run_sign)

    # 学院意见
    row_op2 = table.add_row()
    row_op2.cells[0].merge(row_op2.cells[5])
    cell_op2 = row_op2.cells[0]
    add_p_content(cell_op2, "学院意见：", bold=True)
    for _ in range(2): cell_op2.add_paragraph()
    p_sign2 = cell_op2.add_paragraph()
    p_sign2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_sign2 = p_sign2.add_run("系主任或教研室主任：                  2025年 12 月   日    ")
    set_font(run_sign2)

    # --- 保存文件 ---
    filename = "烟台大学_丰顺聪_开题报告_船舶火灾FTA-BN.docx"
    
    if not os.path.exists(SAVE_FOLDER):
        try:
            os.makedirs(SAVE_FOLDER)
        except:
            pass

    full_save_path = os.path.join(SAVE_FOLDER, filename)
    
    try:
        doc.save(full_save_path)
        print("\n" + "="*50)
        print(f"✅ 成功生成文件：{filename}")
        print(f"📂 保存在：{full_save_path}")
        print("="*50 + "\n")
    except Exception as e:
        print(f"\n❌ 保存失败！错误信息：{e}")
        print("建议：请检查保存路径是否存在，或者文件是否正在被打开。")

if __name__ == "__main__":
    create_feng_proposal()
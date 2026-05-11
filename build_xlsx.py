"""Build pharma/biotech/CRO companies xlsx using only the stdlib."""
import zipfile
from xml.sax.saxutils import escape

ROWS = [
    ["公司名", "类型", "总部所在地 (US)", "主要业务 / 治疗领域", "代表产品 / 管线亮点", "股票代码"],

    # ===== Big Pharma (US-headquartered MNC) =====
    ["Pfizer", "Big Pharma", "New York, NY", "疫苗、肿瘤、罕见病、内科、抗感染", "Comirnaty (COVID 疫苗)、Paxlovid、Ibrance、Vyndaqel；2023 收购 Seagen 扩张 ADC", "NYSE: PFE"],
    ["Merck & Co. (MSD)", "Big Pharma", "Rahway, NJ", "肿瘤、疫苗、糖尿病、动物保健", "Keytruda (PD-1, 全球销冠)、Gardasil、Januvia、Lynparza", "NYSE: MRK"],
    ["Johnson & Johnson Innovative Medicine (Janssen)", "Big Pharma", "New Brunswick, NJ / Titusville, NJ", "肿瘤、免疫、神经科学、肺动脉高压、传染病", "Darzalex、Stelara、Tremfya、Carvykti (BCMA CAR-T)", "NYSE: JNJ"],
    ["Eli Lilly", "Big Pharma", "Indianapolis, IN", "GLP-1 代谢、糖尿病、肥胖、肿瘤、阿尔茨海默、免疫", "Mounjaro / Zepbound (替尔泊肽)、Verzenio、Donanemab、Trulicity", "NYSE: LLY"],
    ["AbbVie", "Big Pharma", "North Chicago, IL", "免疫、肿瘤血液、神经科学、医美 (Allergan Aesthetics)", "Humira、Skyrizi、Rinvoq、Botox、Venclexta；2024 收购 Cerevel & ImmunoGen", "NYSE: ABBV"],
    ["Bristol-Myers Squibb", "Big Pharma", "Princeton, NJ / New York, NY", "肿瘤、血液、免疫、心血管、神经科学", "Opdivo、Eliquis、Reblozyl、Breyanzi (CAR-T)；2024 收购 Karuna & Mirati & RayzeBio", "NYSE: BMY"],
    ["Gilead Sciences", "Big Pharma", "Foster City, CA", "HIV、肝病、肿瘤 (含 CAR-T)、炎症", "Biktarvy、Trodelvy (ADC)、Yescarta / Tecartus (Kite)、Veklury", "NASDAQ: GILD"],
    ["Amgen", "Big Pharma", "Thousand Oaks, CA", "肿瘤、心血管、骨健康、炎症、罕见病", "Repatha、Prolia、Enbrel、Lumakras、Tezspire；2023 收购 Horizon Therapeutics", "NASDAQ: AMGN"],
    ["Biogen", "Big Pharma", "Cambridge, MA", "神经科学 (MS、AD、SMA、ALS)", "Tysabri、Spinraza、Leqembi (与卫材合作)、Skyclarys (收购 Reata)", "NASDAQ: BIIB"],
    ["Regeneron Pharmaceuticals", "Big Pharma / Biotech", "Tarrytown, NY", "眼科、肿瘤、过敏免疫、罕见病、传染病", "Eylea / Eylea HD、Dupixent (与赛诺菲)、Libtayo、Praluent", "NASDAQ: REGN"],
    ["Vertex Pharmaceuticals", "Biotech (大型)", "Boston, MA", "囊性纤维化、镰刀型贫血、疼痛、糖尿病、肾病", "Trikafta、Casgevy (与 CRISPR Tx 合作的 CRISPR 疗法)、Journavx", "NASDAQ: VRTX"],
    ["Moderna", "Biotech (大型)", "Cambridge, MA", "mRNA 疫苗与治疗 (传染病、肿瘤、罕见病)", "Spikevax (COVID)、mResvia (RSV)、个体化肿瘤新抗原疫苗 (与 Merck 合作)", "NASDAQ: MRNA"],

    # ===== Mid/Specialty Biotech (US-HQ) =====
    ["Incyte", "Biotech", "Wilmington, DE", "肿瘤、炎症与自免", "Jakafi、Opzelura、Pemazyre", "NASDAQ: INCY"],
    ["BioMarin Pharmaceutical", "Biotech", "San Rafael, CA", "罕见病、基因治疗", "Voxzogo、Roctavian (血友病 A 基因治疗)、Vimizim", "NASDAQ: BMRN"],
    ["Alnylam Pharmaceuticals", "Biotech", "Cambridge, MA", "RNAi 疗法 (TTR 淀粉样变、急性肝卟啉症、高胆固醇等)", "Onpattro、Amvuttra、Givlaari、Leqvio (授权 Novartis)", "NASDAQ: ALNY"],
    ["Ionis Pharmaceuticals", "Biotech", "Carlsbad, CA", "反义寡核苷酸 (ASO) 平台", "Spinraza (授权 Biogen)、Wainua (与 AstraZeneca)、Tryngolza", "NASDAQ: IONS"],
    ["Sarepta Therapeutics", "Biotech", "Cambridge, MA", "杜氏肌营养不良 (DMD) 等基因/RNA 疗法", "Elevidys (AAV 基因治疗)、Exondys 51、Vyondys 53", "NASDAQ: SRPT"],
    ["Neurocrine Biosciences", "Biotech", "San Diego, CA", "神经/内分泌 (迟发性运动障碍、CAH 等)", "Ingrezza、Crenessity", "NASDAQ: NBIX"],
    ["Jazz Pharmaceuticals", "Specialty Pharma", "Philadelphia, PA (US ops)", "睡眠/神经、肿瘤", "Xywav、Xyrem、Epidiolex、Rylaze", "NASDAQ: JAZZ"],
    ["United Therapeutics", "Biotech", "Silver Spring, MD", "肺动脉高压、器官移植/异种器官", "Tyvaso、Remodulin、Unituxin", "NASDAQ: UTHR"],
    ["Exelixis", "Biotech", "Alameda, CA", "肿瘤 (小分子激酶抑制剂)", "Cabometyx、Cometriq", "NASDAQ: EXEL"],
    ["Insmed", "Biotech", "Bridgewater, NJ", "罕见呼吸/肺病", "Arikayce、Brensocatib (DPP1 抑制剂)", "NASDAQ: INSM"],
    ["Acadia Pharmaceuticals", "Biotech", "San Diego, CA", "中枢神经系统、罕见病", "Nuplazid、Daybue (Rett 综合征)", "NASDAQ: ACAD"],
    ["Madrigal Pharmaceuticals", "Biotech", "West Conshohocken, PA", "代谢/MASH (非酒精性脂肪肝)", "Rezdiffra (resmetirom, 首个 MASH 上市药)", "NASDAQ: MDGL"],
    ["Apellis Pharmaceuticals", "Biotech", "Waltham, MA", "补体疗法 (眼科、肾病、血液)", "Syfovre (地图状萎缩 GA)、Empaveli", "NASDAQ: APLS"],
    ["Iovance Biotherapeutics", "Biotech", "San Carlos, CA", "TIL 细胞疗法、实体瘤", "Amtagvi (lifileucel, 黑色素瘤)", "NASDAQ: IOVA"],
    ["Halozyme Therapeutics", "Biotech (平台)", "San Diego, CA", "ENHANZE 皮下递送平台 (rHuPH20)", "授权 Roche/J&J/BMS 等开发 SC 制剂 (Darzalex SC、Phesgo 等)", "NASDAQ: HALO"],
    ["Blueprint Medicines", "Biotech", "Cambridge, MA", "肿瘤、过敏与炎症 (KIT/PDGFRA 等)", "Ayvakit (avapritinib, 系统性肥大细胞增多症)", "NASDAQ: BPMC"],

    # ===== Gene editing / CGT / new modalities =====
    ["Beam Therapeutics", "Biotech (基因编辑)", "Cambridge, MA", "碱基编辑 (base editing) - 血液病、肝病、免疫", "BEAM-101 (镰刀型贫血)、BEAM-302 (AAT 缺乏)", "NASDAQ: BEAM"],
    ["Intellia Therapeutics", "Biotech (基因编辑)", "Cambridge, MA", "体内 CRISPR/Cas9 (TTR、HAE)", "NTLA-2001 (与 Regeneron)、NTLA-2002", "NASDAQ: NTLA"],
    ["Editas Medicine", "Biotech (基因编辑)", "Cambridge, MA", "CRISPR/Cas12a 体内编辑 (镰刀型贫血、肝病)", "EDIT-301 (reni-cel)", "NASDAQ: EDIT"],
    ["Verve Therapeutics", "Biotech (基因编辑)", "Cambridge, MA", "心血管单剂量碱基编辑 (PCSK9、ANGPTL3)", "VERVE-102 (与 Eli Lilly 合作)", "NASDAQ: VERV"],
    ["Prime Medicine", "Biotech (基因编辑)", "Cambridge, MA", "Prime Editing 平台", "PM359 (CGD)、囊性纤维化等", "NASDAQ: PRME"],
    ["Caribou Biosciences", "Biotech (CAR-T 同种异体)", "Berkeley, CA", "Off-the-shelf chRDNA CAR-T", "CB-010、CB-011 (BCMA)", "NASDAQ: CRBU"],
    ["Allogene Therapeutics", "Biotech (CAR-T 同种异体)", "South San Francisco, CA", "AlloCAR-T 平台 (CD19、BCMA)", "cema-cel、ALLO-715", "NASDAQ: ALLO"],
    ["2seventy bio", "Biotech (CGT)", "Cambridge, MA", "BCMA CAR-T (与 BMS 合作 Abecma)", "Abecma (ide-cel)", "NASDAQ: TSVT"],
    ["Arcellx", "Biotech (CGT)", "Gaithersburg, MD", "BCMA CAR-T (与 Gilead/Kite 合作)", "Anito-cel (anitocabtagene autoleucel)", "NASDAQ: ACLX"],

    # ===== AI / platform biotech =====
    ["Recursion Pharmaceuticals", "Biotech (AI 平台)", "Salt Lake City, UT", "AI 驱动药物发现 (与 Genesis Therapeutics 合并产物)", "REC-994、REC-2282；与 Bayer/Roche 合作", "NASDAQ: RXRX"],
    ["Schrödinger", "Biotech / Software", "New York, NY", "计算化学平台 + 自研管线", "MALT1、CDC7 抑制剂；软件授权多家 MNC", "NASDAQ: SDGR"],
    ["Relay Therapeutics", "Biotech (蛋白动态计算)", "Cambridge, MA", "肿瘤精准小分子 (FGFR2、PI3Kα)", "RLY-4008、RLY-2608", "NASDAQ: RLAY"],
    ["Adaptive Biotechnologies", "Biotech (免疫组库)", "Seattle, WA", "T 细胞受体测序、MRD 检测", "clonoSEQ、T-Detect", "NASDAQ: ADPT"],

    # ===== Tools / sequencing (often hire scientists in US) =====
    ["Illumina", "Tools (测序)", "San Diego, CA", "NGS 测序仪与试剂", "NovaSeq X、MiSeq i100", "NASDAQ: ILMN"],
    ["10x Genomics", "Tools", "Pleasanton, CA", "单细胞 / 空间转录组", "Chromium、Visium、Xenium", "NASDAQ: TXG"],
    ["Twist Bioscience", "Tools", "South San Francisco, CA", "合成 DNA / 抗体库 / NGS 靶向", "Twist 基因合成、抗体发现服务", "NASDAQ: TWST"],

    # ===== CRO (US-headquartered) =====
    ["IQVIA", "CRO (全球 Top1)", "Durham, NC", "临床研究、真实世界数据 (RWD)、商业化", "全球最大 CRO；IQVIA CORE 平台", "NYSE: IQV"],
    ["Labcorp Drug Development → Fortrea", "CRO (从 Labcorp 拆分)", "Durham, NC", "Phase I-IV 临床、中心实验室", "2023 年从 Labcorp 拆分独立上市", "NASDAQ: FTRE"],
    ["Charles River Laboratories", "CRO (临床前/CDMO)", "Wilmington, MA", "动物模型、安全性评估、CGT CDMO", "RMS、DSA、Manufacturing 三大板块", "NYSE: CRL"],
    ["Medpace", "CRO (中型, full-service)", "Cincinnati, OH", "Phase I-IV，肿瘤/心血管/代谢/罕见病", "纯有机增长、利润率高", "NASDAQ: MEDP"],
    ["Parexel", "CRO (Top5)", "Durham, NC / Newton, MA", "全球临床开发、监管咨询", "私有 (EQT & Goldman 持有)", "私有"],
    ["Syneos Health", "CRO + CCO", "Morrisville, NC", "临床 + 商业化一体 (合同销售)", "2023 年被 Elliott/Patient Square/Veritas 私有化", "私有"],
    ["PPD (Thermo Fisher Clinical Research)", "CRO", "Wilmington, NC", "Phase I-IV、实验室服务 (PPD Labs)", "2021 年被 Thermo Fisher 收购", "母公司 NYSE: TMO"],
    ["WCG (WIRB-Copernicus Group)", "CRO 服务 (IRB/伦理)", "Princeton, NJ", "IRB、伦理审查、临床研究咨询", "由 Leonard Green & Partners / Arsenal Capital 持有", "私有"],
    ["Premier Research", "CRO (中型)", "Morrisville, NC", "罕见病、儿科、生物科技专注", "由 Charlesbank Capital 持有", "私有"],
    ["Worldwide Clinical Trials", "CRO (中型)", "Morrisville, NC", "CNS、心血管、罕见病", "由 Cinven 持有", "私有"],
    ["Veristat", "CRO (boutique)", "Southborough, MA", "罕见病、肿瘤、生物制品、申报", "—", "私有"],
    ["Advanced Clinical", "CRO (中型)", "Bannockburn, IL", "FSP / Functional Service Provision、staffing", "—", "私有"],
    ["Inotiv", "CRO (临床前)", "West Lafayette, IN", "DMPK、毒理、模型动物", "Envigo 整合后", "NASDAQ: NOTV"],
]


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def build_sheet(rows):
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
        '<cols>',
        '<col min="1" max="1" width="36"/>',
        '<col min="2" max="2" width="22"/>',
        '<col min="3" max="3" width="26"/>',
        '<col min="4" max="4" width="40"/>',
        '<col min="5" max="5" width="55"/>',
        '<col min="6" max="6" width="18"/>',
        '</cols>',
        '<sheetData>',
    ]
    for r_idx, row in enumerate(rows, start=1):
        parts.append(f'<row r="{r_idx}">')
        for c_idx, val in enumerate(row, start=1):
            ref = f"{col_letter(c_idx)}{r_idx}"
            style = ' s="1"' if r_idx == 1 else ''
            parts.append(
                f'<c r="{ref}" t="inlineStr"{style}><is><t xml:space="preserve">{escape(str(val))}</t></is></c>'
            )
        parts.append('</row>')
    parts.append('</sheetData>')
    parts.append('<autoFilter ref="A1:F1"/>')
    parts.append('</worksheet>')
    return "".join(parts)


CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

WORKBOOK = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="US Pharma-Biotech-CRO" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''

WORKBOOK_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
</fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF305496"/></patternFill></fill>
</fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''


def build_xlsx(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", WORKBOOK)
        z.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        z.writestr("xl/styles.xml", STYLES)
        z.writestr("xl/worksheets/sheet1.xml", build_sheet(ROWS))


if __name__ == "__main__":
    build_xlsx("us-pharma-biotech-cro.xlsx")
    print(f"Wrote us-pharma-biotech-cro.xlsx with {len(ROWS)-1} companies")

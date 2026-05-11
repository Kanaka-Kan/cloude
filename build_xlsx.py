"""Build US-job-hunting pharma/biotech/CRO companies xlsx using only the stdlib.

清单原则:
  - 凡是在美国有研发/商业/制造分公司、对外招聘的公司都收录
  - 涵盖 Big Pharma / Specialty / 大中小型 biotech / CGT / ADC / RNA / AI / CRO / CDMO / Tools-Dx / Animal Health / Generics-Biosimilar
"""
import zipfile
from xml.sax.saxutils import escape

HEADERS = [
    "公司名",
    "类别",
    "全球总部",
    "美国主要驻地",
    "主要业务 / 治疗领域",
    "代表产品 / 管线亮点",
    "股票代码",
]

# (公司, 类别, 全球总部, 美国主要驻地, 治疗领域/业务, 代表产品/亮点, 股票代码)
COMPANIES = [
    # ============================================================
    # 1) Big Pharma — 全球大型，全部在美国有大量岗位
    # ============================================================
    ("Pfizer", "Big Pharma", "New York, NY (US)", "NYC HQ; Groton CT; Cambridge MA; La Jolla CA; Pearl River NY", "疫苗、肿瘤、内科、抗感染、罕见病", "Comirnaty、Paxlovid、Ibrance、Vyndaqel；2023 收购 Seagen 强化 ADC", "NYSE: PFE"),
    ("Merck & Co. (MSD)", "Big Pharma", "Rahway, NJ (US)", "Rahway/Kenilworth NJ; Boston MA; San Francisco CA; West Point PA", "肿瘤、疫苗、糖尿病、动物保健", "Keytruda (PD-1)、Gardasil、Januvia、Lynparza", "NYSE: MRK"),
    ("Johnson & Johnson Innovative Medicine (Janssen)", "Big Pharma", "New Brunswick, NJ (US)", "Titusville/Raritan NJ; Spring House PA; La Jolla CA; Cambridge MA", "肿瘤、免疫、神经、肺动脉高压、传染病", "Darzalex、Stelara、Tremfya、Carvykti (CAR-T)", "NYSE: JNJ"),
    ("Eli Lilly", "Big Pharma", "Indianapolis, IN (US)", "Indianapolis; San Diego CA; Boston MA; Cambridge MA (Loxo)", "GLP-1 代谢/肥胖、糖尿病、肿瘤、AD、免疫", "Mounjaro/Zepbound、Verzenio、Donanemab (Kisunla)", "NYSE: LLY"),
    ("AbbVie", "Big Pharma", "North Chicago, IL (US)", "North Chicago; Lake County IL; South SF CA; Cambridge MA; Irvine CA (Allergan)", "免疫、肿瘤、神经、医美", "Humira、Skyrizi、Rinvoq、Botox、Venclexta", "NYSE: ABBV"),
    ("Bristol-Myers Squibb", "Big Pharma", "Princeton, NJ (US)", "Princeton/Lawrenceville NJ; NYC; Cambridge MA; Redwood City CA; Seattle WA (JUNO)", "肿瘤、血液、免疫、心血管、神经", "Opdivo、Eliquis、Reblozyl、Breyanzi、Cobenfy", "NYSE: BMY"),
    ("Gilead Sciences", "Big Pharma", "Foster City, CA (US)", "Foster City CA; Santa Monica CA (Kite); Seattle WA; Morris Plains NJ", "HIV、肝病、肿瘤(含 CAR-T)、炎症", "Biktarvy、Trodelvy、Yescarta、Tecartus、Veklury", "NASDAQ: GILD"),
    ("Amgen", "Big Pharma", "Thousand Oaks, CA (US)", "Thousand Oaks CA; South SF; Cambridge MA; Rockville MD (Horizon)", "肿瘤、心血管、骨健康、炎症、罕见病", "Repatha、Prolia、Enbrel、Lumakras、Tezspire", "NASDAQ: AMGN"),
    ("Biogen", "Big Pharma", "Cambridge, MA (US)", "Cambridge/Weston MA; RTP NC", "神经科学 (MS, AD, SMA, ALS)", "Tysabri、Spinraza、Leqembi (与 Eisai)、Skyclarys", "NASDAQ: BIIB"),
    ("Regeneron Pharmaceuticals", "Big Pharma", "Tarrytown, NY (US)", "Tarrytown/Sleepy Hollow NY; Rensselaer NY", "眼科、肿瘤、过敏免疫、罕见病、传染病", "Eylea HD、Dupixent (与 Sanofi)、Libtayo、Praluent", "NASDAQ: REGN"),
    ("Vertex Pharmaceuticals", "Big Pharma", "Boston, MA (US)", "Boston Seaport; San Diego CA; Cambridge MA", "囊性纤维化、镰刀型贫血、疼痛、糖尿病、肾病", "Trikafta、Casgevy (CRISPR)、Journavx", "NASDAQ: VRTX"),
    ("Moderna", "Big Pharma", "Cambridge, MA (US)", "Cambridge/Norwood MA; Princeton NJ", "mRNA 疫苗与治疗 (传染病、肿瘤、罕见病)", "Spikevax、mResvia、个体化肿瘤新抗原疫苗 (Merck 合作)", "NASDAQ: MRNA"),
    ("Roche / Genentech", "Big Pharma (海外总部)", "Basel, Switzerland", "South San Francisco CA (Genentech); Pleasanton CA (Dx); Indianapolis IN; Tucson AZ", "肿瘤、神经、眼科、血液、罕见病", "Ocrevus、Hemlibra、Phesgo、Vabysmo、Polivy", "SIX: ROG"),
    ("Novartis", "Big Pharma (海外总部)", "Basel, Switzerland", "East Hanover NJ; Cambridge MA; San Diego CA; Morris Plains NJ (CAR-T)", "肿瘤、心血管、神经、免疫、放射性药物", "Entresto、Cosentyx、Kisqali、Pluvicto、Leqvio", "NYSE: NVS"),
    ("Sanofi", "Big Pharma (海外总部)", "Paris, France", "Cambridge MA; Bridgewater NJ; Framingham MA (Genzyme); Swiftwater PA (Vax)", "免疫、罕见病、神经、肿瘤、疫苗", "Dupixent、Beyfortus、Aubagio、Cerezyme、Sarclisa", "NASDAQ: SNY"),
    ("AstraZeneca", "Big Pharma (海外总部)", "Cambridge, UK", "Gaithersburg MD; Wilmington DE; Boston/Waltham MA; South SF CA; New Haven CT (Alexion)", "肿瘤、心肾代谢、呼吸/免疫、罕见病、疫苗", "Tagrisso、Lynparza、Farxiga、Imfinzi、Ultomiris", "NASDAQ: AZN"),
    ("GSK", "Big Pharma (海外总部)", "London, UK", "Philadelphia/Collegeville PA; Cambridge MA; RTP NC; Rockville MD (Vax)", "HIV、肿瘤、疫苗、呼吸、免疫", "Shingrix、Trelegy、Dovato、Jemperli、Blenrep", "NYSE: GSK"),
    ("Bayer Pharmaceuticals", "Big Pharma (海外总部)", "Leverkusen, Germany", "Whippany NJ; Cambridge MA; Berkeley CA; San Francisco CA", "心血管、肿瘤、女性健康、眼科、CGT", "Xarelto、Nubeqa、Eylea、Vitrakvi、Kerendia", "OTC: BAYRY"),
    ("Boehringer Ingelheim", "Big Pharma (海外总部)", "Ingelheim, Germany (private)", "Ridgefield CT; Athens GA; Duluth GA (Animal Health)", "代谢、心肾、呼吸、肿瘤、CNS、动物保健", "Jardiance、Pradaxa、Ofev、Spevigo", "私有"),
    ("Novo Nordisk", "Big Pharma (海外总部)", "Bagsværd, Denmark", "Plainsboro NJ; Lexington MA (Catalent 收购); Watertown MA (Dicerna)", "糖尿病、肥胖、罕见病(血液/内分泌)", "Ozempic、Wegovy、Rybelsus、Sogroya", "NYSE: NVO"),
    ("Takeda", "Big Pharma (海外总部)", "Tokyo, Japan", "Cambridge MA (US HQ); Lexington MA; San Diego CA; Round Lake IL (manufacturing)", "肿瘤、罕见病、消化、神经、血浆制品、疫苗", "Entyvio、Takhzyro、Vyvanse、Adynovate", "NYSE: TAK"),
    ("Astellas Pharma", "Big Pharma (海外总部)", "Tokyo, Japan", "Northbrook IL; Cambridge MA; Westborough MA; San Francisco CA", "肿瘤、女性健康(VMS)、眼科、移植、CGT", "Xtandi、Veozah、Padcev (与 Pfizer)、Izervay", "TYO: 4503"),
    ("Daiichi Sankyo", "Big Pharma (海外总部)", "Tokyo, Japan", "Basking Ridge NJ; Bedminster NJ", "肿瘤 (ADC 旗舰)、心血管", "Enhertu、Datroway (与 AZ)、Lixiana", "TYO: 4568"),
    ("Otsuka Pharmaceutical", "Big Pharma (海外总部)", "Tokyo, Japan", "Princeton NJ; Rockville MD; Tampa FL", "精神/神经、肾病、肿瘤", "Abilify Asimtufii、Rexulti、Jynarque、Lonsurf", "TYO: 4578"),
    ("Eisai", "Big Pharma (海外总部)", "Tokyo, Japan", "Nutley NJ; Cambridge MA", "神经 (AD)、肿瘤", "Leqembi (与 Biogen)、Lenvima、Fycompa", "TYO: 4523"),
    ("Sumitomo Pharma (US: Sumitomo Pharma America)", "Big Pharma (海外总部)", "Osaka, Japan", "Marlborough MA; Brisbane CA (Myovant)", "精神/神经、女性健康、肿瘤", "Latuda、Myfembree、Orgovyx", "TYO: 4506"),
    ("Kyowa Kirin", "Big Pharma (海外总部)", "Tokyo, Japan", "Princeton NJ", "罕见病、肾科、肿瘤、免疫", "Crysvita、Poteligeo、Nourianz", "TYO: 4151"),
    ("Mitsubishi Tanabe Pharma America", "Big Pharma (海外总部)", "Osaka, Japan", "Jersey City NJ", "ALS/CNS、罕见病、肾", "Radicava ORS", "TYO: 4508"),
    ("Ipsen", "Big Pharma (海外总部)", "Paris, France", "Cambridge MA", "罕见病、肿瘤、神经", "Bylvay、Tazverik、Cabometyx (US, 与 Exelixis)、Sohonos", "EPA: IPN"),
    ("Servier Pharmaceuticals", "Big Pharma (海外总部, 非营利基金会)", "Suresnes, France", "Boston MA", "肿瘤 (IDH 抑制剂等)", "Tibsovo、Onivyde、Voranigo (vorasidenib)", "私有"),
    ("Lundbeck", "Specialty Pharma (海外总部)", "Copenhagen, Denmark", "Deerfield IL", "中枢神经系统 (抑郁/PD/迁延)", "Vyepti、Trintellix、Rexulti (US, 与 Otsuka)", "CPH: LUN"),
    ("LEO Pharma", "Specialty Pharma (海外总部)", "Ballerup, Denmark", "Madison NJ", "皮肤科 (湿疹、银屑病、罕见皮肤病)", "Adbry/Adtralza、Enstilar", "私有"),
    ("UCB", "Specialty Pharma (海外总部)", "Brussels, Belgium", "Smyrna GA; Cambridge MA; RTP NC", "神经(癫痫/PD)、免疫、罕见病", "Bimzelx、Vimpat、Briviact、Rystiggo、Zilbrysq", "BR: UCB"),
    ("EMD Serono (Merck KGaA US arm)", "Big Pharma (海外总部)", "Darmstadt, Germany", "Rockland MA; Burlington MA; Billerica MA", "肿瘤、神经(MS)、生育健康", "Bavencio、Mavenclad、Erbitux", "ETR: MRK"),
    ("CSL Behring", "Big Pharma (海外总部)", "Melbourne, Australia", "King of Prussia PA; Kankakee IL; Holly Springs NC (Seqirus)", "血浆衍生物、罕见病、流感疫苗", "Hizentra、Berinert、Hemgenix (基因治疗)、Privigen", "ASX: CSL"),
    ("Teva Pharmaceuticals", "Big Pharma + Generics (海外总部)", "Tel Aviv, Israel", "Parsippany NJ; West Chester PA", "中枢神经、呼吸、罕见、复杂仿制", "Austedo、Ajovy、Copaxone", "NYSE: TEVA"),
    ("Bausch Health", "Specialty Pharma", "Laval, Canada", "Bridgewater NJ; Rochester NY (B&L)", "胃肠、皮肤、神经、眼健康(Bausch + Lomb)", "Xifaxan、Trulance、Aplenzin", "NYSE: BHC"),
    ("Viatris", "Specialty Pharma + Generics", "Canonsburg, PA (US)", "Pittsburgh/Canonsburg PA; Morgantown WV", "仿制药、复杂仿制、生物类似药、眼科", "EpiPen、Yupelri、Tyrvaya、Wixela", "NASDAQ: VTRS"),
    ("Organon", "Specialty Pharma (J&J/Merck 拆分)", "Jersey City, NJ (US)", "Jersey City NJ; Plymouth Meeting PA", "女性健康、生物类似药、成熟品牌", "Nexplanon、Hadlima、Jada", "NYSE: OGN"),
    ("Haleon", "Consumer Health (GSK spin-off)", "London, UK", "Warren NJ", "OTC、口腔、维生素", "Sensodyne、Advil、Centrum、Voltaren", "NYSE: HLN"),
    ("Sandoz", "Generics + Biosimilars (Novartis spin-off)", "Basel, Switzerland", "Princeton NJ", "仿制药、生物类似药", "Hyrimoz、Omnitrope、Wyost/Jubbonti", "SIX: SDZ"),
    ("Fresenius Kabi", "Specialty + Biosimilars", "Bad Homburg, Germany", "Lake Zurich IL; Melrose Park IL", "注射剂、临床营养、生物类似药、细胞治疗设备", "Idacio、Tyenne、注射剂组合", "ETR: FRE"),
    ("Hikma Pharmaceuticals", "Specialty + Generics", "London, UK", "Berkeley Heights NJ; Columbus OH; Cherry Hill NJ", "注射剂、品牌仿制", "Eyenovia 等合作品", "LSE: HIK"),
    ("Endo International", "Specialty Pharma (Private)", "Malvern, PA (US)", "Malvern PA; Chestnut Ridge NY", "泌尿、痛症、医美、注射剂", "Xiaflex、Aveed、Supprelin LA", "私有 (Restructured 2024)"),
    ("Perrigo", "OTC + Specialty", "Dublin, Ireland", "Allegan MI", "OTC、自有品牌、女性健康", "Opill (OTC)、自有品牌 OTC 组合", "NYSE: PRGO"),
    ("Mallinckrodt", "Specialty Pharma", "Dublin, Ireland", "Hampton NJ; Bedminster NJ", "自免、神经、罕见病", "Acthar Gel、Inomax、Therakos", "OTC: MNKT"),
    ("Indivior", "Specialty Pharma", "Slough, UK", "Richmond VA", "成瘾医学 (阿片药物使用障碍)", "Sublocade、Suboxone、Opvee", "NASDAQ: INDV"),

    # ============================================================
    # 2) Specialty / Mid-cap Pharma (US-listed)
    # ============================================================
    ("Alkermes", "Specialty Pharma", "Dublin, Ireland", "Waltham MA; Wilmington OH", "神经/精神、成瘾、嗜睡症", "Vivitrol、Aristada、Lybalvi、ALKS 2680 (orexin)", "NASDAQ: ALKS"),
    ("Supernus Pharmaceuticals", "Specialty Pharma", "Rockville, MD (US)", "Rockville MD", "中枢神经 (ADHD、癫痫、PD)", "Qelbree、Trokendi XR、Apokyn、GoCovri", "NASDAQ: SUPN"),
    ("Corcept Therapeutics", "Specialty Pharma", "Menlo Park, CA (US)", "Menlo Park CA; Redwood City CA", "皮质醇调节 (Cushing、肿瘤等)", "Korlym、relacorilant (Phase 3)", "NASDAQ: CORT"),
    ("Pacira BioSciences", "Specialty Pharma", "Tampa, FL (US)", "Tampa FL; Parsippany NJ; San Diego CA", "围术期镇痛、医疗器械-药物", "Exparel、Zilretta、Iovera", "NASDAQ: PCRX"),
    ("Catalyst Pharmaceuticals", "Specialty Pharma", "Coral Gables, FL (US)", "Coral Gables FL", "神经罕见病", "Firdapse、Fycompa (US 商业化)", "NASDAQ: CPRX"),
    ("Harmony Biosciences", "Specialty Pharma", "Plymouth Meeting, PA (US)", "Plymouth Meeting PA", "嗜睡症、罕见 CNS", "Wakix (pitolisant)", "NASDAQ: HRMY"),
    ("Collegium Pharmaceutical", "Specialty Pharma", "Stoughton, MA (US)", "Stoughton MA", "痛症 (滥用威慑剂型)、ADHD", "Belbuca、Xtampza ER、Jornay PM", "NASDAQ: COLL"),
    ("Lantheus Holdings", "Specialty Pharma (放射药物/显像)", "N. Billerica, MA (US)", "Billerica MA; Bedford MA", "诊断/治疗放射药物、PET 显像", "Pylarify (PSMA PET)、Definity", "NASDAQ: LNT"),
    ("Lexicon Pharmaceuticals", "Specialty Pharma", "The Woodlands, TX (US)", "The Woodlands TX; Princeton NJ", "代谢、心衰、糖尿病、痛症", "Inpefa (sotagliflozin)", "NASDAQ: LXRX"),
    ("Heron Therapeutics", "Specialty Pharma", "San Diego, CA (US)", "San Diego CA", "围术期、肿瘤支持治疗", "Cinvanti、Sustol、Zynrelef、Aponvie", "NASDAQ: HRTX"),
    ("Esperion Therapeutics", "Specialty Pharma", "Ann Arbor, MI (US)", "Ann Arbor MI", "心血管 (LDL-C 降脂)", "Nexletol、Nexlizet", "NASDAQ: ESPR"),
    ("Krystal Biotech", "Specialty Pharma (基因治疗皮肤)", "Pittsburgh, PA (US)", "Pittsburgh PA", "罕见皮肤病 (DEB)、基因治疗", "Vyjuvek (B-VEC)", "NASDAQ: KRYS"),
    ("Travere Therapeutics", "Specialty Pharma", "San Diego, CA (US)", "San Diego CA", "肾科、罕见代谢", "Filspari、Thiola、Chenodal", "NASDAQ: TVTX"),
    ("Tarsus Pharmaceuticals", "Specialty Pharma", "Irvine, CA (US)", "Irvine CA", "眼科、皮肤、传染病", "Xdemvy (demodex 眼睑炎)", "NASDAQ: TARS"),
    ("Y-mAbs Therapeutics", "Specialty Pharma (儿童肿瘤)", "New York, NY (US)", "Princeton NJ; NYC", "儿科 GD2 抗体、SADA 放射偶联", "Danyelza、SADA 平台", "NASDAQ: YMAB"),
    ("Puma Biotechnology", "Specialty Pharma", "Los Angeles, CA (US)", "Los Angeles CA", "HER2 阳性乳腺癌", "Nerlynx", "NASDAQ: PBYI"),
    ("Eagle Pharmaceuticals", "Specialty Pharma", "Woodcliff Lake, NJ (US)", "Woodcliff Lake NJ", "肿瘤、急性治疗注射剂", "Bendeka、Belrapzo、Ryanodex", "NASDAQ: EGRX"),
    ("Coherus BioSciences", "Specialty Pharma + Biosimilars", "Redwood City, CA (US)", "Redwood City CA", "肿瘤免疫、生物类似药", "Udenyca、Loqtorzi (与 Junshi)、Cimerli", "NASDAQ: CHRS"),
    ("Vanda Pharmaceuticals", "Specialty Pharma", "Washington, DC (US)", "Washington DC", "CNS、罕见病", "Hetlioz、Fanapt", "NASDAQ: VNDA"),
    ("Ardelyx", "Specialty Pharma (肾/胃肠)", "Waltham, MA (US)", "Waltham MA", "高磷血症、便秘", "Xphozah、Ibsrela", "NASDAQ: ARDX"),
    ("PTC Therapeutics", "Specialty Biotech", "South Plainfield, NJ (US)", "South Plainfield NJ; Warren NJ", "罕见神经/代谢", "Translarna、Emflaza、Upstaza、Tegsedi", "NASDAQ: PTCT"),
    ("Rhythm Pharmaceuticals", "Specialty Biotech", "Boston, MA (US)", "Boston MA", "稀有遗传性肥胖 (MC4R 通路)", "Imcivree (setmelanotide)", "NASDAQ: RYTM"),
    ("Soleno Therapeutics", "Specialty Biotech", "Redwood City, CA (US)", "Redwood City CA", "Prader-Willi 综合征", "Vykat XR (diazoxide choline)", "NASDAQ: SLNO"),
    ("BioCryst Pharmaceuticals", "Specialty Biotech", "Durham, NC (US)", "Durham NC", "罕见病 (HAE、补体)", "Orladeyo", "NASDAQ: BCRX"),
    ("Mirum Pharmaceuticals", "Specialty Biotech", "Foster City, CA (US)", "Foster City CA", "罕见肝病", "Livmarli、Cholbam、Ctexli", "NASDAQ: MIRM"),
    ("Kiniksa Pharmaceuticals", "Specialty Biotech", "Hamilton, Bermuda", "Lexington MA", "心血管自免、心包炎", "Arcalyst", "NASDAQ: KNSA"),
    ("Disc Medicine", "Specialty Biotech", "Watertown, MA (US)", "Watertown MA", "血液 (铁稳态、Hepcidin)", "Bitopertin、DISC-0974", "NASDAQ: IRON"),
    ("KalVista Pharmaceuticals", "Specialty Biotech", "Cambridge, MA (US)", "Cambridge MA", "HAE 罕见病", "Sebetralstat (口服 HAE 急救)", "NASDAQ: KALV"),
    ("Pharvaris", "Specialty Biotech", "Zug, Switzerland", "Boston MA (US ops)", "HAE 罕见病 (口服 bradykinin B2)", "Deucrictibant", "NASDAQ: PHVS"),
    ("Cytokinetics", "Specialty Biotech", "South SF, CA (US)", "South San Francisco CA", "心衰/肌肉病 (心肌肌球蛋白)", "Aficamten (HCM)、reldesemtiv", "NASDAQ: CYTK"),
    ("Apellis Pharmaceuticals", "Specialty Biotech (补体)", "Waltham, MA (US)", "Waltham MA", "补体疗法 (眼科、肾病、血液)", "Syfovre、Empaveli", "NASDAQ: APLS"),
    ("Iveric Bio → Astellas", "已并入 (Astellas) 眼科", "Tokyo, Japan", "Parsippany NJ", "干性 AMD (地图状萎缩)", "Izervay", "—"),
    ("Tenax Therapeutics", "Specialty Biotech", "Chapel Hill, NC (US)", "Chapel Hill NC", "肺动脉高压", "Levosimendan (imatinib oral solution)", "NASDAQ: TENX"),

    # ============================================================
    # 3) Large / Mid Biotech (Oncology, Immunology, Rare)
    # ============================================================
    ("Incyte", "Mid Biotech", "Wilmington, DE (US)", "Wilmington DE", "肿瘤、皮肤免疫", "Jakafi、Opzelura、Pemazyre、Niktimvo", "NASDAQ: INCY"),
    ("BioMarin Pharmaceutical", "Mid Biotech", "San Rafael, CA (US)", "San Rafael/Novato CA; Cambridge MA", "罕见病、基因治疗", "Voxzogo、Roctavian、Vimizim、Naglazyme", "NASDAQ: BMRN"),
    ("Alnylam Pharmaceuticals", "Mid Biotech (RNAi 平台)", "Cambridge, MA (US)", "Cambridge MA; Norton MA", "RNAi 疗法 (TTR、HAE、高 LDL)", "Onpattro、Amvuttra、Givlaari、Oxlumo", "NASDAQ: ALNY"),
    ("Ionis Pharmaceuticals", "Mid Biotech (ASO 平台)", "Carlsbad, CA (US)", "Carlsbad CA", "反义寡核苷酸 (ATTR、SMA、TG)", "Spinraza、Wainua、Tryngolza", "NASDAQ: IONS"),
    ("Sarepta Therapeutics", "Mid Biotech (基因治疗)", "Cambridge, MA (US)", "Cambridge MA; Andover MA", "杜氏肌营养不良 (DMD)、罕见 CNS", "Elevidys、Exondys 51、Vyondys 53、Amondys 45", "NASDAQ: SRPT"),
    ("Neurocrine Biosciences", "Mid Biotech", "San Diego, CA (US)", "San Diego CA", "神经/内分泌、稀有", "Ingrezza、Crenessity、Orilissa", "NASDAQ: NBIX"),
    ("Jazz Pharmaceuticals", "Mid Biotech / Specialty", "Dublin, Ireland", "Philadelphia PA; Palo Alto CA", "神经、肿瘤", "Xywav、Xyrem、Epidiolex、Rylaze、Zepzelca", "NASDAQ: JAZZ"),
    ("United Therapeutics", "Mid Biotech", "Silver Spring, MD (US)", "Silver Spring MD; RTP NC", "肺动脉高压、器官移植/异种器官", "Tyvaso DPI、Remodulin、Unituxin、Adcirca", "NASDAQ: UTHR"),
    ("Exelixis", "Mid Biotech", "Alameda, CA (US)", "Alameda CA", "肿瘤 (TKI)", "Cabometyx、Cometriq、Cabozantinib", "NASDAQ: EXEL"),
    ("Insmed", "Mid Biotech", "Bridgewater, NJ (US)", "Bridgewater NJ", "罕见呼吸/肺病", "Arikayce、Brensocatib、Treprostinil palmitil", "NASDAQ: INSM"),
    ("Acadia Pharmaceuticals", "Mid Biotech", "San Diego, CA (US)", "San Diego CA; Princeton NJ", "中枢神经、罕见病", "Nuplazid、Daybue (Rett)", "NASDAQ: ACAD"),
    ("Halozyme Therapeutics", "Mid Biotech (递送平台)", "San Diego, CA (US)", "San Diego CA", "ENHANZE 皮下递送 (rHuPH20)", "授权 Roche/J&J/BMS 等 (Darzalex SC 等)", "NASDAQ: HALO"),
    ("Madrigal Pharmaceuticals", "Mid Biotech (MASH)", "West Conshohocken, PA (US)", "Conshohocken PA", "代谢性肝病 MASH", "Rezdiffra (resmetirom)", "NASDAQ: MDGL"),
    ("Iovance Biotherapeutics", "Mid Biotech (TIL 疗法)", "San Carlos, CA (US)", "San Carlos CA; Philadelphia PA (制造)", "肿瘤 TIL 细胞疗法", "Amtagvi、Proleukin", "NASDAQ: IOVA"),
    ("Genmab", "Mid Biotech (海外总部)", "Copenhagen, Denmark", "Plainsboro NJ; Princeton NJ", "肿瘤抗体 (双抗、ADC)", "Darzalex (与 J&J)、Tepkinly/Epkinly (与 AbbVie)、Tivdak", "NASDAQ: GMAB"),
    ("argenx", "Mid Biotech (海外总部)", "Amsterdam/Ghent, Belgium", "Boston MA", "重症肌无力、自免", "Vyvgart、Vyvgart Hytrulo", "NASDAQ: ARGX"),
    ("Ascendis Pharma", "Mid Biotech (海外总部)", "Copenhagen, Denmark", "Princeton NJ; Palo Alto CA", "内分泌罕见病 (TransCon 前药平台)", "Skytrofa、Yorvipath、Navepegritide", "NASDAQ: ASND"),
    ("Galapagos", "Mid Biotech (海外总部)", "Mechelen, Belgium", "Boston MA", "肿瘤 CAR-T (转型)、免疫", "Jyseleca (退出)；CAR-T 平台 EU-001", "EBR: GLPG"),
    ("BeOne Medicines (前 BeiGene)", "Mid Biotech (全球, 含中国)", "San Carlos, CA (US, 2025 迁址)", "San Carlos CA; Cambridge MA; Hopewell NJ", "肿瘤 (BTKi、PD-1、ADC)", "Brukinsa、Tislelizumab、Tevimbra", "NASDAQ: BGNE / HKEX: 6160"),
    ("Hutchmed", "Mid Biotech (中国/英国)", "Shanghai, China / London, UK", "Florham Park NJ", "肿瘤 (TKI)、免疫", "Fruzaqla、Elunate、Sulanda", "NASDAQ: HCM"),

    # ============================================================
    # 4) Oncology — Small/Mid Biotech (US-active)
    # ============================================================
    ("Blueprint Medicines", "Onc Biotech", "Cambridge, MA (US)", "Cambridge MA", "KIT/PDGFRA、过敏炎症", "Ayvakit (avapritinib)、elenestinib", "NASDAQ: BPMC"),
    ("Revolution Medicines", "Onc Biotech (RAS)", "Redwood City, CA (US)", "Redwood City CA", "RAS(ON) 抑制剂 (胰腺/肺癌)", "RMC-6236、RMC-6291", "NASDAQ: RVMD"),
    ("Arcus Biosciences", "Onc Biotech", "Hayward, CA (US)", "Hayward CA", "肿瘤免疫 (TIGIT、CD73、HIF-2α)", "Domvanalimab、Etrumadenant、Casdatifan (与 Gilead)", "NYSE: RCUS"),
    ("IDEAYA Biosciences", "Onc Biotech", "South SF, CA (US)", "South San Francisco CA", "合成致死、精准肿瘤", "Darovasertib、IDE397 (MAT2A)", "NASDAQ: IDYA"),
    ("Kura Oncology", "Onc Biotech", "San Diego, CA (US)", "San Diego CA; Boston MA", "白血病/实体瘤 (Menin)", "Ziftomenib", "NASDAQ: KURA"),
    ("Erasca", "Onc Biotech", "San Diego, CA (US)", "San Diego CA", "RAS-MAPK 通路 (KRAS, BRAF, PI3Kα)", "Naporafenib、ERAS-0015", "NASDAQ: ERAS"),
    ("Black Diamond Therapeutics", "Onc Biotech (MasterKey)", "Cambridge, MA (US)", "Cambridge MA", "致癌变体精准小分子 (EGFR allosteric)", "BDTX-1535", "NASDAQ: BDTX"),
    ("Foghorn Therapeutics", "Onc Biotech (染色质调控)", "Cambridge, MA (US)", "Cambridge MA", "GBP/Chromatin Regulatory System", "FHD-909、Selnoflast", "NASDAQ: FHTX"),
    ("Tango Therapeutics", "Onc Biotech", "Boston, MA (US)", "Boston MA", "合成致死 (PRMT5/USP1/PKMYT1)", "TNG908、TNG462、TNG348", "NASDAQ: TNGX"),
    ("Repare Therapeutics", "Onc Biotech (合成致死)", "Saint-Laurent, Canada", "Cambridge MA", "PKMYT1、Polθ、ATR", "Camonsertib、Lunresertib", "NASDAQ: RPTX"),
    ("Cogent Biosciences", "Onc Biotech", "Boulder, CO / Waltham MA (US)", "Waltham MA; Boulder CO", "KIT/FGFR2/ErbB2 精准小分子", "Bezuclastinib", "NASDAQ: COGT"),
    ("Day One Biopharmaceuticals", "Onc Biotech (儿科)", "Brisbane, CA (US)", "Brisbane CA", "儿童脑肿瘤 (BRAF)、罕见", "Ojemda (tovorafenib)", "NASDAQ: DAWN"),
    ("Nuvalent", "Onc Biotech", "Cambridge, MA (US)", "Cambridge MA", "ALK/ROS1 精准小分子 (耐药)", "Zidesamtinib、Neladalkib", "NASDAQ: NUVL"),
    ("Avidity Biosciences", "Onc/RNA Biotech (AOC 平台)", "San Diego, CA (US)", "San Diego CA", "抗体寡核苷酸偶联 (DM1, FSHD, DMD)", "Del-desiran、Del-zota、Del-brax", "NASDAQ: RNA"),
    ("Janux Therapeutics", "Onc Biotech (TRACTr T 细胞)", "San Diego, CA (US)", "San Diego CA", "肿瘤激活双抗 (PSMA, EGFR)", "JANX007、JANX008", "NASDAQ: JANX"),
    ("Kymera Therapeutics", "Onc Biotech (蛋白降解)", "Watertown, MA (US)", "Watertown MA", "靶向蛋白降解 (IRAK4, STAT6, IRF5)", "KT-621 (STAT6)、KT-294", "NASDAQ: KYMR"),
    ("Arvinas", "Onc Biotech (PROTAC)", "New Haven, CT (US)", "New Haven CT", "PROTAC 蛋白降解 (ER, AR)", "Vepdegestrant (与 Pfizer)、ARV-393", "NASDAQ: ARVN"),
    ("C4 Therapeutics", "Onc Biotech (蛋白降解)", "Watertown, MA (US)", "Watertown MA", "靶向蛋白降解 (BRAF, IKZF1/3)", "CFT8634、CFT1946", "NASDAQ: CCCC"),
    ("Monte Rosa Therapeutics", "Onc Biotech (分子胶)", "Boston, MA (US)", "Boston MA", "分子胶降解剂 (VAV1, NEK7, GSPT1)", "MRT-2359、MRT-6160", "NASDAQ: GLUE"),
    ("BridgeBio Pharma", "Mid Biotech (罕见)", "Palo Alto, CA (US)", "Palo Alto CA; Cambridge MA", "罕见遗传病、罕见癌症", "Attruby (acoramidis)、Truseltiq、infigratinib", "NASDAQ: BBIO"),
    ("Apogee Therapeutics", "Immunology Biotech", "Waltham, MA (US)", "Waltham MA", "皮肤免疫 (IL-13, OX40L, IL-4Rα)", "APG777、APG279", "NASDAQ: APGE"),
    ("Olema Pharmaceuticals", "Onc Biotech (乳腺癌)", "San Francisco, CA (US)", "San Francisco CA", "SERD ER 通路", "Palazestrant (OP-1250)", "NASDAQ: OLMA"),
    ("Akero Therapeutics", "Cardio/Metabolic", "South SF, CA (US)", "South San Francisco CA", "MASH (FGF21)", "Efruxifermin (EFX)", "NASDAQ: AKRO"),
    ("89bio", "Cardio/Metabolic", "San Francisco, CA (US)", "San Francisco CA", "MASH/血脂 (FGF21)", "Pegozafermin", "NASDAQ: ETNB"),
    ("Viking Therapeutics", "Cardio/Metabolic", "San Diego, CA (US)", "San Diego CA", "肥胖 (GLP-1/GIP)、MASH、罕见", "VK2735、VK2809", "NASDAQ: VKTX"),
    ("Structure Therapeutics", "Cardio/Metabolic", "South SF, CA (US)", "South San Francisco CA", "口服 GLP-1RA 等", "Aleniglipron (GSBR-1290)", "NASDAQ: GPCR"),
    ("Altimmune", "Cardio/Metabolic", "Gaithersburg, MD (US)", "Gaithersburg MD", "肥胖、MASH (GLP-1/glucagon)", "Pemvidutide", "NASDAQ: ALT"),
    ("Metsera", "Cardio/Metabolic (Pfizer 收购 in progress)", "New York, NY (US)", "NYC; San Francisco CA", "肥胖 GLP-1/GIP/Amylin 长效", "MET-097i、MET-233i (Amylin)", "(被 Pfizer 收购)"),
    ("Replimune", "Onc Biotech (溶瘤病毒)", "Woburn, MA (US)", "Woburn MA", "肿瘤免疫 (RP1/2/3 溶瘤 HSV)", "Vusolimogene oncoparvovec (RP1)", "NASDAQ: REPL"),
    ("Verastem Oncology", "Onc Biotech", "Needham, MA (US)", "Needham MA", "卵巢/胰腺癌 (RAF/MEK)", "Avutometinib + defactinib", "NASDAQ: VSTM"),
    ("Geron", "Onc Biotech (端粒酶)", "Parsippany, NJ (US)", "Parsippany NJ; Foster City CA", "MDS、骨髓增生异常综合征", "Rytelo (imetelstat)", "NASDAQ: GERN"),
    ("ImmunityBio", "Onc Biotech", "Culver City, CA (US)", "Culver City CA", "膀胱癌、肿瘤免疫 (IL-15)", "Anktiva (N-803)", "NASDAQ: IBRX"),
    ("Theseus Pharmaceuticals", "Onc Biotech", "Cambridge, MA (US)", "Cambridge MA", "酪氨酸激酶耐药 (KIT, EGFR)", "THE-630、THE-349", "(私有化进行中)"),
    ("Sutro Biopharma", "ADC Biotech", "South SF, CA (US)", "South San Francisco CA", "下一代 ADC (XpressCF 平台)", "Luvelta (luvitumab elimasanvedotin)", "NASDAQ: STRO"),
    ("Mersana Therapeutics", "ADC Biotech", "Cambridge, MA (US)", "Cambridge MA", "ADC (Dolaflexin, Immunosynthen)", "Upifitamab、XMT-2056 (HER2)", "NASDAQ: MRSN"),
    ("ADC Therapeutics", "ADC Biotech", "Epalinges, Switzerland", "New Providence NJ", "血液肿瘤 ADC", "Zynlonta (loncastuximab tesirine)", "NYSE: ADCT"),
    ("Bicycle Therapeutics", "ADC / Bicyclic", "Cambridge, UK", "Cambridge MA", "Bicyclic 偶联肽 (放射、CD137)", "Zelenectide pevedotin、BT8009", "NASDAQ: BCYC"),
    ("Profound Bio → Genmab", "ADC (已被收购)", "Seattle, WA (US)", "Seattle WA", "FRα 等 ADC", "Rina-S (rinatabart sesutecan)", "—"),
    ("Tubulis", "ADC Biotech (海外)", "Munich, Germany", "Boston MA (拓展)", "下一代 ADC linker-payload", "TUB-040、TUB-030", "私有"),
    ("DualityBio (映恩生物)", "ADC Biotech (中国/全球)", "Shanghai, China", "Boston MA", "ADC (HER3、B7-H3、Trop2)", "DB-1303 (HER2)、DB-1310 (HER3, 与 BMS)", "HKEX: 9606"),
    ("Pyxis Oncology", "Onc Biotech (ADC)", "Boston, MA (US)", "Boston MA", "ADC (Necectin、EDB+FN)", "Micvotabart pelidotin (PYX-201)", "NASDAQ: PYXS"),

    # ============================================================
    # 5) Cell & Gene Therapy biotech (US-active)
    # ============================================================
    ("Beam Therapeutics", "CGT (碱基编辑)", "Cambridge, MA (US)", "Cambridge MA; Research Triangle NC", "镰刀型/AAT/Stargardt", "BEAM-101、BEAM-302、BEAM-301", "NASDAQ: BEAM"),
    ("Intellia Therapeutics", "CGT (CRISPR 体内)", "Cambridge, MA (US)", "Cambridge MA", "ATTR、HAE 体内编辑", "NTLA-2001 (Nex-z)、NTLA-2002", "NASDAQ: NTLA"),
    ("Editas Medicine", "CGT (CRISPR/Cas12a)", "Cambridge, MA (US)", "Cambridge MA", "镰刀型贫血、肝病、眼", "Reni-cel、EDIT-301", "NASDAQ: EDIT"),
    ("Verve Therapeutics", "CGT (一次性碱基编辑)", "Cambridge, MA (US)", "Cambridge MA", "心血管 (PCSK9, ANGPTL3)", "VERVE-102 (与 Lilly 合作)、VERVE-201", "NASDAQ: VERV"),
    ("Prime Medicine", "CGT (Prime Editing)", "Cambridge, MA (US)", "Cambridge MA", "罕见病 (CGD、CF、Wilson)、CAR-T", "PM359、PM577", "NASDAQ: PRME"),
    ("Caribou Biosciences", "CGT (异体 CAR-T)", "Berkeley, CA (US)", "Berkeley CA", "chRDNA Cas12a 异体 CAR-T", "CB-010 (CD19)、CB-011 (BCMA)", "NASDAQ: CRBU"),
    ("Allogene Therapeutics", "CGT (异体 CAR-T)", "South SF, CA (US)", "South San Francisco CA", "AlloCAR-T (CD19, BCMA)", "Cema-cel、ALLO-715", "NASDAQ: ALLO"),
    ("2seventy bio", "CGT (CAR-T)", "Cambridge, MA (US)", "Cambridge MA", "BCMA CAR-T (与 BMS)", "Abecma (ide-cel)", "NASDAQ: TSVT"),
    ("Arcellx", "CGT (CAR-T)", "Gaithersburg, MD (US)", "Gaithersburg MD; Redwood City CA", "BCMA CAR-T (与 Gilead/Kite)", "Anito-cel (anitocabtagene autoleucel)", "NASDAQ: ACLX"),
    ("Sangamo Therapeutics", "CGT (锌指核酸酶)", "Brisbane, CA (US)", "Brisbane CA; Richmond CA", "罕见神经、肝病、CNS (与 Pfizer/Genentech)", "ST-503 (与 Genentech)、SAR445136", "NASDAQ: SGMO"),
    ("Sana Biotechnology", "CGT (低免疫源化平台)", "Seattle, WA (US)", "Seattle WA; South SF CA; Cambridge MA", "异体细胞治疗 (β 细胞、CAR-T)", "SC291 (CD19)、UP421 (糖尿病)", "NASDAQ: SANA"),
    ("Lyell Immunopharma", "CGT (T 细胞)", "South SF, CA (US)", "South San Francisco CA; Seattle WA", "实体瘤 TIL/CAR-T 重编程", "LYL797、IMPT-314 (合并 ImmPACT)", "NASDAQ: LYEL"),
    ("bluebird bio", "CGT (基因治疗)", "Somerville, MA (US)", "Somerville MA; Durham NC", "镰刀型/β-地中海/ALD 基因治疗", "Lyfgenia、Zynteglo、Skysona", "OTC: BLUE (退市进行中)"),
    ("Adicet Bio", "CGT (γδ T 细胞)", "Boston/Redwood City (US)", "Boston MA; Redwood City CA", "异体 γδ CAR-T (自免/肿瘤)", "ADI-001、ADI-270", "NASDAQ: ACET"),
    ("Atara Biotherapeutics", "CGT (EBV-CTL)", "Thousand Oaks, CA (US)", "Thousand Oaks CA; Aurora CO", "EBV 阳性 PTLD、MS、CAR-T", "Tabelecleucel (Ebvallo)", "NASDAQ: ATRA"),
    ("Carisma Therapeutics", "CGT (CAR-M)", "Philadelphia, PA (US)", "Philadelphia PA", "CAR-巨噬细胞 (实体瘤、肝纤维化)", "CT-0508、CT-0525", "NASDAQ: CARM"),
    ("Century Therapeutics", "CGT (iPSC NK/T)", "Philadelphia, PA (US)", "Philadelphia PA", "iPSC 异体细胞 (CNTY-101, 自免)", "CNTY-101", "NASDAQ: IPSC"),
    ("Fate Therapeutics", "CGT (iPSC NK/T)", "San Diego, CA (US)", "San Diego CA", "iPSC NK/CAR-NK", "FT522、FT819", "NASDAQ: FATE"),
    ("Precigen", "CGT", "Germantown, MD (US)", "Germantown MD", "重症乳头瘤、肿瘤 UltraCAR-T", "PRGN-2012、PRGN-3006", "NASDAQ: PGEN"),
    ("Poseida Therapeutics → Roche", "CGT (已被收购)", "San Diego, CA (US)", "San Diego CA", "异体 CAR-T、肝 gene therapy", "P-MUC1C-ALLO1、P-BCMA-ALLO1", "—"),
    ("Cabaletta Bio", "CGT (CAR-T 自免)", "Philadelphia, PA (US)", "Philadelphia PA", "自免 CAAR-T/CAR-T (狼疮、肌炎)", "Resecabtagene autoleucel (CABA-201)", "NASDAQ: CABA"),
    ("Kyverna Therapeutics", "CGT (CAR-T 自免)", "Emeryville, CA (US)", "Emeryville CA", "自免 CD19 CAR-T (狼疮、MS)", "KYV-101、KYV-201", "NASDAQ: KYTX"),
    ("Capstan Therapeutics", "CGT (体内 CAR, LNP)", "San Diego, CA (US)", "San Diego CA", "体内 CAR-T、自免、纤维化", "CPTX2309", "私有 (Series B/C)"),
    ("Umoja Biopharma", "CGT (体内 CAR)", "Seattle, WA (US)", "Seattle WA; Louisville CO", "体内 CAR (VivoVec 病毒)", "UB-VV111", "私有"),
    ("Tessera Therapeutics", "CGT (基因写入)", "Somerville, MA (US)", "Somerville MA", "RNA 基因写入 (镰刀型/AAT)", "TSR-001 等管线", "私有 (Flagship)"),
    ("Rocket Pharmaceuticals", "CGT (基因治疗)", "Cranbury, NJ (US)", "Cranbury NJ", "罕见免疫缺陷、Danon、PKD", "Kresladi (LAD-I)、RP-A501", "NASDAQ: RCKT"),
    ("Voyager Therapeutics", "CGT (AAV 衣壳工程)", "Cambridge, MA (US)", "Cambridge MA", "ALS、Huntington、Friedreich", "VY-TAU01、VY-HTT01、VY7523 (与 Novartis)", "NASDAQ: VYGR"),
    ("Solid Biosciences", "CGT (DMD 基因治疗)", "Charlestown, MA (US)", "Charlestown MA", "DMD 微小蛋白、心脏基因治疗", "SGT-003、SGT-501", "NASDAQ: SLDB"),
    ("Taysha Gene Therapies", "CGT (基因治疗)", "Dallas, TX (US)", "Dallas TX", "Rett、罕见 CNS、GAN", "TSHA-102 (MECP2)、TSHA-120", "NASDAQ: TSHA"),
    ("REGENXBIO", "CGT (AAV 平台)", "Rockville, MD (US)", "Rockville MD", "AAV NAV 平台 (DMD、湿性 AMD、MPS)", "Clemidsogene (与 AbbVie)、RGX-202 (DMD)", "NASDAQ: RGNX"),
    ("Passage Bio", "CGT (CNS AAV)", "Philadelphia, PA (US)", "Philadelphia PA", "罕见 CNS (FTD-GRN、克拉伯)", "PBFT02、PBKR03", "NASDAQ: PASG"),
    ("Ultragenyx Pharmaceutical", "Specialty Biotech / CGT", "Novato, CA (US)", "Novato CA; Cambridge MA", "罕见代谢/骨/CNS、基因治疗", "Crysvita (US, 与 Kyowa Kirin)、Mepsevii、Dojolvi、UX111", "NASDAQ: RARE"),
    ("uniQure", "CGT (海外总部)", "Amsterdam, Netherlands", "Lexington MA", "Huntington、ALS、Refractory 癫痫", "AMT-130", "NASDAQ: QURE"),
    ("Vor Biopharma", "CGT (CRISPR + CAR-T)", "Cambridge, MA (US)", "Cambridge MA", "造血干细胞工程 + CAR-T", "Trem-cel、VCAR33", "NASDAQ: VOR"),
    ("Tenaya Therapeutics", "CGT (心脏)", "South SF, CA (US)", "South San Francisco CA", "遗传性心肌病基因治疗", "TN-201 (MYBPC3)、TN-401 (PKP2)", "NASDAQ: TNYA"),
    ("Generation Bio", "CGT (非病毒 ceDNA)", "Cambridge, MA (US)", "Cambridge MA; Waltham MA", "肝/免疫细胞 ceDNA + ctLNP", "GB-200/GB-300", "NASDAQ: GBIO"),

    # ============================================================
    # 6) RNA / mRNA / Oligo
    # ============================================================
    ("BioNTech US", "RNA Biotech (海外总部)", "Mainz, Germany", "Gaithersburg MD; Cambridge MA; Bay Area CA", "mRNA 疫苗、肿瘤个体化", "Comirnaty (与 Pfizer)、BNT122、BNT316", "NASDAQ: BNTX"),
    ("Arrowhead Pharmaceuticals", "RNA Biotech (RNAi 肝外)", "Pasadena, CA (US)", "Pasadena CA; San Diego CA; Madison WI", "RNAi (TG, ANGPTL3, AAT, INHBE)", "Plozasiran、Olpasiran (与 Amgen)、Zodasiran", "NASDAQ: ARWR"),
    ("Wave Life Sciences", "RNA Biotech (oligo)", "Cambridge, MA (US, ops)", "Cambridge MA", "ASO/siRNA (DMD、HD、AATD、肥胖)", "WVE-006、WVE-N531、WVE-007", "NASDAQ: WVE"),
    ("Stoke Therapeutics", "RNA Biotech (TANGO ASO)", "Bedford, MA (US)", "Bedford MA", "Dravet、ADOA、罕见 CNS", "Zorevunersen (STK-001)", "NASDAQ: STOK"),
    ("Korro Bio", "RNA Biotech (RNA 编辑)", "Cambridge, MA (US)", "Cambridge MA", "ADAR-导向 RNA 编辑", "KRRO-110 (AATD)", "NASDAQ: KRRO"),
    ("Acuitas Therapeutics", "LNP 递送平台", "Vancouver, Canada", "(LNP 授权 BioNTech/Pfizer/CureVac)", "脂质纳米颗粒 (LNP)", "授权 Comirnaty 等 LNP", "私有"),
    ("Translate Bio → Sanofi", "mRNA (已并入)", "Lexington, MA (US)", "Lexington MA", "mRNA 疫苗 (流感、SARS-CoV-2)", "MRT5500 等", "—"),
    ("Replicate Bioscience", "RNA (self-amp RNA)", "San Diego, CA (US)", "San Diego CA", "自扩增 RNA (肿瘤、传染病)", "RBI-1000/2000 等", "私有"),
    ("Generation Bio (上节) — 同时属 RNA/LNP", "—", "—", "—", "—", "—", "—"),

    # ============================================================
    # 7) Immunology / Dermatology / Ophthalmology
    # ============================================================
    ("Arcutis Biotherapeutics", "Immunology / Derm", "Westlake Village, CA (US)", "Westlake Village CA", "皮肤 (银屑病、AD、脂溢性皮炎)", "Zoryve (roflumilast)", "NASDAQ: ARQT"),
    ("Galderma", "Derm (海外总部)", "Zug, Switzerland", "Dallas TX; Fort Worth TX", "皮肤、医美", "Nemluvio (nemolizumab)、Restylane、Dysport", "SIX: GALD"),
    ("Dermavant Sciences (Roivant 旗下)", "Immunology / Derm", "Long Beach, CA (US)", "Long Beach CA; Durham NC", "皮肤 (银屑病、湿疹)", "Vtama (tapinarof)", "母公司 NASDAQ: ROIV"),
    ("Aclaris Therapeutics", "Immunology / Derm", "Wayne, PA (US)", "Wayne PA", "免疫炎症 (JAK/MK2)", "ATI-2138、Zunsemetinib", "NASDAQ: ACRS"),
    ("Acelyrin", "Immunology", "Agoura Hills, CA (US)", "Agoura Hills CA", "免疫炎症 (IL-17A 等)", "Izokibep、Lonigutamab", "NASDAQ: SLRN"),
    ("Vera Therapeutics", "Immunology (肾)", "Brisbane, CA (US)", "Brisbane CA", "肾免疫 (IgAN、膜性肾病)", "Atacicept", "NASDAQ: VERA"),
    ("Annexon", "Immunology / 神经", "Brisbane, CA (US)", "Brisbane CA", "经典补体 (GBS、AMD)", "Tanruprubart (ANX005)", "NASDAQ: ANNX"),
    ("Denali Therapeutics", "Neurology", "South SF, CA (US)", "South San Francisco CA", "ALS/PD/MPS (脑血屏障转运平台)", "DNL310 (Hunter)、DNL919 (TREM2)", "NASDAQ: DNLI"),
    ("Axsome Therapeutics", "Neurology", "New York, NY (US)", "New York NY", "抑郁、睡眠、AD 激越", "Auvelity、Sunosi", "NASDAQ: AXSM"),
    ("Cassava Sciences", "Neurology (AD, 已宣告失败)", "Austin, TX (US)", "Austin TX", "AD (simufilam, Phase 3 失败)", "Simufilam", "NASDAQ: SAVA"),
    ("Praxis Precision Medicines", "Neurology", "Boston, MA (US)", "Boston MA", "癫痫、震颤、运动障碍", "Ulixacaltamide、Vormatrigine、Relutrigine", "NASDAQ: PRAX"),
    ("Athira Pharma", "Neurology", "Bothell, WA (US)", "Bothell WA", "AD/PD (HGF/MET)", "Fosgonimeton", "NASDAQ: ATHA"),
    ("Prothena", "Neurology", "Dublin, Ireland", "Brisbane CA", "AD、AL 淀粉样、PD", "Birtamimab、Prasinezumab、PRX012", "NASDAQ: PRTA"),
    ("BioXcel Therapeutics", "Neurology / 精神", "New Haven, CT (US)", "New Haven CT", "精神/激越", "Igalmi (dexmedetomidine)", "NASDAQ: BTAI"),
    ("Tarsus Pharmaceuticals", "Ophth (见上)", "—", "—", "—", "—", "—"),
    ("Aldeyra Therapeutics", "Ophth / Immunology", "Lexington, MA (US)", "Lexington MA", "干眼、过敏性结膜炎、罕见免疫", "Reproxalap、ADX-2191", "NASDAQ: ALDX"),
    ("Belite Bio", "Ophth", "San Diego, CA (US)", "San Diego CA", "干性 AMD、Stargardt、糖尿病视网膜病", "Tinlarebant (LBS-008)", "NASDAQ: BLTE"),
    ("Outlook Therapeutics", "Ophth (biosimilar-ish)", "Iselin, NJ (US)", "Iselin NJ", "湿性 AMD 眼用 bevacizumab", "Lytenava", "NASDAQ: OTLK"),
    ("Ocular Therapeutix", "Ophth", "Bedford, MA (US)", "Bedford MA", "Wet AMD、DR、青光眼", "Axpaxli、Dextenza", "NASDAQ: OCUL"),
    ("Adverum Biotechnologies", "Ophth (CGT)", "Redwood City, CA (US)", "Redwood City CA", "湿性 AMD 基因治疗", "Ixo-vec", "NASDAQ: ADVM"),
    ("4D Molecular Therapeutics", "Ophth + CGT", "Emeryville, CA (US)", "Emeryville CA", "湿性 AMD、PAH、囊性纤维化", "4D-150、4D-310、4D-710", "NASDAQ: FDMT"),
    ("Opthea", "Ophth (海外总部)", "Melbourne, Australia", "Princeton NJ", "湿性 AMD (VEGF-C/D)", "Sozinibercept", "NASDAQ: OPT"),
    ("Larimar Therapeutics", "Rare disease", "Bala Cynwyd, PA (US)", "Bala Cynwyd PA", "Friedreich 共济失调", "Nomlabofusp (CTI-1601)", "NASDAQ: LRMR"),

    # ============================================================
    # 8) Infectious disease / Antimicrobials
    # ============================================================
    ("Vir Biotechnology", "ID Biotech", "San Francisco, CA (US)", "San Francisco CA; Portland ME", "HBV/D、流感、HDV、肿瘤", "Tobevibart、Elebsiran、Sotrovimab", "NASDAQ: VIR"),
    ("Atea Pharmaceuticals", "ID Biotech", "Boston, MA (US)", "Boston MA", "HCV、COVID 抗病毒", "Bemnifosbuvir + ruzasvir", "NASDAQ: AVIR"),
    ("Cidara Therapeutics", "ID Biotech", "San Diego, CA (US)", "San Diego CA", "流感、抗真菌", "Rezzayo、CD388", "NASDAQ: CDTX"),
    ("Paratek Pharmaceuticals (Private)", "ID Biotech", "King of Prussia, PA (US)", "King of Prussia PA", "抗生素", "Nuzyra (omadacycline)", "私有"),
    ("Spero Therapeutics", "ID Biotech", "Cambridge, MA (US)", "Cambridge MA", "复杂 UTI、肺非结核分枝杆菌", "Tebipenem HBr、SPR720", "NASDAQ: SPRO"),
    ("Melinta Therapeutics", "ID Biotech", "Parsippany, NJ (US)", "Parsippany NJ", "抗感染", "Baxdela、Vabomere、Kimyrsa", "私有"),

    # ============================================================
    # 9) AI Drug Discovery / Platform
    # ============================================================
    ("Recursion Pharmaceuticals", "AI Discovery", "Salt Lake City, UT (US)", "Salt Lake City UT; Toronto", "图像 + AI 表型药物发现", "REC-994、REC-2282；与 Roche/Bayer 合作", "NASDAQ: RXRX"),
    ("Schrödinger", "AI Discovery + Software", "New York, NY (US)", "NYC; Cambridge MA", "物理 + ML 计算化学，软件 + 自研管线", "SGR-1505 (MALT1)、SGR-2921 (CDC7)", "NASDAQ: SDGR"),
    ("Relay Therapeutics", "AI Discovery (蛋白动态)", "Cambridge, MA (US)", "Cambridge MA", "动态构象药物 (FGFR2, PI3Kα)", "RLY-4008 (lirafugratinib)、RLY-2608", "NASDAQ: RLAY"),
    ("Insitro", "AI Discovery", "South SF, CA (US)", "South San Francisco CA", "ML + 干细胞 + 真实世界", "ALS、MASH 与 BMS/Lilly 合作", "私有 (a16z)"),
    ("AbCellera Biologics", "AI Discovery (抗体)", "Vancouver, Canada", "Bay Area + Boston (临床扩展)", "AI 驱动抗体发现 (T 细胞引擎、靶点)", "ABCL575 (OX40L)、ABCL635", "NASDAQ: ABCL"),
    ("Nimbus Therapeutics", "AI Discovery", "Cambridge, MA (US)", "Cambridge MA", "计算化学 (TYK2、HPK1、CTPS1)", "NDI-101150 (HPK1)、NDI-219216 (CTPS1)", "私有"),
    ("Generate:Biomedicines", "AI Discovery (生成蛋白)", "Somerville, MA (US)", "Somerville MA", "生成式 AI 蛋白设计 (抗体、肽)", "GB-0669 (RSV)、GB-0895 (TL1A)", "私有 (Flagship)"),
    ("Cellarity", "AI Discovery (细胞行为)", "Somerville, MA (US)", "Somerville MA", "细胞行为模型 + 药物表型", "CLY-124 (SCD)、肥胖管线", "私有 (Flagship)"),
    ("Valo Health", "AI Discovery", "Boston, MA (US)", "Boston MA; Lexington MA; Branford CT", "AI + 真实世界数据 (心血管、AD)", "OPL-0301、OPL-0401", "私有"),
    ("Insilico Medicine", "AI Discovery (海外总部)", "New York / Hong Kong / Shanghai", "New York NY (US ops)", "AI 端到端发现 (IPF、肿瘤)", "Rentosertib (ISM001-055)、ISM5939", "私有"),
    ("Exscientia → Recursion", "AI Discovery (已合并)", "Oxford, UK", "Cambridge MA; Bay Area", "AI 小分子设计 (LSD1, A2A, GTAEXS-617)", "已并入 Recursion (2024)", "—"),
    ("Verge Genomics", "AI Discovery", "South SF, CA (US)", "South San Francisco CA", "ALS、PD、神经退行性", "VRG50635、VRG-LRRK2", "私有"),
    ("Iambic Therapeutics", "AI Discovery", "San Diego, CA (US)", "San Diego CA", "AI 物理化学药物设计 (肿瘤)", "IAM1363 (HER2)", "私有"),
    ("Cradle Bio", "AI Discovery (蛋白工程平台)", "Amsterdam, Netherlands", "(SaaS 全球客户, 含美国)", "生成式蛋白设计 SaaS", "授权多家 biotech", "私有"),
    ("Owkin", "AI / RWD", "Paris, France", "New York, NY (US ops)", "联邦学习 + 病理 AI", "MSIntuit CRC、肿瘤 RWD 合作", "私有"),
    ("Tempus AI", "AI / Diagnostics", "Chicago, IL (US)", "Chicago IL; RTP NC", "肿瘤基因组学 + AI、心脏 ECG-AI", "xT、xR、xF+、ECG-AF", "NASDAQ: TEM"),
    ("Ginkgo Bioworks", "Synth Bio + Bioservices", "Boston, MA (US)", "Boston MA", "细胞编程、CDMO、生物安全", "Foundry + Biosecurity 平台", "NYSE: DNA"),

    # ============================================================
    # 10) Diagnostics / Tools / Sequencing (Life Science Tools)
    # ============================================================
    ("Illumina", "Tools (NGS)", "San Diego, CA (US)", "San Diego CA; Foster City CA; RTP NC", "NGS 测序仪 + 试剂", "NovaSeq X+、MiSeq i100、TruSight Oncology", "NASDAQ: ILMN"),
    ("Pacific Biosciences (PacBio)", "Tools (长读长)", "Menlo Park, CA (US)", "Menlo Park CA", "HiFi 长读长测序", "Revio、Vega、Onso", "NASDAQ: PACB"),
    ("10x Genomics", "Tools (单细胞/空间)", "Pleasanton, CA (US)", "Pleasanton CA", "单细胞、空间组学", "Chromium、Visium、Xenium", "NASDAQ: TXG"),
    ("Bio-Rad Laboratories", "Tools / Diagnostics", "Hercules, CA (US)", "Hercules CA; Pleasanton CA", "ddPCR、电泳、临床诊断", "QX600/QX200、ChemiDoc、临床免疫", "NYSE: BIO"),
    ("Bio-Techne", "Tools / Diagnostics", "Minneapolis, MN (US)", "Minneapolis MN; San Jose CA", "蛋白工具、ProteinSimple、Tula 自动化", "Ella、Wes、ExoTRU", "NASDAQ: TECH"),
    ("Twist Bioscience", "Tools (合成 DNA)", "South SF, CA (US)", "South San Francisco CA; Wilsonville OR", "基因合成、抗体库、NGS 靶向", "Twist 基因合成、Twist Bio 药物发现", "NASDAQ: TWST"),
    ("Element Biosciences", "Tools (NGS)", "San Diego, CA (US)", "San Diego CA", "AVITI 测序", "AVITI、Cloudbreak 化学", "私有 (planning IPO)"),
    ("Singular Genomics", "Tools (NGS)", "San Diego, CA (US)", "San Diego CA", "G4 测序 + 空间多组学", "G4 / G4X", "私有 (Deerfield 2024 收购)"),
    ("Olink (Thermo Fisher)", "Tools (Proteomics)", "Uppsala, Sweden", "Watertown MA; Boston MA", "邻位扩增蛋白组学 (PEA)", "Olink Explore、Reveal", "—"),
    ("Standard BioTools", "Tools", "South SF, CA (US)", "South San Francisco CA", "质谱流式 CyTOF、Biomark 微流控", "CyTOF XT、X9 Real-Time", "NASDAQ: LAB"),
    ("Bionano Genomics", "Tools", "San Diego, CA (US)", "San Diego CA", "光学基因组图谱 (OGM)", "Saphyr、Stratys", "NASDAQ: BNGO"),
    ("Personalis", "Diagnostics", "Fremont, CA (US)", "Fremont CA", "全外显子组、MRD、ctDNA", "NeXT Personal、NeXT Dx", "NASDAQ: PSNL"),
    ("Castle Biosciences", "Diagnostics", "Friendswood, TX (US)", "Friendswood TX; Pittsburgh PA", "皮肤癌、Barrett、心理基因检测", "DecisionDx-Melanoma、TissueCypher", "NASDAQ: CSTL"),
    ("Natera", "Diagnostics", "Austin, TX (US)", "Austin TX; San Carlos CA", "NIPT、肿瘤 MRD、移植", "Signatera、Panorama、Empower", "NASDAQ: NTRA"),
    ("Exact Sciences", "Diagnostics", "Madison, WI (US)", "Madison WI; RTP NC", "结肠癌筛查、MRD、肿瘤", "Cologuard、Oncotype DX", "NASDAQ: EXAS"),
    ("Guardant Health", "Diagnostics", "Redwood City, CA (US)", "Redwood City CA", "肿瘤液体活检、CRC 筛查", "Guardant360、Reveal、Shield", "NASDAQ: GH"),
    ("Veracyte", "Diagnostics", "South SF, CA (US)", "South San Francisco CA", "甲状腺/肺/前列腺基因表达分类", "Afirma、Decipher、Percepta", "NASDAQ: VCYT"),
    ("Foundation Medicine (Roche)", "Diagnostics", "Cambridge, MA (US)", "Cambridge MA; RTP NC", "肿瘤基因组学 (CDx)", "FoundationOne CDx、Liquid CDx、Heme", "母公司 SIX: ROG"),
    ("Caris Life Sciences", "Diagnostics", "Irving, TX (US)", "Irving TX; Phoenix AZ", "肿瘤全外、转录组、ADAPT 生物标志", "MI Profile、MI Cancer Seek、Folio", "(私有, 2025 IPO 计划)"),
    ("Adaptive Biotechnologies", "Diagnostics / Tools", "Seattle, WA (US)", "Seattle WA", "TCR/BCR 测序 + MRD", "clonoSEQ、T-Detect", "NASDAQ: ADPT"),
    ("Hologic", "Diagnostics (女性健康)", "Marlborough, MA (US)", "Marlborough MA; San Diego CA", "分子诊断、影像、外科", "Aptima、Panther、ThinPrep", "NASDAQ: HOLX"),
    ("QuidelOrtho", "Diagnostics", "San Diego, CA (US)", "San Diego CA; Raritan NJ", "POCT、免疫诊断、血型", "Sofia、Triage、Vitros", "NASDAQ: QDEL"),
    ("Qiagen", "Tools / Diagnostics", "Hilden, Germany", "Germantown MD; Beverly MA", "样品制备、PCR、QuantiFERON", "QIAcuity、QuantiFERON-TB", "NYSE: QGEN"),
    ("Danaher Diagnostics (Beckman/Cepheid/Leica/Aldevron)", "Tools / Diagnostics", "Washington, DC (US)", "Brea CA; Sunnyvale CA; Fargo ND", "临床诊断、生命科学、CDMO", "Cepheid GeneXpert、Aldevron 质粒/mRNA", "NYSE: DHR"),
    ("Thermo Fisher Scientific", "Tools / CDMO", "Waltham, MA (US)", "Waltham MA; Carlsbad CA; Wilmington NC", "试剂仪器、Patheon CDMO、PPD CRO", "Ion Torrent、Patheon、PPD、Olink", "NYSE: TMO"),
    ("Agilent Technologies", "Tools / Diagnostics", "Santa Clara, CA (US)", "Santa Clara CA; Lexington MA", "色谱质谱、CDx、NGS 试剂", "SureSelect、Dako PD-L1", "NYSE: A"),
    ("Waters Corporation", "Tools", "Milford, MA (US)", "Milford MA; New Castle DE", "LC/MS、热分析", "Xevo、Acquity", "NYSE: WAT"),

    # ============================================================
    # 11) CRO (Clinical Research)
    # ============================================================
    ("IQVIA", "CRO (Top 1)", "Durham, NC (US)", "Durham NC; Cambridge MA; Plymouth Meeting PA", "全球临床研究、RWD、商业化", "IQVIA CORE 平台、Q² Solutions 中心实验室", "NYSE: IQV"),
    ("ICON plc", "CRO (Top)", "Dublin, Ireland", "Blue Bell PA; North Wales PA; Boston MA", "全球临床、医学影像、PRA Health 整合", "ICON Symphony", "NASDAQ: ICLR"),
    ("Fortrea", "CRO (Labcorp 分拆)", "Durham, NC (US)", "Durham NC", "Phase I-IV、中心实验室", "2023 年从 Labcorp 拆分", "NASDAQ: FTRE"),
    ("Charles River Laboratories", "CRO (临床前 + CDMO)", "Wilmington, MA (US)", "Wilmington MA; Frederick MD; Memphis TN", "动物模型、安全性、CGT CDMO", "RMS、DSA、Manufacturing", "NYSE: CRL"),
    ("Medpace", "CRO (Full-service)", "Cincinnati, OH (US)", "Cincinnati OH", "肿瘤/心血管/代谢/罕见病", "纯有机增长、利润率领先", "NASDAQ: MEDP"),
    ("Parexel", "CRO", "Durham, NC (US)", "Durham NC; Newton MA", "全球临床、监管咨询", "(EQT + Goldman 私有)", "私有"),
    ("Syneos Health", "CRO + CCO", "Morrisville, NC (US)", "Morrisville NC", "临床 + 商业化合同销售", "(Elliott + Patient Square + Veritas 私有化)", "私有"),
    ("PPD (Thermo Fisher Clinical Research)", "CRO", "Wilmington, NC (US)", "Wilmington NC; Highland Heights KY", "Phase I-IV、实验室 (PPD Labs)", "(2021 被 Thermo Fisher 收购)", "母公司 NYSE: TMO"),
    ("WCG Clinical", "CRO 服务 (IRB/伦理)", "Princeton, NJ (US)", "Princeton NJ; Cary NC", "IRB、伦理、 patient-centric 服务", "WIRB-Copernicus 等品牌", "私有"),
    ("Premier Research", "CRO (中型)", "Morrisville, NC (US)", "Morrisville NC", "罕见病、儿科、生物科技", "(Charlesbank Capital 私有)", "私有"),
    ("Worldwide Clinical Trials", "CRO (中型)", "Morrisville, NC (US)", "Morrisville NC", "CNS、心血管、罕见病", "(Kohlberg/Cinven 私有)", "私有"),
    ("Veristat", "CRO (Boutique)", "Southborough, MA (US)", "Southborough MA", "罕见病、肿瘤、生物制品、监管", "—", "私有"),
    ("Advanced Clinical", "CRO + Staffing", "Bannockburn, IL (US)", "Bannockburn IL", "FSP / Functional Service Provision", "—", "私有"),
    ("Inotiv", "CRO (临床前 + 模型动物)", "West Lafayette, IN (US)", "West Lafayette IN; Mt. Vernon IN", "DMPK、毒理、模型动物 (Envigo 收购后)", "—", "NASDAQ: NOTV"),
    ("Linical", "CRO (中型)", "Osaka, Japan", "Newtown PA (US ops)", "肿瘤、神经", "—", "TYO: 2183"),
    ("Novotech", "CRO (亚太/全球)", "Sydney, Australia", "Boston MA; San Diego CA", "亚太临床、肿瘤、罕见病", "(TPG 持有 + IPO 多次)", "私有"),
    ("ProPharma Group", "CRO + Regulatory", "Raleigh, NC / Overland Park KS (US)", "Raleigh NC; Overland Park KS", "MA + 监管 + 药物警戒", "—", "私有"),
    ("Caidya (Clinipace + dMed 合并)", "CRO (中型)", "Morrisville, NC (US)", "Morrisville NC", "全球中型临床", "—", "私有"),
    ("PSI CRO", "CRO (海外总部)", "Zug, Switzerland", "Durham NC (US ops)", "Phase II-IV", "—", "私有"),
    ("Frontage Holdings", "CRO + 临床前", "Exton, PA (US)", "Exton PA; Concord OH", "DMPK、毒理、生物分析、临床", "—", "HKEX: 1521"),
    ("Altasciences", "CRO (临床前 + Phase I)", "Laval, Canada", "Mt. Vernon NY; Columbia MO; Philadelphia PA", "临床前 + EAP + Phase I", "—", "私有"),
    ("BioAgilytix", "CRO (生物分析)", "Durham, NC (US)", "Durham NC; Boston MA; Cambridge MA", "生物分析 + 免疫原性", "—", "私有"),
    ("Catalyst Clinical Research", "CRO (中型, FSP)", "Raleigh, NC (US)", "Raleigh NC", "全球临床、FSP", "—", "私有"),

    # ============================================================
    # 12) CDMO (合同制造)
    # ============================================================
    ("Lonza", "CDMO (Top, 海外总部)", "Basel, Switzerland", "Portsmouth NH; Hayward CA; Visp/Hopkinton MA", "生物制品 + 小分子 + CGT CDMO (Vacaville 收购)", "Ibex Biopark", "SIX: LONN"),
    ("Catalent", "CDMO", "Somerset, NJ (US)", "Bloomington IN; Madison WI; Morrisville NC", "无菌灌装、CGT、OSD、Softgel (Novo Holdings 收购)", "—", "(被 Novo Holdings 私有化)"),
    ("Thermo Fisher Pharma Services / Patheon", "CDMO", "Waltham, MA (US)", "Greenville NC; Cincinnati OH; Florence SC", "API、生物制品、无菌灌装", "—", "NYSE: TMO"),
    ("Samsung Biologics", "CDMO (海外总部)", "Incheon, Korea", "(2024 北美商务办公室)", "生物制品 (大规模)", "—", "KRX: 207940"),
    ("WuXi Biologics", "CDMO (海外总部)", "Shanghai, China / Hong Kong", "Worcester MA; Cranbury NJ", "生物制品 (单抗、ADC、双抗)", "—", "HKEX: 2269"),
    ("WuXi AppTec", "CDMO + 临床前 (海外总部)", "Shanghai, China", "Plymouth Meeting PA; San Diego CA; Atlanta GA", "化学、生物分析、CGT (Advanced Therapies)", "—", "HKEX: 2359 / SHA: 603259"),
    ("Resilience (NTGI)", "CDMO", "San Diego, CA (US)", "San Diego CA; West Chester OH; Boston MA", "生物制品、CGT、mRNA 大规模制造", "—", "私有"),
    ("National Resilience", "CDMO (上行)", "—", "—", "—", "—", "—"),
    ("Avid Bioservices", "CDMO", "Tustin, CA (US)", "Tustin CA", "中型生物制品", "—", "(GHO Capital 私有化)"),
    ("KBI Biopharma (JSR)", "CDMO", "Durham, NC (US)", "Durham NC; Boulder CO; The Woodlands TX", "生物制品 + STA Pharmaceutical 合作", "—", "私有 (JSR)"),
    ("AGC Biologics", "CDMO", "Bothell, WA (US)", "Bothell WA; Boulder CO; Longmont CO", "生物制品、CGT、质粒", "—", "母公司 TYO: 5201"),
    ("FUJIFILM Diosynth Biotechnologies", "CDMO", "Tokyo, Japan", "Research Triangle Park NC; College Station TX; Holly Springs NC", "生物制品、病毒载体、mRNA", "—", "TYO: 4901"),
    ("Cytovance Biologics", "CDMO", "Oklahoma City, OK (US)", "Oklahoma City OK", "微生物 + 哺乳动物细胞", "—", "私有"),
    ("Piramal Pharma Solutions", "CDMO", "Mumbai, India", "Lexington KY; Bethlehem PA; Riverview MI", "API、注射剂、ADC", "—", "NSE: PPLPHARMA"),
    ("Recipharm", "CDMO (海外总部)", "Stockholm, Sweden", "Research Triangle Park NC", "无菌灌装、OSD、吸入", "—", "(EQT/Roar BidCo 私有)"),
    ("Cambrex", "CDMO (API)", "East Rutherford, NJ (US)", "East Rutherford NJ; Charles City IA", "API、HPAPI、控释固体制剂", "—", "私有"),
    ("Aldevron (Danaher)", "CDMO (质粒/mRNA/蛋白)", "Fargo, ND (US)", "Fargo ND; Madison WI; Miami FL", "质粒 DNA、mRNA、蛋白", "—", "母公司 NYSE: DHR"),
    ("Vetter Pharma", "CDMO (海外, 无菌灌装)", "Ravensburg, Germany", "Skokie IL (US ops)", "无菌预灌装注射剂", "—", "私有"),

    # ============================================================
    # 13) Animal Health (US-active)
    # ============================================================
    ("Zoetis", "Animal Health (Top)", "Parsippany, NJ (US)", "Parsippany NJ; Kalamazoo MI", "宠物 + 牲畜", "Apoquel、Cytopoint、Librela、Solensia", "NYSE: ZTS"),
    ("Elanco Animal Health", "Animal Health", "Greenfield, IN (US)", "Greenfield IN", "宠物 + 牲畜", "Credelio、Galliprant、Experior、Increxxa", "NYSE: ELAN"),
    ("Boehringer Ingelheim Animal Health", "Animal Health", "Ingelheim, Germany (private)", "Duluth GA; St. Joseph MO; Athens GA", "宠物 + 牲畜", "NexGard、Heartgard、Frontline", "私有"),
    ("Merck Animal Health", "Animal Health", "Madison, NJ (US)", "Madison NJ; De Soto KS", "宠物 + 牲畜 + 水产", "Bravecto、Nobivac、Sentinel", "(母公司 NYSE: MRK)"),
    ("Phibro Animal Health", "Animal Health", "Teaneck, NJ (US)", "Teaneck NJ", "营养添加剂、疫苗、抗感染", "Stafac、Mecadox、AviPro", "NASDAQ: PAHC"),
    ("Vetoquinol USA", "Animal Health (海外总部)", "Lure, France", "Fort Worth TX", "宠物 + 牲畜", "Zylkene、Felpreva、Profender", "EPA: VETO"),
    ("Dechra Pharmaceuticals", "Animal Health", "Northwich, UK", "Overland Park KS", "宠物 + 牲畜", "Vetoryl、Mirataz", "(EQT 私有化)"),

    # ============================================================
    # 14) Generics / Biosimilars (US presence)
    # ============================================================
    ("Sun Pharmaceutical Industries", "Generics + Specialty", "Mumbai, India", "Princeton NJ; Cranbury NJ", "皮肤、眼、肿瘤、特种仿制", "Ilumya、Cequa、Winlevi、Odomzo", "NSE: SUNPHARMA"),
    ("Dr. Reddy's Laboratories", "Generics + Biosimilars", "Hyderabad, India", "Princeton NJ", "复杂仿制、生物类似药、OTC", "Rituximab biosimilar、Tobramycin", "NYSE: RDY"),
    ("Cipla", "Generics", "Mumbai, India", "Warren NJ; Central Islip NY", "呼吸、抗逆转录病毒、复杂仿制", "Albuterol HFA、Lamivudine", "NSE: CIPLA"),
    ("Lupin Limited", "Generics", "Mumbai, India", "Baltimore MD; Coral Springs FL", "呼吸、神经、心血管、注射", "Solosec、Suprep", "NSE: LUPIN"),
    ("Aurobindo Pharma", "Generics + Biosimilars", "Hyderabad, India", "East Windsor NJ; Durham NC", "复杂仿制、无菌注射、生物类似药", "—", "NSE: AUROPHARMA"),
    ("Zydus Lifesciences", "Generics + Specialty", "Ahmedabad, India", "Pennington NJ", "皮肤、肿瘤、生物类似药", "Mirum、Saroglitazar", "NSE: ZYDUSLIFE"),
    ("Glenmark Pharmaceuticals", "Generics", "Mumbai, India", "Mahwah NJ", "呼吸、皮肤、肿瘤", "Ryaltris、Envafolimab (与 Alphamab)", "NSE: GLENMARK"),
    ("Amneal Pharmaceuticals", "Generics + Specialty", "Bridgewater, NJ (US)", "Bridgewater NJ; Hayward CA", "复杂仿制、注射、品牌 specialty", "Crexont、Lyvispah、Naloxone", "NYSE: AMRX"),
    ("Apotex", "Generics", "Toronto, Canada", "Weston FL", "广谱仿制", "—", "私有"),
    ("Alvogen", "Generics", "Pine Brook, NJ (US)", "Pine Brook NJ", "复杂仿制、特殊剂型", "—", "私有 (CVC)"),
    ("Padagis", "Generics", "Allegan, MI (US)", "Allegan MI", "外用、注射、OTC 自有品牌", "—", "私有 (前 Perrigo Rx)"),
    ("Hetero USA", "Generics (海外总部)", "Hyderabad, India", "Marietta GA; Princeton NJ", "复杂仿制、抗病毒、注射", "—", "私有"),
    ("Granules USA", "Generics (海外总部)", "Hyderabad, India", "Chantilly VA; Princeton NJ", "API + 仿制制剂", "—", "NSE: GRANULES"),
    ("Lannett", "Generics", "Trevose, PA (US)", "Trevose PA", "复杂仿制", "—", "私有 (重组后)"),
    ("ANI Pharmaceuticals", "Specialty + Generics", "Princeton, NJ (US)", "Princeton NJ; Baudette MN", "罕见病 + 注射 + 仿制", "Cortrophin Gel、Iluvien", "NASDAQ: ANIP"),

    # ============================================================
    # 15) 还可能用得到的: 大型工具公司 + Biotech 服务
    # ============================================================
    ("Bristol Myers Squibb 旗下 Celgene (历史)", "—", "—", "—", "—", "已并入 BMS, 仅供参考", "—"),
    ("Roivant Sciences", "Vant 母公司 / 多个子公司在美国招", "London, UK", "New York NY; Durham NC", "多 Vant 平台 (Dermavant, Immunovant, Pulmovant, etc.)", "Vtama、Batoclimab、Brepocitinib", "NASDAQ: ROIV"),
    ("Immunovant", "Immunology Biotech (Roivant Vant)", "New York, NY (US)", "New York NY", "FcRn 抑制剂 (MG、TED、IgAN)", "Batoclimab、IMVT-1402", "NASDAQ: IMVT"),
    ("Summit Therapeutics", "Onc Biotech", "Miami, FL (US)", "Miami FL; Cambridge UK", "肺癌 (PD-1/VEGF 双抗, 与 Akeso)", "Ivonescimab (US 权利)", "NASDAQ: SMMT"),
    ("Coherus BioSciences (上节) — 生物类似药/Loqtorzi 已收录", "—", "—", "—", "—", "—", "—"),
    ("Ascendis (上节)", "—", "—", "—", "—", "—", "—"),
    ("AbbVie 已收录", "—", "—", "—", "—", "—", "—"),

    # 一些重要的临床阶段 biotech (找工作常见)
    ("Replimune (上节)", "—", "—", "—", "—", "—", "—"),
    ("Krystal (上节)", "—", "—", "—", "—", "—", "—"),
    ("Cogent (上节)", "—", "—", "—", "—", "—", "—"),
    ("Disc Medicine (上节)", "—", "—", "—", "—", "—", "—"),
    ("Apellis (上节)", "—", "—", "—", "—", "—", "—"),
    ("Ardelyx (上节)", "—", "—", "—", "—", "—", "—"),

    # 补充: 体外受精/女性健康 / 罕见
    ("Travere (上节)", "—", "—", "—", "—", "—", "—"),
    ("Acadia (上节)", "—", "—", "—", "—", "—", "—"),

    # 补充: 大型工具
    ("Revvity (前 PerkinElmer 拆分)", "Tools / Diagnostics", "Waltham, MA (US)", "Waltham MA; Hopkinton MA", "诊断 + 生命科学", "新生儿筛查、Horizon Discovery、Dharmacon", "NYSE: RVTY"),
    ("MaxCyte", "Tools (电穿孔 CGT 平台)", "Rockville, MD (US)", "Rockville MD", "电穿孔细胞工程平台 (CGT)", "ExPERT GTx、STx", "NASDAQ: MXCT"),
    ("Repligen", "Tools (生物工艺)", "Waltham, MA (US)", "Waltham MA", "蛋白纯化、过滤、生物工艺", "OPUS、KrosFlo、CTech", "NASDAQ: RGEN"),
    ("Maravai LifeSciences", "Tools (mRNA 原料)", "San Diego, CA (US)", "San Diego CA; Carlsbad CA", "CleanCap mRNA 加帽、TriLink", "CleanCap、寡核苷酸", "NASDAQ: MRVI"),
    ("Stevanato Group", "Tools (玻璃容器/灌装)", "Padua, Italy", "Boston/Boston-area MA; Bridgewater NJ (US ops)", "玻璃西林瓶、注射器、自动化", "EZ-fill、Vio-Pro", "NYSE: STVN"),
    ("West Pharmaceutical Services", "Tools (容器/封装)", "Exton, PA (US)", "Exton PA; Williamsport PA", "弹性体瓶塞、自注射 SmartDose", "Daikyo Crystal Zenith、Vial2Bag", "NYSE: WST"),
    ("Charles River 已收录", "—", "—", "—", "—", "—", "—"),
]

# 去重: 跳过 (`—`占位) 行 + 跳过 "见上 / 上节 / 已收录" 等占位
def _is_placeholder(row):
    if row[1] == "—":
        return True
    blob = " ".join(str(x) for x in row)
    return any(tok in blob for tok in ("见上", "上节", "已收录", "已并入", "已被收购"))

COMPANIES = [c for c in COMPANIES if not _is_placeholder(c)]
# 去重同名
_seen = set()
_uniq = []
for c in COMPANIES:
    if c[0] in _seen:
        continue
    _seen.add(c[0])
    _uniq.append(c)
COMPANIES = _uniq


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def build_sheet(rows, header):
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
        '<cols>',
        '<col min="1" max="1" width="42"/>',
        '<col min="2" max="2" width="22"/>',
        '<col min="3" max="3" width="28"/>',
        '<col min="4" max="4" width="46"/>',
        '<col min="5" max="5" width="44"/>',
        '<col min="6" max="6" width="60"/>',
        '<col min="7" max="7" width="20"/>',
        '</cols>',
        '<sheetData>',
    ]
    # header row
    parts.append('<row r="1">')
    for c_idx, val in enumerate(header, start=1):
        ref = f"{col_letter(c_idx)}1"
        parts.append(
            f'<c r="{ref}" t="inlineStr" s="1"><is><t xml:space="preserve">{escape(str(val))}</t></is></c>'
        )
    parts.append('</row>')
    # data rows
    for r_idx, row in enumerate(rows, start=2):
        parts.append(f'<row r="{r_idx}">')
        for c_idx, val in enumerate(row, start=1):
            ref = f"{col_letter(c_idx)}{r_idx}"
            parts.append(
                f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{escape(str(val))}</t></is></c>'
            )
        parts.append('</row>')
    parts.append('</sheetData>')
    last_col = col_letter(len(header))
    parts.append(f'<autoFilter ref="A1:{last_col}1"/>')
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
        z.writestr("xl/worksheets/sheet1.xml", build_sheet(COMPANIES, HEADERS))


if __name__ == "__main__":
    build_xlsx("us-pharma-biotech-cro.xlsx")
    print(f"Wrote us-pharma-biotech-cro.xlsx with {len(COMPANIES)} companies")

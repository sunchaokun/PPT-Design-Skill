window.PPT_CASES = [
  {
    id: 'ai-agent-operating-system',
    title: 'AI Agent Operating System',
    subtitle: '技术编辑型 · 从提示链到受控自治',
    category: '技术叙事',
    year: '2026',
    direction: '深色系统蓝图',
    mode: 'Build Mode',
    description: '一组以系统架构为核心的技术编辑型演示，用深色网格、荧光色线索和分层流程表达 AI Agent 的运行机制。',
    path: 'assets/ai-agent-operating-system',
    pptx: 'downloads/ai_agent_operating_system.pptx',
    pdf: 'downloads/ai_agent_operating_system.pdf',
    slides: ['系统蓝图', '用户意图', '控制平面', '执行平面', '状态与记忆', '工具层', '评估循环', '安全边界', '架构选择', '决策框架', '最小系统', '结论']
  },
  {
    id: 'ai-infrastructure-economics',
    title: 'AI Infrastructure Economics',
    subtitle: '数据编辑型 · 资本、算力与物理堆栈',
    category: '数据叙事',
    year: '2026',
    direction: '纸张与基础设施账本',
    mode: 'Build Mode',
    description: '一组面向战略与平台运营者的来源型数据研究，以纸张质感、层级色彩和物理约束隐喻组织复杂的基础设施叙事。',
    path: 'assets/ai-infrastructure-economics',
    pptx: 'downloads/ai_infrastructure_economics.pptx',
    pdf: 'downloads/ai_infrastructure_economics.pdf',
    slides: ['基础设施账本', '价值链', '资本流', '算力约束', '平台层', '能源与地点', '供给侧', '成本结构', '竞争位置', '风险边界', '行动路径', '结论']
  },
  {
    id: 'car-t-single-cell-atlas',
    title: 'Single-Cell CAR T Atlas',
    subtitle: '学术研究型 · 从临床持续性到细胞状态假设',
    category: '科学叙事',
    year: '2026',
    direction: 'BioTech Blue Research Editorial',
    mode: 'Build Mode',
    description: '基于 Nature 论文的 12 页研究案例，以图证、研究设计和证据边界组织 CAR T 细胞持续性、状态机制与体内验证。',
    path: 'assets/car-t-single-cell-atlas',
    pptx: 'downloads/car_t_single_cell_atlas_blue_editorial.pptx',
    pdf: 'downloads/car_t_single_cell_atlas_blue_editorial.pdf',
    slides: ['研究问题', '临床持续性', '研究设计', '细胞图谱', '持续性梯度', '2 型状态', '候选机制', '血清轨迹', '体内验证', '工程杠杆', '证据链', '结论与边界']
  },
  {
    id: 'louvre-abudhabi',
    title: 'Louvre Abu Dhabi',
    subtitle: '建筑文化型 · 光、水与博物馆城',
    category: '空间叙事',
    year: '2026',
    direction: 'Nocturnal Architectural Editorial',
    mode: 'Build Mode',
    description: '以真实建筑摄影、可编辑几何图形与官方建筑数据组织的 10 页文化建筑案例，呈现 Jean Nouvel 的光、气候与公共空间语言。',
    path: 'assets/louvre-abudhabi',
    pptx: 'downloads/louvre_abudhabi_complete.pptx',
    pdf: 'downloads/louvre_abudhabi_complete.pdf',
    slides: ['封面', '光之雨', '几何与气候', '海上博物馆城', '水作为公共空间', '热环境', '空间序列', '材料逻辑', '建筑立场', '结语']
  },
  {
    id: 'vertical-city-retrofit',
    title: 'Vertical City Retrofit',
    subtitle: '城市策略型 · 在地更新与系统投资',
    category: '空间叙事',
    year: '2026',
    direction: 'Warm Architectural Editorial',
    mode: 'Build Mode',
    description: '面向城市更新、建筑与投资决策者的 14 页策略案例，以建筑剖面、系统图、情景数据和治理框架定义高层住宅改造。',
    path: 'assets/vertical-city-retrofit',
    pptx: 'downloads/vertical_city_retrofit_style_sample.pptx',
    pdf: 'downloads/vertical_city_retrofit_style_sample.pdf',
    slides: ['封面', '显性问题', '五个系统', '改造路径', '能源约束', '共享底层', '日常生活', '改造类型', '在住施工', '顺序经济学', '决策矩阵', '运营模型', '36 个月路径', '结论']
  },
  {
    id: 'couture-color-objects-of-desire',
    title: 'COUTURE COLOR',
    subtitle: 'Luxury Beauty Editorial · Objects of Desire',
    category: '视觉叙事',
    year: '2026',
    direction: 'Nocturnal Beauty Editorial',
    mode: 'Build Mode',
    description: '以全屏妆效肖像、同一模特的上妆动作、可编辑产品结构与材质叙事，构成一组 10 页高定口红作品集案例。',
    path: 'assets/couture-color-objects-of-desire',
    pptx: 'downloads/couture_color_objects_of_desire.pptx',
    pdf: 'downloads/couture_color_objects_of_desire.pdf',
    slides: ['Couture Color', 'The Proposition', 'Chromatic Wardrobe', 'Object Grammar', 'Material Choreography', 'The Gesture', 'The Collection', 'The Ritual', 'Editorial Release', 'Coda']
  }
].map(project => ({
  ...project,
  slides: project.slides.map((label, index) => ({
    file: `slide${String(index + 1).padStart(2, '0')}.png`,
    label
  }))
}));

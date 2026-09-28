"""一次性修订已有演示中的教学假设与边界说明。"""
from pathlib import Path

ROOT = Path(__file__).parent
for name in ('main.py', 'reasoning/rules.py'):
    p = ROOT / name
    s = p.read_text(encoding='utf-8').replace('隐性关联交易', '关联交易待核查')
    s = s.replace('上市公司存在未识别的关联交易 → 预警', '上市公司存在共同控制交易 → 待核查预警')
    s = s.replace('交易双方实际控制人相同 → 关联交易待核查', '交易双方教学规则下的控制人相同 → 关联交易线索')
    s = s.replace('亲属关系视为一致行动人', '教学假设：配偶视为一致行动人（实务须另行取证）')
    s = s.replace('[("Concert", "?p", "?q"), ("Concert", "?q", "?p")], priority=3)', '[("Concert", "?p", "?q"), ("Concert", "?q", "?p")],\n         test=lambda b: b["?rel"] == "配偶", priority=3)')
    s = s.replace('股权穿透（受益所有人 ≥ 25%）', '股权穿透（教学筛查阈值 ≥ 25%，非法律认定）').replace('★受益所有人', '★达到教学筛查阈值')
    p.write_text(s, encoding='utf-8')
p = ROOT / 'reasoning/graph_reasoning.py'
s = p.read_text(encoding='utf-8').replace('threshold=25% 参照《受益所有人信息管理办法》的受益所有人认定标准。', 'threshold=25% 仅为教学筛查阈值，不构成受益所有人的完整法律认定。\n    仅自然人可成为筛查候选；深度截断的企业仍返回但不标记为候选。\n    环路按简单路径截断，不求解交叉持股下的无限级数。')
s = s.replace('beneficial=v["ratio"] >= threshold)', 'beneficial=kg.G.nodes[o]["type"] == "Person" and v["ratio"] >= threshold)')
p.write_text(s, encoding='utf-8')

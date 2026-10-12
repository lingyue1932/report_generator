# -*- coding: utf-8 -*-
"""封面页（模板第 1 页）。

原先这段逻辑内联在 report_generator.fill_pptx() 里，约 40 行，
夹着一堆 has_gt / 姓名 / 日期的启发式判断，与其它页"一页一模块"的
组织方式不一致。抽出来并拆成三个意图明确的小函数。
"""
from datetime import datetime

import parameter as P

PLACEHOLDERS = ("XXX", "xxx")


def _iter_runs(shape):
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            yield run


def _replace_all(shape, old, new):
    """把形状内所有 run 中的 old 替换为 new。"""
    for run in _iter_runs(shape):
        if old in run.text:
            run.text = run.text.replace(old, new)


def _drop_chip_suffix(shape):
    """封面标题里的"芯片"字样去掉（项目名已含具体器件信息）。"""
    for run in _iter_runs(shape):
        if "芯片" in run.text:
            run.text = run.text.replace("芯片", "")


def _fill_identity_block(shape, project_name, author, today):
    """含"姓名""日期"的文本框：首个占位符→姓名，次个→日期，其余→项目名。"""
    name_done = date_done = False
    for run in _iter_runs(shape):
        if not any(ph in run.text for ph in PLACEHOLDERS):
            continue
        if not name_done:
            run.text = run.text.replace("XXX", author).replace("xxx", author)
            name_done = True
        elif not date_done:
            run.text = run.text.replace("XXX", today).replace("xxx", today)
            date_done = True
        else:
            run.text = run.text.replace("XXX", project_name) \
                               .replace("xxx", project_name)


def fill_cover(prs, data, author=None, today=None):
    """填充封面。返回 (project_name, author, today)。

    author 缺省取 parameter.AUTHOR_NAME；today 缺省取今天（%Y-%m-%d）。
    两者都可由 CLI / GUI 显式传入。
    """
    project_name = str(data.get("物料信息", {}).get("project_name") or "").strip()
    author = author or P.AUTHOR_NAME
    today = today or datetime.now().strftime("%Y-%m-%d")

    slide = prs.slides[P.TEMPLATE_PAGE_INDEX["cover"]]
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        full_text = shape.text_frame.text
        if not full_text.strip():
            continue
        # 标题行：XXX 芯片 -> 项目名
        if "GT" in full_text:
            _replace_all(shape, "XXX", project_name)
            _drop_chip_suffix(shape)
        # 姓名 / 日期行
        if "姓名" in full_text and "日期" in full_text:
            _fill_identity_block(shape, project_name, author, today)
        else:
            _replace_all(shape, "XXX", project_name)
            _replace_all(shape, "xxx", project_name)
    return project_name, author, today

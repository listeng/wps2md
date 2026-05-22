import pathlib

import pytest

from wps2md import Paragraph, WpsParseError, parse, to_markdown

TESTS_DIR = pathlib.Path(__file__).parent
SAMPLES = {
    "wps": TESTS_DIR / "sample.wps",
    "doc": TESTS_DIR / "sample.doc",
}


@pytest.fixture(scope="module", params=sorted(SAMPLES))
def sample(request):
    path = SAMPLES[request.param]
    if not path.exists():
        pytest.skip(f"{path.name} not present")
    return path


@pytest.fixture(scope="module")
def doc(sample):
    return parse(sample)


@pytest.fixture(scope="module")
def md(doc):
    return to_markdown(doc.paragraphs)


def test_main_text_nonempty(doc):
    assert len(doc.main_text) > 0


def test_first_paragraph_is_h1(doc):
    assert doc.paragraphs
    assert doc.paragraphs[0].heading_level == 1
    assert doc.paragraphs[0].text == "春季露营活动指南"


def test_markdown_starts_with_h1(md):
    assert md.startswith("# 春季露营活动指南")


def test_multi_level_headings(md):
    # H1, H2, H3 should all appear in the sample.
    assert "\n# " in "\n" + md
    assert "\n## 一、活动筹备清单" in md
    assert "\n### 1. 必备露营装备" in md


def test_unordered_list_rendered_with_dashes(md):
    # "必备露营装备" 下面是无序列表 (nfc=23, bullet).
    assert "- 基础住宿类：自动充气帐篷、防潮垫、羽绒睡袋" in md
    assert "- 照明工具：头灯、营地灯、备用干电池" in md
    assert "- 应急物品：急救包、防蚊液、防晒霜、多功能刀具" in md
    # "安全规范" 下面也是无序列表.
    assert "- 严格在指定区域生火，离开时必须确认火源完全熄灭" in md


def test_ordered_list_rendered_with_numbers(md):
    # "食材准备计划" 下面是有序列表 (nfc=0, Arabic).
    assert "1. 提前处理分装生食：腌制好的牛羊肉串、切好的蔬菜盒" in md
    assert "2. 分装即食食品：面包、真空包装卤味、水果切盒" in md
    assert "3. 饮料分类打包：常温矿泉水、冷藏气泡酒、热饮冲剂" in md
    # "环保要求" 下面也是有序列表, 计数器应重置回 1.
    assert "1. 所有垃圾必须全部打包带走，不留下任何外来污染物" in md
    assert "2. 不随意采摘野生植物，不破坏植被和动物栖息地" in md


def test_table_rendered_as_pipe_table(md):
    assert "| 日期 | 上午行程 | 下午行程 | 晚上安排 |" in md
    assert "| --- | --- | --- | --- |" in md
    assert "| 第一天 | 9点集合出发，11点到达营地搭建帐篷" in md
    assert "| 第三天 |" in md


def test_body_paragraph_present(md):
    assert "营地位于城郊云栖谷" in md
    assert "负氧离子浓度是市区的15倍" in md


def test_paragraph_flags_capture_table_and_list(doc):
    table_paras = [p for p in doc.paragraphs if p.in_table]
    list_paras = [p for p in doc.paragraphs if p.ilfo and not p.in_table]
    assert table_paras, "expected table paragraphs"
    assert any(p.is_row_end for p in table_paras)
    assert list_paras, "expected list paragraphs"
    # Sample has both ordered and unordered list paragraphs.
    assert any(p.list_ordered is True for p in list_paras)
    assert any(p.list_ordered is False for p in list_paras)


def test_markdown_ends_with_single_newline(md):
    assert md.endswith("\n")
    assert not md.endswith("\n\n")


def test_parse_from_bytes_matches_path(sample):
    from_path = to_markdown(parse(sample).paragraphs)
    from_bytes = to_markdown(parse(sample.read_bytes()).paragraphs)
    assert from_path == from_bytes


def test_wps_and_doc_render_identically():
    if not (SAMPLES["wps"].exists() and SAMPLES["doc"].exists()):
        pytest.skip("both samples required")
    md_wps = to_markdown(parse(SAMPLES["wps"]).paragraphs)
    md_doc = to_markdown(parse(SAMPLES["doc"]).paragraphs)
    assert md_wps == md_doc


def test_rejects_unsupported_extension(tmp_path):
    p = tmp_path / "foo.txt"
    p.write_bytes(b"\xd0\xcf\x11\xe0")
    with pytest.raises(WpsParseError):
        parse(p)


def test_rejects_non_ole_bytes():
    with pytest.raises(Exception):
        parse(b"not an OLE2 document")


def test_heading_level_property():
    assert Paragraph(istd=1, text="x").heading_level == 1
    assert Paragraph(istd=9, text="x").heading_level == 9
    assert Paragraph(istd=0, text="x").heading_level == 0
    assert Paragraph(istd=10, text="x").heading_level == 0

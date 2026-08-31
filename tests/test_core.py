import tempfile
import unittest
from pathlib import Path

import openpyxl
import pandas as pd

import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import word_freq_analyzer as analyzer
from llm_sentence_analyzer import LLMAnalyzerConfig, QwenSentenceAnalyzer


class CoreRegressionTests(unittest.TestCase):
    def test_parse_mixed_year_formats_without_dt_failure(self):
        values = pd.Series([
            "2024",
            "2024-03-15",
            "2011/1/1-2011/12/31",
            "2009年",
            45292,  # Excel serial date: 2024-01-01
            "无法识别",
        ])
        self.assertEqual(
            analyzer.parse_year_column(values).tolist(),
            [2024, 2024, 2011, 2009, 2024, 0],
        )

    def test_excel_headers_are_consistent_with_preview(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "blank_header.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["companyid", None, "调研报告内容"])
            ws.append([1, "2024", "公司积极保护生态环境并修复湿地。"])
            wb.save(path)
            wb.close()

            columns = analyzer._read_columns_fast(str(path))
            frame = analyzer.read_data_file(str(path))
            self.assertEqual(columns, list(frame.columns))
            self.assertEqual(columns[1], "Unnamed: 1")

    def test_dictionary_excel_header_is_not_imported_as_keyword(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "dict.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["分类", "关键词"])
            ws.append(["环境", "生态环境"])
            wb.save(path)
            wb.close()

            manager = analyzer.DictionaryManager()
            manager.import_file(str(path))
            self.assertEqual(manager.data, {"环境": ["生态环境"]})

    def test_realistic_xlsx_pipeline_writes_main_and_sentence_outputs(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "调研内容.xlsx"
            output = tmp_path / "result.xlsx"
            pd.DataFrame([
                {"companyid": "000001", "公司简称": "甲公司", "日期": "2024年",
                 "调研报告内容": "公司积极保护生态环境并修复湿地。"},
                {"companyid": "2", "公司简称": "乙公司", "日期": "2024-03-15",
                 "调研报告内容": "公司持续推进绿色发展和环境治理。"},
            ]).to_excel(source, index=False)

            manager = analyzer.DictionaryManager()
            manager.add_category("环境")
            manager.add_word("环境", "生态环境")
            manager.add_word("环境", "绿色发展")

            analyzer.run_analysis(
                files=[str(source)],
                dict_mgr=manager,
                col_stkcd="companyid",
                col_year="日期",
                text_columns=["调研报告内容"],
                output_path=str(output),
                use_regex=True,
                export_sentences=True,
            )

            self.assertTrue(output.exists())
            sentence_output = tmp_path / "result_sentences.xlsx"
            self.assertTrue(sentence_output.exists())
            self.assertEqual(
                pd.ExcelFile(output).sheet_names,
                ["公司年份分类统计", "关键词明细", "分类汇总", "词典诊断", "分析说明"],
            )
            panel = pd.read_excel(output, sheet_name="公司年份分类统计", dtype=str)
            self.assertEqual(set(panel["公司代码"]), {"000001", "000002"})
            self.assertTrue((pd.to_numeric(panel["总计"], errors="coerce") > 0).all())
            sentences = pd.read_excel(sentence_output, sheet_name="命中句子")
            self.assertEqual(len(sentences), 2)
            self.assertFalse((tmp_path / "result_checkpoint").exists())

    def test_llm_partial_json_is_not_reported_as_success(self):
        model = QwenSentenceAnalyzer(LLMAnalyzerConfig.from_inputs(api_key="test-key"))
        partial = model._normalize_result({"rel": 1, "time": 2})
        self.assertEqual(partial["LLM分析状态"], "失败")
        complete = model._normalize_result({
            "rel": 1, "time": 2, "voice": 0, "type": 2,
            "cert": 2, "quant": 1, "tone": 1, "conf": 2,
        })
        self.assertEqual(complete["LLM分析状态"], "成功")


if __name__ == "__main__":
    unittest.main()

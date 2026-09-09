"""
========================================================================================
🧪 ARGOS.UZ - AVTOMATLASHGAN TESTLAR TO'PLAMI (AUTOMATED QA TEST SUITE)
========================================================================================
Ishga tushirish:
    python test_suite.py
========================================================================================
"""
import unittest
import io
import os
import json
import zipfile
import threading
import app as main_app

class ArgosUnitTests(unittest.TestCase):
    """1. Asosiy yordamchi funksiyalar bo'yicha Unit Testlar"""

    def test_extract_location(self):
        vil, tum = main_app.extract_location("Toshkent shahri , Chilonzor tumani")
        self.assertEqual(vil, "Toshkent shahri")
        self.assertEqual(tum, "Chilonzor tumani")

        vil_none, tum_none = main_app.extract_location(None)
        self.assertEqual(vil_none, "Ko'rsatilmagan")
        self.assertEqual(tum_none, "Ko'rsatilmagan")

    def test_format_salary_display(self):
        self.assertEqual(main_app.format_salary_display(None), "Shtat jadvali bo'yicha")
        self.assertEqual(main_app.format_salary_display("0"), "Shtat jadvali bo'yicha")
        self.assertEqual(main_app.format_salary_display(5000000), "5 000 000 UZS")

    def test_format_date_uzbek(self):
        formatted = main_app.format_date_uzbek("2026-09-09T10:00:00")
        self.assertIn("sentabr 2026", formatted)

    def test_normalize_region_id(self):
        # SOATO kodi orqali
        self.assertEqual(main_app.normalize_region_id("1726"), "10")
        # To'g'ridan-to'g'ri qiymat
        self.assertEqual(main_app.normalize_region_id("10"), "10")
        # Nomi orqali
        self.assertEqual(main_app.normalize_region_id("toshkent shahri"), "10")
        # Barchasi / Bo'sh
        self.assertIsNone(main_app.normalize_region_id("0"))
        self.assertIsNone(main_app.normalize_region_id("Barchasi"))
        self.assertIsNone(main_app.normalize_region_id(None))

    def test_test_type_resolution(self):
        client = main_app.api_client
        target, label = client.resolve_test_type_target("boshqaruv")
        self.assertEqual(target, 4)

        target_m, label_m = client.resolve_test_type_target("mutaxassis")
        self.assertEqual(target_m, 5)

        target_none, _ = client.resolve_test_type_target("barchasi")
        self.assertIsNone(target_none)


class ArgosIntegrationTests(unittest.TestCase):
    """2. API Endpointlar bo'yicha Integratsion Testlar"""

    def setUp(self):
        self.client = main_app.app.test_client()

    def test_index_page(self):
        res = self.client.get("/")
        self.assertEqual(res.status_code, 200)
        html = res.data.decode("utf-8")
        self.assertIn("ARGOS", html)
        self.assertIn("filter-sidebar", html)
        self.assertIn("btn-search-top", html)
        self.assertIn("btn-theme-toggle", html)

    def test_favicon(self):
        res = self.client.get("/favicon.ico")
        self.assertEqual(res.status_code, 200)
        self.assertIn("svg", res.mimetype)

    def test_api_regions(self):
        res = self.client.get("/api/regions")
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertIsInstance(data, list)
        self.assertGreater(len(data), 10)

    def test_api_districts_normalized(self):
        # Toshkent shahri kodi bilan
        res1 = self.client.get("/api/districts?region=10")
        self.assertEqual(res1.status_code, 200)
        districts1 = res1.get_json()
        self.assertGreater(len(districts1), 0)

        # SOATO kodi bilan (1726)
        res2 = self.client.get("/api/districts?region=1726")
        self.assertEqual(res2.status_code, 200)
        districts2 = res2.get_json()
        self.assertEqual(len(districts1), len(districts2))

    def test_api_test_types(self):
        res = self.client.get("/api/test-types")
        self.assertEqual(res.status_code, 200)
        types = res.get_json()
        self.assertGreater(len(types), 3)

    def test_api_history(self):
        res = self.client.get("/api/history")
        self.assertEqual(res.status_code, 200)
        self.assertIsInstance(res.get_json(), list)

    def test_api_stop(self):
        res = self.client.post("/api/stop")
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.get_json().get("status"), "ok")

    def test_delete_history_forbidden(self):
        res = self.client.delete("/api/history/sample.docx")
        self.assertEqual(res.status_code, 403)


class ArgosSecurityAndEdgeCaseTests(unittest.TestCase):
    """3. Xavfsizlik va Kutilmagan Chekka Holatlar Testlari"""

    def setUp(self):
        self.client = main_app.app.test_client()

    def test_path_traversal_prevention(self):
        res1 = self.client.get("/api/preview/..%2F..%2Fapp.py")
        self.assertEqual(res1.status_code, 404)

        res2 = self.client.get("/api/docx-raw/..%2F..%2Fapp.py")
        self.assertEqual(res2.status_code, 404)

        res3 = self.client.get("/api/download/..%2F..%2Fapp.py")
        self.assertEqual(res3.status_code, 404)

    def test_upload_non_docx_rejected(self):
        res = self.client.post('/api/upload-docx', data={'file': (io.BytesIO(b'echo test'), 'script.sh')})
        self.assertEqual(res.status_code, 400)
        self.assertIn("Faqat .docx", res.get_json().get("error", ""))

    def test_upload_fake_docx_magic_bytes_rejected(self):
        res = self.client.post('/api/upload-docx', data={'file': (io.BytesIO(b'plain text'), 'test.docx')})
        self.assertEqual(res.status_code, 400)
        self.assertIn("haqiqiy Word", res.get_json().get("error", ""))

    def test_upload_valid_docx_accepted(self):
        bio = io.BytesIO()
        with zipfile.ZipFile(bio, 'w') as zf:
            zf.writestr('[Content_Types].xml', '<Types></Types>')
        bio.seek(0)
        res = self.client.post('/api/upload-docx', data={'file': (bio, 'test_qa_doc.docx')})
        self.assertEqual(res.status_code, 200)
        fname = res.get_json().get("filename")
        self.assertTrue(fname.endswith(".docx"))

        # Tozalash
        target = os.path.join(main_app.EXPORTS_DIR, fname)
        if os.path.exists(target): os.remove(target)
        if os.path.exists(target + ".json"): os.remove(target + ".json")

    def test_non_existent_vacancy_export(self):
        res = self.client.post('/api/export-single/999999999')
        self.assertEqual(res.status_code, 404)
        self.assertIn("topilmadi", res.get_json().get("error", ""))


class ArgosEndToEndFlowTests(unittest.TestCase):
    """4. End-to-End va Foydalanuvchi Jarayonlari Testlari"""

    def setUp(self):
        self.client = main_app.app.test_client()

    def test_vacancies_list_and_single_export_flow(self):
        # 1. Vakansiyalarni ro'yxatdan olish
        r_list = self.client.get('/api/vacancies?limit=1')
        self.assertEqual(r_list.status_code, 200)
        vacs = r_list.get_json().get('vacancies', [])
        if not vacs:
            self.skipTest("Argos portalidan vakansiyalar olinmadi (internet yoki portal vaqtinchalik javob bermayapti)")

        vac_id = vacs[0]['id']
        # 2. Yakka vakansiyani Word qilib eksport qilish
        r_exp = self.client.post(f'/api/export-single/{vac_id}')
        self.assertEqual(r_exp.status_code, 200)
        fname = r_exp.get_json().get('filename')
        self.assertTrue(fname and fname.endswith('.docx'))

        # 3. Preview (Mammoth) ko'rish
        r_prev = self.client.get(f'/api/preview/{fname}')
        self.assertEqual(r_prev.status_code, 200)
        self.assertIn('html', r_prev.get_json())

        # 4. Raw docx yuklab olish
        r_raw = self.client.get(f'/api/docx-raw/{fname}')
        self.assertEqual(r_raw.status_code, 200)
        self.assertGreater(len(r_raw.data), 1000)

        # 5. Word tarixi bo'yicha tekshirish
        r_hist = self.client.get('/api/history')
        self.assertEqual(r_hist.status_code, 200)
        history_files = [h['filename'] for h in r_hist.get_json()]
        self.assertIn(fname, history_files)


class ArgosStressConcurrencyTests(unittest.TestCase):
    """5. Bir vaqtda yuklanish va ko'p oqimli barqarorlik Testlari"""

    def setUp(self):
        self.client = main_app.app.test_client()

    def test_concurrent_api_requests(self):
        results = []
        errors = []

        def worker():
            try:
                r1 = self.client.get('/api/regions')
                r2 = self.client.get('/api/test-types')
                r3 = self.client.get('/api/history')
                results.append((r1.status_code, r2.status_code, r3.status_code))
            except Exception as e:
                errors.append(e)

        threads = [threading.Thread(target=worker) for _ in range(10)]
        for t in threads: t.start()
        for t in threads: t.join()

        self.assertEqual(len(errors), 0)
        self.assertEqual(len(results), 10)
        for r in results:
            self.assertEqual(r, (200, 200, 200))


if __name__ == '__main__':
    print("=" * 65)
    print("🚀 ARGOS.UZ TO'LIQ AVTOMATLASHGAN QA TESTLARINI ISHGA TUSHIRISH")
    print("=" * 65)
    unittest.main(verbosity=2)

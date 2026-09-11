import unittest, os, json, app as main_app

class ArgosSecurityHardeningAudit(unittest.TestCase):
    def setUp(self):
        main_app.rate_limiter.reset_for_test()
        self.client = main_app.app.test_client()

    def test_http_security_headers_present(self):
        res = self.client.get('/')
        self.assertEqual(res.status_code, 200)
        h = res.headers
        self.assertEqual(h.get('X-Content-Type-Options'), 'nosniff')
        self.assertEqual(h.get('X-Frame-Options'), 'SAMEORIGIN')
        self.assertIn('max-age=31536000', h.get('Strict-Transport-Security', ''))
        self.assertEqual(h.get('Referrer-Policy'), 'strict-origin-when-cross-origin')
        self.assertIn('camera=()', h.get('Permissions-Policy', ''))
        self.assertEqual(h.get('Cross-Origin-Opener-Policy'), 'same-origin-allow-popups')
        self.assertEqual(h.get('Cross-Origin-Resource-Policy'), 'cross-origin')
        self.assertIn('default-src', h.get('Content-Security-Policy', ''))
        self.assertIn('no-cache', h.get('Cache-Control', ''))
        self.assertIn('ARGOS-SECURE-GATEWAY', h.get('Server', ''))

    def test_robots_txt_crawler_policy(self):
        res = self.client.get('/robots.txt')
        self.assertEqual(res.status_code, 200)
        content = res.data.decode('utf-8')
        self.assertIn('Disallow: /api/', content)
        self.assertIn('Disallow: /exports/', content)
        self.assertIn('User-agent: GPTBot', content)
        self.assertIn('User-agent: ClaudeBot', content)

    def test_ai_bot_blocking_on_api(self):
        blocked_uas = ['GPTBot', 'ClaudeBot', 'CCBot', 'Bytespider', 'Scrapy', 'PetalBot']
        for ua in blocked_uas:
            res = self.client.get('/api/regions', headers={'User-Agent': ua})
            self.assertEqual(res.status_code, 403)
            self.assertEqual(res.get_json().get('status'), 'error')

        legit_ua = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/130.0.0.0 Safari/537.36'
        res_ok = self.client.get('/api/regions', headers={'User-Agent': legit_ua})
        self.assertEqual(res_ok.status_code, 200)

    def test_rate_limiting_tier1_heavy_endpoints(self):
        main_app.rate_limiter.reset_for_test()
        for _ in range(40):
            res = self.client.post('/api/upload-docx', data={})
            self.assertNotEqual(res.status_code, 429)

        res_limit = self.client.post('/api/upload-docx', data={})
        self.assertEqual(res_limit.status_code, 429)
        self.assertIn('Retry-After', res_limit.headers)
        self.assertEqual(res_limit.get_json().get('status'), 'error')

    def test_path_traversal_hardened_checks(self):
        malicious_filenames = ['../../app.py', '..%2F..%2Fapp.py', '....//....//app.py', 'test.docx%00.exe', '.', '..']
        for bad_name in malicious_filenames:
            r1 = self.client.get('/api/preview/' + bad_name)
            self.assertIn(r1.status_code, (400, 404))
            r2 = self.client.get('/api/docx-raw/' + bad_name)
            self.assertIn(r2.status_code, (400, 404))
            r3 = self.client.get('/api/download/' + bad_name)
            self.assertIn(r3.status_code, (400, 404))

    def test_cors_origin_whitelisting(self):
        res = self.client.get('/api/regions', headers={'Origin': 'https://argostest.web.app'})
        self.assertEqual(res.headers.get('Access-Control-Allow-Origin'), 'https://argostest.web.app')

        res_evil = self.client.get('/api/regions', headers={'Origin': 'https://evil-site.com'})
        self.assertIsNone(res_evil.headers.get('Access-Control-Allow-Origin'))

        res_opt = self.client.open('/api/vacancies', method='OPTIONS', headers={'Origin': 'https://argostest.onrender.com'})
        self.assertEqual(res_opt.status_code, 204)
        self.assertEqual(res_opt.headers.get('Access-Control-Allow-Origin'), 'https://argostest.onrender.com')

    def test_generic_error_handlers_no_leak(self):
        r404 = self.client.get('/api/non-existent-endpoint-xyz')
        self.assertEqual(r404.status_code, 404)
        self.assertEqual(r404.get_json().get('status'), 'error')

        r405 = self.client.delete('/api/regions')
        self.assertEqual(r405.status_code, 405)
        self.assertEqual(r405.get_json().get('status'), 'error')

        r403 = self.client.delete('/api/history/sample.docx')
        self.assertEqual(r403.status_code, 403)
        d403 = r403.get_json()
        self.assertEqual(d403.get('status'), 'error')

    def test_frontend_f12_ctrl_u_right_click_protection(self):
        """Frontend da F12, Ctrl+U va Right-Click bloklanganligini tekshirish"""
        res = self.client.get('/')
        self.assertEqual(res.status_code, 200)
        html = res.data.decode('utf-8')
        self.assertIn('contextmenu', html)
        self.assertIn('F12', html)
        self.assertIn('Ctrl+U', html)
        self.assertIn('argos-security-toast', html)

if __name__ == '__main__':
    unittest.main(verbosity=2)
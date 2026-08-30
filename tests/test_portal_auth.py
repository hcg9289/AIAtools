import importlib
import os
import sys
import types
import unittest
from unittest.mock import Mock, patch


def _install_import_stubs():
    fitz = types.ModuleType('fitz')
    fitz.Matrix = lambda *args: args
    sys.modules.setdefault('fitz', fitz)

    pptx = types.ModuleType('pptx')
    pptx.Presentation = object
    sys.modules.setdefault('pptx', pptx)

    util = types.ModuleType('pptx.util')
    util.Inches = lambda value: value
    util.Pt = lambda value: value
    sys.modules.setdefault('pptx.util', util)

    color = types.ModuleType('pptx.dml.color')
    color.RGBColor = lambda *values: values
    sys.modules.setdefault('pptx.dml', types.ModuleType('pptx.dml'))
    sys.modules.setdefault('pptx.dml.color', color)

    shapes = types.ModuleType('pptx.enum.shapes')
    shapes.MSO_CONNECTOR = types.SimpleNamespace(STRAIGHT=1)
    shapes.MSO_SHAPE = types.SimpleNamespace(RECTANGLE=1)
    text = types.ModuleType('pptx.enum.text')
    text.MSO_ANCHOR = types.SimpleNamespace(TOP=1, MIDDLE=2, BOTTOM=3)
    text.MSO_AUTO_SIZE = types.SimpleNamespace(TEXT_TO_FIT_SHAPE=1)
    text.PP_ALIGN = types.SimpleNamespace(LEFT=1, CENTER=2, RIGHT=3)
    sys.modules.setdefault('pptx.enum', types.ModuleType('pptx.enum'))
    sys.modules.setdefault('pptx.enum.shapes', shapes)
    sys.modules.setdefault('pptx.enum.text', text)

    ns = types.ModuleType('pptx.oxml.ns')
    ns.qn = lambda value: value
    xmlchemy = types.ModuleType('pptx.oxml.xmlchemy')
    xmlchemy.OxmlElement = object
    sys.modules.setdefault('pptx.oxml', types.ModuleType('pptx.oxml'))
    sys.modules.setdefault('pptx.oxml.ns', ns)
    sys.modules.setdefault('pptx.oxml.xmlchemy', xmlchemy)

    medical_pdf = types.ModuleType('gf_medical_pdf')
    medical_pdf.PdfPayloadError = ValueError
    medical_pdf.build_medical_financing_pdf = Mock()
    sys.modules.setdefault('gf_medical_pdf', medical_pdf)


os.environ['PORTAL_SERVICE_KEY_FILE'] = '/definitely/not/a/real/secret'
os.environ['PORTAL_SERVICE_KEY'] = 'test-service-key'
_install_import_stubs()
portal = importlib.import_module('app')


def _response(status_code, payload=None):
    response = Mock()
    response.status_code = status_code
    response.json.return_value = payload or {}
    return response


class PortalAuthTests(unittest.TestCase):
    def setUp(self):
        portal.app.config.update(TESTING=True)
        self.client = portal.app.test_client()

    @patch.object(portal.requests, 'post')
    def test_ott_exchange_creates_fixed_secure_host_cookie(self, post):
        post.return_value = _response(200, {
            'valid': True,
            'session_token': 'central-session-token',
            'uid': 'user-1',
            'expires_at': '2026-08-30T10:20:00Z',
            'expires_in': 1200,
        })

        response = self.client.get('/?ott=one-time-token')

        self.assertEqual(response.status_code, 303)
        self.assertEqual(response.headers['Location'], '/')
        cookie_headers = response.headers.getlist('Set-Cookie')
        portal_cookie = next(value for value in cookie_headers if value.startswith('portal_session='))
        self.assertIn('Max-Age=1200', portal_cookie)
        self.assertIn('Secure', portal_cookie)
        self.assertIn('HttpOnly', portal_cookie)
        self.assertIn('SameSite=Lax', portal_cookie)
        self.assertIn('Path=/', portal_cookie)
        self.assertNotIn('Domain=', portal_cookie)
        post.assert_called_once_with(
            portal.VAULT_PORTAL_EXCHANGE_URL,
            json={'ott': 'one-time-token', 'audience': '1008'},
            headers={'X-Portal-Service-Key': 'test-service-key'},
            timeout=portal.VAULT_AUTH_TIMEOUT_SECONDS,
        )

    @patch.object(portal.requests, 'post')
    def test_index_validates_central_session_and_renders_all_launch_cards(self, post):
        post.return_value = _response(200, {
            'valid': True,
            'uid': 'user-1',
            'expires_at': '2026-08-30T10:20:00Z',
            'expires_in': 900,
        })
        self.client.set_cookie('portal_session', 'central-session-token')

        response = self.client.get('/')

        self.assertEqual(response.status_code, 200)
        page = response.get_data(as_text=True)
        self.assertIn('action="/launch/1003"', page)
        self.assertIn('action="/launch/1004"', page)
        self.assertIn('action="/launch/1005"', page)
        self.assertIn('GF 醫療融資', page)
        self.assertIn('Cold call申請流程簡化', page)
        self.assertNotIn('電話名單生成', page)
        self.assertEqual(page.count('class="tool-card"'), 6)
        self.assertEqual(page.count('class="tool-name"'), 6)
        self.assertIn('grid-template-columns: repeat(2, minmax(0, 1fr))', page)
        self.assertNotIn('整理稅務收據及相關文件', page)
        self.assertNotIn('按醫療保費及提款安排', page)
        self.assertNotIn('class="tool-meta"', page)
        self.assertNotIn('class="pill"', page)

    @patch.object(portal.requests, 'post')
    def test_cold_call_page_uses_simplified_application_name(self, post):
        post.return_value = _response(200, {
            'valid': True,
            'uid': 'user-1',
            'expires_at': '2026-08-30T10:20:00Z',
            'expires_in': 900,
        })
        self.client.set_cookie('portal_session', 'central-session-token')

        response = self.client.get('/tools/cold-call-list')

        self.assertEqual(response.status_code, 200)
        page = response.get_data(as_text=True)
        self.assertIn('<title>AIAtools - Cold call申請流程簡化</title>', page)
        self.assertIn('<h1>Cold call申請流程簡化</h1>', page)
        self.assertNotIn('電話名單生成', page)

    @patch.object(portal.requests, 'post')
    def test_ping_does_not_extend_or_replace_cookie(self, post):
        post.return_value = _response(200, {
            'valid': True,
            'uid': 'user-1',
            'expires_at': '2026-08-30T10:20:00Z',
            'expires_in': 321,
        })
        self.client.set_cookie('portal_session', 'central-session-token')

        response = self.client.get('/api/ping')

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()['expires_in'], 321)
        self.assertNotIn('Set-Cookie', response.headers)
        self.assertEqual(post.call_count, 1)
        self.assertEqual(
            post.call_args.kwargs['json'],
            {'session_token': 'central-session-token'},
        )

    @patch.object(portal.requests, 'post')
    def test_launch_uses_csrf_allowlist_and_fixed_redirect(self, post):
        post.side_effect = [
            _response(200, {
                'valid': True,
                'uid': 'user-1',
                'expires_at': '2026-08-30T10:20:00Z',
                'expires_in': 900,
            }),
            _response(200, {
                'valid': True,
                'ticket': 'A_valid_launch_ticket_123456789',
                'audience': '1004',
                'expires_at': '2026-08-30T10:00:30Z',
            }),
        ]
        session_token = 'central-session-token'
        self.client.set_cookie('portal_session', session_token)
        csrf = portal.portal_csrf_token(session_token)

        response = self.client.post('/launch/1004', data={'csrf_token': csrf})

        self.assertEqual(response.status_code, 303)
        self.assertEqual(
            response.headers['Location'],
            'https://ia.alpha-family.net/?launch=A_valid_launch_ticket_123456789',
        )
        self.assertEqual(post.call_count, 2)
        self.assertEqual(post.call_args.kwargs['json'], {
            'session_token': session_token,
            'audience': '1004',
        })

    @patch.object(portal.requests, 'post')
    def test_unknown_tool_and_bad_csrf_never_request_ticket(self, post):
        valid_session = _response(200, {
            'valid': True,
            'uid': 'user-1',
            'expires_at': '2026-08-30T10:20:00Z',
            'expires_in': 900,
        })
        post.return_value = valid_session
        self.client.set_cookie('portal_session', 'central-session-token')

        unknown = self.client.post('/launch/9999', data={'csrf_token': 'bad'})
        bad_csrf = self.client.post('/launch/1003', data={'csrf_token': 'bad'})

        self.assertEqual(unknown.status_code, 403)
        self.assertEqual(bad_csrf.status_code, 403)
        self.assertEqual(post.call_count, 2)  # central validation only

    @patch.object(portal.requests, 'post')
    def test_direct_access_without_session_is_forbidden(self, post):
        response = self.client.get('/')

        self.assertEqual(response.status_code, 403)
        post.assert_not_called()

    @patch.object(portal.requests, 'post')
    def test_whatsapp_preview_uses_aiatools_branding(self, post):
        response = self.client.get('/', headers={'User-Agent': 'WhatsApp/2.0'})

        self.assertEqual(response.status_code, 200)
        page = response.get_data(as_text=True)
        self.assertIn('<title>AIAtools</title>', page)
        self.assertIn('property="og:title" content="AIAtools"', page)
        self.assertNotIn('PPT Generator', page)
        post.assert_not_called()

    def test_missing_service_key_fails_closed(self):
        with patch.object(portal, 'PORTAL_SERVICE_KEY', ''):
            status, data = portal.exchange_portal_ott('token')

        self.assertEqual(status, 'misconfigured')
        self.assertIsNone(data)


if __name__ == '__main__':
    unittest.main()

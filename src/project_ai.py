# -*- coding: utf-8 -*-
"""可配置的 AI/检索接口；密钥不进入文书、普通配置或日志。"""
from datetime import datetime
from pathlib import Path
from urllib import request, error, parse
import base64
import ctypes
import json
import os
import re
import socket
import tempfile

DEFAULTS = {'base_url': '', 'model': '', 'api_key': '', 'search_key': '',
            'input_chars': 18000, 'output_tokens': 6000, 'timeout': 120,
            'remember_keys': False}


class ServiceError(ValueError):
    pass


class Cancelled(Exception):
    pass


def check_cancel(event):
    if event is not None and event.is_set():
        raise Cancelled('已取消后续处理；已发送的请求仍可能计费。')


def config_path():
    return Path(os.environ.get('LOCALAPPDATA') or Path.home() / '.local' / 'share') / '填表工具' / 'project_writer.json'


def _crypt(data, decrypt=False):
    if os.name != 'nt':
        raise ServiceError('本机密钥加密仅支持Windows，请使用仅本次会话模式。')
    from ctypes import wintypes

    class Blob(ctypes.Structure):
        _fields_ = [('length', wintypes.DWORD), ('data', ctypes.POINTER(ctypes.c_ubyte))]

    buf = ctypes.create_string_buffer(data)
    source = Blob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_ubyte)))
    target = Blob()
    crypt = ctypes.WinDLL('crypt32', use_last_error=True)
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.LocalFree.argtypes = [ctypes.c_void_p]
    kernel.LocalFree.restype = ctypes.c_void_p
    if decrypt:
        fn = crypt.CryptUnprotectData
        fn.argtypes = [ctypes.POINTER(Blob), ctypes.c_void_p, ctypes.c_void_p,
                       ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(Blob)]
        args = (ctypes.byref(source), None, None, None, None, 1, ctypes.byref(target))
    else:
        fn = crypt.CryptProtectData
        fn.argtypes = [ctypes.POINTER(Blob), wintypes.LPCWSTR, ctypes.c_void_p,
                       ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(Blob)]
        args = (ctypes.byref(source), '填表工具项目文书', None, None, None, 1, ctypes.byref(target))
    fn.restype = wintypes.BOOL
    if not fn(*args):
        raise ServiceError('Windows密钥保护失败，请重新输入密钥，或关闭记住密钥。')
    try:
        return ctypes.string_at(target.data, target.length)
    finally:
        kernel.LocalFree(ctypes.cast(target.data, ctypes.c_void_p))


def load_config(path=None):
    result = dict(DEFAULTS)
    warnings = []
    path = Path(path or config_path())
    if not path.exists():
        return result, warnings
    try:
        stored = json.loads(path.read_text(encoding='utf-8'))
        for key in DEFAULTS:
            if key not in ('api_key', 'search_key') and key in stored:
                result[key] = stored[key]
        if result['remember_keys'] and stored.get('protected_keys'):
            secrets = json.loads(_crypt(base64.b64decode(stored['protected_keys']), True))
            for name in ('api_key', 'search_key'):
                result[name] = str(secrets.get(name, ''))
    except Exception:
        warnings.append('未能读取本机设置或解密密钥，请重新配置；没有向外部服务发送数据。')
    return result, warnings


def save_config(settings, path=None):
    path = Path(path or config_path())
    path.parent.mkdir(parents=True, exist_ok=True)
    data = {key: settings.get(key, DEFAULTS[key]) for key in DEFAULTS
            if key not in ('api_key', 'search_key')}
    if data['remember_keys']:
        secrets = json.dumps({key: settings.get(key, '') for key in ('api_key', 'search_key')}).encode()
        data['protected_keys'] = base64.b64encode(_crypt(secrets)).decode('ascii')
    fd, temporary = tempfile.mkstemp(dir=path.parent, suffix='.tmp')
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as stream:
            json.dump(data, stream, ensure_ascii=False, indent=2)
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def endpoint(base):
    parts = parse.urlsplit(base.strip())
    if (parts.scheme != 'https' or not parts.hostname or parts.username or parts.password
            or parts.query or parts.fragment):
        raise ServiceError('服务地址须为不含账号、查询串的HTTPS地址，例如 https://服务域名/v1。')
    path = parts.path.rstrip('/')
    if not path:
        path = '/v1'
    if not path.endswith('/chat/completions'):
        path += '/chat/completions'
    return parse.urlunsplit((parts.scheme, parts.netloc, path, '', ''))


class _NoRedirect(request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def post_json(url, payload, key, timeout):
    data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
    if '\r' in key or '\n' in key:
        raise ServiceError('密钥含无效换行。')
    req = request.Request(url, data=data, method='POST', headers={
        'Content-Type': 'application/json', 'Authorization': 'Bearer ' + key,
        'User-Agent': 'ProjectDocumentWriter/1.0'})
    try:
        with request.build_opener(_NoRedirect()).open(req, timeout=timeout) as response:
            body = response.read(8 * 1024 * 1024 + 1)
            if len(body) > 8 * 1024 * 1024:
                raise ServiceError('服务返回内容过大，请降低单次输出长度。')
            result = json.loads(body.decode('utf-8'))
            if not isinstance(result, dict):
                raise ServiceError('服务返回结构无效。')
            return result
    except error.HTTPError as exc:
        messages = {400: '请求被拒绝，请检查模型名称、上下文上限及接口兼容性。',
                    401: '认证失败，请检查API Key。', 402: '账户余额不足或服务欠费。',
                    403: '当前账号无权调用此服务或模型。', 404: '接口或模型不存在，请检查服务地址。',
                    413: '输入超出服务限制，请降低上下文上限。',
                    429: '请求限流或配额不足，请稍后重试并检查账户额度。'}
        raise ServiceError(messages.get(exc.code, '服务请求失败（HTTP %d），没有自动重试。' % exc.code)) from None
    except (TimeoutError, socket.timeout):
        raise ServiceError('请求超时；远端可能已经计费，请检查后手动继续。') from None
    except error.URLError:
        raise ServiceError('无法安全连接服务，请检查网络、证书或服务地址。') from None
    except (UnicodeError, json.JSONDecodeError):
        raise ServiceError('服务返回的不是有效JSON，请检查接口地址。') from None


def decode_object(text):
    if not isinstance(text, str):
        raise ValueError('返回内容不是文本。')
    text = text.strip()
    if text.startswith('```'):
        text = re.sub(r'^```(?:json)?\s*', '', text, flags=re.I)
        text = re.sub(r'\s*```$', '', text)
    value = json.loads(text)
    if not isinstance(value, dict):
        raise ValueError('返回值必须为JSON对象。')
    return value


class AIClient:
    def __init__(self, settings, cancel=None, transport=None):
        self.settings = dict(DEFAULTS, **settings)
        self.cancel = cancel
        self.transport = transport or post_json
        self.url = endpoint(str(self.settings['base_url']))
        if not self.settings['model'] or not self.settings['api_key']:
            raise ServiceError('请先配置模型名称及API Key。')
        for name, low, high in (('input_chars', 8000, 100000), ('output_tokens', 1000, 16000), ('timeout', 15, 600)):
            try:
                value = int(self.settings[name])
            except (TypeError, ValueError):
                raise ServiceError('%s必须是整数。' % name) from None
            if not low <= value <= high:
                raise ServiceError('%s应在%d至%d之间。' % (name, low, high))
            self.settings[name] = value

    def _chat(self, messages):
        check_cancel(self.cancel)
        size = sum(len(m['content']) for m in messages)
        if size > self.settings['input_chars']:
            raise ServiceError('本次上下文%d字符超过配置上限%d；请提高模型允许的上限或减少提纲/补充输入。'
                               % (size, self.settings['input_chars']))
        response = self.transport(self.url, {'model': self.settings['model'], 'messages': messages,
                                             'max_tokens': self.settings['output_tokens'], 'temperature': 0.25},
                                  self.settings['api_key'], self.settings['timeout'])
        check_cancel(self.cancel)
        try:
            choice = response['choices'][0]
            if choice.get('finish_reason') == 'length':
                raise ServiceError('模型输出被截断；请提高输出token上限或缩短章节，再手动继续。')
            content = choice['message']['content']
            if not isinstance(content, str) or not content.strip():
                raise KeyError('empty')
            return content
        except (KeyError, TypeError, IndexError):
            raise ServiceError('接口没有返回有效文本，请确认兼容Chat Completions。') from None

    def json(self, system, payload, validator=None):
        messages = [{'role': 'system', 'content': system + '\n只返回JSON对象，不添加Markdown代码围栏。'},
                    {'role': 'user', 'content': json.dumps(payload, ensure_ascii=False)}]
        content = self._chat(messages)
        for attempt in range(2):
            try:
                result = decode_object(content)
                if validator:
                    validator(result)
                return result
            except (ValueError, KeyError, TypeError) as exc:
                if attempt:
                    raise ServiceError('返回结构两次校验失败，已停止；请更换模型或简化任务。') from None
                repair = {'任务': '只修复JSON格式与字段结构，不增加事实。', 'schema': payload.get('schema'),
                          'error': str(exc)[:300], 'previous': content[:self.settings['input_chars'] // 2]}
                content = self._chat([messages[0], {'role': 'user', 'content': json.dumps(repair, ensure_ascii=False)}])

    def test(self):
        return self._chat([{'role': 'user', 'content': '这是接口连接测试，只回复：连接成功。'}])


def search_sources(queries, key, cancel=None, transport=None):
    if not key:
        raise ServiceError('联网查证需要单独配置Tavily检索Key。')
    transport = transport or post_json
    results, seen = [], set()
    for query in queries:
        check_cancel(cancel)
        if not query.strip() or len(query) > 200:
            raise ServiceError('每条检索词须为1至200字符，不要输入整篇资料。')
        response = transport('https://api.tavily.com/search', {
            'query': query, 'search_depth': 'advanced', 'max_results': 5,
            'include_raw_content': True, 'include_answer': False,
            'include_domains': ['gov.cn', 'edu.cn']}, key, 60)
        check_cancel(cancel)
        for item in response.get('results', []):
            url = str(item.get('url', ''))
            host = (parse.urlsplit(url).hostname or '').lower()
            if (parse.urlsplit(url).scheme != 'https' or url in seen
                    or not (host.endswith('.gov.cn') or host.endswith('.edu.cn') or host in ('gov.cn', 'edu.cn'))):
                continue
            seen.add(url)
            raw = item.get('raw_content')
            text = raw or item.get('content') or ''
            if not text:
                continue
            # 检索正文过长时明确记录片段范围，不把片段称为完整全文。
            text = str(text)[:12000]
            for offset in range(0, len(text), 2400):
                results.append({'id': 'WEB-%d' % (len(results) + 1), 'file': str(item.get('title', host)),
                                'role': '补充资料', 'location': url, 'url': url,
                                'published': str(item.get('published_date') or '日期未提供'),
                                'retrieved': datetime.now().isoformat(timespec='seconds'),
                                'verification': '官网正文片段' if raw else '仅检索摘要，未充分核验',
                                'text': text[offset:offset + 2400]})
    if not results:
        raise ServiceError('没有取得可引用的政府或院校官网资料；可关闭联网，仅使用上传资料继续。')
    return results

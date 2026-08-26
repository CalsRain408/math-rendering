#!/usr/bin/env python3
"""Formula Canvas 本地 HTTPS 静态服务。

- 服务同目录下的 site/ 目录（插件页面）。
- 使用同目录下的 localhost 自签证书（cert.pem / key.pem），由 install.command 生成。
- 监听 https://localhost:<PORT>/，供 Office 任务窗格加载。

只在本机 127.0.0.1 上监听，不对外暴露；公式渲染全部在浏览器 WebView 内完成，不上传任何内容。
"""
import http.server
import os
import ssl
import sys

PORT = int(os.environ.get("FORMULA_CANVAS_PORT", "3000"))
HERE = os.path.dirname(os.path.abspath(__file__))
SITE_DIR = os.path.join(HERE, "site")
CERT = os.path.join(HERE, "cert.pem")
KEY = os.path.join(HERE, "key.pem")


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=SITE_DIR, **kwargs)

    def end_headers(self):
        # 允许 Office WebView 加载；禁用缓存避免更新后拿到旧文件
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def log_message(self, *args):
        pass  # 静默，避免刷屏


def main():
    if not os.path.isdir(SITE_DIR):
        print(f"[错误] 找不到 site 目录：{SITE_DIR}", file=sys.stderr)
        sys.exit(1)
    if not (os.path.isfile(CERT) and os.path.isfile(KEY)):
        print("[错误] 找不到证书 cert.pem / key.pem，请先运行 install.command", file=sys.stderr)
        sys.exit(1)

    httpd = http.server.ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.load_cert_chain(certfile=CERT, keyfile=KEY)
    httpd.socket = ctx.wrap_socket(httpd.socket, server_side=True)
    print(f"Formula Canvas 本地服务已启动：https://localhost:{PORT}/")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()

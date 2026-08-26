#!/bin/bash
# Formula Canvas 一键安装脚本（macOS）
# 作用：
#   1) 生成本机 localhost HTTPS 证书并加入系统信任
#   2) 安装后台服务（登录时自动启动，立即启动一次）
#   3) 把加载项 manifest 拷到 Word / PowerPoint 的旁加载目录
# 全程只在本机运行，公式渲染在本地完成，不上传任何内容。

set -e

# 切到脚本所在目录（即解压出来的 formula-canvas-mac 目录）
# 用 BASH_SOURCE 解析真实脚本路径，避免双击时 $0 拿到相对/被截断路径导致 cd 失败
SELF="${BASH_SOURCE[0]:-$0}"
SELF_DIR="$(cd "$(dirname "$SELF")" >/dev/null 2>&1 && pwd)"
if [ -z "$SELF_DIR" ] || [ ! -f "$SELF_DIR/server.py" ]; then
  echo "[错误] 无法定位安装目录。请改用终端运行："
  echo "        bash \"<formula-canvas-mac 文件夹的完整路径>/install.command\""
  echo "按回车键关闭..."; read -r _; exit 1
fi
cd "$SELF_DIR"
PKG_DIR="$(pwd)"
PORT=3000
LABEL="com.formulacanvas.localserver"
PLIST="$HOME/Library/LaunchAgents/${LABEL}.plist"

echo "=================================================="
echo " Formula Canvas 安装程序"
echo " 安装目录：$PKG_DIR"
echo "=================================================="
echo

# 找一个可用的 python3
PYTHON="$(command -v python3 || true)"
if [ -z "$PYTHON" ]; then
  PYTHON="/usr/bin/python3"
fi
if [ ! -x "$PYTHON" ] && ! command -v python3 >/dev/null 2>&1; then
  echo "[错误] 未找到 python3。首次可能需要安装命令行工具，系统会弹窗，请点“安装”后重跑本脚本。"
  xcode-select --install 2>/dev/null || true
  exit 1
fi
echo "[1/4] 使用 Python：$PYTHON"

# ---- 1. 生成 localhost 自签证书（带 SAN，WebView 必需）----
echo "[2/4] 生成本机 HTTPS 证书 ..."
OPENSSL="$(command -v openssl || echo /usr/bin/openssl)"
"$OPENSSL" req -x509 -newkey rsa:2048 -nodes \
  -keyout "$PKG_DIR/key.pem" -out "$PKG_DIR/cert.pem" \
  -days 3650 -subj "/CN=localhost" \
  -addext "subjectAltName=DNS:localhost,IP:127.0.0.1" \
  >/dev/null 2>&1
echo "      证书已生成：cert.pem / key.pem"

# 把证书加入系统钥匙串并信任（需要输入一次开机密码）
echo "      将证书加入系统信任（会提示输入你的 Mac 登录密码）..."
sudo security add-trusted-cert -d -r trustRoot \
  -k /Library/Keychains/System.keychain "$PKG_DIR/cert.pem"
echo "      证书已信任。"

# ---- 2. 安装并启动后台服务（LaunchAgent）----
echo "[3/4] 安装后台服务（登录自启）..."
mkdir -p "$HOME/Library/LaunchAgents"
cat > "$PLIST" <<PLISTEOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>${LABEL}</string>
  <key>ProgramArguments</key>
  <array>
    <string>${PYTHON}</string>
    <string>${PKG_DIR}/server.py</string>
  </array>
  <key>EnvironmentVariables</key>
  <dict>
    <key>FORMULA_CANVAS_PORT</key>
    <string>${PORT}</string>
  </dict>
  <key>WorkingDirectory</key>
  <string>${PKG_DIR}</string>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <true/>
  <key>StandardOutPath</key>
  <string>${PKG_DIR}/server.log</string>
  <key>StandardErrorPath</key>
  <string>${PKG_DIR}/server.log</string>
</dict>
</plist>
PLISTEOF

# 重新加载服务
launchctl unload "$PLIST" 2>/dev/null || true
launchctl load "$PLIST"
sleep 2

# 校验服务是否起来
if curl -ks "https://localhost:${PORT}/" -o /dev/null; then
  echo "      服务已启动：https://localhost:${PORT}/"
else
  echo "      [警告] 服务似乎未启动，请查看 ${PKG_DIR}/server.log"
fi

# ---- 3. 拷 manifest 到 Office 旁加载目录 ----
echo "[4/4] 安装加载项到 Word / PowerPoint ..."
WORD_WEF="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
PPT_WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WORD_WEF" "$PPT_WEF"
cp "$PKG_DIR/manifest.xml" "$WORD_WEF/formula-canvas.manifest.xml"
cp "$PKG_DIR/manifest.xml" "$PPT_WEF/formula-canvas.manifest.xml"
echo "      已安装到 Word 和 PowerPoint。"

echo
echo "=================================================="
echo " 安装完成！"
echo
echo " 接下来："
echo "   1. 完全退出 Word 和 PowerPoint（Cmd+Q）。"
echo "   2. 重新打开 Word 或 PowerPoint。"
echo "   3. 菜单「开始」→「加载项」→ 选择 Formula Canvas。"
echo
echo " 如需卸载，运行本目录下的 uninstall.command。"
echo " 注意：请勿删除或移动本文件夹，服务依赖这里的文件。"
echo "=================================================="
echo
echo "按回车键关闭本窗口..."
read -r _

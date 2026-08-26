#!/bin/bash
# Formula Canvas 卸载脚本（macOS）
# 停止并移除后台服务、移除加载项、删除信任的证书。

SELF="${BASH_SOURCE[0]:-$0}"
SELF_DIR="$(cd "$(dirname "$SELF")" >/dev/null 2>&1 && pwd)"
cd "$SELF_DIR" 2>/dev/null || true
PKG_DIR="$(pwd)"
LABEL="com.formulacanvas.localserver"
PLIST="$HOME/Library/LaunchAgents/${LABEL}.plist"

echo "=================================================="
echo " Formula Canvas 卸载程序"
echo "=================================================="
echo

echo "[1/3] 停止并移除后台服务 ..."
launchctl unload "$PLIST" 2>/dev/null || true
rm -f "$PLIST"
echo "      已移除。"

echo "[2/3] 移除 Word / PowerPoint 加载项 ..."
rm -f "$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef/formula-canvas.manifest.xml"
rm -f "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/formula-canvas.manifest.xml"
echo "      已移除。"

echo "[3/3] 移除本机信任的证书 ..."
if [ -f "$PKG_DIR/cert.pem" ]; then
  # 用证书内容匹配删除；需要输入登录密码
  sudo security delete-certificate -c "localhost" /Library/Keychains/System.keychain 2>/dev/null || true
fi
echo "      已处理。"

echo
echo " 卸载完成。你可以删除本文件夹了。"
echo
echo "按回车键关闭本窗口..."
read -r _

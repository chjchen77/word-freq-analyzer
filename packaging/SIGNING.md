# 发布签名说明

当前自动发布流程会生成可直接使用的 Windows 和 macOS 包，但默认不包含开发者证书签名。
因此首次运行时，Windows SmartScreen 或 macOS Gatekeeper 可能显示安全提示。

正式对外分发时建议补充：

- Windows：使用 Authenticode 证书签署 exe 和安装程序，并将证书配置在 GitHub Actions 的安全变量中。
- macOS：使用 Developer ID Application 签名，再使用 `notarytool` 公证并 stapler 到 `.app`。
- 发布页同时保留 `SHA256SUMS.txt`，便于接收者核对文件完整性。

签名需要开发者证书、私钥和相应账号权限，不能由仓库源码自动生成。

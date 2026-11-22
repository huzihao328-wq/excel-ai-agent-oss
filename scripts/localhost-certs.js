/*
 * 为本机 https://localhost 签发并信任证书，不依赖 npm / office-addin-dev-certs。
 * 证书在 %LOCALAPPDATA%\ExcelAIAgent\certs\
 */
const { execFileSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const PASSPHRASE = "excel-ai-agent-localhost";
const certDir = path.join(process.env.LOCALAPPDATA || "", "ExcelAIAgent", "certs");
const pfxPath = path.join(certDir, "localhost.pfx");

function ps(command) {
  return execFileSync("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command], {
    encoding: "utf8",
    windowsHide: true,
  });
}

function pfxReady() {
  if (!fs.existsSync(pfxPath) || fs.statSync(pfxPath).size < 200) return false;
  try {
    ps(`
$ErrorActionPreference = 'Stop'
$pwd = ConvertTo-SecureString -String '${PASSPHRASE}' -Force -AsPlainText
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2('${pfxPath.replace(/\\/g, "\\\\")}', $pwd, 'Exportable')
if ($cert.NotAfter -lt (Get-Date).AddDays(7)) { exit 2 }
if ($cert.DnsNameList.Unicode -notcontains 'localhost') { exit 3 }
`);
    return true;
  } catch {
    return false;
  }
}

function ensureLocalhostCert() {
  fs.mkdirSync(certDir, { recursive: true });
  if (pfxReady()) {
    return { pfx: fs.readFileSync(pfxPath), passphrase: PASSPHRASE };
  }

  const pfxEsc = pfxPath.replace(/'/g, "''");
  ps(`
$ErrorActionPreference = 'Stop'
$certDir = '${certDir.replace(/'/g, "''")}'
New-Item -ItemType Directory -Force -Path $certDir | Out-Null
Get-ChildItem Cert:\\CurrentUser\\My | Where-Object { $_.FriendlyName -eq 'Excel AI Agent localhost' } | ForEach-Object {
  Remove-Item -LiteralPath $_.PSPath -Force
}
$cert = New-SelfSignedCertificate -DnsName 'localhost','127.0.0.1' -CertStoreLocation Cert:\\CurrentUser\\My -FriendlyName 'Excel AI Agent localhost' -NotAfter (Get-Date).AddYears(5) -KeyExportPolicy Exportable -HashAlgorithm SHA256 -KeySpec KeyExchange
$pwd = ConvertTo-SecureString -String '${PASSPHRASE}' -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath '${pfxEsc}' -Password $pwd | Out-Null
$root = New-Object System.Security.Cryptography.X509Certificates.X509Store('Root','CurrentUser')
$root.Open('ReadWrite')
try { $root.Add($cert) } finally { $root.Close() }
`);
  console.log("已生成本机 localhost 证书并写入当前用户受信任根（若弹出安全警告请点「是」）");
  return { pfx: fs.readFileSync(pfxPath), passphrase: PASSPHRASE };
}

module.exports = { ensureLocalhostCert, certDir, pfxPath };

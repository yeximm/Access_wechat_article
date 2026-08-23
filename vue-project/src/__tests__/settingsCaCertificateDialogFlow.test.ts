import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import test from 'node:test'

const settingsPageSource = readFileSync(new URL('../pages/SettingsPage.vue', import.meta.url), 'utf8')

function requireSourceBlock(source: string, pattern: RegExp, label: string) {
  const match = source.match(pattern)

  assert.ok(match, `${label} should exist`)
  return match[0]
}

test('CA 证书检测点击后先打开弹窗，再异步更新检测结果', () => {
  const block = requireSourceBlock(
    settingsPageSource,
    /async function handleCheckCaCertificate\(\)[\s\S]*?\n\}/,
    'handleCheckCaCertificate',
  )

  assert.match(block, /openCaCertificateDialog\(\{[\s\S]*mode:\s*'check'/)
  assert.ok(block.indexOf('openCaCertificateDialog') < block.indexOf('await checkCaCertificate()'))
  assert.doesNotMatch(block, /openDiagnosticResultDialog\(/)
})

test('CA 证书安装先展示项目证书信息，确认后才执行系统安装', () => {
  const openBlock = requireSourceBlock(
    settingsPageSource,
    /async function openInstallCaCertificateDialog\(\)[\s\S]*?\n\}/,
    'openInstallCaCertificateDialog',
  )
  const confirmBlock = requireSourceBlock(
    settingsPageSource,
    /async function confirmInstallCaCertificate\(\)[\s\S]*?\n\}/,
    'confirmInstallCaCertificate',
  )

  assert.match(openBlock, /openCaCertificateDialog\(\{[\s\S]*mode:\s*'install'/)
  assert.ok(openBlock.indexOf('openCaCertificateDialog') < openBlock.indexOf('await checkCaCertificate()'))
  assert.doesNotMatch(openBlock, /await installCaCertificate\(/)
  assert.match(confirmBlock, /caCertificateDialogPhase\.value\s*=\s*'installing'/)
  assert.match(confirmBlock, /await installCaCertificate\(\)/)
  assert.doesNotMatch(confirmBlock, /installCertificateDialogVisible\.value\s*=\s*false/)
})

test('清除 CA 证书先列出系统证书，确认后删除弹窗中列出的指纹', () => {
  const openBlock = requireSourceBlock(
    settingsPageSource,
    /async function handleOpenMitmCertificateDialog\(\)[\s\S]*?\n\}/,
    'handleOpenMitmCertificateDialog',
  )
  const confirmBlock = requireSourceBlock(
    settingsPageSource,
    /async function handleConfirmDeleteMitmCertificates\(\)[\s\S]*?\n\}/,
    'handleConfirmDeleteMitmCertificates',
  )

  assert.match(openBlock, /openCaCertificateDialog\(\{[\s\S]*mode:\s*'delete'/)
  assert.ok(openBlock.indexOf('openCaCertificateDialog') < openBlock.indexOf('await checkCaCertificate()'))
  assert.doesNotMatch(openBlock, /await listMitmCaCertificates\(/)
  assert.match(confirmBlock, /mitmCertificateItems\.value\.map\(\(item\) => item\.thumbprint\)/)
  assert.match(confirmBlock, /await deleteMitmCaCertificates\(thumbprints\)/)
  assert.match(confirmBlock, /caCertificateDialogDeletedItems\.value/)
})

test('CA 证书操作使用同一个业务弹窗承载项目证书和系统证书结果', () => {
  assert.match(settingsPageSource, /caCertificateDialogVisible/)
  assert.match(settingsPageSource, /caCertificateDialogProjectCertificate/)
  assert.match(settingsPageSource, /caCertificateDialogSystemCertificates/)
  assert.match(settingsPageSource, /projectCertificateInstalled/)
  assert.doesNotMatch(settingsPageSource, /<ConfirmDialog[\s\S]*@confirm="confirmInstallCaCertificate"/)
})

test('CA 证书弹窗用来源块区分项目内部证书和系统证书库', () => {
  assert.match(settingsPageSource, /mitm-cert-source-card--project/)
  assert.match(settingsPageSource, /mitm-cert-source-card--system/)
  assert.match(settingsPageSource, /项目内部证书文件/)
  assert.match(settingsPageSource, /Windows 当前用户根证书库/)
  assert.match(settingsPageSource, /项目内/)
  assert.match(settingsPageSource, /系统证书库/)
})

test('项目内部证书块展示证书自身信息，不混入系统位置和文件存在状态', () => {
  const projectBlock = requireSourceBlock(
    settingsPageSource,
    /<section class="mitm-cert-source-card mitm-cert-source-card--project[\s\S]*?<\/section>/,
    'mitm project certificate source card',
  )

  assert.match(projectBlock, /<dt>项目路径<\/dt>/)
  assert.match(projectBlock, /<dt>指纹<\/dt>/)
  assert.match(projectBlock, /<dt>有效期<\/dt>/)
  assert.match(projectBlock, /<dt>颁发者<\/dt>/)
  assert.match(projectBlock, /caCertificateDialogProjectCertificate\?\.notBefore/)
  assert.match(projectBlock, /caCertificateDialogProjectCertificate\?\.notAfter/)
  assert.match(projectBlock, /caCertificateDialogProjectCertificate\?\.issuer/)
  assert.doesNotMatch(projectBlock, /<dt>系统位置<\/dt>/)
  assert.doesNotMatch(projectBlock, /<dt>文件存在<\/dt>/)
  assert.doesNotMatch(projectBlock, /caCertificateStatus\.storePath/)
  assert.doesNotMatch(projectBlock, /caCertificateStatus\.caFileExists/)
})

test('CA certificate dialog install button has footer scoped primary styling', () => {
  const primaryButtonStyle = requireSourceBlock(
    settingsPageSource,
    /\.mitm-cert-dialog-actions\s+:deep\(\.ant-btn\.config-action-button\.primary\)[\s\S]*?\n\}/,
    'mitm certificate dialog primary button style',
  )

  assert.match(primaryButtonStyle, /color:\s*#ffffff/)
  assert.match(primaryButtonStyle, /background:\s*#357FD9/)
  assert.match(primaryButtonStyle, /border-color:\s*#357FD9/)
  assert.doesNotMatch(primaryButtonStyle, /diagnostic-result-dialog/)
})

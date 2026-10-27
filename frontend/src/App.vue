<script setup>
import { ref, computed, watch, nextTick, onMounted, onUnmounted } from 'vue'
const images = ref([]), selected = ref(''), name = ref(''), busy = ref(false), message = ref(''), failed = ref(false)
const zoom = ref(1), pan = ref({x:0,y:0}), cropMode = ref(false), crop = ref(null), photo = ref(null), quality = ref(85), desktop = ref(false)
const token = new URLSearchParams(location.search).get('token') || ''
const brandIcon = '/shitu.png'
const filenameInput = ref(null)
const imageList = ref(null), confirmationDialog = ref(null), pendingRemoval = ref(null), removalError = ref('')
const exportDialog = ref(null)
const canExportFolder = computed(() => desktop.value || typeof window.showDirectoryPicker === 'function' || ['localhost','127.0.0.1','[::1]'].includes(location.hostname))
function openExportDialog() {
  if (busy.value || !images.value.length) return
  failed.value = false
  exportDialog.value.showModal()
}
function closeExportDialog() { if (!busy.value) exportDialog.value.close() }
async function chooseExport(format) {
  const success = await (format === 'folder' ? exportFolder() : exportZip())
  if (success) closeExportDialog()
}
const focusMode = ref(false), showContext = ref(true), notifications = ref([])
const requestedMenu = new URLSearchParams(location.search).get('menu')
const activeMenu = ref(['word','local','certificate'].includes(requestedMenu) ? requestedMenu : 'word')
const menus = [
  {id:'word', label:'Word 提图', hint:'按文档顺序提取图片，同时关联上下文描述。', empty:'导入 Word 文档，开始提取图片', detail:'点击上方“导入 Word 文档”，选择 .docx 文件。'},
  {id:'local', label:'本地图片', hint:'导入已有图片或整个文件夹，统一重命名与编辑。', empty:'添加图片，开始整理', detail:'点击上方“导入图片”或选择图片文件夹。'},
  {id:'certificate', label:'证书截图', hint:'批量处理证书，自选截图字段并查看逐项结果。', empty:'将证书信息变成图片', detail:'点击上方“证书自动截图”，批量添加证书并选择截图字段。'},
]
const menu = computed(() => menus.find(item => item.id === activeMenu.value))
function menuKey(event) {
  if (busy.value || !['ArrowLeft','ArrowRight','Home','End'].includes(event.key)) return
  event.preventDefault()
  const position = menus.findIndex(item => item.id === activeMenu.value)
  const next = event.key === 'Home' ? 0 : event.key === 'End' ? menus.length - 1 : (position + (event.key === 'ArrowRight' ? 1 : -1) + menus.length) % menus.length
  activeMenu.value = menus[next].id
  nextTick(() => document.getElementById(`menu-${activeMenu.value}`)?.focus())
}
const certificateDialog = ref(null), certificate = ref(null)
const certificateQueue = ref([]), certificateErrors = ref([]), certificateJob = ref(null)
const certificateReport = ref(null)
function certificateSummary(job) {
  const success = job.results.filter(row => row.status === 'success').length
  const failedCount = job.results.filter(row => row.status === 'failed').length
  const stopped = job.results.filter(row => row.status === 'cancelled').length
  return `${job.status === 'cancelled' ? '任务已停止' : '截图完成'}：成功 ${success} 份证书，失败 ${failedCount} 份证书${stopped ? `，未处理 ${stopped} 份` : ''}，已加入 ${job.count} 张图片`
}
function certificateFailureText(row) {
  const field = row.failed_field ? `字段“${row.failed_field}”截图失败：` : row.failed_stage === 'import' ? '截图导入失败：' : ''
  return field + (row.error || '未返回具体原因，请重试')
}
const customFields = ref([]), customField = ref('')
const extraOptions = [{id:'serial_number',label:'序列号'},{id:'signature_algorithm',label:'签名算法'},{id:'subject',label:'使用者'},{id:'valid_from',label:'有效期起始'},{id:'valid_to',label:'有效期截止'},{id:'thumbprint',label:'指纹'},{id:'key_usage',label:'密钥用法'},{id:'enhanced_key_usage',label:'增强型密钥用法'},{id:'subject_alt_name',label:'使用者可选名称'},{id:'basic_constraints',label:'基本约束'}]
const captureKinds = computed(() => [...certificateKinds.value, ...customFields.value.map(field => 'field:'+field)])
const certificateStatus = {pending:'等待中',running:'截图中',success:'已完成',failed:'失败',cancelled:'已停止'}
let disposed = false
function addCustomField() {
  const field = customField.value.trim()
  if (!field) return
  if (field.length > 80 || /[\x00-\x1f]/.test(field)) { notify('字段名称需为 1 至 80 个可见字符', 'error'); return }
  if (captureKinds.value.length >= 24) { notify('最多选择 24 个截图项目', 'error'); return }
  if (!customFields.value.includes(field)) customFields.value.push(field)
  customField.value = ''
}
const certificateKinds = ref(['general','issuer','public_key','path'])
const certificateOptions = [{id:'general',label:'常规信息'},{id:'issuer',label:'颁发者'},{id:'public_key',label:'公钥'},{id:'path',label:'证书路径'}]
const notificationTimers = new Map()
let notificationId = 0
function dismissNotification(id) {
  clearTimeout(notificationTimers.get(id))
  notificationTimers.delete(id)
  notifications.value = notifications.value.filter(item => item.id !== id)
}
function notify(text, type = 'success') {
  const id = ++notificationId
  notifications.value.push({id, text, type})
  if (notifications.value.length > 3) dismissNotification(notifications.value[0].id)
  notificationTimers.set(id, setTimeout(() => dismissNotification(id), type === 'error' ? 7000 : 3500))
}
const current = computed(() => images.value.find(i => i.id === selected.value))
const index = computed(() => images.value.findIndex(i => i.id === selected.value))
const url = (image, thumbnail = false) => `/api/images/${image.id}/file?preview=true&thumbnail=${thumbnail}&revision=${image.revision}&token=${encodeURIComponent(token)}`
watch(current, value => { name.value = value?.name.replace(/\.[^.]+$/, '') || ''; zoom.value = 1; pan.value = {x:0,y:0}; crop.value = null; cropMode.value = false })
watch(selected, async () => {
  await nextTick()
  const list = imageList.value, item = list?.querySelector('.image-item.active')
  if (!item) return
  const row = item.closest('.image-row'), rect = row.getBoundingClientRect(), bounds = list.getBoundingClientRect()
  if (rect.top < bounds.top) list.scrollTop -= bounds.top - rect.top
  else if (rect.bottom > bounds.bottom) list.scrollTop += rect.bottom - bounds.bottom
})
async function readJsonResponse(response) {
  let data
  try { data = await response.json() }
  catch {
    const code = response.status ? `（HTTP ${response.status}）` : ''
    throw new Error(response.ok ? '服务返回的数据格式异常，请重试' : `服务端处理失败${code}，请稍后重试；具体原因可查看后端日志`)
  }
  if (!response.ok) throw new Error(typeof data?.detail === 'string' ? data.detail : '请求处理失败，请检查输入或稍后重试')
  return data
}
async function api(path, options = {}) {
  const response = await fetch('/api' + path, {...options, headers: {'x-app-token':token, ...options.headers}})
  return readJsonResponse(response)
}
async function run(action, success = '') {
  if (busy.value) return false
  busy.value = true; message.value = ''; failed.value = false
  try {
    const result = await action()
    if (result?.cancelled) return false
    message.value = typeof success === 'function' ? success(result) : success
    if (message.value) notify(message.value)
    return true
  } catch(e) { message.value = e.message; failed.value = true; notify(e.message, 'error'); return false } finally { busy.value = false }
}
async function refresh() { images.value = await api('/images'); if (!current.value) selected.value = images.value[0]?.id || '' }
async function upload(event) {
  const files = [...event.target.files]; event.target.value = ''; if (!files.length) return
  await run(async () => { const before = images.value.length; const form = new FormData(); files.sort((a,b) => a.name.localeCompare(b.name, undefined, {numeric:true})).forEach(f => form.append('files', f)); images.value = await api('/import', {method:'POST',body:form}); selected.value = images.value[0]?.id || ''; const count = images.value.length - before; if (!count) throw new Error('所选文件中没有找到支持的图片'); return count }, count => `加载完成，已导入 ${count} 张图片`)
}
function move(delta) { const target = images.value[index.value + delta]; if (target && !busy.value) selected.value = target.id }
function useDescription(candidate) {
  if (busy.value) return
  name.value = candidate.name
  failed.value = false
  message.value = '描述已填入文件名称，可修改后点击“保存名称”'
  filenameInput.value?.focus()
}
async function rename() { await run(async () => { images.value = await api(`/images/${selected.value}`, {method:'PATCH',headers:{'Content-Type':'application/json'},body:JSON.stringify({name:name.value})}) }, '名称已保存') }
async function edit(operation) {
  await run(async () => { const result = await api(`/images/${selected.value}/edit`, {method:'POST', headers:{'Content-Type':'application/json'},body:JSON.stringify({operation,box:crop.value,quality:Number(quality.value)})}); images.value = result.images; selected.value = result.selected }, '当前图片已更新')
}
function requestRemoval(image = null) {
  if (busy.value || !images.value.length) return
  pendingRemoval.value = image ? {id:image.id, name:image.name} : {id:null}
  removalError.value = ''
  confirmationDialog.value.showModal()
}
function cancelRemoval() {
  if (busy.value) return
  confirmationDialog.value.close()
  pendingRemoval.value = null
  removalError.value = ''
}
async function confirmRemoval() {
  const action = pendingRemoval.value
  if (!action || busy.value) return
  removalError.value = ''
  const removedIndex = images.value.findIndex(i => i.id === action.id)
  const success = await run(async () => {
    const remaining = await api(action.id ? `/images/${action.id}` : '/images', {method:'DELETE'})
    images.value = remaining
    if (!action.id || selected.value === action.id) {
      selected.value = remaining[Math.min(Math.max(removedIndex, 0), remaining.length - 1)]?.id || ''
    }
    drag = null
    if (!remaining.length) { name.value = ''; zoom.value = 1; pan.value = {x:0,y:0}; crop.value = null; cropMode.value = false }
  }, action.id ? '图片已从工作区删除' : '工作区已清空')
  if (success) cancelRemoval()
  else removalError.value = message.value
}
let drag = null
function point(e) {
  const rect = photo.value.getBoundingClientRect(), item = current.value
  return [Math.round(Math.max(0,Math.min(item.width,(e.clientX-rect.left)/rect.width*item.width))),Math.round(Math.max(0,Math.min(item.height,(e.clientY-rect.top)/rect.height*item.height)))]
}
function down(e) {
  if (!current.value || busy.value) return
  e.currentTarget.setPointerCapture(e.pointerId)
  drag = {x:e.clientX,y:e.clientY,pan:{...pan.value},start:point(e)}
  if (cropMode.value) crop.value = null
}
function motion(e) {
  if (!drag) return
  if (cropMode.value) { const [x,y] = point(e), [sx,sy] = drag.start; crop.value = [Math.min(x,sx),Math.min(y,sy),Math.max(x,sx),Math.max(y,sy)] }
  else pan.value = {x:drag.pan.x+e.clientX-drag.x,y:drag.pan.y+e.clientY-drag.y}
}
const cropStyle = computed(() => { const b=crop.value,i=current.value; return b&&i ? {left:b[0]/i.width*100+'%',top:b[1]/i.height*100+'%',width:(b[2]-b[0])/i.width*100+'%',height:(b[3]-b[1])/i.height*100+'%'} : {} })
function key(e) { if (pendingRemoval.value || certificateDialog.value?.open || exportDialog.value?.open) return; if (e.key === 'Escape') { focusMode.value = false; return } if (['INPUT','TEXTAREA','SELECT','BUTTON'].includes(e.target.tagName)) return; if (['ArrowUp','ArrowLeft','ArrowDown','ArrowRight'].includes(e.key)) { e.preventDefault(); move(['ArrowUp','ArrowLeft'].includes(e.key)?-1:1) } }
async function openCertificateDialog() {
  if (busy.value) return
  const cleared = await run(async () => {
    await api('/certificates', {method:'DELETE'})
    certificateQueue.value = []
    certificate.value = null
    certificateErrors.value = []
    certificateJob.value = null
  })
  if (cleared && !disposed) certificateDialog.value.showModal()
}
function closeCertificateDialog() { if (!busy.value) certificateDialog.value.close() }
async function chooseCertificates(event) {
  const files = [...event.target.files]; event.target.value = ''
  if (!files.length) return
  if (files.length > 50) { notify('每批最多添加 50 张证书', 'error'); return }
  await run(async () => {
    const form = new FormData()
    files.forEach(file => form.append('files', file))
    const result = await api('/certificates/import', {method:'POST', body:form})
    certificateQueue.value.push(...result.certificates)
    certificateErrors.value = result.errors
    certificateJob.value = null
    certificate.value = certificate.value || certificateQueue.value[0] || null
    if (result.errors.length) notify(`${result.errors.length} 个文件未添加，请查看原因`, 'error')
    return result.certificates.length
  }, count => count ? `已添加 ${count} 张证书` : '')
}
async function removeCertificate(item) {
  await run(async () => {
    await api('/certificates/'+item.id, {method:'DELETE'})
    certificateQueue.value = certificateQueue.value.filter(entry => entry.id !== item.id)
    if (certificate.value?.id === item.id) certificate.value = certificateQueue.value[0] || null
    certificateJob.value = null
  })
}
async function cancelCertificateJob() {
  if (!certificateJob.value?.id) return
  try { certificateJob.value = await api(`/certificate-jobs/${certificateJob.value.id}/cancel`, {method:'POST'}) }
  catch(e) { notify(e.message, 'error') }
}
async function captureCertificate() {
  if (!certificateQueue.value.length || !captureKinds.value.length) return
  if (captureKinds.value.length > 24) { notify('最多选择 24 个截图项目', 'error'); return }
  let completed = false
  await run(async () => {
    certificateErrors.value = []
    if (!certificateJob.value || !['running','cancelling'].includes(certificateJob.value.status)) certificateJob.value = await api('/certificate-jobs', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({certificates:certificateQueue.value.map(item => item.id), kinds:captureKinds.value})})
    while (['running','cancelling'].includes(certificateJob.value.status)) {
      await new Promise(resolve => setTimeout(resolve, 750))
      if (disposed) return {cancelled:true}
      certificateJob.value = await api('/certificate-jobs/'+certificateJob.value.id)
    }
    completed = true
    certificateReport.value = JSON.parse(JSON.stringify(certificateJob.value))
    await refresh()
    if (certificateJob.value.selected && images.value.some(item => item.id === certificateJob.value.selected)) selected.value = certificateJob.value.selected
    const job = certificateJob.value, errors = job.results.filter(row => row.status === 'failed')
    const text = certificateSummary(job)
    if (errors.length) throw new Error(`${text}。${errors[0].filename}：${certificateFailureText(errors[0])}${errors.length > 1 ? '；其余失败详情见工作区截图结果。' : ''}`)
    return {summary:text}
  }, result => result?.summary || '')
  if (completed && !disposed) closeCertificateDialog()
}
async function local(action) {
  return run(async () => {
    const result = await window.pywebview.api[action]()
    if (result?.error) throw new Error(result.error)
    if (result?.cancelled) return result
    if (action === 'import_folder') await refresh()
    return result
  }, result => action === 'import_folder' ? `加载完成，已导入 ${result.count} 张图片` : `导出完成，已保存 ${result.count} 张图片`)
}
async function writeExportFile(handle, data) {
  const writer = await handle.createWritable()
  try { await writer.write(data); await writer.close() }
  catch (error) { try { await writer.abort() } catch {} throw error }
}
async function exportFolder() {
  if (!images.value.length || busy.value) return false
  if (desktop.value) return local('export_folder')
  return run(async () => {
    const systemFolder = () => api('/export/folder', {method:'POST'})
    if (typeof window.showDirectoryPicker !== 'function') return systemFolder()
    let directory
    try { directory = await window.showDirectoryPicker({mode:'readwrite'}) }
    catch(error) {
      if (error.name === 'AbortError') return {cancelled:true}
      if (['SecurityError','NotAllowedError','NotSupportedError'].includes(error.name)) return systemFolder()
      throw new Error(`无法选择目标文件夹：${error.message || error.name}。请使用系统文件夹选择框重试`)
    }
    const items = images.value.map(item => ({...item}))
    for (const item of items) {
      try { await directory.getFileHandle(item.name); throw new Error(`目标文件夹已存在“${item.name}”，请选择空文件夹，避免覆盖`) }
      catch(error) { if (error.name !== 'NotFoundError') { if (error.name === 'TypeMismatchError') throw new Error(`目标文件夹已存在同名目录“${item.name}”，请选择空文件夹`); throw error } }
    }
    // Read every source before writing so a missing image creates no partial export.
    const files = []
    for (const item of items) {
      const response = await fetch(`/api/images/${item.id}/file`, {headers:{'x-app-token':token}})
      if (!response.ok) await readJsonResponse(response)
      files.push({name:item.name, data:await response.blob()})
    }
    let saved = 0
    try {
      for (const file of files) {
        const handle = await directory.getFileHandle(file.name, {create:true})
        await writeExportFile(handle, file.data)
        saved += 1
      }
    } catch { throw new Error(`导出中断：已保存 ${saved} / ${files.length} 张图片，请检查文件夹权限或磁盘空间；未完成的文件可能不完整`) }
    return {count:saved, folder:directory.name}
  }, result => `导出完成，已将 ${result.count} 张图片保存到文件夹“${result.folder}”`)
}
async function exportSystemFolder() {
  if (!images.value.length || busy.value) return false
  if (desktop.value) return local('export_folder')
  const success = await run(() => api('/export/folder', {method:'POST'}), result => `导出完成，已将 ${result.count} 张图片保存到文件夹“${result.folder}”`)
  if (success) closeExportDialog()
  return success
}
async function exportZip() {
  if (!images.value.length || busy.value) return
  if (desktop.value) return local('export_zip')
  return run(async () => {
    let handle
    if (typeof window.showSaveFilePicker === 'function') {
      try { handle = await window.showSaveFilePicker({suggestedName:'拾图图片.zip', types:[{description:'ZIP 压缩包', accept:{'application/zip':['.zip']}}]}) }
      catch(error) { if (error.name === 'AbortError') return {cancelled:true}; throw new Error('无法选择 ZIP 保存位置，请检查文件权限') }
    }
    const response = await fetch('/api/export', {headers:{'x-app-token':token}})
    if (!response.ok) await readJsonResponse(response)
    const blob = await response.blob()
    if (handle) { await writeExportFile(handle, blob); return {saved:true} }
    const objectUrl = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = objectUrl; link.download = '拾图图片.zip'; document.body.appendChild(link)
    try { link.click() } finally { link.remove(); setTimeout(() => URL.revokeObjectURL(objectUrl), 60000) }
    return {saved:false}
  }, result => result.saved ? '导出完成，ZIP 压缩包已保存' : 'ZIP 压缩包已生成，已开始下载')
}
function ready() { desktop.value = true }
onMounted(async () => { desktop.value = !!window.pywebview?.api; window.addEventListener('pywebviewready',ready); window.addEventListener('keydown',key); await run(refresh); if (!disposed && requestedMenu === 'certificate') openCertificateDialog() })
onUnmounted(() => { disposed = true; window.removeEventListener('keydown',key); window.removeEventListener('pywebviewready',ready); notificationTimers.forEach(clearTimeout); notificationTimers.clear() })
</script>

<template>
  <div :class="['app', {'focus-mode': focusMode}]">
    <header class="app-header"><div class="brand"><img class="brand-icon" :src="brandIcon" alt="拾图 Shitu"><h1>拾图 Shitu</h1></div><nav class="feature-menu" role="tablist" aria-label="功能菜单"><button v-for="item in menus" :id="`menu-${item.id}`" :key="item.id" :data-menu="item.id" role="tab" :aria-selected="activeMenu===item.id" :aria-controls="`panel-${item.id}`" :tabindex="activeMenu===item.id ? 0 : -1" :disabled="busy" :class="{active:activeMenu===item.id}" @click="activeMenu=item.id" @keydown="menuKey">{{ item.label }}</button></nav><span class="badge">{{ desktop ? '桌面版 · 本地处理' : '本地图片工作台' }}</span></header>
    <div class="command-bar">
      <section :id="`panel-${activeMenu}`" class="feature-panel" role="tabpanel" :aria-labelledby="`menu-${activeMenu}`">
        <template v-if="activeMenu==='word'"><label class="import">＋ 导入 Word 文档<input type="file" accept=".docx" :disabled="busy" @change="upload"></label></template>
        <template v-else-if="activeMenu==='local'"><label class="import">导入图片<input type="file" accept="image/*" multiple :disabled="busy" @change="upload"></label><label v-if="!desktop" class="folder">选择图片文件夹<input type="file" webkitdirectory multiple :disabled="busy" @change="upload"></label><button v-else :disabled="busy" @click="local('import_folder')">选择本地图片文件夹</button></template>
        <template v-else><button class="certificate-entry primary" :disabled="busy" @click="openCertificateDialog">▣ 证书自动截图</button></template>
        <p class="feature-hint">{{ menu.hint }}</p>
      </section>
      <div class="workspace-actions" role="group" aria-label="工作区操作"><button class="primary" :disabled="!images.length||busy" @click="openExportDialog">↓ 导出全部图片</button><button class="clear-workspace" :disabled="!images.length||busy" @click="requestRemoval()">清空工作区</button></div>
    </div>
    <div class="layout">
      <aside aria-label="工作区图片">
        <div class="section-title"><h2>图片列表</h2><span>{{ images.length }} 张</span></div>
        <div ref="imageList" class="image-list" aria-label="图片列表" tabindex="0"><div v-for="(image,i) in images" :key="image.id" class="image-row"><button :class="['image-item',{active:selected===image.id}]" :disabled="busy" :aria-pressed="selected===image.id" @click="selected=image.id"><img :src="url(image,true)" alt="" loading="lazy" decoding="async"><span><b>{{ image.name }}</b><small>{{ image.width }} × {{ image.height }}</small></span><em>{{ String(i+1).padStart(2,'0') }}</em></button><button class="delete-image" :disabled="busy" :aria-label="'删除 '+image.name" :title="'删除 '+image.name" @click="requestRemoval(image)">×</button></div><p v-if="!images.length" class="empty-list">图片将在这里按顺序排列</p></div>
        <p class="list-note">所有菜单共用此图片列表<br>选择图片后，可在右侧重命名和编辑</p>
      </aside>
      <main>
        <div class="workspace-heading"><div><h2>预览与编辑</h2><p class="muted">{{ current ? `${current.width} × ${current.height} px · ${(current.size/1024).toFixed(1)} KB` : '支持 DOCX、PNG、JPG、GIF、BMP、TIFF 和 WebP' }}</p></div><div class="heading-actions"><button v-if="current?.contexts?.length" @click="showContext=!showContext">{{ showContext ? '收起参考描述' : '显示参考描述' }}</button><button :disabled="!current" @click="focusMode=!focusMode">{{ focusMode ? '退出大图' : '大图模式' }}</button><nav class="image-pagination" aria-label="图片翻页"><button :disabled="index<=0||busy" @click="move(-1)">← 上一张</button><span>{{ images.length ? index+1 : 0 }} / {{ images.length }}</span><button :disabled="index>=images.length-1||busy" @click="move(1)">下一张 →</button></nav></div></div>
        <section v-if="certificateReport" class="certificate-report" aria-label="截图处理结果" aria-live="polite">
          <div class="report-heading"><b>{{ certificateSummary(certificateReport) }}</b><button aria-label="关闭截图结果" @click="certificateReport=null">×</button></div>
          <details v-if="certificateReport.results.some(row=>row.status==='failed')" open><summary>失败详情（证书 / 字段 / 原因）</summary><div class="report-failures"><article v-for="row in certificateReport.results.filter(item=>item.status==='failed')" :key="row.id"><b>{{ row.filename }}</b><span>{{ row.failed_field ? '失败字段：'+row.failed_field : row.failed_stage==='import' ? '失败阶段：导入截图' : '失败阶段：证书截图' }}</span><p>{{ row.error || '未返回具体原因，请重试' }}</p></article></div><p class="report-note">失败证书的截图未导入；此前成功证书的结果已保留。请调整字段或证书后重试。</p></details>
        </section>
        <fieldset :disabled="!current||busy" class="rename"><label for="filename">文件名称</label><input id="filename" ref="filenameInput" v-model="name" @keydown.enter="rename"><span>{{ current?.name.match(/\.[^.]+$/)?.[0] }}</span><button class="primary" @click="rename">保存名称</button></fieldset>
        <div :class="['preview-layout', { 'with-context': current?.contexts?.length && showContext }]">
        <div class="stage" :class="{cropping:cropMode}" @pointerdown="down" @pointermove="motion" @pointerup="drag=null" @pointercancel="drag=null" @wheel.prevent="zoom=Math.max(.1,Math.min(4,zoom*($event.deltaY<0?1.1:.9)))">
          <div v-if="current" class="picture" :style="{transform:`translate(${pan.x}px,${pan.y}px) scale(${zoom})`}"><img ref="photo" :src="url(current)" :alt="current.name" draggable="false"><div v-if="crop" class="crop-box" :style="cropStyle"></div></div>
          <div v-else class="placeholder"><div class="placeholder-icon">▧</div><h2>{{ menu.empty }}</h2><p>{{ menu.detail }}</p><small>导入后，可在此预览、重命名、旋转、裁剪和压缩</small></div>
        </div>
        <section v-if="current?.contexts?.length && showContext" class="document-context" aria-label="文档参考描述">
          <h2>文档参考描述</h2>
          <p class="context-hint">点击描述填入名称，再确认保存。文字按文档结构关联，浮动图片请核对实际对应关系。</p>
          <article v-for="(context, occurrence) in current.contexts" :key="occurrence" class="context-occurrence">
            <div class="context-source">{{ context.document }} · 第 {{ context.paragraph }} 段<span v-if="current.contexts.length > 1"> · 第 {{ occurrence + 1 }} 处引用</span></div>
            <template v-if="context.above.length"><h3>图片上方</h3><p v-for="(text, i) in context.above" :key="i" class="context-text">{{ text }}</p></template>
            <template v-if="context.inline_before || context.inline_after"><h3>图片同段文字</h3><p v-if="context.inline_before" class="context-text">前：{{ context.inline_before }}</p><p v-if="context.inline_after" class="context-text">后：{{ context.inline_after }}</p></template>
            <div class="context-position">▧ 当前图片</div>
            <template v-if="context.below.length"><h3>图片下方</h3><p v-for="(text, i) in context.below" :key="i" class="context-text">{{ text }}</p></template>
            <template v-if="context.alt_text.length"><h3>图片替代文字</h3><p v-for="(text, i) in context.alt_text" :key="i" class="context-text">{{ text }}</p></template>
            <div v-if="context.candidates.length" class="name-candidates"><h3>选择命名参考</h3><button v-for="(candidate,i) in context.candidates" :key="i" :disabled="busy" :title="candidate.text" @click="useDescription(candidate)"><small>{{ candidate.source }}<b v-if="candidate.is_caption"> · 疑似题注</b></small><span>{{ candidate.name }}</span><em>填入名称 ↗</em></button></div>
            <p v-else class="context-hint">附近未找到可用文字，请手动命名。</p>
          </article>
        </section>
        </div>
        <div class="view-controls"><span>{{ cropMode ? '在图片上拖动选择裁剪区域' : '滚轮缩放 · 拖动平移 · 方向键切换图片' }}</span><div><button :disabled="!current" @click="zoom=Math.max(.1,zoom-.1)">−</button><span>{{ Math.round(zoom*100) }}%</span><button :disabled="!current" @click="zoom=Math.min(4,zoom+.1)">＋</button><button @click="zoom=1;pan={x:0,y:0}">重置视图</button></div></div>
        <fieldset :disabled="!current||busy" class="tools"><div><span class="tool-label">旋转与翻转</span><button @click="edit('left')">↶ 左转 90°</button><button @click="edit('right')">↷ 右转 90°</button><button @click="edit('horizontal')">水平翻转</button><button @click="edit('vertical')">垂直翻转</button></div><div><span class="tool-label">裁剪与压缩</span><button @click="cropMode=!cropMode;crop=null">{{ cropMode?'取消裁剪':'选区裁剪' }}</button><button :disabled="!crop" @click="edit('crop')">保存裁剪</button><label>JPEG 质量 <input v-model="quality" type="number" min="1" max="100"></label><button @click="edit('compress')">压缩图片</button></div><p>操作直接更新当前图片，保留名称与参考描述；导出时使用修改后的结果。PNG 使用无损优化。</p></fieldset>
        <div role="status" aria-live="polite" :class="['status',{error:failed}]">{{ busy ? '正在处理，请稍候…' : message || '就绪 · 图片在当前工作区处理，请在退出前导出' }}</div>
      </main>
    </div><footer>拾图 SHITU · 2.6.0 <span>让每一张图片，各得其名。</span></footer>
    <dialog ref="confirmationDialog" aria-labelledby="removal-title" class="removal-dialog" @cancel.prevent="cancelRemoval">
      <h2 id="removal-title">{{ pendingRemoval?.id ? '删除这张图片？' : '清空工作区？' }}</h2>
      <p>{{ pendingRemoval?.id ? `将从当前工作区删除“${pendingRemoval.name}”。` : `将清空当前工作区的 ${images.length} 张图片及其编辑结果，请先导出需要的内容。` }}</p>
      <p class="muted">源 Word 文档和导入前的本地文件不会被删除。</p>
      <p v-if="removalError" role="alert" class="removal-error">{{ removalError }}</p>
      <div class="dialog-actions"><button :disabled="busy" autofocus @click="cancelRemoval">取消</button><button class="danger" :disabled="busy" @click="confirmRemoval">{{ busy ? '正在处理…' : pendingRemoval?.id ? '确认删除' : '确认清空' }}</button></div>
    </dialog>
    <dialog ref="certificateDialog" class="certificate-dialog" aria-labelledby="certificate-title" @cancel.prevent="closeCertificateDialog">
      <div class="certificate-heading"><div><h2 id="certificate-title">证书自动截图</h2><p class="muted">自定义字段 · 批量处理 · Windows 原生截图</p></div><button :disabled="busy" aria-label="关闭证书截图" @click="closeCertificateDialog">×</button></div>
      <section v-if="certificateJob" class="certificate-progress" aria-label="批量截图进度" aria-live="polite"><div class="batch-heading"><b>{{ certificateJob.status==='cancelling' ? '正在停止，等待当前证书完成…' : ['running'].includes(certificateJob.status) ? '正在批量截图…' : certificateJob.status==='cancelled' ? '任务已停止' : '批量处理结果' }}</b><span>{{ certificateJob.completed }} / {{ certificateJob.total }} 张证书 · {{ certificateJob.count }} 张图片</span></div><progress :value="certificateJob.completed" :max="certificateJob.total"></progress><div class="batch-results"><div v-for="row in certificateJob.results" :key="row.id" :class="['batch-result',row.status]"><span>{{ row.filename }}</span><b>{{ certificateStatus[row.status] }}{{ row.count ? ` · ${row.count} 张图片` : '' }}</b><small v-if="row.error">{{ certificateFailureText(row) }}</small></div></div><button v-if="['running','cancelling'].includes(certificateJob.status)" :disabled="certificateJob.status==='cancelling'" @click="cancelCertificateJob">停止后续处理</button></section>
      <div class="certificate-import"><label class="import">批量添加证书<input class="certificate-upload" type="file" accept=".crt,.cer,.pem" multiple :disabled="busy" @change="chooseCertificates"></label><span class="muted">支持 .crt、.cer、.pem；每批最多 50 张，每张不超过 2 MB</span></div>
      <div class="certificate-queue" aria-label="证书队列"><div v-for="item in certificateQueue" :key="item.id" :class="['certificate-row',{active:certificate?.id===item.id}]"><button :disabled="busy" @click="certificate=item"><b>{{ item.filename }}</b><small>{{ item.name }}</small></button><button :disabled="busy" :aria-label="'移除证书 '+item.filename" @click="removeCertificate(item)">×</button></div><p v-if="!certificateQueue.length" class="muted">尚未添加证书，可一次选择多个文件。</p></div>
      <ul v-if="certificateErrors.length" class="certificate-errors" aria-label="证书导入错误"><li v-for="(item,i) in certificateErrors" :key="i">{{ item.filename }}：{{ item.error }}</li></ul>
      <dl v-if="certificate" class="certificate-info"><dt>名称</dt><dd>{{ certificate.name }}</dd><dt>颁发者</dt><dd>{{ certificate.issuer }}</dd><dt>公钥</dt><dd>{{ certificate.public_key }}</dd><dt>有效期（UTC）</dt><dd>{{ certificate.not_before }} 至 {{ certificate.not_after }}</dd></dl>
      <fieldset class="certificate-options" :disabled="busy"><legend>选择截图项目（{{ captureKinds.length }} / 24）</legend><label v-for="option in certificateOptions" :key="option.id"><input v-model="certificateKinds" type="checkbox" :value="option.id">{{ option.label }}</label></fieldset>
      <fieldset class="certificate-options extra-options" :disabled="busy"><legend>更多详细字段</legend><label v-for="option in extraOptions" :key="option.id"><input v-model="certificateKinds" type="checkbox" :value="option.id">{{ option.label }}</label></fieldset>
      <fieldset class="custom-fields" :disabled="busy"><label for="custom-certificate-field">自定义详细字段</label><div class="custom-field-input"><input id="custom-certificate-field" v-model="customField" maxlength="80" placeholder="输入 Windows 字段名称，如：证书策略" @keydown.enter.prevent="addCustomField"><button @click="addCustomField">添加字段</button></div><div class="field-tags"><span v-for="field in customFields" :key="field">{{ field }}<button :aria-label="'移除字段 '+field" @click="customFields=customFields.filter(item=>item!==field)">×</button></span></div><p class="certificate-note">按系统“详细信息”中的字段名称精确匹配，不区分英文大小写。最多 24 项；字段缺失时该证书失败，继续下一张。</p></fieldset>

      <p class="certificate-note">截图时会短暂打开证书窗口，请保持 Windows 桌面解锁，避免操作该窗口。每张证书最多等待 75 秒；同一证书全部所选字段完成后导入，成功结果保留。网页模式需在 Windows 本机运行后端。</p>
      <p class="certificate-note">只读查看，不安装证书。证书路径使用本机证书库和缓存，不联网补全；状态提示按 Windows 实际结果保留。</p>
      <p v-if="failed && message" class="removal-error" role="alert">{{ message }}</p>
      <div class="dialog-actions"><button :disabled="busy" @click="closeCertificateDialog">关闭</button><button class="primary" :disabled="busy || !certificateQueue.length || !captureKinds.length || captureKinds.length>24" @click="captureCertificate">{{ busy ? '正在处理，请稍候…' : certificateJob && ['running','cancelling'].includes(certificateJob.status) ? '继续获取进度' : `开始截图${certificateQueue.length ? '（'+certificateQueue.length+' 张证书）' : ''}` }}</button></div>
    </dialog>
    <dialog ref="exportDialog" class="export-dialog" aria-labelledby="export-title" @cancel.prevent="closeExportDialog">
      <h2 id="export-title">导出全部图片</h2><p class="muted">共 {{ images.length }} 张，保留当前文件名称、格式和编辑结果。</p>
      <div class="export-options"><div><button class="primary" :disabled="busy || !canExportFolder" @click="chooseExport('folder')">保存到文件夹</button><p>{{ canExportFolder ? '选择目标文件夹，逐张保存图片；遇到同名文件会停止，避免覆盖。' : '系统文件夹导出仅支持本机服务；也可使用 ZIP 压缩包导出。' }}</p></div><div><button :disabled="busy" @click="chooseExport('zip')">保存为 ZIP 压缩包</button><p>将全部图片打包为一个 ZIP 文件，方便保存和分发。</p></div></div>
      <p v-if="busy" class="export-progress" role="status">正在导出，请在系统弹出的文件夹选择框中确认目标目录；若选择框在后台，请切换到“拾图 Shitu”文件夹选择窗口。</p><p v-if="failed && message" class="removal-error" role="alert">{{ message }}</p>
      <div class="dialog-actions"><button v-if="!desktop && canExportFolder" :disabled="busy" @click="exportSystemFolder">使用系统文件夹选择框</button><button :disabled="busy" @click="closeExportDialog">取消</button></div>
    </dialog>
    <div class="toast-stack" aria-live="polite" aria-atomic="false"><div v-for="item in notifications" :key="item.id" :class="['toast',item.type]" :role="item.type === 'error' ? 'alert' : 'status'"><span class="toast-icon" aria-hidden="true">{{ item.type === 'error' ? '!' : '✓' }}</span><span>{{ item.text }}</span><button :aria-label="'关闭提示：'+item.text" @click="dismissNotification(item.id)">×</button></div></div>
  </div>
</template>

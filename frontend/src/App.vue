<script setup>
import { ref, computed, watch, nextTick, onMounted, onUnmounted } from 'vue'
const images = ref([]), selected = ref(''), name = ref(''), busy = ref(false), message = ref(''), failed = ref(false)
const zoom = ref(1), pan = ref({x:0,y:0}), cropMode = ref(false), crop = ref(null), photo = ref(null), quality = ref(85), desktop = ref(false)
const token = new URLSearchParams(location.search).get('token') || ''
const brandIcon = '/shitu.png'
const filenameInput = ref(null)
const imageList = ref(null), confirmationDialog = ref(null), pendingRemoval = ref(null), removalError = ref('')
const focusMode = ref(false), showContext = ref(true), notifications = ref([])
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
async function api(path, options = {}) {
  const response = await fetch('/api' + path, {...options, headers: {'x-app-token':token, ...options.headers}})
  if (!response.ok) { const error = await response.json(); throw new Error(typeof error.detail === 'string' ? error.detail : '参数无效，请检查输入') }
  return response.json()
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
function key(e) { if (pendingRemoval.value) return; if (e.key === 'Escape') { focusMode.value = false; return } if (['INPUT','TEXTAREA','SELECT','BUTTON'].includes(e.target.tagName)) return; if (['ArrowUp','ArrowLeft','ArrowDown','ArrowRight'].includes(e.key)) { e.preventDefault(); move(['ArrowUp','ArrowLeft'].includes(e.key)?-1:1) } }
async function local(action) {
  await run(async () => {
    const result = await window.pywebview.api[action]()
    if (result?.error) throw new Error(result.error)
    if (result?.cancelled) return result
    if (action === 'import_folder') await refresh()
    return result
  }, result => action === 'import_folder' ? `加载完成，已导入 ${result.count} 张图片` : `导出完成，已保存 ${result.count} 张图片`)
}
async function exportZip() {
  if (!images.value.length || busy.value) return
  if (desktop.value) return local('export_zip')
  await run(async () => {
    const response = await fetch('/api/export', {headers:{'x-app-token':token}})
    if (!response.ok) { const error = await response.json(); throw new Error(error.detail || '导出失败') }
    const blob = await response.blob(), objectUrl = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = objectUrl; link.download = '拾图图片.zip'; document.body.appendChild(link)
    try { link.click() } finally { link.remove(); setTimeout(() => URL.revokeObjectURL(objectUrl), 60000) }
  }, 'ZIP 已生成，已开始下载')
}
function ready() { desktop.value = true }
onMounted(() => { run(refresh); desktop.value = !!window.pywebview?.api; window.addEventListener('pywebviewready',ready); window.addEventListener('keydown',key) })
onUnmounted(() => { window.removeEventListener('keydown',key); window.removeEventListener('pywebviewready',ready); notificationTimers.forEach(clearTimeout); notificationTimers.clear() })
</script>

<template>
  <div :class="['app', {'focus-mode': focusMode}]">
    <header><div class="brand"><img class="brand-icon" :src="brandIcon" alt="拾图 Shitu"><div><h1>拾图 Shitu <small>图片工作台</small></h1><p>从文档到图片，让整理更简单。</p></div></div><span class="badge">{{ desktop ? '桌面版 · 本地处理' : '本地图片工作台' }}</span></header>
    <div class="layout">
      <aside><h2>开始整理</h2><p class="muted">导入 Word，或直接编辑已有图片</p>
        <label class="import">＋ 导入 Word 文档<input type="file" accept=".docx" :disabled="busy" @change="upload"></label>
        <label class="import secondary">导入图片<input type="file" accept="image/*" multiple :disabled="busy" @change="upload"></label>
        <label v-if="!desktop" class="folder">选择图片文件夹<input type="file" webkitdirectory multiple :disabled="busy" @change="upload"></label>
        <button v-if="desktop" :disabled="busy" @click="local('import_folder')">选择本地图片文件夹</button>
        <div class="section-title"><h2>图片列表</h2><span>{{ images.length }} 张</span></div>
        <div ref="imageList" class="image-list" aria-label="图片列表" tabindex="0"><div v-for="(image,i) in images" :key="image.id" class="image-row"><button :class="['image-item',{active:selected===image.id}]" :disabled="busy" :aria-pressed="selected===image.id" @click="selected=image.id"><img :src="url(image,true)" alt="" loading="lazy" decoding="async"><span><b>{{ image.name }}</b><small>{{ image.width }} × {{ image.height }}</small></span><em>{{ String(i+1).padStart(2,'0') }}</em></button><button class="delete-image" :disabled="busy" :aria-label="'删除 '+image.name" :title="'删除 '+image.name" @click="requestRemoval(image)">×</button></div><p v-if="!images.length" class="empty-list">图片将在这里按顺序排列</p></div>
        <div class="exports"><button class="primary" :disabled="!images.length||busy" @click="exportZip">↓ 导出全部图片 ZIP</button><button v-if="desktop" :disabled="!images.length||busy" @click="local('export_folder')">导出到本地文件夹</button><button :disabled="!images.length||busy" @click="requestRemoval()">清空工作区</button></div>
      </aside>
      <main>
        <div class="workspace-heading"><div><h2>预览与编辑</h2><p class="muted">{{ current ? `${current.width} × ${current.height} px · ${(current.size/1024).toFixed(1)} KB` : '支持 DOCX、PNG、JPG、GIF、BMP、TIFF 和 WebP' }}</p></div><div class="heading-actions"><button v-if="current?.contexts?.length" @click="showContext=!showContext">{{ showContext ? '收起参考描述' : '显示参考描述' }}</button><button :disabled="!current" @click="focusMode=!focusMode">{{ focusMode ? '退出大图' : '大图模式' }}</button><nav><button :disabled="index<=0||busy" @click="move(-1)">← 上一张</button><span>{{ images.length ? index+1 : 0 }} / {{ images.length }}</span><button :disabled="index>=images.length-1||busy" @click="move(1)">下一张 →</button></nav></div></div>
        <fieldset :disabled="!current||busy" class="rename"><label for="filename">文件名称</label><input id="filename" ref="filenameInput" v-model="name" @keydown.enter="rename"><span>{{ current?.name.match(/\.[^.]+$/)?.[0] }}</span><button class="primary" @click="rename">保存名称</button></fieldset>
        <div :class="['preview-layout', { 'with-context': current?.contexts?.length && showContext }]">
        <div class="stage" :class="{cropping:cropMode}" @pointerdown="down" @pointermove="motion" @pointerup="drag=null" @pointercancel="drag=null" @wheel.prevent="zoom=Math.max(.1,Math.min(4,zoom*($event.deltaY<0?1.1:.9)))">
          <div v-if="current" class="picture" :style="{transform:`translate(${pan.x}px,${pan.y}px) scale(${zoom})`}"><img ref="photo" :src="url(current)" :alt="current.name" draggable="false"><div v-if="crop" class="crop-box" :style="cropStyle"></div></div>
          <div v-else class="placeholder"><div class="placeholder-icon">▧</div><h2>把文档中的图片，轻松拾起来</h2><p>从左侧导入 .docx 文档，即可按文档顺序提取图片</p><small>也可以导入图片或文件夹，开始重命名与编辑</small></div>
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
    </div><footer>拾图 SHITU · 2.3.0 <span>让每一张图片，各得其名。</span></footer>
    <dialog ref="confirmationDialog" aria-labelledby="removal-title" class="removal-dialog" @cancel.prevent="cancelRemoval">
      <h2 id="removal-title">{{ pendingRemoval?.id ? '删除这张图片？' : '清空工作区？' }}</h2>
      <p>{{ pendingRemoval?.id ? `将从当前工作区删除“${pendingRemoval.name}”。` : `将清空当前工作区的 ${images.length} 张图片及其编辑结果，请先导出需要的内容。` }}</p>
      <p class="muted">源 Word 文档和导入前的本地文件不会被删除。</p>
      <p v-if="removalError" role="alert" class="removal-error">{{ removalError }}</p>
      <div class="dialog-actions"><button :disabled="busy" autofocus @click="cancelRemoval">取消</button><button class="danger" :disabled="busy" @click="confirmRemoval">{{ busy ? '正在处理…' : pendingRemoval?.id ? '确认删除' : '确认清空' }}</button></div>
    </dialog>
    <div class="toast-stack" aria-live="polite" aria-atomic="false"><div v-for="item in notifications" :key="item.id" :class="['toast',item.type]" :role="item.type === 'error' ? 'alert' : 'status'"><span class="toast-icon" aria-hidden="true">{{ item.type === 'error' ? '!' : '✓' }}</span><span>{{ item.text }}</span><button :aria-label="'关闭提示：'+item.text" @click="dismissNotification(item.id)">×</button></div></div>
  </div>
</template>

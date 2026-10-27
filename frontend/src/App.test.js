// @vitest-environment jsdom
import { mount, flushPromises } from '@vue/test-utils'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import App from './App.vue'

const image = id => ({id, name: `${id}.png`, width: 40, height: 20, size: 100, revision: 0, contexts: []})
let wrapper, state, rejectDelete
const findButton = text => wrapper.findAll('button').find(button => button.text() === text)
async function click(text) { await findButton(text).trigger('click'); await flushPromises() }
async function start() { wrapper = mount(App, {attachTo: document.body}); await flushPromises() }

beforeEach(() => {
  state = [image('a'), image('b')]
  rejectDelete = false
  Object.defineProperty(HTMLDialogElement.prototype, 'showModal', {configurable: true, value() { this.setAttribute('open', '') }})
  Object.defineProperty(HTMLDialogElement.prototype, 'close', {configurable: true, value() { this.removeAttribute('open') }})
  vi.stubGlobal('fetch', vi.fn(async (url, options = {}) => {
    const method = options.method || 'GET'
    const path = url.split('?')[0]
    if (path === '/api/certificates') return {ok:true, json:async () => []}
    if (path === '/api/certificate-jobs' && method === 'GET') return {ok:true, json:async () => null}
    if (method === 'DELETE' && rejectDelete) return {ok: false, json: async () => ({detail: '文件暂时无法删除，请重试'})}
    let body
    if (method === 'DELETE') {
      state = path === '/api/images' ? [] : state.filter(item => path !== `/api/images/${item.id}`)
    }
    if (method === 'POST' && path.endsWith('/edit')) {
      const item = state.find(item => path === `/api/images/${item.id}/edit`)
      item.revision += 1
      ;[item.width, item.height] = [item.height, item.width]
      body = {images: structuredClone(state), selected: item.id}
    }
    return {ok: true, json: async () => body || structuredClone(state)}
  }))
})
afterEach(() => { wrapper?.unmount(); document.body.innerHTML = ''; vi.unstubAllGlobals(); vi.useRealTimers(); vi.restoreAllMocks() })

describe('workspace actions', () => {
  it('switches feature menus without losing the selected image or unfinished edits', async () => {
    await start()
    await wrapper.findAll('.image-item')[1].trigger('click')
    await wrapper.get('#filename').setValue('待保存名称')
    await click('选区裁剪')
    expect(wrapper.get('[role="tabpanel"] input').attributes('accept')).toBe('.docx')
    expect(wrapper.get('aside').find('input[type="file"]').exists()).toBe(false)
    await click('本地图片')
    expect(wrapper.get('[role="tabpanel"] input').attributes('accept')).toBe('image/*')
    expect(wrapper.get('[role="tabpanel"]').text()).toContain('选择图片文件夹')
    await click('证书截图')
    expect(wrapper.get('[role="tabpanel"]').text()).toContain('批量处理证书')
    expect(wrapper.find('.certificate-entry').exists()).toBe(true)
    await click('Word 提图')
    expect(wrapper.findAll('.image-row')).toHaveLength(2)
    expect(wrapper.get('.image-item.active').text()).toContain('b.png')
    expect(wrapper.get('#filename').element.value).toBe('待保存名称')
    expect(wrapper.get('.stage').classes()).toContain('cropping')
    expect(fetch).toHaveBeenCalledTimes(1)
  })

  it('supports keyboard menu navigation without changing the current image', async () => {
    await start()
    await wrapper.get('#menu-word').trigger('keydown', {key:'ArrowRight'})
    expect(wrapper.get('#menu-local').attributes('aria-selected')).toBe('true')
    expect(document.activeElement.id).toBe('menu-local')
    await wrapper.get('#menu-local').trigger('keydown', {key:'End'})
    expect(wrapper.get('#menu-certificate').attributes('aria-selected')).toBe('true')
    await wrapper.get('#menu-certificate').trigger('keydown', {key:'ArrowRight'})
    expect(wrapper.get('#menu-word').attributes('aria-selected')).toBe('true')
    expect(wrapper.get('.image-item.active').text()).toContain('a.png')
  })

  it('opens clear confirmation and cancelling preserves images without sending DELETE', async () => {
    await start()
    await click('清空工作区')
    expect(wrapper.get('dialog').element.open).toBe(true)
    expect(wrapper.get('dialog').text()).toContain('2 张图片')
    await click('取消')
    expect(wrapper.get('dialog').element.open).toBe(false)
    expect(wrapper.findAll('.image-row')).toHaveLength(2)
    expect(fetch.mock.calls.filter(([, options]) => options.method === 'DELETE')).toHaveLength(0)
  })

  it('clears list, selection, name, preview and controls after confirming', async () => {
    await start()
    await wrapper.get('#filename').setValue('未保存名称')
    await click('选区裁剪')
    await click('清空工作区')
    await click('确认清空')
    expect(wrapper.findAll('.image-row')).toHaveLength(0)
    expect(wrapper.find('.picture').exists()).toBe(false)
    expect(wrapper.find('.document-context').exists()).toBe(false)
    expect(wrapper.get('#filename').element.value).toBe('')
    expect(wrapper.get('.image-pagination').text()).toContain('0 / 0')
    expect(findButton('清空工作区').element.disabled).toBe(true)
    expect(findButton('压缩图片').element.matches(':disabled')).toBe(true)
    expect(wrapper.get('dialog').element.open).toBe(false)
    expect(wrapper.get('[role="status"]').text()).toBe('工作区已清空')
    expect(wrapper.get('.toast.success').text()).toContain('工作区已清空')
    expect(findButton('选区裁剪')).toBeTruthy()
  })

  it('retains images and shows a retryable error if clear fails', async () => {
    await start()
    rejectDelete = true
    await click('清空工作区')
    await click('确认清空')
    expect(wrapper.findAll('.image-row')).toHaveLength(2)
    expect(wrapper.get('dialog').element.open).toBe(true)
    expect(wrapper.get('.removal-error').text()).toContain('请重试')
    expect(wrapper.get('.toast.error').text()).toContain('请重试')
    rejectDelete = false
    await click('确认清空')
    expect(wrapper.findAll('.image-row')).toHaveLength(0)
  })

  it('auto-dismisses success toasts and allows manually closing them', async () => {
    await start()
    vi.useFakeTimers({toFake:['setTimeout', 'clearTimeout']})
    await click('↷ 右转 90°')
    expect(wrapper.get('.toast.success').text()).toContain('当前图片已更新')
    await vi.advanceTimersByTimeAsync(3600)
    expect(wrapper.find('.toast').exists()).toBe(false)
    await click('水平翻转')
    await wrapper.get('.toast button').trigger('click')
    expect(wrapper.find('.toast').exists()).toBe(false)
  })

  it('does not show success when folder selection is cancelled', async () => {
    vi.stubGlobal('pywebview', {api:{import_folder:vi.fn(async () => ({cancelled:true}))}})
    await start()
    await click('本地图片')
    await click('选择本地图片文件夹')
    expect(wrapper.find('.toast').exists()).toBe(false)
    expect(wrapper.findAll('.image-row')).toHaveLength(2)
  })

  it('shows the actual imported count after loading files', async () => {
    await start()
    fetch.mockResolvedValueOnce({ok:true, json:async () => [image('a'),image('b'),image('c')]})
    const input = wrapper.get('input[type="file"]')
    Object.defineProperty(input.element, 'files', {value:[new File(['picture'], 'c.png', {type:'image/png'})]})
    await input.trigger('change')
    await flushPromises()
    expect(wrapper.get('.toast.success').text()).toContain('加载完成，已导入 1 张图片')
  })

  it('reports an export failure without showing a success toast', async () => {
    await start()
    fetch.mockResolvedValueOnce({ok:false, json:async () => ({detail:'导出失败，请重试'})})
    await click('↓ 导出全部图片')
    await click('保存为 ZIP 压缩包')
    expect(wrapper.get('.toast.error').text()).toContain('导出失败，请重试')
    expect(wrapper.find('.toast.success').exists()).toBe(false)
  })

  it('shows a native export success only after the bridge resolves', async () => {
    let finish
    vi.stubGlobal('pywebview', {api:{export_zip:vi.fn(() => new Promise(resolve => {finish=resolve}))}})
    await start()
    await click('↓ 导出全部图片')
    await findButton('保存为 ZIP 压缩包').trigger('click')
    expect(wrapper.find('.toast').exists()).toBe(false)
    expect(wrapper.findAll('[role="tab"]').every(tab => tab.element.disabled)).toBe(true)
    finish({count:2})
    await flushPromises()
    expect(wrapper.findAll('[role="tab"]').every(tab => !tab.element.disabled)).toBe(true)
    expect(wrapper.get('.toast.success').text()).toContain('导出完成，已保存 2 张图片')
    expect(wrapper.get('.export-dialog').element.open).toBe(false)
  })

  it('exports to a native directory and closes the format chooser only after success', async () => {
    const exportFolder = vi.fn(async () => ({count:2}))
    vi.stubGlobal('pywebview', {api:{export_folder:exportFolder}})
    await start()
    await click('↓ 导出全部图片')
    expect(wrapper.get('.export-dialog').element.open).toBe(true)
    await click('保存到文件夹')
    expect(exportFolder).toHaveBeenCalledOnce()
    expect(wrapper.get('.export-dialog').element.open).toBe(false)
    expect(wrapper.get('.toast.success').text()).toContain('已保存 2 张图片')
  })

  it('uses the local system chooser when the browser has no directory API', async () => {
    vi.stubGlobal('showDirectoryPicker',undefined)
    await start()
    await click('↓ 导出全部图片')
    fetch.mockResolvedValueOnce({ok:true,json:async () => ({count:2,folder:'系统导出目录'})})
    await click('保存到文件夹')
    expect(fetch.mock.calls.some(([path,options]) => path==='/api/export/folder' && options.method==='POST')).toBe(true)
    expect(wrapper.get('.toast.success').text()).toContain('系统导出目录')
    expect(wrapper.get('.export-dialog').element.open).toBe(false)
  })

  it('falls back to the system chooser when a WebView blocks directory picking', async () => {
    vi.stubGlobal('showDirectoryPicker',vi.fn(async () => {throw new DOMException('blocked','SecurityError')}))
    await start()
    await click('↓ 导出全部图片')
    fetch.mockResolvedValueOnce({ok:true,json:async () => ({cancelled:true})})
    await click('保存到文件夹')
    expect(fetch.mock.calls.some(([path]) => path==='/api/export/folder')).toBe(true)
    expect(wrapper.find('.toast').exists()).toBe(false)
    expect(wrapper.get('.export-dialog').element.open).toBe(true)
  })

  it('writes all original image bytes with current names into the browser-selected directory', async () => {
    const writers = ['a.png','b.png'].map(() => ({write:vi.fn(async () => {}),close:vi.fn(async () => {}),abort:vi.fn()}))
    const getFileHandle = vi.fn(async (name, options) => {
      if (!options?.create) throw new DOMException('missing','NotFoundError')
      return {createWritable:async () => writers[name === 'a.png' ? 0 : 1]}
    })
    const picker = vi.fn(async () => ({name:'测试导出',getFileHandle}))
    vi.stubGlobal('showDirectoryPicker',picker)
    await start()
    const blobs = [new Blob(['a bytes']),new Blob(['b bytes'])]
    blobs.forEach(blob => fetch.mockResolvedValueOnce({ok:true,blob:async () => blob}))
    await click('↓ 导出全部图片')
    await click('保存到文件夹')
    expect(picker).toHaveBeenCalledWith({mode:'readwrite'})
    expect(getFileHandle.mock.calls.filter(([,options]) => options?.create).map(([name]) => name)).toEqual(['a.png','b.png'])
    expect(writers[0].write).toHaveBeenCalledWith(blobs[0])
    expect(writers[1].write).toHaveBeenCalledWith(blobs[1])
    expect(writers.every(writer => writer.close.mock.calls.length === 1)).toBe(true)
    expect(wrapper.get('.toast.success').text()).toContain('2 张图片保存到文件夹“测试导出”')
    expect(wrapper.get('.export-dialog').element.open).toBe(false)
  })

  it('never writes when the selected directory contains a conflicting filename', async () => {
    const getFileHandle = vi.fn(async () => ({}))
    vi.stubGlobal('showDirectoryPicker',vi.fn(async () => ({name:'已有文件',getFileHandle})))
    await start()
    await click('↓ 导出全部图片')
    await click('保存到文件夹')
    expect(getFileHandle).toHaveBeenCalledExactlyOnceWith('a.png')
    expect(fetch).toHaveBeenCalledOnce()
    expect(wrapper.get('.toast.error').text()).toContain('已存在“a.png”')
    expect(wrapper.get('.export-dialog').element.open).toBe(true)
  })

  it('keeps export options open without a success notification when picking a folder is cancelled', async () => {
    vi.stubGlobal('showDirectoryPicker',vi.fn(async () => {throw new DOMException('cancelled','AbortError')}))
    await start()
    await click('↓ 导出全部图片')
    await click('保存到文件夹')
    expect(fetch).toHaveBeenCalledOnce()
    expect(wrapper.find('.toast').exists()).toBe(false)
    expect(wrapper.get('.export-dialog').element.open).toBe(true)
  })

  it('reports a partial folder write without claiming the whole export succeeded', async () => {
    const abort = vi.fn(async () => {})
    const getFileHandle = vi.fn(async (name, options) => {
      if (!options?.create) throw new DOMException('missing','NotFoundError')
      return {createWritable:async () => ({write:async () => {if (name==='b.png') throw new Error('disk full')}, close:async () => {}, abort})}
    })
    vi.stubGlobal('showDirectoryPicker',vi.fn(async () => ({name:'目标文件夹',getFileHandle})))
    await start()
    fetch.mockResolvedValueOnce({ok:true,blob:async () => new Blob(['a'])}).mockResolvedValueOnce({ok:true,blob:async () => new Blob(['b'])})
    await click('↓ 导出全部图片')
    await click('保存到文件夹')
    expect(wrapper.get('.toast.error').text()).toContain('已保存 1 / 2 张图片')
    expect(wrapper.find('.toast.success').exists()).toBe(false)
    expect(abort).toHaveBeenCalledOnce()
    expect(wrapper.findAll('.image-row')).toHaveLength(2)
  })

  it('waits until a browser ZIP file has finished saving before showing success', async () => {
    let finish
    const write=vi.fn(async () => {})
    const close=vi.fn(() => new Promise(resolve => {finish=resolve}))
    vi.stubGlobal('showSaveFilePicker',vi.fn(async () => ({createWritable:async () => ({write,close,abort:vi.fn()})})))
    await start()
    fetch.mockResolvedValueOnce({ok:true,blob:async () => new Blob(['zip'])})
    await click('↓ 导出全部图片')
    await findButton('保存为 ZIP 压缩包').trigger('click'); await flushPromises()
    expect(wrapper.find('.toast').exists()).toBe(false)
    expect(wrapper.get('.export-dialog').element.open).toBe(true)
    finish(); await flushPromises()
    expect(write).toHaveBeenCalledOnce()
    expect(wrapper.get('.toast.success').text()).toContain('ZIP 压缩包已保存')
    expect(wrapper.get('.export-dialog').element.open).toBe(false)
  })

  it('can enter and exit the enlarged image workspace', async () => {
    await start()
    await click('大图模式')
    expect(wrapper.get('.app').classes()).toContain('focus-mode')
    window.dispatchEvent(new KeyboardEvent('keydown', {key:'Escape'}))
    await flushPromises()
    expect(wrapper.get('.app').classes()).not.toContain('focus-mode')
    expect(findButton('大图模式')).toBeTruthy()
  })

  it('uploads several certificates and sends preset and custom fields to the batch API', async () => {
    await start()
    await click('证书截图')
    await click('▣ 证书自动截图')
    expect(findButton('开始截图').element.disabled).toBe(true)
    const certificates = [{id:'c1',filename:'one.crt',name:'Example CA',issuer:'Root CA',public_key:'RSA (2048 bits)'},{id:'c2',filename:'two.crt',name:'Second CA'}]
    fetch.mockResolvedValueOnce({ok:true,json:async () => ({certificates,errors:[{filename:'bad.crt',error:'无法解析证书'}]})})
    const input = wrapper.get('.certificate-upload')
    Object.defineProperty(input.element,'files',{value:[new File(['cert'],'one.crt'),new File(['cert'],'two.crt')]})
    await input.trigger('change'); await flushPromises()
    expect(wrapper.findAll('.certificate-row')).toHaveLength(2)
    expect(wrapper.get('.certificate-info').text()).toContain('Example CA')
    expect(wrapper.get('.certificate-errors').text()).toContain('bad.crt')
    expect(wrapper.findAll('.certificate-options input:checked')).toHaveLength(4)
    await wrapper.get('input[value="serial_number"]').setValue(true)
    await wrapper.get('#custom-certificate-field').setValue('证书策略')
    await click('添加字段')
    expect(wrapper.get('.field-tags').text()).toContain('证书策略')
    const job = {id:'job',status:'completed',total:2,completed:2,count:2,selected:'cert-image',results:[{id:'c1',filename:'one.crt',status:'success',count:1},{id:'c2',filename:'two.crt',status:'success',count:1}]}
    fetch.mockResolvedValueOnce({ok:true,json:async () => job})
    state.push(image('cert-image'),image('cert-second'))
    await click('开始截图（2 张证书）')
    const request = fetch.mock.calls.find(([path,options]) => path === '/api/certificate-jobs' && options.method === 'POST')
    expect(JSON.parse(request[1].body)).toEqual({certificates:['c1','c2'],kinds:['general','issuer','public_key','path','serial_number','field:证书策略']})
    expect(wrapper.findAll('.image-row')).toHaveLength(4)
    expect(wrapper.get('.image-item.active').text()).toContain('cert-image')
    expect(wrapper.get('.certificate-progress').text()).toContain('批量处理结果')
    expect(wrapper.get('.certificate-dialog').element.open).toBe(false)
    expect(wrapper.get('.certificate-report').text()).toContain('成功 2 份证书')
    expect(wrapper.get('.toast-stack').text()).toContain('成功 2 份证书')
  })

  it('retains the queue when file selection is cancelled and removes a custom field', async () => {
    await start()
    await click('证书截图')
    await click('▣ 证书自动截图')
    const input = wrapper.get('.certificate-upload')
    Object.defineProperty(input.element,'files',{value:[]})
    await input.trigger('change'); await flushPromises()
    expect(wrapper.find('.certificate-info').exists()).toBe(false)
    expect(wrapper.find('.toast').exists()).toBe(false)
    expect(findButton('开始截图').element.disabled).toBe(true)
    await wrapper.get('#custom-certificate-field').setValue('证书策略')
    await click('添加字段')
    await wrapper.get('[aria-label="移除字段 证书策略"]').trigger('click')
    expect(wrapper.findAll('.field-tags > span')).toHaveLength(0)
  })

  it('shows a readable error for a plain-text server failure and allows retrying certificate import', async () => {
    await start()
    await click('证书截图')
    await click('▣ 证书自动截图')
    const input = wrapper.get('.certificate-upload')
    Object.defineProperty(input.element,'files',{value:[new File(['cert'],'sm2.crt')]})
    fetch.mockResolvedValueOnce({ok:false,status:500,json:async () => {throw new SyntaxError('Unexpected token I in Internal Server Error')}})
    await input.trigger('change'); await flushPromises()
    expect(wrapper.get('.toast.error').text()).toContain('服务端处理失败（HTTP 500）')
    expect(wrapper.get('.toast.error').text()).not.toContain('Unexpected token')
    expect(input.element.disabled).toBe(false)
    fetch.mockResolvedValueOnce({ok:true,json:async () => ({certificates:[{id:'sm2',filename:'sm2.crt',name:'SM2 Test',public_key:'EC'}],errors:[]})})
    await input.trigger('change'); await flushPromises()
    expect(wrapper.findAll('.certificate-row')).toHaveLength(1)
    expect(wrapper.get('.certificate-info').text()).toContain('SM2 Test')
  })

  it('shows progress, requests stop and retains successful images with per-certificate errors', async () => {
    await start()
    await click('证书截图')
    await click('▣ 证书自动截图')
    fetch.mockResolvedValueOnce({ok:true,json:async () => ({certificates:[{id:'c1',filename:'one.crt',name:'CA'}],errors:[]})})
    const input = wrapper.get('.certificate-upload')
    Object.defineProperty(input.element,'files',{value:[new File(['cert'],'one.crt')]})
    await input.trigger('change'); await flushPromises()
    const job = {id:'job',status:'running',total:1,completed:0,count:0,selected:'',results:[{id:'c1',filename:'one.crt',status:'running',count:0}]}
    fetch.mockResolvedValueOnce({ok:true,json:async () => job})
    vi.useFakeTimers({toFake:['setTimeout','clearTimeout']})
    await findButton('开始截图（1 张证书）').trigger('click'); await flushPromises()
    expect(wrapper.get('.certificate-progress').text()).toContain('0 / 1')
    fetch.mockResolvedValueOnce({ok:true,json:async () => ({...job,status:'cancelling'})})
    await click('停止后续处理')
    expect(findButton('停止后续处理').element.disabled).toBe(true)
    expect(fetch.mock.calls.some(([path]) => path.endsWith('/job/cancel'))).toBe(true)
    fetch.mockResolvedValueOnce({ok:true,json:async () => ({...job,status:'cancelled',completed:1,results:[{id:'c1',filename:'one.crt',status:'failed',count:0,error:'未找到唯一字段，请检查当前证书是否包含此字段',failed_field:'证书策略',failed_kind:'field:证书策略'}]})})
    await vi.advanceTimersByTimeAsync(750); await flushPromises()
    expect(wrapper.get('.batch-result.failed').text()).toContain('未找到唯一字段')
    expect(wrapper.get('.toast.error').text()).toContain('失败 1 份证书')
    expect(wrapper.get('.toast.error').text()).toContain('one.crt：字段“证书策略”截图失败')
    expect(wrapper.get('.certificate-dialog').element.open).toBe(false)
    expect(wrapper.get('.report-failures').text()).toContain('one.crt')
    expect(wrapper.get('.report-failures').text()).toContain('失败字段：证书策略')
    expect(wrapper.get('.report-failures').text()).toContain('未找到唯一字段')
    await wrapper.get('[aria-label="关闭截图结果"]').trigger('click')
    expect(wrapper.find('.certificate-report').exists()).toBe(false)
    expect(wrapper.findAll('.image-row')).toHaveLength(2)
  })

  it('starts with an empty certificate queue every time the screenshot dialog is opened', async () => {
    await start()
    await click('证书截图')
    await click('▣ 证书自动截图')
    fetch.mockResolvedValueOnce({ok:true,json:async () => ({certificates:[{id:'old',filename:'previous.crt',name:'Old CA'}],errors:[{filename:'bad.crt',error:'错误示例'}]})})
    const input = wrapper.get('.certificate-upload')
    Object.defineProperty(input.element,'files',{value:[new File(['cert'],'previous.crt')]})
    await input.trigger('change'); await flushPromises()
    await wrapper.get('#custom-certificate-field').setValue('证书策略')
    await click('添加字段')
    expect(wrapper.findAll('.certificate-row')).toHaveLength(1)
    await click('关闭')
    await click('▣ 证书自动截图')
    expect(fetch.mock.calls.filter(([path,options]) => path==='/api/certificates' && options.method==='DELETE')).toHaveLength(2)
    expect(wrapper.findAll('.certificate-row')).toHaveLength(0)
    expect(wrapper.find('.certificate-info').exists()).toBe(false)
    expect(wrapper.find('.certificate-errors').exists()).toBe(false)
    expect(wrapper.find('.certificate-progress').exists()).toBe(false)
    expect(findButton('开始截图').element.disabled).toBe(true)
    expect(wrapper.get('.field-tags').text()).toContain('证书策略')
    expect(wrapper.findAll('.image-row')).toHaveLength(2)
  })

  it('selects the next image when the current one is deleted, then empties the last selection', async () => {
    await start()
    await wrapper.get('[aria-label="删除 a.png"]').trigger('click')
    await click('确认删除')
    expect(wrapper.findAll('.image-row')).toHaveLength(1)
    expect(wrapper.get('#filename').element.value).toBe('b')
    expect(wrapper.get('.image-item.active').text()).toContain('b.png')
    await wrapper.get('[aria-label="删除 b.png"]').trigger('click')
    await click('确认删除')
    expect(wrapper.find('.picture').exists()).toBe(false)
    expect(wrapper.get('#filename').element.value).toBe('')
  })

  it('preserves selection when another image is deleted and cancellation never deletes', async () => {
    await start()
    await wrapper.get('[aria-label="删除 b.png"]').trigger('click')
    await click('取消')
    expect(wrapper.findAll('.image-row')).toHaveLength(2)
    await wrapper.get('[aria-label="删除 b.png"]').trigger('click')
    await click('确认删除')
    expect(wrapper.get('#filename').element.value).toBe('a')
    expect(wrapper.get('.image-pagination').text()).toContain('1 / 1')
  })

  it('updates the existing image URL after editing without creating a list item', async () => {
    await start()
    const originalUrl = wrapper.get('.picture img').attributes('src')
    await click('↷ 右转 90°')
    expect(wrapper.findAll('.image-row')).toHaveLength(2)
    expect(wrapper.get('.picture img').attributes('src')).not.toBe(originalUrl)
    expect(wrapper.get('.picture img').attributes('src')).toContain('revision=1')
    expect(wrapper.get('#filename').element.value).toBe('a')
    expect(wrapper.get('.image-item.active small').text()).toBe('20 × 40')
    expect(wrapper.get('[role="status"]').text()).toBe('当前图片已更新')
  })
})

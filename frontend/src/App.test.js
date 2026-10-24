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
    expect(wrapper.get('nav').text()).toContain('0 / 0')
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
    await click('↓ 导出全部图片 ZIP')
    expect(wrapper.get('.toast.error').text()).toContain('导出失败，请重试')
    expect(wrapper.find('.toast.success').exists()).toBe(false)
  })

  it('shows a native export success only after the bridge resolves', async () => {
    let finish
    vi.stubGlobal('pywebview', {api:{export_zip:vi.fn(() => new Promise(resolve => {finish=resolve}))}})
    await start()
    await findButton('↓ 导出全部图片 ZIP').trigger('click')
    expect(wrapper.find('.toast').exists()).toBe(false)
    finish({count:2})
    await flushPromises()
    expect(wrapper.get('.toast.success').text()).toContain('导出完成，已保存 2 张图片')
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
    expect(wrapper.get('nav').text()).toContain('1 / 1')
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

// @vitest-environment jsdom
import {mount,flushPromises} from '@vue/test-utils'
import {afterEach,beforeEach,expect,it,vi} from 'vitest'
import AiNamingDialog from './AiNamingDialog.vue'

let wrapper,settings,requests,job
const items=[{id:'a',name:'a.png'},{id:'b',name:'b.jpg'}]
const button=text=>wrapper.findAll('button').find(item=>item.text()===text)
async function click(text){await button(text).trigger('click');await flushPromises()}
async function open(){wrapper=mount(AiNamingDialog,{props:{images:items,selected:'a',token:'test'},attachTo:document.body});await wrapper.vm.open();await flushPromises()}
beforeEach(()=>{
 Object.defineProperty(HTMLDialogElement.prototype,'showModal',{configurable:true,value(){this.setAttribute('open','')}})
 Object.defineProperty(HTMLDialogElement.prototype,'close',{configurable:true,value(){this.removeAttribute('open')}})
 requests=[]
 settings={active:'local',can_remember_key:true,profiles:Object.fromEntries(['local','openai','deepseek','custom'].map(provider=>[provider,{provider,label:provider,base_url:provider==='local'?'http://localhost:11434/v1':'https://example.com/v1',model:provider==='local'?'vision-demo':'',send_image:true,use_context:true,numbered:true,timeout:90,prefix:'',instruction:'简短中文',has_api_key:false,remember_key:false}]))}
 job={id:'j1',status:'completed',completed:2,total:2,results:[{id:'a',original_name:'a.png',revision:0,status:'success',name:'网络拓扑',reason:'描述对应',needs_review:false},{id:'b',original_name:'b.jpg',revision:0,status:'success',name:'可能的设备',reason:'内容不清晰',needs_review:true}]}
 vi.stubGlobal('fetch',vi.fn(async(path,options={})=>{
  requests.push([path,options]);let data
  if(path==='/api/ai/settings'&&!options.method)data=structuredClone(settings)
  if(path==='/api/ai/settings'&&options.method==='PUT'){
   const profile=JSON.parse(options.body);settings.active=profile.provider
   settings.profiles[profile.provider]={...settings.profiles[profile.provider],...profile,has_api_key:!!profile.api_key,api_key:undefined}
   data=structuredClone(settings)
  }
  if(path==='/api/ai/test')data={message:'连接成功'}
  if(path==='/api/ai/jobs'&&!options.method)data=null
  if(path==='/api/ai/jobs'&&options.method==='POST'||path==='/api/ai/jobs/j1')data=structuredClone(job)
  if(path==='/api/ai/jobs/j1/cancel')data={...job,status:'cancelling'}
  if(path==='/api/ai/apply')data={count:1,images:[{id:'a',name:'用户确认.png'},items[1]]}
  return{ok:true,json:async()=>data}
 }))
})
afterEach(()=>{wrapper?.unmount();document.body.innerHTML='';vi.unstubAllGlobals();vi.useRealTimers()})

it('supports provider switching and clears the password field after saving',async()=>{
 await open();await wrapper.get('.ai-provider select').setValue('openai')
 await wrapper.get('input[placeholder="填写此服务中可用的模型名称"]').setValue('vision-api')
 await wrapper.get('input[type=password]').setValue('fake-test-key')
 expect(button('生成命名建议').element.disabled).toBe(true)
 await click('保存配置')
 await click('编辑配置')
 expect(wrapper.get('input[type=password]').element.value).toBe('')
 expect(wrapper.text()).not.toContain('fake-test-key')
 const saved=requests.find(([,options])=>options.method==='PUT')
 expect(JSON.parse(saved[1].body).provider).toBe('openai')
 await wrapper.get('.ai-provider select').setValue('local')
 expect(wrapper.get('input[placeholder="填写此服务中可用的模型名称"]').element.value).toBe('vision-demo')
})

it('previews names, leaves uncertain results unchecked and only applies confirmed edited names',async()=>{
 await open();await wrapper.get('.ai-generate select').setValue('all')
 await click('生成命名建议')
 expect(wrapper.findAll('.ai-result')).toHaveLength(2)
 expect(wrapper.get('[aria-label="选择命名 a.png"]').element.checked).toBe(true)
 expect(wrapper.get('[aria-label="选择命名 b.jpg"]').element.checked).toBe(false)
 expect(wrapper.emitted('updated')).toBeUndefined()
 await wrapper.get('[aria-label="建议名称 a.png"]').setValue('用户确认')
 await click('应用所选名称（1）')
 const apply=requests.find(([path])=>path==='/api/ai/apply')
 expect(JSON.parse(apply[1].body)).toEqual({job_id:'j1',changes:[{id:'a',name:'用户确认'}]})
 expect(wrapper.emitted('updated')[0][0][0].name).toBe('用户确认.png')
 expect(wrapper.get('[aria-label="选择命名 a.png"]').element.disabled).toBe(true)
})

it('tests connection without sending workspace images',async()=>{
 await open();await click('测试连接')
 const request=requests.find(([path])=>path==='/api/ai/test')
 expect(JSON.parse(request[1].body)).toEqual({provider:'local'})
 expect(wrapper.get('.ai-info').text()).toBe('连接成功')
})

it('opens settings without naming controls or recovering a previous job',async()=>{
 wrapper=mount(AiNamingDialog,{props:{images:items,selected:'a'},attachTo:document.body})
 await wrapper.vm.openSettings();await flushPromises()
 expect(wrapper.get('h2').text()).toBe('AI 配置')
 expect(wrapper.find('.ai-generate select').exists()).toBe(false)
 expect(button('生成命名建议')).toBeUndefined()
 expect(requests.some(([path])=>path==='/api/ai/jobs')).toBe(false)
 await click('保存配置')
 expect(wrapper.find('.ai-config').exists()).toBe(true)
 await click('测试连接')
 expect(wrapper.get('.ai-info').text()).toBe('连接成功')
})

it('can stop subsequent naming and releases busy state after polling ends',async()=>{
 await open();job={...job,status:'running',completed:0,results:job.results.map(row=>({...row,status:'pending'}))}
 vi.useFakeTimers({toFake:['setTimeout','clearTimeout']})
 await button('生成命名建议').trigger('click');await flushPromises()
 expect(wrapper.get('[aria-label="关闭 AI 命名"]').element.disabled).toBe(true)
 await click('停止后续处理')
 job={...job,status:'cancelled',results:job.results.map(row=>({...row,status:'cancelled'}))}
 await vi.advanceTimersByTimeAsync(650);await flushPromises()
 expect(wrapper.get('[aria-label="关闭 AI 命名"]').element.disabled).toBe(false)
 expect(wrapper.emitted('busy-change').at(-1)).toEqual([false])
})

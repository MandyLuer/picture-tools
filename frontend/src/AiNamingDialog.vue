<script setup>
import {computed,nextTick,onUnmounted,ref} from 'vue'
const props=defineProps({images:{type:Array,default:()=>[]},selected:{type:String,default:''},token:{type:String,default:''}})
const emit=defineEmits(['busy-change','updated','notice'])
const dialog=ref(null),visible=ref(false),working=ref(false),settings=ref(null),provider=ref('local'),form=ref({}),dirty=ref(false)
const expanded=ref(true),scope=ref('current'),job=ref(null),rows=ref([]),error=ref(''),info=ref('')
const mode=ref('batch')
let disposed=false
const labels={pending:'等待',running:'生成中',success:'已生成',failed:'失败',cancelled:'已停止'}
const configured=computed(()=>!!form.value.model?.trim()&&!!form.value.base_url?.trim()&&(!['openai','deepseek'].includes(provider.value)||form.value.has_api_key||form.value.api_key?.trim()))
const chosen=computed(()=>rows.value.filter(row=>row.status==='success'&&row.checked&&!row.applied))
const currentName=computed(()=>props.images.find(item=>item.id===props.selected)?.name||'尚未选择图片')
function extension(name){return name.match(/\.[^.]+$/)?.[0]||''}
async function api(path,options={}){
 const response=await fetch('/api/ai'+path,{...options,headers:{'x-app-token':props.token,...options.headers}})
 let result
 try{result=await response.json()}catch{throw new Error('模型命名服务返回异常，请检查后端日志后重试')}
 if(!response.ok)throw new Error(typeof result.detail==='string'?result.detail:'配置或请求参数无效，请检查输入')
 return result
}
async function run(action){
 if(working.value)return false
 working.value=true;emit('busy-change',true);error.value='';info.value=''
 try{await action();return true}catch(e){error.value=e.message;return false}finally{working.value=false;emit('busy-change',false)}
}
function loadProvider(value){provider.value=value;form.value={...settings.value.profiles[value],api_key:'',clear_key:false};dirty.value=false}
async function open(value='batch'){
 if(working.value)return
 mode.value=value
 scope.value='all'
 visible.value=true;job.value=null;rows.value=[];expanded.value=true
 await nextTick();dialog.value.showModal()
 await run(async()=>{
  const [preferences,recent]=await Promise.all([api('/settings'),value==='batch'?api('/jobs'):null])
  settings.value=preferences;loadProvider(recent?.provider||settings.value.active);expanded.value=value==='settings'||!configured.value
  if(recent){job.value=recent;await poll()}
 })
}
function close(){if(working.value)return;dialog.value.close();visible.value=false;form.value.api_key='';error.value='';info.value=''}
async function save(){
 await run(async()=>{
  settings.value=await api('/settings',{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({...form.value,provider:provider.value,api_key:form.value.api_key.trim()||null})})
  loadProvider(provider.value);info.value='配置已保存';expanded.value=mode.value==='settings'
 })
}
async function test(){await run(async()=>{const result=await api('/test',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({provider:provider.value})});info.value=result.message})}
async function generate(){
 const ids=scope.value==='all'?props.images.map(item=>item.id):[props.selected].filter(Boolean)
 if(!ids.length){error.value='请先导入并选择图片';return}
 if(ids.length>100){error.value='每批最多命名 100 张图片';return}
 if(dirty.value){error.value='请先保存模型配置';expanded.value=true;return}
 await run(async()=>{
  rows.value=[]
  job.value=await api('/jobs',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({provider:provider.value,images:ids})})
  await poll()
 })
}
async function poll(){
  while(['running','cancelling'].includes(job.value.status)){
   rows.value=job.value.results
   await new Promise(resolve=>setTimeout(resolve,650))
   if(disposed)return
   job.value=await api('/jobs/'+job.value.id)
  }
  rows.value=job.value.results.map(row=>({...row,checked:row.status==='success'&&!row.needs_review&&!row.applied,applied:!!row.applied}))
  info.value=`建议生成结束：成功 ${rows.value.filter(row=>row.status==='success').length} 张，失败 ${rows.value.filter(row=>row.status==='failed').length} 张。确认后点击“应用所选名称”。`
}
async function cancel(){try{job.value=await api('/jobs/'+job.value.id+'/cancel',{method:'POST'})}catch(e){error.value=e.message}}
async function apply(){
 await run(async()=>{
  const selected=chosen.value
  const result=await api('/apply',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({job_id:job.value.id,changes:selected.map(row=>({id:row.id,name:row.name}))})})
  selected.forEach(row=>{row.applied=true;row.checked=false})
  emit('updated',result.images);emit('notice',{text:`AI 命名已应用，更新 ${result.count} 张图片`,type:'success'})
  info.value=`已应用 ${result.count} 个名称，图片内容和扩展名保持不变。`
 })
}
function selectAll(event){rows.value.forEach(row=>{if(row.status==='success'&&!row.applied)row.checked=event.target.checked})}
onUnmounted(()=>{disposed=true})
defineExpose({open,openSettings:()=>open('settings'),openBatch:()=>open('batch')})
</script>

<template>
 <dialog ref="dialog" :class="['ai-dialog',{'ai-settings-dialog':mode==='settings'}]" aria-labelledby="ai-title" @cancel.prevent="close">
  <template v-if="visible">
   <div class="ai-heading"><div><h2 id="ai-title">{{ mode==='settings'?'AI 配置':'批量 AI 命名' }}</h2><p>{{ mode==='settings'?'设置模型、接口与命名规则，保存后在图片旁生成名称':'先生成建议，再确认应用' }}</p></div><button :disabled="working" aria-label="关闭 AI 命名" @click="close">×</button></div>
   <div class="ai-provider"><label>模型配置<select :value="provider" :disabled="working||!settings" @change="loadProvider($event.target.value);expanded=true"><option v-for="(profile,key) in settings?.profiles||{}" :key="key" :value="key">{{ profile.label }}</option></select></label><button v-if="mode==='batch'" :disabled="working" @click="expanded=!expanded">{{ expanded?'收起配置':'编辑配置' }}</button></div>
   <fieldset v-if="expanded&&settings" class="ai-config" :disabled="working" @input="dirty=true" @change="dirty=true">
    <label>接口基础地址<input v-model="form.base_url" placeholder="http://127.0.0.1:11434/v1" autocomplete="off"></label>
    <label>模型名称<input v-model="form.model" placeholder="填写此服务中可用的模型名称"></label>
    <label>API Key<input v-model="form.api_key" type="password" :placeholder="form.has_api_key?'已设置，留空保持不变':provider==='local'?'本地服务无认证时可留空':'填写 API Key'" autocomplete="new-password"></label>
    <div class="ai-checks"><label v-if="settings.can_remember_key"><input v-model="form.remember_key" type="checkbox">记住密钥（Windows 加密保存）</label><span v-else>密钥仅在本次服务运行期间保留</span><label v-if="form.has_api_key"><input v-model="form.clear_key" type="checkbox">清除已保存密钥</label></div>
    <div class="ai-checks"><label><input v-model="form.send_image" type="checkbox">发送图片（需要视觉模型）</label><label><input v-model="form.use_context" type="checkbox">使用文档描述</label><label><input v-model="form.numbered" type="checkbox">名称前保留顺序编号</label></div>
    <div class="ai-inline"><label>名称前缀<input v-model="form.prefix" maxlength="30" placeholder="可选，如：项目A_"></label><label>请求超时（秒）<input v-model.number="form.timeout" type="number" min="10" max="180"></label></div>
    <label>命名要求<textarea v-model="form.instruction" rows="2" maxlength="1000" placeholder="如：使用简短中文，突出设备名称或页面用途"></textarea></label>
    <div class="ai-config-actions"><span>Chat Completions 兼容协议；本地可连接 Ollama、LM Studio 等服务。</span><button class="primary" @click="save">保存配置</button></div>
   </fieldset>
   <div v-else-if="settings" class="ai-model-summary">{{ form.model||'尚未配置模型' }} · {{ form.base_url }}<span v-if="dirty"> · 尚未保存</span></div>
   <p class="ai-note">远程 API 会接收所选图片的缩小副本及相关描述（按配置启用）；本地模式连接你填写的服务。ChatGPT 聊天订阅不包含 API 额度。测试连接仅发送合成测试图片或简短文本。</p>
   <div class="ai-generate"><label v-if="mode==='batch'">处理范围<select v-model="scope" :disabled="working"><option value="current">当前图片：{{ currentName }}</option><option value="all">全部图片（{{ images.length }} 张，最多 100 张）</option></select></label><button :disabled="working||dirty||!configured" @click="test">测试连接</button><button v-if="mode==='batch'" class="primary" :disabled="working||dirty||!configured||!images.length" @click="generate">生成命名建议</button></div>
   <div v-if="job" class="ai-job"><div><b>{{ ['running','cancelling'].includes(job.status)?'正在生成建议':'命名建议' }}</b><span>{{ job.completed }} / {{ job.total }}</span><button v-if="['running','cancelling'].includes(job.status)" :disabled="job.status==='cancelling'" @click="cancel">停止后续处理</button></div><progress :max="job.total" :value="job.completed"></progress></div>
   <p v-if="error" class="ai-error" role="alert">{{ error }}</p><p v-if="info" class="ai-info" role="status">{{ info }}</p>
   <div v-if="rows.length" class="ai-results"><label class="ai-select-all"><input type="checkbox" :disabled="working" @change="selectAll">选择全部可用建议</label><article v-for="row in rows" :key="row.id" :class="['ai-result',row.status]"><input v-model="row.checked" type="checkbox" :disabled="working||row.status!=='success'||row.applied" :aria-label="'选择命名 '+row.original_name"><div><div class="ai-original">{{ row.original_name }} <span>{{ row.applied?'已应用':labels[row.status] }}{{ row.needs_review?' · 建议人工核对':'' }}</span></div><div v-if="row.status==='success'" class="ai-proposed"><input v-model="row.name" :disabled="working||row.applied" :aria-label="'建议名称 '+row.original_name"><span>{{ extension(row.original_name) }}</span></div><p v-if="row.reason">{{ row.reason }}</p><p v-if="row.error" class="ai-error">{{ row.error }}</p></div></article></div>
   <div class="ai-actions"><span>{{ mode==='settings'?'保存配置后，点击图片旁的“AI 生成名称”。':'候选名称可修改，应用后才会重命名。' }}</span><button :disabled="working" @click="close">关闭</button><button v-if="mode==='batch'" class="primary" :disabled="working||!chosen.length" @click="apply">应用所选名称（{{ chosen.length }}）</button></div>
  </template>
 </dialog>
</template>

<style scoped>
.ai-dialog{width:min(1000px,95vw);max-height:92vh;overflow-y:auto;border:1px solid #d9e5de;border-radius:12px;padding:24px;color:#243344;box-shadow:0 16px 60px #182c4340}.ai-dialog::backdrop{background:#1c30465c}.ai-heading{display:flex;justify-content:space-between;gap:12px;position:sticky;top:-24px;background:white;margin:-24px -24px 14px;padding:20px 24px 12px;z-index:3;border-bottom:1px solid #edf0f3}.ai-heading p{font-size:12px;color:#82918a;margin:7px 0 0}.ai-provider,.ai-generate{display:flex;align-items:end;gap:10px;flex-wrap:wrap}.ai-provider>label,.ai-generate>label{flex:1;min-width:180px}.ai-dialog label{font-size:12px;color:#586e61}.ai-dialog select,.ai-dialog textarea{font:inherit;border:1px solid #dce3eb;border-radius:6px;padding:9px;min-width:0;background:white;color:#34465a}.ai-provider select,.ai-generate select,.ai-config>label input,.ai-config textarea{display:block;width:100%;margin-top:6px}.ai-config{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:16px;padding:16px;background:#f7faf8;border:1px solid #e1e9e4;border-radius:8px}.ai-checks,.ai-inline,.ai-config-actions,.ai-config>label:has(textarea){grid-column:1/-1}.ai-checks,.ai-inline{display:flex;align-items:center;gap:14px;flex-wrap:wrap}.ai-checks label{display:flex;align-items:center;gap:6px}.ai-checks span{font-size:11px;color:#83938b}.ai-inline>label{flex:1;min-width:140px}.ai-inline input{width:100%;display:block;margin-top:6px}.ai-config-actions{display:flex;align-items:center;justify-content:space-between;gap:12px}.ai-config-actions span{font-size:11px;color:#819188;line-height:1.6}.ai-model-summary{margin-top:12px;background:#f1f7f3;padding:10px;border-radius:6px;font-size:12px;overflow-wrap:anywhere}.ai-note{font-size:11px;line-height:1.8;color:#8b7c64}.ai-job{margin:14px 0}.ai-job>div{display:flex;gap:12px;align-items:center;font-size:12px}.ai-job progress{width:100%;height:8px;accent-color:#126c62}.ai-job button{margin-left:auto;padding:5px 9px;font-size:12px}.ai-error{font-size:12px;color:#b24b3a;line-height:1.7;overflow-wrap:anywhere}.ai-info{font-size:12px;color:#2c7c5a;line-height:1.7}.ai-results{border:1px solid #e0e8e3;border-radius:8px;max-height:340px;overflow-y:auto;margin-top:12px;overscroll-behavior:contain}.ai-select-all{display:flex;align-items:center;gap:7px;padding:10px;border-bottom:1px solid #edf0f3}.ai-result{display:flex;gap:10px;padding:12px;border-top:1px solid #edf0f3}.ai-result>input{align-self:flex-start;margin-top:5px}.ai-result>div{flex:1;min-width:0}.ai-original{display:flex;justify-content:space-between;gap:12px;font-size:12px;overflow-wrap:anywhere}.ai-original span{font-size:11px;color:#87978c;flex-shrink:0}.ai-proposed{display:flex;gap:8px;align-items:center;margin-top:7px}.ai-proposed input{flex:1;min-width:0}.ai-proposed span{font-size:12px;color:#809188}.ai-result p{font-size:11px;line-height:1.6;color:#7e8f84;margin:6px 0 0}.ai-result .ai-error{color:#b24b3a}.ai-actions{display:flex;justify-content:flex-end;align-items:center;gap:10px;position:sticky;bottom:-24px;z-index:2;background:white;padding:14px 24px;margin:14px -24px -24px;border-top:1px solid #edf0f3}.ai-actions>span{margin-right:auto;font-size:11px;color:#87978c}.ai-dialog input[type=checkbox]{accent-color:#126c62}.ai-dialog button{font-size:12px}@media(max-width:650px){.ai-config{grid-template-columns:1fr}.ai-actions{flex-wrap:wrap}.ai-actions>span{width:100%}}
</style>

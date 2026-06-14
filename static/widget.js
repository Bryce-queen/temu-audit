(() => {
if (document.getElementById('ai-chat-widget')) return;
const d = document;
const API = new URL(document.currentScript?.src || window.location.href).origin;
let sid = 'widget_' + Date.now();
let model = 'cloud:Qwen/Qwen3-8B';
let isOpen = false, isStreaming = false;

const css = d.createElement('style');
css.id = 'ai-chat-widget-css';
css.textContent = `
.ai-cw-btn{position:fixed;bottom:20px;right:20px;z-index:9999;width:52px;height:52px;border-radius:50%;background:#2997ec;border:none;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,0.3);display:flex;align-items:center;justify-content:center;transition:transform 0.2s}
.ai-cw-btn:hover{transform:scale(1.08)}.ai-cw-btn svg{width:24px;height:24px;fill:#fff}
.ai-cw-panel{position:fixed;bottom:84px;right:20px;z-index:9998;width:380px;height:520px;background:#1a1a2e;border-radius:14px;box-shadow:0 8px 40px rgba(0,0,0,0.4);display:none;flex-direction:column;overflow:hidden;border:1px solid #2a2a4a;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif}
.ai-cw-panel.open{display:flex}
.ai-cw-header{background:#16213e;padding:12px 16px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #2a2a4a}
.ai-cw-title{color:#e0e0e0;font-size:14px;font-weight:600;display:flex;align-items:center;gap:8px}
.ai-cw-title-dot{width:8px;height:8px;background:#4ade80;border-radius:50%}
.ai-cw-close{background:none;border:none;color:#888;cursor:pointer;font-size:20px;line-height:1}
.ai-cw-msgs{flex:1;overflow-y:auto;padding:12px;display:flex;flex-direction:column;gap:10px}
.ai-cw-msg{max-width:85%;padding:10px 14px;border-radius:12px;font-size:13px;line-height:1.5;white-space:pre-wrap;word-break:break-word}
.ai-cw-msg.user{align-self:flex-end;background:#2997ec;color:#fff;border-bottom-right-radius:4px}
.ai-cw-msg.bot{align-self:flex-start;background:#2a2a4a;color:#d0d0d0;border-bottom-left-radius:4px}
.ai-cw-input{border-top:1px solid #2a2a4a;padding:8px}
.ai-cw-input select{width:100%;padding:6px 8px;margin-bottom:6px;border-radius:6px;border:1px solid #3a3a5a;background:#1e1e3a;color:#c0c0c0;font-size:12px;cursor:pointer;outline:none}
.ai-cw-input-row{display:flex;gap:8px}
.ai-cw-input-row input{flex:1;padding:10px 12px;border-radius:8px;border:1px solid #3a3a5a;background:#12122a;color:#e0e0e0;font-size:13px;outline:none}
.ai-cw-input-row input:focus{border-color:#2997ec}
.ai-cw-send{background:#2997ec;border:none;border-radius:8px;padding:0 14px;color:#fff;cursor:pointer;font-size:13px;font-weight:600}
.ai-cw-send:disabled{opacity:0.5;cursor:default}
.ai-cw-typing{display:flex;gap:4px;padding:10px 14px}
.ai-cw-typing span{width:6px;height:6px;background:#888;border-radius:50%;animation:ai-cw-bounce 1.2s infinite}
.ai-cw-typing span:nth-child(2){animation-delay:0.2s}
.ai-cw-typing span:nth-child(3){animation-delay:0.4s}
@keyframes ai-cw-bounce{0%,60%,100%{transform:translateY(0)}30%{transform:translateY(-6px)}}
`;
d.head.appendChild(css);

const btn = d.createElement('button');
btn.className = 'ai-cw-btn';
btn.id = 'ai-chat-widget';
btn.innerHTML = '<svg viewBox="0 0 24 24"><path d="M20 2H4c-1.1 0-2 .9-2 2v18l4-4h14c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm0 14H5.17L4 17.17V4h16v12z"/><path d="M11 12h2v2h-2zm0-6h2v4h-2z"/></svg>';

const panel = d.createElement('div');
panel.className = 'ai-cw-panel';
panel.id = 'ai-cw-panel';
panel.innerHTML = `
<div class="ai-cw-header">
  <div class="ai-cw-title"><span class="ai-cw-title-dot"></span>AI 客服</div>
  <button class="ai-cw-close" id="ai-cw-close">&times;</button>
</div>
<div class="ai-cw-msgs" id="ai-cw-msgs">
  <div class="ai-cw-msg bot">你好，有什么可以帮你的？</div>
</div>
<div class="ai-cw-input">
  <select id="ai-cw-model">
    <option value="cloud:Qwen/Qwen3-8B">Qwen3-8B (免费)</option>
    <option value="cloud:Qwen/Qwen2.5-7B-Instruct">Qwen2.5-7B (免费)</option>
    <option value="cloud:glm-4.5-air">GLM-4.5-Air (智谱)</option>
  </select>
  <div class="ai-cw-input-row">
    <input id="ai-cw-inp" placeholder="输入消息...">
    <button class="ai-cw-send" id="ai-cw-send">发送</button>
  </div>
</div>`;

d.body.appendChild(btn);
d.body.appendChild(panel);

const msgs = d.getElementById('ai-cw-msgs');
const inp = d.getElementById('ai-cw-inp');
const sendBtn = d.getElementById('ai-cw-send');
const modelSel = d.getElementById('ai-cw-model');

function toggle() {
  isOpen = !isOpen;
  panel.classList.toggle('open', isOpen);
  btn.style.display = isOpen ? 'none' : 'flex';
  if (isOpen) inp.focus();
}
btn.onclick = toggle;
d.getElementById('ai-cw-close').onclick = toggle;

function addMsg(role, text) {
  const div = d.createElement('div');
  div.className = 'ai-cw-msg ' + role;
  div.textContent = text;
  msgs.appendChild(div);
  msgs.scrollTop = msgs.scrollHeight;
  return div;
}

function typing() {
  const div = d.createElement('div');
  div.className = 'ai-cw-typing';
  div.id = 'ai-cw-typing';
  div.innerHTML = '<span></span><span></span><span></span>';
  msgs.appendChild(div);
  msgs.scrollTop = msgs.scrollHeight;
  return div;
}

function rmTyping() {
  const t = d.getElementById('ai-cw-typing');
  if (t) t.remove();
}

async function doSend() {
  const text = inp.value.trim();
  if (!text || isStreaming) return;
  inp.value = '';
  addMsg('user', text);
  const bdiv = addMsg('bot', '');
  const tdiv = typing();
  isStreaming = true;
  sendBtn.disabled = true;
  model = modelSel.value;

  try {
    const r = await fetch(API + '/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ model: model, prompt: text, session_id: sid, stream: true })
    });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    rmTyping();
    const reader = r.body.getReader();
    const dec = new TextDecoder();
    let buf = '', content = '';
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += dec.decode(value, { stream: true });
      const lines = buf.split('\n');
      buf = lines.pop() || '';
      for (const line of lines) {
        if (!line.startsWith('data:')) continue;
        try {
          const j = JSON.parse(line.slice(5).trim());
          if (j.token) {
            content += j.token;
            bdiv.textContent = content;
            msgs.scrollTop = msgs.scrollHeight;
          }
        } catch (e) {}
      }
    }
    if (!content) bdiv.textContent = '[无响应]';
  } catch (e) {
    rmTyping();
    bdiv.textContent = '错误: ' + e.message;
  }
  isStreaming = false;
  sendBtn.disabled = false;
  inp.focus();
}

sendBtn.onclick = doSend;
inp.onkeydown = function (e) {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    doSend();
  }
};
})();

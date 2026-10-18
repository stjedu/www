
const $=id=>document.getElementById(id);const boxes=[...document.querySelectorAll('.done')];const key='ai-video-workshop-v1';
try{const saved=JSON.parse(localStorage.getItem(key)||'[]');boxes.forEach(b=>b.checked=Array.isArray(saved)&&saved.includes(b.dataset.step));}catch{}
function progress(){const n=boxes.filter(b=>b.checked).length;$('progress').style.width=(n/boxes.length*100)+'%';$('progress-text').textContent=`已完成 ${n} / ${boxes.length} 個檢查點`;try{localStorage.setItem(key,JSON.stringify(boxes.filter(b=>b.checked).map(b=>b.dataset.step)));}catch{}}
boxes.forEach(b=>b.addEventListener('change',progress));progress();$('reset').addEventListener('click',()=>{boxes.forEach(b=>b.checked=false);progress();$('status').textContent='檢查進度已重設。';});
function generate(){const topic=$('topic').value.trim()||'我的教學主題';const audience=$('audience').value.trim()||'初學者';$('generated').textContent=`我要把一段「${topic}」錄影做成教學影片，對象是${audience}。

1. 先檢查原片規格、音軌與可讀取內容，依操作變化建立時間碼腳本；不能確認的內容請標示。
2. 每個重要步驟說清楚點擊位置、欄位、數值、操作理由與確認方法。不能漏掉：${$('must').value.trim()||'主要操作與結果檢查'}
3. 旁白採${$('voice').value}。先提供參數設定與接線等代表片段試聽，經我確認再製作全片。
4. 說明超過原片時間時可延長定格，不加速語音；字幕與章節同步更新。
5. 輸出 MP4、繁體中文 SRT 與講解稿，提供重點預覽。保留原片與先前版本。
6. 完成後採「${$('privacy').value}」。若要上傳，先核對我指定的頻道，再操作；登入、驗證、必要權限或條款確認由我接手。
7. 發布後開啟觀看頁，核對新版標題、頻道、瀏覽權限與播放狀態，回報實際連結及尚未完成的處理或檢查。

請先分析與提供腳本，不要跳過預覽和教師確認。`;$('download-prompt').href="data:text/plain;charset=utf-8,"+encodeURIComponent("\ufeff"+$('generated').textContent);}
['topic','audience','voice','privacy','must'].forEach(id=>$(id).addEventListener('input',generate));generate();
document.querySelectorAll('[data-copy]').forEach(b=>b.addEventListener('click',async()=>{try{await navigator.clipboard.writeText($(b.dataset.copy).textContent);$('status').textContent='已複製，可貼到 AI 對話中。';b.textContent='已複製 ✓';}catch{$('status').textContent='瀏覽器未允許複製，請選取指令文字後手動複製。';}}));
function timing(){const a=Number($('screen-sec').value),b=Number($('speech-sec').value);if(!Number.isFinite(a)||!Number.isFinite(b)||a<1||b<1||a>600||b>600){$('timing-result').textContent='請輸入 1–600 秒。';return;}const hold=Math.max(0,b+.5-a);$('timing-result').textContent=hold?`建議延長約 ${hold.toFixed(1)} 秒 → 本段約 ${(a+hold).toFixed(1)} 秒`:`原片時間足夠，保留約 ${(a-b).toFixed(1)} 秒停頓。`;}
['screen-sec','speech-sec'].forEach(id=>$(id).addEventListener('input',timing));timing();$('print').addEventListener('click',()=>window.print());

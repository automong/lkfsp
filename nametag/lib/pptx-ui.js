/* Local-only PPT import/export controls. Uses the app's existing state/store. */
let pptBusy=false;
function setPptBusy(busy){ pptBusy=busy; $('pptPick').disabled=busy; $('pptFile').disabled=busy; $('pptDownload').disabled=busy; }
async function importPpt(file){
  if(!file||pptBusy)return;
  const seq=renderSeq; setPptBusy(true); $('pptStatus').textContent='PPT 확인 중…';delete $('pptStatus').dataset.result;
  try{
    const result=await NametagPpt.inspect(file,msg=>$('pptStatus').textContent=msg);
    if(seq!==renderSeq)throw Error('확인 중에 양식이 바뀌었습니다. PPT를 다시 선택해 주세요.');
    const d=state.tpl.design, messages=[];
    if(result.back){d.back=result.back.data;d.backOn=true;d.backPrinted=true;messages.push('일정표를 넣었습니다.');}
    if(result.strip){d.strip=result.strip.data;d.stripOn=true;d.stripRatio=result.strip.width/result.strip.height;d.stripH=Math.max(5,Math.min(80,result.strip.height));messages.push('띠를 넣었습니다.');}
    for(const [key,label] of [['strip','띠'],['back','일정표']])if(!result[key]){
      const reason=result.ambiguous.includes(key)?`${label}가 여러 종류입니다.`:result.unsupported.includes(key)?`${label} 그림을 읽지 못했습니다.`:`${label}를 찾지 못했습니다.`;
      messages.push(reason+' 아래에서 파일을 선택하세요.');
    }
    fillSettingsUI();render();$('pptStatus').textContent=messages.join(' ');$('pptStatus').dataset.result='ready';
  }catch(error){$('pptStatus').textContent=error.message+' 기존 그림은 유지됩니다.';$('pptStatus').dataset.result='error';}
  finally{setPptBusy(false);$('pptFile').value='';}
}
$('pptFile').onchange=e=>importPpt(e.target.files[0]);
for(const type of ['dragover','drop'])$('pptDrop').addEventListener(type,e=>{e.preventDefault();e.stopPropagation();if(type==='drop')importPpt(e.dataTransfer.files[0]);});
async function downloadPpt(){
  if(pptBusy)return;
  if(!state.map['이름']||!printableRows().length){goStep(1);toast('먼저 이름 열이 있는 명단을 넣어 주세요');return;}
  if(state.headerIndex<0&&!state.headerConfirmed){
    if(!confirm('첫 줄부터 명단으로 저장합니다. 이름·기수·직책 열이 맞습니까?')){goStep(1);return;}
    state.headerConfirmed=true;persist();
  }
  const seq=renderSeq,t=clone(state.tpl),d=t.design;
  // The button exports the staff's original layout, not a photo/custom template.
  setPptBusy(true);$('pptDownloadStatus').textContent='PPT 만드는 중…';delete $('pptDownloadStatus').dataset.result;
  try{
    await document.fonts.ready;
    const rows=printableRows().map((row,i)=>{
      const role=val(row,'직책').trim(),cohort=val(row,'기수').trim();
      const cohortText=/^\d+\s*기?$/.test(cohort)?cohort.replace(/\s*기$/,'')+'기':cohort;
      const title=cohort&&!/^\d+\s*기?$/.test(cohort)?role||cohort:role&&cohortText?role+' ('+cohortText+')':role||cohortText;
      const raw=val(row,'이름').replace(/\s+/g,''),name=Array.from(raw).join(' '),avail=102-PPT_COLUMNS[i%2].textX-7.3;
      return {role:title,name,roleSize:fitOriginalSize(title,21,0,avail),nameSize:fitOriginalSize(name,65*Math.min(1,3/Array.from(raw).length),0,avail)};
    });
    const result=await NametagPpt.build({rows,logo:DC_LOGO,back:d.backOn?d.back:null,backPrinted:d.backPrinted,strip:d.stripOn?d.strip:null,stripHeight:stackH(t)||d.stripH,stripCount:d.stripWhere==='stack'?0:d.stripCount,isOriginal:isOriginalSchedule},msg=>$('pptDownloadStatus').textContent=msg);
    if(seq!==renderSeq)throw Error('만드는 중에 내용이 바뀌었습니다. 다시 눌러 주세요.');
    const a=document.createElement('a'),url=URL.createObjectURL(result.blob);a.href=url;a.download='명찰_'+rows.length+'명.pptx';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);
    $('pptDownloadStatus').textContent=`${rows.length}명 · 명찰 ${result.badgeSlides}장`+(result.stripSlides?` + 띠 ${result.stripSlides}장`:'')+' 저장했습니다.';
    $('pptDownloadStatus').dataset.result='ready';
  }catch(error){$('pptDownloadStatus').textContent=error.message;$('pptDownloadStatus').dataset.result='error';}
  finally{setPptBusy(false);}
}

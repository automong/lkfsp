/* Local-only PPTX reader/writer. No imported XML is inserted into the page.
 * ZIP: bounded, lazy inflation, CRC checks. OOXML: embedded raster images only.
 * Geometry and export template: staff PPT, slide 1. See tools/README.md. */
(function () {
  'use strict';
  const A='http://schemas.openxmlformats.org/drawingml/2006/main';
  const P='http://schemas.openxmlformats.org/presentationml/2006/main';
  const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PKG='http://schemas.openxmlformats.org/package/2006/relationships';
  const MAX_FILE=64*1024*1024, MAX_PART=32*1024*1024, MAX_TOTAL=192*1024*1024;
  const enc=new TextEncoder(), dec=new TextDecoder('utf-8',{fatal:true});
  const tick=()=>new Promise(r=>setTimeout(r,0));
  const all=(node,ns,name)=>Array.from(node.getElementsByTagNameNS(ns,name));
  const one=(node,ns,name)=>all(node,ns,name)[0];
  const children=(node,ns,name)=>Array.from(node.children).filter(e=>e.namespaceURI===ns&&e.localName===name);
  const xml=s=>{
    if(s.length>4*1024*1024||/<!DOCTYPE|<!ENTITY/i.test(s))throw Error('지원하지 않는 PPT 문서입니다.');
    const d=new DOMParser().parseFromString(s,'application/xml');
    if(d.getElementsByTagName('parsererror').length)throw Error('PPT 문서가 손상되었습니다.');
    return d;
  };
  const serialize=d=>new XMLSerializer().serializeToString(d);
  const escape=s=>String(s).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g,'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/"/g,'&quot;');
  const crcTable=Uint32Array.from({length:256},(_,n)=>{for(let k=0;k<8;k++)n=n&1?0xedb88320^(n>>>1):n>>>1;return n>>>0;});
  function crc32(b){let c=0xffffffff;for(const x of b)c=crcTable[(c^x)&255]^(c>>>8);return (c^0xffffffff)>>>0;}
  function resolvePart(base,target){
    if(!target||/[\\?#\x00]/.test(target)||/^[a-z]+:/i.test(target))return null;
    const parts=target.startsWith('/')?[]:base.split('/').slice(0,-1);
    for(const p of target.split('/')){if(!p||p==='.')continue;if(p==='..'){if(!parts.length)return null;parts.pop();}else parts.push(p);}
    return parts.join('/');
  }
  function zipReader(buffer){
    const b=new Uint8Array(buffer),v=new DataView(b.buffer,b.byteOffset,b.byteLength),entries=new Map(),cache=new Map();
    if(b.length>MAX_FILE||b.length<22)throw Error('PPT 파일은 64MB 이하로 선택해 주세요.');
    let end=-1;
    for(let i=b.length-22;i>=Math.max(0,b.length-65557);i--){if(v.getUint32(i,true)===0x06054b50&&i+22+v.getUint16(i+20,true)===b.length){end=i;break;}}
    if(end<0)throw Error('PPTX 파일을 선택해 주세요. (.ppt 파일은 PPTX로 저장)');
    const count=v.getUint16(end+10,true),cd=v.getUint32(end+16,true),cdSize=v.getUint32(end+12,true);
    if(v.getUint16(end+4,true)||v.getUint16(end+6,true)||count!==v.getUint16(end+8,true)||count>6000||cd+cdSize!==end)throw Error('지원하지 않는 PPT 압축 형식입니다.');
    let pos=cd,total=0;
    for(let i=0;i<count;i++){
      if(pos+46>end||v.getUint32(pos,true)!==0x02014b50)throw Error('PPT 압축 정보가 손상되었습니다.');
      const flags=v.getUint16(pos+8,true),method=v.getUint16(pos+10,true),crc=v.getUint32(pos+16,true),compressed=v.getUint32(pos+20,true),size=v.getUint32(pos+24,true);
      const nl=v.getUint16(pos+28,true),extra=v.getUint16(pos+30,true),comment=v.getUint16(pos+32,true),offset=v.getUint32(pos+42,true);
      if(pos+46+nl+extra+comment>end)throw Error('PPT 압축 정보가 손상되었습니다.');
      const name=dec.decode(b.subarray(pos+46,pos+46+nl));total+=size;
      if(flags&1||![0,8].includes(method)||size>MAX_PART||total>MAX_TOTAL||entries.has(name)||resolvePart('',name)!==name||name.startsWith('/')||offset+30>cd)throw Error('읽을 수 없는 PPT 파일입니다. 그림 파일을 직접 선택해 주세요.');
      if(v.getUint32(offset,true)!==0x04034b50||v.getUint16(offset+8,true)!==method||v.getUint16(offset+6,true)!==flags)throw Error('PPT 압축 정보가 손상되었습니다.');
      const localN=v.getUint16(offset+26,true),localX=v.getUint16(offset+28,true),start=offset+30+localN+localX;
      if(start+compressed>cd||dec.decode(b.subarray(offset+30,offset+30+localN))!==name)throw Error('PPT 압축 정보가 손상되었습니다.');
      entries.set(name,{start,compressed,size,method,crc});pos+=46+nl+extra+comment;
    }
    if(pos!==end)throw Error('PPT 압축 정보가 손상되었습니다.');
    async function get(name){
      if(cache.has(name))return cache.get(name);
      const e=entries.get(name);if(!e)throw Error('PPT 안의 그림이나 문서를 찾지 못했습니다.');
      let result;
      if(e.method===0)result=b.slice(e.start,e.start+e.compressed);
      else {
        let stream;try{stream=new DecompressionStream('deflate-raw');}catch{throw Error('최신 Chrome·Edge·Safari에서 다시 열어 주세요.');}
        const reader=new Blob([b.subarray(e.start,e.start+e.compressed)]).stream().pipeThrough(stream).getReader(),chunks=[];let n=0;
        try{while(true){const {done,value}=await reader.read();if(done)break;n+=value.length;if(n>e.size||n>MAX_PART)throw Error('PPT 압축 크기가 올바르지 않습니다.');chunks.push(value);}}
        catch(error){await reader.cancel().catch(()=>{});throw error;}
        result=new Uint8Array(n);let at=0;for(const chunk of chunks){result.set(chunk,at);at+=chunk.length;}
      }
      if(result.length!==e.size||crc32(result)!==e.crc)throw Error('PPT 파일이 손상되었습니다. 다시 저장한 파일을 선택해 주세요.');
      cache.set(name,result);return result;
    }
    return {entries,get,text:async name=>dec.decode(await get(name))};
  }
  function zipWrite(parts){
    const chunks=[],central=[];let offset=0,count=0;
    for(const [name,value] of Object.entries(parts)){
      const n=enc.encode(name),b=typeof value==='string'?enc.encode(value):value,crc=crc32(b);
      const h=new Uint8Array(30+n.length),v=new DataView(h.buffer);
      v.setUint32(0,0x04034b50,true);v.setUint16(4,20,true);v.setUint16(6,0x800,true);v.setUint32(14,crc,true);v.setUint32(18,b.length,true);v.setUint32(22,b.length,true);v.setUint16(26,n.length,true);h.set(n,30);
      const c=new Uint8Array(46+n.length),cv=new DataView(c.buffer);
      cv.setUint32(0,0x02014b50,true);cv.setUint16(4,20,true);cv.setUint16(6,20,true);cv.setUint16(8,0x800,true);cv.setUint32(16,crc,true);cv.setUint32(20,b.length,true);cv.setUint32(24,b.length,true);cv.setUint16(28,n.length,true);cv.setUint32(42,offset,true);c.set(n,46);
      chunks.push(h,b);central.push(c);offset+=h.length+b.length;count++;
    }
    const end=new Uint8Array(22),v=new DataView(end.buffer);v.setUint32(0,0x06054b50,true);v.setUint16(8,count,true);v.setUint16(10,count,true);v.setUint32(12,central.reduce((s,c)=>s+c.length,0),true);v.setUint32(16,offset,true);
    return new Blob([...chunks,...central,end],{type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'});
  }
  async function relationships(z,part){
    const path=part.slice(0,part.lastIndexOf('/')+1)+'_rels/'+part.split('/').pop()+'.rels';
    const out=new Map();if(!z.entries.has(path))return out;
    for(const e of all(xml(await z.text(path)),PKG,'Relationship')){
      if(e.getAttribute('TargetMode')==='External')continue;
      const target=resolvePart(part,e.getAttribute('Target'));if(target)out.set(e.getAttribute('Id'),{target,type:e.getAttribute('Type')});
    }return out;
  }
  const attr=(e,k,d=0)=>e&&e.hasAttribute(k)?Number(e.getAttribute(k)):d;
  const mul=(a,b)=>[a[0]*b[0]+a[2]*b[1],a[1]*b[0]+a[3]*b[1],a[0]*b[2]+a[2]*b[3],a[1]*b[2]+a[3]*b[3],a[0]*b[4]+a[2]*b[5]+a[4],a[1]*b[4]+a[3]*b[5]+a[5]];
  const transform=(m,x,y)=>[m[0]*x+m[2]*y+m[4],m[1]*x+m[3]*y+m[5]];
  function xform(e,group=false){
    const off=one(e,A,'off'),ext=one(e,A,'ext');if(!off||!ext)return null;
    const x=attr(off,'x'),y=attr(off,'y'),w=attr(ext,'cx'),h=attr(ext,'cy'),rad=attr(e,'rot')/60000*Math.PI/180;
    if(!(w>0&&h>0))return null;
    const fx=e.getAttribute('flipH')==='1'||e.getAttribute('flipH')==='true'?-1:1,fy=e.getAttribute('flipV')==='1'||e.getAttribute('flipV')==='true'?-1:1;
    let m=mul([1,0,0,1,x+w/2,y+h/2],[Math.cos(rad)*fx,Math.sin(rad)*fx,-Math.sin(rad)*fy,Math.cos(rad)*fy,0,0]);
    if(group){const co=one(e,A,'chOff'),ce=one(e,A,'chExt'),cw=attr(ce,'cx'),ch=attr(ce,'cy');if(!(cw>0&&ch>0))return null;m=mul(m,[w/cw,0,0,h/ch,-w/2-attr(co,'x')*w/cw,-h/2-attr(co,'y')*h/ch]);}
    else m=mul(m,[w,0,0,h,-w/2,-h/2]);
    return m;
  }
  function pictures(tree,rels,parent=[1,0,0,1,0,0],out=[]){
    for(const e of tree.children){
      if(e.localName==='grpSp'){
        const pr=children(e,P,'grpSpPr')[0],xf=pr&&one(pr,A,'xfrm'),m=xf&&xform(xf,true);if(m)pictures(e,rels,mul(parent,m),out);continue;
      }
      if(!['pic','sp'].includes(e.localName))continue;
      const blip=one(e,A,'blip'),pr=children(e,P,'spPr')[0],xf=pr&&one(pr,A,'xfrm');
      if(!blip||!xf)continue;
      const rel=rels.get(blip.getAttributeNS(R,'embed'));if(!rel||!rel.type.endsWith('/image'))continue;
      const m=xform(xf);if(!m)continue;const matrix=mul(parent,m),points=[[0,0],[1,0],[1,1],[0,1]].map(([x,y])=>transform(matrix,x,y));
      const xs=points.map(p=>p[0]),ys=points.map(p=>p[1]),x=Math.min(...xs),y=Math.min(...ys),w=Math.max(...xs)-x,h=Math.max(...ys)-y;
      if(![x,y,w,h,...matrix].every(Number.isFinite))continue;
      const crop=one(e,A,'srcRect');
      out.push({part:rel.target,x:x/36000,y:y/36000,w:w/36000,h:h/36000,matrix,crop:['l','t','r','b'].map(k=>attr(crop,k)/100000)});
    }return out;
  }
  async function loadImage(url){return new Promise((resolve,reject)=>{const im=new Image();im.onload=()=>resolve(im);im.onerror=()=>reject(Error('PPT 그림을 읽지 못했습니다. 아래에서 그림 파일을 선택해 주세요.'));im.src=url;});}
  function imageHeader(bytes){
    const invalid=()=>{throw Error('PPT 그림의 크기 정보를 읽지 못했습니다. PNG·JPG 파일로 선택해 주세요.');};
    const dimensions=(type,width,height)=>{
      if(!width||!height)invalid();
      if(width>40e6/height)throw Error('그림이 너무 큽니다. 줄인 그림 파일을 선택해 주세요.');
      return {type,width,height};
    };
    const v=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);
    const signature=[137,80,78,71,13,10,26,10];
    if(bytes.length>=8&&signature.every((n,i)=>bytes[i]===n)){
      if(bytes.length<33||v.getUint32(8)!==13||v.getUint32(12)!==0x49484452)invalid();
      return dimensions('image/png',v.getUint32(16),v.getUint32(20));
    }
    if(bytes[0]!==0xff||bytes[1]!==0xd8)invalid();
    // Read marker lengths, never scan arbitrary payload bytes for a fake SOF.
    let p=2;
    while(p<bytes.length){
      if(bytes[p++]!==0xff)invalid();
      while(bytes[p]===0xff)p++;
      if(p>=bytes.length)invalid();
      const marker=bytes[p++];
      if(marker===0xda||marker===0xd9||marker===0x00||marker===0xd8)invalid();
      if(marker===0x01||(marker>=0xd0&&marker<=0xd7))continue;
      if(p+2>bytes.length)invalid();
      const length=v.getUint16(p);
      if(length<2||p+length>bytes.length)invalid();
      if(marker>=0xc0&&marker<=0xcf&&![0xc4,0xc8,0xcc].includes(marker)){
        if(length<11||bytes[p+7]===0||length!==8+3*bytes[p+7])invalid();
        return dimensions('image/jpeg',v.getUint16(p+5),v.getUint16(p+3));
      }
      p+=length;
    }
    invalid();
  }
  async function decodeImage(bytes){
    // The header limit MUST run before constructing an Image or a blob URL.
    const header=imageHeader(bytes);
    const url=URL.createObjectURL(new Blob([bytes],{type:header.type}));let im;
    try{im=await loadImage(url);}finally{URL.revokeObjectURL(url);}
    if(!im.naturalWidth||!im.naturalHeight||im.naturalWidth>40e6/im.naturalHeight)throw Error('그림이 너무 큽니다. 줄인 그림 파일을 선택해 주세요.');
    if(im.naturalWidth!==header.width||im.naturalHeight!==header.height)throw Error('그림의 크기 정보가 맞지 않습니다. PNG·JPG로 다시 저장해 주세요.');
    return im;
  }
  async function raster(z,pic){
    const im=await decodeImage(await z.get(pic.part));
    const [l,t,r,b]=pic.crop,sw=im.naturalWidth*(1-l-r),sh=im.naturalHeight*(1-t-b);
    if(!(sw>0&&sh>0)||pic.crop.some(v=>v<0||v>=1))throw Error('PPT 그림의 자르기 영역을 읽지 못했습니다.');
    const c=document.createElement('canvas'),scale=1400/Math.max(pic.w,pic.h);c.width=Math.max(1,Math.round(pic.w*scale));c.height=Math.max(1,Math.round(pic.h*scale));
    const g=c.getContext('2d');g.fillStyle='#fff';g.fillRect(0,0,c.width,c.height);
    const m=pic.matrix.map(v=>v/36000*scale);m[4]-=pic.x*scale;m[5]-=pic.y*scale;g.setTransform(...m);
    g.drawImage(im,l*im.naturalWidth,t*im.naturalHeight,sw,sh,0,0,1,1);
    return c.toDataURL('image/jpeg',.96);
  }
  async function inspect(file,progress=()=>{}){
    if(!/\.pptx$/i.test(file.name))throw Error('PPTX 파일을 선택해 주세요. (.ppt 파일은 PPTX로 저장)');
    if(file.size>MAX_FILE)throw Error('PPT 파일은 64MB 이하로 선택해 주세요.');
    const z=zipReader(await file.arrayBuffer()),pres=xml(await z.text('ppt/presentation.xml')),size=one(pres,P,'sldSz');
    const width=attr(size,'cx')/36000,height=attr(size,'cy')/36000;
    if(!(width>0&&height>0))throw Error('PPT 용지 크기를 읽지 못했습니다.');
    const rels=await relationships(z,'ppt/presentation.xml'),ids=all(pres,P,'sldId');
    if(ids.length>1000)throw Error('PPT가 너무 깁니다. 1,000장 이하로 나눠 주세요.');
    const found={back:[],strip:[]};
    for(let i=0;i<ids.length;i++){
      const rel=rels.get(ids[i].getAttributeNS(R,'id'));if(!rel||!rel.type.endsWith('/slide'))continue;
      const doc=xml(await z.text(rel.target)),sr=await relationships(z,rel.target),tree=one(doc,P,'spTree');if(!tree)continue;
      for(const p of pictures(tree,sr)){
        if(p.x<-.5||p.y<-.5||p.x+p.w>width+.5||p.y+p.h>height+.5)continue;
        const ratio=p.w/p.h;
        if(p.w>=width*.28&&p.w<=width*.55&&p.h>=height*.22&&p.h<=height*.48&&ratio>.85&&ratio<1.8&&p.y+p.h/2<height*.55)found.back.push(p);
        if(p.w>=width*.28&&p.w<=width*.55&&p.h>=5&&p.h<=40&&ratio>=2.5&&ratio<=15)found.strip.push(p);
      }
      if(i%10===0){progress(`PPT 확인 중… ${i+1} / ${ids.length}장`);await tick();}
    }
    const result={slides:ids.length,back:null,strip:null,ambiguous:[],unsupported:[]};
    for(const kind of ['back','strip']){
      // Different assets/crops indicate multiple designs: do not silently pick one.
      const unique=new Map();for(const p of found[kind]){
        const norm=p.matrix.map((v,i)=>i<4?Math.round(v/(i%2===0?p.w:p.h)/36000*100)/100:0);
        const key=JSON.stringify([p.part,p.crop,norm]);if(!unique.has(key))unique.set(key,p);
      }
      if(unique.size>1){result.ambiguous.push(kind);continue;}
      if(unique.size===1){const p=unique.values().next().value;try{result[kind]={data:await raster(z,p),width:p.w,height:p.h};}catch(error){result.unsupported.push(kind);}}
    }
    return result;
  }
  function dataImage(url){
    const match=/^data:image\/(png|jpeg);base64,([A-Za-z0-9+/=]+)$/.exec(url||'');if(!match)throw Error('그림을 PNG·JPG 파일로 다시 선택해 주세요.');
    return {ext:match[1]==='png'?'png':'jpeg',bytes:Uint8Array.from(atob(match[2]),c=>c.charCodeAt(0))};
  }
  const relDoc=rels=>`<Relationships xmlns="${PKG}">${rels.map(([id,type,target])=>`<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('')}</Relationships>`;
  function shape(doc,id){return all(doc,P,'cNvPr').find(e=>e.getAttribute('id')===String(id))?.parentNode.parentNode;}
  async function printedBack(data,alreadyPrinted,isOriginal){
    if(!data)return null;const im=await loadImage(data);
    if(alreadyPrinted||isOriginal(im))return dataImage(data);
    const c=document.createElement('canvas');c.width=im.naturalWidth;c.height=im.naturalHeight;const g=c.getContext('2d');g.translate(c.width,c.height);g.rotate(Math.PI);g.drawImage(im,0,0);return dataImage(c.toDataURL('image/png'));
  }
  async function build({rows,logo,back,backPrinted,strip,stripHeight=25,stripCount=0,isOriginal=()=>false},progress=()=>{}){
    if(!rows.length)throw Error('먼저 명단을 넣어 주세요.');
    if(rows.length>5000)throw Error('명단은 5,000명 이하로 나눠 주세요.');
    const template=window.NAMETAG_PPT_TEMPLATE;if(!template)throw Error('PPT 양식을 읽지 못했습니다. 페이지를 다시 열어 주세요.');
    const parts={...template.parts},logoImage=dataImage(logo),backImage=await printedBack(back,backPrinted,isOriginal),stripImage=strip?dataImage(strip):null;
    parts['ppt/media/logo.'+logoImage.ext]=logoImage.bytes;if(backImage)parts['ppt/media/back.'+backImage.ext]=backImage.bytes;if(stripImage)parts['ppt/media/strip.'+stripImage.ext]=stripImage.bytes;
    const slides=[];
    const slideRels=[['rId1','slideLayout','../slideLayouts/slideLayout1.xml'],['rId2','image','../media/logo.'+logoImage.ext]];
    if(backImage)slideRels.push(['rId3','image','../media/back.'+backImage.ext]);
    for(let i=0;i<rows.length;i+=2){
      const doc=xml(template.slide);
      for(let col=0;col<2;col++){
        const row=rows[i+col];
        if(!row){[5,8,10,12,14].forEach(id=>shape(doc,id)?.remove());continue;}
        const body=one(shape(doc,11+col),P,'txBody'),ps=children(body,A,'p');
        for(const [p,text,size] of [[ps[0],row.role,row.roleSize||21],[ps[2],row.name,row.nameSize||65]]){
          const ts=all(p,A,'t');ts[0].textContent=text;ts.slice(1).forEach(t=>t.textContent='');
          all(p,A,'rPr').forEach(e=>e.setAttribute('sz',String(Math.round(size*100))));
        }
        if(!backImage)shape(doc,13+col)?.remove();
      }
      slides.push({xml:serialize(doc),rels:relDoc(slideRels)});
      if(i%20===0){progress(`PPT 만드는 중… ${Math.min(i+2,rows.length)} / ${rows.length}명`);await tick();}
    }
    if(stripImage){
      const h=Math.max(5,Math.min(80,stripHeight)),n=stripCount||rows.length,cols=2,per=cols*Math.floor((190.5-16)/h);
      for(let i=0;i<n;i+=per){
        const doc=xml(template.slide),tree=one(doc,P,'spTree');Array.from(tree.children).filter(e=>!['nvGrpSpPr','grpSpPr'].includes(e.localName)).forEach(e=>e.remove());
        for(let j=0;j<Math.min(per,n-i);j++){
          const x=(254-204)/2+j%cols*102,y=8+Math.floor(j/cols)*h,id=j+2;
          const part=xml(`<p:pic xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:nvPicPr><p:cNvPr id="${id}" name="윗칸 띠 ${i+j+1}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${Math.round(x*36000)}" y="${Math.round(y*36000)}"/><a:ext cx="3672000" cy="${Math.round(h*36000)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="7200"><a:solidFill><a:srgbClr val="9CA3AF"/></a:solidFill></a:ln></p:spPr></p:pic>`);
          tree.appendChild(doc.importNode(part.documentElement,true));
        }
        slides.push({xml:serialize(doc),rels:relDoc([['rId1','slideLayout','../slideLayouts/slideLayout1.xml'],['rId4','image','../media/strip.'+stripImage.ext]])});
      }
    }
    const pres=xml(`<p:presentation xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${slides.map((s,i)=>`<p:sldId id="${256+i}" r:id="rId${i+2}"/>`).join('')}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
    parts['ppt/presentation.xml']=serialize(pres);parts['_rels/.rels']=relDoc([['rId1','officeDocument','ppt/presentation.xml']]);
    parts['ppt/_rels/presentation.xml.rels']=relDoc([['rId1','slideMaster','slideMasters/slideMaster1.xml'],...slides.map((s,i)=>['rId'+(i+2),'slide','slides/slide'+(i+1)+'.xml'])]);
    slides.forEach((s,i)=>{parts[`ppt/slides/slide${i+1}.xml`]=s.xml;parts[`ppt/slides/_rels/slide${i+1}.xml.rels`]=s.rels;});
    const types=[['/ppt/presentation.xml','presentation'],['/ppt/slideMasters/slideMaster1.xml','slideMaster'],['/ppt/slideLayouts/slideLayout1.xml','slideLayout'],...slides.map((s,i)=>[`/ppt/slides/slide${i+1}.xml`,'slide'])];
    parts['[Content_Types].xml']=`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>${types.map(([n,t])=>`<Override PartName="${n}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.${t==='presentation'?'presentation.main':t}+xml"/>`).join('')}</Types>`;
    return {blob:zipWrite(parts),slides:slides.length,badgeSlides:Math.ceil(rows.length/2),stripSlides:slides.length-Math.ceil(rows.length/2)};
  }
  window.NametagPpt={inspect,build,zipReader,zipWrite,crc32,imageHeader,decodeImage};
})();

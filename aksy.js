<script>
async function readXlsx(arrayBuffer){
  const files = await unzip(new Uint8Array(arrayBuffer));
  const dec = new TextDecoder('utf-8');
  let shared=[];
  if(files['xl/sharedStrings.xml']){
    const doc=new DOMParser().parseFromString(dec.decode(files['xl/sharedStrings.xml']),'application/xml');
    shared=[...doc.getElementsByTagName('si')].map(si=>[...si.getElementsByTagName('t')].map(t=>t.textContent).join(''));
  }
  const sheetName = Object.keys(files).find(n=>/^xl\/worksheets\/sheet1\.xml$/.test(n)) || Object.keys(files).find(n=>/^xl\/worksheets\/.*\.xml$/.test(n));
  if(!sheetName) throw new Error("Не знайдено аркушів у файлі");
  const doc=new DOMParser().parseFromString(dec.decode(files[sheetName]),'application/xml');
  const rowsEl=[...doc.getElementsByTagName('row')];
  const grid=[]; let maxCol=0;
  for(const r of rowsEl){
    const ri=parseInt(r.getAttribute('r'),10)-1;
    const cells=[...r.getElementsByTagName('c')];
    const arr=[];
    for(const c of cells){
      const ref=c.getAttribute('r'); const col=colNum(ref);
      const t=c.getAttribute('t');
      let v=null;
      const vEl=c.getElementsByTagName('v')[0];
      const isEl=c.getElementsByTagName('is')[0];
      if(t==='s'){ const idx=vEl?parseInt(vEl.textContent,10):-1; v=shared[idx]!=null?shared[idx]:''; }
      else if(t==='inlineStr'&&isEl){ v=[...isEl.getElementsByTagName('t')].map(x=>x.textContent).join(''); }
      else if(vEl){ v=vEl.textContent; }
      arr[col]=v;
      if(col>maxCol)maxCol=col;
    }
    grid[ri]=arr;
  }
  for(const row of grid){ if(row) row.length=Math.max(row.length,maxCol+1); }
  return grid;
}
function colNum(ref){
  const m=/^([A-Z]+)/.exec(ref||''); if(!m)return 0;
  let n=0; for(const ch of m[1]) n=n*26+(ch.charCodeAt(0)-64); return n-1;
}
async function unzip(buf){
  const dv=new DataView(buf.buffer); const out={}; let eocd=-1;
  for(let i=buf.length-22;i>=0;i--){ if(dv.getUint32(i,true)===0x06054b50){eocd=i;break;} }
  if(eocd<0) throw new Error('Не схоже на xlsx');
  const cdOffset=dv.getUint32(eocd+16,true); const cdCount=dv.getUint16(eocd+10,true); let p=cdOffset; const entries=[];
  for(let i=0;i<cdCount;i++){
    if(dv.getUint32(p,true)!==0x02014b50) break;
    const method=dv.getUint16(p+10,true), compSize=dv.getUint32(p+20,true), nameLen=dv.getUint16(p+28,true);
    const extraLen=dv.getUint16(p+30,true), commentLen=dv.getUint16(p+32,true), localOff=dv.getUint32(p+42,true);
    const name=new TextDecoder().decode(buf.subarray(p+46,p+46+nameLen));
    entries.push({name,method,compSize,localOff});
    p+=46+nameLen+extraLen+commentLen;
  }
  for(const e of entries){
    if(!/(sharedStrings|worksheets\/sheet1|worksheets\/.*)\.xml$/.test(e.name)) continue;
    const lp=e.localOff, lNameLen=dv.getUint16(lp+26,true), lExtra=dv.getUint16(lp+28,true);
    const dataStart=lp+30+lNameLen+lExtra, comp=buf.subarray(dataStart,dataStart+e.compSize);
    if(e.method===0){ out[e.name]=comp; } else if(e.method===8){ out[e.name]=await inflateRaw(comp); }
  }
  return out;
}
async function inflateRaw(bytes){
  if(typeof DecompressionStream!=='undefined'){
    try{ const ds=new DecompressionStream('deflate-raw'); const stream=new Response(bytes).body.pipeThrough(ds); const ab=await new Response(stream).arrayBuffer(); return new Uint8Array(ab); }catch(e){}
  }
  return tinfInflate(bytes);
}
function tinfInflate(source){
  const TINF_OK=0; function Tree(){this.table=new Uint16Array(16);this.trans=new Uint16Array(288);} function Data(src){this.s=src;this.i=0;this.tag=0;this.bitcount=0; this.dest=new Uint8Array(src.length*8+1024);this.destLen=0; this.ltree=new Tree();this.dtree=new Tree();} const sltree=new Tree(),sdtree=new Tree(); const length_bits=new Uint8Array(30),length_base=new Uint16Array(30); const dist_bits=new Uint8Array(30),dist_base=new Uint16Array(30); const clcidx=new Uint8Array([16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15]); function buildBitsBase(bits,base,delta,first){let i,sum=first; for(i=0;i<delta;++i)bits[i]=0; for(i=0;i<30-delta;++i)bits[i+delta]=(i/delta)|0; for(i=0;i<30;++i){base[i]=sum;sum+=1<<bits[i];}} function buildFixedTrees(lt,dt){let i; for(i=0;i<7;++i)lt.table[i]=0; lt.table[7]=24;lt.table[8]=152;lt.table[9]=112; for(i=0;i<24;++i)lt.trans[i]=256+i; for(i=0;i<144;++i)lt.trans[24+i]=i; for(i=0;i<8;++i)lt.trans[24+144+i]=280+i; for(i=0;i<112;++i)lt.trans[24+144+8+i]=144+i; for(i=0;i<5;++i)dt.table[i]=0; dt.table[5]=32; for(i=0;i<32;++i)dt.trans[i]=i;} const offs=new Uint16Array(16); function buildTree(t,lengths,off,num){let i,sum; for(i=0;i<16;++i)t.table[i]=0; for(i=0;i<num;++i)t.table[lengths[off+i]]++; t.table[0]=0; for(sum=0,i=0;i<16;++i){offs[i]=sum;sum+=t.table[i];} for(i=0;i<num;++i)if(lengths[off+i])t.trans[offs[lengths[off+i]]++]=i;} function getbit(d){if(!d.bitcount--){d.tag=d.s[d.i++];d.bitcount=7;} const bit=d.tag&1;d.tag>>>=1;return bit;} function readBits(d,num,base){if(!num)return base; while(d.bitcount<24){d.tag|=d.s[d.i++]<<d.bitcount;d.bitcount+=8;} const val=d.tag&(0xffff>>>(16-num));d.tag>>>=num;d.bitcount-=num;return val+base;} function decodeSymbol(d,t){ while(d.bitcount<24){d.tag|=d.s[d.i++]<<d.bitcount;d.bitcount+=8;} let sum=0,cur=0,len=0,tag=d.tag; do{cur=2*cur+(tag&1);tag>>>=1;++len;sum+=t.table[len];cur-=t.table[len];}while(cur>=0); d.tag=tag;d.bitcount-=len;return t.trans[sum+cur];} function decodeTrees(d,lt,dt){const lengths=new Uint8Array(288+32); let hlit,hdist,hclen,i,num,length; hlit=readBits(d,5,257);hdist=readBits(d,5,1);hclen=readBits(d,4,4); for(i=0;i<19;++i)lengths[i]=0; for(i=0;i<hclen;++i){const clen=readBits(d,3,0);lengths[clcidx[i]]=clen;} buildTree(d.codeTree,lengths,0,19); for(num=0;num<hlit+hdist;){const sym=decodeSymbol(d,d.codeTree); switch(sym){ case 16:{let prev=lengths[num-1];for(length=readBits(d,2,3);length;--length)lengths[num++]=prev;break;} case 17:for(length=readBits(d,3,3);length;--length)lengths[num++]=0;break; case 18:for(length=readBits(d,7,11);length;--length)lengths[num++]=0;break; default:lengths[num++]=sym;break;}} buildTree(lt,lengths,0,hlit);buildTree(dt,lengths,hlit,hdist);} function ensure(d,extra){if(d.destLen+extra>d.dest.length){const n=new Uint8Array((d.dest.length+extra)*2);n.set(d.dest);d.dest=n;}} function inflateBlockData(d,lt,dt){ for(;;){let sym=decodeSymbol(d,lt); if(sym===256)return TINF_OK; if(sym<256){ensure(d,1);d.dest[d.destLen++]=sym;} else{sym-=257;const length=readBits(d,length_bits[sym],length_base[sym]); const dist=decodeSymbol(d,dt);const offset=d.destLen-readBits(d,dist_bits[dist],dist_base[dist]); ensure(d,length);for(let i=offset;i<offset+length;++i)d.dest[d.destLen++]=d.dest[i];}}} function inflateUncompressed(d){ while(d.bitcount>8){d.i--;d.bitcount-=8;} let length=d.s[d.i+1];length=256*length+d.s[d.i];d.i+=4; ensure(d,length);for(let i=length;i;--i)d.dest[d.destLen++]=d.s[d.i++]; d.bitcount=0;return TINF_OK;} buildBitsBase(length_bits,length_base,4,3); buildBitsBase(dist_bits,dist_base,2,1); length_bits[28]=0;length_base[28]=258; buildFixedTrees(sltree,sdtree); const d=new Data(source);d.codeTree=new Tree(); let bfinal,btype; do{bfinal=getbit(d);btype=readBits(d,2,0); if(btype===0)inflateUncompressed(d); else if(btype===1)inflateBlockData(d,sltree,sdtree); else{decodeTrees(d,d.ltree,d.dtree);inflateBlockData(d,d.ltree,d.dtree);} }while(!bfinal); return d.dest.subarray(0,d.destLen);
}
</script>
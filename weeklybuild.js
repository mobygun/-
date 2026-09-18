/* 회사 엑셀 원본(template3.js)에 주간업무일지 내용을 채워 넣어 .xlsx 를 만드는 모듈 */
(function(global){
  const T = () => global.WEEKLY_TEMPLATE;

  function b64ToBytes(b64){
    const bin = atob(b64), out = new Uint8Array(bin.length);
    for(let i=0;i<bin.length;i++) out[i]=bin.charCodeAt(i);
    return out;
  }
  function textOf(name){ return new TextDecoder("utf-8").decode(b64ToBytes(T().files[name])); }
  function bytesOfText(s){ return new TextEncoder().encode(s); }

  function escXml(v){
    return String(v==null?"":v).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
  }
  function setCell(rowXml, ref, value, numeric){
    const re = new RegExp('<c r="'+ref+'"([^>]*?)(\\/>|>[\\s\\S]*?<\\/c>)');
    const m = rowXml.match(re);
    if(!m) return rowXml;
    const sm = m[1].match(/s="(\d+)"/);
    const s = sm ? ' s="'+sm[1]+'"' : '';
    let cell;
    if(value===null || value==="" || value===undefined) cell = '<c r="'+ref+'"'+s+'/>';
    else if(numeric) cell = '<c r="'+ref+'"'+s+'><v>'+value+'</v></c>';
    else cell = '<c r="'+ref+'"'+s+' t="inlineStr"><is><t xml:space="preserve">'+escXml(value)+'</t></is></c>';
    return rowXml.slice(0,m.index) + cell + rowXml.slice(m.index+m[0].length);
  }

  function toExcelSerial(dateStr){
    const [y,mo,d] = dateStr.split("-").map(Number);
    const utc = Date.UTC(y, mo-1, d);
    const epoch = Date.UTC(1899,11,30);
    return Math.round((utc-epoch)/86400000);
  }
  function colLetter(n){
    let s="";
    while(n>0){ const r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26); }
    return s;
  }
  function ymd2(dateStr){
    const [y,mo,d] = dateStr.split("-");
    return y.slice(2)+"."+mo+"."+d;
  }

  const MEMBER_COLS = [3,5,7,9,11,13,15]; // C,E,G,I,K,M,O (1-based). 최대 7명까지만 엑셀에 반영됨.

  function buildSheet(weekDoc){
    let sheet = textOf("xl/worksheets/sheet24.xml");
    const mm = sheet.match(/([\s\S]*?<sheetData>)([\s\S]*?)(<\/sheetData>[\s\S]*)/);
    const head = mm[1], sd = mm[2], tail = mm[3];
    const rows = sd.match(/<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g) || [];
    const byNum = {};
    rows.forEach(r=>{ byNum[+r.match(/<row r="(\d+)"/)[1]] = r; });

    const members = (weekDoc.members||[]).slice(0,7);
    const entries = weekDoc.entries||{};
    const remarks = weekDoc.remarks||[];
    const days = weekDoc.days||[];

    byNum[4] = setCell(byNum[4], "C4", weekDoc.dept||"인프라공사부");
    byNum[4] = setCell(byNum[4], "F4", weekDoc.reporter||"");
    byNum[5] = setCell(byNum[5], "C5", ymd2(weekDoc.weekStart)+"~"+ymd2(weekDoc.weekEnd));

    for(let i=0;i<7;i++){
      byNum[11] = setCell(byNum[11], colLetter(MEMBER_COLS[i])+"11", members[i]||"");
    }

    for(let d=0; d<6; d++){
      const rn = 12+d;
      const day = days[d];
      if(day && day.date) byNum[rn] = setCell(byNum[rn], "A"+rn, toExcelSerial(day.date), true);
      for(let i=0;i<7;i++){
        const name = members[i];
        const text = name ? ((entries[name]||[])[d] || "") : "";
        byNum[rn] = setCell(byNum[rn], colLetter(MEMBER_COLS[i])+rn, text);
      }
      byNum[rn] = setCell(byNum[rn], "Q"+rn, remarks[d]||"");
    }

    byNum[18] = setCell(byNum[18], "C18", weekDoc.nextWeekPlan||"");

    const out = [];
    for(let i=1;i<=20;i++) if(byNum[i]) out.push(byNum[i]);
    return head + out.join("") + tail;
  }

  function crcTable(){
    const t = new Uint32Array(256);
    for(let n=0;n<256;n++){
      let c=n;
      for(let k=0;k<8;k++) c = (c&1) ? (0xEDB88320 ^ (c>>>1)) : (c>>>1);
      t[n]=c>>>0;
    }
    return t;
  }
  const CRCT = crcTable();
  function crc32(buf){
    let c = 0xFFFFFFFF;
    for(let i=0;i<buf.length;i++) c = CRCT[(c ^ buf[i]) & 0xFF] ^ (c>>>8);
    return (c ^ 0xFFFFFFFF)>>>0;
  }
  async function deflateRaw(bytes){
    if(typeof CompressionStream==="undefined") return null;
    try{
      const cs=new CompressionStream("deflate-raw");
      const stream=new Blob([bytes]).stream().pipeThrough(cs);
      const buf=await new Response(stream).arrayBuffer();
      return new Uint8Array(buf);
    }catch(e){ return null; }
  }
  async function zipStore(files){
    const enc=new TextEncoder(), chunks=[], central=[]; let offset=0;
    const now=new Date();
    const dosTime=((now.getHours()<<11)|(now.getMinutes()<<5)|(now.getSeconds()>>1))&0xFFFF;
    const dosDate=(((now.getFullYear()-1980)<<9)|((now.getMonth()+1)<<5)|now.getDate())&0xFFFF;
    for(const f of files){
      const nameB=enc.encode(f.name), raw=f.bytes, crc=crc32(raw);
      const packed=await deflateRaw(raw);
      const useDeflate = packed && packed.length < raw.length;
      const data = useDeflate ? packed : raw;
      const method = useDeflate ? 8 : 0;
      const lh=new Uint8Array(30+nameB.length), dv=new DataView(lh.buffer);
      dv.setUint32(0,0x04034b50,true); dv.setUint16(4,20,true); dv.setUint16(6,0x0800,true);
      dv.setUint16(8,method,true); dv.setUint16(10,dosTime,true); dv.setUint16(12,dosDate,true);
      dv.setUint32(14,crc,true); dv.setUint32(18,data.length,true); dv.setUint32(22,raw.length,true);
      dv.setUint16(26,nameB.length,true); dv.setUint16(28,0,true);
      lh.set(nameB,30); chunks.push(lh,data);
      const ch=new Uint8Array(46+nameB.length), cv=new DataView(ch.buffer);
      cv.setUint32(0,0x02014b50,true); cv.setUint16(4,20,true); cv.setUint16(6,20,true);
      cv.setUint16(8,0x0800,true); cv.setUint16(10,method,true);
      cv.setUint16(12,dosTime,true); cv.setUint16(14,dosDate,true);
      cv.setUint32(16,crc,true); cv.setUint32(20,data.length,true); cv.setUint32(24,raw.length,true);
      cv.setUint16(28,nameB.length,true); cv.setUint32(42,offset,true);
      ch.set(nameB,46); central.push(ch);
      offset+=lh.length+data.length;
    }
    const cdSize=central.reduce((a,c)=>a+c.length,0);
    const end=new Uint8Array(22), ev=new DataView(end.buffer);
    ev.setUint32(0,0x06054b50,true);
    ev.setUint16(8,files.length,true); ev.setUint16(10,files.length,true);
    ev.setUint32(12,cdSize,true); ev.setUint32(16,offset,true);
    return new Blob([...chunks, ...central, end],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  }

  global.buildWeeklyXlsx = async function(weekDoc){
    const sheetNew = buildSheet(weekDoc);
    const label = weekDoc.label || "주간업무일지";
    const wbx = textOf("xl/workbook.xml").replace(/2026년 09월 3주차/g, label);

    const overrides = {
      "xl/worksheets/sheet24.xml": sheetNew,
      "xl/workbook.xml": wbx
    };
    const files = T().order.map(name => ({
      name,
      bytes: overrides[name] ? bytesOfText(overrides[name]) : b64ToBytes(T().files[name])
    }));
    return await zipStore(files);
  };
})(typeof window!=="undefined" ? window : globalThis);

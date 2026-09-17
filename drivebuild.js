/* 회사 차량 운행기록부 원본(template2.js)에 운행 내역을 채워 넣어 .xlsx 를 만드는 모듈 */
(function(global){
  const T = () => global.DRIVE_TEMPLATE;

  function b64ToBytes(b64){
    const bin = atob(b64), out = new Uint8Array(bin.length);
    for(let i=0;i<bin.length;i++) out[i]=bin.charCodeAt(i);
    return out;
  }
  function textOf(name){ return new TextDecoder("utf-8").decode(b64ToBytes(T().files[name])); }
  function bytesOfText(s){ return new TextEncoder().encode(s); }
  function escXml(v){ return String(v==null?"":v).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }

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
  // 엑셀 날짜 일련번호 (1899-12-30 기준)
  function serial(y,m,d){ return Math.round((Date.UTC(y,m-1,d) - Date.UTC(1899,11,30))/86400000); }
  function pad2(n){ return String(n).padStart(2,"0"); }

  function crcTable(){
    const t=new Uint32Array(256);
    for(let n=0;n<256;n++){ let c=n; for(let k=0;k<8;k++) c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1); t[n]=c>>>0; }
    return t;
  }
  const CRCT=crcTable();
  function crc32(b){ let c=0xFFFFFFFF; for(let i=0;i<b.length;i++) c=CRCT[(c^b[i])&0xFF]^(c>>>8); return (c^0xFFFFFFFF)>>>0; }
  function zipStore(files){
    const enc=new TextEncoder(), chunks=[], central=[]; let offset=0;
    const now=new Date();
    const dosTime=((now.getHours()<<11)|(now.getMinutes()<<5)|(now.getSeconds()>>1))&0xFFFF;
    const dosDate=(((now.getFullYear()-1980)<<9)|((now.getMonth()+1)<<5)|now.getDate())&0xFFFF;
    files.forEach(f=>{
      const nameB=enc.encode(f.name), data=f.bytes, crc=crc32(data);
      const lh=new Uint8Array(30+nameB.length), dv=new DataView(lh.buffer);
      dv.setUint32(0,0x04034b50,true); dv.setUint16(4,20,true); dv.setUint16(6,0x0800,true);
      dv.setUint16(8,0,true); dv.setUint16(10,dosTime,true); dv.setUint16(12,dosDate,true);
      dv.setUint32(14,crc,true); dv.setUint32(18,data.length,true); dv.setUint32(22,data.length,true);
      dv.setUint16(26,nameB.length,true); dv.setUint16(28,0,true);
      lh.set(nameB,30); chunks.push(lh,data);
      const ch=new Uint8Array(46+nameB.length), cv=new DataView(ch.buffer);
      cv.setUint32(0,0x02014b50,true); cv.setUint16(4,20,true); cv.setUint16(6,20,true);
      cv.setUint16(8,0x0800,true); cv.setUint16(10,0,true);
      cv.setUint16(12,dosTime,true); cv.setUint16(14,dosDate,true);
      cv.setUint32(16,crc,true); cv.setUint32(20,data.length,true); cv.setUint32(24,data.length,true);
      cv.setUint16(28,nameB.length,true); cv.setUint32(42,offset,true);
      ch.set(nameB,46); central.push(ch);
      offset+=lh.length+data.length;
    });
    const cdSize=central.reduce((a,c)=>a+c.length,0);
    const end=new Uint8Array(22), ev=new DataView(end.buffer);
    ev.setUint32(0,0x06054b50,true);
    ev.setUint16(8,files.length,true); ev.setUint16(10,files.length,true);
    ev.setUint32(12,cdSize,true); ev.setUint32(16,offset,true);
    return new Blob([...chunks,...central,end],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  }

  /* days: { "2026-10-03": {before, after, work, commute, dept, name, note} } */
  global.buildDriveXlsx = function(days, st, year, monthNum){
    let sheet = textOf("xl/worksheets/sheet1.xml");
    const mm = sheet.match(/([\s\S]*?<sheetData>)([\s\S]*?)(<\/sheetData>[\s\S]*)/);
    const head=mm[1], sd=mm[2], tail=mm[3];
    const rows = sd.match(/<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g) || [];
    const byNum={};
    rows.forEach(r=>{ byNum[+r.match(/<row r="(\d+)"/)[1]] = r; });

    const y=+year, mo=+monthNum;
    const lastDay = new Date(y, mo, 0).getDate();

    // 상단 정보
    byNum[2] = setCell(byNum[2], "C2", y+". "+pad2(mo)+". 01");
    byNum[3] = setCell(byNum[3], "C3", y+". "+pad2(mo)+". "+pad2(lastDay));
    byNum[4] = setCell(byNum[4], "E4", st.carType||"");
    byNum[4] = setCell(byNum[4], "I4", st.carNo||"");
    byNum[4] = setCell(byNum[4], "M4", st.name||st.writer||"");

    // 데이터 행 10~40 : 그 달의 모든 날짜를 미리 깔고, 운행한 날만 값 채움
    for(let i=0;i<31;i++){
      const rn=10+i, day=i+1;
      let r=byNum[rn];
      if(!r) continue;
      if(day<=lastDay){
        const ser=serial(y,mo,day);
        r=setCell(r,"A"+rn,ser,true);   // 서식 m"/"d 로 표시
        r=setCell(r,"B"+rn,ser,true);   // 서식 aaa 로 요일 표시
        const key=y+"-"+pad2(mo)+"-"+pad2(day);
        const d=days[key];
        if(d){
          r=setCell(r,"C"+rn,d.dept||"");
          r=setCell(r,"D"+rn,d.name||"");
          r=setCell(r,"E"+rn,d.before,true);
          r=setCell(r,"F"+rn,d.after,true);
          r=setCell(r,"H"+rn,d.commute? d.commute : "", !!d.commute);
          r=setCell(r,"I"+rn,d.work? d.work : "", !!d.work);
          r=setCell(r,"J"+rn,d.note||"");
        }
      }
      byNum[rn]=r;
    }

    const out=[];
    Object.keys(byNum).map(Number).sort((a,b)=>a-b).forEach(k=>out.push(byNum[k]));
    const sheetNew = head + out.join("") + tail;

    const tabName = y+"."+pad2(mo)+"월";
    let wbx = textOf("xl/workbook.xml").replace(/2026\.00월/g, tabName);
    let rels = textOf("xl/_rels/workbook.xml.rels").replace(/<Relationship[^>]*calcChain\.xml"\/>/, "");
    let ct = textOf("[Content_Types].xml").replace(/<Override PartName="\/xl\/calcChain\.xml"[^>]*\/>/, "");

    const overrides={
      "xl/worksheets/sheet1.xml": sheetNew,
      "xl/workbook.xml": wbx,
      "xl/_rels/workbook.xml.rels": rels,
      "[Content_Types].xml": ct
    };
    const files=T().order.map(name=>({
      name,
      bytes: overrides[name] ? bytesOfText(overrides[name]) : b64ToBytes(T().files[name])
    }));
    return zipStore(files);
  };
})(typeof window!=="undefined" ? window : globalThis);

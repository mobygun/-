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

  // 수식 셀의 "저장된 계산값"을 갱신 (수식을 못 읽는 폰 앱에서도 숫자가 보이도록)
  function setCachedValue(rowXml, ref, value){
    const re = new RegExp('<c r="'+ref+'"([^>]*?)(\\/>|>([\\s\\S]*?)<\\/c>)');
    const m = rowXml.match(re);
    if(!m) return rowXml;
    let attrs = m[1].replace(/\st="[a-z]+"/g, "");   // 오류(t="e") 표시 제거
    let inner = m[3] || "";
    const f = inner.match(/<f[\s\S]*?<\/f>|<f[^>]*\/>/);
    const fx = f ? f[0] : "";
    const v = (value===""||value===null||value===undefined) ? "" : "<v>"+value+"</v>";
    const cell = '<c r="'+ref+'"'+attrs+'>'+fx+v+'</c>';
    return rowXml.slice(0,m.index) + cell + rowXml.slice(m.index+m[0].length);
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

  global.buildDriveXlsx = async function(days, st, year, monthNum){
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
          r=setCell(r,"K"+rn,d.fuel? d.fuel:"", !!d.fuel);      // 주유비
          r=setCell(r,"L"+rn,d.toll? d.toll:"", !!d.toll);      // 통행료
          r=setCell(r,"M"+rn,d.park? d.park:"", !!d.park);      // 주차비
          r=setCell(r,"N"+rn,d.etc? d.etc:"", !!d.etc);         // 기타
          r=setCachedValue(r,"G"+rn,(Number(d.work)||0)+(Number(d.commute)||0)); // ⑦주행거리
        }
      }
      byNum[rn]=r;
    }

    // 합계 행(41,42)의 계산값도 직접 채움
    let tWork=0,tComm=0,tFuel=0,tToll=0,tPark=0,tEtc=0;
    Object.values(days).forEach(d=>{
      tWork+=Number(d.work)||0; tComm+=Number(d.commute)||0;
      tFuel+=Number(d.fuel)||0; tToll+=Number(d.toll)||0;
      tPark+=Number(d.park)||0; tEtc+=Number(d.etc)||0;
    });
    const tTotal=tWork+tComm, tCost=tFuel+tToll+tPark+tEtc;
    if(byNum[41]){
      let r=byNum[41];
      r=setCachedValue(r,"K41",tFuel); r=setCachedValue(r,"L41",tToll);
      r=setCachedValue(r,"M41",tPark); r=setCachedValue(r,"N41",tEtc);
      byNum[41]=r;
    }
    if(byNum[42]){
      let r=byNum[42];
      r=setCachedValue(r,"E42",tTotal);                                  // ⑪총주행거리
      r=setCachedValue(r,"H42",tTotal);                                  // ⑫업무용 사용거리
      r=setCachedValue(r,"J42",tTotal? 1 : 0);                           // ⑬업무사용비율(출퇴근 포함)
      r=setCachedValue(r,"M42",tCost);                                   // 비용합계
      byNum[42]=r;
    }

    const out=[];
    Object.keys(byNum).map(Number).sort((a,b)=>a-b).forEach(k=>out.push(byNum[k]));
    const sheetNew = head + out.join("") + tail;

    const tabName = y+"."+pad2(mo)+"월";
    let wbx = textOf("xl/workbook.xml").replace(/2026\.00월/g, tabName);
    // 엑셀이 열릴 때 모든 수식을 다시 계산하도록 (소계·합계가 0으로 보이는 문제 방지)
    wbx = wbx.replace(/<calcPr([^>]*?)\/>/, '<calcPr$1 fullCalcOnLoad="1"/>');
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
    return await zipStore(files);
  };
})(typeof window!=="undefined" ? window : globalThis);

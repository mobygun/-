/* 회사 엑셀 원본(template.js)에 영수증 내역만 채워 넣어 .xlsx 를 만드는 모듈 */
(function(global){
  const T = () => global.XLSX_TEMPLATE;

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
  function shiftRow(rowXml, delta){
    if(!delta) return rowXml;
    return rowXml
      .replace(/<row r="(\d+)"/g, (x,n)=>'<row r="'+(+n+delta)+'"')
      .replace(/r="([A-Z]+)(\d+)"/g, (x,c,n)=>'r="'+c+(+n+delta)+'"');
  }
  function pad2(n){ return String(n).padStart(2,"0"); }

  function buildSheet(entries, st, year, monthNum){
    let sheet = textOf("xl/worksheets/sheet3.xml");
    const mm = sheet.match(/([\s\S]*?<sheetData>)([\s\S]*?)(<\/sheetData>[\s\S]*)/);
    const head = mm[1], sd = mm[2], tail = mm[3];
    const rows = sd.match(/<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g) || [];
    const byNum = {};
    rows.forEach(r=>{ byNum[+r.match(/<row r="(\d+)"/)[1]] = r; });

    const n = Math.max(entries.length, 32);
    const delta = n - 32;
    const ds = entries.map(e=>e.date).sort();
    const sDD = ds.length ? ds[0].slice(8) : "00";
    const eDD = ds.length ? ds[ds.length-1].slice(8) : "00";

    byNum[2] = setCell(byNum[2], "A2", "( "+monthNum+" )월 현장 경비 지출대장");
    byNum[3] = setCell(byNum[3], "A3", "["+year+". "+pad2(monthNum)+". "+sDD+" ~ "+pad2(monthNum)+". "+eDD+"]");
    byNum[5] = setCell(byNum[5], "B5", st.siteName||"");
    byNum[5] = setCell(byNum[5], "F5", st.dept||"");
    byNum[6] = setCell(byNum[6], "B6", st.cardNumber||"");
    byNum[6] = setCell(byNum[6], "F6", st.writer||"");

    const out = [];
    for(let i=1;i<=10;i++) if(byNum[i]) out.push(byNum[i]);

    for(let i=0;i<n;i++){
      const rn = 11+i;
      const src = byNum[rn] || byNum[42];
      let r = shiftRow(src, rn - (+src.match(/<row r="(\d+)"/)[1]));
      const e = entries[i];
      if(e){
        const corp = e.pay === "법인카드";
        r = setCell(r, "A"+rn, (+e.date.slice(5,7))+"/"+(+e.date.slice(8)));
        r = setCell(r, "B"+rn, e.cat);
        r = setCell(r, "C"+rn, e.vendor||"");
        r = setCell(r, "D"+rn, corp? e.amount : "", corp);
        r = setCell(r, "E"+rn, corp? "" : e.amount, !corp);
        r = setCell(r, "F"+rn, e.note||"");
      }
      out.push(r);
    }

    const lastData = 10+n;
    Object.keys(byNum).map(Number).filter(k=>k>=43).sort((a,b)=>a-b).forEach(i=>{
      let r = shiftRow(byNum[i], delta);
      if(i===43){
        r = r.replace(/<f>SUM\(D11:D\d+\)<\/f>/, "<f>SUM(D11:D"+lastData+")</f>")
             .replace(/<f>SUM\(E11:E\d+\)<\/f>/, "<f>SUM(E11:E"+lastData+")</f>")
             .replace(/<v>0<\/v>/g, "");
      }
      if(i===44){
        r = r.replace(/<f>D\d+\+E\d+<\/f>/, "<f>D"+(43+delta)+"+E"+(43+delta)+"</f>")
             .replace(/<v>0<\/v>/g, "");
      }
      out.push(r);
    });

    let sheetNew = head + out.join("") + tail;
    sheetNew = sheetNew.replace(/<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/g,
      (x,a,ar,b,br)=>{
        let A=+ar, B=+br;
        if(A>=43) A+=delta;
        if(B>=43) B+=delta;
        return '<mergeCell ref="'+a+A+':'+b+B+'"/>';
      });
    sheetNew = sheetNew.replace(/<dimension ref="A1:J\d+"\/>/, '<dimension ref="A1:J'+(60+delta)+'"/>');
    return {sheetNew, delta};
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
  function zipStore(files){ // files: [{name, bytes}]
    const enc = new TextEncoder();
    const chunks = [], central = [];
    let offset = 0;
    const now = new Date();
    const dosTime = ((now.getHours()<<11)|(now.getMinutes()<<5)|(now.getSeconds()>>1)) & 0xFFFF;
    const dosDate = (((now.getFullYear()-1980)<<9)|((now.getMonth()+1)<<5)|now.getDate()) & 0xFFFF;
    files.forEach(f=>{
      const nameB = enc.encode(f.name), data = f.bytes, crc = crc32(data);
      const lh = new Uint8Array(30+nameB.length), dv = new DataView(lh.buffer);
      dv.setUint32(0,0x04034b50,true); dv.setUint16(4,20,true); dv.setUint16(6,0x0800,true);
      dv.setUint16(8,0,true); dv.setUint16(10,dosTime,true); dv.setUint16(12,dosDate,true);
      dv.setUint32(14,crc,true); dv.setUint32(18,data.length,true); dv.setUint32(22,data.length,true);
      dv.setUint16(26,nameB.length,true); dv.setUint16(28,0,true);
      lh.set(nameB,30);
      chunks.push(lh, data);
      const ch = new Uint8Array(46+nameB.length), cv = new DataView(ch.buffer);
      cv.setUint32(0,0x02014b50,true); cv.setUint16(4,20,true); cv.setUint16(6,20,true);
      cv.setUint16(8,0x0800,true); cv.setUint16(10,0,true);
      cv.setUint16(12,dosTime,true); cv.setUint16(14,dosDate,true);
      cv.setUint32(16,crc,true); cv.setUint32(20,data.length,true); cv.setUint32(24,data.length,true);
      cv.setUint16(28,nameB.length,true); cv.setUint32(42,offset,true);
      ch.set(nameB,46);
      central.push(ch);
      offset += lh.length + data.length;
    });
    const cdSize = central.reduce((a,c)=>a+c.length,0);
    const end = new Uint8Array(22), ev = new DataView(end.buffer);
    ev.setUint32(0,0x06054b50,true);
    ev.setUint16(8,files.length,true); ev.setUint16(10,files.length,true);
    ev.setUint32(12,cdSize,true); ev.setUint32(16,offset,true);
    return new Blob([...chunks, ...central, end], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  }

  global.buildExpenseXlsx = function(entries, settings, year, monthNum){
    const {sheetNew, delta} = buildSheet(entries, settings, year, monthNum);
    let drawing = textOf("xl/drawings/drawing1.xml");
    if(delta) drawing = drawing.replace(/<xdr:row>(\d+)<\/xdr:row>/g,
      (x,n)=> "<xdr:row>"+(+n>=42 ? +n+delta : +n)+"</xdr:row>");
    let wbx = textOf("xl/workbook.xml").replace("2026_00월 상세내역_현장명", year+"_"+monthNum+"월 상세내역");
    let rels = textOf("xl/_rels/workbook.xml.rels").replace(/<Relationship[^>]*calcChain\.xml"\/>/, "");
    let ct = textOf("[Content_Types].xml").replace(/<Override PartName="\/xl\/calcChain\.xml"[^>]*\/>/, "");

    const overrides = {
      "xl/worksheets/sheet3.xml": sheetNew,
      "xl/drawings/drawing1.xml": drawing,
      "xl/workbook.xml": wbx,
      "xl/_rels/workbook.xml.rels": rels,
      "[Content_Types].xml": ct
    };
    const files = T().order.map(name => ({
      name,
      bytes: overrides[name] ? bytesOfText(overrides[name]) : b64ToBytes(T().files[name])
    }));
    return zipStore(files);
  };
})(typeof window!=="undefined" ? window : globalThis);

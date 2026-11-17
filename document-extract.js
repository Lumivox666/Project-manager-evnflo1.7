'use strict';
// Extract searchable text from external documents. Uploaded originals remain unchanged.
// PDF.js is installed by Railpack from package.json. Text-only fallback permits
// local/self-hosted operation even when optional packages are unavailable.
const zlib=require('node:zlib');
const MAX_TEXT=350000;
const clean=s=>String(s||'').replace(/\u0000/g,'').replace(/[ \t]+/g,' ').replace(/\n[ \t]+/g,'\n').replace(/\n{4,}/g,'\n\n').trim().slice(0,MAX_TEXT);
const decodeXml=s=>String(s).replace(/&#x([\da-f]+);/gi,(_,n)=>String.fromCodePoint(parseInt(n,16))).replace(/&#(\d+);/g,(_,n)=>String.fromCodePoint(Number(n))).replace(/&(amp|lt|gt|quot|apos|nbsp);/g,(_,k)=>({amp:'&',lt:'<',gt:'>',quot:'"',apos:"'",nbsp:' '}[k]));
function zipEntries(file, names){
 const entries=new Map(), min=Math.max(0,file.length-65558);let eocd=-1;
 for(let i=file.length-22;i>=min;i--)if(file.readUInt32LE(i)===0x06054b50){eocd=i;break;}
 if(eocd<0)throw new Error('Invalid Office document (ZIP directory missing).');
 const n=Math.min(file.readUInt16LE(eocd+10),2000),central=file.readUInt32LE(eocd+16);let i=central;
 for(let count=0;count<n;count++){
  if(i+46>file.length||file.readUInt32LE(i)!==0x02014b50)throw new Error('Invalid Office document archive.');
  const method=file.readUInt16LE(i+10),packed=file.readUInt32LE(i+20),size=file.readUInt32LE(i+24),length=file.readUInt16LE(i+28),extra=file.readUInt16LE(i+30),comment=file.readUInt16LE(i+32),offset=file.readUInt32LE(i+42);const name=file.subarray(i+46,i+46+length).toString('utf8');
  i+=46+length+extra+comment;
  if(!names(name))continue;
  if(size>6*1024*1024||packed>6*1024*1024)throw new Error('Document XML is too large to index.');
  if(offset+30>file.length||file.readUInt32LE(offset)!==0x04034b50)throw new Error('Invalid Office document entry.');
  const start=offset+30+file.readUInt16LE(offset+26)+file.readUInt16LE(offset+28);
  if(start+packed>file.length)throw new Error('Invalid Office document entry size.');
  const raw=file.subarray(start,start+packed);
  const data=method===8?zlib.inflateRawSync(raw,{maxOutputLength:6*1024*1024}):method===0?raw:null;
  if(!data)throw new Error('Unsupported Office document compression.');
  entries.set(name,data.toString('utf8'));
 }
 return entries;
}
function docxText(file){
 const entries=zipEntries(file,n=>/^word\/(document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/.test(n));
 const raw=[...entries.entries()].sort(([a],[b])=>a==='word/document.xml'?-1:b==='word/document.xml'?1:a.localeCompare(b));
 return clean(raw.map(([,xml])=>decodeXml(xml.replace(/<w:tab\b[^>]*\/?\s*>/g,'\t').replace(/<w:br\b[^>]*\/?\s*>/g,'\n').replace(/<\/w:p\s*>/g,'\n').replace(/<\/w:tr\s*>/g,'\n').replace(/<\/w:tc\s*>/g,'\t').replace(/<[^>]*>/g,''))).join('\n'));
}
function xlsxText(file){
 const entries=zipEntries(file,n=>n==='xl/sharedStrings.xml'||/^xl\/worksheets\/sheet\d+\.xml$/.test(n));
 const sharedRaw=entries.get('xl/sharedStrings.xml')||'';
 const shared=[...sharedRaw.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map(m=>decodeXml([...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(x=>x[1]).join('')));
 const rows=[];
 for(const [name,xml] of entries){if(!name.includes('worksheets'))continue;rows.push(name.replace(/^xl\/worksheets\//,'').toUpperCase());
  for(const r of xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)){
   const cols=[];for(const c of r[1].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)){
    const type=/\bt="([^"]+)"/.exec(c[1])?.[1];const v=/<v\b[^>]*>([\s\S]*?)<\/v>/.exec(c[2])?.[1]||'';
    const inline=[...c[2].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(t=>t[1]).join('');
    const data=type==='s'?shared[Number(v)]||'':inline||v; if(data)cols.push(decodeXml(data));
   }if(cols.length)rows.push(cols.join(' | '));
   if(rows.join(' ').length>MAX_TEXT)break;
  }
 }
 return clean(rows.join('\n'));
}
function stripHtml(s){return clean(decodeXml(String(s).replace(/<(script|style)\b[^>]*>[\s\S]*?<\/\1>/gi,' ').replace(/<\/?(?:p|br|div|li|tr|h\d)\b[^>]*>/gi,'\n').replace(/<[^>]+>/g,' ')));}
function decodePdfTextFallback(buffer){
 // Limited emergency fallback for simple PDF files if PDF.js isn't installed.
 // A status warning is surfaced; never misrepresent it as complete indexing.
 const latin=buffer.toString('latin1');const chunks=[];
 for(const m of latin.matchAll(/stream\r?\n([\s\S]*?)\r?\nendstream/g)){
  let stream=Buffer.from(m[1],'latin1');const dict=latin.slice(Math.max(0,m.index-700),m.index);
  if(/\/FlateDecode/.test(dict)){try{stream=zlib.inflateSync(stream,{maxOutputLength:8*1024*1024});}catch{continue;}}
  const str=stream.toString('latin1');for(const bt of str.matchAll(/BT([\s\S]*?)ET/g)){
   for(const t of bt[1].matchAll(/\((?:\\.|[^\\)])*\)|<[0-9A-Fa-f\s]+>/g)){
    let v=t[0];if(v.startsWith('('))v=v.slice(1,-1).replace(/\\([nrtbf()\\])/g,(_,c)=>({n:'\n',r:'\r',t:'\t',b:'',f:'','(':'(',')':')','\\':'\\'}[c])).replace(/\\(\d{1,3})/g,(_,oct)=>String.fromCharCode(parseInt(oct,8)));
    else v=Buffer.from(v.slice(1,-1).replace(/\s/g,''),'hex').toString('latin1');
    chunks.push(v);
   }
  }
  if(chunks.join('').length>MAX_TEXT)break;
 }
 const result=clean(chunks.join(' '));const controls=(result.match(/[\x01-\x08\x0b\x0c\x0e-\x1f]/g)||[]).length;return controls>result.length*.02?'':result;
}
async function pdfText(buffer){
 try{
  const pdfjs=await import('pdfjs-dist/legacy/build/pdf.mjs');
  const loading=pdfjs.getDocument({data:new Uint8Array(buffer),useSystemFonts:true,disableFontFace:true,isEvalSupported:false});
  const pdf=await loading.promise;let str='';const pages=Math.min(pdf.numPages,100);
  for(let i=1;i<=pages;i++){
   const page=await pdf.getPage(i),c=await page.getTextContent();str+=`\nPage ${i}\n`;
   let lastY=null;for(const item of c.items){if(typeof item.str!=='string')continue;const y=item.transform?.[5];if(lastY!==null&&Number.isFinite(y)&&Math.abs(y-lastY)>5)str+='\n';str+=item.str+(item.hasEOL?'\n':' ');lastY=y;if(str.length>MAX_TEXT)break;}
   page.cleanup();if(str.length>MAX_TEXT)break;
  }
  await pdf.destroy();return {text:clean(str),notice:pdf.numPages>pages?'Only the first 100 pages were indexed.':''};
 }catch(err){const fallback=decodePdfTextFallback(buffer);return {text:fallback,notice:`PDF text extraction was limited (${String(err.message||'PDF reader unavailable').slice(0,120)}). Add useful keywords or a process summary; scanned/image-only PDFs require OCR.`};}
}
async function extract(file,ext){
 if(['.txt','.md','.csv','.json','.log','.xml'].includes(ext))return {text:clean(file.toString('utf8')),notice:''};
 if(['.html','.htm'].includes(ext))return {text:stripHtml(file.toString('utf8')),notice:''};
 if(ext==='.docx')return {text:docxText(file),notice:''};
 if(ext==='.xlsx')return {text:xlsxText(file),notice:''};
 if(ext==='.pdf')return pdfText(file);
 return {text:'',notice:'This file type is stored, but its contents cannot be indexed. Use title, keywords and notes to make it searchable.'};
}
module.exports={extract,docxText,xlsxText,decodePdfTextFallback};

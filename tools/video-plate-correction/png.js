// rgb24 frame -> PNG. mode "plate" recolours so near-white deviation is visible:
//   pure #FFFFFF -> white, 250..254 -> red, 235..249 -> orange, darker -> grey.
const {spawnSync}=require('child_process'); const zlib=require('zlib'); const fs=require('fs');
const FF='C:/Users/sharons/AppData/Local/CapCut/Apps/9.3.0.3970/ffmpeg.exe';
function crc32(b){let c,t=[];for(let n=0;n<256;n++){c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c>>>0;}
 let x=0xFFFFFFFF;for(let i=0;i<b.length;i++)x=t[(x^b[i])&0xFF]^(x>>>8);return (x^0xFFFFFFFF)>>>0;}
function chunk(type,data){const len=Buffer.alloc(4);len.writeUInt32BE(data.length);
 const td=Buffer.concat([Buffer.from(type,'ascii'),data]);const crc=Buffer.alloc(4);crc.writeUInt32BE(crc32(td));
 return Buffer.concat([len,td,crc]);}
function png(w,h,rgb){const raw=Buffer.alloc(h*(w*3+1));
 for(let y=0;y<h;y++){raw[y*(w*3+1)]=0;rgb.copy(raw,y*(w*3+1)+1,y*w*3,(y+1)*w*3);}
 const ihdr=Buffer.alloc(13);ihdr.writeUInt32BE(w,0);ihdr.writeUInt32BE(h,4);ihdr[8]=8;ihdr[9]=2;
 return Buffer.concat([Buffer.from([137,80,78,71,13,10,26,10]),chunk('IHDR',ihdr),
   chunk('IDAT',zlib.deflateSync(raw,{level:9})),chunk('IEND',Buffer.alloc(0))]);}
const [,,file,frameNo,outPng,mode]=process.argv;
const s=spawnSync(FF,['-hide_banner','-i',file,'-f','null','-'],{encoding:'buffer',maxBuffer:1<<26}).stderr.toString('latin1');
const m=s.match(/Video:.*?, (\d+)x(\d+)/); const w=+m[1],h=+m[2];
const n=+frameNo;
const r=spawnSync(FF,['-hide_banner','-loglevel','error','-i',file,'-frames:v',String(n+1),
  '-f','rawvideo','-pix_fmt','rgb24','-vsync','0','-'],{maxBuffer:1<<29});
let buf=r.stdout.subarray(n*w*h*3,(n+1)*w*h*3);
if(buf.length<w*h*3) throw new Error('frame '+n+' not decoded; got '+r.stdout.length+' bytes');
if(mode==='plate'){
  const o=Buffer.alloc(w*h*3);
  for(let i=0;i<w*h;i++){const R=buf[i*3],G=buf[i*3+1],B=buf[i*3+2];const mn=Math.min(R,G,B);
   let c;
   if(R===255&&G===255&&B===255) c=[255,255,255];
   else if(mn>=250) c=[255,0,0];
   else if(mn>=235) c=[255,165,0];
   else c=[205,205,205];
   o[i*3]=c[0];o[i*3+1]=c[1];o[i*3+2]=c[2];}
  buf=o;}
fs.writeFileSync(outPng,png(w,h,buf));
console.log('wrote',outPng,w+'x'+h);

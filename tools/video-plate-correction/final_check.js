// final_check.js <original> <corrected>
// Plate-only purity: rebuild, from the ORIGINAL, the set of pixels that are
// unambiguously background (edge-connected, chroma-neutral, within 3 luma of
// the border plate), erode it 3px so artwork edges cannot contaminate it, then
// report what that same set looks like in the corrected encode.
const {spawnSync,spawn}=require('child_process');
const FF='C:/Users/sharons/AppData/Local/CapCut/Apps/9.3.0.3970/ffmpeg.exe';
function dims(f){const s=spawnSync(FF,['-hide_banner','-i',f,'-f','null','-'],{encoding:'buffer',maxBuffer:1<<26}).stderr.toString('latin1');
 const m=s.match(/Video:.*?, (\d+)x(\d+)/);return {w:+m[1],h:+m[2]};}
function stream(file,fmt,bytes,onFrame){return new Promise(res=>{
 const ff=spawn(FF,['-hide_banner','-loglevel','error','-i',file,'-f','rawvideo','-pix_fmt',fmt,'-vsync','0','-'],{stdio:['ignore','pipe','inherit']});
 let buf=Buffer.alloc(0);const q=[];const out=[];
 ff.stdout.on('data',d=>{q.push(d);if(q.reduce((a,c)=>a+c.length,0)>=bytes){buf=Buffer.concat([buf,...q.splice(0)]);
  while(buf.length>=bytes){out.push(Buffer.from(buf.subarray(0,bytes)));buf=buf.subarray(bytes);}}});
 ff.on('close',()=>res(out));});}
(async()=>{
 const [,,orig,corr]=process.argv;
 const {w,h}=dims(orig), yS=w*h, cW=w>>1, cS=cW*(h>>1);
 const Y=await stream(orig,'yuv420p',yS+2*cS);
 const C=await stream(corr,'rgb24',w*h*3);
 const n=Math.min(Y.length,C.length);
 let tot=0,not255=0,sum=0,min=255; const perFrame=[];
 const mask=new Uint8Array(yS), er=new Uint8Array(yS), queue=new Int32Array(yS);
 for(let f=0;f<n;f++){
  const y=Y[f], c=C[f];
  const bd=[];for(let x=0;x<w;x++){bd.push(y[x]);bd.push(y[(h-1)*w+x]);}
  for(let j=0;j<h;j++){bd.push(y[j*w]);bd.push(y[j*w+w-1]);}
  bd.sort((a,b)=>a-b); const seed=bd[bd.length>>1], floor=seed-3;
  mask.fill(0); let qh=0,qt=0;
  const push=(x,yy)=>{const i=yy*w+x;if(mask[i])return;if(y[i]<floor)return;
   const ci=(yy>>1)*cW+(x>>1);
   if(Math.abs(y[yS+ci]-128)>2||Math.abs(y[yS+cS+ci]-128)>2)return;
   mask[i]=1;queue[qt++]=i;};
  for(let x=0;x<w;x++){push(x,0);push(x,h-1);}
  for(let j=0;j<h;j++){push(0,j);push(w-1,j);}
  while(qh<qt){const i=queue[qh++],x=i%w,yy=(i/w)|0;
   if(x>0)push(x-1,yy);if(x<w-1)push(x+1,yy);if(yy>0)push(x,yy-1);if(yy<h-1)push(x,yy+1);}
  // erode 3px: keep only pixels whose 7x7 neighbourhood is entirely mask
  er.fill(0);
  for(let yy=3;yy<h-3;yy++)for(let x=3;x<w-3;x++){
   let ok=1;for(let dy=-3;dy<=3&&ok;dy++)for(let dx=-3;dx<=3;dx++)if(!mask[(yy+dy)*w+x+dx]){ok=0;break;}
   er[yy*w+x]=ok;}
  let ft=0,fs=0;
  for(let i=0;i<yS;i++){ if(!er[i])continue;
   const r=c[i*3],g=c[i*3+1],b=c[i*3+2];const mn=Math.min(r,g,b);
   tot++;ft++;sum+=mn;fs+=mn;if(mn<min)min=mn;if(!(r===255&&g===255&&b===255))not255++;}
  if(ft)perFrame.push(fs/ft);
 }
 const spread=Math.max(...perFrame)-Math.min(...perFrame);
 let maxJump=0;for(let i=1;i<perFrame.length;i++)maxJump=Math.max(maxJump,Math.abs(perFrame[i]-perFrame[i-1]));
 console.log(JSON.stringify({frames:n,platePx:tot,
  notPureWhite:+(100*not255/tot).toFixed(4)+'%',
  plateMean:+(sum/tot).toFixed(4), plateMin:min,
  flickerSpread:+spread.toFixed(4), maxFrameJump:+maxJump.toFixed(4)}));
})();

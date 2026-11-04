// fidelity.js <original> <candidate> [<control>]
// Mean abs delta over CHARACTER pixels only (original min-channel < 235),
// so background lift is excluded and only artwork damage is measured.
const {spawnSync,spawn}=require('child_process');
const FF='C:/Users/sharons/AppData/Local/CapCut/Apps/9.3.0.3970/ffmpeg.exe';
function dims(f){const s=spawnSync(FF,['-hide_banner','-i',f,'-f','null','-'],{encoding:'buffer',maxBuffer:1<<26}).stderr.toString('latin1');
 const m=s.match(/Video:.*?, (\d+)x(\d+)/);return {w:+m[1],h:+m[2]};}
function frames(file,w,h,cb){
 return new Promise(res=>{const fb=w*h*3;let buf=Buffer.alloc(0);const q=[];
  const ff=spawn(FF,['-hide_banner','-loglevel','error','-i',file,'-f','rawvideo','-pix_fmt','rgb24','-vsync','0','-'],{stdio:['ignore','pipe','inherit']});
  const out=[];ff.stdout.on('data',d=>{q.push(d);if(q.reduce((a,c)=>a+c.length,0)>=fb){buf=Buffer.concat([buf,...q.splice(0)]);
   while(buf.length>=fb){out.push(Buffer.from(buf.subarray(0,fb)));buf=buf.subarray(fb);}}});
  ff.on('close',()=>res(out));});}
(async()=>{
 const [,,orig,cand,ctrl]=process.argv;
 const {w,h}=dims(orig);
 const A=await frames(orig,w,h), B=await frames(cand,w,h);
 const C=ctrl?await frames(ctrl,w,h):null;
 const n=Math.min(A.length,B.length,C?C.length:1e9);
 function cmp(X){let sum=0,cnt=0,gt8=0;
  for(let f=0;f<n;f++){const a=A[f],x=X[f];
   for(let i=0;i<w*h;i++){const r=a[i*3],g=a[i*3+1],b=a[i*3+2];
    if(Math.min(r,g,b)>=235) continue;            // background / near-white: not character
    const d=(Math.abs(r-x[i*3])+Math.abs(g-x[i*3+1])+Math.abs(b-x[i*3+2]))/3;
    sum+=d;cnt++;if(d>8)gt8++;}}
  return {mean:sum/cnt,gt8pct:100*gt8/cnt,cnt};}
 const cb=cmp(B);
 const cc=C?cmp(C):null;
 const o={frames:n,charPx:cb.cnt,corrected:+cb.mean.toFixed(4),correctedGt8:+cb.gt8pct.toFixed(4)};
 if(cc){o.control=+cc.mean.toFixed(4);o.controlGt8=+cc.gt8pct.toFixed(4);
        o.correctionOwnShare=+(cb.mean-cc.mean).toFixed(4);}
 console.log(JSON.stringify(o));
})();

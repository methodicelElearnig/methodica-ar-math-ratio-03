// Character-pixel delta measured AT DISPLAY SIZE (the clip is shown at ~50%),
// which is what a viewer actually sees.
const {spawnSync,spawn}=require('child_process');
const FF='C:/Users/sharons/AppData/Local/CapCut/Apps/9.3.0.3970/ffmpeg.exe';
function frames(file,w,h){return new Promise(res=>{const fb=w*h*3;let buf=Buffer.alloc(0);const q=[];
 const ff=spawn(FF,['-hide_banner','-loglevel','error','-i',file,'-vf',`scale=${w}:${h}:flags=lanczos`,
  '-f','rawvideo','-pix_fmt','rgb24','-vsync','0','-'],{stdio:['ignore','pipe','inherit']});
 const out=[];ff.stdout.on('data',d=>{q.push(d);if(q.reduce((a,c)=>a+c.length,0)>=fb){buf=Buffer.concat([buf,...q.splice(0)]);
  while(buf.length>=fb){out.push(Buffer.from(buf.subarray(0,fb)));buf=buf.subarray(fb);}}});
 ff.on('close',()=>res(out));});}
(async()=>{
 const [,,ref,w,h,...cands]=process.argv;
 const W=+w,H=+h;
 const A=await frames(ref,W,H);
 for(const c of cands){
  const B=await frames(c,W,H);
  const n=Math.min(A.length,B.length);
  let sum=0,cnt=0,gt8=0,gt4=0;
  for(let f=0;f<n;f++){const a=A[f],b=B[f];
   for(let i=0;i<W*H;i++){const r=a[i*3],g=a[i*3+1],bl=a[i*3+2];
    if(Math.min(r,g,bl)>=250) continue;
    const d=(Math.abs(r-b[i*3])+Math.abs(g-b[i*3+1])+Math.abs(bl-b[i*3+2]))/3;
    sum+=d;cnt++;if(d>8)gt8++;if(d>4)gt4++;}}
  console.log(c.split('.').slice(-2)[0].padEnd(16),
   'meanDelta',(sum/cnt).toFixed(4).padStart(8),
   ' >4:',(100*gt4/cnt).toFixed(3)+'%', ' >8:',(100*gt8/cnt).toFixed(3)+'%');
 }
})();

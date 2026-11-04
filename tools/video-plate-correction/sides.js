// Per-side border breakdown + longest contiguous non-255 run, to tell a
// uniform plate edge (long runs on every side) from artwork touching the frame.
const { spawnSync, spawn } = require('child_process');
const FF = 'C:/Users/sharons/AppData/Local/CapCut/Apps/9.3.0.3970/ffmpeg.exe';
function probe(file){
  const r=spawnSync(FF,['-hide_banner','-i',file,'-f','null','-'],{encoding:'buffer',maxBuffer:1<<26});
  const m=r.stderr.toString('latin1').match(/Video:.*?, (\d+)x(\d+)/);
  return {w:+m[1],h:+m[2]};
}
async function run(file){
  const {w,h}=probe(file), fb=w*h*3;
  const sides={top:[],bottom:[],left:[],right:[]};
  const acc={top:{n:0,bad:0,maxRun:0,min:255},bottom:{n:0,bad:0,maxRun:0,min:255},
             left:{n:0,bad:0,maxRun:0,min:255},right:{n:0,bad:0,maxRun:0,min:255}};
  const ff=spawn(FF,['-hide_banner','-loglevel','error','-i',file,'-f','rawvideo','-pix_fmt','rgb24','-vsync','0','-'],{stdio:['ignore','pipe','inherit']});
  let buf=Buffer.alloc(0); const q=[];
  await new Promise(res=>{
    ff.stdout.on('data',d=>{ q.push(d);
      if(q.reduce((a,c)=>a+c.length,0)>=fb){ buf=Buffer.concat([buf,...q.splice(0)]);
        while(buf.length>=fb){ frame(buf.subarray(0,fb)); buf=buf.subarray(fb);} } });
    ff.on('close',()=>res());
  });
  function line(f,pts,a){
    let run=0;
    for(const p of pts){ const r=f[p*3],g=f[p*3+1],b=f[p*3+2]; const mn=Math.min(r,g,b);
      a.n++; if(mn<a.min)a.min=mn;
      if(!(r===255&&g===255&&b===255)){ a.bad++; run++; if(run>a.maxRun)a.maxRun=run; } else run=0; }
  }
  function frame(f){
    const T=[],B=[],L=[],R=[];
    for(let x=0;x<w;x++){T.push(x); B.push((h-1)*w+x);}
    for(let y=0;y<h;y++){L.push(y*w); R.push(y*w+w-1);}
    line(f,T,acc.top); line(f,B,acc.bottom); line(f,L,acc.left); line(f,R,acc.right);
  }
  const pct=(a,b)=>(100*a/b).toFixed(2)+'%';
  console.log(file.replace(/.*03-(0\d).assets.video./,'$1/')+'  '+w+'x'+h);
  for(const k of ['top','bottom','left','right']){
    const a=acc[k];
    console.log(`   ${k.padEnd(7)} not255 ${pct(a.bad,a.n).padStart(7)}  maxRun ${String(a.maxRun).padStart(5)} / ${k==='top'||k==='bottom'?w:h}  min ${a.min}`);
  }
}
(async()=>{ for(const f of process.argv.slice(2)) await run(f); })();

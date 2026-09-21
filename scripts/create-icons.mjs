import fs from 'node:fs';
import zlib from 'node:zlib';
function crc(data){let c=0xffffffff;for(const b of data){c^=b;for(let n=0;n<8;n++)c=(c>>>1)^((c&1)?0xedb88320:0)}return(c^0xffffffff)>>>0}
function chunk(type,data){const name=Buffer.from(type),size=Buffer.alloc(4),checksum=Buffer.alloc(4);size.writeUInt32BE(data.length);checksum.writeUInt32BE(crc(Buffer.concat([name,data])));return Buffer.concat([size,name,data,checksum])}
function png(size,pixel){const raw=Buffer.alloc((size*4+1)*size);for(let y=0;y<size;y++)for(let x=0;x<size;x++){const color=pixel((x+.5)/size,(y+.5)/size);const i=y*(size*4+1)+1+x*4;for(let c=0;c<4;c++)raw[i+c]=color[c]}const header=Buffer.alloc(13);header.writeUInt32BE(size);header.writeUInt32BE(size,4);header[8]=8;header[9]=6;return Buffer.concat([Buffer.from([137,80,78,71,13,10,26,10]),chunk('IHDR',header),chunk('IDAT',zlib.deflateSync(raw)),chunk('IEND',Buffer.alloc(0))])}
function distanceLine(x,y,x1,y1,x2,y2){const dx=x2-x1,dy=y2-y1,t=Math.max(0,Math.min(1,((x-x1)*dx+(y-y1)*dy)/(dx*dx+dy*dy)));return Math.hypot(x-x1-t*dx,y-y1-t*dy)}
const lines=[[.25,.68,.25,.4],[.25,.4,.37,.4],[.37,.4,.37,.62],[.37,.53,.63,.53],[.63,.62,.63,.4],[.63,.4,.75,.4],[.75,.4,.75,.68]];
function mark(x,y){return Math.min(...lines.map(l=>distanceLine(x,y,...l)))}
fs.writeFileSync('build/icon.png',png(1024,(x,y)=>{const d=Math.hypot(Math.max(Math.abs(x-.5)-.31,0),Math.max(Math.abs(y-.5)-.31,0))-.14;if(d>0)return[0,0,0,0];const ink=Math.max(0,Math.min(1,(.028-mark(x,y))*1024+.5));return[70+(255-70)*ink,103+(255-103)*ink,207+(255-207)*ink,255]}));
fs.writeFileSync('resources/trayTemplate.png',png(44,(x,y)=>[0,0,0,Math.max(0,Math.min(1,(.044-mark(x,y))*44+.5))*255]));

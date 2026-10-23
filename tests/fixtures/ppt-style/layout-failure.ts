/** Executes the real host function against a fake presentation; no live Office access. */
import fs from 'node:fs';
import vm from 'node:vm';
export function layoutFailure() {
  const setup:any={SlideWidth:900,SlideSize:1,SlideOrientation:1};
  Object.defineProperty(setup,'SlideHeight',{get:()=>600,set:()=>{throw Error('permission denied: token=secret private-document');}});
  const pres={Name:'synthetic.pptx',PageSetup:setup,Slides:{Count:0},Designs:{Count:1},SlideMaster:{CustomLayouts:{Count:0}}};
  const ctx:any=vm.createContext({getPptPresentation:()=>pres});
  vm.runInContext(fs.readFileSync('wps-addon/src/ppt.js','utf8'),ctx);
  return JSON.parse(JSON.stringify(ctx.configureSlideLayout({}, {slideWidth:1000,slideHeight:700})));
}

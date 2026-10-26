import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { EventEmitter } from 'node:events';
import ts from 'typescript';

// Execute the actual helper's launch function with process creation intercepted;
// no installed app or desktop is touched. Windows visibility still needs real UI acceptance.
const source = ts.createSourceFile('update-helper.ts', fs.readFileSync(new URL('../src/main/update-helper.ts', import.meta.url), 'utf8'), ts.ScriptTarget.Latest, true);
const declaration = source.statements.find(n => ts.isFunctionDeclaration(n) && n.name?.text === 'launch');
assert.ok(declaration);
const code = ts.transpileModule(declaration.getText(source), { compilerOptions: { target: ts.ScriptTarget.ES2022 } }).outputText;

test('updated Windows GUI and old-version recovery launch visibly without a shell', async () => {
  for (const target of ['C:\\Users\\fixture\\Programs\\CharFloat\\2.2.911-aaaaaaaaaaaa', 'C:\\Users\\fixture\\Programs\\CharFloat\\2.2.910-bbbbbbbbbbbb']) {
    let args: any[], options: any, executable: string, unref = false;
    const launch = vm.runInNewContext(code + '\nlaunch', { windows: {}, path: path.win32, process: { env: { ELECTRON_RUN_AS_NODE:'1', WPS_BRIDGE_RESOURCES:'old', KEEP:'yes' } }, spawn: (exe: string, argv: any[], opts: any) => {
      executable=exe;args=argv;options=opts;const child:any=new EventEmitter();child.unref=()=>{unref=true;};queueMicrotask(()=>child.emit('spawn'));return child;
    } });
    await launch(target);
    assert.equal(executable!, path.win32.join(target,'字浮 CharFloat.exe'));
    assert.equal(args!.length,0);assert.notEqual(options.windowsHide,true);assert.notEqual(options.shell,true);
    assert.equal(options.env.ELECTRON_RUN_AS_NODE,undefined);assert.equal(options.env.WPS_BRIDGE_RESOURCES,undefined);assert.equal(options.env.KEEP,'yes');assert.equal(unref,true);
  }
});

test('Mac GUI restart continues through LaunchServices and reports a launch failure', async () => {
  for (const exitCode of [0,1]) {
    let executable='', argument='';
    const launch=vm.runInNewContext(code+'\nlaunch',{windows:undefined,path:path.posix,process:{env:{}},spawn:(exe:string,args:string[])=>{
      executable=exe;argument=args[0];const child:any=new EventEmitter();child.unref=()=>{};queueMicrotask(()=>child.emit('exit',exitCode));return child;
    }});
    if(exitCode)await assert.rejects(launch('/fixture/CharFloat.app'),/未能重新打开/);else await launch('/fixture/CharFloat.app');
    assert.equal(executable,'/usr/bin/open');assert.equal(argument,'/fixture/CharFloat.app');
  }
});

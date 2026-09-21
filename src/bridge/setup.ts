#!/usr/bin/env node
import { AddonInstaller } from '../main/addon-installer.js';
import { InstallerEngine } from '../main/installer-engine.js';
const args = process.argv.slice(2);
if (!args.length) {
  console.log(JSON.stringify(InstallerEngine.detectEnvironment(), null, 2));
  console.log('默认仅检查。部署加载项：npm run setup -- --addon；配置指定客户端：--agent=cursor --skills');
} else {
  const options = { addon: args.includes('--addon'), skills: args.includes('--skills'), agents: args.filter(a => a.startsWith('--agent=')).map(a => a.slice(8)) };
  console.log(JSON.stringify(await InstallerEngine.executeInstall(options), null, 2));
}

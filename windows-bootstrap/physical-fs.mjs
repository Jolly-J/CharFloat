// 安装载荷是物理文件。Electron 默认 fs 会把 app.asar 当虚拟档案，不能用于解压/校验。
import nodeFs from 'node:fs';
import { createRequire } from 'node:module';

const fs = process.versions.electron ? createRequire(process.execPath)('original-fs') : nodeFs;
export default fs;

/** Bounded, in-memory PPTX editing. Never extracts paths or overwrites an open document. */
import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { deflateRawSync } from 'node:zlib';
import yauzl from 'yauzl';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MAX = 256 * 1024 * 1024;
export type Parts = Map<string, Buffer>;
export const digest = (b: Buffer | string) => createHash('sha256').update(b).digest('hex');
export function crc32(data: Buffer) {
  let c = 0xffffffff;
  for (const b of data) { c ^= b; for (let k = 0; k < 8; k++) c = (c >>> 1) ^ ((c & 1) ? 0xedb88320 : 0); }
  return (c ^ 0xffffffff) >>> 0;
}
export async function readPptx(file: string): Promise<Parts> {
  if (fs.statSync(file).size > MAX) throw Error('PPTX 超过 256 MiB 限制');
  const zip = await new Promise<yauzl.ZipFile>((ok, no) => yauzl.open(file, { lazyEntries: true, validateEntrySizes: true }, (e, z) => e ? no(e) : ok(z!)));
  return new Promise((ok, no) => {
    const parts: Parts = new Map(); let total = 0, finished = false;
    const fail = (e: unknown) => { if (!finished) { finished = true; zip.close(); no(e); } };
    zip.on('error', fail);
    zip.on('end', () => { if (!finished) { finished = true; ok(parts); } });
    zip.on('entry', (entry: yauzl.Entry) => {
      const n = entry.fileName;
      if (parts.size >= 10000 || parts.has(n) || n.startsWith('/') || n.includes('\\') || n.split('/').some(x => x === '..' || x === '.') || entry.generalPurposeBitFlag & 1 || ((entry.externalFileAttributes >>> 16) & 0xf000) === 0xa000) return fail(Error('不支持的 ZIP 路径、重复项、加密或链接'));
      total += entry.uncompressedSize;
      if (total > MAX) return fail(Error('PPTX 展开量超过 256 MiB'));
      zip.openReadStream(entry, (e, stream) => {
        if (e) return fail(e);
        const chunks: Buffer[] = []; let size = 0;
        stream!.on('error', fail);
        stream!.on('data', (b: Buffer) => { size += b.length; if (size > entry.uncompressedSize) { stream!.destroy(); fail(Error('ZIP 项大小不符')); } else chunks.push(b); });
        stream!.on('end', () => {
          if (finished) return;
          const data = Buffer.concat(chunks);
          if (data.length !== entry.uncompressedSize || crc32(data) !== entry.crc32) return fail(Error('ZIP 项校验失败'));
          parts.set(n, data); zip.readEntry();
        });
      });
    });
    zip.readEntry();
  });
}
/** Small ZIP writer, bounded by readPptx; all unrelated uncompressed bytes stay identical. */
export function writePptx(parts: Parts, file: string) {
  const locals: Buffer[] = [], central: Buffer[] = []; let offset = 0;
  for (const [name, data] of parts) {
    const n = Buffer.from(name), compressed = deflateRawSync(data), crc = crc32(data);
    const l = Buffer.alloc(30), c = Buffer.alloc(46);
    l.writeUInt32LE(0x04034b50); l.writeUInt16LE(20, 4); l.writeUInt16LE(0x800, 6); l.writeUInt16LE(8, 8);
    l.writeUInt32LE(crc, 14); l.writeUInt32LE(compressed.length, 18); l.writeUInt32LE(data.length, 22); l.writeUInt16LE(n.length, 26);
    c.writeUInt32LE(0x02014b50); c.writeUInt16LE(20, 4); c.writeUInt16LE(20, 6); c.writeUInt16LE(0x800, 8); c.writeUInt16LE(8, 10);
    c.writeUInt32LE(crc, 16); c.writeUInt32LE(compressed.length, 20); c.writeUInt32LE(data.length, 24); c.writeUInt16LE(n.length, 28); c.writeUInt32LE(offset, 42);
    locals.push(l, n, compressed); central.push(c, n); offset += l.length + n.length + compressed.length;
  }
  const cd = Buffer.concat(central), end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50); end.writeUInt16LE(parts.size, 8); end.writeUInt16LE(parts.size, 10); end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(offset, 16);
  fs.writeFileSync(file, Buffer.concat([...locals, cd, end]), { flag: 'wx', mode: 0o600 });
}
export function xml(data: Buffer | string) {
  const text = data.toString();
  if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw Error('拒绝带 DTD/实体的 XML');
  return new DOMParser({ onError: () => { throw Error('无效的 PPTX XML'); } }).parseFromString(text, 'application/xml');
}
const children = (n: any) => Array.from(n?.childNodes || []).filter((x: any) => x.nodeType === 1) as any[];
const child = (n: any, ns: string, name: string) => children(n).find(x => x.namespaceURI === ns && x.localName === name);
export function canonical(n: any): string {
  if (n.nodeType === 3) return n.nodeValue.trim() ? JSON.stringify(n.nodeValue) : '';
  if (n.nodeType !== 1) return '';
  const attrs = Array.from(n.attributes || []).filter((a: any) => !a.name.startsWith('xmlns')).map((a: any) => [a.namespaceURI || '', a.localName, a.value]).sort();
  return JSON.stringify([n.namespaceURI, n.localName, attrs]) + Array.from(n.childNodes || []).map(canonical).join('') + '/';
}
function relTarget(parts: Parts, base: string, id: string): string {
  const rel = path.posix.join(path.posix.dirname(base), '_rels', path.posix.basename(base) + '.rels');
  const matches = children(xml(parts.get(rel)!).documentElement).filter(x => x.getAttribute('Id') === id);
  if (matches.length !== 1 || matches[0].getAttribute('TargetMode') === 'External') throw Error('幻灯片关系不明确');
  const target = path.posix.normalize(path.posix.join(path.posix.dirname(base), matches[0].getAttribute('Target')));
  if (!target.startsWith('ppt/') || !parts.has(target)) throw Error('幻灯片关系越界或缺失');
  return target;
}
export function inspectRadius(parts: Parts, slideIndex: number, names: string[], radiusPt: number) {
  if (!Number.isInteger(slideIndex) || slideIndex < 1 || !Number.isFinite(radiusPt) || radiusPt < 0 || !names.length || new Set(names).size !== names.length) throw Error('页码、名称或半径无效');
  const pres = xml(parts.get('ppt/presentation.xml')!), ids = Array.from(pres.getElementsByTagNameNS(P, 'sldId'));
  const id = ids[slideIndex - 1]; if (!id) throw Error('页码超出文件范围');
  const part = relTarget(parts, 'ppt/presentation.xml', id.getAttributeNS(R, 'id')!);
  const doc = xml(parts.get(part)!), root = doc.documentElement!, tree = child(child(root, P, 'cSld'), P, 'spTree');
  if (!tree) throw Error('缺少原生形状树');
  // Unknown identity dependencies are rejected conservatively, before importing anything.
  if (Array.from(parts.keys()).some(n => /comment|vbaProject|signature/i.test(n)) || root.getElementsByTagNameNS(P, 'timing').length || root.getElementsByTagNameNS(P, 'bldLst').length || root.getElementsByTagNameNS(P, 'extLst').length) throw Error('暂不支持批注、宏、签名、动画或扩展引用');
  const shapes = children(tree), targets = names.map(name => {
    const matches = shapes.filter(s => child(child(s, P, 'nvSpPr'), P, 'cNvPr')?.getAttribute('name') === name);
    if (matches.length !== 1) throw Error('名称必须唯一匹配一个顶层空白 roundRect：' + name);
    const s = matches[0], nv = child(child(s, P, 'nvSpPr'), P, 'cNvPr'), sp = child(s, P, 'spPr');
    const geom = child(sp, A, 'prstGeom'), xfrm = child(sp, A, 'xfrm'), ext = child(xfrm, A, 'ext');
    if (s.localName !== 'sp' || geom?.getAttribute('prst') !== 'roundRect' || !ext || nv.getAttribute('hidden') === '1' || nv.getAttribute('hidden') === 'true' || children(child(child(s, P, 'nvSpPr'), P, 'nvPr')).length || children(nv).length) throw Error('目标不是支持的可见空白原生背景：' + name);
    if (Array.from(s.getElementsByTagNameNS(A, 't')).some((t:any) => t.textContent?.trim()) || s.getElementsByTagNameNS(A, 'fld').length || s.getElementsByTagNameNS(A, 'hlinkClick').length || s.getElementsByTagNameNS(A, 'hlinkMouseOver').length || s.getElementsByTagNameNS(A, 'blip').length || Number(xfrm.getAttribute('rot') || 0) !== 0) throw Error('暂不支持文字、域、超链接、图片填充或旋转背景：' + name);
    const widthPt = Number(ext.getAttribute('cx')) / 12700, heightPt = Number(ext.getAttribute('cy')) / 12700, short = Math.min(widthPt, heightPt);
    if (!Number.isFinite(short) || short <= 0 || radiusPt > short / 2) throw Error('半径不能超过每个形状短边一半：' + name);
    const shapeId = Number(nv.getAttribute('id'));
    for (const n of Array.from(root.getElementsByTagName('*'))) for (const a of Array.from(n.attributes)) {
      if ((/^(spid|spId)$/.test(a.localName || '') || /^(stCxn|endCxn)$/.test(n.localName || '') && a.localName === 'id') && Number(a.value) === shapeId) throw Error('目标有动画或连接器身份引用：' + name);
    }
    const guides = geom.getElementsByTagNameNS(A, 'gd');
    if (guides.length > 1 || guides.length === 1 && guides[0].getAttribute('name') !== 'adj') throw Error('未知几何调整值');
    const fmla = guides[0]?.getAttribute('fmla') || 'val 16667';
    if (!/^val \d+$/.test(fmla)) throw Error('不支持计算型几何调整值');
    const adjustment = Math.round(radiusPt / short * 100000);
    return { name, shapeId, widthPt, heightPt, radiusPt, adjustment, actualRadiusPt: short * adjustment / 100000, currentRadiusPt: short * Number(fmla.slice(4)) / 100000, fingerprint: digest(canonical(s)), node: s, geom };
  });
  return { part, slideId: Number(id.getAttribute('id')), doc, tree, targets };
}
export function patchRadius(parts: Parts, plan: ReturnType<typeof inspectRadius>) {
  const copy = new Map(parts);
  for (const t of plan.targets) {
    let av = child(t.geom, A, 'avLst');
    if (!av) { av = plan.doc.createElementNS(A, 'a:avLst'); t.geom.appendChild(av); }
    let guide = child(av, A, 'gd');
    if (!guide) { guide = plan.doc.createElementNS(A, 'a:gd'); av.appendChild(guide); guide.setAttribute('name', 'adj'); }
    guide.setAttribute('fmla', 'val ' + t.adjustment);
  }
  copy.set(plan.part, Buffer.from(new XMLSerializer().serializeToString(plan.doc)));
  return copy;
}
export const publicTargets = (plan: ReturnType<typeof inspectRadius>) => plan.targets.map(({ node, geom, fingerprint, ...t }) => t);

/** Verify shape identity changes explicitly, and compare all unrelated parts after host save. */
export function verifyRadius(before: Parts, after: Parts, plan: ReturnType<typeof inspectRadius>, receipts: any[]) {
  const saved = inspectRadius(after, Number(Array.from(xml(before.get('ppt/presentation.xml')!).getElementsByTagNameNS(P, 'sldId')).findIndex(n => Number(n.getAttribute('id')) === plan.slideId)) + 1, plan.targets.map(t => t.name), plan.targets[0].radiusPt);
  if (saved.slideId !== plan.slideId || saved.part !== plan.part || receipts.length !== plan.targets.length) throw Error('保存后页面或替换记录不符');
  for (const t of saved.targets) {
    const old = plan.targets.find(o => o.name === t.name)!, receipt = receipts.find(r => r.name === t.name);
    if (!receipt || t.shapeId !== receipt.newId || Math.abs(t.currentRadiusPt - old.radiusPt) > Math.min(old.widthPt, old.heightPt) / 200000 + 1e-8) throw Error('保存后的原生形状、ID 或半径不符');
    const clone = t.node.cloneNode(true), nv = child(child(clone, P, 'nvSpPr'), P, 'cNvPr');
    nv.setAttribute('id', String(old.shapeId));
    const originalVisibility=child(child(old.node,P,'nvSpPr'),P,'cNvPr');
    if(originalVisibility.hasAttribute('hidden')) nv.setAttribute('hidden',originalVisibility.getAttribute('hidden')); else nv.removeAttribute('hidden');
    child(clone, P, 'spPr').replaceChild(old.geom.cloneNode(true), child(child(clone, P, 'spPr'), A, 'prstGeom'));
    if (canonical(clone) !== canonical(old.node)) throw Error('导入后目标几何或样式改变：' + t.name);
    const archived = children(saved.tree).find(s => child(child(s, P, 'nvSpPr'), P, 'cNvPr')?.getAttribute('id') === String(receipt.oldId));
    if (!archived) throw Error('旧背景归档缺失');
    const av = child(child(archived, P, 'nvSpPr'), P, 'cNvPr');
    if (av.getAttribute('name') !== receipt.oldName || !['1','true'].includes(av.getAttribute('hidden'))) throw Error('旧背景未正确隐藏');
    const archiveClone = archived.cloneNode(true), ac = child(child(archiveClone, P, 'nvSpPr'), P, 'cNvPr');
    ac.setAttribute('name', old.name);
    const originalNv=child(child(old.node,P,'nvSpPr'),P,'cNvPr');
    if(originalNv.hasAttribute('hidden')) ac.setAttribute('hidden',originalNv.getAttribute('hidden')); else ac.removeAttribute('hidden');
    if (canonical(archiveClone) !== canonical(old.node)) throw Error('旧背景归档内容改变');
    saved.tree.replaceChild(clone, t.node);saved.tree.removeChild(archived);
  }
  if (canonical(saved.doc.documentElement) !== canonical(xml(before.get(plan.part)!).documentElement)) throw Error('目标页无关对象或图层顺序改变');
  const normalize = (parts: Parts) => {
    const result: Record<string,string> = {};
    for (const [n,b] of parts) {
      if (n === plan.part || n.startsWith('docProps/') || n.startsWith('ppt/media/') || n.endsWith('/')) continue;
      if (n.endsWith('.rels')) {
        const d = xml(b), base = n.replace(/\/_rels\/(.+)\.rels$/, '/$1');
        for (const r of children(d.documentElement)) if (r.getAttribute('TargetMode') !== 'External') {
          const target = path.posix.normalize(path.posix.join(path.posix.dirname(base),r.getAttribute('Target')));
          if (target.startsWith('ppt/media/') && parts.has(target)) r.setAttribute('Target','media-sha256:'+digest(parts.get(target)!));
        }
        result[n]=canonical(d.documentElement);
      } else result[n] = /\.xml$/.test(n) ? canonical(xml(b).documentElement) : digest(b);
    }
    result['$media']=JSON.stringify([...parts].filter(([n])=>n.startsWith('ppt/media/')).map(([,b])=>digest(b)).sort());
    return JSON.stringify(Object.entries(result).sort());
  };
  if (normalize(before) !== normalize(after)) throw Error('无关页、主题、关系或媒体内容改变');
  return saved.targets.map(({name,shapeId,currentRadiusPt})=>({name,shapeId,actualRadiusPt:currentRadiusPt}));
}

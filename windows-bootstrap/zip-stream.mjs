import { PassThrough } from 'node:stream';

/**
 * yauzl 2's fd-slicer marks stored-entry streams destroyed before emitting end,
 * without emitting close. Native async iterators/pipeline can then wait forever.
 * Bridge its data/end/error events through a native stream; retain backpressure.
 */
export function openZipStream(zip, entry) {
  return new Promise((resolve, reject) => zip.openReadStream(entry, (error, source) => {
    if (error) { reject(error); return; }
    const output = new PassThrough();
    let ended = false;
    source.once('end', () => { ended = true; });
    source.on('error', error => output.destroy(error));
    source.once('close', () => { if (!ended) output.destroy(new Error('ZIP 文件流提前关闭。')); });
    output.once('close', () => {
      source.unpipe(output);
      if (!ended && !source.destroyed) source.destroy();
    });
    source.pipe(output);
    resolve(output);
  }));
}

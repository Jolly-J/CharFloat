import { consumeInstallAuthorization } from '../../src/main/update-install.js';

process.once('message', async (value: { file: string; nonce: string }) => {
  try {
    const guard = await consumeInstallAuthorization(value.file, value.nonce);
    await guard();
    process.send?.({ claimed: true }, () => process.disconnect());
  } catch {
    process.send?.({ claimed: false }, () => process.disconnect());
  }
});
process.send?.({ ready: true });

import type { UpdateTrust } from './update-protocol.js';

// Production public key and owned HTTPS origin. A missing/unpublished feed is reported as an error.
// No GitHub proxy, runtime-supplied executable URL, private key, or assumed live endpoint.
export const UPDATE_TRUST: UpdateTrust = { feeds: ['https://charfloat.online/updates/stable.json'], assetOrigins: ['https://charfloat.online'], publicKeys: {"stable-bc17ce8169c16826":"-----BEGIN PUBLIC KEY-----\nMCowBQYDK2VwAyEAEDcobKCS/atWky4zAcV4D6iPiNqkNJ76zWnyH93kY4A=\n-----END PUBLIC KEY-----\n"} };

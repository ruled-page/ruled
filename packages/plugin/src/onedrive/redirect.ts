/**
 * The page Microsoft's sign-in popup returns to (the connector's
 * `redirectUri`) runs this, and nothing else: it hands the sign-in's answer
 * to the editor's page and closes the popup. Bundle it into a small page of
 * your own, served from your origin without a Cross-Origin-Opener-Policy
 * header.
 */
import { broadcastResponseToMainFrame } from '@azure/msal-browser/redirect-bridge';

void broadcastResponseToMainFrame();

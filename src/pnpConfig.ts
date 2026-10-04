import { spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import { WebPartContext } from '@microsoft/sp-webpart-base';

let _sp: SPFI | undefined;

/**
 * Builds the PnPjs instance once per web part instance. Call with the web part
 * context from onInit, then without arguments anywhere else.
 */
export function getSP(context?: WebPartContext): SPFI {
  if (context !== undefined) {
    _sp = spfi().using(SPFx(context));
  }

  if (_sp === undefined) {
    throw new Error('PnPjs has not been initialised. Call getSP(context) from the web part first.');
  }

  return _sp;
}

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFI } from '@pnp/sp';
import { SPFx } from '@pnp/sp/behaviors/spfx';

let _sp: SPFI | undefined;

export const getSP = (context?: WebPartContext): SPFI => {
  if (context !== undefined) {
    _sp = spfi().using(SPFx(context));
  }

  if (_sp === undefined) {
    throw new Error('PnPjs has not been initialized.');
  }

  return _sp;
};
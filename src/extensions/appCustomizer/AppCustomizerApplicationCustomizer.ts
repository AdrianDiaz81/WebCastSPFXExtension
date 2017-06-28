import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
   BaseApplicationCustomizer,
  Placeholder
} from '@microsoft/sp-application-base';

import * as strings from 'appCustomizerStrings';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './AppCustomizer.module.scss';

const LOG_SOURCE: string = 'AppCustomizerApplicationCustomizer';
const HEADER_TEXT: string = "WEBCAST SUGES -> This is the header zone";
const FOOTER_TEXT: string = "WEBCAST SUGES -> This is the footer zone";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAppCustomizerApplicationCustomizerProperties {
 Header: string;
  Footer: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AppCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<IAppCustomizerApplicationCustomizerProperties> {

  private _headerPlaceholder: Placeholder;
  private _footerPlaceholder: Placeholder;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    return Promise.resolve<void>();
  }

  @override
  public onRender(): void {
      if (!this._headerPlaceholder) {
      this._headerPlaceholder = this.context.placeholders.tryAttach(
        'PageHeader',
        {
          onDispose: this._onDispose
        });

      if (!this._headerPlaceholder) {
        console.error('The expected placeholder (PageHeader) was not found.');
        return;
      }

      if (this._headerPlaceholder.domElement) {
        this._headerPlaceholder.domElement.innerHTML = `
                <div class="${styles.app}">
                  <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.header}">
                    <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i>&nbsp; ${escape(HEADER_TEXT)}
                  </div>
                </div>`;
      }
    }

    if (!this._footerPlaceholder) {
      this._footerPlaceholder = this.context.placeholders.tryAttach(
        'PageFooter',
        {
          onDispose: this._onDispose
        });

      if (!this._footerPlaceholder) {
        console.error('The expected placeholder (PageFooter) was not found.');
        return;
      }

      if (this._footerPlaceholder.domElement) {
        this._footerPlaceholder.domElement.innerHTML = `
                <div class="${styles.app}">
                  <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.footer}">
                    <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i>&nbsp; ${escape(FOOTER_TEXT)}
                  </div>
                </div>`;
      }
    }
  }

  private _onDispose(): void {
    console.log('[CustomHeader._onDispose] Disposed custom header.');
  }
}

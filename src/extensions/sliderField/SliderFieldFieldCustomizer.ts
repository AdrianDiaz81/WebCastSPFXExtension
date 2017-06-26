import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';

import { SPPermission } from "@microsoft/sp-page-context";
import pnp, { List, ItemUpdateResult, Item } from 'sp-pnp-js';
import {
  CellFormatter,
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'sliderFieldStrings';
import SliderField, { ISliderFieldProps } from './components/SliderField';

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISliderFieldProperties {
  // This is an example; replace with your own property
 value?: string;
}

const LOG_SOURCE: string = 'SliderFieldFieldCustomizer';

export default class SliderFieldFieldCustomizer
  extends BaseFieldCustomizer<ISliderFieldProperties> {
 private _timerId: number = -1;

  @override
  public onInit(): Promise<void> {
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(LOG_SOURCE, 'Activated SliderFieldFieldCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(LOG_SOURCE, `The following string should be equal: "SliderField" and "${strings.Title}"`);
    return Promise.resolve<void>();
  }

  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {

    const value: string = event.cellValue;
    const id: string = event.row.getValueByName('ID').toString();
    const hasPermissions: boolean = this.context.pageContext.list.permissions.hasPermission(SPPermission.editListItems);
    
    const sliderField: React.ReactElement<{}> =
      React.createElement(SliderField, { value: value, id: id, disabled: !hasPermissions, onChange: this.onSliderValueChanged.bind(this) } as ISliderFieldProps);
    ReactDOM.render(sliderField, event.cellDiv);
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    ReactDOM.unmountComponentAtNode(event.cellDiv);
    super.onDisposeCell(event);
  }

  private onSliderValueChanged(value: number, id: string): void {
    if (this._timerId !== -1)
      clearTimeout(this._timerId);

    this._timerId = setTimeout(() => {
      let etag: string = undefined;
      pnp.sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(parseInt(id)).get(undefined, {
        headers: {
          'Accept': 'application/json;odata=minimalmetadata'
        }
      })
        .then((item: Item): Promise<any> => {
          etag = item["odata.etag"];
          return Promise.resolve((item as any) as any);
        })
        .then((item: any): Promise<ItemUpdateResult> => {
          let updateObj: any = {};
          updateObj[this.context.field.internalName] = value;
          return pnp.sp.web.lists.getByTitle(this.context.pageContext.list.title)
            .items.getById(parseInt(id)).update(updateObj, etag);
        })
        .then((result: ItemUpdateResult): void => {
          Log.info(LOG_SOURCE,`Item with ID: ${id} successfully updated`);
        }, (error: any): void => {
          Log.error(LOG_SOURCE, error);
        });
    }, 1000);
  }
}

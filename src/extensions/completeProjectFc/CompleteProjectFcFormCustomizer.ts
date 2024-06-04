import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { FormDisplayMode } from '@microsoft/sp-core-library';
import {
  BaseFormCustomizer
} from '@microsoft/sp-listview-extensibility';

import CompleteProjectFc from './components/CompleteProjectFc';
import { IListDataRequest } from './models/IListDataRequest';

import {
  SPHttpClient,
} from '@microsoft/sp-http';
import { initializeSpObject } from './services/SharepointServices';
import { ICompleteProjectFcProps } from './components/ICompleteProjectFcProps';

/**
 * If your form customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICompleteProjectFcFormCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

export default class CompleteProjectFcFormCustomizer
  extends BaseFormCustomizer<ICompleteProjectFcFormCustomizerProperties> {

  // Added for the item to show in the form; use with edit and view form
  private _listItem = {} as IListDataRequest;


  public onInit(): Promise<void> {

    initializeSpObject(this.context);

    if (this.displayMode === FormDisplayMode.New) {
      
      // we're creating a new item so nothing to load
      return Promise.resolve();
    }


    // load item to display on the form
    return this.context.spHttpClient
      .get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${this.context.list.title}')/items(${this.context.itemId})`, SPHttpClient.configurations.v1, {
        headers: {
          accept: 'application/json;odata.metadata=none'
        }
      })
      .then(result => {
        if (result.ok) {
          return result.json();
        }
        else {
          return Promise.reject(result.statusText);
        }
      })
      .then(item => {
        this._listItem = item;
        return Promise.resolve();
      });
  }

  public render(): void {

    // Use this method to perform your custom rendering.
    const completeProjectFc: React.ReactElement<{}> =
      React.createElement(CompleteProjectFc, {
        context: this.context,
        displayMode: this.displayMode,
        listGuid: this.context.list.guid,
        itemID: this.context.itemId,
        listItem: this._listItem,
        businessListGuid: "3799fc11-e017-4ca6-b5e9-5faf1158b362",
        onSave: this._onSave,
        onClose: this._onClose
       } as ICompleteProjectFcProps);

    ReactDOM.render(completeProjectFc, this.domElement);
  }

  public onDispose(): void {

    // This method should be used to free any resources that were allocated during rendering.
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  private _onSave = (): void => {

    // You MUST call this.formSaved() after you save the form.
    this.formSaved();
  }

  private _onClose =  (): void => {

    // You MUST call this.formClosed() after you close the form.
    this.formClosed();
  }
}

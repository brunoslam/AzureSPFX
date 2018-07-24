import * as React from 'react';
import styles from './Sar.module.scss';
import { ISarProps } from './ISarProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';

/* Pivot Office Fabric */
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { PivotItem, IPivotItemProps, Pivot, TextField} from 'office-ui-fabric-react';

import axios, { AxiosRequestConfig, AxiosPromise, AxiosResponse } from 'axios';

export interface estados {
  caca: Array<any>;
}
export default class Sar extends React.Component<ISarProps, estados> {
  
  constructor(){
    super();
    SPComponentLoader.loadCss('//dev.office.com/fabric-js/css/fabric.components.min.css');
    this.state = {
      caca : []
    }
  }
  
  public componentDidMount(){
    this.getElements().then((response) => {
      console.log(response.data);
      this.setState({
        caca : response.data.value
      });
    }).catch((err)=>{
        console.log(err);
    });
  }
  public render(): React.ReactElement<ISarProps> {
    //this.getElements();
    return (
      <div>
        <Pivot>
          <PivotItem linkText="Ver elementos"  itemIcon="Emoji2">
            <Label>Pivot #1</Label>
            <table className="ms-Table">
              <thead>
                <tr>
                  <th>Location</th>
                </tr>
              </thead>
              <tbody>
                {
                  this.state.caca.map((value)=>{

                    return <tr>
                            <td>{value.Nombre}</td>
                          </tr>
                  })
                }
              </tbody>
            </table>
          </PivotItem>
          <PivotItem linkText="Insertar elementos" itemCount={23} itemIcon="Recent">
            <Label>Pivot #2</Label>
          </PivotItem>
          <PivotItem itemIcon="Globe">
            <Label>Pivot #3</Label>
          </PivotItem>
          <PivotItem linkText="Shared with me" itemIcon="Ringer" itemCount={1}>
            <Label>Pivot #4</Label>
          </PivotItem>
          <PivotItem
            linkText="Customized Rendering"
            itemIcon="Globe"
            itemCount={10}
            onRenderItemLink={this._customRenderer}
          >
            <Label>Customized Rendering</Label>
          </PivotItem>
        </Pivot>
      </div>
    );
  }
  private _customRenderer(link: IPivotItemProps, defaultRenderer: (link: IPivotItemProps) => JSX.Element): JSX.Element {
    return (
      <span>
        {defaultRenderer(link)}
        <Icon iconName="Airplane" style={{ color: 'red' }} />
      </span>
    );
  }
  private addElement(){
    const tablestorageUrl =  'https://storagelatintest.table.core.windows.net/Persona?sv=2018-03-28&si=Persona-164CE6B1EC9&tn=persona&sig=8dyeUpUnT%2F%2B9XvpEoEjfqepP2yMv6Uw%2F772kLwZg2UM%3D';
    axios.post(tablestorageUrl, {
      "PartitionKey": "123534",
    "RowKey": "12345",
    "Nombre": "asdasdxasd"
            
    }, {
      url: tablestorageUrl,
      method: 'post',
      headers: {
        'Content-Type': 'application/json'
      }
    }).then(function (response) {
      console.log(response);
    })
    .catch(function (error) {
      console.log(error);
    });
  
    
  }
  private getElements()  {

    this.addElement();

    const tablestorageUrl =  'https://storagelatintest.table.core.windows.net/Persona?sv=2018-03-28&si=Persona-164CE6B1EC9&tn=persona&sig=8dyeUpUnT%2F%2B9XvpEoEjfqepP2yMv6Uw%2F772kLwZg2UM%3D';
    

    
    return axios.get(tablestorageUrl, {
      headers: {
        Accepts: 'application/json'
      }})
  }
}

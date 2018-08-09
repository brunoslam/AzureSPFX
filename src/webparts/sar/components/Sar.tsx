import * as React from 'react';
import styles from './Sar.module.scss';
import { ISarProps } from './ISarProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http';
/* Pivot Office Fabric */
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { PivotItem, IPivotItemProps, Pivot, TextField} from 'office-ui-fabric-react';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import axios, { AxiosRequestConfig, AxiosPromise, AxiosResponse } from 'axios';


//import * as http from 'http';
//var fs                = require('fs'),storageManagement = require('azure-asm-storage');
import  * as azure  from 'fast-azure-storage';
var urlSharedKey = 'https://storagelatintest.table.core.windows.net/Persona?st=2018-08-05T22%3A14%3A19Z&se=2018-08-20T22%3A14%3A00Z&sp=raud&sv=2018-03-28&tn=persona&sig=DxU3OGGkO092uET0JPt%2FWdZRmo2Cp3%2FSyCjXcLpP3yY%3D';
var urlAzureFunction = "https://miindicadorapi.azurewebsites.net/api/HttpTriggerJS1?code=HNrWahearYSovl/hZorLwdCmav1uz0eswO5BamXcYvsMHq15Kh5ulg==";
import 'office-ui-fabric/dist/components/DatePicker/DatePicker.min.css';
require("office-ui-fabric/dist/components/DatePicker/DatePicker.min.css")
export interface estados {
  resultados: Array<any>;
  indicadoresDelDia : any;
}
export default class Sar extends React.Component<ISarProps, estados> {
  
  constructor(){
    super();
    SPComponentLoader.loadCss('//dev.office.com/fabric-js/css/fabric.components.min.css');
    this.state = {
      resultados : [],
      indicadoresDelDia: null
    }
  }
  
  public componentDidMount(){
    this.getElements().then((response) => {
      console.log(response.data);
      this.setState({
        resultados : response.data.value
      });
    }).catch((err)=>{
        console.log(err);
    });
    this.getAzureFunction().then((resp)=>{
      this.setState({
        indicadoresDelDia : resp
      })

    })
    
  }
  public render(): React.ReactElement<ISarProps> {
    //this.getElements();
    return (
      <div>
        <Pivot>
          <PivotItem linkText="Ver elementos" itemCount={this.state.resultados.length}  itemIcon="Emoji2">
            <Label>Total Asistentes Charla SPFX:  {this.state.resultados.length}</Label>
            <table className="ms-Table">
              <thead>
                <tr>
                  <th>ID</th>
                  <th>Nombre</th>
                  <th>Comuna</th>
                  <th>Fecha confirmación</th>
                  <th>Teléfono</th>
                </tr>
              </thead>
              <tbody>
                {
                  this.state.resultados.map((value)=>{

                    return <tr>
                            <td>{value.PartitionKey}</td>
                            <td>{value.Nombre}</td>
                            <td>{value.Comuna}</td>
                            <td>{value.FechaConfirmacion}</td>
                            <td>{value.Telefono}</td>
                            
                          </tr>
                  })
                }
              </tbody>
            </table>
          </PivotItem>
          <PivotItem linkText="Insertar elementos" itemIcon="Recent">
            <Label>Pivot #2</Label>
            <div>
              <TextField label="Nombre:" name="Nombre" />
              <TextField label="Comuna:" name="Comuna" />
              <DatePicker
                //firstDayOfWeek={firstDayOfWeek}
                //strings={DayPickerStrings}
                placeholder="Seleccionar fecha confirmación..."
                showMonthPickerAsOverlay={true}
                // tslint:disable:jsx-no-lambda
                onAfterMenuDismiss={() => console.log('onAfterMenuDismiss called')}
                // tslint:enable:jsx-no-lambda
                label="Fecha confirmación:"
              />
              <TextField label="Teléfono:" name="Telefono" />
              <PrimaryButton onClick={this.addElement}>Submit</PrimaryButton>
            </div>
          </PivotItem>
          <PivotItem linkText="Azure API" itemIcon="Globe">
            <Label>Pivot #3</Label>
            <table className="ms-Table">
              <thead>
                <tr>
                  <th>Moneda</th>
                  <th>Valor CLP</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td>{this.state.indicadoresDelDia != null ? this.state.indicadoresDelDia.dolar.nombre : ""}</td>
                  <td>{this.state.indicadoresDelDia != null ? this.state.indicadoresDelDia.dolar.valor: ""}</td>
                </tr>
              </tbody>
            </table>
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
    var options = {
      accountId:          '...',
      accessKey:          '...'
    };
    /*var queue = new azure.Queue(options);
    var table = new azure.Table(options);
    var blob  = new azure.Blob(options);*/
    /*
    msRestAzure.interactiveLogin(function(err, credentials) {
      var client = new storageManagementClient.StorageManagementClient(credentials, 'your-subscription-id');
      client.storageAccounts.list(function(err, result) {
        if (err) console.log(err);
        console.log(result);
      });
     });*/
    /*
    segundo metodo
   var retryOperations = new azure.ExponentialRetryPolicyFilter();
    //var tableSvc = azure.createTableService().withFilter(retryOperations);
    var tableSvc = azure.createTableService("", "")
    var entGen = azure.TableUtilities.entityGenerator;
    var task = {
      PartitionKey: entGen.String('Persona'),
      RowKey: entGen.String('666'),
      description: entGen.String('take out the trash'),
      dueDate: entGen.DateTime(new Date(Date.UTC(2015, 6, 20))),
      Nombre: "asdasdasd"
    };
    var batch = new azure.TableBatch();

    batch.insertEntity(task, {echoContent: true});
    tableSvc.executeBatch('mytable', batch, function (error, result, response) {
      if(!error) {
        // Batch completed
      }
    });
*/
    
    //primer metodo
    const tablestorageUrl = urlSharedKey;
    axios.post(tablestorageUrl, {
      "PartitionKey": new Date(),
      "RowKey": new Date(),
      "Nombre": "asdasdxasd",
      "Comuna": "asdasdxasd",
      "FechaConfirmacion": "2018-08-06T19:15:12.033Z",
      "Telefono": "87654321"
            
    }, {
      url: tablestorageUrl,
      method: 'post',
      headers: {
        'Content-Type': 'application/json'
      }
    }).then(function (response) {
      debugger;
      console.log(response);
    })
    .catch(function (error) {
      debugger;
      console.log(error);
    });
  
    
  }
  private getElements()  {
    this.getAzureFunction();
    const tablestorageUrl = urlSharedKey;
    return axios.get(tablestorageUrl, {
      headers: {
        Accepts: 'application/json'
      }});
  }

  private getAzureFunction():Promise<any>{

    return new Promise<any>((resolve, reject) => {
      const requestHeaders: Headers = new Headers();
      requestHeaders.append("Content-type", "application/jsonp");
      requestHeaders.append("Cache-Control", "no-cache");
      const postOptions: IHttpClientOptions = {
        headers: requestHeaders
        //,credentials: "include"
      };

      this.props.context_.httpClient.get(urlAzureFunction, HttpClient.configurations.v1, postOptions).then((response: HttpClientResponse) => {
        response.json().then((resp: string) => {
            var responseText = resp != null ? JSON.parse(resp) : null;
            resolve(responseText);
          })
          .catch ((response: any) => {
            reject(response);
        });
      });
    });

   
  }

}

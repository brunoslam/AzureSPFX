import * as React from 'react';
import styles from './Sar.module.scss';
import { ISarProps } from './ISarProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http';
/* Pivot Office Fabric */
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { PivotItem, IPivotItemProps, Pivot, TextField} from 'office-ui-fabric-react';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import axios, { AxiosRequestConfig, AxiosPromise, AxiosResponse } from 'axios';


var urlSharedKey = 'https://storagelatintest.table.core.windows.net/Persona?st=2018-08-05T22%3A14%3A19Z&se=2018-08-20T22%3A14%3A00Z&sp=raud&sv=2018-03-28&tn=persona&sig=DxU3OGGkO092uET0JPt%2FWdZRmo2Cp3%2FSyCjXcLpP3yY%3D';
var urlAzureFunction = "https://miindicadorapi.azurewebsites.net/api/HttpTriggerJS1?code=HNrWahearYSovl/hZorLwdCmav1uz0eswO5BamXcYvsMHq15Kh5ulg==";
import 'office-ui-fabric/dist/components/DatePicker/DatePicker.min.css';
import 'office-ui-fabric/dist/components/DatePicker/DatePicker.min.css';
import 'office-ui-fabric-core/dist/css/fabric.min.css';
require("office-ui-fabric-core/dist/css/fabric.min.css");
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons'
export interface estados {
  resultados: Array<any>;
  indicadoresDelDia : any;
  fechaSeleccionada: any;
}
export default class Sar extends React.Component<ISarProps, estados> {
  
  constructor(){
    super();
    SPComponentLoader.loadCss('//dev.office.com/fabric-js/css/fabric.components.min.css');
    SPComponentLoader.loadCss('https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/6.0.0/css/fabric-6.0.0.scoped.css');

    this.state = {
      resultados : [],
      indicadoresDelDia: null,
      fechaSeleccionada: null
    }
  }
  
  public componentWillMount(){
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
                  <th>Confirma</th>
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
                            <td>{value.Confirma ? <i className="ms-Icon ms-Icon--ActivateOrders" aria-hidden="true"></i> : <i className="ms-Icon ms-Icon--EntryDecline" aria-hidden="true"></i>}</td>
                          </tr>
                  })
                }
              </tbody>
            </table>
          </PivotItem>
          <PivotItem linkText="Insertar elementos" itemIcon="Recent">
            <Label>Pivot #2</Label>
            <div>
              <TextField label="Nombre:" name="Nombre" id="txtNombre" />
              <TextField label="Comuna:" name="Comuna" id="txtComuna"/>
              <DatePicker
                //firstDayOfWeek={firstDayOfWeek}
                //strings={DayPickerStrings}
                placeholder="Seleccionar fecha confirmación..."
                showMonthPickerAsOverlay={true}
                // tslint:disable:jsx-no-lambda
                onAfterMenuDismiss={() => console.log('onAfterMenuDismiss called')}
                // tslint:enable:jsx-no-lambda
                label="Fecha confirmación:"
                className="txtFechaConfirmacion"
                formatDate={this._onFormatDate}
              />
              <TextField label="Teléfono:" name="Telefono" id="txtTelefono"/>
              <Toggle
                defaultChecked={false}
                label="Confirma asistencia:"
                onText="Si"
                offText="No"
                id="chkAsistencia"
              />
              <PrimaryButton onClick={this.addElement}>Submit</PrimaryButton>
            </div>
          </PivotItem>
          <PivotItem linkText="Azure API" itemIcon="Globe">
            <Label>Fecha información: {this.state.indicadoresDelDia == null ? "" : this.state.indicadoresDelDia.fecha }</Label>
            <table className="ms-Table">
              <thead>
                <tr>
                  <th>Moneda</th>
                  <th>Valor CLP</th>
                </tr>
              </thead>
              {(() => {
                if(this.state.indicadoresDelDia){
                  return <tbody>
                          <tr>
                            <td>{ this.state.indicadoresDelDia.dolar.nombre }</td>
                            <td>{ this.state.indicadoresDelDia.dolar.valor }</td>
                          </tr>
                          <tr>
                            <td>{ this.state.indicadoresDelDia.euro.nombre }</td>
                            <td>{ this.state.indicadoresDelDia.euro.valor }</td>
                          </tr>
                          <tr>
                            <td>{ this.state.indicadoresDelDia.bitcoin.nombre }</td>
                            <td>{ (this.state.indicadoresDelDia.bitcoin.valor * this.state.indicadoresDelDia.dolar.valor) }</td>
                          </tr>
                          <tr>
                            <td>{ this.state.indicadoresDelDia.uf.nombre }</td>
                            <td>{ this.state.indicadoresDelDia.uf.valor }</td>
                          </tr>
                          <tr>
                            <td>{ this.state.indicadoresDelDia.utm.nombre }</td>
                            <td>{ this.state.indicadoresDelDia.utm.valor }</td>
                          </tr>
                        </tbody>
                }
              })()}
              
            </table>
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
  private _onFormatDate = (date: Date): string => {
    this.setState({fechaSeleccionada : date});
    return date.getDate() + '/' + (date.getMonth() + 1) + '/' + (date.getFullYear() % 100);
  };
  private addElement(){
    
    var nombre = (document.getElementById("txtNombre")  as HTMLInputElement).value;
    var comuna = (document.getElementById("txtComuna")  as HTMLInputElement).value;
    var fechaConfirmacion = (document.querySelector(".txtFechaConfirmacion input[type=text]")  as HTMLInputElement).value;
    var telefono = (document.getElementById("txtTelefono")  as HTMLInputElement).value;
    var confirmacion = (document.getElementById("chkAsistencia")  as HTMLInputElement).getAttribute("aria-pressed");
    debugger;
    //primer metodo
    const tablestorageUrl = urlSharedKey;
    axios.post(tablestorageUrl, {
      "PartitionKey": new Date(),
      "RowKey": new Date(),
      "Nombre": nombre,
      "Comuna": comuna,
      "FechaConfirmacion": fechaConfirmacion,
      "Telefono": telefono,
      "Confirma" : confirmacion
            
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

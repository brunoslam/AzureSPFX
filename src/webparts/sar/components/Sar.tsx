import * as React from 'react';
import styles from './Sar.module.scss';
import { ISarProps } from './ISarProps';
import { escape } from '@microsoft/sp-lodash-subset';

/* Pivo Office Fabric */
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { PivotItem, IPivotItemProps, Pivot, TextField} from 'office-ui-fabric-react';

export default class Sar extends React.Component<ISarProps, {}> {
  public render(): React.ReactElement<ISarProps> {
    return (
      <div>
        <Pivot>
          <PivotItem linkText="Ver elementos"  itemIcon="Emoji2">
            <Label>Pivot #1</Label>
            <table className="ms-Table">
              <thead>
                <tr>
                  <th>Location</th>
                  <th>Modified</th>
                  <th>Type</th>
                  <th>File Name</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td>Location</td>
                  <td>Modified</td>
                  <td>Type</td>
                  <td>File Name</td>
                </tr>
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
}

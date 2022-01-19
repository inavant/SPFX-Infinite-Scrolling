import * as React from 'react';
import styles from './InifiniteGrid.module.scss';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import { IInifiniteGridProps } from './IInifiniteGridProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ConstrainMode, DetailsList, DetailsListLayoutMode, IColumn, Spinner } from 'office-ui-fabric-react';
import { SPPagedResponse } from './PagedResponse';

export interface InifiniteGridState {
  listData:any[];
  nextPageToken: string;
}

export default class InifiniteGrid extends React.Component<IInifiniteGridProps, InifiniteGridState> {
  private _columns: IColumn[];

  constructor(props: IInifiniteGridProps) {
    super(props);

    this._columns = [
      { key: 'Title', name: 'Title', fieldName: 'Title', minWidth: 200, isResizable: true },
      { key: 'Categoria', name: 'Categoria', fieldName: 'Category', minWidth: 200, isResizable: true },
      { key: 'Fecha', name: 'Fecha', fieldName: 'Date', minWidth: 200, isResizable: true }
    ];

    this.state = {
      listData:[],
      nextPageToken:''
    };
  }

  public componentDidMount() {
    //Realizamos la consulta inicial
    this.getSharePointData('').then((pagedResult) => {
      this.setState({
        listData: pagedResult.items,
        nextPageToken: pagedResult.nextPageToken
      });
    });
  }

  //Realizamos la consulta a SharePoint
  private getSharePointData(nextPageToken?:string){
    return new Promise<SPPagedResponse>((resolve, reject): void => {
    sp.web.lists.getByTitle("BigList").renderListDataAsStream({
      ViewXml: `<View>
        <ViewFields>
          <FieldRef Name="Title"/>
          <FieldRef Name="Categoria"/>
          <FieldRef Name="Fecha"/>
        </ViewFields>
        <RowLimit Paged="TRUE">200</RowLimit>
        </View>`,
      Paging: nextPageToken
    }).then(pagedResponse => {
        //Obtenemos los elementos de la consulta
        let resultItems = pagedResponse.Row.map((item)=>{
          let newItem: any = {
            Title: item["Title"],
            Category:item["Categoria"],
            Date:item["Fecha"]
          };
          return newItem;
        });

        //Obtenemos el token de referencia a la siguiente página
        nextPageToken = pagedResponse.NextHref && pagedResponse.NextHref.length ? pagedResponse.NextHref.split('?')[1] : null;

        //En caso de que exista una siguiente página se agrega un elemento nulo, para indicar al grid que se debe cargar más elementos
        if(nextPageToken != null && nextPageToken != "" ){
          //Se agrega un elemento nulo, para indicarle al grid que debe cargar mas elementos
          resultItems.push(null);
        }

        let result = new SPPagedResponse(resultItems,nextPageToken);
        resolve(result);
      });
    });
  }

  public LoadMoreItems = (nextPageData) => {
    let currentListData = this.state.listData;
    if(nextPageData != null && nextPageData != "" ){
      this.getSharePointData(nextPageData).then((pagedResult) => {

        //Quitamos el elemento nulo, para evitar que el grid siga cargando el mismo set de elementos
        currentListData = currentListData.filter((el) => { return el != null; });
        currentListData = [...currentListData, ... pagedResult.items];
        
        //Actualizamos el estado con los nuevos elementos.
        this.setState({
          listData:currentListData,
          nextPageToken:pagedResult.nextPageToken
        });

      });
    }
  }

  public render(): React.ReactElement<IInifiniteGridProps> {
    return (
      <div className={ styles.inifiniteGrid }>
        <div className={styles.gridContainer} data-is-scrollable="true">
          <DetailsList
            items={this.state.listData}
            columns={this._columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.fixedColumns}
            constrainMode={ConstrainMode.unconstrained}
            onShouldVirtualize = { () => true}
            onRenderMissingItem={ (index, rowData) => {
              this.LoadMoreItems(this.state.nextPageToken);
              return null;
          } }
          />
        </div>
      </div>
    );
  }
}

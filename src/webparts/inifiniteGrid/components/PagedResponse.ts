export class SPPagedResponse {
    public items: any[];
    public nextPageToken: string;

    constructor(items:any[], nextPageToken:string) {
      this.items = items;
      this.nextPageToken =nextPageToken;
    }
  }
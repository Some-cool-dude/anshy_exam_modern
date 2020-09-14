import * as React from 'react';
import styles from './AnshyWebPart.module.scss';
import { sp } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAnshyWebPartProps {
  description: string;
  context: WebPartContext;
}

export interface IAnshyWebPartState {
  items: any[];
}

export default class AnshyWebPart extends React.Component<IAnshyWebPartProps, IAnshyWebPartState> {
  constructor(props) {
    super(props);
    this.state = {
      items: []
    };
  }

  public async getData() {
    const res = await sp.web.lists.getByTitle("Holidays").items.select(
      "Id", 
      "Title",
      "Category",
      "IsNonWorkingDay",
      "V4HolidayDate",
      "Worker/Title",
      "Worker/Id",
      "Places/Id",
      "Places/Title").expand("Places", "Worker").get();
      this.setState({items: res});
      console.log(res);
  }

  public async addItem() {
    await sp.web.lists.getByTitle("Holidays").items.add({
      Title: "Default",
      Category: "Summer",
      IsNonWorkingDay: true,
      WorkerId: 12,
      V4HolidayDate: (new Date()).toISOString(),
      PlacesId: {
        results: [1, 2]
      }
    });
    this.getData();
  }

  public async updateItem() {
    let list = sp.web.lists.getByTitle("Holidays");

    const i = await list.items.getById(1).update({
      Title: "My New Title"
    });
  }

  public async componentDidMount() {
    sp.setup({
      spfxContext: this.props.context
    });
    this.getData();
  }

  public render(): React.ReactElement<IAnshyWebPartProps> {
    return (
      <>
        <button className={styles.btn__add} onClick={() => this.addItem()}>Add Default</button>
        <table className={styles.table}>
        <tr className={styles.table__row}>
          <th className={styles.table__head}>Title</th>
          <th className={styles.table__head}>Category</th>
          <th className={styles.table__head}>Non-working day</th>
          <th className={styles.table__head}>Date</th>
          <th className={styles.table__head}>Worker</th>
          <th className={styles.table__head}>Places</th>
        </tr>
        {this.state.items.map((item) =>
          <tr className={styles.table__row} key={item.Id}>
            <td className={styles.table__cell}>{item.Title}</td>
            <td className={styles.table__cell}>{item.Category}</td>
            <td className={styles.table__cell}>{item.IsNonWorkingDay}</td>
            <td className={styles.table__cell}>{new Date(item.V4HolidayDate).toDateString()}</td>
            <td className={styles.table__cell}>{item.Worker.Title}</td>
            <td className={styles.table__cell}>{item.Places.map(place => `${place.Title}; `)}</td>
          </tr>
        )}
      </table>
      </>
    );
  }
}

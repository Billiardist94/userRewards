/* eslint-disable no-case-declarations */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import UserData from "../models/UserData";
import { sp } from "@pnp/sp/presets/all";
export interface IRetrieveListDataService {
    retrieveListItems(): Promise<Array<any>>;
}

export default class RetrieveListDataService implements IRetrieveListDataService {
    private _listTitle: string;
    // private _ctx: WebPartContext;

    constructor(ctx: WebPartContext, listTitle: string) {
        this._listTitle = listTitle;
        // this._ctx = ctx;
    }

    public async getUserData(authorId: number) {
        try {
            const user = await sp.web.getUserById(authorId)();
            return user;
        } catch (error) {
            console.error('Error fetching user data:', error);
            throw error;
        }
    }

    public async retrieveListItems(): Promise<Array<UserData>> {
        try {
            const items = await sp.web.lists.getByTitle(this._listTitle).items.get();
            const result: UserData[] = [];
            for (const item of items) {
                const { EmployeeId: employee, Reward: reward } = item;
                const userData = await this.getUserData(employee);

                const listItem: UserData = {
                    employee: userData,
                    reward: reward
                };
                
                result.push(listItem);
            }
            console.log('Result:', result);
            return result;
        } catch (error) {
            console.error('Error retrieving list items:', error);
            throw error;
        }
    }
}
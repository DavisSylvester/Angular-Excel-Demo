import { Component } from '@angular/core';
import * as OfficeHelpers from '@microsoft/office-js-helpers';
import { isContext } from 'vm';

const template = require('./app.component.html');

@Component({
    selector: 'app-home',
    template
})
export default class AppComponent {
    welcomeMessage = 'Welcome';

    async getRow() {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getActiveWorksheet();
            let range = context.workbook.worksheets.getItem('Sheet1').getRange("A2");
            let range2 = context.workbook.worksheets.getItem('Sheet1').getRange("A10");

            range.load(['values', 'address']);
            

            console.log(`getRow: Value: ${range}`);

            await context.sync();

            console.log(`Value: ${range.address}`);
            this.welcomeMessage = range.values[0][0];
            // range2.values = [[range.values[0][0] ]];
            // range2.values = [[this.welcomeMessage ]];

        
            range2.values = [[range.values[0][0] ]];
            
            await context.sync();

        });
    }

    async setColor() {
        try {
            await Excel.run(async context => {
                const range = context.workbook.getSelectedRange();
                range.load('address');
                range.format.fill.color = 'green';
                await context.sync();
                console.log(`The range address was ${range.address}.`);
            });
        } catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
    }


    async deSetColor() {
        try {
            await Excel.run(async context => {
                const range = context.workbook.getSelectedRange();
                
                range.load('address');
                range.format.fill.color = 'white';
                await this.getRow();
                await context.sync();
                console.log(`The range address was ${range.address}.`);
            });
        } catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
    }

}
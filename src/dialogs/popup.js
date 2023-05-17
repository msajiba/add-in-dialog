/* eslint-disable */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("data-submit").onclick = dataSubmit;
    }
});

export async function dataSubmit() {
    const inputDate = document.getElementById("inputDate").value;
    const inputMerchant = document.getElementById("inputMerchant").value;
    const inputCategory = document.getElementById("inputCategory").value;
    const inputAmount = document.getElementById("inputAmount").value;
    const list = [inputDate, inputMerchant, inputCategory, inputAmount];
    Office.context.ui.messageParent(list.toString());
}



export async function sendStringToParentPage() {
    try {
        await Excel.run(async (context) => {

            const userName = document.getElementById("name-box").value;
            Office.context.ui.messageParent(userName);
        });
    } catch (error) {
        console.error(error);
    }
}
/* eslint-disable */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
    // if (info.host === Office.HostType.Excel) {
    document.getElementById("ok-button").onclick = sendStringToParentPage;
    // }
});



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
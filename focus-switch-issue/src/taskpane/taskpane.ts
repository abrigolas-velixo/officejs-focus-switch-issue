/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import {fromByteArray} from "base64-js";

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("run").onclick = run;
    }
});

export async function run() {
    copyCurrentWorkbook();
}


function copyCurrentWorkbook(): void {
    const sliceSize = 4096;

    Office.context.document.getFileAsync(Office.FileType.Compressed, {sliceSize: sliceSize}, async function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
            throw result.error;
        }

        // Result.value is the File object.
        const fileContents = await getFileContents(result.value);

        let base64string = fromByteArray(new Uint8Array(fileContents));

        await Excel.createWorkbook(base64string);
        await Office.addin.showAsTaskpane();
    });
}

async function getFileContents(
    file: Office.File
): Promise<any[]> {
    let expectedSliceCount = file.sliceCount;
    let fileSlices: Array<Array<number>> = [];

    await new Promise((resolve, reject) => {
        getFileContentsHelper();

        function getFileContentsHelper() {
            file.getSliceAsync(fileSlices.length, function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    file.closeAsync();
                    reject(result.error);
                }

                fileSlices.push(result.value.data);

                if (fileSlices.length == expectedSliceCount) {
                    file.closeAsync();

                    let array = [];
                    fileSlices.forEach((slice) => {
                        array = array.concat(slice);
                    });

                    return resolve(array);
                } else {
                    getFileContentsHelper();
                }
            });
        }
    });

}
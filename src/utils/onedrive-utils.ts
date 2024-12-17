import { graphClient } from "@/app/provider/identity-provider";
import {
    FileUpload,
    OneDriveLargeFileUploadOptions,
    OneDriveLargeFileUploadTask,
    UploadResult,
} from "@microsoft/microsoft-graph-client";
import { DriveItem } from "@microsoft/microsoft-graph-types";
import { readFile } from "fs/promises";
import { basename } from "path";

const filePath = "E:onedrivesdkonedrive-demoms.Thảo.xlsx";

const targetFolderPath = "E:onedrivesdkonedrive-demodownload";
// readFile from fs/promises
const file = await readFile(filePath);
// basename from path
const fileName = basename(filePath);

const options: OneDriveLargeFileUploadOptions = {
    // Relative path from root folder
    path: targetFolderPath,
    fileName: fileName,
    rangeSize: 1024 * 1024,
    uploadEventHandlers: {
        // Called as each "slice" of the file is uploaded
        progress: (range, _) => {
            console.log(
                `Uploaded bytes ${range?.minValue} to ${range?.maxValue}`
            );
        },
    },
};

// Create FileUpload object
const fileUpload = new FileUpload(file, fileName, file.byteLength);
// Create a OneDrive upload task
const uploadTask = await OneDriveLargeFileUploadTask.createTaskWithFileObject(
    graphClient,
    fileUpload,
    options
);

// Do the upload
const uploadResult: UploadResult = await uploadTask.upload();

// The response body will be of the corresponding type of the
// item being uploaded. For OneDrive, this is a DriveItem
const driveItem = uploadResult.responseBody as DriveItem;
console.log(`Uploaded file with ID: ${driveItem.id}`);

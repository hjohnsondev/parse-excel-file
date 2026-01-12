
import * as ExcelJS from 'exceljs';
import { google } from 'googleapis'; 
import { Readable } from 'stream';

export const getFileFromGoogleDrive = async (fileName: string): Promise<Buffer> => {
    const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY!);
    const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID!; // Store your folder ID here

    const auth = new google.auth.GoogleAuth({
        credentials,
        scopes: ['https://www.googleapis.com/auth/drive.readonly']
    });

    const drive = google.drive({ version: 'v3', auth });

    // Search for the file within the specific folder
    const searchResponse = await drive.files.list({
        q: `name='${fileName}' and '${folderId}' in parents and trashed=false`,
        fields: 'files(id, name, mimeType, createdTime, modifiedTime)',
        spaces: 'drive',
        orderBy: 'modifiedTime desc' // Get most recent first
    });

    const files = searchResponse.data.files;
    
    if (!files || files.length === 0)
        return Promise.reject(new Error(`File '${fileName}' not found in the shared folder`));

    // Use the most recently modified file if multiple exist
    const file = files[0]; 
    const fileId = file.id!;
            
    console.log(`Found file: ${file.name} (${fileId}), modified: ${file.modifiedTime}`);

    // Download the file
    const response = await drive.files.get(
        {
            fileId: fileId,
            alt: 'media'
        },
        { responseType: 'arraybuffer' }
    );

    const buffer = Buffer.from(response.data as ArrayBuffer);

    return buffer;
};

export const uploadFileToGoogleDrive = async (fileName: string, buffer: Buffer, mimeType: string): Promise<string> => {
    const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY!);
    const folderId = process.env.GOOGLE_DRIVE_UPLOADS_FOLDER_ID!; // Store your folder ID here

    const auth = new google.auth.GoogleAuth({
        credentials,
        scopes: ['https://www.googleapis.com/auth/drive.file']
    });

    const drive = google.drive({ version: 'v3', auth });

    const fileMetadata = {
        name: fileName,
        parents: [folderId]
    };

    const media = {
        mimeType,
        body: Readable.from(buffer)
    };

    const response = await drive.files.create({
        requestBody: fileMetadata,
        media: media,
        fields: 'id',
        supportsAllDrives: true
    });

    return response.data.id!;
}
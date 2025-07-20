# Bulk Upload SPFx

![alt](/src/webparts/bulkUploadSpFx/assets/SPFx%20solution%20-%20Bulk%20upload%20items%20in%20SharePoint%20list%20using%20Excel.png)

A simple SPFx solution built with React for bulk data addition to SharePoint lists. Originally developed for adding testing data quickly, this solution can also be useful in projects that require bulk data addition.

## Overview

This solution leverages:

- **React** for the user interface.
- **PnPjs** for data operations with SharePoint.
- **SheetJS** to read data from Excel files.

It provides two upload methods:

1. **Sequential Add:** Uploads one item at a time using async/await.
2. **Parallel Add:** Uploads all items concurrently using `Promise.all`.
   To make it more reusable for you guys, I customized it a little bit and added these two uploading methods.

## Requirements

- [Node.js](https://nodejs.org/) (v18.x.x or above)
- [PnPjs](https://pnp.github.io/pnpjs/)
- [SheetJS](https://sheetjs.com/)

## Development Environment

1. Node version = 18.20.3
2. Operating System = Windows 11

## Libraries Installation Commands

1. **PnPjs**
   ```bash
   npm install @pnp/sp @pnp/graph --save
   ```
2. **SheetJS**

   ```bash
   npm i https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz --save
   ```

3. **PnP SPFx Property Controls**

   ```bash
   npm install @pnp/spfx-property-controls --save
   ```

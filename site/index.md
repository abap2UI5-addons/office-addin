---
layout: home

hero:
  name: "abap2UI5 for Excel"
  text: "Show an app in Excel's task pane and read and write the workbook."
  image:
    src: /favicon.png
    alt: abap2UI5
  actions:
    - theme: brand
      text: Get started
      link: /guide
    - theme: alt
      text: GitHub
      link: https://github.com/abap2UI5-addons/office-addin
    - theme: alt
      text: All addons
      link: https://abap2ui5-addons.github.io/

features:
  - title: "Any abap2UI5 app runs in Excel's task pane, started by the manifest's --app option"
  - title: "z2ui5.cc.ExcelBridge writes an ABAP table to a cell, the selection or a new sheet"
  - title: "Reads the selected cells into an ABAP table, with an optional selection-change event"
  - title: "Values keep their types: text, numbers and ISO dates, leading zeros preserved"
  - title: "npm run manifest writes the add-in manifest for your host, client, app and theme"
  - title: "Deployment via the Microsoft 365 admin center or sideloading on Windows, Mac and web"
  - title: "Serve the host page from an approuter or reverse proxy where there is no BSP"
---

## What it does

An Office add-in that opens an abap2UI5 app in the task pane of Microsoft Excel: the real abap2UI5 UI5 frontend, served by your SAP system through a host page in the BSP Z2UI5_XL. The app stays an ABAP class and talks to the workbook through the custom control z2ui5.cc.ExcelBridge, which writes tables into sheets and reads selected cells. The repository holds the host page, a demo app, the manifest generator and the tests. It is a proof of concept: CI tests it against a stand-in for Office.js, nobody has run it in a real Excel yet.

## Requirements

- abap2UI5 with z2ui5.cc.ExcelBridge (pull request #2847, on main since commit bcce7d6; the first release after 1.146.0 that lists it)
- The embed control needs abap2UI5 1.145.0 or later
- UI5 1.71 or later on the system; theme sap_horizon needs UI5 1.102+, else --theme sap_fiori_3
- HTTPS with a certificate the users' machines trust
- Excel: Microsoft 365, Excel 2019 or later on Windows or Mac, Excel on the web best effort (ExcelApi 1.4)
- ABAP Cloud: no BSP applications, pull only src/01 and serve the host page without a BSP
- Node.js 22 for npm run manifest

## Installation

With [abapGit](https://abapgit.org), from [the repository](https://github.com/abap2UI5-addons/office-addin):

- Pull with abapGit into a package of its own (for example $Z2UI5_XL) on the system that runs abap2UI5; it brings Z2UI5_CL_XL_DEMO, the BSP Z2UI5_XL and the ICF nodes /sap/bc/ui5_ui5/sap/z2ui5_xl and /sap/bc/bsp/sap/z2ui5_xl
- In SICF check that /sap/bc/ui5_ui5/sap/z2ui5_xl is active with the same logon settings as /sap/bc/z2ui5
- Try it in a browser first: https://&lt;sap-host&gt;/sap/bc/ui5_ui5/sap/z2ui5_xl/index.html?app=Z2UI5_CL_XL_DEMO
- Write the manifest with npm run manifest -- --host https://my.sap.example --app Z2UI5_CL_XL_DEMO --client 100 --out manifest.xml and deploy it to users
- For Excel on the web set the frame-ancestors and SameSite cookie headers on the system

The [Guide](/guide) has every step.

## A first app

```abap
page->tag( n = `ExcelBridge` ns = `z2ui5`
    )->a( n = `id`               v = `excel`
    )->a( n = `rows`             v = client->_bind( materials )
    )->a( n = `columns`          v = client->_bind( columns )
    )->a( n = `target`           v = `newSheet`
    )->a( n = `sheetName`        v = client->_bind( sheet_name )
    )->a( n = `selection`        v = client->_bind( selection )
    )->a( n = `selectionAddress` v = client->_bind( selection_address )
    )->a( n = `OnWritten`        v = client->_event( val = `XL_WRITTEN` arg = `${$parameters>/address}` )
    )->a( n = `OnRead`           v = client->_event( val = `XL_READ` arg = `${$parameters>/address}` )
    )->a( n = `OnError`          v = client->_event( val   = `XL_ERROR`
                                                     t_arg = VALUE #( ( `${$parameters>/message}` )
                                                                      ( `${$parameters>/code}` ) ) ) ).
```

## Links

- [abap2UI5/embed-control](https://github.com/abap2UI5/embed-control)
- [abap2UI5](https://github.com/abap2UI5/abap2UI5)
- [The repository](https://github.com/abap2UI5-addons/office-addin) - issues and pull requests welcome

# Excel Viewer for VS Code

Open `.xls` and `.xlsx` files as read-only spreadsheets in VS Code. Choose a sheet at the top and page through rows. `.xlsx` supports basic font color, fill color, bold, italic, alignment and number formatting. `.xls` displays formatted cell text, including number formats. Formulas show their saved result; this extension does not calculate formulas or edit workbooks.

The viewer displays up to 200 columns per sheet and 100 rows per page. Some advanced styles, theme colors, merged cells, images and charts are not rendered. Password-protected workbooks are not supported.

## Install

Download the `.vsix` from [GitHub Releases](https://github.com/autotrof/vscode-excel/releases) and run **Extensions: Install from VSIX...** in VS Code. Open an Excel file from the Explorer. You can use **Reopen Editor With... → Excel Viewer** if another editor is selected.

## Development

```sh
npm ci
npm run check
npm test
npm run package
```

Press F5 in VS Code to run an Extension Development Host. To publish a release, update `package.json` version, commit, then push the matching tag (for example `v0.1.0`). GitHub Actions checks and packages the extension and attaches its VSIX to a GitHub Release. The workflow can also be run manually to build a downloadable Actions artifact without creating a Release.

## License

GPL-3.0-only. See [LICENSE](LICENSE).

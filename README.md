# bcp_excel

## Publishing

```powershell
dotnet publish -r win-x64 --sc -c Release
```

## Usage idea

```powershell
bcp_excel import <excel-filename> into <db..table> -S server -U user -P password [other options]
bcp_excel export <db..table> into <excel-filename> -S server -U user -P password [other options]

[other options]
-T table
-D database

[other options export]
-Q "query-text"
```

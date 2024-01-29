## Steps followed in the sample:

1. Load the template document.
2. Hook the BeforeClearFieldEvent.
3. Throw the exception for the field that is not in the data source. Check this with the [HasMappedFieldInDataSource](https://help.syncfusion.com/cr/file-formats/Syncfusion.DocIO.DLS.BeforeClearFieldEventArgs.html#Syncfusion_DocIO_DLS_BeforeClearFieldEventArgs_HasMappedFieldInDataSource) API.
4. Execute Mail merge.
5. Save the document.
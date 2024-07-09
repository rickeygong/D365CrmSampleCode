if (gdh === undefined) {
    var gdh = {};
}
if (gdh.d365 === undefined) {
    gdh.d365 = {};
}
gdh.d365.InvoiceDetail = (function () {
    'use strict';
    return {
        OpenImportData: function (primaryControl) {
            let objFormContext = primaryControl;
            let paramsObject = {
                'recordGuid': objFormContext.data.entity.getId().replace('{', '').replace('}', ''),
                'entityName': objFormContext.data.entity.getEntityName(),
                'subgridName': "Invoice_detail",
            };
            let pageInput = {
                pageType: "webresource",
                webresourceName: "gdh_/html/Import_details.html",
                data: JSON.stringify(paramsObject)
            };
            let navigationOptions = {
                target: 2,
                width: 500, // value specified in pixel
                height: 350, // value specified in pixel
                position: 1,
                title: "Import data"
            };
            Xrm.Navigation.navigateTo(pageInput, navigationOptions);
        }
    }
})();
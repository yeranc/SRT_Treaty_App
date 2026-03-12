sap.ui.define([
    "sap/m/MessageToast",
    "sap/ui/export/Spreadsheet"
], function (MessageToast, Spreadsheet) {
    'use strict';

    return {
        get_result: function (oEvent) {
            var oModel = this.getView().getModel();
            var aContexts = this.extensionAPI.getSelectedContexts();
            var that = this;

            var aResults = [];
            var iCompleted = 0;

            this.getView().setBusy(true);
            for (let index = 0; index < aContexts.length; index++) {
                var oData = aContexts[index].getObject();

                var oPayload = {
                    vtgnr: oData.vtgnr,
                    PeriodStartDate: oData.PeriodStartDate,
                    SectionNumber: oData.SectionNumber
                };

                oModel.callFunction("/show_res", {
                    method: "POST",
                    urlParameters: oPayload,
                    success: function (oResult) {
                        var oResultData = oResult.results;
                        // store result
                        aResults.push(oResultData);
                        iCompleted++;

                        // when all calls finished
                        if (iCompleted === aContexts.length) {
                            that.getView().setBusy(false);
                            that._showResultPopup(aResults);
                        }
                    },

                    error: function () {
                        iCompleted++;
                        if (iCompleted === aContexts.length) {
                            that._showResultPopup(aResults);
                        }
                        MessageToast.show("Error calling backend");
                    }
                });
            }
        },

        _showResultPopup: function (aData) {
            // Flatten nested arrays
            var aFlatData = aData.flat();
            var sthat = this;

            var oModel = new sap.ui.model.json.JSONModel();
            oModel.setData({ results: aFlatData });

            var oTable = new sap.m.Table({
                columns: [
                    new sap.m.Column({
                        header: new sap.m.Text({ text: "Source Treaty Number" }),
                        width: "40%"
                    }),
                    new sap.m.Column({
                        header: new sap.m.Text({ text: "Process Ref ID" }),
                        width: "40%"
                    }),
                    new sap.m.Column({
                        header: new sap.m.Text({ text: "Target Treaty Number" }),
                        width: "40%"
                    }),
                    new sap.m.Column({
                        header: new sap.m.Text({ text: "Target System" }),
                        width: "30%"
                    }),
                    new sap.m.Column({
                        header: new sap.m.Text({ text: "Message Type" }),
                        width: "30%"
                    }),
                    new sap.m.Column({
                        header: new sap.m.Text({ text: "Message" }),
                        width: "80%"
                    })
                ]
            });

            oTable.setModel(oModel);

            oTable.bindItems({
                path: "/results",
                template: new sap.m.ColumnListItem({
                    cells: [
                        new sap.m.Text({ text: "{vtgnr}" }),
                        new sap.m.Text({ text: "{processingID}" }),
                        new sap.m.Text({ text: "{CreatedTreaty}" }),
                        new sap.m.Text({ text: "{syst}" }),
                        new sap.m.Text({ text: "{type}" }),
                        new sap.m.Text({ text: "{message}" })
                    ]
                })
            });

            var oDialog = new sap.m.Dialog({
                title: "Treaty Copy Logs",
                content: [oTable],
                buttons: [
                    new sap.m.Button({
                        text: "Export",
                        type: "Emphasized",
                        press: function () {
                            sthat._exportToExcel(aFlatData);
                        }
                    }),
                    new sap.m.Button({
                        text: "Close",
                        press: function () {
                            oDialog.close();
                        }
                    })
                ],
                afterClose: function () {
                    oDialog.destroy();
                }
            });

            oDialog.open();
        },

        _exportToExcel: function (aData) {
            var aCols = [
                { label: "Source Treaty Number", property: "vtgnr" },
                { label: "Process Ref ID", property: "processingID" },
                { label: "Target Treaty Number", property: "CreatedTreaty" },
                { label: "Target System", property: "syst" },
                { label: "Message Type", property: "type" },
                { label: "Message", property: "message" }
            ];

            var oSettings = {
                workbook: {
                    columns: aCols
                },
                dataSource: aData,
                fileName: "Treaty_Copy_Results.xlsx"
            };

            var oSheet = new Spreadsheet(oSettings);
            oSheet.build().finally(function () {
                oSheet.destroy();
            });
        }
    }
});

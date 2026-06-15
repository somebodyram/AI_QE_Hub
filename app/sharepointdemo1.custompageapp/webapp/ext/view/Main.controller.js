sap.ui.define([
  "sap/ui/core/mvc/Controller",
  "sap/m/MessageBox",
  "sap/ui/model/json/JSONModel",
  "sap/ui/model/Filter",
  "sap/ui/model/FilterOperator"
], function (Controller, MessageBox, JSONModel, Filter, FilterOperator) {
  "use strict";

  return Controller.extend(
    "sharepointdemo1.custompageapp.ext.view.Main",
    {

      /* =========================================================== */
      /* Lifecycle                                                   */
      /* =========================================================== */
      onInit: function () {
        // Holds data shown in the table
        const oExcelModel = new JSONModel([]);
        this.getView().setModel(oExcelModel, "excelModel");

         // Disable Export until Execute has run
        this.getView().addEventDelegate({
          onAfterRendering: function () {
            this.byId("btnExportMenu").setEnabled(false);
          }.bind(this)
        });
      },

      /* =========================================================== */
      /* SEARCH FUNCTIONALITY                                        */
      /* =========================================================== */
      // Internal state to track both filters
      _sSearchQuery: "",
      _sStatusKey: "",

      onSearch: function (oEvent) {
        this._sSearchQuery = oEvent.getParameter("query") || "";
        this._applyFilters();
      },

      onStatusFilter: function (oEvent) {
        this._sStatusKey = oEvent.getParameter("selectedItem").getKey();
        this._applyFilters();
      },

      _applyFilters: function () {
        var oBinding = this.byId("resultTable").getBinding("items");
        if (!oBinding) return;
        var aFilters = [];

        if (this._sSearchQuery) {
          aFilters.push(new Filter({
            filters: [
              new Filter("scenario",    FilterOperator.Contains, this._sSearchQuery),
              new Filter("orderNumber", FilterOperator.Contains, this._sSearchQuery)
            ],
            and: false
          }));
        }

        if (this._sStatusKey) {
          aFilters.push(new Filter("status", FilterOperator.EQ, this._sStatusKey));
        }

        var oBinding = this.byId("resultTable").getBinding("items");
        // AND between status + search, each internally uses OR where needed
        oBinding.filter(aFilters.length ? [new Filter({ filters: aFilters, and: true })] : []);
      },

      /* =========================================================== */
      /* 1. READ EXCEL (STRICT 'all rounder')                        */
      /* =========================================================== */
      onReadExcel: function () {
        const oFile = this._getSelectedFile();
        if (!oFile) {
          MessageBox.warning("Please select an Excel file.");
          return;
        }

        const oFormData = new FormData();
        oFormData.append("file", oFile);

        fetch("/excel/read", {
          method: "POST",
          body: oFormData
        })
          .then(response => response.json())
          .then(data => {
            this.getView()
              .getModel("excelModel")
              .setData(data);

            MessageBox.success(
              "Read Excel completed. Rows: " + data.length
            );
          })
          .catch(err => {
            console.error(err);
            MessageBox.error("Failed to read Excel.");
          });
      },

      /* =========================================================== */
      /* 2. PROCESS EXCEL (USE UI DATA + DUMMY API)                  */
      /* =========================================================== */
      onProcessExcel: function () {
        const oTable = this.byId("resultTable");
        const aSelectedItems = oTable.getSelectedItems();

        if (!aSelectedItems || aSelectedItems.length === 0) {
          MessageBox.warning("Please select at least one row to execute.");
          return;
        }

        // ✅ Disable button immediately
        const oBtn = this.byId("btnExecute");
        oBtn.setEnabled(false);
        oBtn.setText("Executing...");

        // We MUST keep track of the paths so we know which rows to update later!
        const aSelectedData = [];
        const aPaths = []; 

        aSelectedItems.forEach(function(oItem) {
            var oContext = oItem.getBindingContext("excelModel");
            aSelectedData.push(oContext.getObject());
            aPaths.push(oContext.getPath()); // Example: saves "/4" or "/5"
        });

        fetch("/excel/save", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(aSelectedData)
        })
          .then(res => res.json())
          .then(result => { // named it 'result' so result.data and result.success work properly
            if (result.success) {
              MessageBox.success(result.saved + " row(s) saved to DB successfully!");

              // Get the model so we can update the UI
              var oModel = this.getView().getModel("excelModel");

              // Update ONLY the selected rows with the fresh data from HANA
              result.data.forEach(function(savedRow, index) {
                  var sPath = aPaths[index]; 
                  oModel.setProperty(sPath, savedRow); 
              });

              // Remove the checkmarks after successful execution
              oTable.removeSelections();

              // ─── Enable Save button after Execute completes ───────────────────────────
              this.byId("btnExportMenu").setEnabled(true);

            } else {
              MessageBox.error("Save failed: " + result.error);
            }
          })
          .catch(err => {
            console.error(err);
            MessageBox.error("Failed to execute.");
          })
          .finally(() => {
            // ✅ Re-enable button after done
            oBtn.setEnabled(true);
            oBtn.setText("Execute");
          });
      },

      /* =========================================================== */
      /* 3. PRESS REASON                                           */
      /* =========================================================== */
      onPressReason: function (oEvent) {
        // Get the specific row's data
        var oSource = oEvent.getSource();
        var oContext = oSource.getBindingContext("excelModel");
        var sReasonText = oContext.getProperty("reason");

        // Create the dialog lazily
        if (!this._oReasonDialog) {
            this._oReasonDialog = new sap.m.Dialog({
                title: "Reason Details",
                contentWidth: "500px",
                contentHeight: "300px",
                resizable: true,
                content: new sap.m.TextArea({
                    value: sReasonText,
                    width: "100%",
                    height: "100%",
                    editable: false, // Acts like a read-only notepad
                    growing: true
                }),
                endButton: new sap.m.Button({
                    text: "Close",
                    press: function () {
                        this._oReasonDialog.close();
                    }.bind(this)
                })
            });
            // Add dependent so the dialog has access to the view's models
            this.getView().addDependent(this._oReasonDialog);
        } else {
            // If dialog exists, just update the text area value
            this._oReasonDialog.getContent()[0].setValue(sReasonText);
        }

        this._oReasonDialog.open();
    },

    /* =========================================================== */
      /* 4. PRESS RUN STATUS                                         */
      /* =========================================================== */
      onPressRunStatus: function (oEvent) {
        // Get the specific row's data
        var oSource = oEvent.getSource();
        var oContext = oSource.getBindingContext("excelModel");
        var sRunStatusText = oContext.getProperty("runStatus");

        // Create the dialog lazily
        if (!this._oRunStatusDialog) {
            this._oRunStatusDialog = new sap.m.Dialog({
                title: "Execution Run Status Logs",
                contentWidth: "700px",  // Slightly wider for system logs
                contentHeight: "450px", // Slightly taller
                resizable: true,
                content: new sap.m.TextArea({
                    value: sRunStatusText,
                    width: "100%",
                    height: "100%",
                    editable: false, // Read-only terminal view
                    growing: true
                }),
                endButton: new sap.m.Button({
                    text: "Close",
                    press: function () {
                        this._oRunStatusDialog.close();
                    }.bind(this)
                })
            });
            // Add dependent so the dialog has access to the view's models
            this.getView().addDependent(this._oRunStatusDialog);
        } else {
            // If dialog exists, just update the text area value with the new row's data
            this._oRunStatusDialog.getContent()[0].setValue(sRunStatusText);
        }

        this._oRunStatusDialog.open();
      },

      /* =========================================================== */
      /* SHAREPOINT FETCH LOGIC                                      */
      /* =========================================================== */
      onFetchSharepoint: function () {
      var oView  = this.getView();
      var oInput = this.byId("sharepointUrlInput");
      var sUrl   = oInput ? oInput.getValue().trim() : "";

      if (!sUrl) {
        MessageBox.warning("Please paste a SharePoint Excel URL.");
        return;
      }

      oView.setBusy(true);

      fetch("/sharepoint/data", {
        method:  "POST",
        headers: { "Content-Type": "application/json" },
        body:    JSON.stringify({ sharepointUrl: sUrl })
      })
        .then(res => res.json())
        .then(resp => {
          oView.setBusy(false);
          if (resp.error) {
            MessageBox.error("Failed to load: " + resp.error);
            return;
          }
          oView.getModel("excelModel").setData(resp.data);
          MessageBox.success("Loaded " + resp.data.length + " row(s) from SharePoint.");
        })
        .catch(() => {
          oView.setBusy(false);
          MessageBox.error("SharePoint fetch failed.");
        });
    },

    // ─── Save to SharePoint ───────────────────────────────────────────────────
    onSaveToSharePoint: function () {
      var oView  = this.getView();
      var oInput = this.byId("sharepointUrlInput");
      var sUrl   = oInput ? oInput.getValue().trim() : "";

      if (!sUrl) {
        MessageBox.warning("SharePoint URL is missing. Please fetch data first.");
        return;
      }

      // Collect ALL rows that have a status (i.e. were executed)
      var aData = oView.getModel("excelModel").getData();
      var aUpdates = aData.filter(function(row) {
        return row.status && row.status !== "";
      }).map(function(row) {
        return {
          sheet:     row.sheet,
          rowNumber: row.rowNumber,
          status:    row.status,
          reason:    row.reason || ""
        };
      });

      if (!aUpdates.length) {
        MessageBox.warning("No executed rows to save. Please execute rows first.");
        return;
      }

      oView.setBusy(true);
      var oBtn = this.byId("btnExportMenu");
      oBtn.setEnabled(false);
      oBtn.setText("Saving...");

      fetch("/sharepoint/save", {
        method:  "POST",
        headers: { "Content-Type": "application/json" },
        body:    JSON.stringify({ sharepointUrl: sUrl, updates: aUpdates })
      })
        .then(res => res.json())
        .then(resp => {
          oView.setBusy(false);
          oBtn.setText("Export");
          if (resp.error) {
            oBtn.setEnabled(true);
            MessageBox.error("Save failed: " + resp.error);
            return;
          }
          MessageBox.success(
            resp.updated + " row(s) written back to SharePoint successfully!"
          );
          // Keep button disabled — no unsaved changes remain
        })
        .catch(() => {
          oView.setBusy(false);
          oBtn.setEnabled(true);
          oBtn.setText("Export");
          MessageBox.error("Failed to save to SharePoint.");
        });
    },

    /* =========================================================== */
      /* WRITE BACK TO LOCAL EXCEL FILE                              */
      /* =========================================================== */
      onUpdateLocalExcel: function () {
        const oFile = this._getSelectedFile();
        if (!oFile) {
          MessageBox.warning("Please select the original Excel file via the file uploader first.");
          return;
        }

        var oView = this.getView();
        var aData = oView.getModel("excelModel").getData();

        if (!aData || aData.length === 0) {
          MessageBox.warning("No data found in the table to write back.");
          return;
        }

        oView.setBusy(true);

        // 1. Gather current execution status and reasons from the UI model
        var aUpdates = aData.map(function(row) {
          return {
            sheet:     row.sheet,
            rowNumber: row.rowNumber,
            status:    row.status || "",
            reason:    row.reason || ""
          };
        });

        // 2. Wrap file and JSON updates into FormData payload
        const oFormData = new FormData();
        oFormData.append("file", oFile);
        oFormData.append("updates", JSON.stringify(aUpdates));

        // 3. Post to Node backend and handle response as file download blob
        fetch("/excel/writeback", {
          method: "POST",
          body: oFormData
        })
          .then(response => {
            if (!response.ok) throw new Error("Failed to process Excel file update.");
            return response.blob();
          })
          .then(blob => {
            oView.setBusy(false);
            
            // 4. Use the EXACT original file name
            const sNewFileName = oFile.name; 
            const oLink = document.createElement("a");
            oLink.href = URL.createObjectURL(blob);
            oLink.download = sNewFileName; // Will download as "CMEINTAKE-4286- Execution sheet 1.xlsx"
            oLink.click();
            
            MessageBox.success("File generated. If prompted, select your original folder to overwrite it!");
          })
          .catch(err => {
            oView.setBusy(false);
            console.error(err);
            MessageBox.error("Error updating local file: " + err.message);
          });
      },

      /* =========================================================== */
      /*  DOWNLOAD EXCEL (WRITE BACK ENRICHED DATA)                */ 
      /* =========================================================== */
      //NOT IN USE


      /* =========================================================== */
      /* Helper: Get selected file from FileUploader                 */
      /* =========================================================== */
      _getSelectedFile: function () {
        const oUploader = this.byId("excelUploader");
        if (!oUploader) {
          return null;
        }

        const oDomRef = oUploader.getDomRef();
        if (!oDomRef) {
          return null;
        }

        const oInput = oDomRef.querySelector("input[type='file']");
        if (!oInput || !oInput.files || oInput.files.length === 0) {
          return null;
        }

        return oInput.files[0];
      }

    }
  );
});
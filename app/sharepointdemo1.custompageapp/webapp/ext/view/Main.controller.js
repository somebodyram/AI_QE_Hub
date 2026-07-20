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

      onSheetFilter: function (oEvent) {
        this._sSheetKey = oEvent.getParameter("selectedItem").getKey();
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

        if (this._sSheetKey) {
          aFilters.push(new Filter("sheet", FilterOperator.EQ, this._sSheetKey));
        }

        // var oBinding = this.byId("resultTable").getBinding("items");
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
            this.getView().getModel("excelModel").setData(data);
            this._updateSheetFilter(data);

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
            var oRowData = oContext.getObject();
            oRowData.action = "INITIAL_PURCHASE";
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
      /* REFRESH DATA (FETCH LATEST STATUSES FROM HANA)              */
      /* =========================================================== */
      onRefreshData: function () {
        var oView = this.getView();
        var oModel = oView.getModel("excelModel");
        var aCurrentData = oModel.getData();

        // Prevent fetching if the table is completely empty
        if (!aCurrentData || aCurrentData.length === 0) {
          MessageBox.information("No data to refresh. Please load a file first.");
          return;
        }

        var oBtn = this.byId("btnRefresh");
        oView.setBusy(true);
        if (oBtn) oBtn.setEnabled(false);

        // Fetch all recent records from your HANA DB endpoint
        fetch("/excel/results", { method: "GET" })
          .then(res => res.json())
          .then(dbData => {
            if (dbData.error) throw new Error(dbData.error);

            // 1. Create a quick lookup map from the DB data for fast merging
            // We use fileName + sheet + rowNumber to perfectly identify the exact row
            var oDbMap = {};
            dbData.forEach(function(dbRow) {
              if (dbRow.fileName && dbRow.sheet && dbRow.rowNumber) {
                var sKey = dbRow.fileName + "|" + dbRow.sheet + "|" + dbRow.rowNumber;
                oDbMap[sKey] = dbRow;
              }
            });

            // 2. Loop through the UI table and update rows that have newer DB data
            var iUpdatedCount = 0;
            aCurrentData.forEach(function(localRow) {
              if (localRow.fileName && localRow.sheet && localRow.rowNumber) {
                var sKey = localRow.fileName + "|" + localRow.sheet + "|" + localRow.rowNumber;
                var dbMatch = oDbMap[sKey];
                
                // If this row exists in the database, overwrite the UI with the DB's latest status/logs
                if (dbMatch) {
                  localRow.status = dbMatch.status;
                  localRow.reason = dbMatch.reason;
                  localRow.runStatus = dbMatch.runStatus;
                  localRow.id = dbMatch.id; // Sync the DB UUID
                  iUpdatedCount++;
                }
              }
            });

            // 3. Force the table to re-render with the updated data
            oModel.refresh(true);
            
            // Re-apply filters just in case statuses changed (e.g., PENDING to PASS)
            this._applyFilters();

            MessageBox.success("Refresh complete! " + iUpdatedCount + " row(s) synced with the database.");
          })
          .catch(err => {
            console.error(err);
            MessageBox.error("Failed to refresh data: " + err.message);
          })
          .finally(() => {
            oView.setBusy(false);
            if (oBtn) oBtn.setEnabled(true);
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
          this._updateSheetFilter(resp.data);
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
      /* Helper: Populate Dynamic Sheet Dropdown                     */
      /* =========================================================== */
      _updateSheetFilter: function (aData) {
        var oSelect = this.byId("sheetFilter");
        oSelect.removeAllItems();
        
        // Add the default "All Sheets" option back
        oSelect.addItem(new sap.ui.core.Item({ key: "", text: "All Sheets" }));

        if (!aData || !aData.length) return;

        // Find unique sheet names
        var aUniqueSheets = [];
        aData.forEach(function (row) {
          if (row.sheet && aUniqueSheets.indexOf(row.sheet) === -1) {
            aUniqueSheets.push(row.sheet);
          }
        });

        // Add unique sheets to the dropdown
        aUniqueSheets.forEach(function (sheetName) {
          oSelect.addItem(new sap.ui.core.Item({ key: sheetName, text: sheetName }));
        });
      },


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
      },

    /* =========================================================== */
      /* AI CHATBOT WIDGET                                           */
      /* =========================================================== */
      onOpenChatbot: function (oEvent) {
        var oButton = oEvent.getSource();

        // Create the popover lazily (only build it the first time it's clicked)
        if (!this._oChatPopover) {
            
            // 1. The message feed container
            this._oChatLog = new sap.m.VBox({
                items: [
                    new sap.m.MessageStrip({
                        text: "Hi! I am your AI Assistant. How can I help you analyze these execution results?",
                        type: "Information",
                        showIcon: true
                    }).addStyleClass("sapUiSmallMargin")
                ]
            });

            // 2. The input field
            var oChatInput = new sap.m.Input({
                placeholder: "Ask me anything...",
                width: "100%",
                submit: function(oEvent) { 
                  // TODO: Add logic here later to send message to your AI backend
                  sap.m.MessageToast.show("Message sent to AI: " + oEvent.getParameter("value"));
                  oEvent.getSource().setValue(""); // clear input
                }
            });

            // 3. Assemble the Popover window
            this._oChatPopover = new sap.m.Popover({
                title: "QE Hub AI Assistant",
                placement: "Top", // Opens upwards from the footer button
                contentWidth: "350px",
                contentHeight: "450px",
                resizable: true,
                content: [
                    new sap.m.VBox({
                        height: "100%",
                        justifyContent: "SpaceBetween",
                        items: [
                            new sap.m.ScrollContainer({
                                height: "380px",
                                vertical: true,
                                content: [this._oChatLog]
                            }),
                            new sap.m.Toolbar({
                                content: [
                                    oChatInput,
                                    new sap.m.Button({
                                        icon: "sap-icon://paper-plane",
                                        type: "Emphasized",
                                        press: function () {
                                            oChatInput.fireSubmit({ value: oChatInput.getValue() });
                                        }
                                    })
                                ]
                            })
                        ]
                    })
                ]
            });
            // Attach to view so it inherits models
            this.getView().addDependent(this._oChatPopover);
        }

        // Open anchored to the button
        this._oChatPopover.openBy(oButton);
      }

    }
  );
});
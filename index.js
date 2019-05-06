"use strict";
/*
    This file is part of the Microsoft PowerApps code samples.
    Copyright (C) Microsoft Corporation.  All rights reserved.
    This source code is intended only as a supplement to Microsoft Development Tools and/or
    on-line documentation.  See these other materials for detailed information regarding
    Microsoft code samples.

    THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
    EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
    MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
 */
Object.defineProperty(exports, "__esModule", { value: true });
// Define const here
var RowRecordId = "rowRecId";
// Style name of Load More Button
var DataSetControl_LoadMoreButton_Hidden_Style = "DataSetControl_LoadMoreButton_Hidden_Style";
var TSDataSetGrid = /** @class */ (function () {
    /**
     * Empty constructor.
     */
    function TSDataSetGrid() {
    }
    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
     */
    TSDataSetGrid.prototype.init = function (context, notifyOutputChanged, state, container) {
        // Need to track container resize so that control could get the available width. The available height won't be provided even this is true
        context.mode.trackContainerResize(true);
        // Create main table container div. 
        this.mainContainer = document.createElement("div");
        // Create data table container div. 
        this.gridContainer = document.createElement("div");
        this.gridContainer.classList.add("DataSetControl_grid-container");
        // Create data table container div. 
        this.loadPageButton = document.createElement("button");
        this.loadPageButton.setAttribute("type", "button");
        this.loadPageButton.innerText = context.resources.getString("PCF_DataSetControl_LoadMore_ButtonLabel");
        this.loadPageButton.classList.add(DataSetControl_LoadMoreButton_Hidden_Style);
        this.loadPageButton.classList.add("DataSetControl_LoadMoreButton_Style");
        this.loadPageButton.addEventListener("click", this.onLoadMoreButtonClick.bind(this));
        // Adding the main table and loadNextPage button created to the container DIV.
        this.mainContainer.appendChild(this.gridContainer);
        this.mainContainer.appendChild(this.loadPageButton);
        this.mainContainer.classList.add("DataSetControl_main-container");
        container.appendChild(this.mainContainer);
    };
    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    TSDataSetGrid.prototype.updateView = function (context) {
        this.contextObj = context;
        this.toggleLoadMoreButtonWhenNeeded(context.parameters.dataSetGrid);
        if (!context.parameters.dataSetGrid.loading) {
            // Get sorted columns on View
            var columnsOnView = this.getSortedColumnsOnView(context);
            if (!columnsOnView || columnsOnView.length === 0) {
                return;
            }
            while (this.gridContainer.firstChild) {
                this.gridContainer.removeChild(this.gridContainer.firstChild);
            }
            this.gridContainer.appendChild(this.createGridBody(columnsOnView, context.parameters.dataSetGrid));
        }
        // this is needed to ensure the scroll bar appears automatically when the grid resize happens and all the tiles are not visible on the screen.
        this.mainContainer.style.maxHeight = window.innerHeight - this.gridContainer.offsetTop - 75 + "px";
    };
    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    TSDataSetGrid.prototype.getOutputs = function () {
        return {};
    };
    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    TSDataSetGrid.prototype.destroy = function () {
    };
    /**
     * Get sorted columns on view
     * @param context
     * @return sorted columns object on View
     */
    TSDataSetGrid.prototype.getSortedColumnsOnView = function (context) {
        if (!context.parameters.dataSetGrid.columns) {
            return [];
        }
        var columns = context.parameters.dataSetGrid.columns
            .filter(function (columnItem) {
            // some column are supplementary and their order is not > 0
            return columnItem.order >= 0;
        });
        // Sort those columns so that they will be rendered in order
        columns.sort(function (a, b) {
            return a.order - b.order;
        });
        return columns;
    };
    /**
     * funtion that creates the body of the grid and embeds the content onto the tiles.
     * @param columnsOnView columns on the view whose value needs to be shown on the UI
     * @param gridParam data of the Grid
     */
    TSDataSetGrid.prototype.createGridBody = function (columnsOnView, gridParam) {
        var gridBody = document.createElement("div");
        if (gridParam.sortedRecordIds.length > 0) {
            var _loop_1 = function (currentRecordId) {
                var gridRecord = document.createElement("div");
                gridRecord.classList.add("DataSetControl_grid-item");
                gridRecord.addEventListener("click", this_1.onRowClick.bind(this_1));
                var gridTable = document.createElement("table");
                headRow = document.createElement("tr");
                firstCol = document.createElement("td");
                secondCol = document.createElement("td");
                firstCol.style.padding = "20px 10px";
                firstCol.style.textAlign = "center";
                secondCol.style.padding = "20px";
                secondCol.style.textAlign = "justify";
                headRow.appendChild(firstCol);
                headRow.appendChild(secondCol);
                gridTable.appendChild(headRow);
                gridRecord.appendChild(gridTable);
                // Set the recordId on the row dom
                gridRecord.setAttribute(RowRecordId, gridParam.records[currentRecordId].getRecordId());
                columnsOnView.forEach(function (columnItem, index) {
                    var labelPara = document.createElement("p");
                    labelPara.classList.add("DataSetControl_grid-text-label");
                    var valuePara = document.createElement("p");
                    valuePara.classList.add("DataSetControl_grid-text-value");
                    labelPara.textContent = columnItem.displayName + " : ";
                    if (gridParam.records[currentRecordId].getFormattedValue(columnItem.name) != null && gridParam.records[currentRecordId].getFormattedValue(columnItem.name) != "") {
                        if (columnItem.displayName == "Verified" && gridParam.records[currentRecordId].getFormattedValue(columnItem.name) == "Yes") {
                            gridRecord.style.backgroundColor = 'green';
                        }
                        valuePara.textContent = gridParam.records[currentRecordId].getFormattedValue(columnItem.name);
                        if (columnItem.displayName == "Comments") {
                            secondCol.appendChild(labelPara);
                            secondCol.appendChild(valuePara);
                        }
                        else {
                            firstCol.appendChild(labelPara);
                            firstCol.appendChild(valuePara);
                        }
                    }
                    else {
                        valuePara.textContent = "-";
                        firstCol.appendChild(valuePara);
                    }
                });
                gridBody.appendChild(gridRecord);
            };
            var this_1 = this, headRow, firstCol, secondCol;
            for (var _i = 0, _a = gridParam.sortedRecordIds; _i < _a.length; _i++) {
                var currentRecordId = _a[_i];
                _loop_1(currentRecordId);
            }
        }
        else {
            var noRecordLabel = document.createElement("div");
            noRecordLabel.classList.add("DataSetControl_grid-norecords");
            noRecordLabel.style.width = this.contextObj.mode.allocatedWidth - 25 + "px";
            noRecordLabel.innerHTML = this.contextObj.resources.getString("PCF_DataSetControl_No_Record_Found");
            gridBody.appendChild(noRecordLabel);
        }
        return gridBody;
    };
    /**
     * Row Click Event handler for the associated row when being clicked
     * @param event
     */
    TSDataSetGrid.prototype.onRowClick = function (event) {
        var rowRecordId = event.currentTarget.getAttribute(RowRecordId);
        if (rowRecordId) {
            var entityReference = this.contextObj.parameters.dataSetGrid.records[rowRecordId].getNamedReference();
            var entityFormOptions = {
                entityName: entityReference.name,
                entityId: entityReference.id,
            };
            this.contextObj.navigation.openForm(entityFormOptions);
        }
    };
    /**
     * Toggle 'LoadMore' button when needed
     */
    TSDataSetGrid.prototype.toggleLoadMoreButtonWhenNeeded = function (gridParam) {
        if (gridParam.paging.hasNextPage && this.loadPageButton.classList.contains(DataSetControl_LoadMoreButton_Hidden_Style)) {
            this.loadPageButton.classList.remove(DataSetControl_LoadMoreButton_Hidden_Style);
        }
        else if (!gridParam.paging.hasNextPage && !this.loadPageButton.classList.contains(DataSetControl_LoadMoreButton_Hidden_Style)) {
            this.loadPageButton.classList.add(DataSetControl_LoadMoreButton_Hidden_Style);
        }
    };
    /**
     * 'LoadMore' Button Event handler when load more button clicks
     * @param event
     */
    TSDataSetGrid.prototype.onLoadMoreButtonClick = function (event) {
        this.contextObj.parameters.dataSetGrid.paging.loadNextPage();
        this.toggleLoadMoreButtonWhenNeeded(this.contextObj.parameters.dataSetGrid);
    };
    return TSDataSetGrid;
}());
exports.TSDataSetGrid = TSDataSetGrid;
//# sourceMappingURL=index.js.map
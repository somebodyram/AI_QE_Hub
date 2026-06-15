sap.ui.define(['sap/fe/test/ListReport'], function(ListReport) {
    'use strict';

    var CustomPageDefinitions = {
        actions: {},
        assertions: {}
    };

    return new ListReport(
        {
            appId: 'sharepointdemo1.myapplication',
            componentId: 'empdataList',
            contextPath: '/empdata'
        },
        CustomPageDefinitions
    );
});
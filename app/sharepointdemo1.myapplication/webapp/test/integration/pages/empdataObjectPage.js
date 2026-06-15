sap.ui.define(['sap/fe/test/ObjectPage'], function(ObjectPage) {
    'use strict';

    var CustomPageDefinitions = {
        actions: {},
        assertions: {}
    };

    return new ObjectPage(
        {
            appId: 'sharepointdemo1.myapplication',
            componentId: 'empdataObjectPage',
            contextPath: '/empdata'
        },
        CustomPageDefinitions
    );
});